import assert from "node:assert/strict";
import { accessSync, constants } from "node:fs";
import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import { delimiter, extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";
import { chromium } from "playwright-core";

const clientRoot = fileURLToPath(new URL("../dist/client/", import.meta.url));

function resolveChromiumExecutable() {
  if (process.env.CHROMIUM_BIN) return process.env.CHROMIUM_BIN;

  const names =
    process.platform === "win32"
      ? ["chromium.exe", "chrome.exe", "msedge.exe"]
      : ["chromium", "chromium-browser", "google-chrome", "google-chrome-stable"];
  const candidates = (process.env.PATH ?? "")
    .split(delimiter)
    .filter(Boolean)
    .flatMap((directory) => names.map((name) => join(directory, name)));
  if (process.platform === "darwin") {
    candidates.push(
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
      "/Applications/Chromium.app/Contents/MacOS/Chromium",
    );
  }

  for (const candidate of candidates) {
    try {
      accessSync(candidate, constants.X_OK);
      return candidate;
    } catch {
      // Continue through the deterministic candidate list.
    }
  }
  throw new Error("Chromium not found; set CHROMIUM_BIN to an executable path");
}

const chromiumExecutable = resolveChromiumExecutable();

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

async function assetResponse(request) {
  const url = new URL(request.url);
  const pathname = decodeURIComponent(url.pathname).replace(/^\/+/, "");
  const filePath = resolve(clientRoot, pathname);
  if (!filePath.startsWith(clientRoot)) {
    return new Response("Not found", { status: 404 });
  }
  try {
    return new Response(await readFile(filePath), {
      headers: {
        "content-type": contentTypes[extname(filePath)] ?? "application/octet-stream",
      },
    });
  } catch {
    return new Response("Not found", { status: 404 });
  }
}

async function startSiteServer() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("browser-test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  const server = createServer(async (incoming, outgoing) => {
    const origin = `http://127.0.0.1:${server.address().port}`;
    const url = new URL(incoming.url ?? "/", origin);
    try {
      let response;
      if (url.pathname === "/_vinext/image") {
        const source = url.searchParams.get("url");
        response = source
          ? await assetResponse(new Request(new URL(source, origin)))
          : new Response("Not found", { status: 404 });
      } else if (
        url.pathname.startsWith("/assets/") ||
        url.pathname.startsWith("/cases/") ||
        url.pathname.startsWith("/_vinext_fonts/")
      ) {
        response = await assetResponse(new Request(url));
      } else {
        response = await worker.fetch(
          new Request(url, {
            method: incoming.method,
            headers: incoming.headers,
          }),
          { ASSETS: { fetch: assetResponse } },
          {
            waitUntil() {},
            passThroughOnException() {},
          },
        );
      }
      outgoing.writeHead(response.status, Object.fromEntries(response.headers));
      outgoing.end(Buffer.from(await response.arrayBuffer()));
    } catch (error) {
      outgoing.writeHead(500, { "content-type": "text/plain" });
      outgoing.end(String(error));
    }
  });

  await new Promise((resolveReady) => server.listen(0, "127.0.0.1", resolveReady));
  return {
    server,
    url: `http://127.0.0.1:${server.address().port}`,
  };
}

test("filters cases and manages the detail dialog focus", async (context) => {
  const { server, url } = await startSiteServer();
  let browser;
  context.after(async () => {
    await browser?.close();
    await new Promise((resolveClosed) => server.close(resolveClosed));
  });
  browser = await chromium.launch({
    executablePath: chromiumExecutable,
    headless: true,
    args: ["--no-sandbox"],
  });

  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  await page.goto(url, { waitUntil: "networkidle" });

  await page.getByRole("button", { name: "Excel", exact: true }).click();
  assert.equal(await page.locator(".case-card").count(), 3);
  assert.equal(await page.locator(".case-card.xlsx").count(), 3);

  const trigger = page
    .locator(".case-card.xlsx")
    .first()
    .getByRole("button", { name: /查看详情/ });
  await trigger.click();

  const dialog = page.getByRole("dialog");
  await dialog.waitFor({ state: "visible" });
  assert.equal(
    await page.evaluate(() => document.activeElement?.getAttribute("aria-label")),
    "关闭案例详情",
  );

  await page.keyboard.press("Escape");
  await dialog.waitFor({ state: "hidden" });
  assert.equal(
    await page.evaluate(() => document.activeElement?.textContent?.trim()),
    "查看详情 →",
  );

  await page.setViewportSize({ width: 390, height: 844 });
  await page.reload({ waitUntil: "networkidle" });
  assert.equal(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= document.documentElement.clientWidth,
    ),
    true,
  );
});

test("keeps human criterion scores blank, local, and case-specific", async (context) => {
  const { server, url } = await startSiteServer();
  let browser;
  context.after(async () => {
    await browser?.close();
    await new Promise((resolveClosed) => server.close(resolveClosed));
  });
  browser = await chromium.launch({
    executablePath: chromiumExecutable,
    headless: true,
    args: ["--no-sandbox"],
  });

  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  const consoleErrors = [];
  page.on("console", (message) => {
    if (message.type() === "error") consoleErrors.push(message.text());
  });
  await page.goto(`${url}/rubric`, { waitUntil: "networkidle" });

  assert.equal(
    await page.getByRole("button", { name: "5", exact: true }).first().isEnabled(),
    true,
  );
  assert.equal(
    await page.getByRole("button", { name: "5", exact: true }).first().getAttribute(
      "aria-pressed",
    ),
    "false",
  );

  await page.getByRole("button", { name: "5", exact: true }).first().click();
  await page.getByText(/1 \/ 60 已填写/).waitFor();
  await page.reload({ waitUntil: "networkidle" });

  assert.equal(
    await page.getByRole("button", { name: "5", exact: true }).first().getAttribute(
      "aria-pressed",
    ),
    "true",
  );

  await page.getByRole("button", { name: "下一张", exact: true }).click();
  assert.equal(
    await page.getByRole("button", { name: "5", exact: true }).first().getAttribute(
      "aria-pressed",
    ),
    "false",
  );

  await page.getByRole("tab", { name: "Content Accuracy", exact: true }).click();
  assert.equal(await page.locator(".fg-ai-score strong", { hasText: "N/A" }).count(), 20);

  await page.getByRole("tab", { name: "Word · 12", exact: true }).click();
  assert.equal(await page.locator(".fg-case-nav select option").count(), 12);
  assert.match(
    (await page.locator(".fg-slide-sticky img").getAttribute("src")) ?? "",
    /benchmark-units-v3/,
  );
  assert.equal(await page.locator("[data-subquestion-id]").count(), 20);
  assert.equal(await page.locator(".fg-rollup-strip > div:not(.fg-rollup-label)").count(), 5);

  await page.getByRole("tab", { name: "Excel · 12", exact: true }).click();
  assert.equal(await page.locator(".fg-case-nav select option").count(), 12);
  assert.match(
    (await page.locator(".fg-slide-sticky img").getAttribute("src")) ?? "",
    /benchmark-units-v3/,
  );

  await page.setViewportSize({ width: 390, height: 844 });
  await page.reload({ waitUntil: "networkidle" });
  assert.equal(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= document.documentElement.clientWidth,
    ),
    true,
  );
  assert.deepEqual(consoleErrors, []);
});
