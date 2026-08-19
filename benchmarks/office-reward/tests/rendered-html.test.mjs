import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { access, readFile, stat } from "node:fs/promises";
import test from "node:test";

async function render(pathname = "/") {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request(`http://localhost${pathname}`, {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async () => new Response("Not found", { status: 404 }),
      },
    },
    {
      waitUntil() {},
      passThroughOnException() {},
    },
  );
}

test("server-renders the Office Reward progress dashboard", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>Office Reward Build Log<\/title>/i);
  assert.match(html, /Office reward 全链路完成/);
  assert.match(html, /总体完成度/);
  assert.match(html, />678</);
  assert.match(html, /PowerPoint/);
  assert.match(html, /Word/);
  assert.match(html, /Excel/);
  assert.match(html, /XLSX adapter/);
  assert.match(html, /score-office CLI/);
  assert.match(html, /BUILD COMPLETE/);
  assert.match(html, /已合并 main/);
  assert.match(html, /CASE GALLERY/);
  assert.match(html, /全部/);
  assert.match(html, /Executive Review/);
  assert.match(html, /Quarterly Report/);
  assert.match(html, /Revenue Model/);
  assert.match(html, /EVALUATION MECHANISM/);
  assert.match(html, /PROMPT EXPLORER/);
  assert.match(html, /reward_0_1/);
  assert.doesNotMatch(html, /\/_vinext\/image/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/);
});

test("server-renders the four-region real-score annotation workbench", async () => {
  const response = await render("/rubric");
  assert.equal(response.status, 200);

  const html = await response.text();
  assert.match(html, /Office 细粒度评分实验/);
  assert.match(html, /54 个真实 Office 单元/);
  assert.match(html, /2,160/);
  assert.match(html, /PPT.*30/);
  assert.match(html, /Word.*12/);
  assert.match(html, /Excel.*12/);
  assert.match(html, /子问题/);
  assert.match(html, /板块小分/);
  assert.match(html, /4 个直接子问题/);
  assert.match(html, /评分细则/);
  assert.match(html, /真实 Office 单元/);
  assert.match(html, /GPT-5\.5 直接子问题分/);
  assert.match(html, /人工评分/);
  assert.match(html, /Spearman/);
  assert.match(html, /N\/A/);
  assert.match(html, /slide_0437/);
  assert.match(html, /布局与构图/);
  assert.match(html, /1–5/);
  assert.equal((html.match(/data-subquestion-id=/g) ?? []).length, 20);
  assert.doesNotMatch(html, /AI Prompt|当前维度原始分/);
});

test("ships real format assets and removes the starter preview", async () => {
  const [page, layout, css, packageJson, viteConfig, gallery, evaluation, rubric] =
    await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/layout.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
    readFile(new URL("../vite.config.ts", import.meta.url), "utf8"),
    readFile(new URL("../app/case-gallery.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/evaluation-explorer.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/rubric/rubric-workbench.tsx", import.meta.url), "utf8"),
  ]);

  assert.match(page, /ppt-reward-sample\.png/);
  assert.match(page, /word-reward-sample\.png/);
  assert.match(page, /excel-reward-sample\.png/);
  assert.match(layout, /Office Reward Build Log/);
  assert.match(css, /\.format-grid/);
  assert.match(css, /\.case-grid/);
  assert.match(css, /@media \(max-width: 640px\)/);
  assert.match(viteConfig, /"\.trycloudflare\.com"/);
  assert.match(gallery, /role="dialog"/);
  assert.match(gallery, /aria-modal="true"/);
  assert.match(gallery, /Aesthetics/);
  assert.match(gallery, /Content Accuracy/);
  assert.match(gallery, /Communication/);
  assert.match(evaluation, /复制 Prompt/);
  assert.match(evaluation, /multiDimensionInstructions/);
  assert.match(rubric, /localStorage/);
  assert.match(rubric, /officeSubquestionResults/);
  assert.match(rubric, /humanScores/);
  assert.doesNotMatch(rubric, /aiSubscore/);
  assert.doesNotMatch(packageJson, /react-loading-skeleton/);

  await Promise.all([
    access(new URL("../public/ppt-reward-sample.png", import.meta.url)),
    access(new URL("../public/word-reward-sample.png", import.meta.url)),
    access(new URL("../public/excel-reward-sample.png", import.meta.url)),
  ]);
  await assert.rejects(access(new URL("../app/_sites-preview", import.meta.url)));
});

test("publishes authentic prompt snapshots with matching hashes", async () => {
  const content = JSON.parse(
    await readFile(new URL("../app/evaluation-content.json", import.meta.url), "utf8"),
  );

  assert.equal(
    Object.values(content.weights).reduce((sum, value) => sum + value, 0),
    1,
  );
  assert.equal(content.stages.length, 4);
  assert.equal(content.statuses.length, 3);

  for (const format of Object.values(content.prompts)) {
    for (const dimension of Object.values(format.dimensions)) {
      assert.ok(dimension.prompt.length > 1_000);
      assert.equal(
        createHash("sha256").update(dimension.prompt).digest("hex"),
        dimension.hash,
      );
    }
  }
});

test("defines nine internally consistent complete Office reward cases", async () => {
  const cases = JSON.parse(
    await readFile(new URL("../app/case-data.json", import.meta.url), "utf8"),
  );

  assert.equal(cases.length, 9);
  assert.deepEqual(
    cases.reduce((counts, item) => {
      counts[item.format] = (counts[item.format] ?? 0) + 1;
      return counts;
    }, {}),
    { pptx: 3, docx: 3, xlsx: 3 },
  );

  for (const item of cases) {
    const raw =
      item.scores.aesthetics * 0.4 +
      item.scores.content_accuracy * 0.35 +
      item.scores.communication_effectiveness * 0.25;
    assert.equal(item.overall_raw_score_100, Math.round(raw * 100) / 100);
    assert.equal(
      item.reward_0_1,
      Math.round((item.overall_raw_score_100 / 100) * 10_000) / 10_000,
    );
    assert.equal(item.coverage_0_1, 1);
    assert.equal(item.status, "complete");
    const asset = new URL(`../public${item.image}`, import.meta.url);
    await access(asset);
    assert.ok((await stat(asset)).size > 10_000);
  }
});
