# Office Reward Case Gallery Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add nine complete, filterable Office reward cases with real local assets and detailed score dialogs to the existing progress site.

**Architecture:** Store deterministic case fixtures in JSON and expose them through a typed module. Render the comparison grid in a focused client component while keeping the rest of the page server-rendered. Generate all media locally and continue serving the site through the existing Cloudflare Quick Tunnel.

**Tech Stack:** React 19, Next/Vinext, TypeScript, CSS, Node test runner, OfficeCLI, Cloudflare Quick Tunnel.

---

## File Map

- `app/case-data.json`: nine deterministic complete case fixtures.
- `app/case-data.ts`: case types, score derivation, and exported fixtures.
- `app/case-gallery.tsx`: filter controls, dense grid, and accessible dialog.
- `app/page.tsx`: navigation and gallery placement.
- `app/globals.css`: gallery, score bars, dialog, and responsive styling.
- `public/cases/*.png`: nine local case images.
- `tests/rendered-html.test.mjs`: fixture, asset, SSR, and score-consistency tests.
- `.gitignore`: ignore `.superpowers/` visual-companion state.

## Task 1: Case Fixture Contract

**Files:**
- Create: `app/case-data.json`
- Create: `app/case-data.ts`
- Modify: `tests/rendered-html.test.mjs`

- [ ] **Step 1: Write the failing fixture test**

Add a test that loads `app/case-data.json` and asserts:

```js
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
  assert.equal(item.reward_0_1, item.overall_raw_score_100 / 100);
  assert.equal(item.coverage_0_1, 1);
  assert.equal(item.status, "complete");
}
```

- [ ] **Step 2: Run the test and verify RED**

Run: `npm test`

Expected: FAIL because `app/case-data.json` does not exist.

- [ ] **Step 3: Add fixture data and typed exports**

Define:

```ts
export type OfficeFormat = "pptx" | "docx" | "xlsx";

export type RewardCase = {
  id: string;
  format: OfficeFormat;
  title: string;
  subtitle: string;
  image: string;
  unitType: "slide" | "page" | "section" | "sheet";
  unitCount: number;
  scores: {
    aesthetics: number;
    content_accuracy: number;
    communication_effectiveness: number;
  };
  overall_raw_score_100: number;
  reward_0_1: number;
  coverage_0_1: 1;
  status: "complete";
  evidence: string[];
  issueSummary: string;
};
```

Export the JSON as `rewardCases: RewardCase[]`.

- [ ] **Step 4: Run the fixture test**

Run: `npm test`

Expected: PASS for fixture count and score consistency.

- [ ] **Step 5: Commit**

```bash
git add app/case-data.json app/case-data.ts tests/rendered-html.test.mjs
git commit -m "feat: add Office reward case fixtures"
```

## Task 2: Real Case Assets

**Files:**
- Create: `public/cases/ppt-executive-review.png`
- Create: `public/cases/ppt-research-results.png`
- Create: `public/cases/ppt-product-roadmap.png`
- Create: `public/cases/word-quarterly-report.png`
- Create: `public/cases/word-table-analysis.png`
- Create: `public/cases/word-policy-brief.png`
- Create: `public/cases/excel-revenue-model.png`
- Create: `public/cases/excel-chart-dashboard.png`
- Create: `public/cases/excel-operations-sheet.png`
- Modify: `tests/rendered-html.test.mjs`

- [ ] **Step 1: Add failing asset assertions**

For every fixture:

```js
const asset = new URL(`../public${item.image}`, import.meta.url);
await access(asset);
const stats = await stat(asset);
assert.ok(stats.size > 10_000);
```

- [ ] **Step 2: Run the test and verify RED**

Run: `npm test`

Expected: FAIL because `public/cases/` assets are absent.

- [ ] **Step 3: Generate and stage real assets**

- Copy three distinct tracked slide images for PPT cases.
- Create three DOCX files with OfficeCLI and render one screenshot per file.
- Create three XLSX files with OfficeCLI and render one screenshot per file.
- Store only the final PNG files under `public/cases/`.

- [ ] **Step 4: Verify images**

Run:

```bash
identify public/cases/*.png
npm test
```

Expected: nine readable, nonblank PNG files and passing asset tests.

- [ ] **Step 5: Commit**

```bash
git add public/cases tests/rendered-html.test.mjs
git commit -m "feat: add Office reward case media"
```

## Task 3: Filterable Dense Score Grid

**Files:**
- Create: `app/case-gallery.tsx`
- Modify: `app/page.tsx`
- Modify: `app/globals.css`
- Modify: `tests/rendered-html.test.mjs`

- [ ] **Step 1: Write failing UI assertions**

Assert the server output contains:

```js
assert.match(html, /CASE GALLERY/);
assert.match(html, /全部/);
assert.match(html, /PowerPoint/);
assert.match(html, /Word/);
assert.match(html, /Excel/);
assert.match(html, /Executive Review/);
assert.match(html, /Quarterly Report/);
assert.match(html, /Revenue Model/);
```

Also assert source files include `role="dialog"`, `aria-modal="true"`, and the
three score labels.

- [ ] **Step 2: Run the test and verify RED**

Run: `npm test`

Expected: FAIL because the gallery component is absent.

- [ ] **Step 3: Implement the client component**

Implement:

```tsx
"use client";

const filters = [
  ["all", "全部"],
  ["pptx", "PPT"],
  ["docx", "Word"],
  ["xlsx", "Excel"],
] as const;
```

Use state for the active filter and selected case. Render stable three-column
cards with thumbnail, score bars, reward, coverage, unit count, and complete
status. Render an accessible modal dialog for the selected case and close it
with the close button, backdrop click, or Escape.

- [ ] **Step 4: Integrate and style**

- add `Cases` to navigation;
- insert `<CaseGallery />` after the format section;
- use an 8px-or-less card radius;
- use fixed image aspect ratios and score-bar tracks;
- switch to one column below 900px;
- prevent dialog and card text overflow.

- [ ] **Step 5: Run tests**

Run:

```bash
npm test
npm run lint
```

Expected: build, SSR tests, fixture tests, and lint pass.

- [ ] **Step 6: Commit**

```bash
git add app/case-gallery.tsx app/page.tsx app/globals.css tests/rendered-html.test.mjs
git commit -m "feat: add filterable Office case gallery"
```

## Task 4: Cloudflare Verification

**Files:**
- Modify: `.gitignore`

- [ ] **Step 1: Ignore visual-companion state**

Add:

```gitignore
/.superpowers/
```

- [ ] **Step 2: Run complete site verification**

Run:

```bash
npm test
npm run lint
git diff --check
curl -H "Host: exchange-bush-availability-split.trycloudflare.com" \
  http://localhost:3000/
```

Expected: tests and lint pass; local Host request returns the gallery.

- [ ] **Step 3: Verify the Quick Tunnel**

Fetch:

```text
https://exchange-bush-availability-split.trycloudflare.com
```

Expected: HTTP 200 with `CASE GALLERY` and all nine cases.

- [ ] **Step 4: Commit**

```bash
git add .gitignore
git commit -m "chore: finalize Office case gallery delivery"
```
