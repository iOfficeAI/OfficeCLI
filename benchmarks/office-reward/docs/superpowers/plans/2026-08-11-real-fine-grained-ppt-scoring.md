# Real Fine-Grained PPT Scoring Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Run criterion-level model evaluation on 12 real PPT benchmark slides and present the results beside blank persistent human scoring controls.

**Architecture:** A Python experiment script reads the benchmark manifest, submits two schema-constrained multimodal requests per slide, and writes static JSON plus local images. The React workbench consumes that artifact and never performs live model calls.

**Tech Stack:** Python, Responses-compatible API, JSON, React 19, TypeScript, Vinext, Node test runner, Playwright.

---

### Task 1: Experiment Contract and Runner

**Files:**
- Create: `tests/fine-grained-results.test.mjs`
- Create: `scripts/score_fine_grained_ppt.py`
- Create: `app/rubric/fine-grained-types.ts`
- Generate: `app/rubric/fine-grained-results.json`

- [ ] Write a failing test requiring 12 cases, 15 criteria per case, ten real scores, five abstentions, evidence, and provenance.
- [ ] Run `node --test tests/fine-grained-results.test.mjs` and confirm the missing-artifact failure.
- [ ] Implement manifest validation, local image copying, two model calls per slide, schema validation, resume, and atomic output.
- [ ] Run the real experiment against the local endpoint and rerun the contract test.

### Task 2: Four-Region Workbench

**Files:**
- Modify: `app/rubric/rubric-data.ts`
- Modify: `app/rubric/rubric-workbench.tsx`
- Modify: `app/rubric/page.tsx`
- Modify: `app/globals.css`
- Modify: `tests/rendered-html.test.mjs`
- Modify: `tests/case-gallery.browser.test.mjs`

- [ ] Add failing assertions for rubric, slide, real AI result, human score, provenance, abstention, and removal of `aiSubscore`.
- [ ] Run `npm test` and confirm the expected old-layout failure.
- [ ] Render the 12 PPT cases from the experiment artifact with dimension tabs and sticky 16:9 images.
- [ ] Keep human scores blank, local, case-specific, reload-safe, and exportable with model provenance.
- [ ] Add bounded four-track desktop CSS and a no-overflow stacked mobile layout.
- [ ] Run `npm test && npm run lint` and fix failures.

### Task 3: Delivery

- [ ] Run the final build and focused contract/browser tests.
- [ ] Start and retain the healthy local development server.
- [ ] Publish the exact validated build through Sites and verify deployment success.
