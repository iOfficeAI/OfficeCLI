# PPT Subquestion Scoring V2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans and test-driven development.

**Goal:** Score 30 real slides on 40 direct visual/communication subquestions and expose blank human comparison inputs.

**Architecture:** A deterministic manifest builder preserves V1 cases and adds unique annotated cases. A V2 Python runner makes two structured multimodal calls per slide, saves resumable evidence, and emits a static artifact consumed by the existing React page.

**Tech Stack:** Python, Pydantic, OpenAI Responses-compatible API, React, TypeScript, Vinext, Node test runner, Playwright.

---

### Task 1: V2 Contract and Manifest

- Add a failing artifact test for 30 cases, grade distribution `9/9/9/3`, 60
  rows per case, direct subquestion scores, abstentions, and provenance.
- Implement deterministic manifest selection and verify existing V1 IDs remain.

### Task 2: Real V2 Scoring

- Implement four subquestions under each of the ten scoreable parent criteria.
- Run two structured model calls per slide with strict expected-ID validation,
  atomic resume, and model/rubric-bound progress metadata.
- Compute transparent parent means, token totals, and diagnostic rank metrics.

### Task 3: Subquestion Workbench

- Add failing SSR/browser expectations for 20 visible subquestion rows, new
  progress denominator, blank V2 human state, and direct evidence.
- Replace criterion rows with subquestion rows while preserving the four-region
  desktop layout and stacked mobile layout.

### Task 4: Verify and Publish

- Run the full build, tests, lint, desktop/mobile browser checks, and public
  Cloudflare checks.
- Keep the current Quick Tunnel URL active and verify it serves the V2 artifact.
