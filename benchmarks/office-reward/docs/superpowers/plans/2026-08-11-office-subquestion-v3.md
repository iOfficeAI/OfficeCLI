# Office Subquestion Scoring V3 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans and test-driven development.

**Goal:** Add 12 Word and 12 Excel cases with format-specific direct subquestion scoring to the existing 30-case PPT experiment.

**Architecture:** A V3 runner reuses the immutable PPT V2 artifact, renders fixed Word and Excel sources through OfficeCLI, scores two dimensions per new unit, and emits one multi-format artifact consumed by the current workbench.

**Tech Stack:** Python, OfficeCLI, Pillow, Pydantic, OpenAI Responses-compatible API, React, TypeScript, Vinext, Playwright.

---

### Task 1: Multi-format contract
- Add a failing test for 54 cases and `30/12/12` format counts.
- Require 2,160 direct scores, 1,080 abstentions, source hashes, and parent means.

### Task 2: Render and score Word/Excel
- Define fixed 12-document source manifests per format.
- Render one bounded unit per document and reject missing or blank images.
- Define format-specific 5x4 Aesthetics and Communication subquestions.
- Run 48 resumable model calls and combine them with V2 PPT evidence.

### Task 3: Multi-format UI
- Add failing UI checks for PPT/Word/Excel tabs and case counts.
- Select the rubric from the current case format while retaining five parent
  scores, 20 direct rows, and blank V3 human inputs.

### Task 4: Verify and publish
- Run build, all tests, lint, public browser checks, and mobile no-overflow.
- Verify the existing Cloudflare Quick Tunnel serves all three formats.
