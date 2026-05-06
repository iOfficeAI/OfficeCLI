# Visual Diff Thresholds (Phase 36.4)

OfficeCLI tracks HWP/HWPX visual evidence policy without committing to a
specific pixel-diff implementation in CI. This document defines the contract
that any visual claim must respect.

## Hard Fail Conditions

A visual claim is rejected outright when any of the following occur:

```text
page-count-mismatch
missing-expected-svg-page
missing-render-evidence-for-visual-validated-operation
body-marker-in-fixed-layout-exam
```

`body-marker-in-fixed-layout-exam` applies to KICE-style exam sheets and
similar fixed-layout documents. When a document is detected as a fixed-layout
exam sheet, proof text such as `[CU TEMPLATE EDIT ...]`, `VISUAL QA`, or
`edited via Hancom Office HWP UI` must not be inserted into the visible body.
Those markers change the page flow and are hard failures even if the file opens
and save/readback succeeds.

## Thresholded Fail Conditions

Operation-level layout drift is allowed within declared bounds. The current
catalog lives at:

```text
tests/fixtures/common/visual-thresholds.json
```

Allowed metrics:

```text
layout-drift-fraction
blank-page-fraction
svg-hash-equality
```

Defaults:

- text-only mutations may drift up to **2%** of layout area
- any unexpected blank page is a fail
- when a row declares exact SVG hash evidence, any drift is a fail
- fixed-layout exam sheets allow **0%** visible body layout drift unless the
  requested edit explicitly changes exam content

## Fixed-Layout Exam Sheets

Fixed-layout exam sheets are identified by structural signals such as
`NEWSPAPER` two-column layout, exam-title text, and question numbering. Visual
QA for these files requires before/after screenshots and a manual visual review
until a stable renderer is available in CI.

For these documents, QA proof must be stored outside the visible question body:
use screenshots, sidecar evidence, logs, or non-visible metadata. Do not insert
review markers into the first column, question body, answer choices, header
tables, or floating title/page-number objects.

## Visual Validated Operations

Only operations explicitly listed in `visualValidatedOperations` may make
visual claims. As of Phase 36.4 that list is `[render_svg]`. `replace_text`,
`fill_field`, and `set_table_cell` may declare *threshold tolerances* but
cannot declare a visual claim without a render evidence link in
`expected-capabilities.json`.

## Renderer Status

```text
rendererStatus = deferred
```

OfficeCLI CI does not yet ship a stable HWP/HWPX renderer. Until it does, the
threshold contract is enforced declaratively: visual evidence must be linked,
and any operation marked visual-validated without render evidence fails CI.

## Tests

The contract is enforced by:

```text
tests/OfficeCli.Tests/Hwp/HwpVisualDiffThresholdTests.cs
  - PageCountMismatchIsHardFail
  - MissingRenderEvidenceFailsVisualClaim
  - TextOnlyMutationUsesDeclaredThreshold
  - FixedLayoutExamBodyMarkersAreHardFail
  - FixedLayoutExamRuleRejectsAdHocBodyProofMarkers
```

These run on every CI build and do not depend on a renderer being available.
