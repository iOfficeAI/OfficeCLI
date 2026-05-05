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
```

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
```

These run on every CI build and do not depend on a renderer being available.
