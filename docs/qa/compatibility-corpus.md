# HWP/HWPX Compatibility Corpus

This corpus is the evidence ledger for HWP/HWPX parity work. It does not claim
format-wide fidelity. It records which concrete fixtures prove which concrete
operations.

## Manifests

```text
tests/fixtures/hwp/manifest.json
tests/fixtures/hwpx/manifest.json
tests/fixtures/common/expected-capabilities.json
```

Each fixture entry includes:

- stable fixture id;
- repository-relative file path;
- SHA-256;
- byte size;
- document classes;
- verified operations;
- evidence files;
- blocked operations when relevant.

## Claim Policy

An HWP/HWPX operation can move toward DOCX-like agent usability only when the
same operation is present in all of these places:

```text
capabilities --json
schema/help JSON
fixture manifest
expected-capabilities.json
tests or golden evidence
```

Unsupported or unverified operations must fail closed with a typed reason such
as `roundtrip_unverified`, `bridge_missing`, or `unsupported_operation`.

## Current Coverage

Binary HWP coverage is operation-level:

```text
read_text
render_svg
list_fields
read_field
fill_field
replace_text
set_table_cell
```

HWPX coverage is provider-specific:

```text
custom provider remains default
rhwp provider is opt-in for read/render/text replacement paths
set_table_cell remains blocked until package and Hancom compatibility gates pass
```

## Provider Compatibility Matrix

Phase 36.5 adds the cross-provider matrix at:

```text
docs/qa/provider-compatibility-matrix.md
tests/fixtures/common/provider-compatibility.json
```

HWPX `custom` remains the default provider; rhwp-bridge stays opt-in only and
must not be promoted to default until evidence parity is reached. HWP defaults
to `rhwp-bridge`. The matrix covers every expected-capability operation for
both `custom` and `rhwp-bridge`, with blocked provider paths carrying typed
reasons such as `unsupported_engine`, `binary_hwp_mutation_forbidden`, and
`binary_hwp_write_forbidden`. Hancom is `optional` on every row; it can support
a future status promotion but must not be required by normal CI.

## Visual Diff Thresholds

Phase 36.4 adds the visual evidence policy at:

```text
docs/qa/visual-diff-thresholds.md
tests/fixtures/common/visual-thresholds.json
```

Hard fails (page-count mismatch, missing SVG page, missing render evidence
for a visual-validated operation, and body proof markers in fixed-layout exam
sheets) cannot be tolerated. Thresholded fails (text-only layout drift,
unexpected blank render, exact SVG hash mismatch) have declared bounds. KICE
style fixed-layout exam sheets use a stricter rule: proof markers must stay out
of the visible question body and visible layout drift is `0%` unless the
requested edit explicitly changes exam content. As of Phase 36.4 only
`render_svg` is a visual-validated operation; mutation operations may declare
drift tolerance but cannot claim visual validation without a linked render
evidence file. The renderer status remains `deferred` until OfficeCLI ships a
stable in-CI renderer.

## Round-Trip Cases

Phase 36.3 adds an operation-level declarative round-trip catalog at:

```text
tests/fixtures/common/roundtrip-cases.json
tests/fixtures/common/roundtrip-case.v1.schema.json
```

Each case declares a fixtureId, operation, provider, outputMode, args, and the
required checks: `source-unchanged`, `output-created`, `provider-readback`,
`semantic-delta`, `typed-error-if-blocked`. Mutation cases must include
`source-unchanged` and must not run with `outputMode = in-place`. Blocked cases
must include `typed-error-if-blocked` and a typed `expected.error.code`.

Normal CI enforces declarative invariants over the catalog. Real rhwp-backed
execution is opt-in and gated on `OFFICECLI_REAL_RHWP_BIN`; Phase 36 does not
claim a full executor with semantic output comparison in normal CI.

## Fixture Class Coverage

Phase 36.2 records required fixture classes in each manifest under
`fixtureClassCoverage`. Each class must declare a state:

```text
verified         → small in-repo fixture proves the class
blocked          → typed reason explains why the class is not verified
external-manual  → samples are tracked outside the repo
```

Required classes:

```text
multi-section
merged-cell-tables
nested-tables
pictures-bindata
headers-footers
equations
unicode-edge-cases
malformed-hwpx-package
```

`malformed-hwpx-package` is HWPX-only and must remain `blocked` with reason
`fixture_validation_failed`. External-manual entries must not declare
`verifiedOperations` and do not contribute to capability evidence.

## Phase 36 Release Gate

Phase 36 closes when all corpus, declarative round-trip, visual-threshold, and
provider-matrix gates agree. The single source of truth lives at:

```text
docs/qa/phase-36-release-gate.md
```

Allowed claim:

```text
OfficeCLI tracks HWP/HWPX support with corpus-backed operation evidence,
declarative round-trip cases, visual-threshold policy, and provider
compatibility rows.
```

The forbidden claim ("HWP/HWPX have DOCX parity") is enforced by
`HwpCompatibilityCorpusTests.NoDocxParityLanguageBeforeScorecard` and remains
blocked until the later parity scorecard is green.

## Next Gates

The corpus is intentionally small in this first slice. Later Phase 36 patches
should add fixture classes for:

```text
footnotes/endnotes
large documents
```
