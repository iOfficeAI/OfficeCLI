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

## Round-Trip Cases

Phase 36.3 adds an operation-level round-trip catalog at:

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
execution is opt-in and gated on `OFFICECLI_REAL_RHWP_BIN`.

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

## Next Gates

The corpus is intentionally small in this first slice. Later Phase 36 patches
should add fixture classes for:

```text
footnotes/endnotes
large documents
```

