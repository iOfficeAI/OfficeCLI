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

