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

## Next Gates

The corpus is intentionally small in this first slice. Later Phase 36 patches
should add fixture classes for:

```text
multi-section
merged-cell tables
nested tables
pictures/BinData
headers/footers
footnotes/endnotes
equations
large documents
Unicode edge cases
malformed HWPX packages
```

