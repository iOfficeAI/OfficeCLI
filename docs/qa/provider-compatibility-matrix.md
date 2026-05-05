# Provider Compatibility Matrix (Phase 36.5)

OfficeCLI tracks HWP/HWPX provider behaviour without promoting a provider
ahead of the evidence the corpus carries. The authoritative catalog lives at:

```text
tests/fixtures/common/provider-compatibility.json
```

## Schema

Each row has:

```text
format
operation
provider
status
defaultProvider
evidence
blockedReason
hancomLane
```

Allowed `status` values:

```text
unsupported
experimental
fixture-backed
roundtrip-verified
external-manual
```

## Defaults

- HWPX → `custom` is the **default provider**. `rhwp-bridge` is opt-in only and
  must not be promoted to default until evidence parity is reached on every
  HWPX operation it claims.
- HWP → `rhwp-bridge` is the default provider. `custom` is unsupported for
  binary HWP read/render paths under `unsupported_engine` and for mutations
  (`fill_field`, `replace_text`, `set_table_cell`, `save_as_hwp`) under the
  binary-write block.
- Hancom — every row carries `hancomLane = optional`. Hancom evidence may
  *support* a future status promotion but never replaces corpus or round-trip
  evidence and must not be required by normal CI.

## Coverage Contract

The matrix must include one row for each
`expected-capabilities.json` format/operation/provider tuple across:

```text
custom
rhwp-bridge
```

Exactly one provider row per format/operation may set `defaultProvider = true`,
and that provider must match the format's `defaultEngine`.

## Blocked Provider Rows

Any row with `status = unsupported` must carry a typed `blockedReason` drawn
from the typed reason enum in
`src/officecli/Handlers/Hwp/HwpCapabilityReport.cs`:

```text
unsupported_format
unsupported_operation
unsupported_engine
roundtrip_unverified
binary_hwp_mutation_forbidden
binary_hwp_write_forbidden
fixture_validation_failed
capability_schema_invalid
bridge_*
```

## Tests

Enforcement lives in:

```text
tests/OfficeCli.Tests/Hwp/HwpProviderCompatibilityMatrixTests.cs
  - HwpxCustomRemainsDefault
  - RhwpPromotionRequiresEvidenceParity
  - HancomLaneIsOptionalNotCiRequired
  - BlockedProviderRowsHaveTypedReasons
  - RowsAreUnique
  - MatrixCoversExpectedCapabilityProviderPairs
```

These run on every CI build and do not require any external Hancom
installation.
