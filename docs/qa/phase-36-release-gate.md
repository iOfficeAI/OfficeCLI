# Phase 36 Release Gate

Phase 36 closes only when corpus, declarative round-trip, visual-threshold,
and provider-matrix gates all agree. This document is the single source of
truth for what must be green and the language that may be claimed.

## Required Artifacts

- `schemas/interfaces/compatibility-corpus.v1.schema.json`
- `schemas/interfaces/expected-capabilities.v1.schema.json`
- `tests/fixtures/common/expected-capabilities.json`
- `tests/fixtures/common/roundtrip-case.v1.schema.json`
- `tests/fixtures/common/roundtrip-cases.json`
- `tests/fixtures/common/visual-thresholds.json`
- `tests/fixtures/common/provider-compatibility.json`
- `tests/fixtures/hwp/manifest.json`
- `tests/fixtures/hwpx/manifest.json`
- `docs/qa/compatibility-corpus.md`
- `docs/qa/visual-diff-thresholds.md`
- `docs/qa/provider-compatibility-matrix.md`
- `docs/qa/phase-36-release-gate.md`

## Required Tests

```text
HwpCompatibilityCorpusTests
HwpRoundTripCorpusTests
HwpVisualDiffThresholdTests
HwpProviderCompatibilityMatrixTests
```

The release gate adds:

```text
HwpCompatibilityCorpusTests.Phase36ReleaseGateRequiresAllCorpusArtifacts
HwpCompatibilityCorpusTests.NoDocxParityLanguageBeforeScorecard
HwpCompatibilityCorpusTests.BlockedOperationsRemainMachineReadable
```

## Acceptance Commands

```text
dotnet build officecli.slnx
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --filter FullyQualifiedName~Hwp --no-build
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --no-build
git diff --check
git ls-files 'src/rhwp-field-bridge/target/*' | wc -l
```

The last command must return `0`.

## Allowed Claim

```text
OfficeCLI tracks HWP/HWPX support with corpus-backed operation evidence,
declarative round-trip cases, visual-threshold policy, and provider
compatibility rows.
```

## Forbidden Claim

```text
HWP/HWPX have DOCX parity.
```

That claim is forbidden until the later parity scorecard is green and
remains enforced by `NoDocxParityLanguageBeforeScorecard` over corpus,
round-trip, visual, provider, and release-gate documents. The phrase may
appear only in a "forbidden claim" or "must not" guard context, never as
an actual capability statement.
