# Test infrastructure

OfficeCLI uses .NET/xUnit tests plus schema/fixture documents for HWP/HWPX evidence gates.

## Projects

| Project/file | Purpose |
| --- | --- |
| `officecli.slnx` | Solution containing CLI and tests. |
| `src/officecli/officecli.csproj` | Main .NET 10 executable; embeds resources, skills, and help schemas. |
| `src/rhwp-officecli-bridge/rhwp-officecli-bridge.csproj` | Experimental C# bridge project. |
| `tests/OfficeCli.Tests/OfficeCli.Tests.csproj` | xUnit test project with `JsonSchema.Net`, coverlet, and project reference to `officecli`. |

## Test layout

| Path | Coverage |
| --- | --- |
| `tests/OfficeCli.Tests/Hwp` | HWP/rhwp bridge, capabilities, safe-save, provider matrix, corpus, release gate, visual thresholds. |
| `tests/OfficeCli.Tests/Hwpx` | HWPX handler, package validator, corpus tests. |
| `tests/OfficeCli.Tests/CjkHelperTests.cs` | CJK helper regression coverage. |
| `tests/fixtures/common` | Expected capabilities, provider compatibility, round-trip cases, visual thresholds. |
| `tests/fixtures/hwp` | HWP manifest and HWP/rhwp fixture documents. |
| `tests/fixtures/hwpx` | HWPX manifest. |
| `tests/golden` | Golden outputs where present. |

## Phase 36 required tests

`docs/qa/phase-36-release-gate.md` lists these HWP gate classes:

```text
HwpCompatibilityCorpusTests
HwpRoundTripCorpusTests
HwpVisualDiffThresholdTests
HwpProviderCompatibilityMatrixTests
```

The release gate also requires:

```text
HwpCompatibilityCorpusTests.Phase36ReleaseGateRequiresAllCorpusArtifacts
HwpCompatibilityCorpusTests.NoDocxParityLanguageBeforeScorecard
HwpCompatibilityCorpusTests.BlockedOperationsRemainMachineReadable
```

## Recommended gates

For a normal code change:

```bash
dotnet build officecli.slnx
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --no-build
```

For HWP/HWPX evidence work:

```bash
dotnet build officecli.slnx
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --filter FullyQualifiedName~Hwp --no-build
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --no-build
git diff --check
git ls-files 'src/rhwp-field-bridge/target/*' | wc -l
```

The last command must return `0` so Rust bridge build outputs are not tracked.

## Optional real-rhwp coverage

Some real bridge smoke coverage is optional and depends on external binaries/env vars such as `OFFICECLI_REAL_RHWP_BIN` or the HWP bridge variables documented in [04-providers.md](04-providers.md). Normal CI must not require Hancom or local rhwp installs.

## Doc-only smoke

For documentation-only changes, run at minimum:

```bash
git diff --check
```

If .NET tooling is available and fast enough, run the build/test commands above to ensure embedded documentation references have not drifted.
