# OfficeCLI structure index

`structure/` is the agent-facing source-of-truth layer for the current `feat/hwpx` branch. It connects the public README, CLI command surface, file layout, provider contracts, and Phase 36 HWP/HWPX evidence gates.

## Document map

| Doc | Scope | Primary evidence |
| --- | --- | --- |
| [AGENTS.md](AGENTS.md) | Agent onboarding, safe workflows, conventions | `Program.cs`, `CommandBuilder*.cs`, `SKILL.md` |
| [01-file-function-map.md](01-file-function-map.md) | File and directory architecture | `src/`, `schemas/`, `tests/`, `docs/` |
| [02-command-reference.md](02-command-reference.md) | CLI commands and help paths | `Program.cs`, `CommandBuilder.cs`, `CommandBuilder.Help*.cs` |
| [03-format-support.md](03-format-support.md) | DOCX/XLSX/PPTX/HWP/HWPX support matrix | handlers, schemas, Phase 36 QA docs |
| [04-providers.md](04-providers.md) | Provider/backend matrix and boundaries | `Handlers/Hwp`, `docs/providers`, `docs/qa/provider-compatibility-matrix.md` |
| [05-test-infra.md](05-test-infra.md) | Test layout and smoke gates | `tests/OfficeCli.Tests`, `tests/fixtures`, `docs/qa/phase-36-release-gate.md` |

## Read order

1. Start with this index.
2. Read [AGENTS.md](AGENTS.md) before editing docs or using the CLI as an agent.
3. Use [01-file-function-map.md](01-file-function-map.md) to locate implementation files.
4. Use [02-command-reference.md](02-command-reference.md) for command syntax.
5. Use [03-format-support.md](03-format-support.md) and [04-providers.md](04-providers.md) before making any HWP/HWPX capability claim.
6. Use [05-test-infra.md](05-test-infra.md) before running or changing gates.

## Current branch focus

Recent commits on `feat/hwpx` concentrate on HWP/HWPX evidence, not blanket parity claims:

- Phase 36.2-36.6 compatibility corpus, round-trip catalog, visual thresholds, provider matrix, and release gate.
- Experimental HWP/rhwp bridge help, doctor, capabilities, field/text/table operations, and safe in-place text replacement policy.
- HWPX remains evidence-gated: custom provider default, rhwp bridge opt-in, and operation claims tied to `officecli capabilities --json` plus fixture evidence.

## Sync checklist

- Update command docs when `Program.cs` or any `CommandBuilder*.cs` command registration changes.
- Update format/provider docs when `schemas/help/{docx,xlsx,pptx,hwp,hwpx}` or `Handlers/Hwp*` changes.
- Update test docs when `tests/OfficeCli.Tests/Hwp`, `tests/fixtures/common`, or `docs/qa` gates change.
- Do not claim HWP/HWPX DOCX parity until a later parity scorecard explicitly allows it.
