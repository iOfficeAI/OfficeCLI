# Agent onboarding for OfficeCLI

OfficeCLI is a .NET 10 single-binary CLI for AI agents to create, inspect, edit, render, and validate office documents. On `feat/hwpx`, HWP/HWPX work is active and must stay evidence-gated.

## Entry points

| Surface | File(s) | Notes |
| --- | --- | --- |
| Process startup and early-dispatch commands | `src/officecli/Program.cs` | Handles `mcp`, `install`, `skills`, `load_skill`, `config`, help rewrite, logging, auto-install/update. |
| Root command registration | `src/officecli/CommandBuilder.cs` | Registers open/close, watch/mark/view/get/query/set/add/remove/raw/batch/import/create/merge/compare/capabilities/schema/help. |
| Schema-driven help | `src/officecli/CommandBuilder.Help.cs`, `schemas/help/**` | Format aliases and element/property docs live here. |
| HWP bridge help and doctor | `src/officecli/CommandBuilder.Help.Hwp.cs` | Experimental `hwp`/`rhwp` help and readiness probe. |
| Document handlers | `src/officecli/Handlers/**` | DOCX, XLSX, PPTX, HWPX custom, and HWP/rhwp engines. |
| MCP and skills | `src/officecli/Mcp*`, `src/officecli/Core/SkillInstaller.cs`, `skills/**` | Agent integration surfaces. |

## Safe command habits

- Prefer `officecli help`, `officecli help <format>`, and `officecli help <format> <verb> <element>` before guessing properties.
- Use `--json` for machine-readable output.
- For long edit sessions, use `open`/`close`; for multiple mutations, use `batch`.
- For visual QA, use `view <file> html` or `view <file> screenshot` when supported.
- For HWP work, run `officecli hwp doctor --json` and `officecli capabilities --json` before mutation.

## HWP/HWPX claim rules

Allowed current claim from `docs/qa/phase-36-release-gate.md`:

> OfficeCLI tracks HWP/HWPX support with corpus-backed operation evidence, declarative round-trip cases, visual-threshold policy, and provider compatibility rows.

Do not claim HWP/HWPX have DOCX parity. HWP/HWPX operation support must be tied to:

- `officecli capabilities --json`;
- `schemas/help/hwp` or `schemas/help/hwpx`;
- fixture manifests under `tests/fixtures/hwp`, `tests/fixtures/hwpx`, and `tests/fixtures/common`;
- tests under `tests/OfficeCli.Tests/Hwp` or `tests/OfficeCli.Tests/Hwpx`.

## Documentation conventions

- Keep `structure/*.md` under 300 lines.
- Ground statements in files, docs, fixtures, or commits.
- Update `README.md` for public-facing capability changes and `BRANCH_STRATEGY.md` for branch workflow changes.
- Doc-only changes must not edit `.cs`, `.ts`, `.js`, `.py`, or other source files unless explicitly requested.
