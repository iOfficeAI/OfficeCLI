# File/function map

## Top-level layout

| Path | Purpose |
| --- | --- |
| `src/officecli` | Main CLI executable, command registration, document handlers, MCP/skill integration, resources. |
| `src/rhwp-officecli-bridge` | C# sidecar bridge executable/project for experimental rhwp integration. |
| `src/rhwp-field-bridge` | Rust rhwp API sidecar source for field/text/table operations. Build outputs are intentionally ignored. |
| `tests/OfficeCli.Tests` | xUnit test suite for handlers, HWP/HWPX gates, safe-save, schemas, and regression coverage. |
| `tests/fixtures` | HWP/HWPX corpus manifests, common expected-capability/provider/round-trip/visual fixtures, sample documents. |
| `schemas/help` | Embedded schema-driven help by format and element. Formats include `docx`, `xlsx`, `pptx`, `hwp`, `hwpx`. |
| `schemas/interfaces` | JSON schemas for capability, provider, sidecar, validation, diff, edit, and safe-save contracts. |
| `docs/qa` | Phase 36 compatibility corpus, visual thresholds, provider matrix, release-gate documentation. |
| `docs/providers` | Provider boundary docs, currently rhwp sidecar contract. |
| `docs/safety` | Safe-save policy for HWP/HWPX mutations. |
| `skills` | Embedded agent skills including officecli base/specialized skills and `officecli-hwpx`. |
| `examples` / `assets` / `styles` | Sample documents, visual assets, and design style packs. |

## Main command files

| File | Responsibility |
| --- | --- |
| `Program.cs` | Startup, help rewrite, early commands (`mcp`, `install`, `skills`, `load_skill`, `config`), logging/update hooks. |
| `CommandBuilder.cs` | Root command and resident open/close setup; registers partial command builders. |
| `CommandBuilder.View.cs` / `.View.Help.cs` | `view` modes: text, annotated, outline, stats, issues, html, svg, screenshot, HWP/HWPX field modes. |
| `CommandBuilder.GetQuery.cs` | `get` and `query` path/selector reads. |
| `CommandBuilder.Set.cs` | Generic set, selected pseudo-path, document protection, HWP/HWPX mutation dispatch. |
| `CommandBuilder.Set.Hwp*.cs` | HWP/HWPX field/text/table mutation helpers and safe in-place policy errors. |
| `CommandBuilder.Add.cs` | `add`, `remove`, `move`, `swap`, and merge/template helpers. |
| `CommandBuilder.Raw.cs` | `raw`, `raw-set`, `add-part` OpenXML fallback operations. |
| `CommandBuilder.Batch.cs` | JSON batch execution in one open/save cycle. |
| `CommandBuilder.Import.cs` | CSV/TSV import plus `create`/`new`. |
| `CommandBuilder.Capabilities.cs` | Machine-readable HWP/HWPX capability report. |
| `CommandBuilder.Schema.cs` | Schema export/validation helpers for help/interface files. |
| `CommandBuilder.Help*.cs` | Schema-driven help and HWP/rhwp help/doctor. |
| `CommandBuilder.IntegrationStubs.cs` | Help-visible stubs for early-dispatch commands. |

## Handler map

| Handler area | Formats | Notes |
| --- | --- | --- |
| `Handlers/Word*`, `Handlers/Word/**` | `.docx` | OpenXML Word create/read/edit/render/query. |
| `Handlers/Excel*`, `Handlers/Excel/**` | `.xlsx` | Workbook/sheet/cell/formula/chart/pivot/table/data operations. |
| `Handlers/PowerPoint*`, `Handlers/Pptx/**` | `.pptx` | Slide/shape/media/chart/theme/morph/HTML/SVG operations. |
| `Handlers/Hwpx/**` | `.hwpx` | Custom ZIP/XML HWPX handler, validation, view, path, raw, set/diff/import helpers. |
| `Handlers/Hwp/**` | `.hwp`, `.hwpx` via provider | Capability report, engine selection, custom HWPX engine, rhwp bridge engine, typed errors. |
| `Handlers/DocumentHandlerFactory.cs` | all supported extensions | Opens the correct handler by extension. |

## Embedded resources

`src/officecli/officecli.csproj` embeds preview CSS/JS, `Resources/base.hwpx`, chart resources, all `skills/**`, root `SKILL.md`, and `schemas/help/**/*.json` for single-file distribution.
