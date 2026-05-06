# Format support matrix

This file describes current support without over-claiming HWP/HWPX parity. For HWP/HWPX, operation truth comes from `officecli capabilities --json`, `docs/qa/*`, and fixture manifests.

## Public format matrix

| Format | Create | Read/inspect | Modify | Render/view | Status |
| --- | --- | --- | --- | --- | --- |
| `.docx` | Yes | Yes | Yes | HTML, screenshot; native page render on Windows+Word when requested | Stable OpenXML handler. |
| `.xlsx` | Yes | Yes | Yes | HTML, screenshot | Stable OpenXML handler with Excel-specific cells, formulas, charts, pivots, tables, slicers, validations. |
| `.pptx` | Yes | Yes | Yes | HTML, SVG/screenshot flows | Stable OpenXML handler with slides, shapes, charts, media, morph/transition/3D/theme features. |
| `.hwpx` | Yes (`Resources/base.hwpx`) | Yes through custom ZIP/XML handler; rhwp opt-in for selected read/render/field paths | Experimental, operation-gated; do not advertise broad writing | HTML/custom views; rhwp SVG opt-in when enabled | Active branch focus; custom provider default. |
| `.hwp` | No blank creator | Experimental read/render/mutation through rhwp bridge | Experimental output-first and safe in-place text replacement only where gates pass | rhwp SVG bridge | Active branch focus; rhwp bridge default provider for binary HWP. |

## HWP/HWPX Phase 36 status

| Phase/area | Evidence | Current claim boundary |
| --- | --- | --- |
| Compatibility manifests | `tests/fixtures/hwp/manifest.json`, `tests/fixtures/hwpx/manifest.json`, `tests/fixtures/common/expected-capabilities.json` | Fixture-backed operation ledger, not format-wide fidelity. |
| Corpus schema validation | `schemas/interfaces/compatibility-corpus.v1.schema.json`, `expected-capabilities.v1.schema.json` | Manifests must stay machine-readable and schema-valid. |
| Round-trip catalog | `tests/fixtures/common/roundtrip-cases.json`, `roundtrip-case.v1.schema.json` | Declarative invariants: source unchanged, output created, provider readback, semantic delta, typed blocked errors. |
| Visual thresholds | `tests/fixtures/common/visual-thresholds.json`, `docs/qa/visual-diff-thresholds.md` | Visual-validated operations need render evidence; renderer status remains deferred unless evidence says otherwise. |
| Provider matrix | `tests/fixtures/common/provider-compatibility.json`, `docs/qa/provider-compatibility-matrix.md` | HWPX custom default, HWP rhwp-bridge default, Hancom optional lane. |
| Release gate | `docs/qa/phase-36-release-gate.md` | Allows evidence-tracking claim; forbids DOCX parity claim. |

## HWP operation coverage from corpus docs

Binary HWP operation-level coverage currently listed in `docs/qa/compatibility-corpus.md`:

- `read_text`
- `render_svg`
- `list_fields`
- `read_field`
- `fill_field`
- `replace_text`
- `set_table_cell`

HWPX coverage is provider-specific:

- `custom` remains default;
- `rhwp-bridge` is opt-in for read/render/text replacement paths;
- `set_table_cell` remains blocked until package and Hancom compatibility gates pass.

## Current HWPX wording rule

`docs/hwpx-current-operation-inventory.md` requires public wording to stay evidence-gated:

> OfficeCLI advertises only the HWPX operations listed as roundtrip-verified in officecli capabilities --json.

Do not say "OfficeCLI supports HWPX writing" unless the operation is roundtrip-verified with evidence.

## Recent non-HWP capability commits

Recent upstream/main commits reflected in this branch include:

- DOCX section/page rendering improvements: inline `sectPr`, section breaks in dump, `view --render`, and `stats --page-count`.
- Dump serializer improvements for body paragraphs, mixed body content, tables, styles, sections, comments, notes, headers/footers, numbering, settings, theme, pictures, and charts.
- PPTX/XLSX improvements: screenshot mode, `[last()]` selector support, chartEx compatibility fixes, and skill docs cleanup.
- Agent integration: `load_skill` CLI/MCP support and embedded skills setup changes.
