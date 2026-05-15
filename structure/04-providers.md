# Providers and backends

OfficeCLI uses stable OpenXML handlers for DOCX/XLSX/PPTX and provider-gated engines for HWP/HWPX.

## Provider matrix

| Format | Default provider/backend | Optional provider/backend | Evidence files | Notes |
| --- | --- | --- | --- | --- |
| `.docx` | `WordHandler` + DocumentFormat.OpenXml | Windows Word/PDF backend only for requested native render/page-count paths | `schemas/help/docx/**`, Word handler tests | No Office install required for normal operations. |
| `.xlsx` | `ExcelHandler` + DocumentFormat.OpenXml | none documented as provider | `schemas/help/xlsx/**` | Formula evaluator, chart/pivot/table helpers live under `Core` and `Handlers/Excel`. |
| `.pptx` | `PowerPointHandler` + DocumentFormat.OpenXml | HTML/SVG screenshot helpers for visual output | `schemas/help/pptx/**` | PPTX handler owns slides, shapes, media, charts, themes, animations. |
| `.hwpx` | `custom` ZIP/XML provider | `rhwp-bridge` opt-in for selected read/render/field/text paths | `docs/qa/provider-compatibility-matrix.md`, `tests/fixtures/common/provider-compatibility.json` | Custom remains default until evidence parity says otherwise. |
| `.hwp` | `rhwp-bridge` experimental provider | Hancom reopen evidence lane for mutation promotion | Same provider matrix plus `docs/providers/rhwp-sidecar-contract.md` | Binary HWP work is capability-gated through rhwp sidecars. Missing sidecars or unverified operations must fail closed instead of falling back to HWPX. |

## HWP/rhwp bridge

Bridge setup from `CommandBuilder.Help.Hwp.cs`:

```bash
export OFFICECLI_HWP_ENGINE=rhwp-experimental
export OFFICECLI_RHWP_BIN=/path/to/rhwp
export OFFICECLI_RHWP_BRIDGE_PATH=/path/to/rhwp-officecli-bridge.dll
export OFFICECLI_RHWP_API_BIN=/path/to/rhwp-field-bridge
```

Probe commands:

```bash
officecli hwp doctor --json
officecli capabilities --json
```

## rhwp sidecar contract

`docs/providers/rhwp-sidecar-contract.md` defines request/response boundaries against schemas in `schemas/interfaces/`:

- requests include `schemaVersion`, `operation`, `format`, `inputPath`, and `outputPath` for mutations;
- responses include `schemaVersion`, `ok`, `operation`, `format`, `engineVersion`, plus data or typed error;
- errors include `code`, `message`, `format`, `operation`, `engine`, and `nextCommand` when useful.

Supported HWP sidecar operations listed in that contract and current bridge
surface:

- `read-text`, `render-svg`, `render-png`, `export-pdf`, `export-markdown`,
  `document-info`, `diagnostics`, `dump-controls`, `dump-pages`, `thumbnail`,
  `list-fields`, `get-field`, `set-field`, `replace-text`, `insert-text`,
  `get-cell-text`, `scan-cells`, `set-cell-text`, `convert-to-editable`,
  `native-op`, `create-blank`, `save-as-hwp`.

Supported HWPX rhwp operations listed there:

- `read-text`, `render-svg`, `render-png`, `export-pdf`, `export-markdown`,
  `document-info`, `diagnostics`, `dump-controls`, `dump-pages`, `thumbnail`,
  `list-fields`, `get-field`, `set-field`, `replace-text`, `insert-text`,
  `get-cell-text`, `scan-cells`, `set-cell-text`, `native-op`, `save-as-hwp`.

## Safe-save policy

`docs/safety/safe-save-policy.md` separates four categories:

1. documented shape;
2. handler-supported operation;
3. readback-verified output;
4. safe in-place mutation.

Only safe in-place mutations may overwrite input, and only with backup and transaction evidence. Current HWP text replacement exposes in-place mode behind `--in-place --backup --verify`; output-first remains the stable mutation policy.

## Provider compatibility rules

From `docs/qa/provider-compatibility-matrix.md`:

- HWPX `custom` is the default provider.
- HWP `rhwp-bridge` is the default provider.
- Every expected capability must have rows for `custom` and `rhwp-bridge`.
- Unsupported rows carry typed reasons such as `unsupported_engine`,
  `dependency_missing`, `operation_not_wired`, `roundtrip_unverified`, or
  `hancom_reopen_unverified`.
- Hancom evidence is the promotion lane for native HWP write claims. Normal CI
  may keep it manual, but docs must not call rhwp-backed operations impossible
  when the pinned rhwp source exposes the primitive.
