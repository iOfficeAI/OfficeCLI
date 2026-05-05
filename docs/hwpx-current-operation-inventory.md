# HWPX Current Operation Inventory

This file prevents broad HWPX write claims. An operation can be advertised to cli-jaw
only when `officecli capabilities --json` reports it as `roundtrip-verified` and the
evidence files listed here are complete.

| Operation | Current engine | Status | Evidence file(s) | Hancom evidence file | Advertise to cli-jaw? |
|---|---|---|---|---|---|
| `read_text` | `custom` | `experimental` | `tests/fixtures/hwpx/text-basic.golden.txt` planned | none yet | no |
| `render_svg` | `none` | `unsupported` | none | none | no |
| `fill_field` | `custom` | `experimental` | planned | none yet | no |
| `save_original` | `custom` | `experimental` | planned | none yet | no |
| `save_as_hwp` | `none` | `unsupported` | none | none | no |

Current public wording must be:

```text
OfficeCLI advertises only the HWPX operations listed as roundtrip-verified in
officecli capabilities --json.
```

Do not use:

```text
OfficeCLI supports HWPX writing.
OfficeCLI supports current roundtrip-verified HWPX XML-first operations.
```
