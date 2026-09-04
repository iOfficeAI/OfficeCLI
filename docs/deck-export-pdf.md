# `officecli deck export-pdf`

Thin wrap: compile a WorkMate DeckSpec to editable PPTX via `deck build`, then convert that PPTX to PDF through the same **exporter plugin** path as `officecli view <pptx> pdf` (`ExporterInvoker`). **Not** Dashi / AGPL / Chrome HTML export.

## Usage

```bash
officecli deck export-pdf deck.workmate-deck.json -o deck.pdf --json

# Defaults: sibling PDF next to the spec (foo.workmate-deck.json → foo.pdf)
officecli deck export-pdf foo.workmate-deck.json --json

# Also keep the intermediate PPTX
officecli deck export-pdf deck.workmate-deck.json -o deck.pdf --pptx deck.pptx --json
```

### Options

| Flag | Description |
|------|-------------|
| `--output` / `-o` | Output `.pdf` path (default: sibling of spec with `.pdf`) |
| `--pptx` | Keep intermediate PPTX at this path (default: temp file, deleted after export) |
| `--expected-revision` | Reject stale DeckSpec builds (same as `deck build`) |
| `--json` | JSON stdout (deck commands always emit structured JSON envelopes on error) |

Stdout on success (camelCase):

```json
{
  "success": true,
  "output": "/abs/path/deck.pdf",
  "revision": 3,
  "pptx": "/abs/path/deck.pptx",
  "plugin": "officecli-pdf-stub"
}
```

`pptx` is omitted unless `--pptx` was passed. `plugin` is the exporter manifest `name`.

## Exporter dependency

PDF conversion requires an installed **exporter** plugin for `.pdf` (see `plugins/plugin-protocol.md`). Typical production plugins wrap LibreOffice / OnlyOffice / a commercial renderer.

- Discover: `officecli plugins list`
- Override path: `OFFICECLI_PLUGIN_EXPORTER_PDF=/path/to/plugin`
- Missing plugin → `exporter_not_found` (same as `view … pdf`)

OfficeCLI itself does **not** bundle a PDF engine (size / license / platform). CSBU WorkMate ships the CLI; operators install or bundle an exporter separately when PDF export is required.

## Equivalent manual steps

```bash
officecli deck build deck.workmate-deck.json -o deck.pptx --json
officecli view deck.pptx pdf -o deck.pdf --json
```

`deck export-pdf` is those two steps with a temp PPTX (unless `--pptx`) and one JSON result.

## Agent / Studio notes

1. Spec must be `stage: "ready"` (same gate as `deck build`).
2. Prefer validate before export: `officecli deck validate <spec> --json`.
3. Presentation Studio’s primary export remains editable PPTX; PDF is optional once an exporter is available.
