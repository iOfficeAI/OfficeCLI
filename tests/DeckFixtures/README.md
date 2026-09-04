# Native deck fixtures

`business-light` and `data-report` each exercise the full semantic layout catalog.
`technology-dark` and `editorial-report` keep the original twelve-layout smoke coverage.
Catalog size is themes × layouts (currently 6 × 26 = 156). The SVG is the shared local media fixture; no remote assets are fetched during compilation.

`Verify-DeckExportPdf.ps1` covers `deck export-pdf` (missing exporter → `exporter_not_found`, plus a stub exporter under `FakePdfExporter/`). PDF conversion always goes through an exporter plugin — same as `view <pptx> pdf`; not HTML/Chrome.
