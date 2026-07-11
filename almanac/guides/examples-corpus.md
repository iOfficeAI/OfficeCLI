---
title: "Examples Corpus"
summary: "The examples directory is a runnable corpus for Word, Excel, and PowerPoint features, pairing CLI scripts, Python SDK scripts, walkthroughs, and generated Office files."
topics: [guides, examples]
sources:
  - id: examples-readme
    type: file
    path: examples/README.md
  - id: word-examples
    type: file
    path: examples/word/
  - id: excel-examples
    type: file
    path: examples/excel/
  - id: ppt-examples
    type: file
    path: examples/ppt/
  - id: excel-cell-cli
    type: file
    path: examples/excel/cell-formatting.sh
  - id: excel-cell-sdk
    type: file
    path: examples/excel/cell-formatting.py
---

The examples corpus is the practical proving ground for OfficeCLI's document features. It covers Word, Excel, and PowerPoint, and most examples ship as a four-file set: a Markdown walkthrough, a shell script that drives the CLI, a Python script that drives the SDK, and a generated `.docx`, `.xlsx`, or `.pptx` output file [@examples-readme]. Use it to learn command shapes, test feature changes, and confirm that the CLI and SDK paths stay equivalent.

## Pick The Area

Word examples live under `examples/word/` and cover formatting, charts, content controls, fields, pictures, numbering, diagrams, revisions, sections, tables, and text boxes [@examples-readme] [@word-examples]. Excel examples live under `examples/excel/` and include cell formatting, conditional formatting, data validation, sheet and workbook settings, sparklines, pivot tables, slicers, shapes, and a large chart set [@examples-readme] [@excel-examples]. PowerPoint examples live under `examples/ppt/` and cover presentations, settings, diagrams, animations, video, 3D models, charts, tables, transitions, shapes, text boxes, pictures, OLE embedding, and template builders [@examples-readme] [@ppt-examples].

Start with the Markdown walkthrough when you need intent. Run the `.sh` script when testing the command-line surface. Run the `.py` twin when testing the [SDK path](using-sdks). Compare the generated Office file when the visual or structural output matters [@examples-readme].

## Run And Verify

Each standard example can be regenerated with either `bash <name>.sh` or `python3 <name>.py` [@examples-readme]. The shell scripts show direct `officecli create`, `add`, `set`, `get`, `close`, and `validate` usage [@examples-readme]. The Python scripts use the same batch-shaped command dictionaries sent through the SDK resident client [@examples-readme].

`examples/excel/cell-formatting` is a useful template for feature coverage. Its shell version creates a workbook, opens a resident, issues many `set` and `add` commands, closes the file, reads selected cells back with `get --json`, and validates the final workbook [@excel-cell-cli]. Its Python twin creates the same workbook with `officecli.create(..., "--force")` and groups equivalent command dictionaries into `doc.batch(...)` calls [@excel-cell-sdk].

## Add Or Update An Example

When adding feature coverage, keep the example small enough to read but complete enough to regenerate its output. The README's contributor notes ask for a clear script, tested output, placement in the right document-type directory, and an update to the appropriate directory documentation [@examples-readme]. If a feature exists on both CLI and SDK paths, add both scripts so future changes can catch drift between direct commands and resident batch delivery.
