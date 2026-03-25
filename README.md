# OfficeCLI

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

**English** | [中文](README_zh.md)

**The world's first Office suite designed for AI agents.**

**Let AI agents do anything with Office documents — from the command line.**

OfficeCLI is a free, open-source command-line tool for AI agents to read, edit, and automate Word, Excel, and PowerPoint files. Single binary, no dependencies — no Microsoft Office, no WPS, no runtime needed.

> Built for AI agents. Usable by humans.

<video src="assets/ppt-processs.mp4" poster="assets/ppt-process.png" autoplay loop muted playsinline width="100%"></video>

<p align="center"><em>PPT creation process using OfficeCLI on <a href="https://github.com/iOfficeAI/AionUi">AionUi</a></em></p>

## For AI Agents

OfficeCLI ships with a [SKILL.md](SKILL.md) that teaches AI agents how to use it effectively.

Talk to your agent with this first:

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/SKILL.md
```

If your agent supports local skill installation, install it locally instead:

**Claude Code:**

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/SKILL.md -o ~/.claude/skills/officecli.md
```

**Other agents:**

Include the contents of that `SKILL.md` in your agent's system prompt or tool description.

Then install the CLI binary:

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
```

For Windows (PowerShell):

```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

## Why OfficeCLI?

AI agents are great at text — but Office documents are binary blobs of XML. OfficeCLI bridges this gap, letting agents:

- **Create** documents from scratch — blank or with content
- **Read** text, structure, styles, formulas — in plain text or structured JSON
- **Analyze** formatting issues, style inconsistencies, and structural problems
- **Modify** any element — text, fonts, colors, layout, formulas, charts, images
- **Reorganize** content — add, remove, move, copy elements across documents

All through simple CLI commands, with structured JSON output, no Office installation needed.

## Installation

OfficeCLI is a single binary — no runtime, no dependencies. One command to install:

**macOS / Linux:**

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
```

**Windows (PowerShell):**

```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

Or download manually from [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases).

## Quick Start

```bash
# Create a blank document
officecli create report.docx
officecli create budget.xlsx
officecli create deck.pptx

# View document content
officecli view report.docx text

# Check for formatting issues
officecli view report.docx issues --json

# Read a specific cell
officecli get budget.xlsx /Sheet1/B5 --json

# Modify content
officecli set report.docx /body/p[1]/r[1] --prop text="Updated Title" --prop bold=true

# Batch editing with resident mode (keeps doc in memory)
officecli open presentation.pptx
officecli set presentation.pptx /slide[1]/shape[1] --prop text="New Title"
officecli set presentation.pptx /slide[2]/shape[3] --prop text="New Subtitle"
officecli close presentation.pptx
```

## Built-in help

When property names or value formats are unclear, use the nested help instead of guessing. Replace `pptx` with `docx` or `xlsx`; verbs are `view`, `get`, `query`, `set`, `add`, and `raw`.

```bash
officecli pptx set              # All settable elements and properties
officecli pptx set shape        # Detail for one element type
officecli pptx set shape.fill   # One property: format and examples
```

Run `officecli --help` for the same overview. Optional wiki content may be embedded in the binary when built with a `wiki/` directory.

## Three-Layer Architecture

OfficeCLI is designed with a progressive complexity model — start simple, go deep only when needed.

### L1: Read & Inspect

High-level, semantic views of document content.

```bash
# Word — plain text with element paths
officecli view report.docx text

# Word — text with formatting annotations
officecli view report.docx annotated

# Excel — view with column filter
officecli view budget.xlsx text --cols A,B,C --max-lines 50

# Excel — detect formula and style issues
officecli view budget.xlsx issues --json

# PowerPoint — outline all slides
officecli view deck.pptx outline

# PowerPoint — stats on fonts and styles used
officecli view deck.pptx stats
```

### L2: DOM Operations

Modify documents through structured element paths and properties.

```bash
# Word — query headings and set formatting (CSS-like selectors; see help for full syntax)
officecli query report.docx "paragraph[style=Heading1]"
officecli docx query            # Selector reference: attributes, :contains, :has(), etc.
officecli set report.docx /body/p[1]/r[1] --prop bold=true --prop color=FF0000

# Word — add a paragraph, remove another
officecli add report.docx /body --type paragraph --prop text="New paragraph" --index 3
officecli remove report.docx /body/p[5]

# Excel — read and modify cells
officecli get budget.xlsx /Sheet1/B5 --json
officecli set budget.xlsx /Sheet1/A1 --prop formula="=SUM(A2:A10)" --prop numFmt="0.00%"

# Excel — add a new sheet, add rows
officecli add budget.xlsx / --type sheet --prop name="Q2 Report"
officecli add budget.xlsx /Sheet1 --type row --prop values="Name,Amount,Date"

# PowerPoint — modify slide content
officecli set deck.pptx /slide[1]/shape[1] --prop text="New Title"
officecli set deck.pptx /slide[2]/shape[3] --prop fontSize=24 --prop bold=true

# PowerPoint — add a slide, copy a shape from another slide
officecli add deck.pptx / --type slide
officecli add deck.pptx /slide[3] --from /slide[1]/shape[2]

# Move elements
officecli move report.docx /body/p[5] --to /body --index 1
```

### L3: Raw XML

Direct XML access via XPath — the universal fallback for any OpenXML operation.

```bash
# Word — view and modify raw XML
officecli raw report.docx document
officecli raw-set report.docx document \
  --xpath "//w:p[1]" \
  --action append \
  --xml '<w:r><w:t>Injected text</w:t></w:r>'

# Word — add a header
officecli add-part report.docx /body --type header

# Excel — view raw sheet XML
officecli raw budget.xlsx /Sheet1

# Excel — add a chart to a sheet
officecli add-part budget.xlsx /Sheet1 --type chart

# PowerPoint — view raw slide XML
officecli raw deck.pptx /slide[1]

# Validate any document
officecli validate report.docx
officecli validate budget.xlsx
```

## Supported Formats

| Format | Read | Modify | Create |
|--------|------|--------|--------|
| Word (.docx) | ✓ | ✓ | ✓ |
| Excel (.xlsx) | ✓ | ✓ | ✓ |
| PowerPoint (.pptx) | ✓ | ✓ | ✓ |

### Word — Paragraphs, runs, tables, styles, headers/footers, images, equations, comments, lists

### Excel — Cells, formulas, sheets, styles (fonts, fills, borders, number formats), conditional formatting, charts

### PowerPoint — Slides, shapes, text boxes, images, animations, equations, themes, alignment, and shape effects

## Resident Mode

For multi-step workflows, resident mode keeps the document open in a background process, eliminating reload overhead on every command.

```bash
officecli open report.docx        # Start resident process
officecli view report.docx text   # Instant — no file reload
officecli set report.docx ...     # Instant — no file reload
officecli close report.docx       # Save and stop
```

Communication happens via named pipes for near-zero latency between commands.

## Batch mode

Run multiple operations in **one** open/save cycle (or forward to an already resident process) by passing a JSON array of commands on stdin or via `--input`.

```bash
echo '[{"command":"view","mode":"outline"},{"command":"get","path":"/slide[1]","depth":1}]' \
  | officecli batch deck.pptx
```

Use `--stop-on-error` to abort on the first failure. Supported `command` values in each item include `get`, `query`, `set`, `add`, `remove`, `move`, `view`, `raw`, `raw-set`, and `validate` (see `BatchItem` in the source for all fields).

## Updates & configuration

The CLI may check for updates in the background (non-blocking) and align with the latest release from GitHub. Configuration lives under `~/.officecli/config.json`.

- **Disable automatic update checks:** `officecli config autoUpdate false` (read current value with `officecli config autoUpdate`).
- **Skip the background check for a single invocation (e.g. CI):** `OFFICECLI_SKIP_UPDATE=1 officecli ...`

Re-running the install script (`install.sh` / `install.ps1`) still installs or upgrades the binary as before.

## Python Usage

```python
import subprocess, json

def cli(*args): return subprocess.check_output(["officecli", *args], text=True)
def cli_json(*args): return json.loads(cli(*args, "--json"))

cli("create", "deck.pptx")
cli("set", "deck.pptx", "/slide[1]/shape[1]", "--prop", "text=Hello")
shapes = cli_json("query", "deck.pptx", "shape")
```

## JavaScript Usage

```js
const { execFileSync } = require('child_process')

const cli = (...args) => execFileSync('officecli', args, { encoding: 'utf8' })
const cliJson = (...args) => JSON.parse(cli(...args, '--json'))

cli('create', 'deck.pptx')
cli('set', 'deck.pptx', '/slide[1]/shape[1]', '--prop', 'text=Hello')
const shapes = cliJson('query', 'deck.pptx', 'shape')
```

## AI Agent Integration

### Why OfficeCLI for agents?

**Deterministic JSON output** — Every command supports `--json`, returning structured data with consistent schemas. No regex parsing needed.

**Useful validation and diagnostics** — Commands like `validate`, `view issues`, and `raw-set` help agents detect problems and verify document correctness after changes.

**Path-based addressing** — Every element in every document has a stable path. Agents can navigate documents without understanding XML namespaces.

**Progressive complexity** — Agents start with L1 (read), escalate to L2 (modify), and fall back to L3 (raw XML) only when needed. This minimizes token usage while keeping all operations possible.

## Comparison

How does OfficeCLI compare to other approaches for AI agents working with Office documents?

| | OfficeCLI | Microsoft Office | LibreOffice | python-docx / openpyxl |
|---|---|---|---|---|
| Open source & free | ✓ (Apache 2.0) | ✗ (paid license) | ✓ | ✓ |
| AI-friendly CLI | ✓ | ✗ | Partial | ✗ |
| Structured JSON output | ✓ | ✗ | ✗ | ✗ |
| Zero install (single binary) | ✓ | ✗ | ✗ | ✗ (Python + pip) |
| Call from any language | ✓ (CLI) | ✗ (COM/Add-in) | ✗ (UNO API) | ✗ (Python only) |
| Path-based element access | ✓ | ✗ | ✗ | ✗ |
| Raw XML fallback | ✓ | ✗ | ✗ | Partial |
| Resident mode (in-memory) | ✓ | ✗ | ✗ | ✗ |
| Works in headless/CI environments | ✓ | ✗ | Partial | ✓ |
| Cross-platform | ✓ | ✗ (Windows/Mac) | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | Separate libs |
| Read + Write + Create | ✓ | ✓ | ✓ | ✓ |

## Build

Requires [.NET 10 SDK](https://dotnet.microsoft.com/download) for local compilation. From the repository root:

```bash
./build.sh
```

## License

[Apache License 2.0](LICENSE)

## Community

[LINUX DO - The New Ideal Community](https://linux.do/)

---

[OfficeCLI.AI](https://OfficeCLI.AI)
