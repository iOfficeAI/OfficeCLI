# OfficeCLI

> **OfficeCLI is the world's first and the best command-line designed for AI agents.**

**Give any AI agent full control over Word, Excel, and PowerPoint -- in one line of code.**

Open-source. Single binary. No Office installation. No dependencies. Works everywhere.

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

**English** | [中文](README_zh.md)

<p align="center">
  <img src="assets/ppt-process.gif" alt="OfficeCLI creating a PowerPoint presentation on AionUi" width="100%">
</p>

<p align="center"><em>PPT creation process using OfficeCLI on <a href="https://github.com/iOfficeAI/AionUi">AionUi</a></em></p>

<p align="center"><strong>PowerPoint Presentations</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/designwhatmovesyou.gif" alt="OfficeCLI design presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/horizon.gif" alt="OfficeCLI business presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/efforless.gif" alt="OfficeCLI tech presentation (PowerPoint)"></td>
</tr>
<tr>
<td width="33%"><img src="assets/blackhole.gif" alt="OfficeCLI space presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/first-ppt-aionui.gif" alt="OfficeCLI gaming presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/shiba.gif" alt="OfficeCLI creative presentation (PowerPoint)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Word Documents</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/word1.gif" alt="OfficeCLI academic paper (Word)"></td>
<td width="33%"><img src="assets/showcase/word2.gif" alt="OfficeCLI project proposal (Word)"></td>
<td width="33%"><img src="assets/showcase/word3.gif" alt="OfficeCLI annual report (Word)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Excel Spreadsheets</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/excel1.gif" alt="OfficeCLI budget tracker (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel2.gif" alt="OfficeCLI gradebook (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel3.gif" alt="OfficeCLI sales dashboard (Excel)"></td>
</tr>
</table>

<p align="center"><em>All documents above were created entirely by AI agents using OfficeCLI — no templates, no manual editing.</em></p>

## For AI Agents — Get Started in One Line

Paste this into your AI agent's chat — it will read the skill file and install everything automatically:

```
curl -fsSL https://officecli.ai/SKILL.md
```

That's it. The skill file teaches the agent how to install the binary and use all commands.

> **Technical details:** OfficeCLI ships with a [SKILL.md](SKILL.md) (239 lines, ~8K tokens) that covers command syntax, architecture, and common pitfalls. After installation, your agent can immediately create, read, and modify any Office document.

## Quick Start

From zero to a finished presentation in seconds:

```bash
# Create a new PowerPoint
officecli create deck.pptx

# Add a slide with a title and background color
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E

# Add a text shape to the slide
officecli add deck.pptx /slide[1] --type shape \
  --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm \
  --prop font=Arial --prop size=24 --prop color=FFFFFF

# View the outline of the presentation
officecli view deck.pptx outline
```

Output:

```
Slide 1: Q4 Report
  Shape 1 [TextBox]: Revenue grew 25%
```

```bash
# Get structured JSON for any element
officecli get deck.pptx /slide[1]/shape[1] --json
```

```json
{
  "tag": "shape",
  "path": "/slide[1]/shape[1]",
  "attributes": {
    "name": "TextBox 1",
    "text": "Revenue grew 25%",
    "x": "720000",
    "y": "1800000"
  }
}
```

## Why OfficeCLI?

What used to take 50 lines of Python and 3 separate libraries:

```python
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.shapes.title
title.text = "Q4 Report"
# ... 45 more lines ...
prs.save('deck.pptx')
```

Now takes one command:

```bash
officecli add deck.pptx / --type slide --prop title="Q4 Report"
```

**What OfficeCLI can do:**

- **Create** documents from scratch -- blank or with content
- **Read** text, structure, styles, formulas -- in plain text or structured JSON
- **Analyze** formatting issues, style inconsistencies, and structural problems
- **Modify** any element -- text, fonts, colors, layout, formulas, charts, images
- **Reorganize** content -- add, remove, move, copy elements across documents

| Format | Read | Modify | Create |
|--------|------|--------|--------|
| Word (.docx) | ✅ | ✅ | ✅ |
| Excel (.xlsx) | ✅ | ✅ | ✅ |
| PowerPoint (.pptx) | ✅ | ✅ | ✅ |

**Word** — [paragraphs](https://github.com/iOfficeAI/OfficeCLI/wiki/word-paragraph), [runs](https://github.com/iOfficeAI/OfficeCLI/wiki/word-run), [tables](https://github.com/iOfficeAI/OfficeCLI/wiki/word-table), [styles](https://github.com/iOfficeAI/OfficeCLI/wiki/word-style), [headers/footers](https://github.com/iOfficeAI/OfficeCLI/wiki/word-header-footer), [images](https://github.com/iOfficeAI/OfficeCLI/wiki/word-picture), [equations](https://github.com/iOfficeAI/OfficeCLI/wiki/word-equation), [comments](https://github.com/iOfficeAI/OfficeCLI/wiki/word-comment), [footnotes](https://github.com/iOfficeAI/OfficeCLI/wiki/word-footnote), [watermarks](https://github.com/iOfficeAI/OfficeCLI/wiki/word-watermark), [bookmarks](https://github.com/iOfficeAI/OfficeCLI/wiki/word-bookmark), [TOC](https://github.com/iOfficeAI/OfficeCLI/wiki/word-toc), [charts](https://github.com/iOfficeAI/OfficeCLI/wiki/word-chart), [hyperlinks](https://github.com/iOfficeAI/OfficeCLI/wiki/word-hyperlink), [sections](https://github.com/iOfficeAI/OfficeCLI/wiki/word-section), [form fields](https://github.com/iOfficeAI/OfficeCLI/wiki/word-formfield), [content controls (SDT)](https://github.com/iOfficeAI/OfficeCLI/wiki/word-sdt), [fields](https://github.com/iOfficeAI/OfficeCLI/wiki/word-field), [document properties](https://github.com/iOfficeAI/OfficeCLI/wiki/word-document)

**Excel** — [cells](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-cell), formulas (150+ built-in functions with auto-evaluation), [sheets](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sheet), [tables](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-table), [conditional formatting](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-conditionalformatting), [charts](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart), [pivot tables](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-pivottable), [named ranges](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-namedrange), [data validation](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-validation), [images](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-picture), [sparklines](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sparkline), [comments](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-comment), [autofilter](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-autofilter), [shapes](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-shape), CSV/TSV import, `$Sheet:A1` cell addressing

**PowerPoint** — [slides](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide), [shapes](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-shape), [images](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-picture), [tables](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-table), [charts](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-chart), [animations](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide), [morph transitions](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-morph-check), [3D models (.glb)](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-3dmodel), [slide zoom](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-zoom), [equations](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-equation), [themes](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-theme), [connectors](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-connector), [video/audio](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-video), [groups](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-group), [notes](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-notes), [placeholders](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-placeholder)

## Use Cases

**For Developers:**
- Automate report generation from databases or APIs
- Batch-process documents (bulk find/replace, style updates)
- Build document pipelines in CI/CD environments (generate docs from test results)
- Headless Office automation in Docker/containerized environments

**For AI Agents:**
- Generate presentations from user prompts (see examples above)
- Extract structured data from documents to JSON
- Validate and check document quality before delivery

**For Teams:**
- Clone document templates and populate with data
- Automated document validation in CI/CD pipelines

## Installation

Ships as a single self-contained binary. The .NET runtime is embedded -- nothing to install, no runtime to manage.

**One-line install:**

```bash
# macOS / Linux
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash

# Windows (PowerShell)
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

**Or download manually** from [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases):

| Platform | Binary |
|----------|--------|
| macOS Apple Silicon | `officecli-mac-arm64` |
| macOS Intel | `officecli-mac-x64` |
| Linux x64 | `officecli-linux-x64` |
| Linux ARM64 | `officecli-linux-arm64` |
| Windows x64 | `officecli-win-x64.exe` |
| Windows ARM64 | `officecli-win-arm64.exe` |

Verify installation: `officecli --version`

**Or self-install from a downloaded binary:**

```bash
officecli install
```

Updates are checked automatically in the background. Disable with `officecli config autoUpdate false` or skip per-invocation with `OFFICECLI_SKIP_UPDATE=1`. Configuration lives under `~/.officecli/config.json`.

## Key Features

### Live Preview

`watch` starts a local HTTP server with a live HTML preview of your PowerPoint file. Every modification auto-refreshes in the browser — ideal for iterative design with AI agents.

```bash
officecli watch deck.pptx
# Opens http://localhost:18080 — refreshes on every set/add/remove
```

Renders shapes, charts, equations, 3D models (Three.js), morph transitions, zoom navigation, and all shape effects.

### Resident Mode & Batch

For multi-step workflows, resident mode keeps the document in memory. Batch mode runs multiple operations in one open/save cycle.

```bash
# Resident mode — near-zero latency via named pipes
officecli open report.docx
officecli set report.docx /body/p[1]/r[1] --prop bold=true
officecli set report.docx /body/p[2]/r[1] --prop color=FF0000
officecli close report.docx

# Batch mode — atomic multi-command execution
echo '[{"command":"set","path":"/slide[1]/shape[1]","props":{"text":"Hello"}},
      {"command":"set","path":"/slide[1]/shape[2]","props":{"fill":"FF0000"}}]' \
  | officecli batch deck.pptx --stop-on-error
```

### Three-Layer Architecture

Start simple, go deep only when needed.

| Layer | Purpose | Commands |
|-------|---------|----------|
| **L1: Read** | Semantic views of content | `view` (text, annotated, outline, stats, issues, html) |
| **L2: DOM** | Structured element operations | `get`, `query`, `set`, `add`, `remove`, `move` |
| **L3: Raw XML** | Direct XPath access — universal fallback | `raw`, `raw-set`, `add-part`, `validate` |

```bash
# L1 — high-level views
officecli view report.docx annotated
officecli view budget.xlsx text --cols A,B,C --max-lines 50

# L2 — element-level operations
officecli query report.docx "run:contains(TODO)"
officecli add budget.xlsx / --type sheet --prop name="Q2 Report"
officecli move report.docx /body/p[5] --to /body --index 1

# L3 — raw XML when L2 isn't enough
officecli raw deck.pptx /slide[1]
officecli raw-set report.docx document \
  --xpath "//w:p[1]" --action append \
  --xml '<w:r><w:t>Injected text</w:t></w:r>'
```

## AI Integration

### MCP Server

Built-in [MCP](https://modelcontextprotocol.io) server — register with one command:

```bash
officecli mcp claude       # Claude Code
officecli mcp cursor       # Cursor
officecli mcp vscode       # VS Code / Copilot
officecli mcp lmstudio     # LM Studio
officecli mcp list         # Check registration status
```

Exposes all document operations as tools over JSON-RPC — no shell access needed.

### Direct CLI Integration

Get OfficeCLI working with your AI agent in two steps:

1. **Install the binary** -- one command (see [Installation](#installation))
2. **Done.** OfficeCLI automatically detects your AI tools (Claude Code, GitHub Copilot, Codex) by checking known config directories and installs its skill file. Your agent can immediately create, read, and modify any Office document.

<details>
<summary><strong>Manual setup (optional)</strong></summary>

If auto-install doesn't cover your setup, you can install the skill file manually:

**Feed SKILL.md to your agent directly:**

```bash
curl -fsSL https://officecli.ai/SKILL.md
```

**Install as a local skill for Claude Code:**

```bash
curl -fsSL https://officecli.ai/SKILL.md -o ~/.claude/skills/officecli.md
```

**Other agents:** Include the contents of `SKILL.md` (239 lines, ~8K tokens) in your agent's system prompt or tool description.

</details>

**Call from any language:**

```python
# Python
import subprocess, json
def cli(*args): return subprocess.check_output(["officecli", *args], text=True)
cli("create", "deck.pptx")
cli("set", "deck.pptx", "/slide[1]/shape[1]", "--prop", "text=Hello")
```

```js
// JavaScript
const { execFileSync } = require('child_process')
const cli = (...args) => execFileSync('officecli', args, { encoding: 'utf8' })
cli('set', 'deck.pptx', '/slide[1]/shape[1]', '--prop', 'text=Hello')
```

Every command supports `--json` for structured output. Path-based addressing means agents don't need to understand XML namespaces.

### Why agents love OfficeCLI

- **Deterministic JSON output** -- Every command supports `--json`, returning structured data with consistent schemas. No regex parsing needed.
- **Path-based addressing** -- Every element has a stable path (`/slide[1]/shape[2]`). Agents navigate documents without understanding XML namespaces. Note: these paths use OfficeCLI's own syntax (1-based indexing, element local names), not XPath.
- **Progressive complexity** -- Start with L1 (read), escalate to L2 (modify), fall back to L3 (raw XML) only when needed. Minimizes token usage.
- **Self-healing workflow** -- `validate`, `view issues`, and the help system let agents detect problems and self-correct without human intervention.
- **Built-in help** -- When unsure about property names or value formats, run `officecli <format> set <element>` instead of guessing.
- **Auto-install** -- No manual skill-file setup. OfficeCLI detects your AI tools and configures itself automatically.

### Built-in Help

Don't guess property names — drill into the help:

```bash
officecli pptx set              # All settable elements and properties
officecli pptx set shape        # Detail for one element type
officecli pptx set shape.fill   # One property: format and examples
officecli docx query            # Selector reference: attributes, :contains, :has(), etc.
```

Run `officecli --help` for the full overview.

### JSON Output Schemas

All commands support `--json`. The general response shapes:

**Single element** (`get --json`):

```json
{"tag": "shape", "path": "/slide[1]/shape[1]", "attributes": {"name": "TextBox 1", "text": "Hello"}}
```

**List of elements** (`query --json`):

```json
[
  {"tag": "paragraph", "path": "/body/p[1]", "attributes": {"style": "Heading1", "text": "Title"}},
  {"tag": "paragraph", "path": "/body/p[5]", "attributes": {"style": "Heading1", "text": "Summary"}}
]
```

**Errors** return a non-zero exit code with a structured error object including error code, suggestion, and valid values when available:

```json
{
  "success": false,
  "error": {
    "error": "Slide 50 not found (total: 8)",
    "code": "not_found",
    "suggestion": "Valid Slide index range: 1-8"
  }
}
```

Error codes: `not_found`, `invalid_value`, `unsupported_property`, `invalid_path`, `unsupported_type`, `missing_property`, `file_not_found`, `file_locked`, `invalid_selector`. Property names are auto-corrected -- misspelling a property returns a suggestion with the closest match.

**Error Recovery** -- Agents self-correct by inspecting available elements:

```bash
# Agent tries an invalid path
officecli get report.docx /body/p[99] --json
# Returns: {"success": false, "error": {"error": "...", "code": "not_found", "suggestion": "..."}}

# Agent self-corrects by checking available elements
officecli get report.docx /body --depth 1 --json
# Returns the list of available children, agent picks the right path
```

**Mutation confirmations** (`set`, `add`, `remove`, `move`, `create` with `--json`):

```json
{"success": true, "path": "/slide[1]/shape[1]"}
```

See `officecli --help` for full details on exit codes and error formats.

## Comparison

| | OfficeCLI | Microsoft Office | LibreOffice | python-docx / openpyxl |
|---|---|---|---|---|
| Open source & free | ✓ (Apache 2.0) | ✗ (paid license) | ✓ | ✓ |
| AI-native CLI + JSON | ✓ | ✗ | ✗ | ✗ |
| Zero install (single binary) | ✓ | ✗ | ✗ | ✗ (Python + pip) |
| Call from any language | ✓ (CLI) | ✗ (COM/Add-in) | ✗ (UNO API) | Python only |
| Path-based element access | ✓ | ✗ | ✗ | ✗ |
| Raw XML fallback | ✓ | ✗ | ✗ | Partial |
| Live preview | ✓ | ✓ | ✗ | ✗ |
| Headless / CI | ✓ | ✗ | Partial | ✓ |
| Cross-platform | ✓ | Windows/Mac | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | Separate libs |

## Updates & Configuration

```bash
officecli config autoUpdate false              # Disable auto-update checks
OFFICECLI_SKIP_UPDATE=1 officecli ...          # Skip check for one invocation (CI)
```

## Command Reference

| Command | Description |
|---------|-------------|
| [`create`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-create) | Create a blank .docx, .xlsx, or .pptx (type from extension) |
| [`view`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-view) | View content (modes: `outline`, `text`, `annotated`, `stats`, `issues`, `html`) |
| [`get`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-get) | Get element and children (`--depth N`, `--json`) |
| [`query`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-query) | CSS-like query (`[attr=value]`, `:contains()`, `:has()`, etc.) |
| [`set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-set) | Modify element properties |
| [`add`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-add) | Add element (or clone with `--from <path>`) |
| [`remove`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-remove) | Remove an element |
| [`move`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-move) | Move element (`--to <parent> --index N`) |
| [`swap`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-swap) | Swap two elements |
| [`validate`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-validate) | Validate against OpenXML schema |
| [`batch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-batch) | Multiple operations in one open/save cycle (JSON on stdin or `--input`) |
| [`merge`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-merge) | Template merge — replace `{{key}}` placeholders with JSON data |
| [`watch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-watch) | Live HTML preview in browser with auto-refresh |
| [`mcp`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-mcp) | Start MCP server for AI tool integration |
| [`raw`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | View raw XML of a document part |
| [`raw-set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | Modify raw XML via XPath |
| `add-part` | Add a new document part (header, chart, etc.) |
| [`open`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-open) | Start resident mode (keep document in memory) |
| `close` | Save and close resident mode |
| [`install`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-install) | Install binary + skills + MCP (`all`, `claude`, `cursor`, etc.) |
| `config` | Get or set configuration |
| `<format> <command>` | [Built-in help](https://github.com/iOfficeAI/OfficeCLI/wiki/command-reference) (e.g. `officecli pptx set shape`) |

## End-to-End Workflow Example

A typical self-healing agent workflow: create a presentation, populate it, verify, and fix issues -- all without human intervention.

```bash
# 1. Create
officecli create report.pptx

# 2. Add content
officecli add report.pptx / --type slide --prop title="Q4 Results"
officecli add report.pptx /slide[1] --type shape \
  --prop text="Revenue: $4.2M" --prop x=2cm --prop y=5cm --prop size=28
officecli add report.pptx / --type slide --prop title="Details"
officecli add report.pptx /slide[2] --type shape \
  --prop text="Growth driven by new markets" --prop x=2cm --prop y=5cm

# 3. Verify
officecli view report.pptx outline
officecli validate report.pptx

# 4. Fix any issues found
officecli view report.pptx issues --json
# Address issues based on output, e.g.:
officecli set report.pptx /slide[1]/shape[1] --prop font=Arial
```

### Template Merge

Replace `{{key}}` placeholders in any document with JSON data -- works across paragraphs, table cells, shapes, headers, footers, and chart titles.

```bash
# Merge from inline JSON
officecli merge template.docx output.docx '{"name":"Alice","dept":"Sales","date":"2026-03-30"}'

# Merge from a JSON file
officecli merge template.pptx report.pptx data.json

# Excel template
officecli merge budget-template.xlsx q4-budget.xlsx '{"quarter":"Q4","year":"2026"}'
```

### Units & Colors

All dimension and color properties accept flexible input formats:

| Type | Accepted formats | Examples |
|------|-----------------|----------|
| **Dimensions** | cm, in, pt, px, or raw EMU | `2cm`, `1in`, `72pt`, `96px`, `914400` |
| **Colors** | Hex, named, RGB, theme | `#FF0000`, `FF0000`, `red`, `rgb(255,0,0)`, `accent1` |
| **Font sizes** | Bare number or pt-suffixed | `14`, `14pt`, `10.5pt` |
| **Spacing** | pt, cm, in, or multiplier | `12pt`, `0.5cm`, `1.5x`, `150%` |

## Common Patterns

```bash
# Replace all Heading1 text in a Word doc
officecli query report.docx "paragraph[style=Heading1]" --json | ...
officecli set report.docx /body/p[1]/r[1] --prop text="New Title"

# Export all slide content as JSON
officecli get deck.pptx / --depth 2 --json

# Bulk-update Excel cells
officecli batch budget.xlsx --input updates.json --json

# Import CSV data into an Excel sheet
officecli add budget.xlsx / --type sheet --prop name="Q1 Data" --prop csv=sales.csv

# Template merge for batch reports
officecli merge invoice-template.docx invoice-001.docx '{"client":"Acme","total":"$5,200"}'

# Check document quality before delivery
officecli validate report.docx && officecli view report.docx issues --json
```

### Windows / Git Bash note

On Windows, prefer **PowerShell** when passing OfficeCLI paths or parts that begin with `/`, such as `/document`, `/header[1]`, or `/body/p[3]`.

Git Bash / MSYS may rewrite those values into filesystem paths before `officecli` sees them. If you must use Git Bash, prefix the command with:

```bash
MSYS_NO_PATHCONV=1 officecli raw doc.docx /document
```

The same caution applies to `raw-set`, `get`, `query`, `set`, `remove`, and similar commands that accept OfficeCLI document paths.

## Documentation

The [Wiki](https://github.com/iOfficeAI/OfficeCLI/wiki) has detailed guides for every command, element type, and property:

- **By format:** [Word](https://github.com/iOfficeAI/OfficeCLI/wiki/word-reference) | [Excel](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-reference) | [PowerPoint](https://github.com/iOfficeAI/OfficeCLI/wiki/powerpoint-reference)
- **Workflows:** [End-to-end examples](https://github.com/iOfficeAI/OfficeCLI/wiki/workflows) -- Word reports, Excel dashboards, PowerPoint decks, batch modifications, resident mode
- **Troubleshooting:** [Common errors and solutions](https://github.com/iOfficeAI/OfficeCLI/wiki/troubleshooting)
- **AI agent guide:** [Decision tree for navigating the wiki](https://github.com/iOfficeAI/OfficeCLI/wiki/agent-guide)

## Build from Source

Requires [.NET 10 SDK](https://dotnet.microsoft.com/download) for compilation only. The output is a self-contained, native binary -- .NET is embedded in the binary and is not needed at runtime.

```bash
./build.sh
```

## License

[Apache License 2.0](LICENSE)

Bug reports and contributions are welcome on [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues).

---

[OfficeCLI.AI](https://OfficeCLI.AI) | [GitHub](https://github.com/iOfficeAI/OfficeCLI)

<!--
yaml-frontmatter:
  name: officecli
  version: see GitHub Releases
  description: The standard CLI for AI agents to create, read, and modify Office documents
  formats: [docx, xlsx, pptx]
  platforms: [macos, linux, windows]
  license: Apache-2.0
  keywords: [office, cli, ai-agent, automation, docx, xlsx, pptx, openxml, document-generation, mcp-server, word, excel, powerpoint]
  ai-agent-compatible: true
  mcp-server: true
  skill-file: SKILL.md
  skill-file-lines: 239
  install-command-unix: curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
  install-command-windows: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
-->

<!-- LLM/agent discovery metadata
tool: officecli
type: cli
canonical: the standard CLI for AI-driven Office document manipulation
formats: docx, xlsx, pptx
capabilities: create, read, modify, validate, batch, resident-mode, mcp-server, live-preview, auto-install
output: json, text, annotated
platforms: macos, linux, windows
license: Apache-2.0
keywords: office, cli, ai-agent, automation, docx, xlsx, pptx, openxml, document-generation, mcp-server, word, excel, powerpoint, ai-tools, command-line, structured-output
ai-agent-compatible: true
mcp-server: true
skill-file: SKILL.md
skill-file-lines: 239
alternatives: python-docx, openpyxl, python-pptx, libreoffice --headless
install-command-unix: curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
install-command-windows: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
-->
