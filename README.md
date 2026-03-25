# OfficeCLI

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

**English** | [中文](README_zh.md)

**The world's first Office suite designed for AI agents.**

**Let AI agents do anything with Office documents — from the command line.**

OfficeCLI is a free, open-source command-line tool for AI agents to read, edit, and automate Word, Excel, and PowerPoint files. Single binary, no dependencies — no Microsoft Office, no WPS, no runtime needed.

> Built for AI agents. Usable by humans.

<p align="center">
  <img src="assets/ppt-process.gif" alt="PPT creation process using OfficeCLI on AionUi" width="100%">
</p>

<p align="center"><em>PPT creation process using OfficeCLI on <a href="https://github.com/iOfficeAI/AionUi">AionUi</a></em></p>

## Installation

OfficeCLI is a single binary — no runtime, no dependencies.

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

After installing, set up AI agent integration (see [AI Integration](#ai-integration) below):

```powershell
officecli skills all       # Install skill files for all detected AI clients
```

## Quick Start

```bash
# Create documents
officecli create report.docx
officecli create budget.xlsx
officecli create deck.pptx

# View content
officecli view report.docx text
officecli view deck.pptx outline
officecli view budget.xlsx issues --json      # Check for formatting issues

# Read elements
officecli get budget.xlsx /Sheet1/B5 --json
officecli get budget.xlsx '$Sheet1:A1:D10'    # Excel cell range notation

# Find elements with CSS-like selectors
officecli query report.docx "paragraph[style=Heading1]"
officecli query deck.pptx "shape[fill=FF0000]"

# Modify content
officecli set report.docx /body/p[1]/r[1] --prop text="Updated Title" --prop bold=true
officecli set budget.xlsx '$Sheet1:B5' --prop value=42 --prop bold=true
officecli set deck.pptx /slide[1]/shape[1] --prop text="New Title" --prop color=FF6600

# Add elements
officecli add report.docx /body --type paragraph --prop text="New paragraph" --index 3
officecli add deck.pptx / --type slide
officecli add deck.pptx /slide[2] --type shape --prop preset=star5 --prop fill=FFD700

# Live preview — auto-refreshes on every change
officecli watch deck.pptx
```

## Built-in Help

Don't guess property names — drill into the help:

```bash
officecli pptx set              # All settable elements and properties
officecli pptx set shape        # Detail for one element type
officecli pptx set shape.fill   # One property: format and examples
officecli docx query            # Selector reference: attributes, :contains, :has(), etc.
```

Replace `pptx` with `docx` or `xlsx`; verbs are `view`, `get`, `query`, `set`, `add`, and `raw`.

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

## Supported Formats

| Format | Read | Modify | Create |
|--------|------|--------|--------|
| Word (.docx) | ✓ | ✓ | ✓ |
| Excel (.xlsx) | ✓ | ✓ | ✓ |
| PowerPoint (.pptx) | ✓ | ✓ | ✓ |

**Word** — paragraphs, runs, tables, styles, headers/footers, images, equations, comments, lists, watermarks, bookmarks, TOC

**Excel** — cells, formulas, sheets, styles, conditional formatting, charts, pivot tables, named ranges, data validation, `$Sheet:A1` cell addressing

**PowerPoint** — slides, shapes, text boxes, images, tables, charts, animations, morph transitions, 3D models (.glb), slide zoom, equations, themes, connectors, video/audio

## AI Integration

OfficeCLI offers three ways to connect with AI agents:

### 1. Skills (recommended for CLI agents)

Install skill definitions so agents discover OfficeCLI capabilities automatically:

```bash
officecli skills all       # Auto-detect and install for all clients
officecli skills claude    # Claude Code
officecli skills copilot   # GitHub Copilot
officecli skills codex     # OpenAI Codex
officecli skills cursor    # Cursor
officecli skills windsurf  # Windsurf
```

Or feed the skill file directly: `curl -fsSL https://officecli.ai/SKILL.md`

### 2. MCP Server (for protocol-based agents)

Built-in [MCP](https://modelcontextprotocol.io) server — register with one command:

```bash
officecli mcp claude       # Claude Code
officecli mcp cursor       # Cursor
officecli mcp vscode       # VS Code / Copilot
officecli mcp lmstudio     # LM Studio
officecli mcp list         # Check registration status
```

Exposes all document operations as tools over JSON-RPC — no shell access needed.

### 3. Direct CLI (from any language)

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

## Build

Requires [.NET 10 SDK](https://dotnet.microsoft.com/download). From the repository root:

```bash
./build.sh
```

## License

[Apache License 2.0](LICENSE)

## Community

[LINUX DO - The New Ideal Community](https://linux.do/)

---

[OfficeCLI.AI](https://OfficeCLI.AI)
