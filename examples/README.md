# OfficeCLI Examples

Comprehensive examples demonstrating OfficeCLI capabilities for Word, Excel, and PowerPoint automation.

For a standalone web-to-document pipeline, see [Web research to a Word report](mcp/).
It uses Parallel Search MCP to collect source excerpts, then writes them with OfficeCLI.

## 📂 Directory Structure

Every example below ships all four files — `.{md,sh,py,<ext>}` — even where the
tree abbreviates. Only `ppt/templates/styles/*/build*.sh` are `.sh`-only.

```
examples/
├── README.md                              # This file
├── word/                                  # 📄 Word examples — *.{md,sh,py,docx}
│   ├── formulas.{md,sh,py,docx}             # LaTeX math/chemistry/physics formulas
│   ├── tables.{md,sh,py,docx}               # styled tables
│   ├── textbox.{md,sh,py,docx}              # formatted text boxes
│   ├── charts.{md,sh,py,docx}               # inline charts (14 types incl. treemap/waterfall)
│   ├── run-formatting.{md,sh,py,docx}       # run/character property surface
│   ├── paragraph-formatting.{md,sh,py,docx} # paragraph property surface
│   ├── document-formatting.{md,sh,py,docx}  # document-level property surface
│   ├── sections.{md,sh,py,docx}             # section layout — multi-column, footnote/endnote props, per-section page setup
│   ├── content-controls.{md,sh,py,docx}     # SDT content controls — text/dropdown/combobox/date/picture/group intake form
│   ├── fields.{md,sh,py,docx}               # field codes (PAGE/DATE/REF/IF/HYPERLINK) + auto-populating table of contents
│   ├── pictures.{md,sh,py,docx}             # inline/floating images — crop, alt, wrap, behind-text, absolute position
│   ├── numbering.{md,sh,py,docx}            # list/numbering styles
│   ├── diagram.{md,sh,py,docx}              # Mermaid diagrams — native editable shapes + full-fidelity PNG
│   └── revisions.{md,sh,py,docx}            # tracked-change (revision) API
├── excel/                                 # 📊 Excel examples — *.{md,sh,py,xlsx}
│   ├── cell-formatting.{md,sh,py,xlsx}      # full cell property surface (fonts/fills/borders/numFmt/data)
│   ├── conditional-formatting.{md,sh,py,xlsx}
│   ├── data-validation.{md,sh,py,xlsx}        # dropdown lists, number/date/text/custom rules, input & error messages
│   ├── sheet-settings.{md,sh,py,xlsx}         # freeze panes, print area/titles, headers/footers, display & protection
│   ├── sparklines.{md,sh,py,xlsx}             # in-cell line/column/win-loss mini charts + point markers
│   ├── workbook-settings.{md,sh,py,xlsx}
│   ├── pivot-tables.{md,sh,py,xlsx}
│   ├── slicers.{md,sh,py,xlsx}                # pivot-table slicers (field/caption/columnCount/rowHeight)
│   ├── shapes.{md,sh,py,xlsx}                 # drawing shapes — geometry, flip, glow, gradient, reflection, outline
│   ├── charts.{md,sh,py,xlsx}               # master chart showcase
│   └── charts/                              # per-type chart scripts — charts-<type>.{md,sh,py,xlsx}
│           (demo, basic, advanced, extended, area, bar, boxwhisker,
│            bubble, column, combo, histogram, line, pie, radar,
│            scatter, stock, waterfall)
└── ppt/                                   # 🎨 PowerPoint examples — *.{md,sh,py,pptx}
    ├── presentation.{md,sh,py,pptx}
    ├── presentation-settings.{md,sh,py,pptx}
    ├── diagram.{md,sh,py,pptx}              # Mermaid diagrams — native editable shapes + full-fidelity PNG
    ├── animations.{md,sh,py,pptx}
    ├── video.{md,sh,py,pptx}
    ├── 3d-model.{md,sh,py,pptx}
    ├── charts/                              # charts-<type>.{md,sh,py,pptx}
    │       (column, bar, line, pie, doughnut, area, scatter,
    │        bubble, radar, stock, combo, waterfall, 3d, advanced)
    ├── tables/                              # tables-<topic>.{md,sh,py,pptx}
    │       (basic, styled, merged, borders, rows-cols, financial, nested)
    ├── transitions/                         # transitions-<topic>.{md,sh,py,pptx}
    │       (basic, directional, shapes, bands, dynamic, modern, random, timing, morph)
    ├── shapes/                              # shapes-<topic>.{md,sh,py,pptx}
    │       (basic, connectors, effects, typography)
    ├── textboxes/                           # textboxes-<topic>.{md,sh,py,pptx}
    │       (basic, advanced)
    ├── pictures/                            # pictures-basic.{md,sh,py,pptx}
    ├── ole/                                 # ole-embed.{md,sh,py,pptx} — embedded Excel/Word OLE objects
    └── templates/styles/*/build*.sh         # full-deck template generators (.sh only)
```

Each example ships the same **four-file set**:

| File | Role |
|------|------|
| `<name>.md` | Walkthrough — what the example demonstrates and the key techniques |
| `<name>.sh` | **CLI** build script — drives the `officecli` binary directly (`officecli create / add / set / close / validate`) |
| `<name>.py` | **SDK** build script — drives the [`officecli` Python SDK](../sdk/python/) (`import officecli`; `with officecli.create(...) as doc: doc.batch([...])`) |
| `<name>.<ext>` | Pre-generated output (`.docx` / `.xlsx` / `.pptx`) |

The `.sh` and `.py` are **equivalent twins** — both regenerate the same output document, one via the command line and one via the Python SDK. Run whichever fits your workflow.

> The full-deck template generators under `ppt/templates/styles/*/build*.sh` are
> standalone deck builders, not four-file API examples, so they ship as `.sh` only.

---

## 🚀 Quick Start

Every example runs **two equivalent ways** — pick one:

```bash
bash <name>.sh        # via the officecli CLI binary
python3 <name>.py     # via the officecli Python SDK  (pip install officecli-sdk)
```

Both regenerate the same output document. The commands below show one form per
example; swap `bash …​.sh` ⇄ `python3 …​.py` freely.

### By Document Type

**Word (.docx):**
```bash
cd word
bash run-formatting.sh         # Run/character formatting: bold/underline/strike/caps/super-sub/fonts/effects
bash paragraph-formatting.sh   # Paragraph formatting: align/indent/spacing/pagination/shading/markRPr
bash formulas.sh             # LaTeX math formulas
bash tables.sh               # Styled tables
bash textbox.sh              # Formatted text boxes
bash numbering.sh   # List/numbering styles
bash revisions.sh            # Tracked-change (revision) API — ins/del/format/move/cellChange
```

**Excel (.xlsx):**
```bash
cd excel
python cell-formatting.py    # Full cell property surface: fonts, fills, borders, number formats, formulas/links
bash charts.sh               # Master chart showcase (8 chart types in one workbook)
bash charts/charts-basic.sh  # Per-type high-level examples (any charts/charts-<type>.sh)
python charts/charts-line.py # Single-type example (any charts/charts-<type>.py)
python pivot-tables.py       # Pivot tables
```

**PowerPoint (.pptx):**
```bash
cd ppt
bash presentation.sh         # Morph transitions / full deck
bash animations.sh           # Animation effects
python video.py              # Video embedding
bash 3d-model.sh             # 3D model embedding
bash diagram.sh              # Mermaid diagrams — render=native (editable shapes) + render=image (PNG)
python charts/charts-column.py      # PowerPoint chart examples (any charts/charts-<type>.py)
bash tables/tables-basic.sh         # Tables — minimal create + populate
bash tables/tables-styled.sh        # 9 built-in styles + banding flags + rowHeight/name=
bash tables/tables-merged.sh        # gridSpan horizontal merge
bash tables/tables-borders.sh       # Per-side / per-cell borders
bash tables/tables-rows-cols.sh     # add row/column, per-row height, gridSpan + merge.down
bash tables/tables-financial.sh     # End-to-end financial deck
bash transitions/transitions-basic.sh        # cut/fade/dissolve/flash + 'none' clear
bash transitions/transitions-directional.sh  # push/wipe/cover/uncover × direction matrix
bash transitions/transitions-shapes.sh       # circle/diamond/wedge/wheel/zoom
bash transitions/transitions-bands.sh        # blinds/strips/split/checker
bash transitions/transitions-dynamic.sh      # 2010+ Exciting gallery (vortex/flip/...)
bash transitions/transitions-modern.sh       # 2013+ Exciting gallery (pageCurl/airplane/origami/...)
bash transitions/transitions-random.sh       # newsflash / random
bash transitions/transitions-timing.sh       # speed, duration, advanceTime, advanceClick
bash transitions/transitions-morph.sh        # 2016+ Morph tweening
bash shapes/shapes-basic.sh                  # geometries, fills, outlines, rotation, basic effects
bash shapes/shapes-connectors.sh             # straight/elbow/curve connectors + groups
bash shapes/shapes-effects.sh                # autoFit, flip, image fill, 3D, softEdge, links, zorder
bash shapes/shapes-typography.sh             # spacing, kern, case, RTL direction, font.cs, lang
bash textboxes/textboxes-basic.sh            # alignment, bullets, runs, per-script fonts
bash textboxes/textboxes-advanced.sh         # per-paragraph overrides, indents, per-run typography
python pictures/pictures-basic.py            # picture src/crop/rotation/links (needs Pillow)
```

---

## 📚 Documentation by Type

### 📄 [Word Examples →](word/)
- Run / character formatting — weight, underline variants, strike/dstrike, caps/smallCaps, super/subscript, color/size/highlight, per-script fonts, text effects, character spacing, language
- Paragraph formatting — alignment, indentation, spacing, pagination flags, paragraph-level run formatting, shading, paragraph-mark (markRPr) formatting, outline level
- Mathematical formulas (LaTeX)
- Complex tables
- Text boxes and styling
- Numbering / list showcases
- Mermaid diagrams — `render=native` (editable flowchart / sequence shapes) and `render=image` (inline full-fidelity PNG of every mermaid type)

### 📊 [Excel Examples →](excel/)
- Cell formatting — the full `cell` property surface across 5 sheets: fonts (name/size/bold/italic/color/underline/strike), fills (hex/named/rgb) + alignment (h/v/wrap/RTL), borders (shorthand/all/per-side/color), number formats (thousands/%/currency/date/scientific/accounting), and data (value/type/formula/link/locked/merge)
- Master and per-type chart scripts (line, bar, pie, scatter, stock, waterfall, …)
- Pivot tables
- Number formatting and styling

### 🎨 [PowerPoint Examples →](ppt/)
- Slide / shape construction
- Morph transitions and animations
- Video and 3D model embedding
- Mermaid diagrams — `render=native` (editable flowchart / sequence shapes + connectors) and `render=image` (full-fidelity PNG via mermaid.js, covering every mermaid type); `mermaid` / `text` / `dsl` / `src` source, placement box, `get` / `set` / `remove` the whole diagram as one group
- Native chart examples (column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo, waterfall, 3D, advanced)
- Tables — basic, built-in styles, merged cells, borders, row/column ops, real-world financial deck
- Slide transitions — all 59 schema tokens covered across 9 trios: basic, directional, shape, band, dynamic 3D (p14), modern (p15 — Page Curl, Airplane, Origami, …), random, timing, and Morph
- Shapes — full pptx/shape property surface across 4 trios: geometries + fills + outlines + rotation + basic effects (basic), straight/elbow/curve connectors + groups (connectors), autoFit + flip + image-fill + 3D scene + softEdge + click links + zorder (effects), paragraph/char spacing + kerning + smallCaps + RTL + complex-script font + BCP-47 lang (typography)
- Textboxes — alignment, bulleted/numbered lists, run-by-run rich text (bold/italic/color/super/sub/strike), per-script fonts (Latin/EastAsian), vertical alignment and padding
- Pictures — file path / URL / data-URI / `name=` for `src=`, all crop forms (symmetric, V,H, per-edge L/T/R/B), rotation, clickable links (URL / slide jump / named action)

---

## 🔧 Common Patterns

### Create and Populate

```bash
#!/bin/bash
set -e

FILE="document.docx"
officecli create "$FILE"
officecli add "$FILE" /body --type paragraph --prop text="Hello World"
officecli validate "$FILE"
```

### Batch Operations

```bash
cat << 'EOF' > commands.json
[
  {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Para 1"}},
  {"command":"set","path":"/body/p[1]","props":{"bold":"true","size":"24"}}
]
EOF
officecli batch document.docx < commands.json
```

### Resident Mode (3+ operations)

```bash
officecli open document.docx
officecli add document.docx /body --type paragraph --prop text="Fast operation"
officecli set document.docx /body/p[1] --prop bold=true
officecli close document.docx
```

### Query and Modify

```bash
# Find all Heading1 paragraphs
officecli query report.docx "paragraph[style=Heading1]" --json

# Change their color
officecli set report.docx /body/p[1] --prop color=FF0000
```

---

## 📊 Quick Reference

### Document Types

| Format | Extension | Create | View | Modify |
|--------|-----------|--------|------|--------|
| Word | .docx | ✓ | ✓ | ✓ |
| Excel | .xlsx | ✓ | ✓ | ✓ |
| PowerPoint | .pptx | ✓ | ✓ | ✓ |

### Common Commands

| Command | Purpose | Example |
|---------|---------|---------|
| `create` | Create blank document | `officecli create file.docx` |
| `view` | View content | `officecli view file.docx text` |
| `get` | Get element | `officecli get file.docx /body/p[1]` |
| `set` | Modify element | `officecli set file.docx /body/p[1] --prop bold=true` |
| `add` | Add element | `officecli add file.docx /body --type paragraph` |
| `remove` | Remove element | `officecli remove file.docx /body/p[5]` |
| `query` | CSS-like query | `officecli query file.docx "paragraph[style=Normal]"` |
| `batch` | Multiple operations | `officecli batch file.docx < commands.json` |
| `validate` | Check schema | `officecli validate file.docx` |

### View Modes

| Mode | Description | Usage |
|------|-------------|-------|
| `text` | Plain text | `officecli view file.docx text` |
| `annotated` | Text with formatting | `officecli view file.docx annotated` |
| `outline` | Structure | `officecli view file.docx outline` |
| `stats` | Statistics | `officecli view file.docx stats` |
| `issues` | Problems | `officecli view file.docx issues` |
| `html` | HTML preview | `officecli view file.docx html` |
| `svg` | SVG preview | `officecli view file.docx svg` |
| `forms` | Form fields | `officecli view file.docx forms` |

---

## 💡 Tips

1. **Explore before modifying:**
   ```bash
   officecli view document.docx outline
   officecli get document.docx /body --depth 2
   ```

2. **Use `--json` for automation:**
   ```bash
   officecli query data.xlsx "cell[formula~=SUM]" --json | jq
   ```

3. **Check help for properties** (schema reference is under the `help` verb):
   ```bash
   officecli help docx set paragraph
   officecli help xlsx set cell
   officecli help pptx set shape
   ```

4. **Validate after changes:**
   ```bash
   officecli validate document.docx
   ```

5. **Use resident mode for performance** (3+ operations on same file):
   ```bash
   officecli open file.pptx
   # ... multiple commands ...
   officecli close file.pptx
   ```

---

## 🤝 Contributing Examples

1. **Create script** with clear comments
2. **Test and verify** output
3. **Add to appropriate directory** (word/excel/ppt)
4. **Update directory README**
5. **Submit PR**

**Example format:**
```bash
#!/bin/bash
# Brief description of what this demonstrates
# Key techniques: list them here

set -e

FILE="output.docx"
officecli create "$FILE"
# ... your commands ...
officecli validate "$FILE"
echo "Created: $FILE"
```

---

## 📖 More Resources

- **[SKILL.md](../SKILL.md)** - Complete command reference for AI agents
- **[README.md](../README.md)** - Project overview and installation

---

## 🆘 Getting Help

**Top-level help:**
```bash
officecli --help                       # CLI usage
officecli help                         # Schema reference entry point
officecli help docx                    # All docx elements
officecli help docx set                # Elements that support `set` for docx
officecli help docx set paragraph      # Settable properties on paragraph
officecli help docx paragraph --json   # Raw schema JSON
officecli help all                     # Flat dump of every (format, element, property)
```

Format aliases: `word→docx`, `excel→xlsx`, `ppt`/`powerpoint→pptx`.
Verbs: `add`, `set`, `get`, `query`, `remove`.

---

**Happy automating! 🚀**

For questions or issues, visit [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues).
