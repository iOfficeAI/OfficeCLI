---
name: officecli
description: Create, analyze, proofread, and modify Office documents (.docx, .xlsx, .pptx) using the officecli CLI tool. Use when the user wants to create, inspect, check formatting, find issues, add charts, or modify Office documents.
---

# officecli

AI-friendly CLI for .docx, .xlsx, .pptx. Single binary, no dependencies, no Office installation needed.

## Install

If `officecli` is not installed:

`macOS / Linux`

```bash
if ! command -v officecli >/dev/null 2>&1; then
    curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
fi
```

`Windows (PowerShell)`

```powershell
if (-not (Get-Command officecli -ErrorAction SilentlyContinue)) {
    irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
}
```

Verify: `officecli --version`

If `officecli` is still not found after first install, open a new terminal and run the verify command again.

---

## Strategy

**L1 (read) → L2 (DOM edit) → L3 (raw XML)**. Always prefer higher layers. Add `--json` for structured output.

---

## Help System (IMPORTANT)

**When unsure about property names, value formats, or command syntax, ALWAYS run help instead of guessing.** One help query is faster than guess-fail-retry loops.

**Three-layer navigation** — start from the deepest level you know:
```bash
officecli pptx set              # All settable elements and their properties
officecli pptx set shape        # Shape properties in detail
officecli pptx set shape.fill   # Specific property format and examples
```

Replace `pptx` with `docx` or `xlsx`. Commands: `view`, `get`, `query`, `set`, `add`, `raw`.

---

## Performance: Resident Mode

**Every command auto-starts a resident on first access** (60s idle timeout) — file-lock conflicts are automatically avoided. Explicit `open`/`close` is still recommended for longer sessions (12min idle):
```bash
officecli open report.docx       # explicitly keep in memory
officecli set report.docx ...    # no file I/O overhead
officecli close report.docx      # save and release
```

Opt out of auto-start: `OFFICECLI_NO_AUTO_RESIDENT=1`. Skipping `close` still works (resident exits on idle), but explicit `close` guarantees the file is flushed before the next command reads it.

---

## Quick Start

**PPT:**
```bash
officecli create slides.pptx
officecli add slides.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add slides.pptx '/slide[1]' --type shape --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm --prop font=Arial --prop size=24 --prop color=FFFFFF
```

**Word:**
```bash
officecli create report.docx
officecli add report.docx /body --type paragraph --prop text="Executive Summary" --prop style=Heading1
officecli add report.docx /body --type paragraph --prop text="Revenue increased by 25% year-over-year."
```

**Excel:**
```bash
officecli create data.xlsx
officecli set data.xlsx /Sheet1/A1 --prop value="Name" --prop bold=true
officecli set data.xlsx /Sheet1/A2 --prop value="Alice"
```

---

## L1: Create, Read & Inspect

```bash
officecli create <file>               # Create blank .docx/.xlsx/.pptx (type from extension)
officecli view <file> <mode>          # outline | stats | issues | text | annotated
officecli get <file> <path> --depth N # Get a node and its children [--json]
officecli query <file> <selector>     # CSS-like query
officecli validate <file>             # Validate against OpenXML schema
```

### view modes

| Mode | Description | Useful flags |
|------|-------------|-------------|
| `outline` | Document structure | |
| `stats` | Statistics (pages, words, shapes) | |
| `issues` | Formatting/content/structure problems | `--type format\|content\|structure`, `--limit N` |
| `text` | Plain text extraction | `--start N --end N`, `--max-lines N` |
| `annotated` | Text with formatting annotations | |
| `html` | Static HTML snapshot (.docx/.xlsx/.pptx) — writes to stdout | `--browser` (open in default browser), `--page N` (docx), `--start N --end N` (pptx slide range) |

**`view html` vs `watch`** — both render the same HTML (shared `*.HtmlPreview.cs` renderer). Use `view html` for one-shot snapshots (CI artifacts, archival, diffing, piping to files); use `watch` when you need live refresh or browser-side click-to-select. `view html` needs no server/port.

```bash
officecli view report.docx html > snapshot.html       # snapshot to file
officecli view report.docx html --browser             # open in default browser
```

### get

Any XML path via element localName. Use `--depth N` to expand children. Add `--json` for structured output. Default text output is grep-friendly single-line per node: `path (type) "text" key=val key=val ...`

```bash
officecli get report.docx '/body/p[3]' --depth 2 --json
officecli get slides.pptx '/slide[1]' --depth 1          # list all shapes on slide 1
officecli get data.xlsx '/Sheet1/B2' --json
```

Run `officecli docx get` / `officecli xlsx get` / `officecli pptx get` for all available paths.

### Stable ID Addressing

Elements with stable IDs return `@attr=value` paths instead of positional indices. Prefer these in multi-step workflows — positional indices shift on insert/delete, stable IDs do not.

```
/slide[1]/shape[@id=550950021]                    # PPT shape
/slide[1]/table[@id=1388430425]/tr[1]/tc[2]       # PPT table
/body/p[@paraId=1A2B3C4D]                         # Word paragraph
/comments/comment[@commentId=1]                    # Word comment
```

Use returned paths directly for subsequent `set`/`remove`. PPT also accepts `@name=` (e.g. `shape[@name=Title 1]`), with morph `!!` prefix awareness. Elements without stable IDs (slide, run, tr/tc, row) fall back to positional indices. Run `officecli <format> get` for the full list.

### query

CSS-like selectors: `[attr=value]`, `[attr!=value]`, `[attr~=text]`, `[attr>=value]`, `[attr<=value]`, `:contains("text")`, `:empty`, `:has(formula)`, `:no-alt`.

```bash
officecli query report.docx 'paragraph[style=Normal] > run[font!=Arial]'
officecli query slides.pptx 'shape[fill=FF0000]'
```

### validate

```bash
officecli validate report.docx    # Check for schema errors
officecli validate slides.pptx    # Must pass before delivery
```

**For large documents**, ALWAYS use `--max-lines` or `--start`/`--end` to limit output.

---

## Watch & Interactive Selection

Live HTML preview that auto-refreshes on every file change. Browsers can click / shift-click / box-drag to select shapes; the CLI can read the current browser selection and act on it.

```bash
officecli watch <file> [--port N]      # Start preview server (default port 18080)
officecli unwatch <file>               # Stop the preview server
```

Open the printed `http://localhost:N` URL in a browser. Click any element to select; shift/cmd/ctrl+click to multi-select; drag from empty space to box-select (rubber-band). PPT/Word uses blue outline; Excel uses native-style green selection with crosshair and rectangular range selection. **Excel extras:** double-click a cell to edit inline (shows formula, commits on Enter/Tab); drag a chart to reposition it.

### `get <file> selected` — read what the user clicked

```bash
officecli get <file> selected [--json]
```

Returns the DocumentNodes for whatever is currently selected in the watching browser(s). Empty result if nothing selected. Exit code != 0 if no watch is running for this file.

**Workflow** — agent acts on what the user visually selected:

```bash
# User clicks shapes in the browser, then asks "make these red"
PATHS=$(officecli get deck.pptx selected --json | jq -r '.data.Results[].path')
for p in $PATHS; do
  officecli set deck.pptx "$p" --prop fill=FF0000
done
```

### Key properties

- **Selection survives file edits.** Paths use the stable `@id=` form (e.g. `/slide[1]/shape[@id=10000]`), so editing other shapes — or even the selected one — does not lose the selection.
- **All connected browsers share one selection.** Opening the watch URL in two tabs gives a shared cursor; clicking in one updates highlights in the other. Last-write-wins.
- **Same-file single-watch.** A given file can have only one watch process at a time; the second `watch <file>` errors.
- **Group shapes select as a whole.** Clicking any shape inside a `<group>` selects the group container, not the inner shape. The CLI sees `/slide[1]/group[@id=N]`. Drilling into individual children of a group is not supported in v1.
- **PPT and top-level Word.** Selection / mark works on `.pptx` shapes, pictures, tables, charts, connectors, groups, and on `.docx` top-level paragraphs (`<p>`/`<h1-6>`/`<li>`/`.empty`) and top-level `<table>`. Inherited layout/master decorations (footers, logos) and Word nested elements (table cells, run-level) are not addressable. **Excel `.xlsx` does not emit `data-path`** — `mark`/`selection` on xlsx will always resolve to `stale=true`. Excel support is a v2 candidate.

## Marks — edit proposals waiting for review

**Marks are edit proposals waiting for review.** Use `mark` when you (or the user) want to see, evaluate, and approve changes BEFORE they hit the file. Marks live in the watch process only — nothing is written to disk until a separate `set` pipeline applies them.

**Decision tree — pick one:**

- User doesn't need to confirm? → **`set`** directly (straight to disk). Marks are overkill for one-shot changes.
- User wants to review before changes apply? → **`mark`** (propose → review → `set` → mark goes stale).
- Just leaving a permanent annotation in the file? → **`add --type comment`** (Word native, persists in file).

**Four-step lifecycle:**

1. **Propose** — agent scans and creates marks with `find` + `tofix` + `note`.
2. **Review** — human opens the watch URL, sees highlights, decides what to accept.
3. **Apply** — a pipeline reads `get-marks --json` and runs real `set` commands for accepted items.
4. **Stale** — after the underlying text changes, the mark's `find` no longer matches; `stale=true` signals "this proposal has been handled".

```bash
officecli mark <file> <path> [--prop find=...] [--prop color=...] [--prop note=...] [--prop tofix=...] [--prop regex=true] [--json]
officecli unmark <file> [--path <p> | --all] [--json]
officecli get-marks <file> [--json]
```

| Prop | Meaning |
|------|---------|
| `find` | Literal text to highlight (or regex when `regex=true`; raw form `find='r"[abc]"'` also accepted). 500ms match timeout. |
| `color` | CSS color from whitelist: hex, `rgb(...)`, or one of 22 named colors. Invalid rejected. |
| `note` | Free-form reviewer comment. |
| `tofix` | Structured proposed replacement value (drives the apply pipeline). |
| `regex` | `true` to switch `find` to regex. |

**Path** must be `data-path` format from watch HTML: Word `/body/p[N]` or `/body/table[N]`; PPT `/slide[N]/shape[@id=ID]` (preferred) or `/slide[N]/shape[N]`. Excel is not supported in v1 (marks always resolve `stale=true`). Native query paths like `/body/p[@paraId=...]` will NOT resolve.

**Worked example:**

```bash
officecli watch report.docx &
# Propose
officecli mark report.docx /body/p[3] --prop find="资钱" --prop tofix="资金" --prop color=red --prop note="术语错误"
# Review — human eyeballs highlights in browser, unmarks bad proposals
# Apply — read accepted marks, run real set commands
officecli get-marks report.docx --json \
  | jq -r '(.marks // []) | .[] | select(.tofix != null) | [.path, .find, .tofix] | @tsv' \
  | while IFS=$'\t' read -r path find tofix; do
      officecli set report.docx "$path" --prop "find=$find" --prop "replace=$tofix"
    done
```

All mark commands support `--json`. For >3 mutations, wrap the apply loop in `batch` or `open`/`close` for performance.

---

## L2: DOM Operations

### set — modify properties

```bash
officecli set <file> <path> --prop key=value [--prop ...]
```

**Any XML attribute is settable** via element path (found via `get --depth N`) — even attributes not currently present.

Without `find=`, `set` applies format to the entire element. To target specific text within a paragraph, use `find=` (see **find** section below).

Run `officecli <format> set` for all settable elements. Run `officecli <format> set <element>` for detail.

**Value formats:**

| Type | Format | Examples |
|------|--------|---------|
| Colors | Hex, named, RGB, theme | `FF0000`, `red`, `rgb(255,0,0)`, `accent1`..`accent6` |
| Spacing | Unit-qualified | `12pt`, `0.5cm`, `1.5x`, `150%` |
| Dimensions | EMU or suffixed | `914400`, `2.54cm`, `1in`, `72pt`, `96px` |

### find — format or replace matched text

Use `find=` with `set` to target specific text for formatting or replacement. Works the same in Word and PPT — just swap paths. Format props are separate `--prop` flags — do NOT nest them.

```bash
# Format matched text (auto-splits runs)
officecli set doc.docx '/body/p[1]' --prop find=weather --prop bold=true --prop color=red

# Regex matching
officecli set doc.docx '/body/p[1]' --prop 'find=\d+%' --prop regex=true --prop color=red

# Replace text (use `/` for whole-document scope)
officecli set doc.docx / --prop find=draft --prop replace=final

# Replace + format
officecli set doc.docx '/body/p[1]' --prop find=TODO --prop replace=DONE --prop bold=true

# PPT — same syntax, different paths
officecli set slides.pptx / --prop find=draft --prop replace=final
```

**Path controls search scope:** `/` = whole document, `/body/p[1]` or `/slide[N]/shape[M]` = specific element, `/header[1]` = header, `/footer[1]` = footer.

**Notes:**
- Case-sensitive by default. Case-insensitive: `--prop 'find=(?i)error' --prop regex=true`
- Matches work across run boundaries
- No match = no error (silent success). `--json` includes `"matched": N`
- Batch JSON regex: `{"props":{"find":"\\d+%","regex":"true","color":"FF0000"}}`
- **Excel:** only `find` + `replace` supported (no find + format props)

### add — add elements or clone

```bash
officecli add <file> <parent> --type <type> [--prop ...]
officecli add <file> <parent> --type <type> --after <path> [--prop ...]   # insert after anchor
officecli add <file> <parent> --type <type> --before <path> [--prop ...]  # insert before anchor
officecli add <file> <parent> --type <type> --index N [--prop ...]        # insert at position (0-based, legacy)
officecli add <file> <parent> --from <path>                               # clone existing element
```

`--after`, `--before`, `--index` are mutually exclusive. No position flag = append to end.

**Element types (with aliases):**

| Format | Types |
|--------|-------|
| **pptx** | slide, shape (textbox), picture (image/img), chart, table, row (tr), connector (connection/line), group, video (audio/media), equation (formula/math), notes, paragraph (para), run, zoom (slidezoom), ole (oleobject/object/embed) |
| **docx** | paragraph (para), run, table, row (tr), cell (td), image (picture/img), header, footer, section, bookmark, comment, footnote, endnote, formfield (text/checkbox/dropdown), sdt (contentcontrol), chart, equation (formula/math), field, hyperlink, style, toc, watermark, break (pagebreak/columnbreak), ole (oleobject/object/embed). Document protection: `set / --prop protection=forms\|readOnly\|comments\|trackedChanges\|none` |
| **xlsx** | sheet, row, cell, chart, image (picture), comment, table (listobject), namedrange (definedname), pivottable (pivot), sparkline, validation (datavalidation), autofilter, shape, textbox, databar/colorscale/iconset/formulacf (conditional formatting), ole (oleobject/object/embed — no Remove yet), csv (tsv). `value="=SUM(...)"` auto-detects as formula. Formulas auto-evaluated on write (150+ functions including VLOOKUP, SUMIF, IF, DATE, PMT, etc.) |

### Pivot tables (xlsx)

```bash
officecli add data.xlsx /Sheet1 --type pivottable \
  --prop source="Sheet1!A1:E100" --prop rows=Region,Category \
  --prop cols=Year --prop values="Sales:sum,Qty:count" \
  --prop grandTotals=rows --prop subtotals=off --prop sort=asc
```

Key props: `rows`, `cols`, `values` (Field:func[:showDataAs]), `filters`, `source`, `position`, `layout` (compact/outline/tabular), `repeatLabels` (true/false — repeat outer row labels on every data row), `blankRows` (true/false — insert blank line after each group), `aggregate`, `showDataAs` (percent_of_total/row/col, running_total), `grandTotals` (both/rows/cols/none), `subtotals` (on/off), `sort` (asc/desc/locale/locale-desc). Aggregators: sum, count, average, max, min, product, stdDev, stdDevp, var, varp, countNums. Date columns auto-group. Multiple data fields and N×N row/col hierarchies supported. Run `officecli xlsx set pivottable` for full property list.

### Document-level properties (all formats)

```bash
officecli set doc.docx / --prop docDefaults.font=Arial --prop docDefaults.fontSize=11pt
officecli set doc.docx / --prop protection=forms --prop evenAndOddHeaders=true
officecli set data.xlsx / --prop calc.mode=manual --prop calc.refMode=r1c1
officecli set slides.pptx / --prop defaultFont=Arial --prop show.loop=true --prop print.what=handouts
```

Run `officecli <format> set /` for all available document-level properties (docDefaults, docGrid, CJK spacing, calc, print, show, theme, extended).

**Text-anchored insert** (`--after find:X` / `--before find:X`):

The `--after` and `--before` flags accept a `find:` prefix to locate an insertion point by text match within a paragraph.

```bash
# Insert run after matched text (inline, within the same paragraph)
officecli add doc.docx '/body/p[1]' --type run --after find:weather --prop text=" (sunny)"

# Insert table after matched text (block — auto-splits the paragraph)
officecli add doc.docx '/body/p[1]' --type table --after "find:First sentence." --prop rows=2 --prop cols=2

# Insert before matched text
officecli add doc.docx '/body/p[1]' --type run --before find:weather --prop text="["

```

- Inline types (run, picture, hyperlink...) insert within the paragraph
- Block types (table, paragraph) auto-split the paragraph and insert between the two halves

**PPT text-anchored insert** — same as Word, but PPT only supports **inline** types (`run`); block-type insertion is not supported.

```bash
officecli add slides.pptx '/slide[1]/shape[1]' --type run --after find:weather --prop text=" (sunny)"
```

**Clone:** `officecli add <file> / --from '/slide[1]'` — copies with all cross-part relationships.

Run `officecli <format> add` for all addable types and their properties.

### move, swap, remove

```bash
officecli move <file> <path> [--to <parent>] [--index N] [--after <path>] [--before <path>]
officecli swap <file> <path1> <path2>
officecli remove <file> '/body/p[4]'
```

When using `--after` or `--before`, `--to` can be omitted — the target container is inferred from the anchor path.

### batch — multiple operations in one save cycle

Stops on first error by default. Use `--force` to continue past errors.

```bash
# Via stdin
echo '[
  {"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
  {"command":"set","path":"/Sheet1/B1","props":{"value":"Score","bold":"true"}}
]' | officecli batch data.xlsx --json

# Via --commands (inline) or --input (file)
officecli batch data.xlsx --commands '[{"op":"set","path":"/Sheet1/A1","props":{"value":"Done"}}]' --json
officecli batch data.xlsx --input updates.json --force --json
```

Batch supports: `add`, `set`, `get`, `query`, `remove`, `move`, `swap`, `view`, `raw`, `raw-set`, `validate`. Fields: `command` (or `op`), `path`, `parent`, `type`, `from`, `to`, `index`, `after`, `before`, `props`, `selector`, `mode`, `depth`, `part`, `xpath`, `action`, `xml`.

---

## L3: Raw XML

Use when L2 cannot express what you need. No xmlns declarations needed — prefixes auto-registered.

```bash
officecli raw <file> <part>                          # view raw XML
officecli raw-set <file> <part> --xpath "..." --action replace --xml '<w:p>...</w:p>'
officecli add-part <file> <parent>                   # create new document part (returns rId)
```

**raw-set actions:** `append`, `prepend`, `insertbefore`, `insertafter`, `replace`, `remove`, `setattr`.

Run `officecli <format> raw` for available parts per format.

---

## Common Pitfalls

| Pitfall | Correct Approach |
|---------|-----------------|
| `--name "foo"` | ❌ Use `--prop name="foo"` — all attributes go through `--prop` |
| `x=-3cm` | ❌ Negative coordinates not supported. Use `x=0cm` or `x=36cm` |
| PPT `shape[1]` for content | ❌ `shape[1]` is typically the title placeholder. Use `shape[2]` or higher for content shapes |
| `/shape[myname]` | ❌ Name indexing not supported. Use numeric index: `/shape[3]` |
| Guessing property names | ❌ Run `officecli <format> set <element>` to see exact names |
| Modifying an open file | ❌ Close the file in PowerPoint/WPS first |
| `\n` in shell strings | ❌ Use `\\n` for newlines in `--prop text="..."` |
| `officecli set f.pptx /slide[1]` | ❌ Shell glob expands brackets. Always single-quote paths: `'/slide[1]'` |

---

## Specialized Skills

For complex scenarios, load the dedicated skill from `skills/<skill-name>/SKILL.md`:

| Skill | Scope |
|-------|-------|
| `officecli-hwpx` | HWPX / Hancom — Korean government/OWPML documents |
| `officecli-docx` | Word documents — reports, letters, memos |
| `officecli-academic-paper` | Academic papers with TOC, equations, footnotes, bibliography |
| `officecli-pptx` | Presentations — general slide decks |
| `officecli-pitch-deck` | Investor/product/sales decks with charts and callouts |
| `morph-ppt` | Morph-animated cinematic presentations |
| `officecli-xlsx` | Excel — financial models, trackers, formulas |
| `officecli-data-dashboard` | CSV/tabular data → Excel dashboards with charts, sparklines |

---

## Notes

- Paths are **1-based** (XPath convention): `'/body/p[3]'` = third paragraph
- `--index` is **0-based** (array convention): `--index 0` = first position
- After modifications, verify with `validate` and/or `view issues`
- **When unsure**, run `officecli <format> <command> [element[.property]]` instead of guessing
