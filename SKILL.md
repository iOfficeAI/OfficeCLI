---
name: officecli
description: Create, analyze, proofread, and modify Office documents (.docx, .xlsx, .pptx) using the officecli CLI tool. Use when the user wants to create, inspect, check formatting, find issues, add charts, or modify Office documents.
---

# officecli

AI-friendly CLI for .docx, .xlsx, .pptx.

**First, check if officecli is available:**
```bash
officecli --version
```
If the command is not found, install it:
```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
```
For Windows (PowerShell):
```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

**Strategy:** L1 (read) → L2 (DOM edit) → L3 (raw XML). Always prefer higher layers. Add `--json` for structured output.

**IMPORTANT: When unsure about property names, value formats, or command syntax, run the help command below instead of guessing. When a command fails, ALWAYS check help before retrying.** One help query is faster than guess-fail-retry loops.

**Help — three-layer navigation (start from the deepest level you know):**
```bash
officecli pptx set              # All settable elements and their properties
officecli pptx set shape        # Shape properties in detail
officecli pptx set shape.fill   # Specific property format and examples
```
Replace `pptx` with `docx` or `xlsx`. Commands: `view`, `get`, `query`, `set`, `add`, `raw`.

**Performance:** For multi-step workflows, use `open`/`close` to keep the document in memory:
```bash
officecli open report.docx       # keep in memory
officecli set report.docx ...    # fast — no file I/O
officecli close report.docx      # save and release
```

**Quick start — create a PPT from scratch:**
```bash
officecli create slides.pptx
officecli add slides.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add slides.pptx /slide[1] --type shape --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm --prop font=Arial --prop size=24 --prop color=FFFFFF
officecli set slides.pptx /slide[1] --prop transition=fade --prop advanceTime=3000
```

---

## L1: Create, Read & Inspect

```bash
officecli create <file>          # create blank .docx/.xlsx/.pptx (type inferred from extension)
officecli view <file> outline|stats|issues|text|annotated [--start N --end N] [--max-lines N] [--cols A,B]
officecli get <file> '/body/p[3]' --depth 2 [--json]
officecli query <file> 'paragraph[style=Normal] > run[font!=宋体]'
```

**get** supports any XML path via element localName. Use `--depth N` to expand children. Run `officecli docx get` / `officecli xlsx get` / `officecli pptx get` for all available paths.

**view modes:** `outline` (structure), `stats` (statistics), `issues` (`--type format|content|structure`, `--limit N`), `text` (plain), `annotated` (with formatting)

**query selectors:** `[attr=value]`, `[attr!=value]`, `[attr~=text]`, `[attr>=value]`, `[attr<=value]`, `:contains("text")`, `:empty`, `:has(formula)`, `:no-alt`. Run `officecli docx query` / `officecli pptx query` for all selector types.

For large documents, ALWAYS use `--max-lines` or `--start`/`--end` to limit output.

---

## L2: DOM Operations

### set — `officecli set <file> <path> --prop key=value [--prop ...]`

**Any XML attribute is settable via element path** (found via `get --depth N`) — even attributes not currently present. Use this before reaching for L3.

Run `officecli <format> set` for all settable elements and properties. Run `officecli <format> set <element>` for detail (e.g. `officecli pptx set shape`, `officecli docx set paragraph`).

Colors: hex RGB (`FF0000`, `#FF0000`), named colors (`red`, `blue`), `rgb(255,0,0)`, or theme names (`accent1`..`accent6`, `dk1`, `dk2`, `lt1`, `lt2`)

Spacing: unit-qualified — `12pt`, `0.5cm`, `1.5x` (multiplier), `150%`, `18pt` (fixed)

Dimensions: raw EMU or suffixed `cm`/`in`/`pt`/`px`

### add — `officecli add <file> <parent> --type <type> [--index N] [--prop ...]` or `--from <path>`

Run `officecli <format> add` for all addable element types and properties.

**Copy from existing:** `officecli add <file> <parent> --from <path> [--index N]` — clones the element. Cross-part relationships handled automatically. Either `--type` or `--from` is required.

**Clone entire slide:** `officecli add <file> / --from /slide[1] [--index 0]`

### move — `officecli move <file> <path> [--to <parent>] [--index N]`

### swap — `officecli swap <file> <path1> <path2>`

### remove — `officecli remove <file> '/body/p[4]'`

### batch — For 3+ mutations (one open/save cycle)

```bash
echo '[{"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
      {"command":"set","path":"/Sheet1/B1","props":{"value":"Score","bold":"true"}}]' | officecli batch data.xlsx --json
```

Batch fields: `command`(add/set/get/query/remove/move/view/raw/raw-set/validate), `path`, `parent`, `type`, `from`, `to`, `index`, `props`(dict), `selector`, `mode`, `depth`, `part`, `xpath`, `action`, `xml`.

---

## L3: Raw XML

Use when L2 cannot express what you need. No xmlns declarations needed — prefixes auto-registered.

```bash
officecli raw <file> /document           # view raw XML (Word/Excel/PPT parts vary — run officecli <format> raw for details)
officecli raw-set <file> /document --xpath "//w:body/w:p[1]" --action replace --xml '<w:p>...</w:p>'
# actions: append, prepend, insertbefore, insertafter, replace, remove, setattr
```

---

## Notes

- Paths are **1-based** (XPath convention), quote brackets: `'/body/p[3]'`
- `--index` is **0-based** (array convention): `--index 0` = first position
- After modifications, verify with `validate` and/or `view issues`
- **When unsure about any property or format**, run `officecli <format> <command> [element[.property]]` instead of guessing. Example: `officecli pptx set chart` shows all chart properties and accepted values
