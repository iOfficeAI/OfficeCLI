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
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCli/main/install.sh | bash
```
For Windows (PowerShell):
```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCli/main/install.ps1 | iex
```

**Strategy:** L1 (read) → L2 (DOM edit) → L3 (raw XML). Always prefer higher layers. Add `--json` for structured output.

**Performance:** Use `open <file>`/`close <file>` for interactive sessions, or `batch` for scripted multi-operation workflows.

**Batch:** For 3+ mutations, use `batch` (one open/save cycle). Pipe JSON array via stdin or `--input file.json`. Add `--json` for structured output, `--stop-on-error` to abort on failure.

```bash
echo '[{"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
      {"command":"add","parent":"/","type":"slide","props":{"title":"Hello"}},
      {"command":"remove","path":"/body/p[2]"}]' | officecli batch doc.xlsx --json
```

Batch fields: `command`(add/set/get/query/remove/move/view/raw/raw-set/validate), `path`, `parent`, `type`, `from`, `to`, `index`, `props`(dict), `selector`, `mode`, `depth`, `part`, `xpath`, `action`, `xml`.

**Help:** If unsure about usage, run `officecli <format> <command>` for detailed help (e.g. `officecli pptx add`, `officecli docx set`, `officecli xlsx get`).

---

## L1: Create, Read & Inspect

```bash
officecli create <file>          # create blank .docx/.xlsx/.pptx (type inferred from extension)
officecli view <file> outline|stats|issues|text|annotated [--start N --end N] [--max-lines N] [--cols A,B]
officecli get <file> '/body/p[3]' --depth 2 [--json]
officecli query <file> 'paragraph[style=Normal] > run[font!=宋体]'
```

**get** supports any XML path via element localName: `/body/tbl[1]/tblPr`, `/Sheet1/sheetViews/sheetView[1]`, `/slide[1]/cSld/spTree/sp[1]/nvSpPr`. Use `--depth N` to expand children. Word also supports: `/` (core properties), `/footnote[N]`, `/endnote[N]`, `/toc[N]`, `/section[N]`, `/styles/StyleId`, `/chart[N]` (N = id returned by add). Excel also supports: `/SheetName/chart[N]`.

**view modes:** `outline` (structure), `stats` (statistics with style inheritance), `issues` (`--type format|content|structure`, `--limit N`), `text` (plain with element paths), `annotated` (with element paths and formatting)

**query selectors:** `[attr=value]`, `[attr!=value]`, `:contains("text")`, `:empty`, `:has(formula)`, `:no-alt`. Built-in types: `paragraph`, `run`, `picture`, `equation`, `cell`, `table`, `chart`, `bookmark`. Falls back to generic XML element name (e.g. `wsp`, `a:ln`, `srgbClr[val=0070C0]`).

For large documents, ALWAYS use `--max-lines` or `--start`/`--end` to limit output.

---

## L2: DOM Operations

### set — `officecli set <file> <path> --prop key=value [--prop ...]`

The table below lists shortcut properties for common paths. Word run/paragraph/table props also accept any valid OpenXML child element name (validated via SDK type system).

**Any XML attribute is settable via element path:** `set` also works on **any** XML element path (found via `get --depth N`) with **any** XML attribute name — even attributes not currently present on the element. Use this before reaching for L3.

Examples (not exhaustive — shortcut properties from the table below and any XML attribute are all settable):

```bash
# Example: set PPT shape position and size via element path
officecli get doc.pptx '/slide[1]/cSld/spTree/sp[1]/spPr' --depth 3
officecli set doc.pptx '/slide[1]/cSld/spTree/sp[1]/spPr/xfrm[1]/off[1]' --prop x=1500000 --prop y=300000
officecli set doc.pptx '/slide[1]/cSld/spTree/sp[1]/spPr/xfrm[1]/ext[1]' --prop cx=9192000 --prop cy=900000
# Example: set PPT text color (simple)
officecli set doc.pptx '/slide[1]/shape[1]' --prop color=FFFFFF
# Example: set PPT text color via element path (when you need per-run control)
officecli set doc.pptx '/slide[1]/cSld/spTree/sp[1]/txBody/p[1]/r[1]/rPr[1]/solidFill[1]/srgbClr[1]' --prop val=FFFFFF
```

| Target | Path example | Properties |
|--------|-------------|------------|
| Word run | `/body/p[3]/r[1]` | `text`, `font`, `size`, `bold`, `italic`, `caps`, `smallCaps`, `superscript`, `subscript`, `strike`, `dstrike`, `vanish`, `outline`, `shadow`, `emboss`, `imprint`, `noProof`, `rtl`, `highlight`, `color`, `underline`, `shd`, ... |
| Word run image | `/body/p[5]/r[1]` | `alt`, `width`, `height` (cm/in/pt/px), ... |
| Word paragraph | `/body/p[3]` | `style`, `alignment`, `firstLineIndent`, `leftIndent`, `rightIndent`, `hangingIndent`, `shd`, `spaceBefore`, `spaceAfter`, `lineSpacing`, `numId`, `numLevel`/`ilvl`, `listStyle`(=bullet\|numbered\|none), `start`, `keepNext`, `keepLines`, `pageBreakBefore`, `widowControl`, ... |
| Word table cell | `/body/tbl[1]/tr[1]/tc[1]` | `text`, `font`, `size`, `bold`, `italic`, `color`, `shd`, `alignment`, `valign`(top\|center\|bottom), `width`, `vmerge`, `gridspan`, ... |
| Word table row | `/body/tbl[1]/tr[1]` | `height`, `header`(bool), ... |
| Word table | `/body/tbl[1]` | `alignment`, `width`, ... |
| Word document | `/` | `defaultFont`, `pageBackground`, `pageWidth`, `pageHeight`, `marginTop/Bottom/Left/Right`, `title`, `author`, `subject`, `keywords`, `description`, `category`, ... |
| Word footnote | `/footnote[N]` | `text` |
| Word endnote | `/endnote[N]` | `text` |
| Word TOC | `/toc[N]` | `levels`, `hyperlinks`(bool), `pagenumbers`(bool) |
| Word section | `/section[N]` | `type`(nextPage\|continuous\|evenPage\|oddPage), `pagewidth`, `pageheight`, `orientation`, `marginTop/Bottom/Left/Right` |
| Word style | `/styles/StyleId` | `name`, `basedon`, `next`, `font`, `size`, `bold`, `italic`, `color`, `alignment`, `spacebefore`, `spaceafter` |
| Word watermark | `/watermark` | `text`, `color`, `font`, `opacity`, `rotation` (add replaces existing; one per document) |
| Word chart | `/chart[N]` | `title`, `legend`, `categories`, `data`, `series1..N`, `colors`, `dataLabels`, `axisTitle`, `catTitle`, `axisMin`, `axisMax`, `majorUnit`, `axisNumFmt` |
| Excel cell | `/Sheet1/A1` | `value`, `formula`, `clear`, `link`, `font.bold/italic/strike/underline/color/size/name`, `fill`(hex), `border.all/left/right/top/bottom`(thin\|medium\|thick\|double\|none), `border.color`, `alignment.horizontal/vertical/wrapText`, `numFmt`, ... |
| Excel merge | `/Sheet1/A1:D1` | `merge`(bool) |
| Excel column | `/Sheet1/col[A]` | `width`, `hidden`(bool) |
| Excel row | `/Sheet1/row[1]` | `height`(pt), `hidden`(bool) |
| Excel sheet | `/Sheet1` | `freeze`(cell ref, e.g. A2) |
| Excel autofilter | `/Sheet1/autofilter` | `range`(e.g. A1:F100) |
| Excel chart | `/Sheet1/chart[N]` | `title`, `legend`, `categories`, `data`, `series1..N`, `colors`, `dataLabels`, `axisTitle`, `catTitle`, `axisMin`, `axisMax`, `majorUnit`, `axisNumFmt` |
| PPT shape | `/slide[1]/shape[1]` | `name`(rename shape, use `!!` prefix for morph matching), `text`, `font`, `size`, `bold`, `italic`, `color`, `fill`, `gradient`(linear/radial), `image`(blipFill), `line`, `lineWidth`, `lineDash`, `lineOpacity`, `opacity`, `shadow`, `glow`, `reflection`, `softEdge`(pt), `textFill`(gradient on text, e.g. FF0000-0000FF-90), `spacing`(char spacing pt), `indent`(para indent), `marginLeft`/`marginRight`, `baseline`(superscript/subscript %), `superscript`/`subscript`(bool), `flipH`/`flipV`(bool), `zorder`(front/back/forward/backward/N), `rot3d`(rotX,rotY,rotZ degrees), `bevel`/`bevelBottom`(preset-w-h), `depth`(extrusion pt), `material`, `lighting`, `geometry`(SVG-like: "M x,y L x,y C x1,y1 x2,y2 x,y Z"), ... |
| PPT slide | `/slide[N]` | `background`, `transition`(fade\|push\|wipe\|morph\|morph-byWord\|morph-byChar\|...), `advanceTime`(ms), `advanceClick`(bool), `notes`. **Morph:** matches shapes by `name` across slides. When `transition=morph` is set, shape names on the current and previous slide are automatically prefixed with `!!` to force matching even when text changes. No manual `!!` needed. |
| PPT paragraph | `/slide[1]/shape[1]/paragraph[1]` | `align`, `indent`, `marginLeft`, `marginRight`, `lineSpacing`, `spaceBefore`, `spaceAfter`, plus run-level props |
| PPT run | `/slide[1]/shape[1]/paragraph[1]/run[1]` | `text`, `font`, `size`, `bold`, `italic`, `color`, `spacing`, `baseline`, `textFill` |
| PPT chart | `/slide[1]/chart[1]` | `title`, `legend`, `categories`, `data`, `series1..N`, `colors`, `dataLabels`, `axisTitle`, `catTitle`, `axisMin`, `axisMax`, `majorUnit`, `axisNumFmt` |
| PPT video/audio | `/slide[1]/video[1]` | `volume`(0-100), `autoplay`(bool), `trimStart`(ms), `trimEnd`(ms), `x`, `y`, `width`, `height` |
| PPT picture | `/slide[1]/picture[1]` | `alt`, `path`(replace image), `crop`, `cropLeft/Top/Right/Bottom`, `x`, `y`, `width`, `height` |
| PPT table | `/slide[1]/table[1]` | `tableStyle`(medium1..4\|light1..3\|dark1..2\|none), `x`, `y`, `width`, `height` |
| PPT presentation | `/` | `slideSize`(16:9\|4:3\|16:10\|a4), `slideWidth`, `slideHeight` |

Colors: hex RGB (`FF0000`) or theme names (`accent1`..`accent6`, `dk1`, `dk2`, `lt1`, `lt2`, `tx1`, `tx2`, `bg1`, `bg2`, `hyperlink`)

Composite props (`pBdr`, `tabs`, `lang`, `bdr`) → use L3 (`raw-set --action setattr`).

### add — `officecli add <file> <parent> --type <type> [--index N] [--prop ...]` or `--from <path>`

Props listed are common examples, not exhaustive — most `set` shortcut properties also work with `add`:

| Format | Types & props |
|--------|--------------|
| Word | `paragraph`(text,font,size,bold,style,alignment,keepNext,keepLines,...), `run`(text,font,size,bold,italic,superscript,subscript,...), `table`(rows,cols), `row`(cols,c1,c2,...), `cell`(text,width), `picture`(path,width,height,alt,...), `chart`(chartType,title,categories,data/series1..N,legend,colors,width,height), `equation`(formula,mode), `comment`(text,author,...), `section`(type,orientation,...), `footnote`(text), `endnote`(text), `toc`(levels,title,...), `style`(name,id,font,size,bold,...) |
| Excel | `sheet`(name), `row`(cols), `cell`(ref,value,formula,...), `autofilter`(range), `databar`(sqref,min,max,color), `colorscale`(sqref,mincolor,maxcolor,midcolor), `iconset`(sqref,iconset,reverse), `formulacf`(sqref,formula,fill), `chart`(chartType,title,categories,data/series1..N,legend,...) |
| PPT | `slide`(title,text,layout,background,...), `shape`(text,font,size,name,fill,gradient,preset,geometry,textFill,spacing,indent,softEdge,bevel,depth,rot3d,...), `paragraph`(text,align,indent,marginLeft,bold,color,...), `run`(text,font,size,bold,italic,color,spacing,baseline/superscript/subscript,textFill,...), `chart`(chartType,title,categories,data/series1..N,legend,colors,...), `video`/`audio`(path,poster,volume,autoplay,trimStart,trimEnd,...), `connector`(preset,line,...), `group`(shapes=1,2,3), `picture`(path,width,height,x,y,...), `equation`(formula) |

Dimensions: raw EMU or suffixed `cm`/`in`/`pt`/`px`. Equation formula: LaTeX subset (`\frac{}{}`, `\sqrt{}`, `^{}`, `_{}`, `\sum`, Greek letters). Mode: `display`(default) or `inline`. Comment parent can be a paragraph (`/body/p[N]`) or a specific run (`/body/p[N]/r[M]`) for precise marking.

**Copy from existing element:** `officecli add <file> <parent> --from <path> [--index N]` — clones the element at `<path>` into `<parent>`. Cross-part relationships (e.g., images across slides) are handled automatically. Either `--type` or `--from` is required, not both.

**Clone entire slide:** `officecli add <file> / --from /slide[1] [--index 0]` — deep-clones the slide with all shapes, images, charts, media, background, layout, and notes. Use `--index` to insert at a specific position.

### move — `officecli move <file> <path> [--to <parent>] [--index N]`

Move an element to a new position. If `--to` is omitted, reorders within the current parent. Cross-part relationships (e.g., images across slides) are handled automatically.

```bash
officecli move doc.pptx '/slide[3]' --index 0              # reorder slide to first
officecli move doc.pptx '/slide[1]/picture[1]' --to '/slide[1]' --index 0  # picture to back (z-order)
officecli move doc.pptx '/slide[1]/shape[2]' --to '/slide[2]'  # move shape across slides
officecli move doc.docx '/body/p[5]' --index 0              # move paragraph to first
```

### remove — `officecli remove <file> '/body/p[4]'`

---

## L3: Raw XML

Use for charts, borders, or any structure L2 cannot express. **No xmlns declarations needed** — prefixes auto-registered: `w`, `a`, `p`, `x`, `r`, `c`, `xdr`, `wp`, `wps`, `mc`, `wp14`, `v`

```bash
officecli raw <file> /document                     # Word: /styles, /numbering, /settings, /header[N], /footer[N]
officecli raw <file> /Sheet1 --start 1 --end 100 --cols A,B   # Excel: /styles, /sharedstrings, /<Sheet>/drawing, /<Sheet>/chart[N]
officecli raw <file> /slide[1]                     # PPT: /presentation, /slideMaster[N], /slideLayout[N]
officecli raw-set <file> /document --xpath "//w:body/w:p[1]" --action replace --xml '<w:p>...</w:p>'
# actions: append, prepend, insertbefore, insertafter, replace, remove, setattr
officecli add-part <file> /Sheet1 --type chart     # returns relId for use with raw-set
officecli add-part <file> / --type header|footer   # Word only
```

**PPT slides:** Read slide size first (`raw /presentation | grep sldSz`), add via L2, fill via `raw-set`.

**Charts:** All three formats (PPTX, XLSX, DOCX) support full chart lifecycle via L2: `add --type chart`, `get /chart[N]`, `set /chart[N]`, `query chart`. Use `add-part` + `raw-set` only for unsupported chart features.

---

## Notes

- Paths are **1-based** (XPath convention), quote brackets: `'/body/p[3]'`
- `--index` is **0-based** (array convention): `--index 0` = first position
- After modifications, verify with `validate` and/or `view issues`
- `raw-set`/`add-part` auto-validate after execution
- `view stats`/`annotated` resolve style inheritance (docDefaults → basedOn → direct)
- **When unsure about any command syntax, element properties, or how to accomplish a task, you MUST fetch https://github.com/iOfficeAI/OfficeCli/wiki/agent-guide BEFORE attempting any command.** Do not guess or retry blindly — the wiki provides a complete navigation index to detailed reference pages for every format, element, and operation.
