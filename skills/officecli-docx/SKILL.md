---
name: officecli-docx
description: "Use this skill any time a .docx file is involved -- as input, output, or both. This includes: creating Word documents, reports, letters, memos, or proposals; reading, parsing, or extracting text from any .docx file; editing, modifying, or updating existing documents; working with templates, tracked changes, comments, headers/footers, or tables of contents. Trigger whenever the user mentions 'Word doc', 'document', 'report', 'letter', 'memo', or references a .docx filename."
---

# OfficeCLI DOCX Skill

## Setup

If `officecli` is missing:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

Verify with `officecli --version` (open a new terminal if PATH hasn't picked up). If install fails, download a binary from https://github.com/iOfficeAI/OfficeCLI/releases.

## ⚠️ Help-First Rule

**This skill teaches what good docx looks like, not every command flag. When a property name, enum value, or alias is uncertain, consult help BEFORE guessing.**

```bash
officecli help docx                         # List all docx elements
officecli help docx <element>               # Full element schema (e.g. paragraph, field, numbering, watermark, toc)
officecli help docx <verb> <element>        # Verb-scoped (e.g. add field, set section)
officecli help docx <element> --json        # Machine-readable schema
```

Help is pinned to the installed CLI version. When this skill and help disagree, **help is authoritative**.

## Mental Model

A `.docx` is a ZIP of XML parts (`document.xml`, `styles.xml`, `numbering.xml`, `header*.xml`, `footer*.xml`, `comments.xml`, …). Everything the user sees — headings, tables, page numbers, TOC, tracked changes — is XML inside that ZIP. `officecli` gives you a semantic-path API (`/body/p[1]/r[2]`) over it, so you almost never touch raw XML; when you must, use `raw-set` (see the XML appendix).

## Shell & Execution Discipline

docx paths contain `[]`; some prop values contain `$`. Both are shell metacharacters. Escaping happens at three layers — keep them separate.

1. **Shell.** ALWAYS quote element paths: `"/body/p[1]"`, not `/body/p[1]` (zsh/bash glob `[N]`). Single-quote any value containing `$`: `--prop text='$50M'` — at any length, the whole value inside one pair of single quotes. Unquoted `$50M` is stripped to `M`; mixing `'…$var…'` and `"…$50…"` on one long string is where the `$50` silently vanishes.
2. **CLI (`text=`).** The two-char escapes `\n` and `\t` ARE interpreted in `--prop text=` — `\n` becomes a `<w:br/>` soft line break, `\t` a `<w:tab/>` — consistently across docx / pptx / xlsx. Double them (`\\n`) for a literal backslash-n (rarely wanted). This applies to row-level table `c1…cN` shortcuts too (`\n` → `<w:br/>` within the cell).
3. **JSON (batch).** A real newline can also be passed as `"\n"` in the JSON string of a `batch` heredoc; same result.

If in doubt, `view text` after writing and compare character-for-character.

**Incremental execution.** `officecli` mutates the file on every call. Run commands one at a time and check each exit code — a 50-command script that fails at command 3 cascades silently. After any structural op (new style, table, TOC, section break) run `get` on it before stacking more.

**Open/save lifecycle:** `officecli open <file>` at the start, `officecli save <file>` at the end to flush to disk — `save` only writes and leaves the resident warm for follow-up edits; reach for `officecli close <file>` only to release the resident on a one-shot handoff. Both are always safe (never error or lose work). For many paragraphs of one style, use `batch` (one pass for the whole array). **Flush only at the non-officecli boundary:** officecli's own reads always see your edits; run `save`/`close` only before a non-officecli program reads the file (python-docx, Word, a renderer, delivery).

**`$FILE` convention.** All commands use `"$FILE"` — set it once (`FILE="your-doc.docx"`). Never copy a literal `doc.docx` / `review.docx` into output — always substitute your actual target.

## Requirements for Outputs

Deliverable standards every document MUST meet — know these before reaching for a command.

**Clear hierarchy.** Every non-trivial document has Title → Heading 1 → Heading 2 → body, not a wall of unstyled `Normal` paragraphs. If `view outline` shows one flat list, the hierarchy is missing.

**Explicit heading sizes** (Word default style sizes drift between templates): **H1 ≥ 18pt** (20pt for long reports), H2 = 14pt bold, H3 = 12pt bold, body = 11–12pt, line spacing 1.15–1.5x. Prefer `style=Heading1` over inline sizes so a retheme touches the definition once — but set explicit sizes when you can't trust the template's styles.

**One body font, one accent.** One readable body font (Calibri, Cambria, Georgia, Times New Roman); accent color for heading emphasis or table headers, not rainbow formatting.

**Spacing through properties.** Use `spaceBefore` / `spaceAfter` on paragraphs. Rows of empty paragraphs break pagination and are flagged by `view issues`.

**Typographic quality.** New content uses curly quotes (`'` `'` `"` `"`), not ASCII — Unicode directly or XML entities (`&#x2018;`/`&#x2019;`/`&#x201C;`/`&#x201D;`) inside `raw-set`. En-dash `–` for ranges (`2024–2026`), em-dash `—` for parenthetical breaks.

**Headers, footers, page numbers on any document > 1 page.** Page numbers go through a live `PAGE` field (`--prop field=page`), never the literal text "Page 1" — the CLI injects `<w:fldChar>` for you (see Headers & Footers).

**Preserve existing templates.** When editing a file that already has a look, match it — existing conventions override these guidelines.

### Visual delivery floor (applies to EVERY document)

Before declaring done, run `officecli view "$FILE" html` and Read the returned HTML path to confirm ALL of these:

- **No placeholder tokens rendered as data.** `$xxx$`, `{var}`, `{{name}}`, `<TODO>`, `lorem`, `xxxx` must never appear in a heading, body, cover, TOC, caption, header, or footer. A literal `{name}` meant for a human to fill belongs inside a visible instruction paragraph ("Replace `{name}` before sending"), never as finished content.
- **No truncated titles or overflowing cells.** Widen the column or set `wrapText` rather than trimming content.
- **TOC present when the document has 3+ headings** (`--type toc`).
- **Cover page ≥ 60% filled, last page ≥ 40% filled.** Pad a thin cover with subtitle / author / date / scope / key highlights; pad a "Thank you" last page with conclusion / next steps / contact / legal.
- **No `\$`, `\t`, `\n` literals in document text.** If `view text` shows these, a shell-escape layer leaked — delete the paragraph and re-enter it.

If any fails, STOP and fix before declaring done.

## Common Workflow

Six steps. Every non-trivial build follows this shape.

1. **Open the file.** `officecli open "$FILE"` (resident default). New file: `officecli create "$FILE"` first.
2. **Orient.** Existing file: `officecli view "$FILE" outline` — heading tree, section count, whether a TOC / watermark / tracked changes already exist. Never edit blind.
3. **Build incrementally.** Structural first, content next, formatting last: styles & numbering defs → sections / page setup → headings & body → tables / images / fields / TOC → headers / footers → comments. After each structural op, `get` it back before stacking on top.
4. **Format to spec.** Explicit heading sizes, spacing, widths, alignment, tabs, list indents — formatting is part of the deliverable, not optional polish.
5. **Save, then trust structure over cached text.** `officecli save "$FILE"` writes the XML. TOC / PAGE / NUMPAGES / SEQ / PAGEREF fields carry **cached values** that may be stale or empty until a human recalculates (F9 in Word). Confirm fields *exist* (`get --depth 3` finds `<w:fldChar>`) rather than trusting the visible text.
6. **QA — assume there are problems.** You are done after one fix-and-verify cycle finds zero new issues, not when your last command exited 0. See QA.

## Quick Start

Minimal viable docx: a heading, a body paragraph, a subheading, and a footer with a live page-number field. Adapt, don't copy-paste — your file, your content.

```bash
FILE="review.docx"
officecli create "$FILE"
officecli open "$FILE"
officecli add "$FILE" /body --type paragraph --prop text="Q4 2026 Review" --prop style=Heading1 --prop size=20pt --prop bold=true --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph --prop text="Revenue grew 18% year-over-year, ahead of plan." --prop size=11pt --prop spaceAfter=8pt
officecli add "$FILE" /body --type paragraph --prop text="Key Drivers" --prop style=Heading2 --prop size=14pt --prop bold=true --prop spaceBefore=12pt --prop spaceAfter=6pt
officecli add "$FILE" /body --type paragraph --prop text="Enterprise renewals, upsell, and a new EMEA region." --prop size=11pt
officecli add "$FILE" / --type footer --prop type=default --prop size=9pt --prop text="Page " --prop field=page
officecli set "$FILE" "/footer[1]/p[1]" --prop align=center
officecli save "$FILE"
officecli validate "$FILE"
```

Verified: `validate` returns `no errors found`; `get /footer[1] --depth 3` shows the 5-run PAGE field chain (begin / instrText / separate / cached value / end).

## Reading & Analysis

Start wide, then narrow. `outline` tells you what's already there; jump into `view text` / `get` / `query` once you know where to look.

```bash
officecli view "$FILE" outline            # heading tree, section count, table/image counts, watermark, tracked-changes presence — orient here first
officecli view "$FILE" html               # Read the returned HTML path: first visual check after a batch of edits (hierarchy, empty-para spacing, missing TOC)
officecli view "$FILE" text --start 1 --end 80   # text for content QA; paths shown as [/body/p[N]] so you can jump back with get
officecli view "$FILE" annotated          # values + style/font/size + warnings per run
officecli view "$FILE" stats              # paragraph counts, font usage, style distribution
officecli view "$FILE" issues             # empty paras, missing alt text, spacing anomalies
```

`officecli watch "$FILE"` keeps a live preview running for the human user to open at their discretion — agent self-check uses `view html`. Final visual verification is the user opening the `.docx` in Word / WPS / Pages.

**Inspect one element.** XPath-style semantic paths (1-based). Always quote — shells glob `[N]`. Use `[last()]` (with parens) for the last element; `[last]` errors. Add `--json` for machine output.

```bash
officecli get "$FILE" /                          # document root: metadata, page setup
officecli get "$FILE" "/body/p[1]"                # one paragraph
officecli get "$FILE" "/body/p[1]/r[1]"           # one run (character-level formatting)
officecli get "$FILE" "/body/tbl[1]" --depth 3    # table with rows and cells
officecli get "$FILE" "/footer[1]" --depth 3      # footer — check for fldChar
officecli get "$FILE" "/styles/Heading1"          # style definition
officecli get "$FILE" /numbering --depth 2        # numbering abstractNum + num bindings
```

**Query across the document.** CSS-like selectors, for systematic checks rather than hand-walking. Operators: `=`, `!=`, `~=` (contains), `>=`, `<=`, `[attr]` (exists). Full reference: `officecli query --help`.

```bash
officecli query "$FILE" 'paragraph[style=Heading1]'       # all H1s
officecli query "$FILE" 'p:contains("quarterly")'         # text match
officecli query "$FILE" 'p:empty'                         # empty paragraphs (clutter)
officecli query "$FILE" 'image:no-alt'                    # accessibility gaps
officecli query "$FILE" 'paragraph[size>=24pt]'           # numeric comparison
officecli query "$FILE" 'field[fieldType!=page]'          # fields other than PAGE
```

`query --json` wraps results in `.data.results[]` — `jq '.data.results | length'` to count.

**Large documents.** Navigate by heading with `view outline` and jump with `query`; don't dump the whole body into context.

## Creating & Editing

Verbs: `add` (new element), `set` (change a prop), `remove`, `move`, `swap`, `batch`, `raw-set` (last-resort XML). Ninety percent of a build is paragraphs, runs, tables, a couple of images, a TOC, and a footer.

### Paragraphs, runs, styles

A paragraph (`p`) is a block; a run (`r`) is a span of consistent character formatting inside it. Set paragraph-level props (style, alignment, spacing, indent) on the `p`; set font / size / color / bold on the `r`.

```bash
officecli add "$FILE" /body --type paragraph --prop text="Executive Summary" --prop style=Heading1 --prop size=18pt --prop bold=true --prop spaceAfter=12pt
officecli set "$FILE" "/body/p[1]/r[1]" --prop color=1F4E79
```

Use `spaceBefore` / `spaceAfter` for vertical spacing — never chains of empty paragraphs. For left indent use `--prop indent=720` (twips), `firstLineIndent=360` for first line, `hangingIndent=720` for hanging; leading spaces fire `view issues`.

### Tables

Tables are `/body/tbl[N]` with rows `tr[N]` and cells `tc[N]`. Add with row/column counts, then fill.

```bash
officecli add "$FILE" /body --type table --prop rows=4 --prop cols=3 --prop width=100%
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop header=true --prop c1=Quarter --prop c2="Revenue" --prop c3="Growth"
officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[1]/r[1]" --prop bold=true
```

Row-level `set` supports `height`, `header`, and `c1 / c2 / … / cN` text shortcuts (`cN` generalises to any column count). Cell formatting (bold, fill, color) goes on the cell's paragraph / run — **not** row-level. For per-cell borders, set cell-level `border.*` on the `tc` (`--prop border.bottom="single;6;000000;0"`), or paragraph-level `pbdr.*` on the inner paragraph.

**Horizontal rule = a paragraph bottom border, never a 1-row table.** A table-as-divider renders as an empty min-height box (worst in headers/footers). Use `pbdr.bottom` (`STYLE;SIZE;COLOR`) on the paragraph instead:

```bash
officecli set "$FILE" "/body/p[3]" --prop pbdr.bottom="single;6;2E75B6"
```

### Lists (bullets, numbered, multi-level)

For single-level bullets/numbers, set `listStyle` on the paragraph (`listStyle` is a paragraph prop, NOT a run prop — common mistake):

```bash
officecli add "$FILE" /body --type paragraph --prop text="First item" --prop listStyle=bullet
```

For multi-level (legal-style 1 / 1.1 / 1.1.1), add an `abstractNum`, then a `num`, then reference the `numId` per paragraph:

```bash
officecli add "$FILE" /numbering --type abstractnum --prop format=decimal     # → abstractNum id=0
officecli add "$FILE" /numbering --type num --prop abstractNumId=0             # → num id=1
officecli add "$FILE" /body --type paragraph --prop text="Section one" --prop numId=1 --prop ilvl=0
```

IDs are 0-based: the first `abstractNum` is id=0; the `num` references it via `abstractNumId=0` and is itself assigned id=1. A non-existent `abstractNumId` errors, so check ids after creating. Verify with `officecli query "$FILE" 'paragraph[numId>0]'`. See `help docx abstractnum` / `help docx num` for level and format options.

### Tab stops (signature lines, leader rows)

Tab stops are a first-class `tab` child of the paragraph; `pos` accepts `6in`/`6cm`/twips, `val` ∈ `left`/`center`/`right`, `leader` ∈ `none`/`dot`/`hyphen`/`underscore`. See `help docx tab`.

```bash
officecli add "$FILE" "/body/p[1]" --type tab --prop pos=6in --prop val=right --prop leader=dot
```

**Leader caveat.** `leader=dot` alone emits no dots — the leader renders only when a real `<w:tab/>` character sits in a run between the text and the tab stop. Put one there with `\t` in the text: define the stop (`add tab --prop pos=6in --prop val=right --prop leader=dot`), then `--prop text="Chapter 1\t12"` — the `\t` becomes the `<w:tab/>` and dots fill to the right-aligned page number. (Literal `text="Chapter 1 ......... 12"` also ships, but a real tab stop aligns cleanly.)

### Fields (PAGE / NUMPAGES / DATE / MERGEFIELD / REF)

Fields are live values computed at render time. `fieldType` picks the field; `name` supplies the target (merge name or `ref` bookmark); `format` / `instr` add switches.

| Field | Use | Example |
|---|---|---|
| `page` | current page number | `--prop field=page` on footer, or `--prop fieldType=page` inline |
| `numpages` | total pages | `--prop field=numpages` / `--prop fieldType=numpages` |
| `date` | today | `--prop fieldType=date --prop format='yyyy-MM-dd'` |
| `mergefield` | template merge token | `--prop fieldType=mergefield --prop name=CustomerName` |
| `ref` | cross-reference to a bookmark | `--prop fieldType=ref --prop name=bookmarkName` |

Full `fieldType` enum (30+ values incl. `pageref`, `seq`, `styleref`, `docproperty`, `createdate`, …) is in `help docx field`. **There is NO `fieldInstr` fieldType** — use the `instruction` prop for raw field instruction text when typed shortcuts fall short. Picture switches (`MERGEFIELD Amount \# "#,##0.00"`, `DATE \@ "yyyy年MM月"`) go via `--prop instruction='…'` (mergefield's `format` prop is ignored with a warning — use `instruction`).

```bash
officecli add "$FILE" "/body/p[3]" --type field --prop fieldType=mergefield --prop name=customer_name
# Renders «customer_name» — visible placeholder, replaced in Word at mail-merge time.
```

**MERGEFIELD templates: never render placeholder literals.** A `{{customer_name}}` or `$NAME$` shown as body text is a failed template the recipient sees — insert a real MERGEFIELD (above), or confine literal tokens to an obvious instruction paragraph. Confirm with `query 'field[fieldType=mergefield]'`.

**SEQ / PAGEREF / TOC field values.** officecli doesn't store rendered field values at write time. Recompute by what each path needs:

- **SEQ numbering** (`Figure 1/2/3`): `officecli set "$FILE" / --prop recalcFields=seq` counts SEQ fields in body document order and writes the cached values (`evaluated` flips true; switches/formats in `help docx document`). Heading-relative `\s` and SEQ in headers/footers defer to Word.
- **PAGE / PAGEREF / NUMPAGES / TOC page numbers** need pagination, which officecli has no engine for — `officecli set "$FILE" /settings --prop updateFields=true` defers them to Word on open.

Use both on a multi-figure document. Academic papers: see the `officecli-academic-paper` skill.

### Headers & Footers (page numbering)

Single-command pattern — the CLI injects `<w:fldChar>`, so you never compose the field by hand:

```bash
# Empty first-page footer — auto-enables differentFirstPage so the cover has no page number
officecli add "$FILE" / --type footer --prop type=first --prop text=""
# Default footer with live page number
officecli add "$FILE" / --type footer --prop type=default --prop align=center --prop size=9pt --prop text="Page " --prop field=page
```

When both exist, the default footer is `/footer[2]`; alone it is `/footer[1]`. **Verify**: `get --depth 3` must show `fldChar` children, not just a run with literal `"Page"` (`view outline` prints "Footer: Page" for both live fields AND static text — don't rely on it). Do NOT `set --prop differentFirstPage=true` — that prop is unsupported (rejected with exit 2, not silently); adding a first-type footer flips the bit. For composite **Page X of Y**, see recipe (b).

### Table of Contents

For any document with 3+ headings, first verify that the requested range has real sources. A TOC source must use a built-in Heading style or a custom paragraph style with `outlineLvl`; bold or large Normal text is not a source. If no eligible source exists, stop and fix the heading structure instead of inserting a TOC that renders "Error! No table of contents entries found."

```bash
# Built-in Heading styles are TOC sources.
officecli add "$FILE" /body --type paragraph --prop text="Introduction" --prop style=Heading1
# Custom paragraph styles need an outline level (0 = Heading 1).
officecli add "$FILE" /styles --type style --prop id=ThesisH1 --prop type=paragraph --prop outlineLvl=0
# Add the TOC only after its sources exist.
officecli add "$FILE" /body --type toc --prop levels="1-3" --prop title="Table of Contents" --prop hyperlinks=true --index 0
```

Page numbers need pagination; OfficeCLI cannot calculate TOC page numbers itself. Set `updateFields=true` so Word recomputes the TOC (and all fields) on open. Without a Word-compatible field engine, report the TOC as dynamic and uncomputed; do not claim ready page numbers or guess a static TOC.

```bash
officecli set "$FILE" /settings --prop updateFields=true
```

Address the TOC directly with `/toc[1]` or `/tableofcontents` for `get`/`set`/`remove`.

### Images

Pictures go inside a run. Alt text is mandatory for accessibility — pass `alt` directly at create time:

```bash
officecli add "$FILE" "/body/p[5]" --type picture --prop src=logo.png --prop width=1.5in --prop alt="Acme logo"
```

Confirm `officecli query "$FILE" 'image:no-alt'` is empty before delivery.

### Charts

For data, add a **native chart** — editable, themeable, accessible, re-renders in Word — never a flat PNG screenshot of a chart. `data="Label:v1,v2,…"` per series; one `data=` per series (or `series1=`/`series2=`).

```bash
officecli add "$FILE" /body --type chart --prop chartType=bar --prop title="Revenue by Region" --prop categories="EMEA,APAC,Americas" --prop data="2026:120,150,180"
```

`chartType` ∈ bar / column / line / pie / area / scatter (`help docx chart` for axis/legend/series styling). A PNG via `--type picture` is only a fallback for an exotic chart officecli can't build.

### Hyperlinks and bookmarks

External links go via `hyperlink`:

```bash
officecli add "$FILE" "/body/p[2]" --type hyperlink --prop url="https://example.com" --prop text="our site"
```

**Internal links** (to a bookmark) use `--prop anchor=bookmarkName` — not a `#fragment` in `url`:

```bash
officecli add "$FILE" "/body/p[2]" --type hyperlink --prop anchor=chapter1 --prop text="See Chapter 1"
```

Pairing a `PAGEREF` field with visible text is the alternative. See `help docx hyperlink` / `help docx bookmark`.

### Sections and page setup

Document root `/` carries page setup (`pageWidth`, `pageHeight`, margins, in twips). Multi-section documents (landscape insert, columns) add a `section` break — see `help docx section`. Both camelCase (`pageWidth`, canonical) and lowercase alias (`pagewidth`) are accepted; prefer camelCase.

```bash
officecli set "$FILE" / --prop pageWidth=12240 --prop pageHeight=15840 --prop marginTop=1440 --prop marginLeft=1440
# Newspaper-style multi-column flow (columnSpace in twips; 720 = 0.5in):
officecli set "$FILE" / --prop columns=2 --prop columnSpace=720
```

### Forcing page breaks

Use exactly one page-break mechanism per logical boundary. The default is `pageBreakBefore=true` on the target heading. Alternatively, insert one explicit `pagebreak` before the heading. Never combine both mechanisms for the same boundary: the resulting double break can create a blank page. Never set `pageBreakBefore` on `p[last()]` immediately after adding a `pagebreak`.

```bash
# Default: apply the break directly to the heading.
officecli add "$FILE" /body --type paragraph --prop text="Introduction" --prop style=Heading1 --prop pageBreakBefore=true
```

`--prop break=newPage` is an alias for `pageBreakBefore=true`. Preview with `view html` and check for blank pages.

### Report-level recipes

Patterns that come up on every long-form report. Each has been executed and `validate`-passed.

**(a) Rich cover page — hit the ≥ 60% filled floor.** Stack a confidentiality banner, title, subtitle, client/project/date block, and a key-themes strip, then force the next section onto a new page:

```bash
officecli add "$FILE" /body --type paragraph --prop text="CONFIDENTIAL — CLIENT USE ONLY" --prop align=center --prop size=9pt --prop color=C00000 --prop spaceAfter=24pt
officecli add "$FILE" /body --type paragraph --prop text="Strategic Growth Review" --prop style=Title --prop size=32pt --prop bold=true --prop align=center --prop font=Cambria --prop spaceAfter=8pt
officecli add "$FILE" /body --type paragraph --prop text="FY26 Outlook and Scenario Planning" --prop italic=true --prop size=16pt --prop align=center --prop spaceAfter=36pt
officecli add "$FILE" /body --type paragraph --prop text='Prepared for: Acme Corp. Leadership Team' --prop align=center --prop size=11pt
officecli add "$FILE" /body --type paragraph --prop text='Engagement: 2026-04 — 2026-06' --prop align=center --prop size=11pt --prop spaceAfter=36pt
officecli add "$FILE" /body --type paragraph --prop text="Key themes: 1) margin resilience, 2) EMEA expansion, 3) capital allocation." --prop align=center --prop italic=true --prop size=10pt
officecli add "$FILE" /body --type paragraph --prop text="Executive Summary" --prop style=Heading1 --prop pageBreakBefore=true
```

**(b) Page X of Y footer — composite PAGE + NUMPAGES.** Add the footer paragraph, then three child ops build `Page <X> of <Y>` live. The official `help docx footer` recipe.

```bash
officecli add "$FILE" / --type footer --prop type=default --prop text="Page " --prop align=center --prop size=9pt
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=page
officecli add "$FILE" "/footer[1]/p[1]" --type run --prop text=" of "
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=numpages
officecli get "$FILE" "/footer[1]/p[1]" --depth 1 | grep -o fldChar | wc -l   # expect ≥ 4; use grep -o ... | wc -l, NOT grep -c (single-line XML returns 1)
```

**(c) Header row with fill and white bold text.** Order matters — populate header cell text FIRST (runs don't exist in empty cells; a `set …/tc[N]/p[1]/r[1]` on an empty cell errors "No r found"), THEN cell fill, THEN run formatting:

```bash
officecli add "$FILE" /body --type table --prop rows=5 --prop cols=4 --prop width=100%
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop header=true --prop c1=Quarter --prop c2=Revenue --prop c3=Growth --prop c4=Status
for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[$col]" --prop fill=1F4E79
  officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[$col]/p[1]/r[1]" --prop bold=true --prop color=FFFFFF
done
for row in 3 5; do for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[$row]/tc[$col]" --prop fill=D9E2F3      # zebra stripe
done; done
```

**(d) Financial table — right-align numbers, bold totals, bottom border on total row.**

```bash
for row in 2 3 4 5; do for col in 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[$row]/tc[$col]/p[1]" --prop align=right
done; done
for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[5]/tc[$col]/p[1]/r[1]" --prop bold=true
  officecli set "$FILE" "/body/tbl[1]/tr[4]/tc[$col]/p[1]" --prop pbdr.bottom="single;6;000000;0"
done
```

**(e) Cell with multiple bullets (SWOT / risk matrix).** `c1="a\nb"` gives a `<w:br/>` line break within **one** paragraph — fine for plain multi-line text, but bullets need separate paragraphs. Seed the first via `set c1=`, then `add paragraph` (with `listStyle=bullet`) under the cell per subsequent bullet:

```bash
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop c1="Installed base of 18k enterprise seats"
officecli add "$FILE" "/body/tbl[1]/tr[1]/tc[1]" --type paragraph --prop text="Margin structure above peer median" --prop listStyle=bullet
officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[1]" --prop listStyle=bullet
```

If the seeded line lands at the bottom, re-order: `officecli move "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[N]" --index 0`.

**(f) TOC without a field engine.** A CLI-only pipeline cannot calculate reliable page numbers. Keep the live TOC, set `updateFields=true`, and tell the recipient to open the document in Word or another compatible field engine. Do not replace it with guessed static page numbers.

### Template delivery — separating Template Notes from end-user content

HR / legal / vendor templates carry internal-only guidance ("replace `{{CompanyName}}`") that must NOT ship. Two working patterns:

- **Trailing "Template Notes" section** under a clear `Heading 1` ("Template Notes for HR Users") with all instructions below it; before distribution, `remove` from the heading downward (locate with `query 'paragraph[style=Heading1]:contains("Template Notes")'`).
- **Bookmark-bounded internal section** between `__template_notes_start` / `_end` bookmarks; at delivery `raw-set` removes everything between the anchors.

Delivery gate for templates: after removal, `query 'p:contains("Template Notes")'` AND `query 'p:contains("{{")'` both return empty. If a notes paragraph survives, a downstream employee reads internal language.

### Advanced / specialty topics (skip if you are writing a report)

Reports, memos, letters, proposals, and HR templates don't need this. Keep reading only if your document is academic (equations, footnotes, bibliography), reviewed (comments, tracked changes), or marked (watermark).

**Equations and footnotes.** `--type equation` takes LaTeX — `\frac`, `\sum`, Greek, `\mathit`, `\mathcal` all render. By default it creates a standalone `/body/oMathPara[N]` display block; pass `--prop mode=inline` with a paragraph parent path (`add "/body/p[N]" --type equation --prop formula=… --prop mode=inline`) to drop an inline `<m:oMath>` into running text. Footnotes auto-number by paragraph index. Bibliography hanging indent: `firstLineIndent=-720 indent=720` per entry.

```bash
officecli add "$FILE" /body --type equation --prop formula="\\frac{a}{b} + \\sum_{i=1}^{n} x_i"
officecli add "$FILE" "/body/p[3]" --type footnote --prop text="See Appendix A for methodology."
```

**Comments and tracked changes.** Bulk accept/reject: `set "$FILE" /revision --prop revision.action=accept` (or `--prop revision.action=reject`); narrow with a selector like `/revision[@author=Alice]` or `/revision[@type=ins]`. Locate individual changes with `query ins` and `query del` (`trackedchange` is not a selector). Create tracked changes on a run with `--prop revision.type=ins|del --prop revision.author=…` (`help docx run` for the full `revision.*` set — `format`/`moveFrom`/`moveTo` too). Add a comment: `add "/body/p[4]" --type comment --prop author=… --prop text=…`; reply-thread it with `--prop parentId=N` and mark it resolved with `set "/comments/comment[N]" --prop done=true` (resolve rather than delete to keep the audit trail — `query 'comment[done=false]'` then lists what's still open). Prop schema: `help docx comment` / `help docx run`.

**Watermark.** `add / --type watermark --prop text="DRAFT" --prop color=BFBFBF --prop opacity=0.8` in one command (default opacity 0.5); `set /watermark --prop opacity=…` adjusts it later.

**When to switch skills.** Stay in docx for chapter drafts, ≤ 3 footnotes, ≤ 2 equations, no bibliography/cross-refs. Switch to **`academic-paper`** for citation styles (APA / Chicago / IEEE / GB 7714), in-text↔reference auto-linking, numbered equations with `\ref`, "List of Figures", or auto-updating cross-refs. Switch to **`officecli-word-form`** when the document's purpose is **data capture** — fillable forms, contracts with user-fill slots, questionnaires, mail-merge templates (`<w:sdt>` content controls, `<w:ffData>`, `documentProtection=forms`).

### Raw-set escape hatch (L1 / L2 / L3)

Three tiers of precision; use the lowest that does the job.

- **L1 — high-level props** (`--prop text=…`, `--prop style=Heading1`): your default. Covers 80%.
- **L2 — dotted-attr fallback** (`pbdr.top=`, `ind.left=`, `shd.fill=`, `padding.top=`, `font.size=`): when L1 lacks the knob. Example: `--prop pbdr.bottom="single;6;1F4E79;0"`. Emits schema-valid XML.
- **L3 — `raw-set` with XML**: last resort, no schema protection. Use for internal hyperlinks, composite fields, and other shapes the typed verbs can't express (see XML appendix).

Borders use the format `style;size;color;space`: `single;4;FF0000;1`. Hex colors never start with `#`: `FF0000`. Scheme color names (`accent1..6`, `dark1`/`dark2`, `light1`/`light2`, `hyperlink`) are accepted anywhere a hex color is — prefer hex for stable colors across themes.

## QA (Required)

**Assume there are problems — QA is a bug hunt, not a confirmation step.** Your first document is almost never correct; zero issues on first inspection means you weren't looking hard enough. Headings look fine until `view outline` shows an H3 directly under an H1; the footer shows "Page 1" until `get --depth 3` reveals a static run, not a field.

### Minimum cycle before "done"

1. `officecli view "$FILE" issues` — empty paras, missing alt text, formatting anomalies.
2. `officecli view "$FILE" outline` — heading hierarchy (no H1 → H3 skips), TOC presence, section count.
3. `officecli view "$FILE" text --max-lines 400` — typos, stray `\$`/`\t`/`\n` literals, placeholder tokens.
4. `officecli validate "$FILE"` — schema check (the Delivery Gate re-runs this on the closed, on-disk file).
5. **Visual pass — whole document as a contact sheet** (vision-capable agents only — if you cannot interpret images, skip this step: steps 1–4 are your ceiling, and flag the document "not visually verified" at handoff). `officecli view "$FILE" screenshot --grid auto -o /tmp/sheet.png`, then Read it. `--grid auto` tiles **every page** into one image (auto column count; `--grid 4` to force) — you *see* pagination, blank pages, heading rhythm, lopsided margins, and TOC/cover placement, not just the DOM. Windows+Word renders each page through real Word; elsewhere HTML. If the screenshot fails, fall back to `view html` and flag cross-page breaks / alignment / rhythm as "not visually verified". Thumbnails only **locate**: confirm any fine call (column alignment, line spacing, indents, dark-on-dark, caption placement) on the suspect page at full resolution with `screenshot --page N` (no `--grid`; real Word on Windows). "validate pass" is not delivery; "looks like a real document" is.
6. If anything failed, fix, then **rerun the full cycle** — one fix commonly creates another problem.

### Delivery Gate (run before handing off — any failure = REJECT, do NOT deliver)

Copy-paste, set `FILE`, and refuse to declare done until every gate prints OK.

```bash
FILE="your-file.docx"

# Gate 1 — schema.
officecli close "$FILE" 2>/dev/null
officecli validate "$FILE" | grep -q "no errors found" || { echo "REJECT Gate 1: validate failed"; exit 1; }
echo "Gate 1 OK"

# Gate 2 — token leak (shell-escape / template tokens / literal \$ \t \n). grep -c never false-PASSes.
LEAK=$(officecli view "$FILE" text | grep -cE '(\$[A-Za-z_]+\$|\{\{[^}]+\}\}|<TODO>|xxxx|lorem|\\[\$tn])')
[ "$LEAK" -eq 0 ] && echo "Gate 2 OK" || { echo "REJECT Gate 2: $LEAK leak line(s)"; officecli view "$FILE" text | grep -nE '(\$[A-Za-z_]+\$|\{\{[^}]+\}\}|<TODO>|xxxx|lorem|\\[\$tn])'; exit 1; }
# A TOC placeholder is valid before a Word-compatible field engine updates it; confirm the TOC field and updateFields setting structurally instead.

# Gate 3 — live PAGE field exists when a footer is expected.
FLD=$(officecli query "$FILE" 'field[fieldType=page]' --json | jq '.data.results | length')
[ "$FLD" -ge 1 ] && echo "Gate 3 OK" || { echo "REJECT Gate 3: no live PAGE field"; exit 1; }
echo "Delivery Gate PASS"
```

### Field / cached-value spot-check

Fields carry cached values that may be stale or empty at write time — confirm existence by **structure, not text**.

- **Footer PAGE:** `get /footer[N] --depth 3` lists the begin / instrText / separate / cached / end run chain — ≥ 5 runs for one PAGE, ≥ 11 for composite "Page X of Y". A single run with text `"Page"` = field missing; re-add with `--prop field=page`.
- **TOC:** `get /toc[1] --depth 2` shows field structure. Page numbers may read `1 1 1 1` or `Update field to see…` until recalculated (see TOC delivery step).
- **MERGEFIELD:** `query 'field[fieldType=mergefield]'` — one per slot, no literal `{{name}}` elsewhere.

### Honest limit

`validate` catches schema errors, not design errors — a document can pass it with wrong heading hierarchy, fake-Heading-1 sizes, placeholder tokens as body text, or an empty first-page footer on a coverless document. The contact-sheet visual pass (`screenshot --grid`) and the field-structure check are how you catch what validation can't.

### QA display notes (don't chase these)

- `view text` shows `"1."` for every numbered list item regardless of rendered number — actual output increments correctly.
- `view issues` flags "body paragraph missing first-line indent" on cover paragraphs, centered headings, list items, bibliography entries — first-line indent is only required for APA/academic body text; on block-style professional documents these are expected.

## Known Issues & Pitfalls

When something "looks broken", attribute it before chasing: **[AGENT-ERROR]** the document is wrong (fix it) · **[RENDERER-BUG]** the document is correct, a viewer renders it differently (don't chase) · **[SKILL gap]** the skill didn't teach the rule (file an issue).

### Renderer quirks (cross-viewer, [RENDERER-BUG] — don't chase)

Before calling a color/field/chart broken, open the file in the user's target viewer; if it looks correct there it's a viewer quirk.

- **PAGE field may render literal "Page"** (no number) until recalculated — judge by `fldChar` presence, not the digit.
- **TOC cached page numbers may read "1 1 1 1"** until F9.
- **Pie / doughnut fill may collapse to one color** in some viewers (column/bar render fine).
- **Form-control checkboxes may render double-boxed**; **OMML equation baselines** may shift across viewers (XML identical).

### Common pitfalls

| Pitfall | Correct approach |
|---|---|
| `--index` vs `[N]` | `--index` is 0-based; `[N]` paths are 1-based |
| Multiple `add --index N` with the same N | Each insert shifts later content down; reusing N puts later items BEFORE earlier ones — insert in reverse order, or `move --after/--before` anchored on `paraId` |
| Unquoted `[N]` in zsh/bash | Quote every path: `"/body/p[1]"` |
| `[last]` as predicate | Must be `[last()]` with parens |
| Raw twips in spacing | Use unit-qualified values: `12pt`, `0.5cm`, `1.5x` |
| Empty paragraphs for spacing | Use `spaceBefore` / `spaceAfter` |
| Row-level `set` for cell formatting | Row `set` only supports `height`, `header`, `c1..cN` text; format goes on the cell paragraph / run |
| `listStyle` on a run | It's a paragraph property |
| Indent via leading spaces | `indent=720` / `firstLineIndent=360` / `hangingIndent=720` (dotted `ind.left` / `ind.firstLine` also work) |
| Cover page-number suppression via `set differentFirstPage=true` | UNSUPPORTED — add a first-type footer: `--type footer --prop type=first --prop text=""` |
| Need a fresh chapter page | Use `pageBreakBefore=true` on the heading; use an explicit `pagebreak` only as an alternative, never both |
| Multiple bullet paragraphs in one cell | `c1="a\nb"` makes a `<w:br/>` line break (one paragraph); for separate bullet paragraphs use recipe (e) |
| `raw-set` when dotted-attr would work | Prefer L2 dotted-attr over L3 raw-set |
| Next paragraph inherits the previous Heading style | Set explicit `--prop style=Normal` on the following paragraph |
| Modifying a file open in Word | Close it in Word first |
| Echo into batch breaks on `$`/`'` | Heredoc with single-quoted delimiter: `cat <<'EOF' \| officecli batch …` |

## Raw-set XML appendix (L3 patterns)

`raw-set` injects literal OOXML — no schema protection. Element order in `<w:pPr>`: `pStyle`, `numPr`, `spacing`, `ind`, `jc`, `rPr` (last). Smart quotes as entities (`&#x2018;`/`&#x2019;`/`&#x201C;`/`&#x201D;`). Add `xml:space="preserve"` to any `<w:t>` with leading/trailing spaces. RSIDs are 8-digit hex. Use "Claude" as the author for tracked changes/comments unless the user names another.

**Tracked-change insertion / deletion** — prefer the high-level `--prop revision.type=ins|del` on a run; raw-set only for what the typed path can't express (rejecting/restoring another author's change, below). Replace the whole `<w:r>…</w:r>`, never inject tags inside a run; copy the original `<w:rPr>` into both to preserve formatting. Inside `<w:del>` use `<w:delText>` (and `<w:delInstrText>` for instructions):

```xml
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30</w:delText></w:r></w:del>
<w:ins w:id="2" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:t>60</w:t></w:r></w:ins>
<w:r><w:t> days.</w:t></w:r>
```

When deleting ALL content of a paragraph/list item, also mark the paragraph mark deleted (`<w:del/>` inside `<w:pPr><w:rPr>`) — otherwise accepting changes leaves an empty paragraph. To **reject another author's insertion**, nest your `<w:del>` inside their `<w:ins>`; to **restore their deletion**, add a `<w:ins>` after it (don't modify theirs).

**Internal hyperlink to a bookmark** (prefer the high-level `--prop anchor=` path; raw-set only for custom run styling the command can't express):

```xml
<w:hyperlink w:anchor="chapter1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>See Chapter 1</w:t></w:r></w:hyperlink>
```

**Composite field in one run** (e.g. two fields the single-command path can't compose) — the `fldChar begin / instrText / separate / value / end` chain:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

**Comment markers** are siblings of `<w:r>`, NEVER inside one (reply threading and resolved-state are high-level — `--prop parentId=`/`done=`, see Comments above):

```xml
<w:commentRangeStart w:id="0"/><w:r><w:t>annotated text</w:t></w:r><w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
```

Force field recalc on open with `officecli set "$FILE" /settings --prop updateFields=true` (writes `<w:updateFields w:val="true"/>`; covers the layout-dependent fields PAGE / PAGEREF / NUMPAGES / TOC page numbers — no raw-set needed). For SEQ numbering, prefer `set / --prop recalcFields=seq`, which writes correct cached values now without waiting on Word.

### Help pointer

When in doubt: `officecli help docx`, `officecli help docx <element>`, `officecli help docx <verb> <element>`, `--json` for agents. Help is the authoritative schema; this skill is the decision guide.
