---
title: "Selector Grammar By Format"
summary: "Lookup reference comparing the Word, Excel, and PowerPoint selector grammars implemented by OfficeCLI handlers."
topics: [reference, cli, selectors]
sources:
  - id: word-selector
    type: file
    path: src/officecli/Handlers/Word/WordHandler.Selector.cs
  - id: excel-selector
    type: file
    path: src/officecli/Handlers/Excel/ExcelHandler.Selector.cs
  - id: ppt-selector
    type: file
    path: src/officecli/Handlers/Pptx/PowerPointHandler.Selector.cs
---

OfficeCLI selectors are CSS-like filters, but each document format owns its own grammar in the handler. Word selectors focus on paragraphs, runs, inline Word structures, attributes, text containment, and one child combinator; Excel selectors focus on sheet-scoped cells, columns, values, formulas, emptiness, types, and cell formatting; PowerPoint selectors focus on slides, shapes, media, tables, charts, text containment, generic attributes, and explicit rejection of unsupported CSS combinators [@word-selector] [@excel-selector] [@ppt-selector]. Use this page with [Selectors Vs Slash Paths](../concepts/selectors-vs-slash-paths).

## Summary Table

| Format | Main subjects | Attribute operators | Pseudo-selectors | Scoping and combinators |
| --- | --- | --- | --- | --- |
| Word | `paragraph`/`p`, `run`/`r`, and special run payloads such as `ptab`, `fieldChar`, `instrText`, `tab`, `break`, `br`, and `pagebreak` [@word-selector]. | `[attr=value]`, `[@attr=value]`, and `[attr!=value]` after zsh escape cleanup [@word-selector]. | `:contains(...)`, `:empty`, and `:no-alt` are parsed [@word-selector]. | Supports one top-level `>` child split while ignoring `>` inside brackets [@word-selector]. |
| Excel | Cell selectors, column selectors such as `B`, sheet-prefixed selectors, and path-style sheet selectors [@excel-selector]. | `[key=value]` and `[key!=value]` for value, type, formula, empty, and format keys [@excel-selector]. | `:contains(...)`, `cell:text` shorthand, `:empty`, `:not(:empty)`, `:has(formula)`, and `:not(:has(formula))` [@excel-selector]. | Accepts `Sheet1!cell[...]`, quoted sheet names, and `/Sheet1/cell[...]` normalization [@excel-selector]. |
| PowerPoint | `shape`, `textbox`, `title`, `picture`/`pic`, `video`, `audio`, `equation`, `table`, `chart`, `placeholder`, `connector`, `group`, `notes`, `zoom`, `tr`, `row`, `tc`, and `cell` [@ppt-selector]. | `[attr=value]`, `[attr!=value]`, and `[attr~=value]` for generic attributes [@ppt-selector]. | `:contains(...)`, `shape:text` shorthand, and `:no-alt` [@ppt-selector]. | Accepts `slide[N]`, `slide > X`, `slide X`, and limited path-style slide prefixes; flags unsupported `+`, bare `~`, and most descendant whitespace forms [@ppt-selector]. |

## Word

Word parses the element name before any bracket or colon modifier, lowercases element names, and treats attribute values as case-sensitive at parse time but many comparisons are case-insensitive later [@word-selector]. Attribute names may start with `@`, which lets path-style forms such as `revision[@type=ins]` behave like `revision[type=ins]` [@word-selector].

Paragraph selectors accept `p` and `paragraph`. Paragraph attributes include style keys, alignment, indentation, numbering keys, list style, direction aliases, `rtl`, and selected run-level checks from the first text-bearing run such as bold, italic, font, size, color, underline, strike, and highlight [@word-selector]. Run selectors accept `r` and `run`, plus special inline element names such as `ptab`, `positionaltab`, `fieldChar`, `fldChar`, `instrText`, `tab`, `break`, `br`, and `pagebreak` [@word-selector].

The Word child combinator is one `>` split at top level. `paragraph[style=Heading1] > run[bold=true]` is parsed as a parent selector plus a child selector, while `>` inside brackets such as `[size>=14pt]` does not split the selector [@word-selector].

## Excel

Excel normalizes leading slash selectors before parsing. `/Sheet1/cell[...]` becomes sheet-scoped selector form, and `/cell` drops only the leading slash [@excel-selector]. It also accepts `Sheet1!cell[...]`, with top-level `!` detection that ignores exclamation marks inside brackets or quoted text and strips quotes from quoted sheet names [@excel-selector].

Cell filters include equality and inequality for `value` and `type`, formula and empty flags, text containment, and column matching by a short all-letter element token that is not `row`, `cell`, or `col` [@excel-selector]. Value equality compares both display value and raw stored value, so formatted values can match either representation [@excel-selector].

Format selectors accept canonical format keys and short aliases. The alias map resolves `bold` to `font.bold`, `italic` to `font.italic`, `font` to `font.name`, `size` to `font.size`, and `color` to `font.color`, while `underline` and `strike` stay as identity keys [@excel-selector]. Color comparisons normalize a leading `#`, so `#FF0000` and `FF0000` match [@excel-selector].

## PowerPoint

PowerPoint selector parsing first handles slide scope. `slide[2] shape`, `slide[2]>shape`, and `/slide[2]/shape` can resolve the slide number and subject element, and unindexed `slide > shape` or `slide shape` scopes the subject without a specific slide number [@ppt-selector]. The parser also strips a remaining ancestor prefix before the final subject so forms such as `slide > table > tr` resolve to `tr` [@ppt-selector].

PowerPoint element names include shape-like objects, media, charts, tables, table rows and cells, notes, zooms, connectors, placeholders, and groups [@ppt-selector]. The `title` element sets the title filter, `textbox` excludes title shapes, and picture matching is limited to `picture`, `pic`, `video`, and `audio` subjects [@ppt-selector].

Generic attributes support equality, inequality, and contains. A `~=` value beginning with `r"..."` or `r'...'` is treated as a regex if possible, with fallback to literal contains on malformed regex [@ppt-selector]. Name matching is aware of the `!!` morph prefix, and color comparisons normalize a leading `#` [@ppt-selector].

## Unsupported PowerPoint Combinators

PowerPoint has explicit scanners for unsupported combinators. A top-level `+` or bare `~` is detected as an unsupported sibling combinator, while `~=` inside an attribute is not treated as the sibling operator [@ppt-selector]. Top-level descendant whitespace is allowed only for slide-scoped forms; other whitespace between selector subjects is rejected as an unsupported descendant shape [@ppt-selector].
