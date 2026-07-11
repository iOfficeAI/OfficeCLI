---
title: "Selectors Vs Slash Paths"
summary: "Selectors and slash paths are two addressing models in OfficeCLI: exact node paths identify known elements, while selectors find elements by type, content, or attributes."
topics: [concepts, handlers, cli]
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
  - id: command-query
    type: file
    path: src/officecli/CommandBuilder.GetQuery.cs
  - id: command-set
    type: file
    path: src/officecli/CommandBuilder.Set.cs
---

Selectors and slash paths are the two ways OfficeCLI names document content. A slash path such as `/body/p[1]`, `/Sheet1/A1`, or `/slide[1]/shape[2]` points at a specific node. A selector such as `paragraph[style=Heading1]`, `cell[bold=true]`, or `shape:Revenue` finds nodes by type, content, or attributes. The distinction matters because `get`, mutation commands, batch items, and handler internals often cross from one model to the other [@command-query] [@command-set].

## Exact Paths

Slash paths are best when the caller already knows the target. They are stable enough to pass between commands, and they match the document-node model used by `get` and many mutation operations. In practice, a path is a concrete address inside one document: a Word paragraph, an Excel cell or sheet element, or a PowerPoint slide element.

The CLI treats query as selector-oriented. Its `query` command accepts a CSS-like selector argument and then asks the current handler to produce matching `DocumentNode` results [@command-query]. That keeps broad searches separate from exact-node reads.

## Selectors

Selectors are intentionally format-specific. Word parses element names, bracket attributes, `:contains(...)`, `:empty`, `:no-alt`, and a `>` child combinator while matching paragraphs and runs against Word-specific properties such as style, numbering, font, and run content [@word-selector]. Excel accepts sheet prefixes, path-style selector normalization, cell value filters, formula and empty pseudo-selectors, type filters, and short format aliases such as `bold` mapping to `font.bold` [@excel-selector]. PowerPoint accepts slide scoping, shape/media/table/chart-oriented element types, `:contains(...)`, `:no-alt`, generic attributes, and rejects unsupported combinators with explicit selector errors [@ppt-selector].

This is why [selector grammar by format](../reference/selector-grammar-by-format) is a separate reference page. The shared idea is CSS-like filtering, not one universal CSS engine.

## Where They Meet

Handlers bridge selectors to paths by returning matching `DocumentNode` values. Once a selector has matched nodes, downstream code can use each node's concrete path for edits or output. The set command has a special selector-set path: if the target is a selector instead of a slash path, matching nodes are edited through the same filtering path used by query, and the handler records how many elements were touched [@command-set].

The bridge has safety rules. The CLI rejects bare unscoped selector-style mutation targets in cases where a short selector could accidentally update too much content [@command-set]. Query, by contrast, is meant for discovery and accepts broad selectors such as type names or attribute filters [@command-query].

## Format-Specific Normalization

Excel shows the most explicit path-selector overlap. Its selector parser normalizes path-style selectors such as `/Sheet1/cell[...]` into sheet-scoped selector form and also accepts Excel-native forms such as `Sheet1!cell[...]` [@excel-selector]. PowerPoint allows selected path-style slide scopes such as `/slide[1]/...`, but rejects other leading-slash selectors in query because a raw path like `/theme` is not a selector subject [@ppt-selector]. Word keeps its selector parser focused on element and attribute matching rather than treating arbitrary slash paths as selectors [@word-selector].

The useful rule is simple: use slash paths when you have an exact node, and use selectors when you need to find nodes. Pages such as [document node and paths](document-node-and-paths), [selector grammar by format](../reference/selector-grammar-by-format), and [adding handler mutations](../guides/adding-handler-mutations) depend on that distinction.
