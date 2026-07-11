---
title: "Rich Object Helper Map"
summary: "Lookup map for the shared chart, diagram, OLE, formula, and PowerPoint table-style helpers used by OfficeCLI handlers."
topics: [reference, handlers]
sources:
  - id: chart-core
    type: file
    path: src/officecli/Core/Chart/
  - id: diagram-core
    type: file
    path: src/officecli/Core/Diagram/
  - id: ole-helper
    type: file
    path: src/officecli/Core/OleHelper.cs
  - id: formula-core
    type: file
    path: src/officecli/Core/Formula/
  - id: table-styles
    type: file
    path: src/officecli/Core/TableStyles/
---

The rich object helper map identifies the shared subsystems that handlers use for Office objects more complex than plain text, cells, slides, and paragraphs. Charts, diagrams, OLE objects, formulas, and PowerPoint table styles live under `Core` so Word, Excel, and PowerPoint behavior can share parsing, generation, rendering, and validation rules instead of duplicating them per handler [@chart-core] [@diagram-core] [@ole-helper] [@formula-core] [@table-styles]. For the architectural overview, see [rich object helpers](../architecture/handlers/rich-object-helpers).

## Helper directories

| Area | Primary path | Main responsibility |
|---|---|---|
| Charts | `src/officecli/Core/Chart/` | Build, read, set, render, style, and preset chart XML across host formats [@chart-core] |
| Diagrams | `src/officecli/Core/Diagram/` | Parse Mermaid, lay out editable flowcharts and sequence diagrams, and optionally render Mermaid to PNG [@diagram-core] |
| OLE | `src/officecli/Core/OleHelper.cs` | Detect ProgIDs, classify embedded payloads, create embedded parts, normalize OLE properties, and populate OLE document nodes [@ole-helper] |
| Formulas | `src/officecli/Core/Formula/` | Convert LaTeX to OMML, evaluate Excel formulas, qualify modern functions, and guard formula recursion [@formula-core] |
| Table styles | `src/officecli/Core/TableStyles/` | Resolve PowerPoint built-in table style GUIDs, aliases, theme colors, regions, borders, fills, and emphasis bands [@table-styles] |

## Charts

`ChartHelper` is the shared chart build, read, and set helper for PowerPoint, Excel, and Word handlers, and its methods operate on Open XML chart parts rather than a specific host document [@chart-core]. It recognizes chart kinds such as column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo, waterfall, funnel, treemap, sunburst, box-whisker, histogram, and pareto [@chart-core].

`ChartExBuilder` handles Office 2016 extended chart types such as funnel, treemap, sunburst, box-whisker, histogram, and pareto [@chart-core]. `ChartSvgRenderer` extracts regular and extended chart information and renders chart previews using theme accent colors when available, falling back to Office default chart palette colors [@chart-core].

## Diagrams

`DiagramCompiler` is the shared entry point for `add --type diagram`; it sniffs the Mermaid header and dispatches to flowchart or sequence layout, while unsupported explicit diagram types are rejected with a clear error [@diagram-core]. `MermaidParser` handles a practical flowchart subset, including common node shapes, chained edges, edge labels, grouped endpoints, Unicode identifiers, and ignored style directives [@diagram-core].

`MermaidImageRenderer` is the optional high-fidelity path. It tries `mmdc`, then a Chrome-family browser, caches Mermaid JavaScript under `~/.officecli/cache`, stamps source text with the `mermaid:` alt-text prefix, and leaves native editable synthesis as the fallback when no image backend works [@diagram-core].

## OLE objects

`OleHelper` maps source extensions to default ProgIDs such as `Word.Document.12`, `Excel.Sheet.12`, `PowerPoint.Show.12`, `AcroExch.Document`, `Visio.Drawing`, or `Package` [@ole-helper]. It classifies Office-family payloads as embedded package parts and arbitrary binaries as generic embedded object parts [@ole-helper].

The same helper adds embedded parts to Word, Excel, PowerPoint, header, or footer host parts; handles self-embedding by writing a placeholder payload; decodes data URIs; validates ProgIDs; normalizes display values; creates placeholder icon images; wraps generic payloads as Ole10Native compound-file data; and unwraps Ole10Native payloads when reading [@ole-helper].

## Formulas

`FormulaParser` converts a LaTeX subset to Office Math Markup Language and includes a lenient parse path that records diagnostics and emits a valid placeholder instead of failing a whole batch [@formula-core]. It uses `DocumentLimits.MaxRecursionDepth` as the formula group-depth cap and collects unrecognized commands for warning output [@formula-core].

`FormulaEvaluator` represents numeric, string, boolean, error, array, range, blank, and lambda results, and its session object memoizes formula-cell results, cross-sheet evaluators, sheet data, materialized ranges, row and column extents, and circular-reference tracking [@formula-core]. `ModernFunctionQualifier` prefixes modern Excel functions with `_xlfn.` or `_xlfn._xlws.`, detects dynamic-array formulas that need spill metadata, and can unqualify formulas for user-facing display [@formula-core].

## Table styles

`TableStyleRegistry` maps PowerPoint's 74 built-in table-style GUIDs to family and accent pairs, and also exposes CLI aliases such as `medium2`, `light1`, `dark2`, and `none` [@table-styles]. `TableStyleResolver` resolves a style, cell position, and theme color map into concrete fill, text color, border, and bold values [@table-styles].

Table-style resolution uses region priority from lowest to highest: `wholeTbl`, banded columns, banded rows, first row, last row, first column, and last column [@table-styles]. The data model represents fills, border edges, table regions, table definitions, cell position flags, and resolved cell values [@table-styles].
