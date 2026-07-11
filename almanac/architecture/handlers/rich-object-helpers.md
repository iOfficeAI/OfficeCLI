---
title: "Rich Object Helpers"
summary: "Shared helper subsystems that keep charts, diagrams, OLE objects, and formulas consistent across OfficeCLI handlers."
topics: [architecture, handlers, charts]
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
---

Rich object helpers are the shared subsystems that stop each handler from inventing its own chart, diagram, OLE, and formula behavior. They live under `Core`, operate on Open XML parts or neutral intermediate models, and are called by Word, Excel, and PowerPoint handlers when a document contains objects more complex than text, cells, slides, or paragraphs [@chart-core] [@diagram-core] [@ole-helper] [@formula-core]. This makes rich objects part of the normal [document handler lifecycle](document-handler-lifecycle) instead of isolated per-format features.

## Charts

The chart helpers own chart parsing, construction, mutation, and preview rendering. `ChartHelper` is documented as shared chart build, read, and set logic for PowerPoint, Excel, and Word handlers, and its methods operate on `ChartPart`, `C.Chart`, and `C.PlotArea` rather than on a host document type [@chart-core]. It parses chart types, series data, cell references, categories, colors, axes, and chart properties before building or updating chart XML [@chart-core].

Extended chart support is split into `ChartExBuilder`, `ChartExResources`, and related style builders for newer chart families such as histogram, funnel, treemap, sunburst, box-and-whisker, and waterfall [@chart-core]. `ChartSvgRenderer` reads regular and extended chart information and renders chart previews as SVG/HTML-friendly output, using theme-derived colors when available and fallback Office palette colors otherwise [@chart-core]. The result is one chart vocabulary across add, set, query, dump, and preview flows.

## Diagrams

Diagram helpers translate Mermaid input into either editable Office shapes or raster images. `DiagramCompiler` is the native entry point for `add --type diagram`; it recognizes flowchart and sequence diagram headers, dispatches to the corresponding layout engine, and rejects unsupported diagram types with a clear error [@diagram-core]. `MermaidParser` parses a practical Mermaid flowchart subset into a semantic graph, including common node shapes, chained edges, labels, grouped endpoints, Unicode identifiers, and ignored style directives [@diagram-core].

When native editable shapes are not enough, `MermaidImageRenderer` can render Mermaid through `mmdc` or a Chrome-family browser and return a PNG [@diagram-core]. It caches Mermaid JavaScript under the user's OfficeCLI cache, stamps source text into image alt text with a sentinel tag, and falls back to the native compiler when no image backend is available [@diagram-core]. That gives handlers a choice between editable shape synthesis and higher-fidelity raster output.

## OLE objects

`OleHelper` is the shared OLE boundary for embedded objects. It detects default ProgIDs from source extensions, maps embedded package part types, classifies Office files versus generic object payloads, creates embedded parts on Word, Excel, or PowerPoint host parts, and populates canonical `DocumentNode` fields from an embedded part [@ole-helper]. It also validates ProgIDs, normalizes display modes, creates placeholder icon images, supports data URI payloads, and warns about unknown OLE properties [@ole-helper].

The helper also handles failure-sensitive embedding details. It permits self-embedding by writing a placeholder payload when the host package lock would otherwise prevent reading the source file, wraps generic object payloads in an Ole10Native compound-file container, and deletes newly-created parts if feeding bytes fails [@ole-helper]. Those behaviors keep OLE semantics consistent across handlers and reduce the chance of orphan parts.

## Formulas

Formula helpers cover two different domains. `FormulaParser` converts a LaTeX subset to Office Math Markup Language and can convert OMML back to readable or LaTeX text [@formula-core]. It records unrecognized LaTeX commands for user-facing warnings, guards recursion depth, flattens nested Office Math, and has a lenient parse mode that writes a valid placeholder instead of failing an entire batch [@formula-core].

Excel formulas use `FormulaEvaluator` and related partial files. The evaluator represents numeric, string, boolean, error, array, range, blank, and lambda results; it shares evaluation session state for cross-sheet references, memoized formula-cell results, range materialization, row and column extents, and circular-reference tracking [@formula-core]. `ModernFunctionQualifier` prefixes modern Excel functions with `_xlfn.` or `_xlfn._xlws.` when emitting OOXML, and identifies formulas that need dynamic-array spill metadata [@formula-core].

Together, these helpers form the rich-object map for maintainers. Handler code should delegate shared object rules here, then keep host-specific placement, anchoring, and path resolution inside the handler. The intended lookup companion is the rich object helper map reference.
