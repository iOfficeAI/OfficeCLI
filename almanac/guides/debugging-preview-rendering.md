---
title: "Debugging Preview Rendering"
summary: "How to trace HTML, screenshot, SVG, and watch-preview rendering issues across the view command, renderer registry, handlers, and preview assets."
topics: [guides, rendering, preview]
sources:
  - id: view-command
    type: file
    path: src/officecli/CommandBuilder.View.cs
  - id: rendering-core
    type: file
    path: src/officecli/Core/Rendering/
  - id: handler-rendering
    type: file
    path: src/officecli/Handlers/Rendering/
  - id: word-preview
    type: file
    path: src/officecli/Handlers/Word/WordHandler.HtmlPreview.cs
  - id: excel-preview
    type: file
    path: src/officecli/Handlers/Excel/ExcelHandler.HtmlPreview.cs
  - id: ppt-preview
    type: file
    path: src/officecli/Handlers/Pptx/PowerPointHandler.HtmlPreview.cs
  - id: preview-css
    type: file
    path: src/officecli/Resources/preview.css
  - id: preview-js
    type: file
    path: src/officecli/Resources/preview.js
---

Debug preview rendering by following the artifact from `view` into the renderer registry and then into the format-specific HTML generator. `view html`, `view screenshot`, and `view svg` all start in `CommandBuilder.View.cs`; Word, Excel, and PowerPoint HTML normally go through registered basic renderers that forward to the handler preview methods [@view-command] [@handler-rendering]. The successful outcome is knowing whether the failure is command option parsing, renderer selection, handler markup, screenshot capture, or shared preview assets.

## Reproduce The Smallest Artifact

Start with `officecli view <file> html --out preview.html` for HTML problems. The command opens the handler, calls `RenderViaRegistry`, and writes either stdout, the requested output path, or an unpredictable temporary file when `--browser` needs one [@view-command].

For screenshot problems, first decide whether the issue exists in HTML. Screenshot mode renders the same HTML preview unless a native Windows backend or a PNG-capable renderer provides pixels first [@view-command]. If the HTML is correct but the screenshot is wrong, focus on viewport size, page filtering, grid mode, clipping, or headless browser capture [@view-command].

## Check Renderer Selection

The registry chooses the highest-priority available renderer whose capabilities cover the requested format, output kind, and mode [@rendering-core]. The built-in bootstrap registers Word HTML, Excel HTML, and PowerPoint HTML/SVG renderers at priority `0`, then gives out-of-tree renderers a composition hook [@handler-rendering].

If a renderer is not being called, inspect the requested `RenderOptions`. `RenderOutputKind` separates HTML, SVG, PNG, and PDF; `RenderMode` separates static output from watch output; page ranges, Word page filters, grid columns, viewport size, and raster dimensions travel through `RenderOptions` [@rendering-core]. A renderer that does not advertise the requested output or watch support will be skipped [@rendering-core].

## Inspect Handler Markup

Word preview builds a self-contained HTML document from the live Word model, catches malformed XML during preview, emits page and grid behavior, and places `data-path` attributes on rendered paragraphs and tables [@word-preview]. Excel preview renders visible sheets as spreadsheet tables, materializes style arrays for efficient lookup, and emits `data-path` attributes for column headers, row headers, and cells [@excel-preview]. PowerPoint preview renders slides as absolutely positioned HTML, supports page ranges and grid overrides, loads shared CSS and JS resources, and emits slide and shape path data used by interaction and clipping [@ppt-preview].

When a visual element is missing, compare the handler's `get` or `query` output with the HTML. If the node exists but has no `data-path`, clipping and watch selection may fail even though the pixels render [@view-command] [@word-preview] [@excel-preview] [@ppt-preview].

## Debug Screenshot And Clip Failures

Screenshot mode defaults to a bounded visual unit: PowerPoint slide 1, Word page 1, and Excel's active sheet behavior through the HTML preview [@view-command]. Use `--page 1-N` or `--grid` when the bug requires a multi-page or multi-slide capture [@view-command].

For `--range` clipping, use a path that the HTML preview actually emits. The command resolves clip targets to `data-path` selectors and returns a targeted error suggesting queryable paths when no rendered region matches [@view-command]. This makes the preview HTML the authority for clipped screenshots, not just the document tree.

If the error says no screenshot backend is available, install a supported headless browser or Playwright Chromium. The command's fallback path writes temporary HTML and captures it with `HtmlScreenshot`; if that backend is missing, no HTML-based screenshot can be produced [@view-command].

## Check Shared Preview Assets

For PowerPoint sidebar, thumbnail, fullscreen, scaling, or keyboard-navigation issues, inspect the embedded preview assets after confirming the handler emitted valid slide markup. The PowerPoint preview loads `Resources.preview.css` and `Resources.preview.js` through helper methods [@ppt-preview]. The CSS owns shared slide viewer layout and headless hiding rules, while the JavaScript owns scaling, thumbnail construction, navigation, and sidebar behavior [@preview-css] [@preview-js].

For live preview behavior, connect the rendering issue to [rendering and preview stack](../architecture/handlers/rendering-and-preview-stack) and [watch preview contract](../architecture/handlers/watch-preview-contract). Watch depends on the same rendered HTML contract, so a broken one-shot preview usually needs to be fixed before debugging the refresh loop.
