---
title: "Rendering And Preview Stack"
summary: "How OfficeCLI turns live handlers into HTML, SVG, screenshots, and preview assets through a renderer registry."
topics: [architecture, handlers, rendering, preview]
sources:
  - id: rendering-core
    type: file
    path: src/officecli/Core/Rendering/
  - id: handler-rendering
    type: file
    path: src/officecli/Handlers/Rendering/
  - id: view-command
    type: file
    path: src/officecli/CommandBuilder.View.cs
  - id: preview-css
    type: file
    path: src/officecli/Resources/preview.css
  - id: preview-js
    type: file
    path: src/officecli/Resources/preview.js
---

The rendering and preview stack converts an already-open document handler into visual artifacts. Its center is a renderer registry: renderers advertise supported formats, output kinds, watch support, and priority, and the registry selects the highest-priority available renderer that covers a request [@rendering-core]. The built-in adapters forward to the existing Word, Excel, and PowerPoint preview methods, but they receive the parsed handler model through `HandlerRenderInput` rather than reopening files [@handler-rendering]. That decision keeps rendering inside the [document handler lifecycle](document-handler-lifecycle) while leaving room for alternate renderers.

## Registry contract

Rendering is artifact-oriented. `RenderOutputKind` names HTML, SVG, PNG, and PDF outputs, while `RenderOptions` carries page ranges, Word page filters, grid settings, viewport size, and raster dimensions [@rendering-core]. `IRenderInput` exposes the format id, the parsed model, and read-only `Get` and `Query` access to the document tree [@rendering-core]. A renderer can therefore work from the live in-memory document instead of re-parsing a file path.

The built-in renderer bootstrap registers three low-priority renderers: Word HTML, Excel HTML, and PowerPoint HTML/SVG [@handler-rendering]. Those adapters are intentionally thin. Word calls `ViewAsHtml` with page and grid options, Excel calls `ViewAsHtml`, and PowerPoint calls either `ViewAsHtml` or `ViewAsSvg` depending on the requested output [@handler-rendering]. Higher-priority renderers can register during registry composition and replace the built-in behavior without changing command dispatch [@handler-rendering].

## View command flow

`view html` opens a handler through the handler factory, determines the format, and calls `RenderViaRegistry` for Word, Excel, or PowerPoint [@view-command]. If the caller asks for a browser or output file, the command writes HTML to either the requested path or an unpredictable temporary path before opening it [@view-command]. `view svg` uses the same registry path for PowerPoint SVG and falls back to plugin SVG when a format-handler plugin can provide it [@view-command].

`view screenshot` builds on the same preview stack. It first tries native Word or PowerPoint raster backends on Windows when allowed by `--render`, then asks the renderer registry for PNG bytes, and finally falls back to HTML captured by a headless browser [@view-command]. Screenshot mode can crop to rendered `data-path` targets, tile PowerPoint or Word pages into a grid, and size single-slide or single-page captures to their native 96 DPI pixel dimensions [@view-command].

## Preview assets

The preview HTML is supported by embedded static resources. `preview.css` defines the shared slide viewer layout, sidebar thumbnails, main scrolling area, fullscreen state, headless capture behavior, and responsive sidebar behavior [@preview-css]. `preview.js` implements slide scaling, thumbnail construction, fullscreen navigation, keyboard navigation, and sidebar toggling [@preview-js]. These assets are not the renderer API, but they are part of the default visual contract because built-in PowerPoint previews rely on their classes and browser behavior.

The stack is also the input to the [watch preview contract](watch-preview-contract). Watch uses rendered HTML snapshots and patches, not a second renderer, so the same renderer output supports one-shot `view`, screenshots, and live browser refresh [@rendering-core].
