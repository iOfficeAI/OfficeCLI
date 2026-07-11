---
title: "Renderers Receive Parsed Models"
summary: "OfficeCLI renderers receive a live render input backed by the open handler instead of reopening documents from file paths."
topics: [decisions, rendering, handlers]
sources:
  - id: renderer-contract
    type: file
    path: src/officecli/Core/Rendering/IRenderer.cs
  - id: handler-input
    type: file
    path: src/officecli/Handlers/Rendering/HandlerRenderInput.cs
  - id: basic-renderers
    type: file
    path: src/officecli/Handlers/Rendering/BasicRenderers.cs
---

OfficeCLI renderers receive parsed models through `IRenderInput`, not document file paths. The renderer boundary exposes the format id, optional in-memory model, and read-only document-tree access, while `HandlerRenderInput` binds that boundary to a live document handler; future renderers should render from this live input instead of reopening files and building a second view of the document [@renderer-contract] [@handler-input].

## Context

Rendering sits downstream of the [document handler lifecycle](../architecture/handlers/document-handler-lifecycle). By the time a view, screenshot, or preview is requested, OfficeCLI already has a handler that owns the parsed document package and the current in-session document state. `IRenderInput` therefore defines `FormatId`, `Model`, `Get`, and `Query` as the renderer-facing input surface [@renderer-contract].

The handler adapter is deliberately narrower than the full mutation surface. `HandlerRenderInput` stores an `IDocumentHandler`, exposes the format id, projects `Model` from handlers that implement `IRenderModelHost`, and forwards `Get` and `Query` to the handler [@handler-input]. Its comment states the intent: renderers see read access and the parsed model, not the file path or mutation API [@handler-input].

This decision is the core constraint behind the [rendering and preview stack](../architecture/handlers/rendering-and-preview-stack).

## Decision

Renderers must accept `IRenderInput` and `RenderOptions`. The renderer contract says the model is the already parsed in-memory document and that `Get` reflects in-session edits, so renderers work from handler-owned state rather than re-reading a file [@renderer-contract].

Built-in renderers remain thin adapters over the existing handler preview methods. The Word renderer casts the input back to `HandlerRenderInput`, obtains the `WordHandler`, and calls `ViewAsHtml`; the Excel renderer does the same for `ExcelHandler`; the PowerPoint renderer calls `ViewAsHtml` or `ViewAsSvg` based on the requested output [@basic-renderers]. The adapters are registered as low-priority built-ins, leaving room for higher-priority renderers without changing the input decision [@basic-renderers].

## Status

This decision is active. `IRenderer.Render` takes `IRenderInput`, `HandlerRenderInput` bridges live handlers into that contract, and the built-in Word, Excel, and PowerPoint renderers all consume handler-backed inputs [@renderer-contract] [@handler-input] [@basic-renderers].

## Consequences

The main benefit is consistency with live document state. A renderer can use `Get` and `Query` against the current handler, and the contract explicitly says that this read access reflects in-session edits [@renderer-contract]. That avoids a second parser observing stale bytes or missing resident-session changes.

The boundary also keeps renderer plugins from depending on the mutation-facing handler API. Renderers receive a model object when one is attached and format-neutral read helpers, but the narrow contract is independent of any one concrete handler [@renderer-contract] [@handler-input].

The tradeoff is that built-in renderers that still need handler-specific preview methods must validate that the input is a `HandlerRenderInput` with the expected native handler type, and they throw if a different input is passed [@basic-renderers]. Future alternate renderers should prefer the `IRenderInput` surface directly when possible, and should treat direct file reopening as a regression unless a separate decision creates that path.
