---
title: "Document Handler Lifecycle"
summary: "How OfficeCLI opens Office files, chooses native or plugin handlers, and exposes one document contract to commands."
topics: [architecture, handlers]
sources:
  - id: factory
    type: file
    path: src/officecli/Handlers/DocumentHandlerFactory.cs
  - id: contract
    type: file
    path: src/officecli/Core/IDocumentHandler.cs
  - id: word-handler
    type: file
    path: src/officecli/Handlers/WordHandler.cs
  - id: excel-handler
    type: file
    path: src/officecli/Handlers/ExcelHandler.cs
  - id: ppt-handler
    type: file
    path: src/officecli/Handlers/PowerPointHandler.cs
---

The document handler lifecycle is the path from a file name to a live object that commands can query, mutate, render, save, and dispose. `DocumentHandlerFactory.Open` is the shared entry point: it validates the path, rejects unsafe packages, repairs a few known producer defects, then returns either a native Word, Excel, or PowerPoint handler, or a plugin-backed handler for non-native formats [@factory]. The returned object implements `IDocumentHandler`, which is the common command contract for semantic views, structured document nodes, raw XML access, validation, binary extraction, and save [@contract]. This shape matters because higher layers can talk in [command layers](../../concepts/command-layers) without owning Word, Excel, or PowerPoint package details.

## Opening boundary

The factory treats an empty path, a missing file, and a zero-byte file as separate user-facing failures before any Open XML package is opened [@factory]. For native `.docx`, `.xlsx`, and `.pptx` files, it also inspects the zip directory for entry count, uncompressed size, and compression ratio, rejecting packages that look like decompression bombs [@factory]. These checks happen before handler construction, so every command path inherits the same safety behavior.

The factory also performs two repair steps before and during open. It scans native packages for dangling internal relationships and strips relationships whose target part is missing, because some Office producers leave broken relationships that Microsoft Office tolerates but the SDK can reject or crash on later traversal [@factory]. It also catches unsupported XML encoding declarations, rewrites them to UTF-8, and retries the open [@factory]. These repairs are intentionally centralized so individual handlers do not each carry their own open-time compatibility policy.

## Handler dispatch

After validation, extension dispatch is simple: `.docx` creates `WordHandler`, `.xlsx` creates `ExcelHandler`, and `.pptx` creates `PowerPointHandler` [@factory]. Unknown extensions are not immediately fatal. The factory first asks the plugin registry for a dump-reader plugin, which can convert a foreign source into a native sibling file, or for a format-handler plugin, which is wrapped in a proxy session [@factory]. If no plugin handles the extension, the factory returns an unsupported-type error that names the native extensions and points users toward plugins [@factory]. That is the bridge from native handler lifecycle to the [plugin system](../plugins/plugin-system).

## Shared contract

`IDocumentHandler` divides handler behavior into three layers. The semantic layer returns text, annotated text, outline, stats, issues, and JSON view variants [@contract]. The structured layer exposes `Get`, `Query`, `Set`, `Add`, `Remove`, `Move`, and `CopyFrom` over `DocumentNode` paths and selectors [@contract]. The raw layer exposes package parts through `Raw`, `RawSet`, and `AddPart`, with validation, binary extraction, and `Save` alongside them [@contract]. This is the same conceptual split described by [raw XML access](../../concepts/raw-xml-access) and [document nodes and paths](../../concepts/document-node-and-paths).

The native handlers all implement this contract and also expose their parsed SDK package to the rendering layer through a narrow render-model interface [@word-handler] [@excel-handler] [@ppt-handler]. Word wraps a `WordprocessingDocument`, Excel wraps a `SpreadsheetDocument`, and PowerPoint wraps a `PresentationDocument` [@word-handler] [@excel-handler] [@ppt-handler]. The command layer therefore sees one handler interface, while each handler keeps format-specific package state private.

## Save and disposal

Editable handlers open package streams with sharing choices that let external readers observe saved snapshots while a session is alive [@word-handler] [@excel-handler] [@ppt-handler]. Mutating paths mark the handler as modified, and `Save` or `Dispose` flushes package changes and stamps the OfficeCLI audit metadata only when a document was actually changed [@word-handler] [@excel-handler] [@ppt-handler]. Word also has a deferred-save mode used by batch replay so many mutations serialize once rather than after every operation [@word-handler].

The lifecycle consequence is that command code should open through the factory, perform work through `IDocumentHandler`, call `Save` only when it needs an explicit mid-session flush, and otherwise rely on disposal for final cleanup. Format-specific details belong inside the handler, while the factory owns file admission, repair, and handler selection.
