---
title: "Command Layers"
summary: "OfficeCLI commands sit on semantic, document-node, and raw XML layers so agents can start with stable document operations and drop lower only when needed."
topics: [concepts, cli, handlers]
sources:
  - id: readme-purpose
    type: file
    path: README.md
  - id: handler-contract
    type: file
    path: src/officecli/Core/IDocumentHandler.cs
---

Command layers are the three levels at which OfficeCLI exposes an Office document: a semantic view layer for readable summaries, a document-node layer for structured operations, and a raw layer for direct OOXML access. The layers matter because OfficeCLI is built for agents that need to read, change, validate, and render Word, Excel, and PowerPoint files without Office installed, while still having an escape hatch when a curated operation is not enough [@readme-purpose].

The [command dispatch](../architecture/cli/command-dispatch) surface is easier to understand through this model. A command is not just a CLI verb. It also chooses how much of the document model the caller wants to handle.

## Semantic Layer

The semantic layer is for looking at a document without committing to its internal tree shape. The shared handler contract exposes text, annotated text, outline, stats, issue, and JSON variants as first-class view operations [@handler-contract]. These operations let callers inspect a document before deciding which node to edit.

This layer is intentionally high-level. It answers questions like "what does the document say?", "what sections or slides exist?", and "what problems are visible?". It is the safest layer for orientation because it does not require a path, selector, or XML part name.

## Document-Node Layer

The document-node layer is the normal editing layer. The same handler contract exposes `get`, `query`, `set`, `add`, `remove`, `move`, and `copyFrom` operations over paths, selectors, element types, insertion positions, and property dictionaries [@handler-contract].

This layer gives OfficeCLI its cross-format command vocabulary. A PowerPoint shape, a Word paragraph, and an Excel cell are different objects internally, but they can all be addressed as document nodes and changed through the same command family. The [document node and paths](document-node-and-paths) concept explains that address model in more detail.

The layer is also where most user-facing safety checks belong. A curated `set` operation can reject an unsupported property, a `remove` can return a warning, and an `add` can resolve an insertion position before anything touches raw XML [@handler-contract]. That keeps common edits inside the handler lifecycle rather than pushing callers into package internals.

## Raw Layer

The raw layer is the fallback below curated document operations. The handler contract includes raw part reads, XPath-based raw writes, part creation, validation, binary extraction, and save operations [@handler-contract].

This layer exists because OOXML is larger than the curated command surface. It lets maintainers and advanced agents reach specific package parts while still using the same file-opening and validation machinery described in the [document handler lifecycle](../architecture/handlers/document-handler-lifecycle). The [command surface](../reference/command-surface) reference should therefore be read as layered: semantic commands first, structured node mutations next, and raw XML only when the stable vocabulary does not cover the case.
