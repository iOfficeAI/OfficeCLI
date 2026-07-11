---
title: "Raw XML Access"
summary: "Raw XML access is OfficeCLI's lowest command layer, exposing Office package parts for inspection and last-resort mutation when curated operations are not enough."
topics: [concepts, cli, handlers]
sources:
  - id: raw-command
    type: file
    path: src/officecli/CommandBuilder.Raw.cs
  - id: raw-helper
    type: file
    path: src/officecli/Core/RawXmlHelper.cs
  - id: excel-handler
    type: file
    path: src/officecli/Handlers/ExcelHandler.cs
  - id: ppt-handler
    type: file
    path: src/officecli/Handlers/PowerPointHandler.cs
---

Raw XML access is the universal fallback under OfficeCLI's curated document commands. The `raw` command reads XML from a document part, `raw-set` applies an XPath-based mutation to a part, and `add-part` creates a new Office package part for later raw editing [@raw-command]. This layer exists because not every OpenXML feature has a first-class `add`, `set`, or `remove` operation, but the file format is still XML inside an Office package.

## The Command Layer

`raw` takes a file and a part path, with optional row and column filters for Excel sheets [@raw-command]. `raw-set` takes a file, part path, XPath, action, and optional XML fragment or attribute assignment [@raw-command]. Its accepted action tokens are `append`, `prepend`, `insertbefore`, `insertafter`, `replace`, `remove`, and `setattr` [@raw-helper].

The command builder still uses the normal OfficeCLI safety frame. It tries resident delivery first, opens the handler directly when no resident owns the file, validates before and after mutation, reports new validation problems as warnings, and notifies live watch previews after a raw edit [@raw-command]. That makes raw access a low-level command layer, not a separate toolchain.

## Semantic Parts And Zip URIs

Raw part names can be semantic or literal. Semantic names are handler-defined shortcuts such as `/workbook`, `/styles`, `/Sheet1`, `/presentation`, `/slide[1]`, `/slideMaster[1]`, or `/notesMaster` [@excel-handler] [@ppt-handler]. Literal zip-URI paths are paths ending in `.xml` or `.rels`, such as `/xl/worksheets/sheet1.xml` or `/ppt/slides/slide1.xml`; `RawXmlHelper` treats those as package-internal URIs and resolves them across the package part graph [@raw-helper].

This design replaced incomplete per-handler alias tables. The helper comments state the rule directly: semantic short names still route through handler switches, while `.xml` and `.rels` paths use zip-URI lookup so arbitrary package parts can be reached [@raw-helper].

## XML Mutation Semantics

Raw XML mutation is XPath-based. `RawXmlHelper` parses the target XML, registers common OpenXML namespace prefixes, finds matching elements, applies the requested action, and writes the modified XML back to the OpenXML part [@raw-helper]. If the XPath matches no elements, raw-set raises an error rather than returning a successful no-op, because silent raw mutations are dangerous in batch and resident flows [@raw-helper].

The helper also handles OpenXML details that ordinary XML string replacement would miss. It preserves namespace declarations for SDK round-tripping, copies root attributes that the `InnerXml` setter would otherwise drop, normalizes whitespace-only leaf text so meaningful spaces survive, and reorders known PowerPoint containers into schema order after raw appends [@raw-helper]. It also updates markup-compatibility `mc:Ignorable` data when alternate-content choices require extension prefixes [@raw-helper].

## When To Use It

Raw XML is related to [command layers](command-layers), [document handler lifecycle](../architecture/handlers/document-handler-lifecycle), and the [command surface](../reference/command-surface). It should be understood as the escape hatch below the semantic API. Curated commands are easier to validate and easier for agents to reason about. Raw access is for gaps: unusual OpenXML features, package parts without a curated wrapper, replay of advanced dumps, or precise repairs that the high-level handler cannot express.

The mental model is therefore layered. Use `get`, `query`, `add`, `set`, and `remove` when they cover the task. Use `raw` to inspect the underlying XML when the semantic view is insufficient. Use `raw-set` only when the caller knows the target part, XPath, and XML effect well enough to own the package-level consequences.
