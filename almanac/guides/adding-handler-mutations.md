---
title: "Adding Handler Mutations"
summary: "How to add or change add, set, and remove behavior while keeping handlers, command warnings, and schemas aligned."
topics: [guides, handlers]
sources:
  - id: command-set
    type: file
    path: src/officecli/CommandBuilder.Set.cs
  - id: command-add
    type: file
    path: src/officecli/CommandBuilder.Add.cs
  - id: word-set
    type: file
    path: src/officecli/Handlers/Word/WordHandler.Set.cs
  - id: excel-set
    type: file
    path: src/officecli/Handlers/Excel/ExcelHandler.Set.cs
  - id: ppt-set
    type: file
    path: src/officecli/Handlers/Pptx/PowerPointHandler.Set.cs
  - id: word-add
    type: file
    path: src/officecli/Handlers/Word/WordHandler.Add.cs
  - id: excel-add
    type: file
    path: src/officecli/Handlers/Excel/ExcelHandler.Add.cs
  - id: excel-remove
    type: file
    path: src/officecli/Handlers/Excel/ExcelHandler.Remove.cs
  - id: schemas-readme
    type: file
    path: schemas/README.md
---

Add or change a handler mutation when OfficeCLI needs a curated operation instead of raw XML editing. The safe path is to put format-specific behavior inside the Word, Excel, or PowerPoint handler, let the command layer keep parsing, guard, and warning behavior, and update the matching help schema in the same change [@command-add] [@command-set] [@schemas-readme]. The successful outcome is that direct CLI, resident, batch, and schema-driven agent help all describe the same mutation.

## Start At The Command Contract

Use the existing command shape unless you are intentionally changing the public CLI. `add` owns parent path, type, copy-from, insert position, properties, protection checks, resident forwarding, unsupported-property warnings, and watch notifications [@command-add]. `remove` shares the same command file and owns path parsing, optional Excel `--shift`, selector guardrails, resident forwarding, handler dispatch, and watch notification [@command-add]. `set` owns path, `--find` and `--replace` sugar, missing-property errors, protection checks, selector guardrails, resident forwarding, unsupported-property correction, and result warnings [@command-set].

The command layer should not learn the OOXML details of the new operation. It should parse user input, pass a property dictionary to the handler, and report what the handler says was unsupported [@command-add] [@command-set]. This keeps behavior aligned with the [document handler lifecycle](../architecture/handlers/document-handler-lifecycle).

## Put Dispatch Near The Format

Add the mutation to the handler that owns the file format. Word `Add` resolves the parent, validates parent-child legality, resolves insert anchors, and dispatches by element type [@word-add]. Excel `Add` normalizes paths, rejects wrong-format element types early, and dispatches by type such as sheet, row, cell, picture, chart, pivot table, or slicer [@excel-add].

For `set`, follow the existing order. Word routes revision selectors, selector-style sets, range formatting, find/replace, and then concrete path dispatch [@word-set]. Excel normalizes sheet paths, handles selector sets, validates find/replace, and then dispatches by workbook, sheet, cell, table, chart, pivot, and related path forms [@excel-set]. PowerPoint normalizes path casing, id paths, and last predicates before selector sets, range formatting, find/replace, presentation-level properties, and slide or shape paths [@ppt-set].

## Preserve Selector And Path Semantics

Do not bypass the selector bridge for broad edits. The handlers route bare selectors or slash-scoped content filters through the same filtering engine used by query, then call `Set` on each matched concrete path [@word-set] [@excel-set] [@ppt-set]. Excel remove follows the same query-to-concrete-path pattern and orders row removals descending so shift-deletes do not retarget later rows [@excel-remove]. This is why selector edits can report a match count or fail clearly on zero matches instead of silently doing nothing [@command-set] [@excel-remove].

Use exact slash paths for deterministic node edits and selectors for discovery or multi-target edits. The distinction is explained in [selectors vs slash paths](../concepts/selectors-vs-slash-paths); new mutation code should preserve it instead of inventing a third addressing model.

## Report Unsupported Properties Honestly

Handlers are the source of truth for supported properties. `add` wraps properties in a tracking dictionary and reports keys the handler never reads; Word can also add handler-internal unsupported properties for helpers that must iterate the dictionary [@command-add]. `set` applies through a shared correction path and returns unsupported properties with scoped suggestions based on handler type [@command-set].

When adding a property, make sure the handler actually consumes it. If the operation must reject a combination, throw a targeted error near the dispatcher, as existing Word and PowerPoint code do for contradictory `find` and `range` requests [@word-set] [@ppt-set]. If a property is accepted only for some paths, return it as unsupported for the other paths so callers see a warning instead of a false success [@command-add] [@command-set].

## Update Schemas And Verify

Update the matching file under `schemas/help/` in the same change. The schema README states that any PR changing `Add`, `Set`, or `Get` behavior for an element must update the matching schema, and contract tests verify schema claims against the real handler implementation [@schemas-readme].

Verify through the command surface, not only a helper method. Exercise the new direct command, the same operation through resident mode, and the same item in a batch when the mutation should be batch-addressable. Then run a readback command such as `get`, `query`, `view`, or `validate` so the saved document and the agent-facing result agree with the handler behavior [@command-add] [@command-set].
