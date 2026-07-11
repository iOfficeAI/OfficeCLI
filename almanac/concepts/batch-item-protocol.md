---
title: "Batch Item Protocol"
summary: "The batch item protocol is OfficeCLI's shared JSON command object for multi-step CLI batches, resident RPC, MCP execution, SDKs, and in-process embedding."
topics: [concepts, batch, cli, resident, sdk]
sources:
  - id: batch-types
    type: file
    path: src/officecli/BatchTypes.cs
  - id: batch-command
    type: file
    path: src/officecli/CommandBuilder.Batch.cs
  - id: batch-executor
    type: file
    path: src/officecli/Core/BatchExecutor.cs
---

The batch item protocol is the JSON shape OfficeCLI uses when a command must be represented as data instead of argv. A batch item has a bare verb in `command` or `op`, and the verb arguments live beside it as fields such as `path`, `parent`, `type`, `props`, `selector`, `text`, `mode`, `part`, `xpath`, `action`, and `xml` [@batch-types]. This is the shared object behind CLI batch files, resident batch delivery, MCP and SDK-style command objects, and in-process embedding.

The protocol exists to prevent every transport from inventing its own command vocabulary. `BatchItem` can convert itself into a `ResidentRequest`, carrying scalar arguments through `Args` and property dictionaries through `Props` [@batch-types]. The in-process `BatchExecutor` also accepts the same JSON array and returns the same output shape as the CLI batch command [@batch-executor].

## Item Shape

A batch item is not a shell command string. The batch command help explicitly calls out that `command` must be the bare verb, while arguments are sibling fields such as `parent`, `path`, `selector`, `type`, `props`, `to`, `after`, `before`, and `path2` [@batch-command].

The JSON reader is intentionally lenient where that helps agents. It accepts `op` as an alias for `command`, accepts `path` as a query selector alias in the executor path, accepts `path2` or legacy `to` for swap, and accepts `props` either as an object or as an array of `key=value` strings [@batch-types].

## Shared Replay

The central replay loop is `ApplyBatchItems`. It walks the item list, executes each item, captures per-item success or error, and stops early only when `stopOnError` is true [@batch-command]. The CLI non-resident path, MCP-style batch surface, and resident server all reuse this loop so failure handling does not drift between transports [@batch-command].

That shared replay model is why [batch execution](../architecture/cli/batch-execution), [using batch safely](../guides/using-batch-safely), and [batch item fields](../reference/batch-item-fields) all describe the same protocol from different angles. The concept is the object model; the architecture page explains runtime flow; the reference page names every field.

## Save Boundaries

A batch groups mutations so the document can be opened once and saved once in the non-resident path [@batch-command]. When a resident process already owns the file, the CLI sends the whole batch to the resident as one `batch` request, and the resident applies the items in memory while disk flushing is deferred to resident save, close, idle autosave, or an explicit flush policy [@batch-command].

This is the main behavioral difference between a batch item and an argv command. The item describes the same operation, but the surrounding batch controls replay, error collection, and save timing.
