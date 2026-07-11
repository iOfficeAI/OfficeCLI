---
title: "Batch Execution"
summary: "Batch execution normalizes a JSON array of command items, replays them through one shared item dispatcher, and formats per-item results consistently across CLI, resident, MCP, SDK, and in-process callers."
topics: [architecture, cli, batch]
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
  - id: command-builder
    type: file
    path: src/officecli/CommandBuilder.cs
  - id: resident-server
    type: file
    path: src/officecli/ResidentServer.cs
---

Batch execution is the shared replay path for many OfficeCLI operations packed into one JSON array. The CLI accepts batch input from `--commands`, `--input`, or stdin, validates the item shape, opens or routes to the target document once, then calls the same item dispatcher used by resident mode and in-process embedding. This keeps command semantics, result formatting, and continue-on-error behavior aligned across transports while avoiding repeated open/save cycles [@batch-command] [@batch-executor] [@resident-server].

## Batch Item Shape

A batch item is an object whose `command` is a bare verb such as `add`, `set`, `remove`, `move`, `swap`, `get`, or `query`. Arguments are sibling fields like `path`, `parent`, `type`, `props`, `selector`, `to`, and `path2`, not a whole CLI command line stuffed into `command` [@batch-command] [@batch-types].

`BatchItemConverter` accepts both `command` and `op`, accepts `props` either as an object or as an array of `key=value` strings, and ignores unknown JSON properties during normal deserialization [@batch-types]. The CLI batch command adds a stronger pre-validation layer before deserialization: each object is scanned for unknown fields and rejected with a list of valid fields, so mistakes are caught before replay begins [@batch-command].

The item model also converts directly to `ResidentRequest`. This is the bridge that lets the same batch-shaped object travel through the CLI, resident pipe, MCP command surface, SDKs, and in-process callers without inventing a second command vocabulary [@batch-types] [@batch-executor].

## Input Normalization

The CLI enforces one primary input source. `--commands` and `--input` are mutually exclusive, and redirected stdin produces a warning when it would otherwise be ignored; `--input -` explicitly means stdin [@batch-command]. Stdin input has a UTF-8 BOM stripped so piped input behaves like `File.ReadAllText()` on an input file [@batch-command].

The command also unwraps the normal JSON output envelope from `dump --json` when the root object has a `data` array. That lets a dump result feed back into batch without requiring an extra transformation step [@batch-command]. Non-array roots are rejected with a clear error that tells the caller to wrap a single item in an array, and explicit `null` items are rejected before replay [@batch-command].

## Replay Semantics

`ApplyBatchItems()` is the central loop. It iterates items, calls `ExecuteBatchItem()`, records a `BatchResult` for each item, and either stops at the first failure or continues according to `stopOnError` [@batch-command]. The default CLI behavior is continue-on-error; `--stop-on-error` opts back into strict abort-on-first-failure mode [@batch-command].

Each result records the item index, success flag, output, and error. Failed rows also include the original item so an agent can inspect and retry the failed command [@batch-types]. JSON output writes valid object or array output as raw JSON rather than double-encoding it as a string [@batch-types].

`ExecuteBatchItem()` applies the same command verbs that single-command surfaces expose, but it first rejects NUL bytes in caller-controlled strings and props. That boundary check prevents invalid XML characters from reaching the OOXML save path after earlier batch items have already mutated the document [@command-builder].

## File Lifetime

Non-resident batch opens the file once, checks document protection once against the live in-memory DOM, runs the shared replay loop with save deferral, and relies on handler disposal to finalize and save once [@batch-command]. For Word handlers, this avoids serializing the document after every item and keeps large replays from becoming dominated by repeated whole-document saves [@batch-command].

If a resident already owns the file, the CLI sends the whole batch as a single `batch` request on the resident pipe. The resident applies the items in memory and defers the disk flush to `save`, `close`, idle autosave, or the configured per-command flush policy [@batch-command] [@resident-server]. In resident mode, embedded `open` and `close` items are skipped because the resident already holds the document [@batch-command].

The public `BatchExecutor` gives in-process hosts the same behavior without spawning a CLI process. It deserializes the same item array, calls the same non-resident replay helper, and returns the same text or JSON envelope that the CLI would write to stdout [@batch-executor].

## Output And Exit Contract

Batch is treated as a judgment command. In JSON mode, the outer envelope succeeds only when every item succeeds; per-item success remains available under the result list [@batch-command]. Text mode prints each item result and a batch summary, while JSON mode wraps the formatted result body in the standard `{success, data, warnings}` envelope [@batch-command].

Warnings for unrecognized LaTeX commands are collected per item and surfaced after the replay, matching the single-command add/set behavior. If any item fails, exit code `1` takes precedence; if the only issue is an unrecognized-LaTeX warning, the command exits `2` [@batch-command]. Resident batch mirrors this by tracking whether the most recent batch had any failed row and mapping that verdict back into the response exit code and JSON envelope [@resident-server].

## Consequences

The batch architecture makes batch a protocol, not just a convenience flag. A caller can use the same item fields through the standalone CLI, a live resident, MCP, SDK-style send/batch calls, or direct embedding, and the replay rules stay tied to one implementation [@batch-types] [@batch-command] [@batch-executor]. The resident-specific flush behavior is explained in [Resident Process And Pipes](../runtime/resident-process-and-pipes).
