---
title: "Dump Reader Buffered Replay"
summary: "Dump-reader plugins must stream JSONL, but OfficeCLI buffers those lines before replay so OpenXML mutations run synchronously on the caller thread."
topics: [decisions, plugins, batch]
sources:
  - id: dump-invoker
    type: file
    path: src/officecli/Core/Plugins/DumpReaderInvoker.cs
  - id: plugin-protocol
    type: file
    path: plugins/plugin-protocol.md
---

OfficeCLI keeps the dump-reader protocol streaming but buffers the received JSONL lines before replaying them into a native document. This preserves the plugin contract and idle-watchdog activity signal while avoiding unsafe OpenXML mutation from the stdout reader thread. Future dump-reader work must keep plugin output line-oriented, but replay should remain on the caller thread unless the document package layer is made safe for concurrent mutation [@dump-invoker] [@plugin-protocol].

## Context

A dump-reader converts a foreign source file into OfficeCLI batch items. The protocol requires one JSON object per line on stdout, with each line matching a `batch --commands` item [@plugin-protocol]. Top-level JSON arrays are rejected with `corrupt_batch` [@plugin-protocol].

The streaming requirement exists for two reasons: it gives the plugin watchdog per-item activity, and it avoids making the plugin protocol depend on one large JSON array [@plugin-protocol]. The broader flow is described in [Dump Reader And Exporter Plugins](../architecture/plugins/dump-reader-and-exporter-plugins).

## Decision

`DumpReaderInvoker.Run` still invokes the plugin as a streaming producer and processes stdout line by line [@dump-invoker]. Each non-empty line is trimmed, a per-line UTF-8 BOM is tolerated, and a line beginning with `[` is rejected as the old JSON-array shape [@dump-invoker].

Instead of replaying each line inside the stdout callback, the invoker appends valid raw lines to an in-memory list. After the plugin exits successfully, the invoker opens the generated native temp file on the caller thread and deserializes and replays each buffered line through `CommandBuilder.ExecuteBatchItem` [@dump-invoker].

The reason is package safety. The previous design replayed from the process stdout reader task. The code records that heavy OpenXML updates from that background thread could trip non-thread-safe package state, especially with many parts being created or reopened in update mode [@dump-invoker].

## Status

This is the current implementation in `DumpReaderInvoker.cs`. The plugin still streams. OfficeCLI still uses the plugin process runner and idle timeout while lines arrive. Replay is deferred until after plugin exit [@dump-invoker].

## Consequences

The decision makes replay behavior closer to normal batch execution, where items are replayed synchronously through one handler [@dump-invoker]. It also keeps item-indexed errors for invalid JSON, null items, and failed commands during replay [@dump-invoker].

The cost is host memory proportional to the number and size of emitted JSONL lines. That weakens one original benefit of full streaming replay, so dump-reader authors should still emit compact batch items and avoid unnecessary payload bloat [@dump-invoker] [@plugin-protocol].

The boundary remains clear: dump-reader plugins own the foreign source and emit commands; OfficeCLI owns the target native file and the replay. Edits after conversion belong to the native output, not to the original foreign file [@plugin-protocol] [@dump-invoker].
