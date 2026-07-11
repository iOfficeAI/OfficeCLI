---
title: "Resident Save Close Flush"
summary: "When to use resident open, save, close, and flush settings so in-memory edits become visible on disk."
topics: [guides, resident]
sources:
  - id: command-builder
    type: file
    path: src/officecli/CommandBuilder.cs
  - id: save-command
    type: file
    path: src/officecli/CommandBuilder.Save.cs
  - id: flush-policy
    type: file
    path: src/officecli/Core/ResidentFlushPolicy.cs
  - id: resident-server
    type: file
    path: src/officecli/ResidentServer.cs
  - id: readme
    type: file
    path: README.md
---

A resident session keeps one document open in memory so repeated commands avoid reopening the Office package. Use `open` to start or reuse that session, `save` to flush in-memory edits while keeping it warm, and `close` to flush and stop it [@command-builder] [@save-command]. The successful outcome is simple: OfficeCLI commands see current edits immediately, and non-OfficeCLI readers see current disk bytes after an explicit or automatic flush.

## Open The Session

Run `officecli open <file>` before a long edit sequence. If a resident is already running, `open` reuses it and upgrades its idle timeout to the normal 12-minute window [@command-builder]. `create` can also auto-start a short-lived resident; a later `open` turns that into the longer-lived session [@command-builder].

Keep passing the file path on every command. The CLI uses that path to find the per-file resident pipe, so resident mode changes transport and lifetime, not the command vocabulary [@command-builder].

## Save Before External Reads

Run `officecli save <file>` when another program needs to read the package directly while the resident should stay alive. The save command probes for an existing resident instead of auto-starting one, sends a resident `save` request, and returns the resident's stdout, stderr, and exit code [@save-command].

Calling `save` with no resident is safe. In non-resident mode each mutation has already saved to disk, so the command reports that the file is already saved and exits successfully [@save-command]. Inside the resident, `save` skips work when there are no pending changes and otherwise calls the handler save, clears the dirty flag, records save duration, and prints `Saved <file>` [@resident-server].

## Close When Ownership Ends

Run `officecli close <file>` when the editing session is done. A live resident receives a close request, flushes through ordered shutdown, releases the handler, and stops responding only after the handler has disposed [@command-builder] [@resident-server]. If no resident owns the file, `close` is also a successful no-op because the disk is already current in non-resident mode [@command-builder].

Treat close warnings as important. If the resident reports an error during shutdown, the CLI returns that error; if the backing path vanished after a successful dispose, the resident emits a warning explaining that a rename may be safe but a delete may lose changes [@command-builder] [@resident-server].

## Tune Flush Policy

The default flush policy is `auto`: after mutations, an idle resident autosaves after an adaptive debounce between 2 and 10 seconds, based on measured save duration [@flush-policy] [@resident-server]. This makes direct disk readers see recent edits shortly after the session goes idle without paying a save after every command [@resident-server].

Set `OFFICECLI_RESIDENT_FLUSH=each` when deterministic disk visibility matters more than throughput. Set it to an integer number of seconds for a fixed idle debounce, or to `off` or `0` when only explicit `save`, `close`, or shutdown should write the disk snapshot [@flush-policy] [@resident-server]. The legacy `OFFICECLI_RESIDENT_IDLE_SAVE_SECONDS` variable is still honored when the newer flush variable is unset [@resident-server].

## Verify The Flush

Use an OfficeCLI read, such as `get`, `query`, or `view`, to verify logical content while the resident is active; these commands read the resident's in-memory handler [@resident-server]. Use `save` or `close` before checking with third-party tools such as Office, Python package readers, or external renderers, because those tools read disk bytes [@resident-server].

The README's quick-start sequence ends edits with `close`, which is the safest final handoff step because it both saves and releases the resident-owned file [@readme]. For deeper background, see [resident sessions](../concepts/resident-sessions) and [resident process and pipes](../architecture/runtime/resident-process-and-pipes).
