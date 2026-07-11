---
title: "Resident Process And Pipes"
summary: "The resident runtime keeps one document open behind per-file named pipes, serializes commands, and controls when in-memory changes are flushed to disk."
topics: [architecture, runtime, resident]
sources:
  - id: command-builder
    type: file
    path: src/officecli/CommandBuilder.cs
  - id: resident-server
    type: file
    path: src/officecli/ResidentServer.cs
  - id: resident-client
    type: file
    path: src/officecli/ResidentClient.cs
  - id: flush-policy
    type: file
    path: src/officecli/Core/ResidentFlushPolicy.cs
---

The resident runtime is OfficeCLI's long-lived document session. A resident process owns one file, keeps its handler in memory, exposes a main named pipe for commands, and exposes a separate ping pipe for liveness, close, and idle-timeout control. This avoids repeated file open/save cost while preserving one-writer ownership: clients must route through the resident when it is alive, and direct file access is used only when no resident owns the file [@command-builder] [@resident-server] [@resident-client].

## Process Ownership

`open` starts a resident by spawning the current executable with the hidden `__resident-serve__` command. The child is given the target file path through `ProcessStartInfo.ArgumentList`, and the parent waits up to five seconds for the ping pipe to respond before reporting success [@command-builder].

The server command takes a per-file singleton lock in the temp directory before opening the document. The lock name is derived from the same pipe name as the resident, and a competing process exits quietly if another resident already starts serving the file [@command-builder]. This prevents multiple resident processes from holding separate in-memory copies and overwriting each other on flush [@command-builder].

Pipe names are based on a SHA-256 hash of the full path, with path casing normalized on Windows and macOS. The main pipe is named `officecli-<hash>` and the ping pipe adds `-ping` [@resident-server] [@resident-client].

## Two-Pipe Protocol

The ping pipe is intentionally separate from the main command pipe. `ResidentClient.TryConnect()` talks to `officecli-<hash>-ping`, sends `__ping__`, and verifies that the returned full path matches the requested file [@resident-client]. The ping pipe also serves `__set-idle-timeout__` and `__close__`, so these control operations remain responsive even when the main command queue is busy [@resident-server] [@resident-client].

The main pipe carries business commands as one JSON request per connection. The server pre-creates the next pipe instance before handing an accepted connection to a worker, which reduces connection gaps during bursts of clients [@resident-server]. Actual command execution is serialized by `_commandLock`, so only one handler operation mutates or reads the in-memory document at a time [@resident-server].

The client retries only during the connect phase. Once a request has been written, it does not resend on broken replies or bad JSON, because the resident may already have applied a non-idempotent mutation such as add, remove, move, swap, or batch [@resident-client]. That gives the pipe protocol at-most-once delivery after write, with visible failure preferred over silent double-application [@resident-client].

## Dispatch And Mutation State

The resident opens handlers read-only by default. The first mutating command promotes the handler to editable mode, enables Word save deferral where applicable, and marks the session dirty [@resident-server]. Mutating commands include set, add, remove, move, swap, refresh, raw-set, add-part, and batch; read-only commands such as get, query, view, raw, validate, and dump can run without promoting the handler [@resident-server].

`CommandBuilder.TryResident()` is the normal front door for single commands. It probes the ping pipe first. If no resident is found, it may auto-start a short-lived resident unless `OFFICECLI_NO_AUTO_RESIDENT` disables that behavior. If a resident is alive but the main pipe cannot accept the command after longer retries, it returns a busy error instead of falling back to direct file access [@command-builder].

Batch uses the same ownership rule. When a resident is alive, the CLI sends the whole batch as one `batch` request, and the resident applies the items against the in-memory handler [@command-builder] [@resident-server]. That path is described from the CLI side in [Batch Execution](../cli/batch-execution).

## Flush And Shutdown

A resident separates in-memory visibility from on-disk visibility. OfficeCLI commands routed through the resident read the live in-memory handler immediately, but external programs only see changes after a save, close, idle autosave, shutdown, or `each` flush [@resident-server] [@flush-policy].

The flush policy is controlled by `OFFICECLI_RESIDENT_FLUSH`, with a legacy fallback environment variable. Supported modes are `each`, `auto`, fixed seconds, and `off` or `0` [@resident-server] [@flush-policy]. `auto` is the default. It uses an idle-debounced interval based on an exponential moving average of measured save duration, clamped between two and ten seconds [@resident-server] [@flush-policy].

Shutdown ordering is part of the correctness contract. The server uses separate cancellation tokens for the main command loop and the ping responder. During shutdown, the main loop stops accepting new commands first, the in-flight command is drained, the handler is disposed, and only then is the ping responder cancelled [@resident-server]. This preserves the invariant that a successful ping implies the resident still owns the handler and file [@resident-server].

`close` sends `__close__` over the ping pipe and waits for the resident's acknowledgement. If shutdown detects that the backing file disappeared before saving, the close response can report a non-zero exit; if the file vanishes after a successful dispose, it reports a warning without flipping the exit code [@resident-server] [@resident-client].

## Failure Behavior

The server captures stdout and stderr around each request, then maps warnings, batch verdicts, validation failures, and JSON envelopes into a `ResidentResponse` with an exit code [@resident-server]. JSON-mode responses avoid double-wrapping already wrapped envelopes and merge resident-level warnings when needed [@resident-server].

Malformed or unreadable pipe input gets an explicit error response instead of a silent close. Pipe reads also have a large message-size ceiling so runaway messages fail visibly instead of being truncated and misreported as undelivered [@resident-server] [@resident-client].

## Consequences

The resident runtime makes OfficeCLI fast for repeated operations without giving up a single-writer model. The cost is that callers must understand the flush boundary: OfficeCLI sees live changes immediately, while non-OfficeCLI readers may need `save`, `close`, or an idle flush first [@resident-server] [@flush-policy]. Normal CLI dispatch is the entry point for this behavior; see [Command Dispatch](../cli/command-dispatch).
