---
title: "Resident Sessions"
summary: "Resident sessions are per-file in-memory OfficeCLI handlers served over named pipes, with serialized commands and explicit or policy-driven disk flushing."
topics: [concepts, resident, runtime]
sources:
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

A resident session is a long-lived OfficeCLI process that keeps one document handler in memory for one file. Clients find it through a per-file named pipe, send commands over the main pipe, and use a separate ping pipe to check liveness or request close without waiting behind document work [@resident-server]. The model exists so repeated edits can reuse the same parsed document and avoid reopening or resaving the package for every command.

The server derives its pipe name from a normalized full file path hash, so the resident is scoped to the file rather than to a terminal or project directory [@resident-server]. The client verifies that a ping response names the same full path before treating a resident as the owner of a file [@resident-client].

## Serialized Ownership

Resident sessions are single-owner sessions. The server accepts pipe connections concurrently, but command execution is guarded by `_commandLock`, so document operations run one at a time against the in-memory handler [@resident-server]. The client only retries connection establishment; after it writes a command, it does not resend on read failure because the resident may already have applied a non-idempotent mutation [@resident-client].

That at-most-once rule is part of the safety model. A visible delivery failure is better than applying `add`, `remove`, `move`, `swap`, or `batch` twice.

## Dirty State And Flush Policy

Resident mutation commands first promote the handler to editable mode, mark the in-memory document dirty, and defer Word saves inside the resident lifetime [@resident-server]. The dirty flag is cleared by explicit save, close, shutdown, or idle autosave after a successful handler save [@resident-server].

The flush policy is controlled by one parser that recognizes `each`, `auto`, fixed seconds, and `off` or `0` [@flush-policy]. In `auto` mode, the debounce interval is derived from an exponential moving average of measured save duration and clamped between two and ten seconds [@flush-policy]. This lets ordinary resident sessions feel current without forcing every mutation to pay a full package save.

## Ping Liveness

The ping pipe has a specific invariant: a successful ping should mean the handler still owns the file [@resident-server]. During shutdown, the server cancels the main command loop before disposing the handler, and it cancels the ping responder only after dispose finishes [@resident-server]. This ordering supports the [resident ping liveness invariant](../decisions/resident-ping-liveness-invariant) and keeps clients from silently falling back to direct file access while a live resident still holds unsaved state.

Related runtime behavior is covered by [resident process and pipes](../architecture/runtime/resident-process-and-pipes) and the operational guide for [resident save, close, and flush](../guides/resident-save-close-flush).
