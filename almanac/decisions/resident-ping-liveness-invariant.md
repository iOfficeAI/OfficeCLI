---
title: "Resident Ping Liveness Invariant"
summary: "Resident shutdown keeps the ping pipe alive until the handler is disposed so a successful ping means the resident still owns the file."
topics: [decisions, resident, runtime]
sources:
  - id: resident-server
    type: file
    path: src/officecli/ResidentServer.cs
  - id: resident-client
    type: file
    path: src/officecli/ResidentClient.cs
---

OfficeCLI defines resident liveness as more than "a process exists": if the ping pipe responds, the resident must still hold the document handler and file. Shutdown therefore cancels the main command loop first, drains in-flight work, disposes the handler, and only then cancels the ping responder. Future resident changes must keep that ordering, because clients use ping availability to decide whether direct file access is safe [@resident-server].

## Context

The resident runtime uses two named pipes for one file: a main pipe for document commands and a `-ping` pipe for liveness, idle-timeout updates, and close [@resident-server] [@resident-client]. `ResidentClient.TryConnect()` connects to the ping pipe, sends `__ping__`, and verifies that the returned path matches the requested file [@resident-client].

That ping result affects ownership. If a client believes no resident is alive, it may open the file directly. If a resident is alive, clients should route through it or report busy rather than racing the resident's in-memory handler.

## Decision

The invariant is: ping responds if and only if the resident handler still owns the file [@resident-server]. `ResidentServer` enforces this with separate cancellation tokens. `_mainCts` gates the main command loop, while `_pingCts` gates the ping responder and idle watchdog [@resident-server].

Shutdown is ordered. The server cancels `_mainCts`, kicks the main pipe, waits for the command lock so the current command can drain, disposes the handler, then cancels `_pingCts` and kicks the ping pipe [@resident-server]. During the slow handler-dispose step, the ping pipe intentionally remains live, so a probing client still sees "resident owns this file" rather than falling back to direct access [@resident-server].

`__close__` also uses this ordering. The close handler calls shutdown before acknowledging the client, so a caller that receives the close acknowledgement can assume the handler has been released or that the response reports the shutdown problem [@resident-server].

## Status

This invariant is documented and implemented in `ResidentServer.cs`. `ResidentClient.cs` is written around it: ping is the liveness probe, close uses the ping pipe, and command retry is limited to the connect phase so a delivered mutation is not resent after a broken response [@resident-client].

## Consequences

The invariant makes direct-mode fallback safer. A live ping means "do not open the file directly"; a dead ping means the resident has released the handler or is no longer a valid owner [@resident-server].

The tradeoff is that ping remains available during part of shutdown. That is intentional. It may make a client wait or report busy while disposal finishes, but it avoids the worse outcome where a client sees no resident and opens a file that is still being flushed [@resident-server].

Future shutdown, idle, close, or pipe changes must not cancel the ping responder before handler disposal. Doing so would break the ownership signal used by [Resident Process And Pipes](../architecture/runtime/resident-process-and-pipes) [@resident-server].
