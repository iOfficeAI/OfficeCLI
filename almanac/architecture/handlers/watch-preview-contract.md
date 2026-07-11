---
title: "Watch Preview Contract"
summary: "How OfficeCLI live preview uses pre-rendered HTML, named pipes, SSE updates, marks, and goto without making watch own document parsing."
topics: [architecture, handlers, watch, preview]
sources:
  - id: watch-command
    type: file
    path: src/officecli/CommandBuilder.Watch.cs
  - id: mark-command
    type: file
    path: src/officecli/CommandBuilder.Mark.cs
  - id: goto-command
    type: file
    path: src/officecli/CommandBuilder.Goto.cs
  - id: watch-server
    type: file
    path: src/officecli/Core/Watch/WatchServer.cs
  - id: watch-notifier
    type: file
    path: src/officecli/Core/Watch/WatchNotifier.cs
  - id: command-builder
    type: file
    path: src/officecli/CommandBuilder.cs
---

The watch preview contract is a live-browser layer over rendered HTML, not a document parser. `officecli watch` starts a loopback HTTP server for one file, seeds it with initial HTML from a resident session or direct handler open, and then serves that HTML with watch scripts injected [@watch-command] [@watch-server]. Later document commands render fresh HTML or slide fragments while the handler is still open and send those artifacts to the watch process over a named pipe [@command-builder] [@watch-notifier]. This keeps watch aligned with the [rendering and preview stack](rendering-and-preview-stack) while preserving a hard boundary: watch does not open Office files or reference handlers [@watch-server] [@watch-notifier].

## Process ownership

One watched file maps to one hashed pipe name and one marker file in the temp directory [@watch-server]. The marker stores process id and port so callers can check whether a watch is already running without probing the pipe [@watch-server]. `WatchServer.RunAsync` refuses duplicate watch processes for the same file, starts a loopback `TcpListener`, writes the marker, starts the pipe listener, and starts an idle watchdog [@watch-server]. Shutdown is centralized through `StopAsync`, which cancels loops, stops the TCP listener, closes SSE streams, kicks the named pipe listener, deletes the marker, and removes stale Unix pipe sockets [@watch-server].

The HTTP server is local-only by default. It rejects requests whose `Host` header is not a loopback host, with an environment-variable allowlist for reverse-proxy cases [@watch-server]. State-changing HTTP endpoints also check `Origin`, allowing absent or loopback origins and rejecting cross-origin requests [@watch-server].

## Refresh protocol

The pipe protocol carries `WatchMessage` objects with an action, optional slide number, optional full HTML, optional slide HTML, optional scroll target, version, base version, and optional Word patches [@watch-notifier]. `WatchNotifier.NotifyIfWatching` serializes the message, writes it to the named pipe, waits for an acknowledgement, and gives up silently if no watch process answers [@watch-notifier].

Mutation commands call `NotifyWatch` or `NotifyWatchRoot` after a change. Excel and Word send full HTML snapshots plus an optional scroll selector; PowerPoint sends a slide-level `replace` when a changed slide can be rendered in isolation, otherwise it sends a full deck snapshot [@command-builder]. Root changes such as adding or removing slides may send `add`, `remove`, or full messages with the new full HTML cache [@command-builder].

Inside the watch server, a full HTML message updates the cached snapshot, increments the version, reconciles marks, and then broadcasts either a full SSE update or a smaller patch [@watch-server]. Word can use block-level patches when hidden block markers make the diff safe; Excel can use row-level patches when table chrome and chart overlay signatures are unchanged [@watch-server]. PowerPoint slide updates patch cached HTML by `data-slide` [@watch-server]. If a patch is unsafe or too broad, the server falls back to a full refresh [@watch-server].

## Marks, selection, and goto

Marks are in-memory annotations owned by the watch process. The `watch mark`, `watch unmark`, and `watch marks` commands talk to the pipe through `WatchNotifier`, and the server stores marks with ids, paths, colors, notes, matched text, stale state, and a monotonic version [@mark-command] [@watch-server] [@watch-notifier]. Mark paths must start with `/`, colors are validated server-side, and mark resolution uses the current cached HTML snapshot rather than reopening the document [@watch-server].

Browser selection is reported to the server through `POST /api/selection`, and commands can query that selection through the pipe [@watch-server] [@watch-notifier]. `watch goto` resolves supported Word paths to either an anchor id or a `data-path` selector, asks the server to validate the selector against cached HTML, and then broadcasts a scroll-only SSE event without changing cached HTML or version [@goto-command] [@watch-server] [@watch-notifier]. The public command list is tracked by the command surface reference.

The invariant is simple: commands and handlers render documents; watch relays, patches, scrolls, and annotates the rendered result.
