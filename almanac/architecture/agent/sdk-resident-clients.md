---
title: "SDK Resident Clients"
summary: "The Node and Python SDKs are thin resident-pipe clients that forward batch-shaped OfficeCLI commands without inventing a second document vocabulary."
topics: [architecture, sdk, resident, agent]
sources:
  - id: node-sdk
    type: file
    path: sdk/node/index.js
  - id: node-types
    type: file
    path: sdk/node/index.d.ts
  - id: python-sdk
    type: file
    path: sdk/python/officecli.py
  - id: python-package
    type: file
    path: sdk/python/pyproject.toml
  - id: node-package
    type: file
    path: sdk/node/package.json
---

The Node and Python SDKs are resident clients, not independent Office document libraries. They create or open a document by starting or reusing an OfficeCLI resident, then send batch-item-shaped commands over the resident pipe. This keeps SDK usage aligned with the [batch item protocol](../../concepts/batch-item-protocol) and the [resident process and pipes](../runtime/resident-process-and-pipes) architecture [@node-sdk] [@python-sdk].

## Shared Model

Both SDKs state the same rule in code comments: there is no second vocabulary. A command passed to `send` is the same object a user would put into an OfficeCLI batch list, with `command` or `op` selecting the verb, `props` carrying property values, and the remaining keys forwarded as command arguments [@node-sdk] [@python-sdk].

The SDKs expose two surfaces. `create` and `open` are bootstrap operations that may spawn one CLI process because no resident may exist yet. `send` and `batch` are the hot path and use pipe round trips instead of per-command process spawns [@node-sdk] [@python-sdk].

The Node type definitions make that contract public. `BatchItem` has common fields such as `command`, `op`, `path`, `parent`, `type`, `selector`, and `props`, while also allowing additional keys so new OfficeCLI batch fields can pass through without SDK type changes [@node-types]. The Python package description says the distribution is a thin SDK for the resident pipe and has no runtime dependencies beyond the standard library [@python-package].

## Pipe Addressing And Framing

Both SDKs reproduce the resident pipe address convention. They hash the canonical full path with SHA-256, uppercase the first 16 hex characters, and build an `officecli-<hash>` pipe name, with a `-ping` companion pipe. macOS and Windows uppercase the full path before hashing; Linux leaves the path case-sensitive [@node-sdk] [@python-sdk].

Requests are one JSON line per connection. The request uses PascalCase fields such as `Command`, `Args`, `Props`, and `Json`, and the response contains `ExitCode`, `Stdout`, and `Stderr` [@node-sdk] [@python-sdk].

The SDKs parse response `Stdout` as JSON only when it is an object or array. Bare JSON scalars stay text, so a content reply like `42` is not confused with a numeric API value [@node-sdk] [@python-sdk].

## Command Delivery

Command delivery mirrors the resident client's busy policy. Connect is bounded by a generous timeout and a small retry loop with backoff. Once connected, the reply read blocks until the resident answers, which avoids cutting off slow but valid operations [@node-sdk] [@python-sdk].

Retries happen only before a command is delivered. If a resident accepts a connection and then closes without a complete reply, the SDKs raise instead of resending, because the command may already have been applied [@node-sdk] [@python-sdk].

On delivery failure, both SDKs probe the ping pipe to distinguish a dead resident from a live but unresponsive one. A dead resident can be restarted with `officecli open` and retried once. A live but busy resident is not bypassed, because touching the file directly would race the resident's eventual save [@node-sdk] [@python-sdk].

## Lifecycle

`open` is idempotent. It reuses a resident that is already serving the file or starts one if needed, then upgrades the resident idle timeout to the interactive window used by OfficeCLI open sessions [@node-sdk] [@python-sdk].

`create` runs `officecli create`, then binds a document handle to the resident that create auto-started. The returned handle is the same kind of handle returned by `open` [@node-sdk] [@python-sdk].

`close` sends the resident close command over the ping pipe. The SDKs treat a missing acknowledgement as acceptable only if a follow-up liveness check says the resident is gone; otherwise they surface the error so the caller does not falsely believe the file was released [@node-sdk] [@python-sdk].

## Packaging Boundary

The Node package depends on `@officecli/officecli`, so it can prefer a bundled installer-package binary before checking `PATH` or the official install location [@node-package] [@node-sdk]. The Python package is named `officecli-sdk` but installs a single `officecli` module; its metadata notes that pip cannot install the OfficeCLI binary and that the CLI must be installed separately [@python-package].

This packaging difference does not change the architecture. In both languages, the SDK is a thin transport shell over the resident process. It forwards OfficeCLI commands, returns parsed envelopes or raw text, and leaves document semantics to the CLI and its handlers [@node-sdk] [@python-sdk].
