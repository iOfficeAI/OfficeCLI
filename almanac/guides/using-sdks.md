---
title: "Using SDKs"
summary: "The Node and Python SDKs are thin resident clients that open or create one OfficeCLI session, send batch-shaped commands over the pipe, and leave command semantics to the CLI."
topics: [guides, sdk]
sources:
  - id: node-readme
    type: file
    path: sdk/node/README.md
  - id: node-sdk
    type: file
    path: sdk/node/index.js
  - id: python-readme
    type: file
    path: sdk/python/README.md
  - id: python-sdk
    type: file
    path: sdk/python/officecli.py
---

Use the SDKs when a program needs repeated OfficeCLI operations on the same file. Both SDKs start or reuse a resident process for `open` or `create`, then send the same object shape used by the [batch item protocol](../concepts/batch-item-protocol) over the resident pipe [@node-readme] [@python-readme]. They do not define a second command vocabulary; new CLI fields can pass through as command arguments or `props` without adding SDK-specific methods [@node-sdk] [@python-sdk].

## Open The Session

In Node, install and import `@officecli/sdk`, then call `create(path, args?, options?)` or `open(path, options?)` [@node-readme]. In Python, install `officecli-sdk`, import `officecli`, and use `officecli.create(...)` or `officecli.open(...)` [@python-readme]. `create` accepts extra CLI flags such as `--force` or `--type`, because the SDK forwards them to the CLI process [@node-sdk] [@python-sdk].

Treat `with officecli.create(...) as doc:` in Python or `try/finally { await doc.close(); }` in Node as ownership. Closing stops the resident and flushes the in-memory document to disk [@node-sdk] [@python-sdk]. If you only borrow a resident owned by another process, avoid the context manager and do not call `close()` [@node-readme] [@python-readme].

## Send Commands

Use `doc.send({ command, path, props })` for one operation and `doc.batch([...])` for many operations in one pipe round-trip [@node-readme] [@python-readme]. `command` or `op` selects the CLI command, `props` becomes the property map, and every other key is forwarded as a command argument [@node-sdk] [@python-sdk].

For JSON-returning commands, the SDKs parse object or array stdout into the returned value [@node-sdk] [@python-sdk]. For text-style commands such as `view`, `raw`, or `dump`, pass `asJson = false` in Node or `as_json=False` in Python to request plain text [@node-readme] [@python-readme].

## Handle Failure

Transport and process failures raise `OfficeCliError`, but business outcomes remain in the returned envelope's `success` field [@node-readme] [@python-readme]. If a resident is gone, both SDKs restart it with `officecli open` and retry the command once [@node-sdk] [@python-sdk]. If the ping pipe proves the resident is alive but the main pipe is busy or unresponsive, the SDK raises instead of bypassing the live resident and risking a save race [@node-sdk] [@python-sdk].

## Installation Notes

The Node SDK prefers the bundled `@officecli/officecli` binary when it exists, then checks `PATH`, then the official install location, and can provision the bundled binary or run the official installer when `autoInstall` is enabled [@node-sdk]. The Python SDK checks `PATH`, then the official install location, and can run `install.sh` or `install.ps1` on first use unless `auto_install=False` is passed [@python-sdk]. See [installing and platform detection](installing-and-platform-detection) when changing those locations.
