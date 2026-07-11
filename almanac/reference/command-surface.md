---
title: "Command Surface"
summary: "Lookup reference for OfficeCLI's public commands, early-dispatch commands, hidden compatibility commands, and their owning source files."
topics: [reference, cli]
sources:
  - id: program
    type: file
    path: src/officecli/Program.cs
  - id: root-builder
    type: file
    path: src/officecli/CommandBuilder.cs
  - id: add-commands
    type: file
    path: src/officecli/CommandBuilder.Add.cs
  - id: set-command
    type: file
    path: src/officecli/CommandBuilder.Set.cs
  - id: get-query
    type: file
    path: src/officecli/CommandBuilder.GetQuery.cs
  - id: watch-command
    type: file
    path: src/officecli/CommandBuilder.Watch.cs
  - id: raw-commands
    type: file
    path: src/officecli/CommandBuilder.Raw.cs
  - id: view-command
    type: file
    path: src/officecli/CommandBuilder.View.cs
  - id: refresh-command
    type: file
    path: src/officecli/CommandBuilder.Refresh.cs
  - id: validate-command
    type: file
    path: src/officecli/CommandBuilder.Check.cs
  - id: save-command
    type: file
    path: src/officecli/CommandBuilder.Save.cs
  - id: batch-command
    type: file
    path: src/officecli/CommandBuilder.Batch.cs
  - id: dump-command
    type: file
    path: src/officecli/CommandBuilder.Dump.cs
  - id: import-create-merge
    type: file
    path: src/officecli/CommandBuilder.Import.cs
  - id: plugins-command
    type: file
    path: src/officecli/CommandBuilder.Plugins.cs
  - id: help-command
    type: file
    path: src/officecli/CommandBuilder.Help.cs
  - id: integration-stubs
    type: file
    path: src/officecli/CommandBuilder.IntegrationStubs.cs
---

The OfficeCLI command surface is split between process-level commands handled before the root parser and normal subcommands registered by `CommandBuilder.BuildRootCommand()`. The normal surface covers resident lifecycle, document viewing, node lookup, selectors, mutations, raw OOXML access, validation, batch replay, import/create/merge utilities, plugins, and schema-aware help; early dispatch keeps installer, MCP, skill, config, and schema CRC behavior outside the reusable document parser [@program] [@root-builder]. See [Command Dispatch](../architecture/cli/command-dispatch) for the runtime boundary and [Command Layers](../concepts/command-layers) for the semantic layers behind the verbs.

## Early Dispatch

These commands are recognized in `Program.cs` before the System.CommandLine root is invoked [@program].

| Token | Behavior | Notes |
| --- | --- | --- |
| `--output-schema-crc` | Prints the CRC32 fingerprint for embedded `schemas/help/**` resources. | This is a single flag, not a root subcommand [@program]. |
| `--help`, `-h`, `-?` | Rewrites to `help`, preserving trailing tokens. | Only the first or second argv token is rewritten so property values named `--help` are not corrupted [@program]. |
| `mcp` | Starts the MCP stdio server with no target, registers a target, unregisters a target, or lists target status. | `mcp-serve` is a legacy alias for starting the server [@program]. |
| `install` | Runs one-step binary, skill, and MCP setup. | It accepts optional target arguments through the installer path [@program]. |
| `skills`, `skill` | Lists skills, installs the base skill, installs a named skill, or installs to one agent target. | Singular and plural forms route to the same handler [@program]. |
| `load_skill` | Prints the skill catalog, a skill's `SKILL.md`, or a bundled reference file via `--path`. | This is read-only and mirrors the MCP `load_skill` behavior [@program]. |
| `config` | Handles update-check configuration. | It is dispatched before the root parser [@program]. |
| `__update-check__` | Runs the internal update refresh process. | This is an internal child-process route [@program]. |

`mcp`, `skills`, and `install` also appear in root help through stub commands. The stubs do not own execution; they print the same usage text used by early dispatch and `officecli help <cmd>` [@integration-stubs] [@help-command].

## Root Commands

Every command below is registered by `BuildRootCommand()` unless noted as a nested subcommand [@root-builder].

| Command | Purpose | Owner |
| --- | --- | --- |
| `open` | Starts or reuses a resident process for a document. | `CommandBuilder.cs` [@root-builder] |
| `close` | Flushes a resident, stops it, and treats no-resident as a successful no-op. | `CommandBuilder.cs` [@root-builder] |
| `watch` | Starts live preview; owns nested `mark`, `unmark`, `marks`, and `goto` commands. | `CommandBuilder.Watch.cs` [@watch-command] |
| `unwatch` | Stops a watch preview server for a document. | `CommandBuilder.Watch.cs` [@watch-command] |
| `view` | Renders document views such as text, annotated, outline, stats, issues, html, svg, screenshot, pdf, and forms. | `CommandBuilder.View.cs` [@view-command] |
| `get` | Reads a document node by path, defaulting the path to `/`, with optional depth and binary extraction. | `CommandBuilder.GetQuery.cs` [@get-query] |
| `query` | Runs CSS-like selectors with optional text filtering, compact output, and extra fields. | `CommandBuilder.GetQuery.cs` [@get-query] |
| `set` | Mutates a node or selector target with `--prop`, `--find`, and `--replace`. | `CommandBuilder.Set.cs` [@set-command] |
| `add` | Adds an element under a parent path, optionally from another element, at an index, after a path, or before a path. | `CommandBuilder.Add.cs` [@add-commands] |
| `remove` | Removes an element, with Excel cell shifting and Word tracked-delete modifier support. | `CommandBuilder.Add.cs` [@add-commands] |
| `move` | Moves an element to a target parent or relative position. | `CommandBuilder.Add.cs` [@add-commands] |
| `swap` | Swaps two document nodes. | `CommandBuilder.Add.cs` [@add-commands] |
| `refresh` | Recalculates derived Word field values; the description marks Word plus Windows as required. | `CommandBuilder.Refresh.cs` [@refresh-command] |
| `raw` | Reads a raw document part, with Excel row and column filters. | `CommandBuilder.Raw.cs` [@raw-commands] |
| `raw-set` | Mutates a raw XML part by XPath and action. | `CommandBuilder.Raw.cs` [@raw-commands] |
| `add-part` | Creates a document part and returns a relationship id and path. | `CommandBuilder.Raw.cs` [@raw-commands] |
| `validate` | Runs OpenXML validation. | `CommandBuilder.Check.cs` [@validate-command] |
| `save` | Flushes resident in-memory changes while keeping the resident running. | `CommandBuilder.Save.cs` [@save-command] |
| `batch` | Replays a JSON array of batch items in one pass. | `CommandBuilder.Batch.cs` [@batch-command] |
| `dump` | Serializes a subtree into replayable batch output. | `CommandBuilder.Dump.cs` [@dump-command] |
| `import` | Imports CSV or TSV data into an Excel sheet. | `CommandBuilder.Import.cs` [@import-create-merge] |
| `create` | Creates a blank `.docx`, `.xlsx`, or `.pptx` file. | `CommandBuilder.Import.cs` [@import-create-merge] |
| `merge` | Merges a template with JSON data by replacing `{{key}}` placeholders. | `CommandBuilder.Import.cs` [@import-create-merge] |
| `plugins` | Groups plugin inspection commands. | `CommandBuilder.Plugins.cs` [@plugins-command] |
| `help` | Shows schema-driven help, command help, flat schema dumps, raw schema JSON, and JSONL output for flat dumps. | `CommandBuilder.Help.cs` [@help-command] |

## Hidden Root Commands

Hidden commands are registered for internal process control or compatibility. `__resident-serve__` runs the resident server child process and is marked hidden; it takes a per-file singleton lock before opening the document [@root-builder]. Top-level `mark`, `unmark`, `get-marks`, and `goto` are hidden compatibility aliases for the corresponding `watch` subcommands [@root-builder].

## Shared Flags And Help

The root command defines one shared `--json` option and passes it into normal command builders, so document commands can use the same AI-friendly output switch [@root-builder]. The `help` command recognizes `add`, `set`, `get`, `query`, and `remove` as schema verbs; it can list formats, list elements, filter elements by verb, render one element schema, emit raw schema JSON, or produce flat corpus dumps through `help all`, `help all --json`, and `help all --jsonl` [@help-command].
