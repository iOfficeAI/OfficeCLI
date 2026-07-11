---
title: "Command Dispatch"
summary: "OfficeCLI dispatches a small set of process-level commands before handing normal document commands to a shared System.CommandLine root."
topics: [architecture, cli]
sources:
  - id: program
    type: file
    path: src/officecli/Program.cs
  - id: command-builder
    type: file
    path: src/officecli/CommandBuilder.cs
  - id: integration-stubs
    type: file
    path: src/officecli/CommandBuilder.IntegrationStubs.cs
  - id: help-command
    type: file
    path: src/officecli/CommandBuilder.Help.cs
---

Command dispatch in OfficeCLI is split between a process bootstrapper and a reusable command tree. `Program.cs` owns startup policy, help rewriting, update and install side effects, and commands that must run before System.CommandLine. `CommandBuilder.BuildRootCommand()` owns the normal CLI surface for document operations, shared `--json` behavior, resident commands, batch, help, and registered integration stubs. This shape matters because the terminal CLI and MCP bridge can share the same root command while still keeping setup commands and stdio server startup out of the document-command parser [@program] [@command-builder].

## Startup Boundary

The process first sets UTF-8 output and pins the default culture to invariant, while preserving the original OS locale for document creation. That makes numeric and machine-readable output stable across regional settings before any command handler runs [@program].

Before the root command is built, `Program.cs` handles process-level routes. It returns a schema CRC for `--output-schema-crc`, rewrites `--help`, `-h`, and `-?` into the schema-aware `help` command, starts or installs MCP targets under `mcp`, runs the installer under `install`, accepts the legacy `mcp-serve` alias, handles `skill` and `skills`, serves `load_skill`, and handles update-check configuration [@program]. These routes either start a different long-lived process, mutate user installation state, or need help behavior that would be risky or awkward if they were treated as ordinary document subcommands [@program] [@help-command].

After early dispatch, the program logs the command, runs auto-install and background update checks, builds the root command, and invokes it through System.CommandLine. If no arguments are provided, it parses `help`; otherwise it parses the original argument vector with response-file expansion disabled so literal `@...` values can reach document handlers unchanged [@program].

## Root Command

`BuildRootCommand()` constructs the reusable command surface. It creates a shared `--json` option, adds `open`, `close`, and the hidden internal `__resident-serve__` command, then registers document and utility command families from partial `CommandBuilder` files: watch, view, get, query, set, add, remove, move, swap, refresh, raw access, validate, save, batch, dump, import, create, merge, plugins, and help [@command-builder].

The root command is also where resident-process dispatch enters the normal command tree. `open` starts or reuses a resident and can upgrade an auto-started resident's idle timeout. `close` asks a resident to flush and shut down, but treats the absence of a resident as an idempotent success because non-resident mutations already write to disk [@command-builder]. The hidden `__resident-serve__` command is only for child processes; it takes a per-file singleton lock before opening the document and exits quietly if another resident already owns the file [@command-builder].

The command tree is deliberately reusable. [MCP Shared CLI Root](../runtime/mcp-shared-cli-root) invokes this same root command in-process, so parsing and validation behavior stays aligned between terminal users and MCP clients [@command-builder].

## Help As Dispatch

Help is normalized before normal dispatch. `officecli --help`, `officecli <cmd> --help`, and related aliases are rewritten to `officecli help ...`, with trailing tokens preserved so schema help can drill into formats, verbs, and elements [@program]. The `help` command then decides whether the request is for the root command, an early-dispatch command, a registered System.CommandLine command, a flat schema dump, or an element schema [@help-command].

Early-dispatch commands still appear in root help through stubs. `BuildIntegrationStubCommands()` registers visible `mcp`, `skills`, and `install` commands, but their actions only print the same usage text used by `officecli help <cmd>` and early-dispatch error paths [@integration-stubs] [@help-command]. This keeps the command list complete without letting setup commands accidentally run through the document-command parser [@integration-stubs].

## Error And Output Contract

Most root-command actions run through `SafeRun`. In text mode it writes a concise `Error:` message to stderr; in JSON mode it writes a structured error envelope to stdout, matching the rest of the AI-facing output contract [@command-builder]. The same helper also captures command output when CLI logging is enabled, so logging does not change command behavior [@command-builder].

Resident delivery has a stricter rule. `TryResident()` first probes the fast ping pipe. If no resident owns the file, it may auto-start a short-lived resident unless disabled by environment. If a resident is alive but the main command cannot be delivered after retries, it returns a distinct busy failure instead of falling back to direct file access, because direct access could race the resident's in-memory copy [@command-builder]. That same resident path is one of the major collaborators of [Resident Process And Pipes](../runtime/resident-process-and-pipes).

## Consequences

The dispatch boundary gives OfficeCLI three useful properties. Process-level commands can remain simple and side-effectful, normal document commands can share one System.CommandLine grammar, and help can act as a schema-aware reference instead of a generic usage banner [@program] [@command-builder] [@help-command]. Batch execution builds on the same root-command and handler conventions; see [Batch Execution](batch-execution) for the multi-command path.
