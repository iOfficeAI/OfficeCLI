---
title: "MCP Shared CLI Root"
summary: "The MCP server exposes one officecli tool that parses command strings or argv arrays and invokes the same root command used by the terminal CLI."
topics: [architecture, runtime, mcp, cli]
sources:
  - id: mcp-server
    type: file
    path: src/officecli/McpServer.cs
  - id: mcp-installer
    type: file
    path: src/officecli/McpInstaller.cs
  - id: program
    type: file
    path: src/officecli/Program.cs
  - id: command-builder
    type: file
    path: src/officecli/CommandBuilder.cs
---

The MCP runtime is a stdio JSON-RPC server that exposes a single tool named `officecli`. Instead of defining one MCP tool per document operation, it accepts an OfficeCLI command line as a string or pre-split argv array, strips an optional leading binary name, and invokes the same `CommandBuilder.BuildRootCommand()` tree used by the terminal CLI. This keeps MCP parsing, validation, envelopes, and document behavior aligned with normal command dispatch [@mcp-server] [@command-builder].

## Server Entry Point

`Program.cs` handles `officecli mcp` before normal root-command dispatch. With no target, it starts `McpServer.RunAsync()`; with supported target arguments, it registers or unregisters the server with an MCP client through `McpInstaller` [@program] [@mcp-installer]. The legacy `mcp-serve` alias also starts the MCP server [@program].

The server reads newline-delimited JSON-RPC requests from stdin and writes JSON-RPC responses to stdout. It implements `initialize`, `notifications/initialized`, `tools/list`, `tools/call`, and `ping`, and it rejects unsupported JSON-RPC batch requests with an invalid-request error [@mcp-server].

At startup, the MCP process sets `OFFICECLI_NO_AUTO_RESIDENT=1` unless the environment already defines it. That prevents a lone MCP command from auto-spawning a resident whose deferred flush might surprise the caller, while still allowing commands to route through an already-running resident for the file [@mcp-server]. It also defaults `OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT=1` because stdin is the JSON-RPC transport for MCP, not a user-supplied batch payload [@mcp-server].

## One Tool Surface

`tools/list` advertises exactly one tool, `officecli`. Its input schema has one required property, `command`, whose type may be a string or an array of strings [@mcp-server]. The tool description tells agents to use the same CLI verbs, paths, props, help, validation, screenshot, and save flow that terminal examples use [@mcp-server].

`tools/call` rejects any tool name other than `officecli`. For the accepted tool, `ExecuteCommandLine()` extracts argv, handles special command families that live outside the shared root, handles screenshot output, or invokes the shared CLI root [@mcp-server].

The string form is tokenized without invoking a shell. It honors single and double quotes, preserves most backslashes inside double quotes, and never creates a command-injection surface because tokens go directly to the in-process System.CommandLine parser [@mcp-server]. The array form preserves empty string arguments, which is necessary for values such as `--prop text=` [@mcp-server].

## Shared Root Invocation

The MCP server keeps a cached `RootCommand` from `CommandBuilder.BuildRootCommand()` and calls `RootCommand.Parse(argv).Invoke()` while capturing stdout, stderr, and the exit code [@mcp-server] [@command-builder]. Parse and validation failures are allowed to render through System.CommandLine so MCP users receive the same usage blocks that terminal users see [@mcp-server].

This root-command path means most document commands need no MCP-specific marshalling. New CLI flags and command validation become available to MCP through the shared grammar as soon as they are added to the root command [@mcp-server] [@command-builder].

Some commands are not in the root tree because `Program.cs` early-dispatches them. MCP mirrors the read-only skill path by serving `load_skill`, `skill`, and `skills` from `SkillInstaller` directly [@mcp-server] [@program]. MCP also special-cases screenshots: it runs the CLI screenshot command, injects a temporary output path when needed, reads the PNG, returns an MCP image content block, and removes the auto-created temp file [@mcp-server].

## Result Mapping

CLI stdout and stderr are surfaced together when both exist, so a caller does not lose success text or advisory warnings [@mcp-server]. A non-zero CLI exit normally sets MCP `isError=true`, but exit code `2` with stdout is treated as "applied with caveats" rather than a hard MCP error, because the underlying add or set operation may already have landed while reporting unsupported properties [@mcp-server].

The server returns business verdicts verbatim. For example, a batch or validate command may write a structured envelope with `success:false` to stdout; MCP preserves that text and uses `isError` as the transport-level signal rather than rewriting the payload [@mcp-server].

## Registration

`McpInstaller` registers the command as `officecli mcp` for supported clients. It prefers a stable binary path: the canonical self-install path, then an `officecli` executable found on `PATH`, then the current process path as a fallback [@mcp-installer].

Supported installation targets include LM Studio, Claude Code, Cursor, and VS Code Copilot. LM Studio gets a plugin directory with a manifest and bridge config; Cursor and VS Code use JSON config files; Claude Code prefers the `claude mcp` CLI and falls back to editing `~/.claude.json` if the CLI is unavailable [@mcp-installer].

## Consequences

The MCP architecture chooses one command-string tool over a large generated MCP surface. The benefit is low drift: the CLI root remains the authoritative grammar, help remains the discovery path, and MCP behavior follows [Command Dispatch](../cli/command-dispatch) instead of a parallel tool schema [@mcp-server] [@command-builder]. Resident routing still applies when another OfficeCLI process already owns a file, as described in [Resident Process And Pipes](resident-process-and-pipes) [@mcp-server].
