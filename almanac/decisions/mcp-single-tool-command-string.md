---
title: "MCP Single Tool Command String"
summary: "OfficeCLI exposes MCP as one command-string tool so agent calls share the terminal CLI grammar instead of a separate per-command MCP schema."
topics: [decisions, mcp, cli, runtime]
sources:
  - id: mcp-server
    type: file
    path: src/officecli/McpServer.cs
---

OfficeCLI chooses a single MCP tool named `officecli` whose required `command` field is either a CLI command string or a pre-split argv array. The choice keeps MCP behavior tied to the same command root, validation, help text, screenshots, skills, stdout, stderr, and exit-code mapping that terminal users get. Future MCP work must preserve this shared grammar instead of adding a parallel set of per-operation MCP tools [@mcp-server].

## Context

The MCP server is a stdio JSON-RPC process that implements `initialize`, `tools/list`, `tools/call`, and `ping` [@mcp-server]. It needs to let agents create, inspect, mutate, validate, and render Office documents without drifting away from the normal CLI surface described in [MCP Shared CLI Root](../architecture/runtime/mcp-shared-cli-root) and [MCP And Skills](../architecture/agent/mcp-and-skills).

The alternative would be a generated MCP tool for each OfficeCLI verb or document operation. That would duplicate argument schemas, parsing rules, and help behavior. It would also require MCP-specific updates whenever a CLI flag or document command changed.

## Decision

`tools/list` advertises exactly one tool, `officecli` [@mcp-server]. Its input schema has one required property, `command`, and that property accepts either a string or an array of strings [@mcp-server].

`tools/call` rejects any name other than `officecli`, so misrouted calls do not mutate files under an unexpected tool name [@mcp-server]. Accepted calls are converted to argv and run through the shared root command, with special handling only for command families that are outside the root dispatch path, such as `load_skill` and screenshot image return [@mcp-server].

The string form is tokenized inside the process and never passed to a shell. The array form preserves empty string arguments, which matters for intentionally empty values such as clearing text with a property value [@mcp-server]. A leading `officecli` binary name is optional and is stripped when present, so examples copied from terminal-oriented skill text still work [@mcp-server].

## Status

This decision is implemented in `McpServer.cs`. The server builds a cached root command, parses argv with that root, invokes it while capturing stdout and stderr, and maps the result into MCP content blocks plus `isError` [@mcp-server].

## Consequences

The main benefit is low drift. New CLI verbs, flags, validation messages, and usage output become available to MCP through the shared root rather than a second schema layer [@mcp-server].

The cost is that agents must learn CLI syntax through `help`, loaded skills, and examples rather than through many narrow MCP tools. The server compensates by putting usage guidance in the single tool description and by surfacing CLI stdout and stderr together when both are present [@mcp-server].

MCP-specific code is allowed only where the transport requires it: JSON-RPC framing, tool validation, argv extraction, skill early dispatch, screenshot image content, and result mapping. Document behavior should remain owned by the CLI command path [@mcp-server].
