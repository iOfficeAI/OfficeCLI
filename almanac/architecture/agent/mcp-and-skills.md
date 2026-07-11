---
title: "MCP And Skills"
summary: "OfficeCLI's agent integration combines a single-command MCP server with lazy skill loading so agents use the same CLI grammar and load detailed workflow guidance only when needed."
topics: [architecture, mcp, skills, agent]
sources:
  - id: mcp-server
    type: file
    path: src/officecli/McpServer.cs
  - id: mcp-installer
    type: file
    path: src/officecli/McpInstaller.cs
  - id: skill-installer
    type: file
    path: src/officecli/Core/SkillInstaller.cs
  - id: root-skill
    type: file
    path: SKILL.md
  - id: program
    type: file
    path: src/officecli/Program.cs
---

OfficeCLI's agent layer exposes the CLI to AI tools through MCP and skills. The MCP server advertises one `officecli` tool whose input is a command string or argv array, and skills provide lazy, task-specific guidance for agents before they create or edit Office files. This keeps MCP, terminal usage, and skill examples aligned around the same command grammar [@mcp-server] [@root-skill].

## MCP Server

`officecli mcp` starts a stdio JSON-RPC server. The server implements `initialize`, `tools/list`, `tools/call`, and `ping`, and advertises protocol version `2024-11-05` [@mcp-server]. `Program.cs` handles `mcp` as an early-dispatch command before building the normal CLI root [@program].

The server exposes exactly one tool named `officecli`. A `tools/call` with any other name is rejected, so a misrouted request cannot silently mutate files under a bogus tool name [@mcp-server].

The tool is a thin shell over the CLI. It accepts `command` as either a string or a pre-split argv array, removes an optional leading `officecli`, tokenizes strings without invoking a shell, and invokes the same `System.CommandLine` root as the terminal CLI [@mcp-server]. This is the same design described by [MCP shared CLI root](../runtime/mcp-shared-cli-root): argument validation and command behavior come from the shared root, not from a second MCP-specific command model [@mcp-server].

MCP sets two environment defaults for its process. It defaults `OFFICECLI_NO_AUTO_RESIDENT=1` so a lone MCP command does not auto-spawn a lingering resident, while still routing through an existing resident if one already owns the file. It also defaults `OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT=1` because MCP uses stdin as the JSON-RPC transport and should not warn about redirected stdin on every batch call [@mcp-server].

## Skill Loading

Skills are served through both the CLI and MCP. `Program.cs` treats `load_skill` as a read-only early-dispatch command that prints a catalog, a full skill, or one bundled reference file without installing anything [@program]. The MCP server mirrors that path by intercepting `load_skill`, `skill`, and `skills` before invoking the CLI root, then serving content through `SkillInstaller` [@mcp-server].

`SkillInstaller` owns the available skill map. It maps names such as `pptx`, `word`, `excel`, `morph-ppt`, `pitch-deck`, and `financial-model` to embedded skill folders, builds a catalog with routing descriptions, and appends a reference-file manifest to loaded skill content [@skill-installer].

The root skill is the always-on guide. It tells agents to use OfficeCLI for `.docx`, `.xlsx`, and `.pptx`, to prefer read, DOM edit, and raw XML layers in that order, and to run `load_skill` for specialized document tasks before proceeding [@root-skill].

The MCP tool description includes a compact skill-trigger summary from `SkillInstaller.BuildSkillTriggerSummary`. Full skill text stays lazy behind `load_skill`, which keeps the tool description small while still prompting agents to load the right guide before mutating documents [@mcp-server] [@skill-installer]. The conceptual role is covered by [OfficeCLI skills](../../concepts/officecli-skills) and the bundled skills reference.

## Installation Into Agents

`officecli mcp <target>` registers the MCP server in supported clients. The installer supports LM Studio, Claude Code, Cursor, and VS Code/Copilot targets, and writes each config to launch the stable OfficeCLI path with the `mcp` argument [@mcp-installer].

The MCP installer chooses a stable command path before falling back to the running binary. It prefers the canonical self-install path, then a `PATH` wrapper or symlink, and only then `Environment.ProcessPath`, because versioned binary paths can rot after upgrades [@mcp-installer].

Skill installation is separate. `officecli skills install` installs the base skill to detected agents, while `officecli skills install <name>` installs a named sub-skill to detected agents. Supported agent directories include Claude Code, GitHub Copilot, Codex CLI, Cursor, Pi, Windsurf, MiniMax CLI, OpenCode, Hermes, OpenClaw, NanoBot, and ZeroClaw [@skill-installer].

The result is a two-part agent contract. MCP provides a narrow execution bridge into the same CLI commands humans use; skills provide the workflow knowledge that tells an agent which commands, schemas, checks, and delivery gates matter for a particular document task [@mcp-server] [@root-skill].
