---
title: "OfficeCLI Skills"
summary: "OfficeCLI skills are lazy, agent-facing guides that teach document workflows while the CLI help schemas remain the command authority."
topics: [concepts, skills, agent]
sources:
  - id: root-skill
    type: file
    path: SKILL.md
  - id: docx-skill
    type: file
    path: skills/officecli-docx/SKILL.md
  - id: skill-installer
    type: file
    path: src/officecli/Core/SkillInstaller.cs
  - id: program
    type: file
    path: src/officecli/Program.cs
---

OfficeCLI skills are the human-readable operating manuals that an AI agent loads when it needs to work on Office files. They explain workflow, safety gates, and format-specific habits, while `officecli help` remains the source for command syntax, element properties, and schema details [@root-skill]. This matters because the repo supports both broad agent entry points and specialized document tasks: the base skill teaches the overall layer model, and sub-skills such as `officecli-docx` teach deeper rules for one format [@docx-skill].

## Why Skills Exist

The CLI already has a schema-driven help system, but schemas are not enough for an agent building a real document. A schema can list valid properties. It cannot reliably teach when to prefer read commands, when to run visual QA, or why a Word footer should use a live `PAGE` field instead of literal text. The base skill therefore tells agents to use the command layers in order: read first, use DOM edits when possible, and fall back to raw XML only when needed [@root-skill].

Specialized skills carry the operational knowledge that would be too large or too situational for every session. The DOCX skill, for example, explains shell quoting, incremental execution, document hierarchy, output quality requirements, the open/save lifecycle, and QA expectations for Word files [@docx-skill]. That makes skills a guidance layer around the command surface rather than a replacement for it.

## Lazy Loading

OfficeCLI keeps skill detail lazy. The `SkillInstaller` maps short names such as `word`, `excel`, `pptx`, `morph-ppt`, `academic-paper`, `data-dashboard`, and `financial-model` to bundled skill folders [@skill-installer]. It also builds a compact trigger summary for MCP tool descriptions, so agents see when to load a skill without receiving every full manual up front [@skill-installer].

The same installer can return a catalog, return one skill's `SKILL.md`, list bundled reference files, or fetch a referenced text file from a skill bundle [@skill-installer]. Binary assets are deliberately excluded from text-channel loading and are only made available through skill installation on disk [@skill-installer]. This split lets the MCP and CLI surfaces expose guidance without corrupting non-text assets.

## Loading Versus Installing

OfficeCLI separates reading a skill from installing a skill. `load_skill` is a read-only command that prints the skill catalog, a named skill, or one bundled reference file; `skills install` writes skill files into detected agent skill directories [@program]. Program dispatch keeps `skill` and `skills` as accepted command tokens, but routes named skill reading through `load_skill` so loading and installation have matching CLI and MCP semantics [@program].

Installation is targeted at AI clients. The installer knows directory conventions for Claude Code, GitHub Copilot, Codex CLI, Cursor, Pi, Windsurf, MiniMax CLI, OpenCode, Hermes Agent, OpenClaw, NanoBot, and ZeroClaw [@skill-installer]. A base install writes the umbrella OfficeCLI skill, while a named install writes a specific sub-skill and any bundled files for that sub-skill [@skill-installer].

## Relationship To Schemas And MCP

Skills are closely related to [MCP and skills](../architecture/agent/mcp-and-skills) and [bundled skills](../reference/bundled-skills), but they solve a different problem from schemas. The base skill explicitly tells agents to run `officecli help` when property names, value formats, or command syntax are uncertain [@root-skill]. The DOCX skill repeats that help-first rule and states that help is authoritative when a skill and help disagree [@docx-skill].

That division is the key mental model: skills teach how to work, and help schemas define what the installed binary accepts. A contributor changing OfficeCLI behavior should update schemas for command truth and update skills when the agent workflow, quality gate, or format-specific discipline changes.
