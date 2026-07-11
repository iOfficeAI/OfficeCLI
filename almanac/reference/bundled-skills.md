---
title: "Bundled Skills"
summary: "Lookup reference for the OfficeCLI skill catalog, aliases, trigger summaries, bundled files, and supported installation targets."
topics: [reference, skills, agent]
sources:
  - id: skill-installer
    type: file
    path: src/officecli/Core/SkillInstaller.cs
  - id: root-skill
    type: file
    path: SKILL.md
  - id: skills-dir
    type: file
    path: skills/
---

Bundled skills are the agent-facing workflow guides embedded with OfficeCLI. The umbrella `officecli` skill explains the general OfficeCLI workflow, while the named sub-skills are lazy-loaded or installed when a document type or specialized document task needs more detailed rules [@root-skill] [@skill-installer]. The conceptual role is covered in [OfficeCLI skills](../concepts/officecli-skills) and [MCP and skills](../architecture/agent/mcp-and-skills).

## Catalog names

`SkillInstaller` exposes these named skills through `load_skill`, `skills list`, and specific skill installation [@skill-installer].

| Public name | Bundled folder | Trigger summary |
|---|---|---|
| `pptx` | `officecli-pptx` | Slide decks and presentations [@skill-installer] |
| `word` | `officecli-docx` | Word documents, reports, letters, and memos [@skill-installer] |
| `excel` | `officecli-xlsx` | Spreadsheets, financial models, and dashboards [@skill-installer] |
| `morph-ppt` | `morph-ppt` | Cross-slide Morph animation and continuous motion [@skill-installer] |
| `morph-ppt-3d` | `morph-ppt-3d` | 3D Morph decks with GLB models and camera work [@skill-installer] |
| `pitch-deck` | `officecli-pitch-deck` | Fundraising and investor decks [@skill-installer] |
| `academic-paper` | `officecli-academic-paper` | Academic papers and research reports [@skill-installer] |
| `data-dashboard` | `officecli-data-dashboard` | Data dashboards [@skill-installer] |
| `financial-model` | `officecli-financial-model` | Financial models and projections [@skill-installer] |
| `word-form` | `officecli-word-form` | Fillable forms, content controls, and protected documents [@skill-installer] |

The `skills/` tree contains these folders plus their `SKILL.md` entry points and any bundled references or helper files, such as `morph-ppt/reference/decision-rules.md` and helper scripts [@skills-dir]. The umbrella `officecli` skill is installed from `skills/officecli/SKILL.md` in embedded-resource builds, but it is intentionally kept out of the public `SkillMap` so `load_skill` without a name lists only sub-skills [@skill-installer].

## Root skill

The root skill covers `.docx`, `.xlsx`, and `.pptx`; tells agents to prefer read, DOM edit, then raw XML layers; and directs agents to check specialized skills before document work [@root-skill]. It also says `officecli help` is the authority for property names, value formats, command syntax, and schema details [@root-skill].

## Loading behavior

`load_skill` with no name returns a generated catalog containing each public skill name and its full routing description [@skill-installer]. `load_skill <name>` returns that skill's `SKILL.md` content with its `## Setup` section stripped and a reference-file manifest appended [@skill-installer]. `load_skill <name> --path <relpath>` returns one bundled text reference file after rejecting traversal segments, missing paths, and binary assets [@skill-installer].

Binary skill assets are excluded from text-channel loading. The blocked extensions include `.pptx`, `.docx`, `.xlsx`, common image formats, `.glb`, `.pdf`, `.zip`, and `.ico`; users must install the skill to get those files on disk [@skill-installer].

## Installation targets

Skill installation writes to detected agent skill directories under the user's home directory [@skill-installer].

| Agent | Aliases | Skill directory |
|---|---|---|
| Claude Code | `claude`, `claude-code` | `~/.claude/skills` [@skill-installer] |
| GitHub Copilot | `copilot`, `github-copilot` | `~/.copilot/skills` [@skill-installer] |
| Codex CLI | `codex`, `openai-codex` | `~/.agents/skills` [@skill-installer] |
| Cursor | `cursor` | `~/.cursor/skills` [@skill-installer] |
| Pi | `pi`, `pi-agent` | `~/.pi/agent/skills` [@skill-installer] |
| Windsurf | `windsurf` | `~/.windsurf/skills` [@skill-installer] |
| MiniMax CLI | `minimax`, `minimax-cli` | `~/.minimax/skills` [@skill-installer] |
| OpenCode | `opencode` | `~/.opencode/skills` [@skill-installer] |
| Hermes Agent | `hermes`, `hermes-agent` | `~/.hermes/skills` [@skill-installer] |
| OpenClaw | `openclaw` | `~/.openclaw/skills` [@skill-installer] |
| NanoBot | `nanobot` | `~/.nanobot/workspace/skills` [@skill-installer] |
| ZeroClaw | `zeroclaw` | `~/.zeroclaw/workspace/skills` [@skill-installer] |

`officecli skills install` and `officecli skills install all` install the umbrella skill to all detected agents [@skill-installer]. `officecli skills install <name>` installs all embedded files for one named sub-skill to all detected agents, and `officecli skills install <skill> <agent>` or the reversed argument order installs one named skill to one supported agent target [@skill-installer].

## Installed file handling

Specific skill installation copies every embedded file for the selected folder, rewriting Markdown cross-skill references where needed and leaving scripts or other non-Markdown files unchanged [@skill-installer]. After a binary upgrade, `RefreshInstalled` updates only skills that already have `SKILL.md` present in detected agent directories; it does not add new agents or new sub-skills [@skill-installer].
