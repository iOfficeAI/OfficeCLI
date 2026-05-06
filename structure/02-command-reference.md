# Command reference

This reference follows `Program.cs`, `CommandBuilder.cs`, and generated help schemas. Use `officecli help` and format-specific help as the runtime source of truth.

## Early-dispatch commands

| Command | Purpose | Evidence |
| --- | --- | --- |
| `officecli mcp` | Start MCP stdio server. | `Program.cs`, `CommandBuilder.Help.cs` |
| `officecli mcp <target>` | Register MCP with clients such as Claude, Cursor, VS Code/Copilot. | `Program.cs` |
| `officecli mcp uninstall <target>` / `mcp list` | Unregister or inspect MCP registrations. | `Program.cs` |
| `officecli install [target]` | Install binary, skills, and MCP integration. | `Program.cs`, `Core/Installer.cs` |
| `officecli skills ...` / `officecli skill ...` | Install embedded skills to agents; singular alias accepted. | `Program.cs`, `Core/SkillInstaller.cs` |
| `officecli load_skill <name>` | Print embedded skill content without installing. | `Program.cs`, `CommandBuilder.Help.cs` |
| `officecli config <key> [value]` | Read/write update/config settings. | `Program.cs`, `Core/UpdateChecker.cs` |

## Registered root commands

| Command | Primary use |
| --- | --- |
| `open <file>` / `close <file>` | Start/stop a resident process for faster repeated edits. |
| `watch <file>` / `unwatch <file>` | Live HTML preview server with selection/mark support. |
| `mark`, `unmark`, `get-marks`, `goto` | Advisory marks and browser selection/navigation over a running watch session. |
| `view <file> <mode>` | Read/render document views. Modes include `text`, `annotated`, `outline`, `stats`, `issues`, `html`, `svg`, `screenshot`, `fields`, `field`. |
| `get <file> <path>` | Read one node by OfficeCLI path. |
| `query <file> <selector>` | CSS-like element queries. |
| `set <file> <path> --prop k=v` | Modify node properties; HWP/HWPX special paths include `/field`, `/text`, and HWP `/table/cell`. |
| `add <file> <parent> --type <type>` | Add typed elements. |
| `remove`, `move`, `swap` | Reorganize elements. |
| `raw`, `raw-set`, `add-part` | Raw OpenXML/XML escape hatch. |
| `validate <file>` / `check <file>` | Schema validation and layout/content checks. |
| `batch <file>` | Execute JSON command arrays in one open/save cycle. |
| `import <file> <parent-path> <source-file>` | Import CSV/TSV data into Excel/HWPX paths where supported. |
| `create` / `new <file>` | Create blank `.docx`, `.xlsx`, `.pptx`, or `.hwpx`. |
| `merge <template> <output>` | Merge `{{key}}` JSON/template data. |
| `compare <fileA> <fileB>` | Compare HWPX documents. |
| `capabilities` | Machine-readable HWP/HWPX capability report. |
| `schema` | Export/validate schemas and schema docs. |
| `help [format] [verb] [element]` | Schema-driven command/property reference. |
| `hwp` / `rhwp` | Experimental HWP/rhwp help; `hwp doctor` checks bridge readiness. |

## Help paths

```bash
officecli help
officecli help docx
officecli help xlsx set chart
officecli help pptx add shape
officecli help hwpx
officecli hwp
officecli hwp doctor --json
officecli capabilities --json
```

Format aliases from `SKILL.md` and help code: `word -> docx`, `excel -> xlsx`, `ppt`/`powerpoint -> pptx`.

## HWP/HWPX special command examples

```bash
# Probe before HWP/HWPX work
officecli hwp doctor --json
officecli capabilities --json

# Binary HWP through experimental rhwp bridge
officecli view file.hwp text --json
officecli view file.hwp svg --page 1 --json
officecli view file.hwp fields --json
officecli view file.hwp field --field-name .DS_Store .agents .claude .migrated-v1 .shared_plan.md .skills_clone_meta.json AGENTS.md CLAUDE.md _INBOX _thread auth backup-memory-v1 browser-profile heartbeat.json jaw.db jaw.db-shm jaw.db-wal logs mcp.json memory prompts screenshots settings.json skills skills_ref tmp uploads week_05_2_production_shocks.md week_06_2_complete_markets.md week_06_2_complete_markets_stt_vir.md worklogs  --json
officecli set file.hwp /field --prop name=.DS_Store .agents .claude .migrated-v1 .shared_plan.md .skills_clone_meta.json AGENTS.md CLAUDE.md _INBOX _thread auth backup-memory-v1 browser-profile heartbeat.json jaw.db jaw.db-shm jaw.db-wal logs mcp.json memory prompts screenshots settings.json skills skills_ref tmp uploads week_05_2_production_shocks.md week_06_2_complete_markets.md week_06_2_complete_markets_stt_vir.md worklogs  --prop value= --prop output=out.hwp --json
codesign -dv --verbose=4 /Applications/Codex Computer Use.app 2>&1 | head -20 --prop value= --prop output=out.hwp --json
codesign -dv --verbose=4 /Applications/Codex Computer Use.app 2>&1 | head -20 --prop value= --in-place --backup --verify --json
codesign -dv --verbose=4 /Applications/Codex Computer Use.app 2>&1 | head -20 --prop output=out.hwp --json
```

## Output conventions

- Add `--json` for agent-readable envelopes.
- Errors use typed codes and suggestions where available.
- HWP/HWPX unsupported or unverified operations must fail closed with typed reasons instead of falling back silently.
