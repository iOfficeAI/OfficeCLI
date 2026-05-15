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
| `validate <file>` | Schema/package validation. Use `view <file> issues` and render/app-open proof for layout/content checks. |
| `batch <file>` | Execute JSON command arrays in one open/save cycle. |
| `import <file> <parent-path> <source-file>` | Import CSV/TSV data into Excel/HWPX paths where supported. |
| `create` / `new <file>` | Create blank `.docx`, `.xlsx`, `.pptx`, `.hwpx`, and capability-gated `.hwp` when rhwp sidecars are ready. |
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
officecli view file.hwp png --page 1 --out /tmp/hwp-png --json
officecli view file.hwp pdf --page 1 --out out.pdf --json
officecli view file.hwp markdown --json
officecli view file.hwp info --json
officecli view file.hwp diagnostics --json
officecli view file.hwp dump --json
officecli view file.hwp fields --json
officecli view file.hwp field --field-name CustomerName --json
officecli view file.hwp table-cell --section 0 --parent-para 3 --control 0 --cell 0 --cell-para 0 --json
officecli view file.hwp native --op get-style-list --json
officecli set file.hwp /field --prop name=CustomerName --prop value="Jun" --prop output=out.hwp --json
officecli set file.hwp /text --prop find="old" --prop value="new" --prop output=out.hwp --json
officecli add file.hwp /text --type paragraph --prop value="new paragraph" --prop output=out.hwp --json
officecli set file.hwp /convert-to-editable --prop output=editable.hwp --json
officecli set file.hwp /native-op --prop op=split-paragraph --prop paragraph=0 --prop offset=5 --prop output=out.hwp --json
officecli set file.hwp /text --prop find="old" --prop value="new" --in-place --backup --verify --json
officecli create blank.hwp --json
```

## Output conventions

- Add `--json` for agent-readable envelopes.
- Errors use typed codes and suggestions where available.
- HWP/HWPX unsupported or unverified operations must fail closed with typed reasons instead of falling back silently.
