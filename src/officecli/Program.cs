// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0


// Ensure UTF-8 output on all platforms (Windows defaults to system codepage e.g. GBK)
Console.OutputEncoding = System.Text.Encoding.UTF8;

// Snapshot the OS user culture BEFORE we pin the thread to Invariant.
// Read by `create --locale` (when no explicit tag is given) to bake the
// user's actual locale into the new doc, mirroring what Word / Pages /
// other producers do on the host machine. Stashed once at startup so
// subsequent reads of CultureInfo.CurrentCulture (which we deliberately
// force to Invariant below) don't lose the original signal.
OfficeCli.Core.LocaleFontRegistry.OsLocaleSnapshot = System.Globalization.CultureInfo.CurrentCulture.Name;

// officecli is a CLI / AI tool that produces machine-format output (OOXML,
// JSON, CSS, canonical DocumentNode.Format values) and accepts the same on
// input. All three formats fix '.' as the decimal separator. Pin the
// process-wide culture to invariant so every interpolated double, ToString,
// and TryParse uses '.' regardless of the user's regional settings —
// otherwise nl-NL / de-DE / fr-FR users see '141,73pt' in CSS, '0,5' in
// `get --json` output, and silently-zero numCache reads in charts.
System.Globalization.CultureInfo.DefaultThreadCurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = System.Globalization.CultureInfo.InvariantCulture;

// Internal commands (spawned as separate processes, not user-facing)
if (args.Length == 1 && args[0] == "__update-check__")
{
    OfficeCli.Core.UpdateChecker.RunRefresh();
    return 0;
}

// Unify `--help` with `help` so AI agents see one help surface, not two.
//   officecli [--help|-h|-?]              → officecli help
//   officecli <cmd> [--help|-h|-?] [...]  → officecli help <cmd>
// The `help` command renders schema details for docx/xlsx/pptx, EarlyDispatchHelp
// for mcp/skills/install, and forwards to the SCL `<cmd> --help` for everything
// else — making `help` the single source of truth, with `--help` as a compatibility
// alias. Done before any other dispatch so it overrides early-dispatch + SCL.
//
// Restricted to args[0] and args[1] only — a blanket scan over all args would
// also rewrite cases where `--help` appears as an option *value* (e.g.
// `officecli set foo.docx /body --prop --help`), silently corrupting the
// command into a help dump.
if (args.Length > 0)
{
    if (args[0] is "--help" or "-h" or "-?")
    {
        // `officecli --help docx [add chart]` → `officecli help docx [add chart]`.
        // Preserve trailing tokens so flag-style invocations can drill into
        // schema details, not just the root banner.
        var tail = args.Skip(1).ToArray();
        args = tail.Length == 0
            ? new[] { "help" }
            : new[] { "help" }.Concat(tail).ToArray();
    }
    else if (args.Length >= 2 && args[1] is "--help" or "-h" or "-?")
    {
        // `officecli set --help chart` → `officecli help set chart`.
        // Mirror the args[0] branch above: preserve tokens after the help
        // flag so '<cmd> --help <element>' drills into the element schema
        // (verb-filtered) instead of just listing the verb's elements.
        var tail = args.Skip(2).ToArray();
        args = tail.Length == 0
            ? new[] { "help", args[0] }
            : new[] { "help", args[0] }.Concat(tail).ToArray();
    }
}

// MCP commands: officecli mcp [target]
if (args.Length >= 1 && args[0] == "mcp")
{
    if (args.Length == 1)
    {
        // officecli mcp → start MCP server
        await OfficeCli.McpServer.RunAsync();
        return 0;
    }
    if (args.Length == 2 && args[1] == "list")
    {
        OfficeCli.McpInstaller.Install("list");
        return 0;
    }
    if (args.Length == 3 && args[1] == "uninstall")
    {
        return OfficeCli.McpInstaller.Uninstall(args[2]) ? 0 : 1;
    }
    if (args.Length == 2)
    {
        // officecli mcp <target> → register + show instructions
        return OfficeCli.McpInstaller.Install(args[1]) ? 0 : 1;
    }
    OfficeCli.CommandBuilder.WriteEarlyDispatchUsage("mcp", Console.Error);
    return 1;
}

// Setup AionUI command: officecli setup-aionui [--verbose] [--dry-run] [--force]
if (args.Length >= 1 && args[0] == "setup-aionui")
{
    return OfficeCli.Core.AionuiInstaller.Run(args.Skip(1).ToArray());
}

// Install command: officecli install [target]
if (args.Length >= 1 && args[0] == "install")
{
    return OfficeCli.Core.Installer.Run(args.Skip(1).ToArray());
}

// Legacy alias
if (args.Length == 1 && args[0] == "mcp-serve")
{
    await OfficeCli.McpServer.RunAsync();
    return 0;
}

// Skill[s] commands. `skill` and `skills` are interchangeable to forgive
// the singular/plural typo; routing is by the second token, not the first.
if (args.Length >= 1 && (args[0] == "skills" || args[0] == "skill"))
{
    if (args.Length == 2 && args[1] == "list")
    {
        // officecli skills list → list all available skills
        OfficeCli.Core.SkillInstaller.ListSkills();
        return 0;
    }
    if (args.Length == 2 && args[1] == "install")
    {
        // officecli skills install → base SKILL.md to all detected agents
        OfficeCli.Core.SkillInstaller.Install("install");
        return 0;
    }
    if (args.Length == 3 && args[1] == "install")
    {
        // officecli skills install morph-ppt → specific skill to all detected agents
        var result = OfficeCli.Core.SkillInstaller.InstallSkill(args[2]);
        return result.Count > 0 ? 0 : 1;
    }
    if (args.Length == 4 && args[1] == "install")
    {
        // officecli skills install <skill> <agent>  OR  <agent> <skill>
        // Token order is auto-detected — skill names and agent aliases don't overlap.
        var result = OfficeCli.Core.SkillInstaller.InstallSkillToAgentTarget(args[2], args[3]);
        return result.Count > 0 ? 0 : 1;
    }
    if (args.Length == 2)
    {
        // 2-arg form: install base SKILL.md to a specific agent
        // (officecli skills <agent-alias>). The previous "if it's a known skill
        // name → ensure-install + print" branch was removed in favor of the
        // dedicated `officecli load_skill <name>` command, so CLI matches MCP:
        // load = pure read, install = explicit `skills install <name>`.
        var result = OfficeCli.Core.SkillInstaller.Install(args[1]);
        return result.Count > 0 ? 0 : 1;
    }
    OfficeCli.CommandBuilder.WriteEarlyDispatchUsage("skills", Console.Error);
    return 1;
}

// load_skill: read-only counterpart of `skills install <name>`. Prints the
// embedded SKILL.md content for a named skill to stdout with no install
// side-effect. Mirrors the MCP `load_skill` tool exactly so CLI and MCP have
// the same semantics.
if (args.Length >= 1 && args[0] == "load_skill")
{
    if (args.Length == 2)
    {
        try
        {
            Console.Out.Write(OfficeCli.Core.SkillInstaller.LoadSkillContent(args[1]));
            return 0;
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }
    OfficeCli.CommandBuilder.WriteEarlyDispatchUsage("load_skill", Console.Error);
    return 1;
}

// Config command: officecli config <key> [value]
if (args.Length >= 2 && args[0] == "config")
{
    OfficeCli.Core.CliLogger.LogCommand(args);
    return OfficeCli.Core.UpdateChecker.HandleConfigCommand(args.Skip(1).ToArray());
}

// Log command
OfficeCli.Core.CliLogger.LogCommand(args);

// Auto-install: if running outside ~/.local/bin/officecli, copy self there.
// Fresh install → full Run() (binary + skills + MCP). Upgrade → binary only.
OfficeCli.Core.Installer.MaybeAutoInstall(args);

// Non-blocking update check: spawns background upgrade if stale
if (Environment.GetEnvironmentVariable("OFFICECLI_SKIP_UPDATE") != "1")
    OfficeCli.Core.UpdateChecker.CheckInBackground();

var rootCommand = OfficeCli.CommandBuilder.BuildRootCommand();

if (args.Length == 0)
{
    rootCommand.Parse("help").Invoke();
    return 0;
}

var parseResult = rootCommand.Parse(args);
return parseResult.Invoke();
