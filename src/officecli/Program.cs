// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0


// Ensure UTF-8 output on all platforms (Windows defaults to system codepage e.g. GBK)
Console.OutputEncoding = System.Text.Encoding.UTF8;

// Internal commands (spawned as separate processes, not user-facing)
if (args.Length == 1 && args[0] == "__update-check__")
{
    OfficeCli.Core.UpdateChecker.RunRefresh();
    return 0;
}

// `--help`/`-h` on early-dispatch commands (mcp/skills/install) should
// reach System.CommandLine so the registered Command stubs print proper
// usage. Without this, e.g. `install --help` would otherwise run
// InstallBinary() before erroring "Unknown target: --help".
static bool HasHelpFlag(string[] a)
{
    for (int i = 1; i < a.Length; i++)
        if (a[i] is "--help" or "-h" or "-?") return true;
    return false;
}

// MCP commands: officecli mcp [target]
if (args.Length >= 1 && args[0] == "mcp" && !HasHelpFlag(args))
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
        OfficeCli.McpInstaller.Uninstall(args[2]);
        return 0;
    }
    if (args.Length == 2)
    {
        // officecli mcp <target> → register + show instructions
        OfficeCli.McpInstaller.Install(args[1]);
        return 0;
    }
    Console.Error.WriteLine("Usage: officecli mcp              Start MCP server");
    Console.Error.WriteLine("       officecli mcp <target>     Register (lms, claude, cursor, vscode)");
    Console.Error.WriteLine("       officecli mcp uninstall <target>  Unregister");
    Console.Error.WriteLine("       officecli mcp list         Show registration status");
    return 1;
}

// Install command: officecli install [target]
if (args.Length >= 1 && args[0] == "install" && !HasHelpFlag(args))
{
    return OfficeCli.Core.Installer.Run(args.Skip(1).ToArray());
}

// Legacy alias
if (args.Length == 1 && args[0] == "mcp-serve")
{
    await OfficeCli.McpServer.RunAsync();
    return 0;
}

// Skills commands: officecli skills install [skill-name]
if (args.Length >= 1 && args[0] == "skills" && !HasHelpFlag(args))
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
    if (args.Length == 2)
    {
        // Legacy: officecli skills claude → base SKILL.md to specific agent
        OfficeCli.Core.SkillInstaller.Install(args[1]);
        return 0;
    }
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  officecli skills install                Install base SKILL.md to all detected agents");
    Console.Error.WriteLine("  officecli skills install <skill-name>   Install a specific skill to all detected agents");
    Console.Error.WriteLine("  officecli skills <agent>                Install base SKILL.md to a specific agent");
    Console.Error.WriteLine($"Skills: {string.Join(", ", new[] { "pptx", "word", "excel", "morph-ppt", "pitch-deck", "academic-paper", "data-dashboard", "financial-model" })}");
    Console.Error.WriteLine("Agents: claude, copilot, codex, cursor, windsurf, minimax, openclaw, nanobot, zeroclaw, all");
    return 1;
}

// Config command: officecli config <key> [value]
if (args.Length >= 2 && args[0] == "config")
{
    OfficeCli.Core.CliLogger.LogCommand(args);
    OfficeCli.Core.UpdateChecker.HandleConfigCommand(args.Skip(1).ToArray());
    return 0;
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
    rootCommand.Parse("--help").Invoke();
    return 0;
}

var parseResult = rootCommand.Parse(args);
return parseResult.Invoke();
