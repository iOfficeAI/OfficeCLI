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

// MCP commands: officecli mcp [target]
if (args.Length >= 1 && args[0] == "mcp")
{
    if (args.Length == 1)
    {
        // officecli mcp → start MCP server
        await OfficeCli.Core.McpServer.RunAsync();
        return 0;
    }
    if (args.Length == 2 && args[1] == "list")
    {
        OfficeCli.Core.McpInstaller.Install("list");
        return 0;
    }
    if (args.Length == 3 && args[1] == "uninstall")
    {
        OfficeCli.Core.McpInstaller.Uninstall(args[2]);
        return 0;
    }
    if (args.Length == 2)
    {
        // officecli mcp <target> → register + show instructions
        OfficeCli.Core.McpInstaller.Install(args[1]);
        return 0;
    }
    Console.Error.WriteLine("Usage: officecli mcp              Start MCP server");
    Console.Error.WriteLine("       officecli mcp <target>     Register (lms, claude, cursor, vscode, minimax)");
    Console.Error.WriteLine("       officecli mcp uninstall <target>  Unregister");
    Console.Error.WriteLine("       officecli mcp list         Show registration status");
    return 1;
}

// Install command: officecli install [target]
if (args.Length >= 1 && args[0] == "install")
{
    return OfficeCli.Core.Installer.Run(args.Skip(1).ToArray());
}

// Legacy alias
if (args.Length == 1 && args[0] == "mcp-serve")
{
    await OfficeCli.Core.McpServer.RunAsync();
    return 0;
}

// Skills commands: officecli skills install [skill-name]
if (args.Length >= 1 && args[0] == "skills")
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

// Non-blocking update check: spawns background upgrade if stale
if (Environment.GetEnvironmentVariable("OFFICECLI_SKIP_UPDATE") != "1")
    OfficeCli.Core.UpdateChecker.CheckInBackground();

var rootCommand = OfficeCli.CommandBuilder.BuildRootCommand();

if (args.Length == 0)
{
    rootCommand.Parse("--help").Invoke();
    return 0;
}

// Handle help commands (docx/xlsx/pptx) before System.CommandLine parsing
// so that --help also shows our custom output instead of the default help
if (OfficeCli.HelpCommands.TryHandle(args))
    return 0;

// Rewrite format-prefixed commands: "xlsx add cell <file> <path> ..." → "add <file> <path> --type cell ..."
// This allows users to type "officecli xlsx add cell file.xlsx /Sheet1 --prop ..."
// instead of "officecli add file.xlsx /Sheet1 --type cell --prop ..."
if (args.Length >= 4 && args[0].ToLowerInvariant() is "docx" or "xlsx" or "pptx"
    && args[1].ToLowerInvariant() is "add" or "set" or "get" or "query" or "remove" or "view" or "raw")
{
    var verb = args[1];
    var elementType = args[2];
    var rest = args.Skip(3).ToList();
    // Only rewrite if the next arg looks like a file path (not a flag)
    if (rest.Count > 0 && !rest[0].StartsWith("--"))
    {
        var newArgs = new List<string> { verb };
        newArgs.AddRange(rest);
        if (verb.Equals("add", StringComparison.OrdinalIgnoreCase))
            newArgs.InsertRange(2, ["--type", elementType]);
        args = newArgs.ToArray();
    }
}

var parseResult = rootCommand.Parse(args);
return parseResult.Invoke();
