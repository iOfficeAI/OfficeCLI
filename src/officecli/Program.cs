// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;

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
    Console.Error.WriteLine("       officecli mcp <target>     Register (lms, claude, cursor, vscode)");
    Console.Error.WriteLine("       officecli mcp uninstall <target>  Unregister");
    Console.Error.WriteLine("       officecli mcp list         Show registration status");
    return 1;
}

// Legacy alias
if (args.Length == 1 && args[0] == "mcp-serve")
{
    await OfficeCli.Core.McpServer.RunAsync();
    return 0;
}

// Skills commands: officecli skills <target>
if (args.Length >= 1 && args[0] == "skills")
{
    if (args.Length == 2)
    {
        OfficeCli.Core.SkillInstaller.Install(args[1]);
        return 0;
    }
    Console.Error.WriteLine("Usage: officecli skills <target>     Install skills");
    Console.Error.WriteLine("Targets: claude, copilot, codex, cursor, windsurf, minimax, openclaw, nanobot, zeroclaw, all");
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

var parseResult = rootCommand.Parse(args);
return parseResult.Invoke();
