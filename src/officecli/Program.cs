// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0


// Ensure UTF-8 output on all platforms (Windows defaults to system codepage
// e.g. GBK). How that is reached depends on where the stream actually goes:
//
//  * Redirected into a pipe or a file — write through our own UTF-8 writers.
//    The bytes reach the consumer untouched and the console object is never
//    touched, so `officecli … | tool` and `officecli … > out.txt` mutate
//    nothing that outlives the process.
//  * A real console — the terminal decodes bytes with whatever code page it
//    is set to, so putting CJK on a CP437/CP936 console means switching that
//    code page. Console.OutputEncoding does exactly that, but through the
//    console object, which is shared with the parent shell and outlives this
//    process: left as-is, every officecli run pinned the user's terminal to
//    CP 65001 and the next legacy program they ran wrote mojibake into it.
//    Switch it, then put it back on the way out. A hard kill (taskkill /F)
//    still skips the restore — best effort is the most a CLI can do here.
//
// Stdin gets the same treatment one layer down; see CommandBuilder.StdIn.
{
    var utf8NoBom = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
    if (!Console.IsOutputRedirected || !Console.IsErrorRedirected)
    {
        var previous = Console.OutputEncoding;
        if (previous.CodePage != System.Text.Encoding.UTF8.CodePage)
        {
            Console.OutputEncoding = utf8NoBom;
            void Restore(object? sender, EventArgs e)
            {
                try { Console.OutputEncoding = previous; } catch { /* console already gone */ }
            }
            AppDomain.CurrentDomain.ProcessExit += Restore;
            Console.CancelKeyPress += Restore;
        }
    }
    // Order matters: the assignment above resets Console.Out/Error, so install
    // our writers after it. AutoFlush + TextWriter.Synchronized match what the
    // framework's own console writers give — the resident and watch servers
    // write from several threads at once.
    if (Console.IsOutputRedirected)
        Console.SetOut(System.IO.TextWriter.Synchronized(
            new System.IO.StreamWriter(Console.OpenStandardOutput(), utf8NoBom) { AutoFlush = true }));
    if (Console.IsErrorRedirected)
        Console.SetError(System.IO.TextWriter.Synchronized(
            new System.IO.StreamWriter(Console.OpenStandardError(), utf8NoBom) { AutoFlush = true }));
}

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

// On Unix, named-pipe sockets live under $TMPDIR; a long (sandbox-nested)
// TMPDIR pushes them past the kernel's socket-path cap and resident/watch
// startup breaks (issue #263). Must run before any pipe endpoint or
// Path.GetTempPath()-keyed .lock/.port file is touched, including the
// __resident-serve__ dispatch below.
OfficeCli.Core.PipeTempDirGuard.EnsurePipePathFits();

// Internal commands (spawned as separate processes, not user-facing)
if (args.Length == 1 && args[0] == "__update-check__")
{
    OfficeCli.Core.UpdateChecker.RunRefresh();
    return 0;
}

// Schema fingerprint: officecli --output-schema-crc → CRC32 of the embedded
// schemas/help tree. Downstream automation pins this to detect property-
// surface drift across binary upgrades (same crc → schemas identical,
// safe to upgrade blind). A flag, not a version: no ordering semantics.
if (args.Length == 1 && args[0] == "--output-schema-crc")
{
    Console.WriteLine(OfficeCli.Help.SchemaCrc.Compute());
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
    // Optional --path <relpath> fetches one bundled reference file the skill's
    // SKILL.md points to (manifest listed at the end of the no-path output).
    // Mirrors the MCP load_skill path= argument.
    string? skillRelPath = null;
    var positional = new List<string>();
    for (int ai = 1; ai < args.Length; ai++)
    {
        if (args[ai] == "--path" && ai + 1 < args.Length) { skillRelPath = args[++ai]; continue; }
        positional.Add(args[ai]);
    }
    // No skill name → print the catalog (name + when-to-use), mirroring MCP.
    if (positional.Count == 0 && string.IsNullOrEmpty(skillRelPath))
    {
        Console.Out.Write(OfficeCli.Core.SkillInstaller.BuildSkillCatalog());
        return 0;
    }
    if (positional.Count == 1)
    {
        try
        {
            Console.Out.Write(string.IsNullOrEmpty(skillRelPath)
                ? OfficeCli.Core.SkillInstaller.LoadSkillContent(positional[0])
                : OfficeCli.Core.SkillInstaller.LoadSkillFile(positional[0], skillRelPath));
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

// Response-file token replacement is OFF: a bare `@…` token must reach the
// handler verbatim (`set row[N] --prop @height=25` forces the ROW-PROPERTY
// side of a column-shadow collision, same escape as `query row[@height…]`);
// the default replacer would reject it as "response file not found".
var parseResult = rootCommand.Parse(args,
    new System.CommandLine.ParserConfiguration { ResponseFileTokenReplacer = null });
return parseResult.Invoke();
