// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Help;

namespace OfficeCli;

static partial class CommandBuilder
{
    // Recognized verbs that route help through the operation-scoped filter.
    // Matches IDocumentHandler's public surface — keep in sync if new verbs
    // are added to the handler API.
    private static readonly string[] HelpVerbs =
        { "add", "set", "get", "query", "remove" };

    // Commands that are NOT registered as System.CommandLine subcommands but
    // are instead early-dispatched in Program.cs. They do not understand
    // `--help` (install would actually run InstallBinary!), so the help
    // dispatcher must print their usage itself rather than shell out.
    // Keep these usage blurbs in sync with the Console.Error.WriteLine
    // blocks in Program.cs (mcp: ~line 40, skills: ~line 87, install path:
    // documented via Installer.Run).
    private static readonly Dictionary<string, string[]> EarlyDispatchHelp =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["mcp"] = new[]
            {
                "Usage:",
                "  officecli mcp                    Start MCP stdio server (for AI agents)",
                "  officecli mcp <target>           Register officecli with an MCP client",
                "  officecli mcp uninstall <target> Unregister officecli from an MCP client",
                "  officecli mcp list               Show registration status across all clients",
                "",
                "Targets: lms (LM Studio), claude (Claude Code), cursor, vscode (Copilot)",
            },
            ["skills"] = new[]
            {
                "Usage:",
                "  officecli skills install                Install base SKILL.md to all detected agents",
                "  officecli skills install <skill-name>   Install a specific skill to all detected agents",
                "  officecli skills <agent>                Install base SKILL.md to a specific agent",
                "  officecli skills list                   List all available skills",
                "",
                "Skills: pptx, word, excel, morph-ppt, pitch-deck, academic-paper, data-dashboard, financial-model",
                "Agents: claude, copilot, codex, cursor, windsurf, minimax, openclaw, nanobot, zeroclaw, all",
            },
            ["install"] = new[]
            {
                "Usage:",
                "  officecli install           One-step setup: install binary + skills + MCP to all detected agents",
                "  officecli install <target>  Install to a specific agent (claude, copilot, cursor, vscode, ...)",
                "",
                "Equivalent to: installing the binary, then `officecli skills install` and `officecli mcp <target>`.",
                "Targets: claude, copilot, codex, cursor, windsurf, vscode, minimax, openclaw, nanobot, zeroclaw, all",
            },
        };

    /// <summary>
    /// `officecli help [format] [verb] [element] [--json]` — schema-driven help.
    ///
    /// Argument forms accepted:
    ///   help                         → list formats
    ///   help &lt;format&gt;                → list all elements
    ///   help &lt;format&gt; &lt;verb&gt;         → list elements supporting that verb
    ///   help &lt;format&gt; &lt;element&gt;      → full element detail
    ///   help &lt;format&gt; &lt;verb&gt; &lt;element&gt; → verb-filtered element detail
    ///
    /// The middle arg is interpreted as verb iff it matches HelpVerbs.
    /// Mirrors the actual CLI structure: `officecli &lt;verb&gt; &lt;file&gt; ...`, so
    /// `officecli help docx add chart` reads exactly like the command you
    /// are about to run.
    /// </summary>
    public static Command BuildHelpCommand(Option<bool> jsonOption, RootCommand? rootCommand = null)
    {
        var formatArg = new Argument<string?>("format")
        {
            Description = "Document format: docx/xlsx/pptx (aliases: word, excel, ppt, powerpoint). Omit to list formats.",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var secondArg = new Argument<string?>("verb-or-element")
        {
            Description = "Verb (add/set/get/query/remove) or element name. Omit to list all elements.",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var thirdArg = new Argument<string?>("element")
        {
            Description = "Element name when a verb was given (e.g. 'help docx add chart').",
            Arity = ArgumentArity.ZeroOrOne,
        };

        var command = new Command("help", "Show schema-driven capability reference for officecli.");
        command.Add(formatArg);
        command.Add(secondArg);
        command.Add(thirdArg);
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            var format = result.GetValue(formatArg);
            var second = result.GetValue(secondArg);
            var third = result.GetValue(thirdArg);

            // Disambiguate middle arg: is it a verb or an element?
            string? verb = null;
            string? element = null;
            if (second != null)
            {
                if (third != null)
                {
                    // 3 args: format, verb, element — second MUST be a verb.
                    verb = second;
                    element = third;
                }
                else if (HelpVerbs.Contains(second, StringComparer.OrdinalIgnoreCase))
                {
                    // 2 args where second is a verb: filter listing by verb.
                    verb = second;
                }
                else
                {
                    // 2 args where second is NOT a verb: treat as element.
                    element = second;
                }
            }

            return SafeRun(() => RunHelp(format, verb, element, json, rootCommand), json);
        });

        return command;
    }

    private static int RunHelp(string? format, string? verb, string? element, bool json, RootCommand? rootCommand)
    {
        // Case 1: no args — list formats, commands, and usage banner.
        if (string.IsNullOrEmpty(format))
        {
            Console.WriteLine("officecli help — capability reference (schema-driven)");
            Console.WriteLine();
            Console.WriteLine("Formats:");
            foreach (var f in SchemaHelpLoader.ListFormats())
                Console.WriteLine($"  {f}");
            Console.WriteLine();

            Console.WriteLine("Commands (use 'officecli help <command>' for details):");
            WriteCommandIndex(rootCommand);
            Console.WriteLine();

            Console.WriteLine("Usage:");
            Console.WriteLine("  officecli help <format>                         List all elements");
            Console.WriteLine("  officecli help <format> <verb>                  Elements supporting the verb");
            Console.WriteLine("  officecli help <format> <element>               Full element detail");
            Console.WriteLine("  officecli help <format> <verb> <element>        Verb-filtered element detail");
            Console.WriteLine("  officecli help <format> <element> --json        Raw schema JSON");
            Console.WriteLine("  officecli help <command>                        Forward to '<command> --help'");
            Console.WriteLine();
            Console.WriteLine("Verbs: add, set, get, query, remove");
            Console.WriteLine("Aliases: word→docx, excel→xlsx, ppt/powerpoint→pptx");
            Console.WriteLine();
            Console.WriteLine("Tip: most shells expand [brackets] — quote paths: officecli get doc.docx \"/body/p[1]\"");
            return 0;
        }

        // Case 1b: not a format — try command help.
        //   - Early-dispatch commands (mcp/skills/install) don't understand
        //     --help (install would actually run InstallBinary!), so print
        //     a hardcoded usage blurb.
        //   - Registered SCL subcommands get their --help forwarded.
        if (!SchemaHelpLoader.IsKnownFormat(format)
            && verb == null
            && element == null)
        {
            if (EarlyDispatchHelp.TryGetValue(format, out var lines))
            {
                foreach (var line in lines) Console.WriteLine(line);
                return 0;
            }

            if (rootCommand != null)
            {
                var match = rootCommand.Subcommands.FirstOrDefault(
                    c => string.Equals(c.Name, format, StringComparison.OrdinalIgnoreCase)
                         && !c.Hidden
                         && c.Name != "help");
                if (match != null)
                    return rootCommand.Parse(new[] { match.Name, "--help" }).Invoke();
            }
        }

        // Validate verb if supplied.
        if (verb != null && !HelpVerbs.Contains(verb, StringComparer.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine($"error: unknown verb '{verb}'. Valid: {string.Join(", ", HelpVerbs)}.");
            return 1;
        }

        var canonicalFormat = SchemaHelpLoader.NormalizeFormat(format);

        // Case 2: format (+ optional verb) only — list elements.
        if (string.IsNullOrEmpty(element))
        {
            var all = SchemaHelpLoader.ListElements(canonicalFormat);
            var filtered = verb == null
                ? all
                : all.Where(el => SchemaHelpLoader.ElementSupportsVerb(canonicalFormat, el, verb!)).ToList();

            if (filtered.Count == 0 && verb != null)
            {
                Console.WriteLine($"No elements in {canonicalFormat} support '{verb}'.");
                return 0;
            }

            var header = verb == null
                ? $"Elements for {canonicalFormat}:"
                : $"Elements for {canonicalFormat} supporting '{verb}':";
            Console.WriteLine(header);

            // Build parent → children map for tree rendering. Children whose
            // declared parent isn't itself in the filtered set float back up
            // to top-level so nothing disappears under a filter.
            var filteredSet = new HashSet<string>(filtered, StringComparer.Ordinal);
            var parentOf = filtered.ToDictionary(
                el => el,
                el => SchemaHelpLoader.GetParentForTree(canonicalFormat, el),
                StringComparer.Ordinal);

            var topLevel = new List<string>();
            var byParent = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            foreach (var el in filtered)
            {
                var pr = parentOf[el];
                if (pr != null && filteredSet.Contains(pr))
                {
                    if (!byParent.TryGetValue(pr, out var list))
                        byParent[pr] = list = new List<string>();
                    list.Add(el);
                }
                else
                {
                    topLevel.Add(el);
                }
            }

            void WriteNode(string el, int depth)
            {
                Console.WriteLine($"{new string(' ', 2 + depth * 2)}{el}");
                if (byParent.TryGetValue(el, out var kids))
                    foreach (var kid in kids)
                        WriteNode(kid, depth + 1);
            }
            foreach (var el in topLevel)
                WriteNode(el, 0);
            Console.WriteLine();

            var detailHint = verb == null
                ? $"Run 'officecli help {canonicalFormat} <element>' for detail."
                : $"Run 'officecli help {canonicalFormat} {verb} <element>' for verb-filtered detail.";
            Console.WriteLine(detailHint);
            return 0;
        }

        // Case 3: format + (optional verb) + element — render schema.
        using var doc = SchemaHelpLoader.LoadSchema(format, element);
        Console.WriteLine(json
            ? SchemaHelpRenderer.RenderJson(doc)
            : SchemaHelpRenderer.RenderHuman(doc, verb));
        return 0;
    }

    // Groups of top-level commands for the no-arg `officecli help` banner.
    // The same command may not appear twice; anything registered on the root
    // but not listed here falls into "Other". Early-dispatch commands
    // (mcp/skills/install) are hardcoded into the Integration group since
    // they never appear in RootCommand.Subcommands.
    private static readonly (string Label, string[] Names)[] CommandGroups =
    {
        ("Document",     new[] { "create", "view", "get", "query", "set", "add", "remove", "move", "swap" }),
        ("Advanced",     new[] { "raw", "raw-set", "add-part", "validate", "batch", "import", "merge" }),
        ("Resident",     new[] { "open", "close" }),
        ("Preview",      new[] { "watch", "unwatch", "mark", "unmark", "get-marks" }),
        ("Integration",  new[] { "mcp", "skills", "install" }),
    };

    private static void WriteCommandIndex(RootCommand? rootCommand)
    {
        // Collect live subcommand names (non-hidden, excluding 'help' itself).
        var registered = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (rootCommand != null)
        {
            foreach (var c in rootCommand.Subcommands)
            {
                if (c.Hidden || c.Name == "help") continue;
                registered.Add(c.Name);
            }
        }
        // Early-dispatch commands are always available even though they are
        // not SCL subcommands.
        foreach (var k in EarlyDispatchHelp.Keys) registered.Add(k);

        var printed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (label, names) in CommandGroups)
        {
            var present = names.Where(n => registered.Contains(n)).ToList();
            if (present.Count == 0) continue;
            Console.WriteLine($"  {label,-12} {string.Join(", ", present)}");
            foreach (var n in present) printed.Add(n);
        }

        var leftover = registered.Where(n => !printed.Contains(n)).OrderBy(n => n).ToList();
        if (leftover.Count > 0)
            Console.WriteLine($"  {"Other",-12} {string.Join(", ", leftover)}");
    }
}
