// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
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
    /// <summary>
    /// Print the verbose usage block for an early-dispatch command
    /// (mcp/skills/install) to the given writer. Single source of truth shared
    /// between `officecli help &lt;cmd&gt;`, the integration stubs' SetAction, and
    /// Program.cs's invalid-args error path. Returns true if the command name
    /// was recognized.
    /// </summary>
    internal static bool WriteEarlyDispatchUsage(string name, TextWriter writer)
    {
        // `skill` is the singular alias of `skills` (Program.cs accepts both as
        // the early-dispatch token). Normalize here so `officecli skill --help`
        // and `officecli help skill` resolve to the same usage block.
        if (string.Equals(name, "skill", StringComparison.OrdinalIgnoreCase))
            name = "skills";
        if (!EarlyDispatchHelp.TryGetValue(name, out var lines)) return false;
        foreach (var line in lines) writer.WriteLine(line);
        return true;
    }

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
                "  officecli skills install <skill-name> <agent>  Install a specific skill to a single agent (either order works)",
                "  officecli skills <agent>                Install base SKILL.md to a specific agent",
                "  officecli skills list                   List all available skills",
                "",
                "Skills: pptx, word, excel, morph-ppt, pitch-deck, academic-paper, data-dashboard, financial-model",
                "Agents: claude, copilot, codex, cursor, windsurf, minimax, openclaw, nanobot, zeroclaw, hermes, all",
            },
            ["load_skill"] = new[]
            {
                "Usage:",
                "  officecli load_skill <name>   Print the named skill's SKILL.md to stdout (no install)",
                "",
                "Skills: pptx, word, excel, morph-ppt, morph-ppt-3d, pitch-deck, academic-paper, data-dashboard, financial-model",
                "To install a skill on disk, run: officecli skills install <name>",
            },
            ["install"] = new[]
            {
                "Usage:",
                "  officecli install           One-step setup: install binary + skills + MCP to all detected agents",
                "  officecli install <target>  Install to a specific agent (claude, copilot, cursor, vscode, ...)",
                "",
                "Equivalent to: installing the binary, then `officecli skills install` and `officecli mcp <target>`.",
                "Targets: claude, copilot, codex, cursor, windsurf, vscode, minimax, openclaw, nanobot, zeroclaw, hermes, all",
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
        // Scoped to `help` only — `help all`/`help <fmt> all` can emit either:
        //   --json   one envelope-wrapped JSON document (matches other CLI
        //            commands; one parse for the whole corpus)
        //   --jsonl  NDJSON (one self-contained JSON object per line, no
        //            envelope, streaming-friendly)
        // Mutually exclusive on `help all`. Other help forms ignore --jsonl
        // since they're either single documents (use --json) or human-readable
        // listings with no JSON form.
        var jsonlOption = new Option<bool>("--jsonl")
        {
            Description = "(help all only) Emit NDJSON: one JSON object per line, no envelope.",
        };

        var command = new Command("help", "Show schema-driven capability reference for officecli.");
        command.Add(formatArg);
        command.Add(secondArg);
        command.Add(thirdArg);
        command.Add(jsonOption);
        command.Add(jsonlOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            var jsonl = result.GetValue(jsonlOption);
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
                    // 3 args: format, verb, element — second is a verb only if it
                    // actually looks like one. If format is itself a HelpVerb (from
                    // the `<cmd> --help <format> <element>` rewrite) then second is
                    // a document format token, not a verb; leave verb=null so Case 1b
                    // handles it by showing SCL help for the command.
                    // CONSISTENCY(args-rewrite): mirrors the 2-arg guard below.
                    if (HelpVerbs.Contains(second, StringComparer.OrdinalIgnoreCase))
                    {
                        verb = second;
                        element = third;
                    }
                    else if (SchemaHelpLoader.IsKnownFormat(format!))
                    {
                        // format is a real schema format AND third is provided, but
                        // second isn't a verb — surface the error instead of
                        // silently falling through to Case 2 (which would list all
                        // elements, ignoring user input).
                        Console.Error.WriteLine(
                            $"error: unknown verb '{second}'. Valid: {string.Join(", ", HelpVerbs)}.");
                        return 1;
                    }
                    // else: format is a HelpVerb (CRUD-verb-as-format from the
                    // `<verb> --help <fmt> <element>` rewrite), second is the format
                    // token, third is the element — fall through with verb=null,
                    // element=null so Case 1b shows SCL command help.
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

            return SafeRun(() => RunHelp(format, verb, element, json, jsonl, rootCommand), json);
        });

        return command;
    }

    private static int RunHelp(string? format, string? verb, string? element, bool json, bool jsonl, RootCommand? rootCommand)
    {
        // --json and --jsonl are mutually exclusive on `help all` / `help <fmt>
        // all`: the first emits one envelope-wrapped JSON document, the second
        // emits NDJSON. Combining them has no coherent meaning. Reject early
        // with a clear message rather than silently picking one.
        if (json && jsonl)
        {
            Console.Error.WriteLine("error: --json and --jsonl are mutually exclusive.");
            return 1;
        }

        // Case 1: no args — print SCL's default help (Description, Usage,
        // Options, full Commands list with arg signatures + descriptions),
        // then append the schema-driven reference block. The SCL output is
        // the single source of truth for the command surface; this command
        // only adds what SCL doesn't know about (formats, schema verbs,
        // aliases, drill-in usage).
        // Use `== null` (not IsNullOrEmpty) so an explicit empty-string format
        // (`help '' docx paragraph`) falls through to NormalizeFormat → proper
        // "unknown format ''" error, instead of silently discarding the
        // trailing tokens by routing into the no-args banner.
        // CONSISTENCY(empty-arg) — mirrors the Case 2 element guard.
        // Case 0: `help all` — flat, grep-friendly dump of every (format,
        // element, property) row across the schema corpus. One self-contained
        // line per record so `officecli help all | grep <term>` returns
        // intelligible matches without context loss.
        if (string.Equals(format, "all", StringComparison.OrdinalIgnoreCase))
        {
            if (verb != null || element != null)
            {
                Console.Error.WriteLine(
                    "error: 'help all' takes no further arguments. Pipe to grep to filter.");
                return 1;
            }
            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    SchemaHelpFlatRenderer.RenderAllJsonArray()));
                return 0;
            }
            Console.Write(jsonl
                ? SchemaHelpFlatRenderer.RenderAllJsonl()
                : SchemaHelpFlatRenderer.RenderAll());
            return 0;
        }

        // Case 0b: `help <format> all` — same flat dump but filtered to one
        // format. "all" isn't a CRUD verb so it lands in `element` after the
        // upstream disambiguation. Saves the user a `| grep ^<format>`.
        if (format != null
            && SchemaHelpLoader.IsKnownFormat(format)
            && verb == null
            && string.Equals(element, "all", StringComparison.OrdinalIgnoreCase))
        {
            var canonical = SchemaHelpLoader.NormalizeFormat(format);
            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    SchemaHelpFlatRenderer.RenderAllJsonArray(canonical)));
                return 0;
            }
            Console.Write(jsonl
                ? SchemaHelpFlatRenderer.RenderAllJsonl(canonical)
                : SchemaHelpFlatRenderer.RenderAll(canonical));
            return 0;
        }

        if (format == null)
        {
            if (rootCommand != null)
            {
                // rootCommand.Parse(["--help"]) routes to SCL's HelpOption,
                // which writes Description/Usage/Options/Commands directly to
                // Console. Note Program.cs's `--help` → `help` rewrite only
                // runs once at process startup on the original args, so this
                // programmatic Parse goes straight to SCL and does not loop.
                rootCommand.Parse(new[] { "--help" }).Invoke();
                Console.WriteLine();
            }

            Console.WriteLine("Schema Reference (docx/xlsx/pptx):");
            Console.WriteLine("  officecli help <format>                         List all elements");
            Console.WriteLine("  officecli help <format> <verb>                  Elements supporting the verb");
            Console.WriteLine("  officecli help <format> <element>               Full element detail");
            Console.WriteLine("  officecli help <format> <verb> <element>        Verb-filtered element detail");
            Console.WriteLine("  officecli help <format> <element> --json        Raw schema JSON");
            Console.WriteLine("  officecli help all                              Flat dump of every (format,element,property) — pipe to grep");
            Console.WriteLine("  officecli help all --json                       Same dump as one envelope-wrapped JSON document");
            Console.WriteLine("  officecli help all --jsonl                      Same dump as NDJSON (one JSON object per line)");
            Console.WriteLine();
            Console.Write("  Formats: ");
            Console.WriteLine(string.Join(", ", SchemaHelpLoader.ListFormats()));
            Console.WriteLine("  Verbs:   add, set, get, query, remove");
            Console.WriteLine("  Aliases: word→docx, excel→xlsx, ppt/powerpoint→pptx");
            Console.WriteLine();
            Console.WriteLine("Tip: most shells expand [brackets] — quote paths: officecli get doc.docx \"/body/p[1]\"");
            return 0;
        }

        // Case 1b: not a format — try command help.
        //   - Early-dispatch commands (mcp/skills/install) don't understand
        //     --help (install would actually run InstallBinary!), so print
        //     a hardcoded usage blurb.
        //   - Registered SCL subcommands get their --help forwarded.
        //
        // CONSISTENCY(args-rewrite): `officecli set --help chart` is rewritten to
        // `officecli help set chart` by Program.cs. "set" is not a document format,
        // so we fall into this branch. The trailing element token ("chart") has no
        // meaning in SCL command-help context — ignore it and show SCL help for "set".
        // Guard drops `element == null` for CRUD verbs so the rewrite case is handled.
        if (!SchemaHelpLoader.IsKnownFormat(format)
            && verb == null
            && (element == null || HelpVerbs.Contains(format, StringComparer.OrdinalIgnoreCase)
                || EarlyDispatchHelp.ContainsKey(format)
                || string.Equals(format, "skill", StringComparison.OrdinalIgnoreCase)))
        {
            if (WriteHwpBridgeHelp(format, Console.Out, json))
                return 0;

            if (WriteEarlyDispatchUsage(format, Console.Out))
                return 0;

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
        // Use `== null` (not IsNullOrEmpty) so that an explicit empty-string
        // arg (`help docx ''`) falls through to Case 3 where LoadSchema raises
        // a proper "unknown element ''" error. CONSISTENCY(empty-arg).
        if (element == null)
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

}
