// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using OfficeCli;

namespace OfficeCli.Core;

/// <summary>
/// Parser for the batch text channel (`batch --from script.txt`): one
/// CLI-style command per line, converted to the same <see cref="BatchItem"/>
/// model the JSON channel deserializes into, so both sources feed one
/// execution pipeline with identical atomicity and error reporting.
///
/// Line grammar:
///   line := verb operand [--type T | --prop k=v | --from P | --to P |
///                          --index N | --after P | --before P | --selector S |
///                          --mode M | --depth N]...
///   verb := set | add | remove | move | swap
///
/// A line whose first non-blank character is '#' is a comment; a trailing '\'
/// continues onto the next physical line (line numbers in errors report the
/// first physical line of the logical line). Quoting: single quotes are
/// literal content (the documented --prop text='$15M' convention), double
/// quotes group only; an unterminated quote rejects the line.
///
/// Every parse error is collected and the whole script is rejected up front
/// (code batch_parse_error) — a half-parsed script never executes.
/// Property-name typos are deliberately NOT rejected here: the parse layer
/// has no per-element wordlist, and the executor's unsupported_property
/// error already lists the valid props for the exact element being targeted.
/// </summary>
internal static class BatchScriptParser
{
    internal static readonly string[] SupportedVerbs = ["set", "add", "remove", "move", "swap"];

    private static readonly string[] ValueFlags =
        ["--type", "--from", "--to", "--after", "--before", "--selector", "--mode"];

    public static List<BatchItem> ParseFile(string path) => Parse(File.ReadAllText(path));

    public static List<BatchItem> Parse(string content)
    {
        var errors = new List<string>();
        var items = new List<BatchItem>();

        var physical = content.Replace("\r\n", "\n").Split('\n');
        var lineNo = 1;
        while (lineNo <= physical.Length)
        {
            var first = lineNo;
            var text = physical[lineNo - 1];
            while (text.TrimEnd().EndsWith('\\') && lineNo < physical.Length)
            {
                text = text.TrimEnd()[..^1] + physical[lineNo];
                lineNo++;
            }
            lineNo++;

            var trimmed = text.Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#'))
                continue;
            try
            {
                items.Add(ParseLine(trimmed, first));
            }
            catch (CliException ex)
            {
                errors.Add($"line {first}: {ex.Message}");
            }
        }

        if (errors.Count > 0)
            throw new CliException(
                $"batch script rejected: {errors.Count} parse error(s), nothing executed.\n"
                + string.Join("\n", errors))
            {
                Code = "batch_parse_error",
                Suggestion = "Fix the listed lines and re-run. Per line: <verb> <path> [--type T] [--prop k=v ...] " +
                             $"with verb ∈ {string.Join("/", SupportedVerbs)}; '#' comments; '\\' continues a line."
            };
        return items;
    }

    private static BatchItem ParseLine(string line, int lineNo)
    {
        var tokens = Tokenize(line, lineNo);
        var verb = tokens[0].ToLowerInvariant();
        if (!SupportedVerbs.Contains(verb))
            throw new CliException($"unknown verb '{tokens[0]}' (supported: {string.Join(", ", SupportedVerbs)})")
            { Code = "invalid_argument" };

        var item = new BatchItem { Command = verb };
        var positionals = new List<string>();
        for (var i = 1; i < tokens.Count; i++)
        {
            var tok = tokens[i];
            if (!tok.StartsWith("--"))
            {
                positionals.Add(tok);
                continue;
            }

            var flag = tok.ToLowerInvariant();
            if (flag == "--prop")
            {
                if (i + 1 >= tokens.Count)
                    throw new CliException("--prop needs a key=value value") { Code = "invalid_argument" };
                var kv = tokens[++i];
                var eq = kv.IndexOf('=');
                if (eq <= 0)
                    throw new CliException($"--prop value '{kv}' is not key=value") { Code = "invalid_argument" };
                (item.Props ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))[kv[..eq]] =
                    kv[(eq + 1)..];
            }
            else if (flag is "--index" or "--depth")
            {
                if (i + 1 >= tokens.Count)
                    throw new CliException($"{flag} needs a number") { Code = "invalid_argument" };
                var raw = tokens[++i];
                if (!int.TryParse(raw, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var n))
                    throw new CliException($"{flag} value '{raw}' is not an integer") { Code = "invalid_argument" };
                if (flag == "--index") item.Index = n;
                else item.Depth = n;
            }
            else if (ValueFlags.Contains(flag))
            {
                if (i + 1 >= tokens.Count)
                    throw new CliException($"{flag} needs a value") { Code = "invalid_argument" };
                var value = tokens[++i];
                switch (flag)
                {
                    case "--type": item.Type = value; break;
                    case "--from": item.From = value; break;
                    case "--to": item.To = value; break;
                    case "--after": item.After = value; break;
                    case "--before": item.Before = value; break;
                    case "--selector": item.Selector = value; break;
                    case "--mode": item.Mode = value; break;
                }
            }
            else
            {
                throw new CliException(
                    $"unknown option '{tok}' (supported: --prop, --type, --from, --to, --index, --after, --before, --selector, --mode, --depth)")
                { Code = "invalid_argument" };
            }
        }

        switch (verb)
        {
            case "set":
                item.Path = RequirePositional(positionals, "set <path>");
                RequireNoExtra(positionals, line);
                if (item.Props == null || item.Props.Count == 0)
                    throw new CliException("set carries no --prop k=v — a set line without properties would do nothing")
                    { Code = "invalid_argument" };
                break;
            case "add":
                item.Parent = RequirePositional(positionals, "add <parent>");
                RequireNoExtra(positionals, line);
                if (item.Type == null && item.From == null)
                    throw new CliException("add needs --type <type> (or --from <path> to clone an element)")
                    { Code = "invalid_argument" };
                break;
            case "remove":
                item.Path = RequirePositional(positionals, "remove <path>");
                RequireNoExtra(positionals, line);
                break;
            case "move":
                item.Path = RequirePositional(positionals, "move <path>");
                RequireNoExtra(positionals, line);
                if (item.To == null && !item.Index.HasValue && item.After == null && item.Before == null)
                    throw new CliException("move needs a destination: --to / --index / --after / --before")
                    { Code = "invalid_argument" };
                break;
            case "swap":
                if (positionals.Count != 2)
                    throw new CliException("swap needs exactly two paths: swap <path1> <path2>")
                    { Code = "invalid_argument" };
                item.Path = positionals[0];
                item.Path2 = positionals[1];
                break;
        }
        return item;
    }

    private static string RequirePositional(List<string> positionals, string usage)
    {
        if (positionals.Count == 0)
            throw new CliException($"missing path — usage: {usage}") { Code = "invalid_argument" };
        return positionals[0];
    }

    private static void RequireNoExtra(List<string> positionals, string line)
    {
        if (positionals.Count > 1)
            throw new CliException(
                $"unexpected token '{positionals[1]}' — quote it into a --prop value if it belongs to one: {line}")
            { Code = "invalid_argument" };
    }

    private static List<string> Tokenize(string line, int lineNo)
    {
        var tokens = new List<string>();
        var current = new StringBuilder();
        var hasToken = false;
        bool inSingle = false, inDouble = false;
        foreach (var ch in line)
        {
            if (inSingle) { if (ch == '\'') inSingle = false; else current.Append(ch); }
            else if (inDouble) { if (ch == '"') inDouble = false; else current.Append(ch); }
            else if (ch == '\'') { inSingle = true; hasToken = true; }
            else if (ch == '"') { inDouble = true; hasToken = true; }
            else if (char.IsWhiteSpace(ch))
            {
                if (current.Length > 0 || hasToken) { tokens.Add(current.ToString()); current.Clear(); hasToken = false; }
            }
            else current.Append(ch);
        }
        if (inSingle || inDouble)
            throw new CliException("unterminated quote") { Code = "invalid_argument" };
        if (current.Length > 0 || hasToken)
            tokens.Add(current.ToString());
        return tokens;
    }
}
