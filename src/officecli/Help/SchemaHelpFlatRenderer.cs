// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;

namespace OfficeCli.Help;

/// <summary>
/// Flat, grep-friendly dump of every (format, element, property) row across the
/// schema corpus. One self-contained line per record so external tools like
/// grep / awk / fzf can match against the full record without context loss.
/// Two row tags: ELEM (element summary) and PROP (property detail).
///
/// Default (lean): name/type/ops/aliases/enum-values only — ~80 chars/line.
/// Verbose: adds description and first example — ~200 chars/line.
///
/// Example:
///   docx paragraph     ELEM  ops:[asgqr]  paths:/body/p[@paraId=ID];/body/p[N]
///   docx paragraph     PROP  alignment    enum    ops:[asg]  values:left|center|...  aliases:halign,align
/// </summary>
internal static class SchemaHelpFlatRenderer
{
    private static readonly string[] Verbs = { "add", "set", "get", "query", "remove" };

    internal static string RenderAll(bool verbose = false)
    {
        var sb = new StringBuilder();
        if (verbose)
        {
            sb.AppendLine("# officecli help all verbose — grep-friendly schema dump (full descriptions)");
            sb.AppendLine("# Columns: <format> <element> <ELEM|PROP> <name> <type> ops:[asgqr] <details> <description> ex:<example>");
        }
        else
        {
            sb.AppendLine("# officecli help all — grep-friendly schema dump");
            sb.AppendLine("# Columns: <format> <element> <ELEM|PROP> <name> <type> ops:[asgqr] <details>");
            sb.AppendLine("# For full descriptions and examples: officecli help all verbose");
        }
        sb.AppendLine("# ops letters: a=add s=set g=get q=query r=remove (- = not supported)");
        // Use placeholder tokens (<PROP>, <ELEM>) instead of bare PROP/ELEM
        // so this Tips comment does not match `grep " PROP "` / `grep " ELEM "`
        // alongside the real data rows.
        sb.AppendLine("# Tips: grep '^docx paragraph'  |  grep ' <PROP> '  |  grep alignment  |  grep aliases:halign");
        sb.AppendLine();

        foreach (var format in SchemaHelpLoader.ListFormats())
        {
            foreach (var element in SchemaHelpLoader.ListElements(format))
            {
                JsonDocument doc;
                try { doc = SchemaHelpLoader.LoadSchema(format, element); }
                catch { continue; }

                using (doc)
                {
                    AppendElementRow(sb, format, element, doc);
                    AppendPropertyRows(sb, format, element, doc, verbose);
                }
            }
        }
        return sb.ToString();
    }

    private static void AppendElementRow(StringBuilder sb, string format, string element, JsonDocument doc)
    {
        var root = doc.RootElement;
        var ops = FormatOps(root);
        var paths = FormatPaths(root);

        // <format> <element-padded> ELEM ops:[...] paths:...
        sb.Append(format).Append(' ');
        sb.Append(PadRight(element, 16)).Append("  ELEM  ");
        sb.Append("ops:[").Append(ops).Append(']');
        if (!string.IsNullOrEmpty(paths))
            sb.Append("  paths:").Append(paths);
        sb.AppendLine();
    }

    private static void AppendPropertyRows(StringBuilder sb, string format, string element, JsonDocument doc, bool verbose = false)
    {
        if (!doc.RootElement.TryGetProperty("properties", out var props)
            || props.ValueKind != JsonValueKind.Object) return;

        foreach (var prop in props.EnumerateObject())
        {
            if (prop.Value.ValueKind != JsonValueKind.Object) continue;

            var name = prop.Name;
            var type = TryGetString(prop.Value, "type") ?? "any";
            var ops = FormatOps(prop.Value);

            sb.Append(format).Append(' ');
            sb.Append(PadRight(element, 16)).Append("  PROP  ");
            sb.Append(PadRight(name, 22)).Append(' ');
            sb.Append(PadRight(type, 8)).Append(' ');
            sb.Append("ops:[").Append(ops).Append(']');

            // type-specific detail
            if (string.Equals(type, "enum", StringComparison.OrdinalIgnoreCase)
                && prop.Value.TryGetProperty("values", out var values)
                && values.ValueKind == JsonValueKind.Array)
            {
                sb.Append("  values:");
                bool first = true;
                foreach (var v in values.EnumerateArray())
                {
                    if (!first) sb.Append('|');
                    sb.Append(v.GetString());
                    first = false;
                }
            }

            // aliases (a frequent search target — surface inline)
            if (prop.Value.TryGetProperty("aliases", out var aliases)
                && aliases.ValueKind == JsonValueKind.Array)
            {
                sb.Append("  aliases:");
                bool first = true;
                foreach (var a in aliases.EnumerateArray())
                {
                    if (!first) sb.Append(',');
                    sb.Append(a.GetString());
                    first = false;
                }
            }

            if (verbose)
            {
                // description (truncated, single-line) or readback as fallback
                var desc = TryGetString(prop.Value, "description")
                           ?? TryGetString(prop.Value, "readback");
                if (!string.IsNullOrEmpty(desc))
                {
                    sb.Append("  ");
                    sb.Append(SingleLine(desc!, 120));
                }

                // first example (also a high-value search target)
                if (prop.Value.TryGetProperty("examples", out var examples)
                    && examples.ValueKind == JsonValueKind.Array)
                {
                    var first = examples.EnumerateArray().FirstOrDefault();
                    if (first.ValueKind == JsonValueKind.String)
                    {
                        sb.Append("  ex:");
                        sb.Append(SingleLine(first.GetString()!, 80));
                    }
                }
            }

            sb.AppendLine();
        }
    }

    private static string FormatOps(JsonElement scope)
    {
        // Supports either top-level "operations" object (element) or per-property
        // boolean flags named after the verbs (property).
        var sb = new StringBuilder(5);
        JsonElement opsObj = default;
        bool hasOpsObj = scope.ValueKind == JsonValueKind.Object
                         && scope.TryGetProperty("operations", out opsObj)
                         && opsObj.ValueKind == JsonValueKind.Object;

        foreach (var v in Verbs)
        {
            bool supported = false;
            if (hasOpsObj && opsObj.TryGetProperty(v, out var bv) && bv.ValueKind == JsonValueKind.True)
                supported = true;
            else if (!hasOpsObj && scope.TryGetProperty(v, out var pv) && pv.ValueKind == JsonValueKind.True)
                supported = true;
            sb.Append(supported ? v[0] : '-');
        }
        return sb.ToString();
    }

    private static string FormatPaths(JsonElement root)
    {
        if (!root.TryGetProperty("paths", out var paths)
            || paths.ValueKind != JsonValueKind.Object) return "";
        var parts = new List<string>();
        foreach (var kind in new[] { "stable", "positional" })
        {
            if (paths.TryGetProperty(kind, out var arr) && arr.ValueKind == JsonValueKind.Array)
                foreach (var p in arr.EnumerateArray())
                    if (p.ValueKind == JsonValueKind.String) parts.Add(p.GetString()!);
        }
        return string.Join(";", parts);
    }

    private static string? TryGetString(JsonElement obj, string name) =>
        obj.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString() : null;

    private static string SingleLine(string s, int max)
    {
        var collapsed = s.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
        while (collapsed.Contains("  ")) collapsed = collapsed.Replace("  ", " ");
        collapsed = collapsed.Trim();
        return collapsed.Length <= max ? collapsed : collapsed.Substring(0, max - 1) + "…";
    }

    private static string PadRight(string s, int width) =>
        s.Length >= width ? s : s + new string(' ', width - s.Length);
}
