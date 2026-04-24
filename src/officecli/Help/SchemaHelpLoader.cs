// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;

namespace OfficeCli.Help;

/// <summary>
/// Locates and loads help schemas from the schemas/help tree. Resolves format
/// aliases (word/excel/ppt) and element aliases declared inside each schema.
/// </summary>
internal static class SchemaHelpLoader
{
    private static readonly string[] CanonicalFormats = { "docx", "xlsx", "pptx" };

    private static readonly Dictionary<string, string> FormatAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["docx"] = "docx",
            ["word"] = "docx",
            ["xlsx"] = "xlsx",
            ["excel"] = "xlsx",
            ["pptx"] = "pptx",
            ["ppt"] = "pptx",
            ["powerpoint"] = "pptx",
        };

    private static string? _cachedRoot;

    internal static string LocateSchemasRoot()
    {
        if (_cachedRoot != null) return _cachedRoot;

        // 1. AppContext.BaseDirectory direct: schemas ship as Content next to
        //    the built binary (bin/Debug/.../ or published single-file location).
        var baseDir = AppContext.BaseDirectory;
        var direct = Path.Combine(baseDir, "schemas", "help");
        if (Directory.Exists(direct))
        {
            _cachedRoot = direct;
            return direct;
        }

        // 2. Walk up from AppContext.BaseDirectory looking for schemas/help
        //    (same logic as SchemaContractTests). Handles dev-tree `dotnet run`
        //    where bin/ is several levels below the repo root.
        var dir = baseDir;
        for (int i = 0; i < 10 && dir is not null; i++)
        {
            var candidate = Path.Combine(dir, "schemas", "help");
            if (Directory.Exists(candidate))
            {
                _cachedRoot = candidate;
                return candidate;
            }
            dir = Path.GetDirectoryName(dir);
        }

        throw new DirectoryNotFoundException(
            "Could not locate schemas/help/ starting from " + baseDir);
    }

    internal static IReadOnlyList<string> ListFormats() => CanonicalFormats;

    /// <summary>
    /// True if <paramref name="input"/> is a known format alias (docx/xlsx/pptx
    /// or word/excel/ppt/powerpoint). Used by the help dispatcher to decide
    /// whether to treat the token as a schema format or fall through to
    /// top-level command forwarding.
    /// </summary>
    internal static bool IsKnownFormat(string input) =>
        !string.IsNullOrEmpty(input) && FormatAliases.ContainsKey(input);

    /// <summary>
    /// Normalize a user-supplied format token to canonical docx/xlsx/pptx.
    /// Throws InvalidOperationException with a suggestion if unknown.
    /// </summary>
    internal static string NormalizeFormat(string input)
    {
        if (FormatAliases.TryGetValue(input, out var canonical)) return canonical;

        // Suggest closest format alias
        var best = ClosestMatch(input, FormatAliases.Keys);
        var suggestion = best != null ? $" Did you mean: {best}?" : "";
        throw new InvalidOperationException(
            $"error: unknown format '{input}'.{suggestion}\n" +
            $"Use: officecli help");
    }

    internal static IReadOnlyList<string> ListElements(string format)
    {
        var canonical = NormalizeFormat(format);
        var dir = Path.Combine(LocateSchemasRoot(), canonical);
        if (!Directory.Exists(dir))
            throw new DirectoryNotFoundException($"Schema directory missing: {dir}");

        return Directory.GetFiles(dir, "*.json")
            .Select(Path.GetFileNameWithoutExtension)
            .Where(n => !string.IsNullOrEmpty(n))
            .Select(n => n!)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// Load a schema for (format, element). Element can be the filename stem
    /// or any alias declared in another schema's "aliases" entry (rare, mostly
    /// a property-level concept, but checked for completeness).
    /// </summary>
    internal static JsonDocument LoadSchema(string format, string element)
    {
        var canonical = NormalizeFormat(format);
        var dir = Path.Combine(LocateSchemasRoot(), canonical);
        var elements = ListElements(canonical);

        // 1. Exact filename match (case-insensitive).
        var match = elements.FirstOrDefault(
            e => string.Equals(e, element, StringComparison.OrdinalIgnoreCase));
        if (match != null)
        {
            var full = Path.Combine(dir, match + ".json");
            return JsonDocument.Parse(File.ReadAllText(full));
        }

        // 2. Unknown element — suggest closest match.
        var best = ClosestMatch(element, elements);
        var suggestion = best != null ? $"\nDid you mean: {best}?" : "";
        throw new InvalidOperationException(
            $"error: unknown element '{element}' for format '{canonical}'.{suggestion}\n" +
            $"Use: officecli help {canonical}");
    }

    /// <summary>
    /// Read the canonical parent of an element from its schema and resolve it
    /// to a filename in the same format directory. Returns null if the schema
    /// has no parent declaration or the parent is a root-ish container
    /// (body / slide / sheet / document / workbook / presentation) — those
    /// cases are treated as "top-level" for listing purposes.
    ///
    /// Schema 'parent' values use element-semantic names (e.g. "row" inside
    /// table-cell.json), while the listing works over filenames
    /// (e.g. "table-row"). This method bridges the two namespaces by scanning
    /// the format's schemas for any whose internal "element" field matches
    /// the declared parent — that schema's filename is the returned parent.
    /// </summary>
    internal static string? GetParentForTree(string format, string element)
    {
        // Root-ish parents are treated as "no parent" so top-level elements
        // (paragraph, table, section, sheet, slide, cell...) don't get buried
        // under container schemas.
        var rootLike = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "body", "document", "slide", "sheet", "workbook", "presentation", "styles", "numbering"
        };

        string? rawParent;
        try
        {
            using var doc = LoadSchema(format, element);
            if (!doc.RootElement.TryGetProperty("parent", out var p)) return null;

            rawParent = p.ValueKind switch
            {
                JsonValueKind.String => p.GetString(),
                JsonValueKind.Array => p.EnumerateArray()
                                        .Select(a => a.GetString())
                                        .FirstOrDefault(s => !string.IsNullOrEmpty(s)),
                _ => null,
            };
        }
        catch
        {
            return null;
        }

        if (string.IsNullOrEmpty(rawParent)) return null;

        // Parent can be "paragraph|body" — take the first element-typed segment
        // (i.e. the first segment that isn't a root-like container).
        var parts = rawParent!.Split('|', StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim())
            .Where(s => !string.IsNullOrEmpty(s) && !rootLike.Contains(s))
            .ToList();
        if (parts.Count == 0) return null;

        var parentName = parts[0];

        // Resolve element-name → filename. Look for a schema file whose stem
        // matches verbatim first (common case), then fall back to scanning
        // for any schema whose internal "element" field matches.
        var siblings = ListElements(NormalizeFormat(format));
        if (siblings.Contains(parentName, StringComparer.OrdinalIgnoreCase))
            return parentName;

        foreach (var sib in siblings)
        {
            try
            {
                using var sibDoc = LoadSchema(format, sib);
                if (sibDoc.RootElement.TryGetProperty("element", out var elEl)
                    && string.Equals(elEl.GetString(), parentName, StringComparison.OrdinalIgnoreCase))
                {
                    return sib;
                }
            }
            catch { /* skip bad schemas */ }
        }

        // Couldn't resolve — surface the raw name; caller will treat it as
        // top-level (since it's not in the filename set), which is safe.
        return parentName;
    }

    /// <summary>
    /// Check whether a schema's top-level operations[verb] is true. Used by
    /// `officecli help &lt;format&gt; &lt;verb&gt;` to filter the element list.
    /// </summary>
    internal static bool ElementSupportsVerb(string format, string element, string verb)
    {
        try
        {
            using var doc = LoadSchema(format, element);
            if (doc.RootElement.TryGetProperty("operations", out var ops)
                && ops.TryGetProperty(verb, out var v)
                && v.ValueKind == JsonValueKind.True)
            {
                return true;
            }
        }
        catch
        {
            // Swallow — a bad schema shouldn't kill the filter.
        }
        return false;
    }

    /// <summary>
    /// Suggest the closest candidate from <paramref name="candidates"/> to
    /// <paramref name="input"/> using substring + Levenshtein. Returns null
    /// if no candidate is close enough.
    /// </summary>
    private static string? ClosestMatch(string input, IEnumerable<string> candidates)
    {
        var lower = input.ToLowerInvariant();

        // Prefer substring hit (common for user typos like `paragrah`).
        var substringHit = candidates.FirstOrDefault(
            c => c.Contains(lower, StringComparison.OrdinalIgnoreCase)
                 || lower.Contains(c, StringComparison.OrdinalIgnoreCase));

        string? best = null;
        int bestDist = int.MaxValue;
        foreach (var c in candidates)
        {
            var dist = LevenshteinDistance(lower, c.ToLowerInvariant());
            // Accept distance up to max(2, len/3) — same rule CommandBuilder uses.
            var maxDist = Math.Max(2, lower.Length / 3);
            if (dist <= maxDist && dist < bestDist)
            {
                best = c;
                bestDist = dist;
            }
        }

        return best ?? substringHit;
    }

    private static int LevenshteinDistance(string s, string t)
    {
        if (s.Length == 0) return t.Length;
        if (t.Length == 0) return s.Length;

        var d = new int[s.Length + 1, t.Length + 1];
        for (int i = 0; i <= s.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= t.Length; j++) d[0, j] = j;

        for (int i = 1; i <= s.Length; i++)
        {
            for (int j = 1; j <= t.Length; j++)
            {
                int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }

        return d[s.Length, t.Length];
    }
}
