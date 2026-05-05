// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli.Help;

/// <summary>
/// Locates and loads help schemas from the schemas/help tree. Resolves format
/// aliases (word/excel/ppt) and element aliases declared inside each schema.
/// </summary>
internal static class SchemaHelpLoader
{
    private static readonly string[] CanonicalFormats = { "docx", "xlsx", "pptx", "hwpx", "hwp" };

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
            ["hwpx"] = "hwpx",
            ["hwp"] = "hwp",
            ["hangeul"] = "hwp",
            ["hanword"] = "hwp",
        };

    // Manifest index: canonical key "schemas/help/{format}/{element}.json"
    // (lowercased, forward slashes) → the actual resource name as MSBuild
    // emitted it. MSBuild may use either '/' or '\' in %(RecursiveDir) on
    // Windows; we normalize both forms at index-build time.
    private static Dictionary<string, string>? _manifestIndex;
    private static readonly object _manifestLock = new();

    private static Dictionary<string, string> ManifestIndex
    {
        get
        {
            if (_manifestIndex != null) return _manifestIndex;
            lock (_manifestLock)
            {
                if (_manifestIndex != null) return _manifestIndex;
                var asm = typeof(SchemaHelpLoader).Assembly;
                var idx = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var name in asm.GetManifestResourceNames())
                {
                    var canonical = name.Replace('\\', '/');
                    if (canonical.StartsWith("schemas/help/", StringComparison.OrdinalIgnoreCase))
                        idx[canonical] = name;
                }
                _manifestIndex = idx;
                return idx;
            }
        }
    }

    private static Stream? OpenSchemaStream(string format, string element)
    {
        var key = $"schemas/help/{format}/{element}.json";
        if (!ManifestIndex.TryGetValue(key, out var resourceName)) return null;
        return typeof(SchemaHelpLoader).Assembly.GetManifestResourceStream(resourceName);
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
        var prefix = $"schemas/help/{canonical}/";
        var elements = new List<string>();
        foreach (var key in ManifestIndex.Keys)
        {
            if (!key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
            var rest = key.Substring(prefix.Length);
            // Skip nested entries (none today, but future-proof).
            if (rest.Contains('/')) continue;
            if (!rest.EndsWith(".json", StringComparison.OrdinalIgnoreCase)) continue;
            elements.Add(rest.Substring(0, rest.Length - ".json".Length));
        }
        elements.Sort(StringComparer.Ordinal);
        return elements;
    }

    /// <summary>
    /// Load a schema for (format, element). Element can be the filename stem
    /// or any alias declared in another schema's "aliases" entry (rare, mostly
    /// a property-level concept, but checked for completeness).
    /// </summary>
    internal static JsonDocument LoadSchema(string format, string element)
    {
        var canonical = NormalizeFormat(format);
        var elements = ListElements(canonical);

        // 1. Exact filename match (case-insensitive).
        var match = elements.FirstOrDefault(
            e => string.Equals(e, element, StringComparison.OrdinalIgnoreCase));
        if (match != null)
        {
            using var stream = OpenSchemaStream(canonical, match)
                ?? throw new InvalidOperationException(
                    $"Embedded schema resource missing: schemas/help/{canonical}/{match}.json");
            // Read into memory so we can inspect for `extends` and merge with a
            // shared base if present. Most schemas have no extends and skip the
            // merge path entirely.
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            ms.Position = 0;
            var doc = JsonDocument.Parse(ms);
            var bases = ReadExtendsList(doc).ToList();
            if (bases.Count > 0)
            {
                doc.Dispose();
                ms.Position = 0;
                using var mainReader = new StreamReader(ms);
                var mainJson = mainReader.ReadToEnd();
                // Compose: start with first base, layer in each subsequent base,
                // then apply the override file last.
                string composed = LoadSharedBaseRaw(bases[0])
                    ?? throw new InvalidOperationException(
                        $"Schema {canonical}/{match}.json extends '{bases[0]}' but base not found at schemas/help/{bases[0]}.json");
                for (int i = 1; i < bases.Count; i++)
                {
                    var nextBase = LoadSharedBaseRaw(bases[i])
                        ?? throw new InvalidOperationException(
                            $"Schema {canonical}/{match}.json extends '{bases[i]}' but base not found at schemas/help/{bases[i]}.json");
                    composed = MergeSchemaJson(composed, nextBase);
                }
                var merged = MergeSchemaJson(composed, mainJson);
                return JsonDocument.Parse(merged);
            }
            return doc;
        }

        // 2. Unknown element — suggest closest match.
        var best = ClosestMatch(element, elements);
        var suggestion = best != null ? $"\nDid you mean: {best}?" : "";
        // CONSISTENCY(mcp-error): truncate user-supplied value in error messages to prevent
        // response amplification (caller echoes arbitrary-length input back unchanged).
        throw new InvalidOperationException(
            $"error: unknown element '{TruncateForError(element, 64)}' for format '{canonical}'.{suggestion}\n" +
            $"Use: officecli help {canonical}");
    }

    /// <summary>
    /// Read the `extends` field — either a single string or an array of
    /// strings — and yield the base refs in declaration order. Empty enumerable
    /// when no extends is declared.
    /// </summary>
    private static IEnumerable<string> ReadExtendsList(JsonDocument doc)
    {
        if (doc.RootElement.ValueKind != JsonValueKind.Object) yield break;
        if (!doc.RootElement.TryGetProperty("extends", out var extEl)) yield break;
        if (extEl.ValueKind == JsonValueKind.String)
        {
            var s = extEl.GetString();
            if (!string.IsNullOrEmpty(s)) yield return s!;
        }
        else if (extEl.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in extEl.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.String) continue;
                var s = item.GetString();
                if (!string.IsNullOrEmpty(s)) yield return s!;
            }
        }
    }

    /// <summary>
    /// Load the raw text of a shared base schema by reference like
    /// `_shared/chart`. Returns null when not found.
    /// </summary>
    private static string? LoadSharedBaseRaw(string baseRef)
    {
        var key = $"schemas/help/{baseRef}.json";
        if (!ManifestIndex.TryGetValue(key, out var resourceName)) return null;
        using var stream = typeof(SchemaHelpLoader).Assembly.GetManifestResourceStream(resourceName);
        if (stream == null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// Deep-merge a base schema JSON with an override schema JSON, producing
    /// the resolved bytes. Override semantics:
    ///   - Top-level scalar/array fields in override replace base.
    ///   - Top-level `properties` object: union of keys; same-name property
    ///     in override replaces the base entry entirely (no per-attribute deep
    ///     merge — properties are atomic).
    ///   - The synthetic `extends` and `shared_base` markers are stripped.
    /// </summary>
    private static string MergeSchemaJson(string baseJson, string overrideJson)
    {
        var baseNode = JsonNode.Parse(baseJson) as JsonObject
            ?? throw new InvalidOperationException("Shared base must be a JSON object.");
        var overrideNode = JsonNode.Parse(overrideJson) as JsonObject
            ?? throw new InvalidOperationException("Schema override must be a JSON object.");

        var merged = new JsonObject();

        // Start from base top-level (excluding shared_base marker).
        foreach (var kv in baseNode)
        {
            if (kv.Key == "shared_base") continue;
            merged[kv.Key] = kv.Value?.DeepClone();
        }

        // Apply override top-level (excluding extends marker).
        foreach (var kv in overrideNode)
        {
            if (kv.Key == "extends") continue;
            if (kv.Key == "properties")
            {
                // Properties order: override-declared first (preserve dev-authored
                // ordering of the format file), then base-only properties appended
                // in base order. Same-name in override replaces base entry atomically.
                var basedProps = merged["properties"] as JsonObject;
                var overProps = kv.Value as JsonObject;
                var combined = new JsonObject();
                if (overProps != null)
                {
                    foreach (var pkv in overProps)
                    {
                        combined[pkv.Key] = pkv.Value?.DeepClone();
                    }
                }
                if (basedProps != null)
                {
                    foreach (var pkv in basedProps)
                    {
                        if (combined.ContainsKey(pkv.Key)) continue;
                        // Re-clone to detach from basedProps before reassigning.
                        combined[pkv.Key] = pkv.Value?.DeepClone();
                    }
                }
                merged["properties"] = combined;
            }
            else
            {
                merged[kv.Key] = kv.Value?.DeepClone();
            }
        }

        return merged.ToJsonString();
    }

    /// <summary>
    /// Truncate a user-supplied string for safe display in error messages,
    /// avoiding split UTF-16 surrogate pairs (which serialize as U+FFFD).
    /// Used by error sites that echo caller input back verbatim.
    /// </summary>
    internal static string TruncateForError(string s, int maxChars)
    {
        if (s.Length <= maxChars) return s;
        var cut = maxChars;
        if (cut > 0 && char.IsHighSurrogate(s[cut - 1])) cut--;
        return s[..cut] + "…";
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
    /// Generic keys that are never declared as schema properties but are
    /// always legitimate on add/set — they describe how the element is
    /// created/located rather than the element's own OOXML properties.
    /// </summary>
    private static readonly HashSet<string> GenericVerbKeys =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "from", "copyFrom", "path", "positional", "text",
        };

    /// <summary>
    /// Dotted prefixes that indicate a sub-property namespace. If a property
    /// key starts with any of these (e.g. "font.", "alignment."), we accept
    /// it even if the schema doesn't enumerate every sub-key individually.
    /// This is the same leniency the existing handlers already apply at the
    /// property-key level.
    /// </summary>
    private static readonly string[] SubPropertyPrefixes =
    {
        "font.", "alignment.", "border.", "fill.", "shadow.", "glow.",
        "plotArea.", "chartArea.", "legend.", "title.", "datalabels.",
        // Chart sub-property namespaces — handled by ChartHelper.Setter /
        // SetterHelpers (series/trendline/errbar/point/dataLabel{N}/
        // dataTable/displayUnitsLabel/trendlineLabel/combo/area).
        // NOTE: axis./cataxis./valaxis./xaxis./yaxis. are deliberately NOT
        // listed here. The handler only supports a small fixed subset
        // (axis.font, axis.line, axis.visible, cataxis.{visible,line},
        // valaxis.{visible,line,labelrotation}, xaxis.labelrotation,
        // yaxis.labelrotation) — these are wired in as explicit aliases on
        // axisfont/axisline/axisvisible/cataxisline/valaxisline/
        // cataxisvisible/valaxisvisible/labelrotation in chart.json. A
        // blanket "axis." prefix would silently swallow typos like
        // axis.color and let Add succeed while the value is dropped.
        "series.", "trendline.", "errbars.", "errbar.",
        "datatable.", "displayunitslabel.", "trendlinelabel.",
        "combo.", "area.", "style.",
        // Word OOXML "element.attr" dotted keys for the generic typed-attr
        // fallback (TypedAttributeFallback.TrySet). Each entry corresponds
        // to a wordprocessing element whose attrs the fallback can write.
        // Schema validation is delegated to OpenXML SDK at write time, so
        // typos like `ind.notAttr` reach the handler and get rejected
        // there with a precise message — unlike unknown bare keys, which
        // are filtered upstream.
        "ind.", "shd.", "u.", "spacing.", "pbdr.",
        // Section-level: page size / margins / cols / type / etc.
        "pgsz.", "pgmar.", "cols.", "docgrid.", "lnnumtype.",
        // Table / row / cell containers: borders, margins, height, etc.
        "tblborders.", "tblcellmar.", "tcborders.", "tcmar.", "trheight.",
        "tcw.", "tblw.", "tbllayout.", "tblpr.", "tblpprex.",
    };

    /// <summary>
    /// Lenient prefixes that match indexed dotted keys (e.g. "series1.color",
    /// "dataLabel3.text", "point2.fill", "legendEntry1.delete"). Matched
    /// case-insensitively and only when followed by digits-then-dot.
    /// </summary>
    private static readonly string[] IndexedSubPropertyPrefixes =
    {
        "series", "datalabel", "point", "legendentry",
        // autofilter per-column criteria keys: criteria0.equals,
        // criteria3.gt, criteria12.contains, etc.
        "criteria",
        // table per-column override keys: columns.1.dxfId, etc.
        "columns.",
    };

    /// <summary>
    /// Validate a --prop dictionary against the schema for a given
    /// (format, element, verb). Returns the keys that are not recognized
    /// by the schema. Empty list means everything is declared.
    ///
    /// Lenient by design:
    ///   - Unknown format/element → return empty (don't break new elements
    ///     whose schema hasn't landed yet).
    ///   - Case-insensitive key comparison.
    ///   - Accepts a key if it matches a declared property name, any of that
    ///     property's "aliases", or a generic add/set key (from / copyFrom /
    ///     text / path / positional).
    ///   - Accepts dotted sub-property keys (font.*, alignment.*, border.*,
    ///     etc.) even when not enumerated — handlers already treat these as
    ///     a namespace.
    ///
    /// CONSISTENCY(schema-prop-validation): same validator is shared between
    /// CommandBuilder.Add (inline) and ResidentServer.ExecuteAdd so both
    /// execution paths report "bogus" props with matching semantics.
    /// </summary>
    internal static IReadOnlyList<string> ValidateProperties(
        string format,
        string element,
        string verb,
        IReadOnlyDictionary<string, string>? props)
    {
        if (props == null || props.Count == 0) return Array.Empty<string>();
        if (string.IsNullOrEmpty(format) || string.IsNullOrEmpty(element))
            return Array.Empty<string>();

        JsonDocument doc;
        try
        {
            // NormalizeFormat also throws on unknown formats; treat any
            // schema resolution failure as "don't know → be lenient".
            doc = LoadSchema(NormalizeFormat(format), element);
        }
        catch
        {
            return Array.Empty<string>();
        }

        using (doc)
        {
            // Build the allowed-key set once.
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var k in GenericVerbKeys) allowed.Add(k);

            if (doc.RootElement.TryGetProperty("properties", out var propsEl)
                && propsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in propsEl.EnumerateObject())
                {
                    // Only count the property as valid for this verb if the
                    // schema declares operations[verb]=true on it, OR if the
                    // schema is silent (defensive: some older entries omit
                    // the per-verb flags, treat those as allowed).
                    bool verbOk = true;
                    if (prop.Value.ValueKind == JsonValueKind.Object
                        && prop.Value.TryGetProperty(verb, out var verbFlag))
                    {
                        verbOk = verbFlag.ValueKind == JsonValueKind.True;
                    }

                    if (!verbOk) continue;

                    allowed.Add(prop.Name);

                    if (prop.Value.ValueKind == JsonValueKind.Object
                        && prop.Value.TryGetProperty("aliases", out var aliases)
                        && aliases.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var a in aliases.EnumerateArray())
                        {
                            var s = a.GetString();
                            if (!string.IsNullOrEmpty(s)) allowed.Add(s!);
                        }
                    }
                    // Some enum-typed schemas use object-form `aliases` for
                    // value-level synonyms and reserve a separate `propAliases`
                    // array for prop-name aliases (e.g. section.type accepts
                    // --prop break=… as a more intuitive name). bt-4.
                    if (prop.Value.ValueKind == JsonValueKind.Object
                        && prop.Value.TryGetProperty("propAliases", out var propAliases)
                        && propAliases.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var a in propAliases.EnumerateArray())
                        {
                            var s = a.GetString();
                            if (!string.IsNullOrEmpty(s)) allowed.Add(s!);
                        }
                    }
                }
            }
            else
            {
                // Schema has no "properties" block — don't second-guess.
                return Array.Empty<string>();
            }

            var unknown = new List<string>();
            foreach (var kv in props)
            {
                var key = kv.Key;
                if (string.IsNullOrEmpty(key)) continue;
                if (allowed.Contains(key)) continue;

                // Accept dotted sub-property namespaces.
                bool prefixOk = false;
                foreach (var pref in SubPropertyPrefixes)
                {
                    if (key.StartsWith(pref, StringComparison.OrdinalIgnoreCase))
                    {
                        prefixOk = true;
                        break;
                    }
                }
                if (prefixOk) continue;

                // Indexed dotted prefixes: "series1.color", "dataLabel3.text",
                // "point2.fill", "legendEntry1.delete". Match
                // <prefix><digits>. case-insensitively.
                bool indexedOk = false;
                var keyLower = key.ToLowerInvariant();
                foreach (var pref in IndexedSubPropertyPrefixes)
                {
                    if (!keyLower.StartsWith(pref)) continue;
                    int p = pref.Length;
                    int digitStart = p;
                    while (p < keyLower.Length && char.IsDigit(keyLower[p])) p++;
                    if (p > digitStart && p < keyLower.Length && keyLower[p] == '.')
                    {
                        indexedOk = true;
                        break;
                    }
                }
                if (indexedOk) continue;

                unknown.Add(key);
            }
            return unknown;
        }
    }

    /// <summary>
    /// Phase-1 schema/handler parity helper. Given a set of keys (e.g.
    /// the <c>DocumentNode.Format</c> keys returned by a handler's Get),
    /// return those that the schema doesn't declare as valid for
    /// <paramref name="verb"/>. Reuses <see cref="ValidateProperties"/> so
    /// alias / propAlias / dotted-sub-prefix / indexed-prefix leniency
    /// stays in one place.
    ///
    /// Lenient on unknown format/element (returns empty), matching the
    /// rest of the validator — tests on brand-new elements without a
    /// landed schema don't regress to hard failures.
    /// </summary>
    internal static IReadOnlyList<string> FindUnknownKeys(
        string format, string element, string verb, IEnumerable<string> keys)
    {
        if (keys == null) return Array.Empty<string>();
        var seen = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var k in keys)
        {
            if (string.IsNullOrEmpty(k)) continue;
            seen[k] = "";
        }
        if (seen.Count == 0) return Array.Empty<string>();
        return ValidateProperties(format, element, verb, seen);
    }

    /// <summary>
    /// Map a file extension (".docx"/".xlsx"/".pptx"/".hwpx"/".hwp") to the canonical
    /// schema format name, or null if the extension isn't an Office one.
    /// Small helper so CLI add/set sites don't duplicate the mapping.
    /// </summary>
    internal static string? FormatForExtension(string extension)
    {
        if (string.IsNullOrEmpty(extension)) return null;
        return extension.ToLowerInvariant() switch
        {
            ".docx" => "docx",
            ".xlsx" => "xlsx",
            ".pptx" => "pptx",
            ".hwpx" => "hwpx",
            ".hwp" => "hwp",
            _ => null,
        };
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
