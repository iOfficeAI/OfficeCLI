// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Section-based assembler for the cx chartStyle sidecar (an OOXML
/// chartEx auxiliary part defined by ECMA-376 / ISO/IEC 29500). Iterates
/// the canonical chartStyle section tags in schema-required order and
/// emits, for each section, either a curated fragment looked up by the
/// caller's (chartType, variant) key or a minimal schema-compliant
/// fallback provided by <see cref="MinimalScaffold"/>.
///
/// The result is a single byte stream suitable for feeding directly
/// into <c>ChartStylePart.FeedData</c>.
/// </summary>
internal static class ChartExStyleBuilder
{
    /// <summary>
    /// Canonical chartStyle section order. Must match the CT_ChartStyle
    /// schema sequence — Excel silently repairs (drops) the whole chart
    /// if a section is missing, reordered, or unknown.
    /// </summary>
    internal static readonly string[] Sections = new[]
    {
        "axisTitle",
        "categoryAxis",
        "chartArea",
        "dataLabel",
        "dataLabelCallout",
        "dataPoint",
        "dataPoint3D",
        "dataPointLine",
        "dataPointMarker",
        "dataPointMarkerLayout",
        "dataPointWireframe",
        "dataTable",
        "downBar",
        "dropLine",
        "errorBar",
        "floor",
        "gridlineMajor",
        "gridlineMinor",
        "hiLoLine",
        "leaderLine",
        "legend",
        "plotArea",
        "plotArea3D",
        "seriesAxis",
        "seriesLine",
        "title",
        "trendline",
        "trendlineLabel",
        "upBar",
        "valueAxis",
        "wall",
    };

    private const string CsNs = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
    private const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Build a cx chartStyle.xml stream for the given chart type and
    /// optional style variant. Caller feeds the stream into
    /// <c>ChartStylePart.FeedData</c>.
    /// </summary>
    /// <param name="chartType">
    /// The cx chart type name (case-insensitive, whitespace/dash/underscore
    /// tolerated via <see cref="NormalizeTypeForLookup"/>). Used as part
    /// of the section lookup key.
    /// </param>
    /// <param name="variant">
    /// Optional style variant name. Defaults to <c>"default"</c>. Also
    /// accepts <c>"style1"</c>..<c>"style10"</c> or bare integers
    /// <c>"1"</c>..<c>"10"</c>.
    /// </param>
    internal static Stream BuildChartStyleXml(
        string chartType, string variant = "default")
    {
        var normalizedType = NormalizeTypeForLookup(chartType);
        var normalizedVariant = NormalizeVariantForLookup(variant);

        var entry = GalleryIndex.TryGet(normalizedType, normalizedVariant);
        var styleId = entry?.StyleId ?? 410;

        var sb = new StringBuilder(4096);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append(
            $"<cs:chartStyle xmlns:cs=\"{CsNs}\" xmlns:a=\"{ANs}\" id=\"{styleId}\">");

        foreach (var section in Sections)
        {
            string? fragment = null;
            if (entry != null
                && entry.Fragments.TryGetValue(section, out var fragId))
            {
                fragment = FragmentStore.TryLoad(fragId);
            }
            // Any missing section falls through to the minimal
            // schema-compliant scaffold below.
            fragment ??= MinimalScaffold.For(section);
            sb.Append(fragment);
        }

        sb.Append("</cs:chartStyle>");
        return new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString()));
    }

    /// <summary>
    /// Normalize a chart type name to the lookup key used by the
    /// internal style index. Matches <c>ChartExBuilder.IsExtendedChartType</c>
    /// so "Box Whisker" / "box-whisker" / "BOXWHISKER" / "box_whisker"
    /// all resolve to the same entry.
    /// </summary>
    internal static string NormalizeTypeForLookup(string chartType)
    {
        return chartType.ToLowerInvariant()
            .Replace(" ", "")
            .Replace("_", "")
            .Replace("-", "");
    }

    /// <summary>
    /// Normalize a variant name to the lookup key used by the internal
    /// style index. Accepts <c>default</c>, <c>style{N}</c>, bare
    /// integers (<c>"3"</c> → <c>"style3"</c>), and any case.
    /// </summary>
    internal static string NormalizeVariantForLookup(string variant)
    {
        if (string.IsNullOrWhiteSpace(variant)) return "default";
        var v = variant.Trim().ToLowerInvariant();
        if (v == "default" || v == "0") return "default";
        if (int.TryParse(v, out var n) && n >= 1 && n <= 10) return $"style{n}";
        return v;
    }
}

/// <summary>
/// Minimal schema-compliant default fragments for cx chartStyle sections.
/// Every fragment is a self-contained <c>&lt;cs:section&gt;</c> element
/// with zero chart-type dependencies — safe to emit for any cx chart.
/// Each child of <c>cs:styleEntry</c> is <c>minOccurs=0</c> per
/// <c>CT_StyleEntry</c>, so the generic 4-ref form is the smallest
/// schema-valid content Excel accepts.
/// </summary>
internal static class MinimalScaffold
{
    /// <summary>
    /// Return the minimal default fragment for a given chartStyle section
    /// name. Specific sections need enriched content to keep the chart
    /// visually coherent; the rest get the generic 4-ref scaffold.
    /// </summary>
    internal static string For(string section) => section switch
    {
        // chartArea needs a visible background + outline for the chart
        // rectangle to render at all.
        "chartArea" =>
            "<cs:chartArea mods=\"allowNoFillOverride allowNoLineOverride\">" +
                "<cs:lnRef idx=\"0\"/>" +
                "<cs:fillRef idx=\"0\"/>" +
                "<cs:effectRef idx=\"0\"/>" +
                "<cs:fontRef idx=\"minor\">" +
                    "<a:schemeClr val=\"tx1\"/>" +
                "</cs:fontRef>" +
                "<cs:spPr>" +
                    "<a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill>" +
                    "<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" +
                        "<a:solidFill>" +
                            "<a:schemeClr val=\"tx1\">" +
                                "<a:lumMod val=\"15000\"/>" +
                                "<a:lumOff val=\"85000\"/>" +
                            "</a:schemeClr>" +
                        "</a:solidFill>" +
                        "<a:round/>" +
                    "</a:ln>" +
                "</cs:spPr>" +
            "</cs:chartArea>",

        // dataPoint uses the phClr placeholder fill so the accent color
        // from the accompanying chartColorStyle sidecar flows through.
        "dataPoint" =>
            "<cs:dataPoint>" +
                "<cs:lnRef idx=\"0\"/>" +
                "<cs:fillRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:fillRef>" +
                "<cs:effectRef idx=\"0\"/>" +
                "<cs:fontRef idx=\"minor\">" +
                    "<a:schemeClr val=\"tx1\"/>" +
                "</cs:fontRef>" +
                "<cs:spPr>" +
                    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                "</cs:spPr>" +
            "</cs:dataPoint>",

        // dataPointMarkerLayout is a self-closing element with
        // symbol/size attributes per CT_MarkerLayoutProperties — unlike
        // every other section it's not a CT_StyleEntry composite.
        "dataPointMarkerLayout" =>
            "<cs:dataPointMarkerLayout symbol=\"circle\" size=\"5\"/>",

        // plotArea / plotArea3D carry the `mods` attribute so Excel
        // honors user fill/line overrides emitted into chart.xml via
        // the plotareafill / plotarea.border knobs.
        "plotArea" =>
            "<cs:plotArea mods=\"allowNoFillOverride allowNoLineOverride\">" +
                "<cs:lnRef idx=\"0\"/>" +
                "<cs:fillRef idx=\"0\"/>" +
                "<cs:effectRef idx=\"0\"/>" +
                "<cs:fontRef idx=\"minor\"/>" +
            "</cs:plotArea>",

        "plotArea3D" =>
            "<cs:plotArea3D mods=\"allowNoFillOverride allowNoLineOverride\">" +
                "<cs:lnRef idx=\"0\"/>" +
                "<cs:fillRef idx=\"0\"/>" +
                "<cs:effectRef idx=\"0\"/>" +
                "<cs:fontRef idx=\"minor\"/>" +
            "</cs:plotArea3D>",

        // Generic 4-ref scaffold — the smallest schema-valid form per
        // CT_StyleEntry (every child is minOccurs=0).
        _ =>
            $"<cs:{section}>" +
                "<cs:lnRef idx=\"0\"/>" +
                "<cs:fillRef idx=\"0\"/>" +
                "<cs:effectRef idx=\"0\"/>" +
                "<cs:fontRef idx=\"minor\"/>" +
            $"</cs:{section}>"
    };
}

/// <summary>
/// In-memory lookup table mapping <c>(chartType, variant)</c> to a set
/// of per-section fragment IDs consumed by <see cref="ChartExStyleBuilder"/>.
/// Backed by an optional embedded resource; if the resource isn't
/// present, <see cref="TryGet"/> always returns null and the builder
/// emits <see cref="MinimalScaffold"/> everywhere.
///
/// Lazy-loaded on first access, cached for process lifetime, thread-safe
/// via double-checked lock.
/// </summary>
internal static class GalleryIndex
{
    private const string IndexResourceName =
        "OfficeCli.Resources.cx-gallery.index.json";

    private static Dictionary<string, GalleryEntry>? _cache;
    private static readonly object _cacheLock = new();

    /// <summary>
    /// Look up the style entry for a given (chartType, variant) pair.
    /// Returns null when the index has nothing for that key, in which
    /// case <see cref="ChartExStyleBuilder"/> falls back to
    /// <see cref="MinimalScaffold"/> for every section.
    /// </summary>
    internal static GalleryEntry? TryGet(string chartType, string variant)
    {
        var cache = EnsureLoaded();
        if (cache == null) return null;
        var key = $"{chartType.ToLowerInvariant()}/{variant.ToLowerInvariant()}";
        return cache.TryGetValue(key, out var entry) ? entry : null;
    }

    /// <summary>
    /// Expose the set of known (type, variant) keys for diagnostics.
    /// </summary>
    internal static IReadOnlyCollection<string> KnownKeys()
    {
        var cache = EnsureLoaded();
        return cache?.Keys ?? (IReadOnlyCollection<string>)Array.Empty<string>();
    }

    private static Dictionary<string, GalleryEntry>? EnsureLoaded()
    {
        if (_cache != null) return _cache;
        lock (_cacheLock)
        {
            if (_cache != null) return _cache;
            _cache = LoadFromEmbeddedResource() ?? new Dictionary<string, GalleryEntry>();
        }
        return _cache;
    }

    private static Dictionary<string, GalleryEntry>? LoadFromEmbeddedResource()
    {
        var assembly = typeof(GalleryIndex).Assembly;
        using var stream = assembly.GetManifestResourceStream(IndexResourceName);
        if (stream == null)
        {
            // No index resource embedded — TryGet returns null and the
            // builder falls back to minimal scaffolds for every section.
            return null;
        }

        using var doc = JsonDocument.Parse(stream);
        var root = doc.RootElement;
        if (!root.TryGetProperty("entries", out var entriesEl)
            || entriesEl.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        var result = new Dictionary<string, GalleryEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entriesEl.EnumerateObject())
        {
            var key = entry.Name.ToLowerInvariant();
            var val = entry.Value;
            if (val.ValueKind != JsonValueKind.Object) continue;

            int styleId = 410;
            if (val.TryGetProperty("styleId", out var styleIdEl)
                && styleIdEl.ValueKind == JsonValueKind.Number)
            {
                styleId = styleIdEl.GetInt32();
            }

            var fragMap = new Dictionary<string, string>(StringComparer.Ordinal);
            if (val.TryGetProperty("fragments", out var fragsEl)
                && fragsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var frag in fragsEl.EnumerateObject())
                {
                    if (frag.Value.ValueKind == JsonValueKind.String)
                    {
                        fragMap[frag.Name] = frag.Value.GetString()!;
                    }
                }
            }

            result[key] = new GalleryEntry(styleId, fragMap);
        }
        return result;
    }
}

/// <summary>
/// Record holding one (chartType, variant) entry: the numeric
/// <c>cs:chartStyle @id</c> and a map from section name to fragment ID.
/// Sections not in the map fall through to <see cref="MinimalScaffold"/>.
/// </summary>
internal sealed record GalleryEntry(
    int StyleId,
    IReadOnlyDictionary<string, string> Fragments);

/// <summary>
/// Loads individual chartStyle section fragments by their content-hash
/// ID from embedded resources. Fragments are lazily loaded on first
/// request and cached for the process lifetime. Thread-safe via a
/// lock-free <see cref="System.Collections.Concurrent.ConcurrentDictionary{TKey,TValue}"/>.
/// </summary>
internal static class FragmentStore
{
    private const string FragmentResourcePrefix =
        "OfficeCli.Resources.cx-gallery.fragments.";

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string?> _cache
        = new(StringComparer.Ordinal);

    /// <summary>
    /// Load the raw XML text of a single chartStyle section fragment
    /// by its content-hash ID. Returns null if the fragment isn't
    /// embedded — caller (<see cref="ChartExStyleBuilder"/>) then falls
    /// back to <see cref="MinimalScaffold.For"/>.
    /// </summary>
    internal static string? TryLoad(string fragmentId)
    {
        return _cache.GetOrAdd(fragmentId, LoadFromEmbeddedResource);
    }

    private static string? LoadFromEmbeddedResource(string fragmentId)
    {
        var assembly = typeof(FragmentStore).Assembly;
        var resourceName = FragmentResourcePrefix + fragmentId + ".xml";
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null) return null;
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
