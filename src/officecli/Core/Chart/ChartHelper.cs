// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Shared chart build/read/set logic used by PPTX, Excel, and Word handlers.
/// All methods operate on ChartPart / C.Chart / C.PlotArea — independent of host document type.
/// </summary>
internal static partial class ChartHelper
{
    // ==================== Parse Helpers ====================

    internal static (string kind, bool is3D, bool stacked, bool percentStacked) ParseChartType(string chartType)
    {
        var ct = SchemaKeyNormalizer.Normalize(chartType);
        var is3D = ct.EndsWith("3d") || ct.Contains("3d");
        ct = ct.Replace("3d", "");

        // OOXML has no Scatter3D or Radar3D variant — CT_ScatterChart and
        // CT_RadarChart are 2D-only. Previously `scatter3d`/`radar3d` silently
        // had the `3d` stripped and became plain scatter/radar, losing caller
        // intent (round-trip returned `scatter`/`radar`, real PowerPoint
        // rendered flat). Reject these forms explicitly with a helpful error
        // pointing at the supported alternatives.
        if (is3D && (ct == "scatter" || ct == "xy" || ct == "radar" || ct == "spider"))
        {
            throw new ArgumentException(
                $"Chart type '{chartType}' is not supported. " +
                "OOXML has no Scatter3D or Radar3D variant — both CT_ScatterChart and CT_RadarChart " +
                "are 2D-only. Use 'scatter' / 'radar' (optionally with radarStyle=filled|marker|standard).");
        }

        var stacked = ct.Contains("stacked") && !ct.Contains("percent");
        var percentStacked = ct.Contains("percentstacked") || ct.Contains("pstacked");
        ct = ct.Replace("percentstacked", "").Replace("pstacked", "").Replace("stacked", "");

        var kind = ct switch
        {
            "bar" => "bar",
            "column" or "col" => "column",
            "line" => "line",
            "pie" => "pie",
            "pieofpie" => "pieofpie",
            "barofpie" => "barofpie",
            "doughnut" or "donut" => "doughnut",
            "area" => "area",
            "scatter" or "xy" => "scatter",
            "bubble" => "bubble",
            "radar" or "spider" => "radar",
            "stock" or "ohlc" => "stock",
            "combo" => "combo",
            "waterfall" or "wf" => "waterfall",
            _ => throw new ArgumentException(
                $"Unknown chart type: '{chartType}'. Supported types: " +
                "column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo, waterfall, " +
                "funnel, treemap, sunburst, boxWhisker, histogram, pareto. " +
                "Modifiers: 3d (e.g. column3d), stacked (e.g. stackedColumn), percentStacked (e.g. percentStackedBar).")
        };

        return (kind, is3D, stacked, percentStacked);
    }

    /// <summary>
    /// Extended series info that may contain cell references instead of literal data.
    /// </summary>
    internal class SeriesInfo
    {
        public string Name { get; set; } = "";
        public double[]? Values { get; set; }
        public string? ValuesRef { get; set; }       // e.g. "Sheet1!$B$2:$B$13"
        public string? CategoriesRef { get; set; }    // e.g. "Sheet1!$A$2:$A$13"
        // R52 bt-3: bubble charts store per-point sizes in a third series
        // dimension. The literal data round-trips via Format["bubbleSize"];
        // BubbleSizeRef carries the external cell range so dump→replay
        // preserves the source <c:numRef> instead of falling back to the
        // BuildBubbleChart default (size = y-values).
        public double[]? BubbleSizeValues { get; set; }
        public string? BubbleSizeRef { get; set; }    // e.g. "[Book1.xlsx]Sheet1!$D$1:$D$3"
    }

    /// <summary>
    /// Returns true if the value looks like a cell range reference (contains '!' or matches A1:B2 pattern).
    /// </summary>
    internal static bool IsRangeReference(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (value.Contains('!')) return true;
        // Match patterns like A1:B13, $A$1:$B$13, AA1:ZZ999
        return System.Text.RegularExpressions.Regex.IsMatch(value.Trim(),
            @"^\$?[A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d+$");
    }

    /// <summary>
    /// Returns true if the value looks like a single cell reference with an
    /// explicit sheet prefix (Sheet1!A1, Sheet1!$A$1, 'My Sheet'!A1:A1). Used
    /// to detect when a series.name / chart title parameter should be emitted
    /// as a c:strRef instead of literal c:v.
    ///
    /// Bare cell-shaped tokens (e.g. "Q1", "A1", "B2") are deliberately NOT
    /// treated as cell references — they collide with common literal labels
    /// (quarter codes, product names) and emitting a strRef without an
    /// external workbook backing causes real PowerPoint to render no title /
    /// no series name at all (data loss). Callers wanting a cell reference
    /// must qualify with the sheet name.
    /// </summary>
    internal static bool IsCellReference(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        var trimmed = value.Trim();
        // Mandatory sheet prefix (Sheet1! or 'Sheet with spaces'!), single
        // cell A1 or $A$1, optionally followed by :A1 range of size 1.
        return System.Text.RegularExpressions.Regex.IsMatch(trimmed,
            @"^(?:'[^']+'!|[A-Za-z_][\w\.]*!)\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$");
    }



    /// <summary>
    /// Wraps a sheet-name prefix (the part ending with '!') in single quotes.
    /// ECMA-376 permits sheet-name prefixes to be quoted unconditionally, and
    /// Office/WPS accepts references such as 'Sheet1'!A1, '测试'!A1, and
    /// '1透视'!A1 without issue. Already-quoted prefixes are left unchanged.
    /// </summary>
    private static string QuoteSheetNamePrefix(string sheetPartWithBang)
    {
        if (string.IsNullOrEmpty(sheetPartWithBang)) return sheetPartWithBang;
        // sheetPartWithBang ends with '!'
        var prefix = sheetPartWithBang[..^1];
        if (string.IsNullOrEmpty(prefix)) return sheetPartWithBang;

        // Already quoted — leave as-is.
        if (prefix.Length >= 2 && prefix[0] == '\'' && prefix[^1] == '\'')
            return sheetPartWithBang;

        // Escape embedded single quotes by doubling them, then wrap the whole
        // prefix in single quotes. This is valid for any sheet name per ECMA-376.
        var escaped = prefix.Replace("'", "''");
        return "'" + escaped + "'!";
    }

    /// <summary>
    /// Normalizes a single-cell reference for use inside a chart's c:strRef/c:f.
    /// Ensures absolute ($col$row) form, preserves any sheet prefix, and quotes
    /// the sheet prefix unconditionally per ECMA-376.
    /// </summary>
    internal static string NormalizeCellReference(string value)
    {
        var trimmed = value.Trim();
        string sheetPart = "";
        string cellPart = trimmed;
        var bangIdx = trimmed.IndexOf('!');
        if (bangIdx >= 0)
        {
            sheetPart = trimmed[..(bangIdx + 1)];
            cellPart = trimmed[(bangIdx + 1)..];
        }
        var parts = cellPart.Split(':');
        for (int i = 0; i < parts.Length; i++)
            parts[i] = AddAbsoluteMarkers(parts[i]);
        return QuoteSheetNamePrefix(sheetPart) + string.Join(":", parts);
    }


    /// <summary>
    /// Normalizes a range reference by adding $ signs for absolute references.
    /// If no sheet prefix, prepends defaultSheet. Sheet-name prefixes are quoted
    /// unconditionally per ECMA-376.
    /// </summary>
    internal static string NormalizeRangeReference(string value, string? defaultSheet = null)
    {
        var trimmed = value.Trim();
        string sheetPart = "";
        string rangePart = trimmed;

        var bangIdx = trimmed.IndexOf('!');
        if (bangIdx >= 0)
        {
            sheetPart = trimmed[..(bangIdx + 1)];
            rangePart = trimmed[(bangIdx + 1)..];
        }
        else if (!string.IsNullOrEmpty(defaultSheet))
        {
            sheetPart = defaultSheet + "!";
        }

        // Add $ signs to cell refs if not already present
        var parts = rangePart.Split(':');
        for (int i = 0; i < parts.Length; i++)
            parts[i] = AddAbsoluteMarkers(parts[i]);

        return QuoteSheetNamePrefix(sheetPart) + string.Join(":", parts);
    }

    private static string AddAbsoluteMarkers(string cellRef)
    {
        // Already has $ signs — return as-is
        if (cellRef.Contains('$')) return cellRef;

        // Split into column letters and row digits
        int firstDigit = 0;
        for (int i = 0; i < cellRef.Length; i++)
        {
            if (char.IsDigit(cellRef[i])) { firstDigit = i; break; }
        }
        if (firstDigit == 0) return cellRef; // no digits found

        var col = cellRef[..firstDigit];
        var row = cellRef[firstDigit..];
        return $"${col}${row}";
    }

    /// <summary>
    /// Parse series data supporting both legacy format and new dotted syntax with cell references.
    /// Dotted syntax: series1.name=Sales, series1.values=Sheet1!B2:B13, series1.categories=Sheet1!A2:A13
    /// Legacy: series1=Sales:10,20,30 or data=Sales:10,20,30;Cost:5,8,12
    /// </summary>
    internal static List<(string name, double[] values)> ParseSeriesData(Dictionary<string, string> properties)
    {
        // CONSISTENCY(chart-series-name-alias): `series{N}Name=` flat form
        // is a natural alias for the dotted `series{N}.name=`. Rewrite so
        // both the dotted and legacy branches below see the canonical
        // `series{N}.name` key. TryGetValue on the original `series{N}Name`
        // key still fires (preserved tracking) — we add the dotted alias.
        NormalizeFlatSeriesNameAliases(properties);

        // Check for dotted syntax first
        var extSeries = ParseSeriesDataExtended(properties);
        if (extSeries != null && extSeries.Count > 0 && extSeries.Any(s => s.ValuesRef != null || s.CategoriesRef != null))
        {
            // Dotted syntax with references. Literal values may still ride
            // alongside via `data=` / `series{N}=` (the batch emitter dumps
            // both the cached point values AND the source cell-range ref) —
            // merge them positionally so the built numLit carries the real
            // data and the numRef rewrite preserves it as numCache. Without
            // the merge a ref-bearing series came back with an EMPTY value
            // list → numRef with no numCache → real PowerPoint silently
            // renders a blank chart.
            var literalSeries = ParseLiteralSeriesData(properties);
            var merged = new List<(string name, double[] values)>(extSeries.Count);
            for (int i = 0; i < extSeries.Count; i++)
            {
                var s = extSeries[i];
                var name = s.Name;
                var vals = s.Values ?? Array.Empty<double>();
                if (i < literalSeries.Count)
                {
                    if (vals.Length == 0) vals = literalSeries[i].values;
                    if (!properties.ContainsKey($"series{i + 1}.name")
                        && literalSeries[i].name.Length > 0)
                        name = literalSeries[i].name;
                }
                merged.Add((name, vals));
            }
            for (int i = extSeries.Count; i < literalSeries.Count; i++)
                merged.Add(literalSeries[i]);
            return merged;
        }

        return ParseLiteralSeriesData(properties);
    }

    /// <summary>
    /// Parse literal series data from `data=Name1:v1,v2;...` or the legacy
    /// `series{N}=Name:v1,v2` / dotted `series{N}.values=v1,v2` keys.
    /// </summary>
    private static List<(string name, double[] values)> ParseLiteralSeriesData(Dictionary<string, string> properties)
    {
        var result = new List<(string name, double[] values)>();

        if (properties.TryGetValue("data", out var dataStr))
        {
            // Determine series delimiter: use ';' if present, otherwise detect
            // comma-separated name:value pairs (e.g. "Q1:40,Q2:55,Q3:70")
            string[] seriesParts;
            if (dataStr.Contains(';'))
            {
                seriesParts = dataStr.Split(';', StringSplitOptions.RemoveEmptyEntries);
            }
            else
            {
                // Check if comma-separated parts each contain a colon (name:value pairs)
                var commaParts = dataStr.Split(',', StringSplitOptions.RemoveEmptyEntries);
                if (commaParts.Length > 1 && commaParts.All(p => p.Contains(':')))
                    seriesParts = commaParts;
                else
                    seriesParts = new[] { dataStr };
            }

            foreach (var seriesPart in seriesParts)
            {
                // Split on the LAST colon: the value list is colon-free, so the
                // rightmost colon separates a (possibly colon-containing) series
                // name from its values, e.g. "Persons (Data year: 2021):1,2,3".
                var colonIdx = seriesPart.LastIndexOf(':');
                if (colonIdx < 0) continue;
                var name = seriesPart[..colonIdx].Trim();
                var valStr = seriesPart[(colonIdx + 1)..].Trim();
                if (string.IsNullOrEmpty(valStr))
                    throw new ArgumentException($"Series '{name}' has no data values. Expected format: 'Name:1,2,3'");
                var vals = ParseSeriesValues(valStr, name);
                result.Add((name, vals));
            }
            return result;
        }

        for (int i = 1; i <= 20; i++)
        {
            // Read both keys up front so TrackingPropertyDictionary marks each
            // consumed regardless of which branch supplies the values. Without
            // this, `series{i}=` combined with `series{i}.name=` fell through
            // to UNSUPPORTED + silent value drop (interview-edit R4 major).
            var hasDotName = properties.TryGetValue($"series{i}.name", out var dotName);
            var hasDotValues = properties.TryGetValue($"series{i}.values", out var dotValues);
            var hasLegacy = properties.TryGetValue($"series{i}", out var legacyStr);

            // Check for dotted syntax first: series1.name, series1.values
            if (hasDotName || hasDotValues)
            {
                var name = dotName ?? $"Series {i}";
                // CONSISTENCY(chart-series-mixed): when dotted .name is given
                // without dotted .values, fall back to the legacy `series{i}=`
                // key as the values source rather than silently dropping it.
                var valuesStr = !string.IsNullOrEmpty(dotValues) ? dotValues
                              : (hasLegacy ? legacyStr : "");
                if (!string.IsNullOrEmpty(valuesStr) && !IsRangeReference(valuesStr))
                {
                    var vals = ParseSeriesValues(valuesStr, name);
                    result.Add((name, vals));
                }
                else
                {
                    // Reference-based — add empty placeholder (actual ref handled by BuildChartSpace)
                    result.Add((name, Array.Empty<double>()));
                }
                continue;
            }

            // Legacy format: series1=Sales:10,20,30
            if (!hasLegacy) continue;
            var seriesStr = legacyStr!;
            // CONSISTENCY(chart-series-rangeref): mirror the dotted-syntax
            // guard above (line 253) — if the legacy value is a range
            // reference (e.g. "Sheet1!B2:B7"), don't try to parse it as
            // literal comma-separated numbers. ApplySeriesReferences picks
            // it up later via the series{N} key. Without this guard
            // ParseSeriesValues throws "Invalid data value 'B7'".
            if (IsRangeReference(seriesStr))
            {
                result.Add(($"Series {i}", Array.Empty<double>()));
                continue;
            }
            // Split on the LAST colon: values are colon-free comma numbers, so
            // the rightmost colon separates a (possibly colon-containing) series
            // name from its values (e.g. "Persons (Data year: 2021):1,2,3").
            var colonIdx = seriesStr.LastIndexOf(':');
            if (colonIdx < 0)
            {
                var vals = ParseSeriesValues(seriesStr, $"series{i}");
                result.Add(($"Series {i}", vals));
            }
            else
            {
                var name = seriesStr[..colonIdx].Trim();
                var vals = ParseSeriesValues(seriesStr[(colonIdx + 1)..], name);
                result.Add((name, vals));
            }
        }

        return result;
    }

    /// <summary>
    /// Parse extended series data with cell references support.
    /// Returns null if no dotted syntax series found.
    /// </summary>
    internal static List<SeriesInfo>? ParseSeriesDataExtended(Dictionary<string, string> properties)
    {
        // Same flat-name alias rewrite as ParseSeriesData entry; idempotent
        // when called second time.
        NormalizeFlatSeriesNameAliases(properties);

        var result = new List<SeriesInfo>();

        for (int i = 1; i <= 20; i++)
        {
            var hasName = properties.TryGetValue($"series{i}.name", out var nameStr);
            var hasValues = properties.TryGetValue($"series{i}.values", out var valuesStr);
            // `series{N}.valuesRef=<range>` carries the source cell-range ref
            // WITHOUT displacing literal values (unlike `.values=<range>`).
            // The batch emitter dumps it from the Reader's per-series
            // Format["valuesRef"] so dump→replay rebuilds numRef + numCache.
            var hasValuesRef = properties.TryGetValue($"series{i}.valuesRef", out var valuesRefStr);
            var hasCats = properties.TryGetValue($"series{i}.categories", out var catsStr);
            var hasBubbleSize = properties.TryGetValue($"series{i}.bubbleSize", out var bsStr);
            var hasBubbleSizeRef = properties.ContainsKey($"series{i}.bubbleSizeRef");

            // CONSISTENCY(chart-series-rangeref): legacy `series{N}=Sheet1!B2:B3`
            // range-ref form mirrors `series{N}.values=<range>`. ParseSeriesData
            // emits an empty literal placeholder for it (line 289); pick the range
            // up here as a ValuesRef so the series gets a numRef, not a blank val.
            string? legacyRangeRef = null;
            if (!hasValues
                && properties.TryGetValue($"series{i}", out var legacyStr)
                && !string.IsNullOrEmpty(legacyStr)
                && IsRangeReference(legacyStr))
            {
                legacyRangeRef = legacyStr;
            }

            if (!hasName && !hasValues && !hasValuesRef && !hasCats && !hasBubbleSize && !hasBubbleSizeRef && legacyRangeRef == null) continue;

            var info = new SeriesInfo { Name = nameStr ?? $"Series {i}" };

            if (!string.IsNullOrEmpty(valuesStr))
            {
                if (IsRangeReference(valuesStr))
                    info.ValuesRef = NormalizeRangeReference(valuesStr);
                else
                    info.Values = ParseSeriesValues(valuesStr, info.Name);
            }
            else if (legacyRangeRef != null)
            {
                info.ValuesRef = NormalizeRangeReference(legacyRangeRef);
            }

            if (info.ValuesRef == null && !string.IsNullOrEmpty(valuesRefStr)
                && IsRangeReference(valuesRefStr))
            {
                info.ValuesRef = NormalizeRangeReference(valuesRefStr);
            }

            if (!string.IsNullOrEmpty(catsStr))
            {
                if (IsRangeReference(catsStr))
                    info.CategoriesRef = NormalizeRangeReference(catsStr);
            }

            // R52 bt-3: bubble series carry per-point sizes. `series{N}.bubbleSize`
            // accepts either a comma-separated literal list or a cell range; the
            // range form preserves the source <c:numRef> so PowerPoint shows the
            // correct bubble pixel-geometry from the linked workbook data.
            if (!string.IsNullOrEmpty(bsStr))
            {
                if (IsRangeReference(bsStr))
                    info.BubbleSizeRef = NormalizeRangeReference(bsStr);
                else
                    info.BubbleSizeValues = ParseSeriesValues(bsStr, info.Name + ".bubbleSize");
            }
            // `series{N}.bubbleSizeRef` is the explicit ref key emitted by the
            // batch emitter when a source <c:bubbleSize><c:numRef> is detected.
            // It coexists with the literal `series{N}.bubbleSize=cached,values`
            // — the ref takes precedence, the literal becomes the numCache.
            if (properties.TryGetValue($"series{i}.bubbleSizeRef", out var bsRefStr)
                && !string.IsNullOrEmpty(bsRefStr))
            {
                info.BubbleSizeRef = NormalizeRangeReference(bsRefStr);
            }

            result.Add(info);
        }

        return result.Count > 0 ? result : null;
    }

    /// <summary>
    /// Parse the top-level categories property, supporting both literal and reference values.
    /// Returns the reference string if it's a range reference, null otherwise (literal handled separately).
    /// </summary>
    internal static string? ParseCategoriesRef(Dictionary<string, string> properties)
    {
        // Explicit `categoriesRef=<range>` (the Reader/batch-emitter form:
        // literal labels in `categories=` for the strCache, the source cell
        // range here) wins over range-form `categories=`.
        if (properties.TryGetValue("categoriesRef", out var catRefStr)
            && IsRangeReference(catRefStr))
            return NormalizeRangeReference(catRefStr);
        if (!properties.TryGetValue("categories", out var catStr)) return null;
        if (IsRangeReference(catStr)) return NormalizeRangeReference(catStr);
        return null;
    }

    // CONSISTENCY(chart-series-name-alias): `series{N}Name=Revenue` flat form
    // → `series{N}.name=Revenue` dotted form. Records the read on the
    // original flat key (via TryGetValue) so handler-as-truth tracking still
    // marks the user input consumed.
    private static void NormalizeFlatSeriesNameAliases(Dictionary<string, string> properties)
    {
        for (int i = 1; i <= 20; i++)
        {
            var flatKey = $"series{i}Name";
            if (properties.TryGetValue(flatKey, out var nameVal))
            {
                var dottedKey = $"series{i}.name";
                if (!properties.ContainsKey(dottedKey))
                    properties[dottedKey] = nameVal;
                // Remove the flat key so SetChartProperties' dispatch loop
                // doesn't also iterate it — the legacy `series{N}=` branch
                // does int.TryParse on the suffix ("2Name") and reports it
                // as unsupported. TryGetValue above already marked the
                // original input consumed for handler-as-truth tracking.
                properties.Remove(flatKey);
            }
        }
    }

    private static double[] ParseSeriesValues(string valStr, string seriesName)
    {
        return valStr.Split(',').Select(v =>
        {
            var trimmed = v.Trim();
            if (!double.TryParse(trimmed, System.Globalization.CultureInfo.InvariantCulture, out var num)
                || double.IsNaN(num) || double.IsInfinity(num))
                throw new ArgumentException($"Invalid data value '{trimmed}' in series '{seriesName}'. Expected comma-separated finite numbers (e.g. '1,2,3').");
            return num;
        }).ToArray();
    }

    internal static string[]? ParseCategories(Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("categories", out var catStr)) return null;
        // If the value is a cell range reference, don't treat as literal categories
        if (IsRangeReference(catStr)) return null;
        return catStr.Split(',').Select(c => c.Trim()).ToArray();
    }

    // BUG-DUMP-R36-3: series indices (0-based) whose dump carried per-point
    // <c:dPt> styling (and/or a verbatim series <c:spPr>) but NO series-level
    // fill key (series{N}.color / .gradient / .spPr). For these the source had
    // no series-level <c:spPr>; the chart builder must therefore NOT inject the
    // Office accent1 default — doing so plants a spurious solidFill that a
    // partial-dPt series would visibly show. Mirrors the spec "only emit a
    // series spPr when one was actually captured". Fresh `add chart` (no dPt
    // keys) is unaffected and keeps its default palette.
    internal static HashSet<int> SeriesWithNoCapturedFill(Dictionary<string, string> properties)
    {
        var result = new HashSet<int>();
        foreach (var kv in properties)
        {
            var k = kv.Key;
            if (!k.StartsWith("series", StringComparison.OrdinalIgnoreCase)) continue;
            int dot = k.IndexOf('.');
            if (dot <= 6) continue;
            var mid = k.Substring(6, dot - 6);
            if (!int.TryParse(mid, out var idx1) || idx1 < 1) continue;
            var sub = k.Substring(dot + 1);
            // A series carrying verbatim per-point styling (dPt) or a verbatim
            // series spPr is the replay signal. We only flag it for default
            // suppression when it does NOT also carry a series-level fill key.
            if (!sub.Equals("dPt", StringComparison.OrdinalIgnoreCase)
                && !sub.Equals("spPr", StringComparison.OrdinalIgnoreCase)
                && !sub.Equals("inheritFill", StringComparison.OrdinalIgnoreCase)) continue;
            // `series{N}.inheritFill=true`: the dump saw a source series with
            // no <c:spPr> at all — the fill is theme-inherited, so the replay
            // must not pin a DefaultSeriesColors palette entry over it. The
            // TryGetValue registers the key as consumed (handler-as-truth
            // tracking doesn't fire on Dictionary-typed foreach iteration).
            if (sub.Equals("inheritFill", StringComparison.OrdinalIgnoreCase))
                properties.TryGetValue(k, out _);
            int idx0 = idx1 - 1;
            // A series{N}.spPr key IS a captured series fill — keep it (the
            // verbatim spPr is applied post-build), so do not flag spPr-bearing
            // series. Only dPt-only series (no series-level fill) get flagged.
            if (sub.Equals("spPr", StringComparison.OrdinalIgnoreCase)) continue;
            bool hasSeriesFill =
                properties.ContainsKey($"series{idx1}.color")
                || properties.ContainsKey($"series{idx1}.gradient")
                || properties.ContainsKey($"series{idx1}.spPr");
            if (!hasSeriesFill) result.Add(idx0);
        }
        return result;
    }

    internal static string[]? ParseSeriesColors(Dictionary<string, string> properties)
    {
        // CONSISTENCY(chart-series-color): Add path accepts both the
        // compact `colors=red,blue,green` form and per-series dotted
        // `series{N}.color=<hex>` keys (same vocabulary that `set chart`
        // already supports via ApplySeriesColor). When both are supplied,
        // dotted keys override positions in the `colors` array.
        string[]? arr = null;
        if (properties.TryGetValue("colors", out var colorsStr) || properties.TryGetValue("seriesColors", out colorsStr))
            arr = colorsStr.Split(',').Select(c => c.Trim()).ToArray();

        // Collect per-series dotted color keys
        var dotted = new Dictionary<int, string>();
        foreach (var kv in properties)
        {
            var k = kv.Key;
            if (!k.StartsWith("series", StringComparison.OrdinalIgnoreCase)) continue;
            if (!k.EndsWith(".color", StringComparison.OrdinalIgnoreCase)) continue;
            var mid = k.Substring(6, k.Length - 6 - ".color".Length);
            if (!int.TryParse(mid, out var idx) || idx < 1) continue;
            // Mark this series{N}.color key consumed. A `foreach` over a
            // Dictionary<>-typed parameter binds to the base struct enumerator,
            // which bypasses TrackingPropertyDictionary's consumption tracking
            // (only TryGetValue/ContainsKey/indexer route through the recording
            // comparer). Without this ContainsKey, the Add path reports an
            // applied series color as UNSUPPORTED and exits non-zero.
            _ = properties.ContainsKey(k);
            if (!string.IsNullOrWhiteSpace(kv.Value))
                dotted[idx] = kv.Value.Trim();
        }
        if (dotted.Count == 0) return arr;

        var maxIdx = dotted.Keys.Max();
        var size = Math.Max(maxIdx, arr?.Length ?? 0);
        var merged = new string[size];
        for (int i = 0; i < size; i++)
        {
            if (dotted.TryGetValue(i + 1, out var c))
                merged[i] = c;
            else if (arr != null && i < arr.Length && !string.IsNullOrEmpty(arr[i]))
                merged[i] = arr[i];
            else
                merged[i] = DefaultSeriesColors[i % DefaultSeriesColors.Length];
        }
        return merged;
    }

    /// <summary>
    /// Like ParseSeriesColors but returns ONLY the per-series dotted color
    /// keys (`series{N}.color=<hex>`), without merging the positional
    /// `colors=` array. Used by single-series chart builders
    /// (pie/doughnut/stock) where positional `colors=` is per-data-point but
    /// `series{N}.color` should still tint the whole series.
    /// </summary>
    internal static Dictionary<int, string> ParseDottedSeriesColorsOnly(Dictionary<string, string> properties)
    {
        var dotted = new Dictionary<int, string>();
        foreach (var kv in properties)
        {
            var k = kv.Key;
            if (!k.StartsWith("series", StringComparison.OrdinalIgnoreCase)) continue;
            if (!k.EndsWith(".color", StringComparison.OrdinalIgnoreCase)) continue;
            var mid = k.Substring(6, k.Length - 6 - ".color".Length);
            if (!int.TryParse(mid, out var idx) || idx < 1) continue;
            // Mark consumed — see ParseSeriesColors: foreach over a
            // Dictionary<>-typed param bypasses the tracking enumerator.
            _ = properties.ContainsKey(k);
            if (!string.IsNullOrWhiteSpace(kv.Value))
                dotted[idx] = kv.Value.Trim();
        }
        return dotted;
    }

    /// <summary>
    /// Positional-only `colors=` / `seriesColors=` array, returned unmerged.
    /// Used by pie / doughnut / pieofpie / barofpie / stock builders where the
    /// positional list maps to per-data-point dPt overrides, but `series{N}.
    /// color` MUST NOT bleed into the dPt list (it routes through
    /// ApplyDottedSeriesColors as a series-level fill instead). The merged
    /// array from ParseSeriesColors conflates the two, so a doughnut with
    /// only `series1.color=#C00000` emitted a spurious dPt#0 with the same
    /// fill — Get readback then surfaced both `series1.color` AND
    /// `point1.color`, breaking dump→replay byte-equality.
    /// </summary>
    internal static string[]? ParsePositionalColorsOnly(Dictionary<string, string> properties)
    {
        if (properties.TryGetValue("colors", out var colorsStr)
            || properties.TryGetValue("seriesColors", out colorsStr))
            return colorsStr.Split(',').Select(c => c.Trim()).ToArray();
        return null;
    }

    // ==================== ManualLayout Helpers ====================

    /// <summary>
    /// Ensures the given element has a Layout > ManualLayout child and sets the specified
    /// positional property (x/y/w/h). Creates Layout and ManualLayout if missing.
    /// For plotArea, LayoutTarget is set to Inner; for others it is omitted.
    /// </summary>
    internal static void SetManualLayoutProperty(OpenXmlCompositeElement parent, string prop, double value, bool isPlotArea = false)
    {
        var layout = parent.GetFirstChild<C.Layout>();
        if (layout == null)
        {
            layout = new C.Layout();
            // Insert layout after structural elements to respect schema order.
            // c:title  → tx, [layout], overlay, ...
            // c:legend → legendPos, legendEntry*, [layout], overlay, ...
            // c:dLbl   → idx, delete, [layout], ...
            // c:plotArea → layout is first child (InsertAt 0 is correct)
            if (isPlotArea)
            {
                parent.InsertAt(layout, 0);
            }
            else if (parent is C.DataLabel)
            {
                // CT_DLbl: idx, delete, [layout], tx, numFmt, spPr, ...
                var insertAfter = parent.GetFirstChild<C.Delete>() as OpenXmlElement
                    ?? parent.GetFirstChild<C.Index>() as OpenXmlElement;
                if (insertAfter != null)
                    insertAfter.InsertAfterSelf(layout);
                else
                    parent.InsertAt(layout, 0);
            }
            else
            {
                // c:title  → tx, [layout], overlay, ...
                // c:legend → legendPos, legendEntry*, [layout], overlay, ...
                var insertAfter = parent.GetFirstChild<C.ChartText>() as OpenXmlElement
                    ?? parent.ChildElements.LastOrDefault(
                        e => e.LocalName is "legendPos" or "legendEntry") as OpenXmlElement;
                if (insertAfter != null)
                    insertAfter.InsertAfterSelf(layout);
                else
                    parent.InsertAt(layout, 0);
            }
        }
        var ml = layout.GetFirstChild<C.ManualLayout>();
        if (ml == null)
        {
            ml = new C.ManualLayout();
            if (isPlotArea)
                ml.LayoutTarget = new C.LayoutTarget { Val = C.LayoutTargetValues.Inner };
            ml.LeftMode = new C.LeftMode { Val = C.LayoutModeValues.Edge };
            ml.TopMode = new C.TopMode { Val = C.LayoutModeValues.Edge };
            layout.AppendChild(ml);
        }
        // Use typed properties to guarantee schema order (OneSequence)
        switch (prop)
        {
            case "x": ml.Left = new C.Left { Val = value }; break;
            case "y": ml.Top = new C.Top { Val = value }; break;
            case "w": ml.Width = new C.Width { Val = value }; break;
            case "h": ml.Height = new C.Height { Val = value }; break;
        }
    }

    /// <summary>
    /// Read ManualLayout x/y/w/h from an element that has Layout as a child.
    /// Writes results into node.Format with the given prefix (e.g. "plotArea", "title", "legend").
    /// </summary>
    internal static void ReadManualLayout(OpenXmlCompositeElement parent, DocumentNode node, string prefix)
    {
        var layout = parent.GetFirstChild<C.Layout>();
        var ml = layout?.GetFirstChild<C.ManualLayout>();
        if (ml == null) return;

        var x = ml.Left?.Val?.Value;
        var y = ml.Top?.Val?.Value;
        var w = ml.Width?.Val?.Value;
        var h = ml.Height?.Val?.Value;

        if (x != null) node.Format[$"{prefix}.x"] = x.Value.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture);
        if (y != null) node.Format[$"{prefix}.y"] = y.Value.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture);
        if (w != null) node.Format[$"{prefix}.w"] = w.Value.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture);
        if (h != null) node.Format[$"{prefix}.h"] = h.Value.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture);
    }
}
