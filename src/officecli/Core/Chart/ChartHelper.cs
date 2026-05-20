// Copyright 2025 OfficeCLI (officecli.ai)
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
    /// Normalizes a single-cell reference for use inside a chart's c:strRef/c:f.
    /// Ensures absolute ($col$row) form and preserves any sheet prefix. If the
    /// input is a A1:A1 style single-cell range, the range form is kept so the
    /// output matches what Excel writes when a user points the Name field at a
    /// single cell via the dialog.
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
        return sheetPart + string.Join(":", parts);
    }

    /// <summary>
    /// Normalizes a range reference by adding $ signs for absolute references.
    /// If no sheet prefix, prepends defaultSheet.
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

        return sheetPart + string.Join(":", parts);
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
        // Check for dotted syntax first
        var extSeries = ParseSeriesDataExtended(properties);
        if (extSeries != null && extSeries.Count > 0 && extSeries.Any(s => s.ValuesRef != null || s.CategoriesRef != null))
        {
            // Dotted syntax with references — return literal values where available, empty arrays for references
            return extSeries.Select(s => (s.Name, s.Values ?? Array.Empty<double>())).ToList();
        }

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
                var colonIdx = seriesPart.IndexOf(':');
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
            // Check for dotted syntax first: series1.name, series1.values
            if (properties.ContainsKey($"series{i}.values") || properties.ContainsKey($"series{i}.name"))
            {
                var name = properties.GetValueOrDefault($"series{i}.name") ?? $"Series {i}";
                var valuesStr = properties.GetValueOrDefault($"series{i}.values") ?? "";
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
            if (!properties.TryGetValue($"series{i}", out var seriesStr)) continue;
            var colonIdx = seriesStr.IndexOf(':');
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
        var result = new List<SeriesInfo>();

        for (int i = 1; i <= 20; i++)
        {
            var hasName = properties.TryGetValue($"series{i}.name", out var nameStr);
            var hasValues = properties.TryGetValue($"series{i}.values", out var valuesStr);
            var hasCats = properties.TryGetValue($"series{i}.categories", out var catsStr);

            if (!hasName && !hasValues && !hasCats) continue;

            var info = new SeriesInfo { Name = nameStr ?? $"Series {i}" };

            if (!string.IsNullOrEmpty(valuesStr))
            {
                if (IsRangeReference(valuesStr))
                    info.ValuesRef = NormalizeRangeReference(valuesStr);
                else
                    info.Values = ParseSeriesValues(valuesStr, info.Name);
            }

            if (!string.IsNullOrEmpty(catsStr))
            {
                if (IsRangeReference(catsStr))
                    info.CategoriesRef = NormalizeRangeReference(catsStr);
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
        if (!properties.TryGetValue("categories", out var catStr)) return null;
        if (IsRangeReference(catStr)) return NormalizeRangeReference(catStr);
        return null;
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

    internal static string[]? ParseSeriesColors(Dictionary<string, string> properties)
    {
        // CONSISTENCY(chart-series-color): Add path accepts both the
        // compact `colors=red,blue,green` form and per-series dotted
        // `series{N}.color=<hex>` keys (same vocabulary that `set chart`
        // already supports via ApplySeriesColor). When both are supplied,
        // dotted keys override positions in the `colors` array.
        string[]? arr = null;
        if (properties.TryGetValue("colors", out var colorsStr))
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
            if (!string.IsNullOrWhiteSpace(kv.Value))
                dotted[idx] = kv.Value.Trim();
        }
        return dotted;
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
