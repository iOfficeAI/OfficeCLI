// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Advanced chart features: reference lines, conditional coloring, waterfall simulation.
/// </summary>
internal static partial class ChartHelper
{
    // ==================== Reference Line ====================

    /// <summary>
    /// Add a reference (target/average) line to a chart by inserting a hidden line series.
    /// Format (positional, ':'-separated):
    ///   value
    ///   value:color
    ///   value:color:label
    ///   value:color:width:dash      (4 parts, if parts[2] is numeric and parts[3] is a known dash style)
    ///   value:color:label:dash      (4 parts, legacy — parts[2] is non-numeric)
    ///   value:color:width:dash:label (5 parts, canonical — parts[2] may be empty for default width)
    /// Width is in points (default 1.5pt). Dash style: solid/dot/dash/dashdot/longdash/longdashdot/longdashdotdot.
    /// e.g. "50", "75:FF0000", "100:00AA00:Target", "80:0000FF:Average:dash",
    ///      "50:FF0000:2.5:dash", "50:FF0000:2:dash:Target", "50:FF0000::dash:Target"
    /// </summary>
    internal static void AddReferenceLine(C.Chart chart, string spec, bool removeExisting = true)
    {
        const double DefaultWidthPt = 1.5;
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        // Caller may suppress the sweep when accumulating multiple lines from
        // a semicolon-joined value (see Setter `case "referenceline"`).
        if (removeExisting)
            RemoveExistingReferenceLines(plotArea);

        var parts = spec.Split(':');
        if (!double.TryParse(parts[0].Trim(),
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var refValue))
            throw new ArgumentException(
                $"Invalid referenceLine value '{parts[0]}'. Expected: number or number:color:label:dash (e.g. '50:FF0000:Target:dash') or number:color:width:dash (e.g. '50:FF0000:2:dash').");

        var color = parts.Length > 1 ? parts[1].Trim() : "FF0000";
        double widthPt = DefaultWidthPt;
        string label = $"Ref ({refValue.ToString("G", System.Globalization.CultureInfo.InvariantCulture)})";
        string dash = "dash";

        // Positional parse — see doc comment above. parts[0..1] already consumed.
        if (parts.Length == 3)
        {
            label = parts[2].Trim();
        }
        else if (parts.Length == 4)
        {
            var p2 = parts[2].Trim();
            var p3 = parts[3].Trim();
            // Disambiguate: "50:FF0000:2.5:dash" (width form) vs "50:FF0000:Target:dash" (legacy label form).
            // Only treat p2 as width if it parses as a number AND p3 is a recognized dash keyword — both
            // conditions together make the "ergonomic" width interpretation unambiguous.
            if (double.TryParse(p2, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var w4)
                && IsKnownDashStyle(p3))
            {
                widthPt = w4;
                dash = p3;
            }
            else
            {
                label = p2;
                dash = p3;
            }
        }
        else if (parts.Length >= 5)
        {
            // Canonical 5-part form: value:color:width:dash:label (extra parts after label are joined
            // back with ':' so labels containing literal colons survive a round-trip).
            var widthStr = parts[2].Trim();
            if (widthStr.Length > 0)
            {
                if (!double.TryParse(widthStr, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out widthPt))
                    throw new ArgumentException(
                        $"Invalid referenceLine width '{widthStr}'. Expected a number in points (e.g. '1.5'), or empty for default {DefaultWidthPt}pt.");
            }
            dash = parts[3].Trim();
            label = string.Join(':', parts.Skip(4)).Trim();
        }

        if (widthPt <= 0 || widthPt > 100)
            throw new ArgumentException(
                $"Invalid referenceLine width '{widthPt.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}'. Expected a positive number of points, typically 0.25–10.");

        // Warn: percent-stacked value axis is 0-1 (displayed 0%-100%). A refValue > 1
        // is almost always a mistake — user likely forgot to convert 50 → 0.5.
        // Without this check, Excel silently stretches the val axis to fit (e.g. 5000%),
        // producing a chart where the real bars are compressed to a thin sliver on the left.
        if (refValue > 1.0 && IsPercentStackedChart(plotArea))
        {
            var refMsg =
                $"referenceLine value {refValue.ToString("G", System.Globalization.CultureInfo.InvariantCulture)} "
                + "on a percent-stacked chart. The value axis is 0-1 (0%-100%); "
                + $"did you mean {(refValue / 100.0).ToString("G", System.Globalization.CultureInfo.InvariantCulture)}? "
                + "Excel will auto-scale the axis to fit, compressing the real bars.";
            // CONSISTENCY(numfmt-warning): JSON mode → envelope warnings[];
            // plain mode keeps the stderr line.
            if (WarningContext.IsActive)
                WarningContext.Add(refMsg, "referenceline_out_of_scale",
                    $"Use {(refValue / 100.0).ToString("G", System.Globalization.CultureInfo.InvariantCulture)} for a 0-1 percent axis");
            else
                Console.Error.WriteLine($"Warning: {refMsg}");
        }

        // Find max data point count from existing series (after removing old ref lines)
        var existingSerCount = CountSeries(plotArea);
        var maxDataPoints = 0;
        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
        {
            var vals = ser.GetFirstChild<C.Values>();
            var numLit = vals?.GetFirstChild<C.NumberLiteral>();
            var ptCount = numLit?.GetFirstChild<C.PointCount>()?.Val?.Value ?? 0;
            if ((int)ptCount > maxDataPoints) maxDataPoints = (int)ptCount;
            var numRef = vals?.GetFirstChild<C.NumberReference>();
            var cacheCount = numRef?.GetFirstChild<C.NumberingCache>()?.GetFirstChild<C.PointCount>()?.Val?.Value ?? 0;
            if ((int)cacheCount > maxDataPoints) maxDataPoints = (int)cacheCount;
        }
        if (maxDataPoints == 0) maxDataPoints = 3;

        // Create a flat line series (all values = refValue)
        var refValues = Enumerable.Repeat(refValue, maxDataPoints).ToArray();
        var seriesIdx = (uint)existingSerCount;

        // Find or create a LineChart in the plot area for the reference line
        var lineChart = plotArea.GetFirstChild<C.LineChart>();
        if (lineChart == null)
        {
            // Create a new line chart overlay — shares axes with existing chart
            uint catAxisId = 1, valAxisId = 2;
            // Try to find existing axis IDs
            var existingCatAx = plotArea.GetFirstChild<C.CategoryAxis>()?.GetFirstChild<C.AxisId>()?.Val?.Value;
            var existingValAx = plotArea.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.AxisId>()?.Val?.Value;
            if (existingCatAx != null) catAxisId = existingCatAx.Value;
            if (existingValAx != null) valAxisId = existingValAx.Value;

            lineChart = new C.LineChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false }
            );
            lineChart.AppendChild(new C.ShowMarker { Val = false });
            lineChart.AppendChild(new C.AxisId { Val = catAxisId });
            lineChart.AppendChild(new C.AxisId { Val = valAxisId });

            // Insert before axes
            var firstAxis = plotArea.Elements<C.CategoryAxis>().FirstOrDefault() as OpenXmlElement
                ?? plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (firstAxis != null)
                plotArea.InsertBefore(lineChart, firstAxis);
            else
                plotArea.AppendChild(lineChart);
        }

        // Build the reference line series
        var refSer = new C.LineChartSeries();
        refSer.AppendChild(new C.Index { Val = seriesIdx });
        refSer.AppendChild(new C.Order { Val = seriesIdx });
        refSer.AppendChild(new C.SeriesText(new C.NumericValue(label)));

        // Style: colored dashed line, no markers. Width is pt → EMU (1pt = 12700 EMU).
        var spPr = new C.ChartShapeProperties();
        var outline = new Drawing.Outline { Width = (int)Math.Round(widthPt * EmuConverter.EmuPerPoint) };
        var sf = new Drawing.SolidFill();
        sf.AppendChild(BuildChartColorElement(color));
        outline.AppendChild(sf);
        outline.AppendChild(new Drawing.PresetDash { Val = ParseDashStyle(dash) });
        spPr.AppendChild(outline);
        refSer.AppendChild(spPr);

        // No marker
        refSer.AppendChild(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.None }));

        // Flat data — same value repeated
        var numLitRef = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)refValues.Length });
        for (int i = 0; i < refValues.Length; i++)
            numLitRef.AppendChild(new C.NumericPoint(
                new C.NumericValue(refValue.ToString("G"))) { Index = (uint)i });
        refSer.AppendChild(new C.Values(numLitRef));

        // Insert ser before dLbls/dropLines/hiLowLines/upDownBars/marker/smooth/axId
        // per CT_LineChart schema: grouping, varyColors, ser*, dLbls?, ...
        var insertBeforeEl = lineChart.GetFirstChild<C.DataLabels>() as OpenXmlElement
            ?? lineChart.GetFirstChild<C.DropLines>()
            ?? lineChart.GetFirstChild<C.HighLowLines>()
            ?? lineChart.GetFirstChild<C.UpDownBars>()
            ?? lineChart.GetFirstChild<C.ShowMarker>()
            ?? lineChart.GetFirstChild<C.Smooth>()
            ?? (OpenXmlElement?)lineChart.GetFirstChild<C.AxisId>();
        if (insertBeforeEl != null)
            lineChart.InsertBefore(refSer, insertBeforeEl);
        else
            lineChart.AppendChild(refSer);
    }

    /// <summary>
    /// Remove existing reference line series from a plot area.
    /// A reference line series is identified as a LineChartSeries in a LineChart
    /// where all data points have the same value (flat line), the series has a dashed
    /// outline style, and the marker is set to None.
    /// </summary>
    internal static void RemoveExistingReferenceLines(C.PlotArea plotArea)
    {
        var lineChart = plotArea.GetFirstChild<C.LineChart>();
        if (lineChart == null) return;

        var toRemove = new List<C.LineChartSeries>();
        foreach (var ser in lineChart.Elements<C.LineChartSeries>())
        {
            // Check for reference line markers: no marker (None) and dashed outline
            var marker = ser.GetFirstChild<C.Marker>();
            var markerSymbol = marker?.GetFirstChild<C.Symbol>()?.Val?.Value;
            if (markerSymbol != C.MarkerStyleValues.None) continue;

            var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
            var outline = spPr?.GetFirstChild<Drawing.Outline>();
            var hasDash = outline?.GetFirstChild<Drawing.PresetDash>() != null;
            if (!hasDash) continue;

            // Check if all values are the same (flat line = reference line)
            var vals = ser.GetFirstChild<C.Values>();
            var numLit = vals?.GetFirstChild<C.NumberLiteral>();
            if (numLit != null)
            {
                var points = numLit.Elements<C.NumericPoint>().Select(p => p.InnerText).Distinct().ToList();
                if (points.Count == 1)
                    toRemove.Add(ser);
            }
        }

        foreach (var ser in toRemove)
            ser.Remove();

        // If the LineChart is now empty (no series left), remove it entirely
        if (!lineChart.Elements<C.LineChartSeries>().Any())
            lineChart.Remove();
    }

    /// <summary>
    /// Returns true if any chart in the plot area uses percent-stacked grouping.
    /// BarChart/Bar3DChart use BarGrouping; LineChart/AreaChart use Grouping.
    /// </summary>
    private static bool IsPercentStackedChart(C.PlotArea plotArea)
    {
        foreach (var el in plotArea.Elements<OpenXmlCompositeElement>())
        {
            var barGrouping = el.GetFirstChild<C.BarGrouping>()?.Val?.Value;
            if (barGrouping == C.BarGroupingValues.PercentStacked) return true;

            var grouping = el.GetFirstChild<C.Grouping>()?.Val?.Value;
            if (grouping == C.GroupingValues.PercentStacked) return true;
        }
        return false;
    }

    /// <summary>
    /// Returns true if the given token matches a dash style accepted by ParseDashStyle
    /// (see ChartHelper.Setter.cs). Used for the referenceLine numeric-label heuristic.
    /// </summary>
    private static bool IsKnownDashStyle(string token)
    {
        return token.ToLowerInvariant() switch
        {
            "solid" or "dot" or "sysdot" or "dash" or "sysdash"
                or "dashdot" or "sysdash_dot" or "sysdashdot"
                or "sysdashdotdot" or "sysdash_dot_dot"
                or "longdash" or "lgdash"
                or "longdashdot" or "lgdashdot"
                or "longdashdotdot" or "lgdashdotdot" => true,
            _ => false
        };
    }

    // ==================== Conditional Coloring ====================

    /// <summary>
    /// Apply conditional coloring to data points based on value thresholds.
    /// Format: "threshold:belowColor:aboveColor" or "low:lowColor:mid:midColor:high:highColor"
    /// Simple: "0:FF0000:00AA00" — below 0 = red, above 0 = green
    /// Three-tier: "0:FF0000:50:FFAA00:100:00AA00" — red/orange/green zones
    /// </summary>
    internal static void ApplyColorRule(C.PlotArea plotArea, string spec)
    {
        var parts = spec.Split(':');
        if (parts.Length < 3)
            throw new ArgumentException(
                $"Invalid colorRule '{spec}'. Expected: threshold:belowColor:aboveColor (e.g. '0:FF0000:00AA00') " +
                "or low:lowColor:mid:midColor:high:highColor (e.g. '0:FF0000:50:FFAA00:100:00AA00').");

        var rules = new List<(double threshold, string color)>();
        string topColor;

        if (parts.Length == 3)
        {
            // Simple two-zone: threshold:belowColor:aboveColor
            if (!double.TryParse(parts[0], System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var t))
                throw new ArgumentException($"Invalid threshold '{parts[0]}' in colorRule. Expected a number.");
            rules.Add((t, parts[1].Trim()));
            topColor = parts[2].Trim();
        }
        else
        {
            // Multi-zone: t1:c1:t2:c2:...:cN
            for (int i = 0; i < parts.Length - 1; i += 2)
            {
                if (!double.TryParse(parts[i], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var t))
                    throw new ArgumentException($"Invalid threshold '{parts[i]}' in colorRule.");
                rules.Add((t, parts[i + 1].Trim()));
            }
            topColor = parts.Length % 2 == 1 ? parts[^1].Trim() : rules[^1].color;
            if (parts.Length % 2 == 0)
                rules.RemoveAt(rules.Count - 1); // Last pair has no "above" — use as topColor
        }

        // Apply to each data point in each series
        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
        {
            var values = ReadNumericData(ser.GetFirstChild<C.Values>())
                ?? ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal"));
            if (values == null) continue;

            for (int pi = 0; pi < values.Length; pi++)
            {
                var val = values[pi];
                string pointColor = topColor;
                foreach (var (threshold, color) in rules)
                {
                    if (val < threshold) { pointColor = color; break; }
                    pointColor = color; // at or above this threshold, use this color
                }
                // If above all thresholds, use topColor
                if (rules.Count > 0 && val >= rules[^1].threshold)
                    pointColor = topColor;

                ApplyDataPointColor(ser, pi, pointColor);
            }
        }
    }

    // ==================== Waterfall Chart (Stacked Bar Simulation) ====================

    /// <summary>
    /// Build a waterfall chart using stacked bar technique:
    /// - Invisible "base" series for the running total
    /// - Visible "increase" series (positive changes) and "decrease" series (negative changes)
    /// - Last bar shows the total
    ///
    /// Input: categories and a single series of change values.
    /// e.g. categories=Revenue,Cost,Tax,Profit  data=Cashflow:100,-30,-15,55
    /// The last value can be auto-calculated as the total if "auto" or omitted.
    /// </summary>
    internal static C.ChartSpace BuildWaterfallChart(
        string? title,
        string[]? categories,
        double[] values,
        string? increaseColor,
        string? decreaseColor,
        string? totalColor,
        Dictionary<string, string> properties)
    {
        increaseColor ??= "4472C4"; // blue
        decreaseColor ??= "FF0000"; // red
        totalColor ??= "2E75B6";    // dark blue

        var n = values.Length;
        var baseVals = new double[n];
        var incVals = new double[n];
        var decVals = new double[n];

        double running = 0;
        for (int i = 0; i < n; i++)
        {
            var v = values[i];
            if (i == n - 1 && properties.GetValueOrDefault("waterfallTotal", "true")
                .Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                // Last bar = total (starts from 0, shows cumulative running total)
                // The user's value for the last point is ignored — the total is computed automatically.
                baseVals[i] = 0;
                incVals[i] = running;
                decVals[i] = 0;
            }
            else if (v >= 0)
            {
                baseVals[i] = running;
                incVals[i] = v;
                decVals[i] = 0;
                running += v;
            }
            else
            {
                baseVals[i] = running + v; // base drops by |v|
                incVals[i] = 0;
                decVals[i] = -v;
                running += v;
            }
        }

        categories ??= Enumerable.Range(1, n).Select(i => i.ToString()).ToArray();

        var chartSpace = new C.ChartSpace();
        var chart = new C.Chart();
        if (!string.IsNullOrEmpty(title))
            chart.AppendChild(BuildChartTitle(title));

        var plotArea = new C.PlotArea(new C.Layout());
        uint catAxisId = 1, valAxisId = 2;

        var barChart = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Stacked },
            new C.VaryColors { Val = false }
        );

        // Series 0: invisible base
        var baseSer = BuildBarSeries(0, "Base", categories, baseVals, null);
        // Make base series invisible: no fill, no border
        baseSer.RemoveAllChildren<C.ChartShapeProperties>();
        var baseSpPr = new C.ChartShapeProperties();
        baseSpPr.AppendChild(new Drawing.NoFill());
        var baseOutline = new Drawing.Outline();
        baseOutline.AppendChild(new Drawing.NoFill());
        baseSpPr.AppendChild(baseOutline);
        baseSer.InsertAfter(baseSpPr, baseSer.GetFirstChild<C.SeriesText>());
        barChart.AppendChild(baseSer);

        // Series 1: increase (positive values)
        barChart.AppendChild(BuildBarSeries(1, "Increase", categories, incVals, increaseColor));

        // Series 2: decrease (negative values)
        barChart.AppendChild(BuildBarSeries(2, "Decrease", categories, decVals, decreaseColor));

        barChart.AppendChild(new C.GapWidth { Val = 80 });
        barChart.AppendChild(new C.Overlap { Val = 100 });
        barChart.AppendChild(new C.AxisId { Val = catAxisId });
        barChart.AppendChild(new C.AxisId { Val = valAxisId });

        plotArea.AppendChild(barChart);
        plotArea.AppendChild(BuildCategoryAxis(catAxisId, valAxisId));
        plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));

        chart.AppendChild(plotArea);

        // Hide base series from legend
        var legend = new C.Legend(
            new C.LegendPosition { Val = C.LegendPositionValues.Bottom },
            new C.Overlay { Val = false }
        );
        // Delete legend entry for base series (index 0)
        // CT_Legend schema order: legendPos, legendEntry+, layout, overlay — insert after legendPos
        var leBase = new C.LegendEntry();
        leBase.AppendChild(new C.Index { Val = 0 });
        leBase.AppendChild(new C.Delete { Val = true });
        var legendPosEl = legend.GetFirstChild<C.LegendPosition>();
        if (legendPosEl != null)
            legendPosEl.InsertAfterSelf(leBase);
        else
            legend.PrependChild(leBase);
        chart.AppendChild(legend);

        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chart.AppendChild(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        chartSpace.AppendChild(chart);

        // Color the total bar differently (last data point of increase series)
        if (properties.GetValueOrDefault("waterfallTotal", "true")
            .Equals("true", StringComparison.OrdinalIgnoreCase) && n > 0)
        {
            var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser").ToList();
            if (allSer.Count >= 2)
                ApplyDataPointColor(allSer[1], n - 1, totalColor);
        }

        return chartSpace;
    }

    // ==================== Flexible Combo Chart ====================

    /// <summary>
    /// Build a combo chart with per-series chart type assignment.
    /// comboTypes property: "column,column,line,area" — one type per series.
    /// </summary>
    internal static void RebuildComboChart(C.Chart chart, string comboTypes)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        var typeList = comboTypes.Split(',').Select(t => t.Trim().ToLowerInvariant()).ToArray();

        // Validate every token BEFORE any mutation: unknown tokens used to fall
        // through to the default LineChart arm, silently coercing garbage
        // (combotypes=asdf,qwer) into line,line — the only mini-language prop
        // that accepted typos. Also keeps the rebuild atomic on bad input.
        foreach (var t in typeList)
        {
            var baseToken = t.EndsWith("percentstacked", StringComparison.Ordinal) ? t[..^14]
                : t.EndsWith("stacked", StringComparison.Ordinal) ? t[..^7]
                : t;
            if (baseToken is not ("bar" or "column" or "col" or "line" or "area" or "scatter"))
                throw new ArgumentException(
                    $"Invalid comboTypes token '{t}'. Expected bar/column/line/area/scatter, " +
                    "optionally with a stacked/percentstacked suffix (e.g. 'column,line' or 'columnstacked,line').");
        }

        // Read all existing series data
        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        if (allSer.Count == 0) return;

        // Read series data
        var seriesInfo = new List<(OpenXmlCompositeElement original, string targetType)>();
        for (int i = 0; i < allSer.Count; i++)
        {
            var targetType = i < typeList.Length ? typeList[i] : typeList[^1];
            seriesInfo.Add((allSer[i], targetType));
        }

        // Find axis IDs
        uint catAxisId = plotArea.GetFirstChild<C.CategoryAxis>()?.GetFirstChild<C.AxisId>()?.Val?.Value ?? 1;
        uint valAxisId = plotArea.Elements<C.ValueAxis>()
            .FirstOrDefault(axis => axis.GetFirstChild<C.AxisId>()?.Val?.Value != catAxisId)
            ?.GetFirstChild<C.AxisId>()?.Val?.Value ?? 2;

        // Remove existing chart type elements (but keep axes, layout, etc.)
        foreach (var ct in plotArea.ChildElements
            .Where(e => e.LocalName.EndsWith("Chart") || e.LocalName.EndsWith("chart"))
            .OfType<OpenXmlCompositeElement>().ToList())
        {
            ct.Remove();
        }

        // R26-1 — drop any SECONDARY axis declarations (catAx/valAx with an id
        // other than the primary cat/val ids, i.e. the 3/4 pair created by a
        // prior ApplySecondaryAxis). The rebuild below re-binds every series to
        // the primary axIds, so a leftover secondary axis would be referenced by
        // no chart container and Excel rejects the orphaned declaration. This is
        // order-independent defense: it holds even if secondaryaxis somehow runs
        // before combotypes (the PropOrder=0 schedule normally prevents that).
        foreach (var ax in plotArea.ChildElements
            .Where(e => e.LocalName is "catAx" or "valAx" or "serAx" or "dateAx")
            .OfType<OpenXmlCompositeElement>().ToList())
        {
            var id = ax.GetFirstChild<C.AxisId>()?.Val?.Value;
            if (id.HasValue && id.Value != catAxisId && id.Value != valAxisId)
                ax.Remove();
        }

        // Group series by target chart type
        var groups = seriesInfo.GroupBy(s => s.targetType).ToList();
        foreach (var group in groups)
        {
            // Grouping-qualified tokens (columnstacked / areapercentstacked …)
            // — parse the suffix so a stacked combo group doesn't rebuild as
            // clustered/standard (and doesn't fall through to the default
            // LineChart branch).
            var groupToken = group.Key;
            string comboGrpSuffix = "";
            if (groupToken.EndsWith("percentstacked", StringComparison.Ordinal))
            { comboGrpSuffix = "percentstacked"; groupToken = groupToken[..^14]; }
            else if (groupToken.EndsWith("stacked", StringComparison.Ordinal))
            { comboGrpSuffix = "stacked"; groupToken = groupToken[..^7]; }
            var comboBarGrp = comboGrpSuffix switch
            {
                "percentstacked" => C.BarGroupingValues.PercentStacked,
                "stacked" => C.BarGroupingValues.Stacked,
                _ => C.BarGroupingValues.Clustered,
            };
            var comboStdGrp = comboGrpSuffix switch
            {
                "percentstacked" => C.GroupingValues.PercentStacked,
                "stacked" => C.GroupingValues.Stacked,
                _ => C.GroupingValues.Standard,
            };
            OpenXmlCompositeElement chartTypeEl;
            switch (groupToken)
            {
                case "bar":
                    chartTypeEl = new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Bar },
                        new C.BarGrouping { Val = comboBarGrp },
                        new C.VaryColors { Val = false });
                    break;
                case "column" or "col":
                    chartTypeEl = new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Column },
                        new C.BarGrouping { Val = comboBarGrp },
                        new C.VaryColors { Val = false });
                    break;
                case "line":
                    chartTypeEl = new C.LineChart(
                        new C.Grouping { Val = comboStdGrp },
                        new C.VaryColors { Val = false });
                    break;
                case "area":
                    chartTypeEl = new C.AreaChart(
                        new C.Grouping { Val = comboStdGrp },
                        new C.VaryColors { Val = false });
                    break;
                case "scatter":
                    chartTypeEl = new C.ScatterChart(
                        new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                        new C.VaryColors { Val = false });
                    break;
                default:
                    chartTypeEl = new C.LineChart(
                        new C.Grouping { Val = comboStdGrp },
                        new C.VaryColors { Val = false });
                    break;
            }

            foreach (var (original, _) in group)
            {
                // Don't clone original directly — original is a BarChartSeries, but
                // chartTypeEl may be LineChart/AreaChart/ScatterChart which require
                // LineChartSeries / AreaChartSeries / ScatterChartSeries respectively.
                // Schema validation rejects mismatched series. Convert to the right type.
                chartTypeEl.AppendChild(ConvertSeriesToType(original, groupToken));
            }

            // Bar/column groups get the same explicit gapWidth the builder
            // stamps (150, the spec default). Omitting it renders identically
            // today but made a combotypes-rebuilt chart differ from a
            // directly-built one, breaking first-round dump idempotency.
            if (chartTypeEl is C.BarChart)
                chartTypeEl.AppendChild(new C.GapWidth { Val = 150 });
            chartTypeEl.AppendChild(new C.AxisId { Val = catAxisId });
            chartTypeEl.AppendChild(new C.AxisId { Val = valAxisId });

            // Insert before axes
            var firstAxis = plotArea.Elements<C.CategoryAxis>().FirstOrDefault() as OpenXmlElement
                ?? plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (firstAxis != null)
                plotArea.InsertBefore(chartTypeEl, firstAxis);
            else
                plotArea.AppendChild(chartTypeEl);
        }
    }

    /// <summary>
    /// Convert a chart series element (BarChartSeries, LineChartSeries, etc.) to the
    /// series type required by a target chart type (bar/column/line/area/scatter).
    /// The OOXML schema requires each chart container to host its own series subclass —
    /// a LineChart cannot host a BarChartSeries even though the field set is identical.
    /// Copies idx, order, tx, spPr, cat, val (and x/yVal for scatter) from the source.
    /// </summary>
    private static OpenXmlCompositeElement ConvertSeriesToType(OpenXmlCompositeElement source, string targetType)
    {
        // Extract identity + data children by local name so we can move them across
        // schema namespaces without depending on the source's concrete type.
        OpenXmlElement? Take(string localName)
        {
            return source.ChildElements.FirstOrDefault(e => e.LocalName == localName);
        }

        var idx = Take("idx");
        var order = Take("order");
        var tx = Take("tx");
        var spPr = Take("spPr");
        var marker = Take("marker");
        var cat = Take("cat");
        var val = Take("val");
        var xVal = Take("xVal");
        var yVal = Take("yVal");
        var smooth = Take("smooth");
        var invertIfNegative = Take("invertIfNegative");

        OpenXmlCompositeElement target = targetType switch
        {
            "bar" or "column" or "col" => new C.BarChartSeries(),
            "line" => new C.LineChartSeries(),
            "area" => new C.AreaChartSeries(),
            "scatter" => new C.ScatterChartSeries(),
            _ => new C.LineChartSeries(),
        };

        // CT_SerXxx schema order: idx, order, tx, spPr, [invertIfNegative|marker], [dPt*],
        // [dLbls], [trendline*], [errBars], cat, val (or xVal/yVal for scatter), smooth.
        if (idx != null) target.AppendChild(idx.CloneNode(true));
        if (order != null) target.AppendChild(order.CloneNode(true));
        if (tx != null) target.AppendChild(tx.CloneNode(true));
        if (spPr != null)
        {
            var spPrClone = (OpenXmlCompositeElement)spPr.CloneNode(true);
            // Line-based series carry their color on the stroke
            // (<a:ln><a:solidFill>); a bare <a:solidFill> cloned from an
            // area/bar source is stroke-inert (real Office ignores it and
            // renders the theme color), AND the dump reads it as the series
            // color and replays it as an <a:ln> stroke — a first-round
            // dump→replay drift. Rewrap the fill as the stroke here.
            if (targetType is "line" or "scatter")
            {
                var bareFill = spPrClone.ChildElements
                    .FirstOrDefault(e => e.LocalName == "solidFill");
                var hasLn = spPrClone.ChildElements.Any(e => e.LocalName == "ln");
                if (bareFill != null && !hasLn)
                {
                    bareFill.Remove();
                    var outline = new Drawing.Outline { Width = 25400 }; // 2pt, same as ApplySeriesColor
                    outline.AppendChild(bareFill);
                    // CT_ShapeProperties order puts <a:ln> after the fill
                    // group; the fill was just removed, so insert before any
                    // effect/3d/ext tail else append.
                    var lnBefore = spPrClone.ChildElements.FirstOrDefault(e =>
                        e.LocalName is "effectLst" or "effectDag" or "scene3d" or "sp3d" or "extLst");
                    if (lnBefore != null) spPrClone.InsertBefore(outline, lnBefore);
                    else spPrClone.AppendChild(outline);
                }
            }
            target.AppendChild(spPrClone);
        }

        // invertIfNegative only valid on bar series; marker on line/scatter
        if (targetType is "bar" or "column" or "col")
        {
            if (invertIfNegative != null) target.AppendChild(invertIfNegative.CloneNode(true));
        }
        else if (targetType is "line" or "scatter")
        {
            if (marker != null) target.AppendChild(marker.CloneNode(true));
        }

        if (targetType == "scatter")
        {
            // Scatter needs xVal + yVal; synthesize from cat/val if source was non-scatter.
            if (xVal != null)
                target.AppendChild(xVal.CloneNode(true));
            else if (cat != null)
            {
                // Reuse cat data as numeric x-values where possible; otherwise omit.
                // Scatter without xVal is legal — Excel auto-indexes.
            }
            if (yVal != null) target.AppendChild(yVal.CloneNode(true));
            else if (val != null)
            {
                // Convert c:val -> c:yVal by re-parenting the numRef/numLit child.
                var inner = val.ChildElements.FirstOrDefault(e =>
                    e.LocalName == "numRef" || e.LocalName == "numLit");
                if (inner != null)
                    target.AppendChild(new C.YValues(inner.CloneNode(true)));
            }
        }
        else
        {
            if (cat != null) target.AppendChild(cat.CloneNode(true));
            if (val != null) target.AppendChild(val.CloneNode(true));
        }

        if (targetType is "line" or "scatter" && smooth != null)
            target.AppendChild(smooth.CloneNode(true));

        return target;
    }
}
