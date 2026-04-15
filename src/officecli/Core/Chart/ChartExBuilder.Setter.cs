// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeCli.Core;

/// <summary>
/// Set-side (mutate-in-place) implementation for cx:chart extended chart
/// types. Covers the same vocabulary as the Add path in ChartExBuilder.cs
/// so charts created via Add can be fully re-styled via Set.
///
/// The shape of each case mirrors ChartHelper.Setter.cs for regular cChart:
/// remove the existing styled element, rebuild it via a shared helper (or
/// mutate in place), and save. All tree mutations respect the CT_Axis /
/// CT_Chart schema order.
/// </summary>
internal static partial class ChartExBuilder
{
    /// <summary>
    /// Mutate an existing <see cref="ExtendedChartPart"/> to apply the given
    /// properties. Returns the list of keys that weren't recognized (caller
    /// surfaces these to the user). Unknown keys are never an error — same
    /// convention as ChartHelper.SetChartProperties.
    /// </summary>
    internal static List<string> SetChartProperties(
        ExtendedChartPart chartPart, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var chartSpace = chartPart.ChartSpace;
        var chart = chartSpace?.GetFirstChild<CX.Chart>();
        if (chart == null) { unsupported.AddRange(properties.Keys); return unsupported; }

        var plotArea = chart.GetFirstChild<CX.PlotArea>();
        var plotAreaRegion = plotArea?.GetFirstChild<CX.PlotAreaRegion>();
        var allSeries = plotAreaRegion?.Elements<CX.Series>().ToList() ?? new List<CX.Series>();
        var allAxes = plotArea?.Elements<CX.Axis>().ToList() ?? new List<CX.Axis>();
        var catAxis = allAxes.FirstOrDefault();          // Id=0 — category axis (histogram/boxWhisker)
        var valAxis = allAxes.ElementAtOrDefault(1);      // Id=1 — value axis

        // Process structural properties (title text, axis title creation) before
        // styling properties (title.color, axisTitle.color) so the target element
        // always exists by the time the styling case runs. Same trick as the
        // regular cChart setter.
        static int PropOrder(string k)
        {
            var lower = k.ToLowerInvariant();
            if (lower is "title" or "xaxistitle" or "yaxistitle" or "legend") return 0;
            return 1;
        }

        foreach (var (key, value) in properties.OrderBy(kv => PropOrder(kv.Key)))
        {
            var handled = HandleSetKey(chart, plotArea, allSeries, allAxes, catAxis, valAxis,
                key, value, properties);
            if (!handled) unsupported.Add(key);
        }

        chartPart.ChartSpace?.Save();
        return unsupported;
    }

    // The per-key dispatch lives in its own method so the surrounding loop
    // stays readable. Returns true if the key was recognized (regardless of
    // whether anything could actually be mutated — e.g. styling a non-existent
    // title is a silent no-op, not an unsupported-key report, matching regular
    // cChart semantics).
    private static bool HandleSetKey(
        CX.Chart chart,
        CX.PlotArea? plotArea,
        List<CX.Series> allSeries,
        List<CX.Axis> allAxes,
        CX.Axis? catAxis,
        CX.Axis? valAxis,
        string key,
        string value,
        Dictionary<string, string> allProperties)
    {
        switch (key.ToLowerInvariant())
        {
            // ==================== Chart title ====================

            case "title":
            {
                chart.RemoveAllChildren<CX.ChartTitle>();
                if (!string.IsNullOrEmpty(value)
                    && !value.Equals("none", StringComparison.OrdinalIgnoreCase)
                    && !value.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    // cx:title must be the first child of cx:chart per schema.
                    chart.PrependChild(BuildChartTitle(value, allProperties));
                }
                return true;
            }

            case "title.color" or "titlecolor":
            case "title.size" or "titlesize":
            case "title.font" or "titlefont":
            case "title.bold" or "titlebold":
            {
                var ctitle = chart.GetFirstChild<CX.ChartTitle>();
                if (ctitle == null) return true; // silent no-op
                foreach (var run in ctitle.Descendants<Drawing.Run>())
                {
                    var rPr = run.RunProperties
                              ?? (run.RunProperties = new Drawing.RunProperties { Language = "en-US" });
                    ChartHelper.ApplyRunStyleProperties(rPr, allProperties, keyPrefix: "title");
                }
                return true;
            }

            case "title.shadow" or "titleshadow":
            {
                // Apply an a:outerShdw effect to the title run's rPr. Same
                // vocabulary as regular cChart (ChartHelper.Setter.cs:63):
                // "COLOR-BLUR-ANGLE-DIST-OPACITY" or "none" to clear.
                var ctitle = chart.GetFirstChild<CX.ChartTitle>();
                if (ctitle == null) return true;
                foreach (var run in ctitle.Descendants<Drawing.Run>())
                {
                    var rPr = run.RunProperties
                              ?? (run.RunProperties = new Drawing.RunProperties { Language = "en-US" });
                    ApplyRunEffectShadow(rPr, value);
                }
                return true;
            }

            // ==================== Legend ====================

            case "legend":
            {
                chart.RemoveAllChildren<CX.Legend>();
                if (!string.IsNullOrEmpty(value)
                    && !value.Equals("none", StringComparison.OrdinalIgnoreCase)
                    && !value.Equals("false", StringComparison.OrdinalIgnoreCase)
                    && !value.Equals("off", StringComparison.OrdinalIgnoreCase))
                {
                    // Legend goes after plotArea per cx:chart schema.
                    chart.AppendChild(BuildLegend(value, allProperties));
                }
                return true;
            }

            case "legend.overlay" or "legendoverlay":
            {
                var legend = chart.GetFirstChild<CX.Legend>();
                if (legend == null) return true;
                legend.Overlay = ParseHelpers.IsTruthy(value);
                return true;
            }

            case "legendfont" or "legend.font":
            {
                // Compound form "size:color:fontname" styles the legend text.
                // Mirrors ChartHelper.Setter.cs:118 "legendfont" for regular
                // cChart. Wraps an a:defRPr in cx:txPr on the legend.
                var legend = chart.GetFirstChild<CX.Legend>();
                if (legend == null) return true;
                legend.RemoveAllChildren<CX.TxPrTextBody>();
                if (!string.IsNullOrEmpty(value)
                    && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    var txPr = BuildAxisTickLabelStyle(value);
                    if (txPr != null) legend.AppendChild(txPr);
                }
                return true;
            }

            // ==================== Axis titles (text) ====================

            case "xaxistitle":
            {
                if (catAxis == null) return true;
                catAxis.RemoveAllChildren<CX.AxisTitle>();
                if (!string.IsNullOrEmpty(value)
                    && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    InsertAxisTitle(catAxis, BuildAxisTitle(value, allProperties));
                }
                return true;
            }

            case "yaxistitle":
            {
                if (valAxis == null) return true;
                valAxis.RemoveAllChildren<CX.AxisTitle>();
                if (!string.IsNullOrEmpty(value)
                    && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    InsertAxisTitle(valAxis, BuildAxisTitle(value, allProperties));
                }
                return true;
            }

            case "axistitle.color" or "axistitlecolor":
            case "axistitle.size" or "axistitlesize":
            case "axistitle.font" or "axistitlefont":
            case "axistitle.bold" or "axistitlebold":
            {
                foreach (var axis in allAxes)
                {
                    var axisTitle = axis.GetFirstChild<CX.AxisTitle>();
                    if (axisTitle == null) continue;
                    foreach (var run in axisTitle.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties
                                  ?? (run.RunProperties = new Drawing.RunProperties { Language = "en-US" });
                        ChartHelper.ApplyRunStyleProperties(rPr, allProperties, keyPrefix: "axisTitle");
                    }
                }
                return true;
            }

            // ==================== Tick-label font (axis-level cx:txPr) ====================

            case "axisfont" or "axis.font":
            {
                foreach (var axis in allAxes)
                {
                    // cx:txPr must remain the last axis child (per CT_Axis schema:
                    // ... → tickLabels → numFmt → spPr → txPr → extLst).
                    axis.RemoveAllChildren<CX.TxPrTextBody>();
                    var txPr = BuildAxisTickLabelStyle(value);
                    if (txPr != null) axis.AppendChild(txPr);
                }
                return true;
            }

            // ==================== Gridlines ====================

            case "gridlines":
            {
                if (valAxis == null) return true;
                valAxis.RemoveAllChildren<CX.MajorGridlinesGridlines>();
                if (ParseHelpers.IsTruthy(value))
                    InsertGridlinesInAxisOrder(valAxis, new CX.MajorGridlinesGridlines());
                return true;
            }

            case "xgridlines":
            {
                if (catAxis == null) return true;
                catAxis.RemoveAllChildren<CX.MajorGridlinesGridlines>();
                if (ParseHelpers.IsTruthy(value))
                    InsertGridlinesInAxisOrder(catAxis, new CX.MajorGridlinesGridlines());
                return true;
            }

            case "gridlinecolor" or "gridline.color":
            {
                var gl = valAxis?.GetFirstChild<CX.MajorGridlinesGridlines>();
                if (gl != null) gl.ShapeProperties = BuildGridlineShapeProperties(value);
                return true;
            }

            case "xgridlinecolor" or "xgridline.color":
            {
                var gl = catAxis?.GetFirstChild<CX.MajorGridlinesGridlines>();
                if (gl != null) gl.ShapeProperties = BuildGridlineShapeProperties(value);
                return true;
            }

            // ==================== Value-axis scaling (axismin/max/majorunit) ====================
            // CONSISTENCY(chart-axis-scaling): same prop names as regular cChart
            // (ChartHelper.Setter.cs:357). CX.ValueAxisScaling stores Min/Max/
            // MajorUnit/MinorUnit as StringValue attributes, not typed doubles,
            // but we still parse + re-format as invariant double for
            // consistency with cChart behavior (reject NaN/Infinity).
            case "axismin" or "min":
            {
                var valScaling = valAxis?.GetFirstChild<CX.ValueAxisScaling>();
                if (valScaling == null) return true;
                valScaling.Min = ParseHelpers.SafeParseDouble(value, "axismin")
                    .ToString("G", CultureInfo.InvariantCulture);
                return true;
            }

            case "axismax" or "max":
            {
                var valScaling = valAxis?.GetFirstChild<CX.ValueAxisScaling>();
                if (valScaling == null) return true;
                valScaling.Max = ParseHelpers.SafeParseDouble(value, "axismax")
                    .ToString("G", CultureInfo.InvariantCulture);
                return true;
            }

            case "majorunit":
            {
                var valScaling = valAxis?.GetFirstChild<CX.ValueAxisScaling>();
                if (valScaling == null) return true;
                valScaling.MajorUnit = ParseHelpers.SafeParseDouble(value, "majorunit")
                    .ToString("G", CultureInfo.InvariantCulture);
                return true;
            }

            case "minorunit":
            {
                var valScaling = valAxis?.GetFirstChild<CX.ValueAxisScaling>();
                if (valScaling == null) return true;
                valScaling.MinorUnit = ParseHelpers.SafeParseDouble(value, "minorunit")
                    .ToString("G", CultureInfo.InvariantCulture);
                return true;
            }

            // ==================== Axis visibility (hidden flag) ====================
            // CONSISTENCY(chart-axis-visibility): same prop names as regular
            // cChart (ChartHelper.Setter.cs:795). CX uses a simple @hidden
            // attribute on cx:axis, unlike cChart's c:delete child element.
            case "axisvisible" or "axis.visible" or "axis.delete":
            {
                var hide = key.Contains("delete")
                    ? ParseHelpers.IsTruthy(value)
                    : !ParseHelpers.IsTruthy(value);
                foreach (var axis in allAxes) axis.Hidden = hide;
                return true;
            }

            case "cataxisvisible" or "cataxis.visible":
            {
                if (catAxis != null) catAxis.Hidden = !ParseHelpers.IsTruthy(value);
                return true;
            }

            case "valaxisvisible" or "valaxis.visible":
            {
                if (valAxis != null) valAxis.Hidden = !ParseHelpers.IsTruthy(value);
                return true;
            }

            // ==================== Axis line styling ====================
            // CONSISTENCY(chart-axis-line): "color" | "color:width" | "color:width:dash"
            // | "none". Same vocabulary as regular cChart (ChartHelper.Setter.cs:1471),
            // reuses ChartHelper.BuildOutlineElement for parsing.
            case "axisline" or "axis.line":
            {
                foreach (var axis in allAxes) ApplyCxAxisLine(axis, value);
                return true;
            }

            case "cataxisline" or "cataxis.line":
            {
                if (catAxis != null) ApplyCxAxisLine(catAxis, value);
                return true;
            }

            case "valaxisline" or "valaxis.line":
            {
                if (valAxis != null) ApplyCxAxisLine(valAxis, value);
                return true;
            }

            // ==================== Tick labels (on/off, both axes) ====================

            case "ticklabels":
            {
                var enable = ParseHelpers.IsTruthy(value);
                foreach (var axis in allAxes)
                {
                    axis.RemoveAllChildren<CX.TickLabels>();
                    if (enable) InsertTickLabelsInAxisOrder(axis, new CX.TickLabels());
                }
                return true;
            }

            // ==================== Data labels (series-level) ====================

            case "datalabels" or "labels":
            {
                var enable = ParseHelpers.IsTruthy(value);
                foreach (var series in allSeries)
                {
                    series.RemoveAllChildren<CX.DataLabels>();
                    if (!enable) continue;
                    var dl = new CX.DataLabels { Pos = CX.DataLabelPos.OutEnd };
                    dl.AppendChild(new CX.DataLabelVisibilities
                    {
                        Value = true, SeriesName = false, CategoryName = false,
                    });
                    // dataLabels goes before cx:dataId per cx:series schema.
                    var dataId = series.GetFirstChild<CX.DataId>();
                    if (dataId != null) series.InsertBefore(dl, dataId);
                    else series.AppendChild(dl);
                }
                return true;
            }

            case "datalabels.numfmt" or "labelnumfmt" or "datalabels.format" or "labelformat":
            {
                // CONSISTENCY(chart-datalabel-numfmt): same prop names as
                // regular cChart (ChartHelper.Setter.cs:1181). Applies a
                // cx:numFmt element to every series' cx:dataLabels. Silent
                // no-op if a series has no dataLabels block (use `dataLabels=true`
                // to enable them first, same as regular cChart semantics).
                foreach (var series in allSeries)
                {
                    var dl = series.GetFirstChild<CX.DataLabels>();
                    if (dl == null) continue;
                    dl.NumberFormat = new CX.NumberFormat
                    {
                        FormatCode = value,
                        SourceLinked = false,
                    };
                }
                return true;
            }

            // ==================== Series fill / multi-series colors ====================

            case "fill":
            {
                foreach (var series in allSeries)
                    ReplaceSeriesFill(series, value);
                return true;
            }

            case "colors":
            {
                var colorList = value.Split(',').Select(c => c.Trim()).ToArray();
                for (int i = 0; i < Math.Min(allSeries.Count, colorList.Length); i++)
                    ReplaceSeriesFill(allSeries[i], colorList[i]);
                return true;
            }

            // ==================== Series effects (shadow) ====================
            // CONSISTENCY(chart-series-shadow): same vocabulary as regular cChart
            // (ChartHelper.Setter.cs:642 / SetterHelpers.cs:374). Format
            // "COLOR-BLUR-ANGLE-DIST-OPACITY" or "none" to clear. Applied to
            // every series by attaching an a:effectLst inside the existing
            // cx:spPr (or creating one if the series has no fill yet).
            case "series.shadow" or "seriesshadow":
            {
                foreach (var series in allSeries)
                    ApplyCxSeriesShadow(series, value);
                return true;
            }

            // ==================== Histogram binning ====================

            case "bincount":
            {
                SetHistogramBinSpec(allSeries, kind: "binCount", rawValue: value);
                return true;
            }

            case "binsize":
            {
                SetHistogramBinSpec(allSeries, kind: "binSize", rawValue: value);
                return true;
            }

            case "intervalclosed":
            {
                foreach (var series in allSeries)
                {
                    var binning = series.Descendants<CX.Binning>().FirstOrDefault();
                    if (binning == null) continue;
                    binning.IntervalClosed = value.ToLowerInvariant() == "l"
                        ? CX.IntervalClosedSide.L
                        : CX.IntervalClosedSide.R;
                }
                return true;
            }

            case "underflowbin":
            {
                foreach (var series in allSeries)
                {
                    var binning = series.Descendants<CX.Binning>().FirstOrDefault();
                    if (binning != null)
                        binning.Underflow = string.IsNullOrEmpty(value) ? null : value;
                }
                return true;
            }

            case "overflowbin":
            {
                foreach (var series in allSeries)
                {
                    var binning = series.Descendants<CX.Binning>().FirstOrDefault();
                    if (binning != null)
                        binning.Overflow = string.IsNullOrEmpty(value) ? null : value;
                }
                return true;
            }

            case "gapwidth":
            {
                var catScaling = catAxis?.GetFirstChild<CX.CategoryAxisScaling>();
                if (catScaling != null) catScaling.GapWidth = value;
                return true;
            }

            // ==================== Other extended-type layoutPr ====================

            case "parentlabellayout":  // treemap
            {
                foreach (var series in allSeries)
                {
                    var parentLabel = series.Descendants<CX.ParentLabelLayout>().FirstOrDefault();
                    if (parentLabel == null) continue;
                    parentLabel.ParentLabelLayoutVal = value.ToLowerInvariant() switch
                    {
                        "none" => CX.ParentLabelLayoutVal.None,
                        "banner" => CX.ParentLabelLayoutVal.Banner,
                        _ => CX.ParentLabelLayoutVal.Overlapping,
                    };
                }
                return true;
            }

            case "quartilemethod":  // boxwhisker
            {
                foreach (var series in allSeries)
                {
                    var stats = series.Descendants<CX.Statistics>().FirstOrDefault();
                    if (stats == null) continue;
                    stats.QuartileMethod = value.ToLowerInvariant() == "inclusive"
                        ? CX.QuartileMethod.Inclusive
                        : CX.QuartileMethod.Exclusive;
                }
                return true;
            }

            // ==================== Plot area / chart area fill + border ====================
            // CONSISTENCY(chart-area-fill): same prop names as regular cChart
            // (ChartHelper.Setter.cs:476,491,1220,1232). Both PlotArea and
            // ChartSpace accept a cx:spPr child; we attach a solidFill for
            // the background and an a:ln outline for the border.
            case "plotareafill" or "plotfill":
            {
                if (plotArea == null) return true;
                ApplyCxAreaFill(plotArea, value);
                return true;
            }

            case "plotarea.border" or "plotborder":
            {
                if (plotArea == null) return true;
                ApplyCxAreaBorder(plotArea, value);
                return true;
            }

            case "chartareafill" or "chartfill":
            {
                var chartSpace = chart.Parent as CX.ChartSpace;
                if (chartSpace == null) return true;
                ApplyCxAreaFill(chartSpace, value);
                return true;
            }

            case "chartarea.border" or "chartborder":
            {
                var chartSpace = chart.Parent as CX.ChartSpace;
                if (chartSpace == null) return true;
                ApplyCxAreaBorder(chartSpace, value);
                return true;
            }
        }
        return false;
    }

    // ==================== Schema-aware insertion helpers ====================

    /// <summary>
    /// Insert a <see cref="CX.AxisTitle"/> into an axis, respecting the
    /// CT_Axis sequence: catScaling/valScaling → title → units → gridlines → ...
    /// </summary>
    private static void InsertAxisTitle(CX.Axis axis, CX.AxisTitle title)
    {
        // Title goes immediately after catScaling/valScaling.
        var scaling = axis.GetFirstChild<CX.CategoryAxisScaling>() as OpenXmlElement
                   ?? axis.GetFirstChild<CX.ValueAxisScaling>();
        if (scaling != null) scaling.InsertAfterSelf(title);
        else axis.PrependChild(title);
    }

    /// <summary>
    /// Insert majorGridlines after title (or scaling) but before tickLabels /
    /// spPr / txPr, matching the CT_Axis schema sequence.
    /// </summary>
    private static void InsertGridlinesInAxisOrder(CX.Axis axis, CX.MajorGridlinesGridlines gl)
    {
        var insertAfter = (OpenXmlElement?)axis.GetFirstChild<CX.AxisTitle>()
                       ?? (OpenXmlElement?)axis.GetFirstChild<CX.CategoryAxisScaling>()
                       ?? axis.GetFirstChild<CX.ValueAxisScaling>();
        if (insertAfter != null) insertAfter.InsertAfterSelf(gl);
        else axis.PrependChild(gl);
    }

    /// <summary>
    /// Insert tickLabels after gridlines (or earlier children) but before
    /// axis-level spPr / txPr.
    /// </summary>
    private static void InsertTickLabelsInAxisOrder(CX.Axis axis, CX.TickLabels tickLabels)
    {
        // cx:txPr is what our Set path appends to the axis for tick-label
        // styling; tickLabels must come BEFORE any existing txPr.
        var existingTxPr = axis.GetFirstChild<CX.TxPrTextBody>();
        if (existingTxPr != null)
        {
            axis.InsertBefore(tickLabels, existingTxPr);
            return;
        }
        var insertAfter = (OpenXmlElement?)axis.GetFirstChild<CX.MajorGridlinesGridlines>()
                       ?? (OpenXmlElement?)axis.GetFirstChild<CX.AxisTitle>()
                       ?? (OpenXmlElement?)axis.GetFirstChild<CX.CategoryAxisScaling>()
                       ?? axis.GetFirstChild<CX.ValueAxisScaling>();
        if (insertAfter != null) insertAfter.InsertAfterSelf(tickLabels);
        else axis.AppendChild(tickLabels);
    }

    // ==================== Series-level helpers ====================

    /// <summary>
    /// Replace the series fill color (single solid fill). Used by both
    /// `fill` and `colors` cases.
    /// </summary>
    private static void ReplaceSeriesFill(CX.Series series, string color)
    {
        if (string.IsNullOrEmpty(color)) return;
        series.RemoveAllChildren<CX.ShapeProperties>();
        var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(color);
        var spPr = new CX.ShapeProperties(
            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb }));
        // spPr goes right after cx:tx per cx:series schema sequence.
        var tx = series.GetFirstChild<CX.Text>();
        if (tx != null) tx.InsertAfterSelf(spPr);
        else series.PrependChild(spPr);
    }

    /// <summary>
    /// Replace a histogram's <c>cx:binCount</c> / <c>cx:binSize</c> with the
    /// given value. Binning is XOR — setting one removes the other. Uses the
    /// same OpenXmlUnknownElement workaround as the Add path (SDK's typed
    /// binCount is a leaf-text element but Excel wants a <c>val</c> attribute).
    /// </summary>
    private static void SetHistogramBinSpec(
        IReadOnlyList<CX.Series> allSeries, string kind, string rawValue)
    {
        const string cxNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";

        foreach (var series in allSeries)
        {
            var lp = series.GetFirstChild<CX.SeriesLayoutProperties>();
            if (lp == null) continue;
            var binning = lp.GetFirstChild<CX.Binning>();
            if (binning == null) continue;

            // Remove any existing binCount / binSize (XOR with the new one).
            foreach (var existing in binning.ChildElements.ToList())
                if (existing.LocalName is "binCount" or "binSize") existing.Remove();

            if (string.IsNullOrEmpty(rawValue)) continue; // bare "bincount=" clears

            if (kind == "binCount" && uint.TryParse(rawValue, out var binCount))
            {
                var el = new OpenXmlUnknownElement("cx", "binCount", cxNs);
                el.SetAttribute(new OpenXmlAttribute("val", "", binCount.ToString()));
                binning.AppendChild(el);
            }
            else if (kind == "binSize"
                     && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture,
                                        out var binSize))
            {
                var el = new OpenXmlUnknownElement("cx", "binSize", cxNs);
                el.SetAttribute(new OpenXmlAttribute("val", "",
                    binSize.ToString("G", CultureInfo.InvariantCulture)));
                binning.AppendChild(el);
            }
        }
    }
}
