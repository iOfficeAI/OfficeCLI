// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeCli.Core;

/// <summary>
/// Builder for cx:chart (Office 2016 extended chart types):
/// funnel, treemap, sunburst, boxWhisker, histogram, waterfall (native).
///
/// Split into two files:
///   ChartExBuilder.cs        — BuildExtendedChartSpace (Add path)
///   ChartExBuilder.Setter.cs — SetChartProperties      (Set path)
/// Both halves share the same private helpers defined here.
/// </summary>
internal static partial class ChartExBuilder
{
    internal static readonly HashSet<string> ExtendedChartTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "funnel", "treemap", "sunburst", "boxwhisker", "histogram", "pareto"
        // Pareto is a 2-series structure: clusteredColumn (sorted bars) +
        // paretoLine (cumulative-% overlay). PreparePareto pre-sorts desc
        // and computes cumulative %. The value axis is forced to 0-100 so
        // both bars and cumulative line share the same 0-100 range.
        // DetectExtendedChartType handles both OfficeCli-authored and
        // MSO-authored (same 2-series shape) forms.
    };

    internal static bool IsExtendedChartType(string chartType)
    {
        var normalized = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
        return ExtendedChartTypes.Contains(normalized);
    }

    /// <summary>
    /// Build a cx:chartSpace for an extended chart type.
    /// </summary>
    internal static CX.ChartSpace BuildExtendedChartSpace(
        string chartType,
        string? title,
        string[]? categories,
        List<(string name, double[] values)> seriesData,
        Dictionary<string, string> properties)
    {
        var normalized = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");

        // Pareto pre-sorts descending and keeps a single series. The
        // paretoLine series is appended after the main loop with ownerIdx=0
        // (derives from the clusteredColumn series — no separate data needed).
        if (normalized == "pareto")
            (categories, seriesData) = PreparePareto(categories, seriesData);

        var chartSpace = new CX.ChartSpace();

        // 1. Build ChartData
        var chartData = new CX.ChartData();

        // boxWhisker: native Excel structure is one cx:data per group (numDim only,
        // no strDim) + one cx:series per group. The category axis positions each
        // group automatically by series order. Any strDim causes Excel to stack
        // all boxes onto the same X position.
        var dataBlockCount = seriesData.Count;
        for (int si = 0; si < dataBlockCount; si++)
        {
            CX.Data data = normalized == "boxwhisker"
                ? BuildBoxWhiskerGroupDataBlock((uint)si, seriesData[si].values, seriesData[si].name)
                : BuildDataBlock((uint)si, normalized, categories, seriesData[si].values);
            chartData.AppendChild(data);
        }
        chartSpace.AppendChild(chartData);

        // 2. Build Chart
        var chart = new CX.Chart();

        if (!string.IsNullOrEmpty(title))
        {
            chart.AppendChild(BuildChartTitle(title, properties));
        }

        var plotArea = new CX.PlotArea();
        var plotAreaRegion = new CX.PlotAreaRegion();

        var layoutId = normalized switch
        {
            "funnel" => "funnel",
            "treemap" => "treemap",
            "sunburst" => "sunburst",
            "boxwhisker" => "boxWhisker",
            "histogram" => "clusteredColumn",
            "pareto" => "clusteredColumn",
            _ => "funnel"
        };

        // Parse series fill colors — reuse the `colors=RED,BLUE,GREEN`
        // convention from regular charts, or accept a single `fill=COLOR`
        // for one-series charts like histogram.
        var seriesColors = ChartHelper.ParseSeriesColors(properties);
        if (seriesColors == null && properties.TryGetValue("fill", out var fillStr))
            seriesColors = new[] { fillStr };

        // dataLabels: off by default. Accept "true" / "on" / "1" / "value"
        // (any explicit truthy value enables). "false" / "off" / "0" disables.
        var showDataLabels = IsTruthyProp(properties, "dataLabels", defaultValue: false);

        // All chart types including boxWhisker: one cx:series per data set.
        // boxWhisker gets one series per group, matching the one-cx:data-per-group
        // structure above. Colors are set per-series via cx:spPr.
        for (int si = 0; si < seriesData.Count; si++)
        {
                var series = new CX.Series { LayoutId = new EnumValue<CX.SeriesLayout>(
                    ParseSeriesLayout(layoutId)) };

                // Schema order for cx:series:
                //   tx → spPr → valueColors → valueColorPositions → dataPoint*
                //   → dataLabels → dataId → layoutPr → axisId* → extLst
                series.AppendChild(new CX.Text(
                    new CX.TextData(
                        new CX.Formula(""),
                        new CX.VXsdstring(seriesData[si].name))));

                // Per-series solid fill
                if (seriesColors != null && si < seriesColors.Length && !string.IsNullOrEmpty(seriesColors[si]))
                {
                    var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(seriesColors[si]);
                    series.AppendChild(new CX.ShapeProperties(
                        new Drawing.SolidFill(
                            new Drawing.RgbColorModelHex { Val = rgb })));
                }

                // Optional series.shadow (applied to every series). Reuses the
                // ApplyCxSeriesShadow helper so the Add and Set paths emit
                // identical trees.
                var seriesShadow = properties.GetValueOrDefault("series.shadow")
                                ?? properties.GetValueOrDefault("seriesshadow");
                if (!string.IsNullOrEmpty(seriesShadow))
                    ApplyCxSeriesShadow(series, seriesShadow);

                // Data labels (value count above each bar)
                if (showDataLabels)
                {
                    var dl = new CX.DataLabels { Pos = CX.DataLabelPos.OutEnd };
                    dl.AppendChild(new CX.DataLabelVisibilities
                    {
                        Value = true,
                        SeriesName = false,
                        CategoryName = false,
                    });
                    // Optional number format (datalabels.numfmt / labelnumfmt).
                    var dlNumFmt = properties.GetValueOrDefault("datalabels.numfmt")
                                ?? properties.GetValueOrDefault("labelnumfmt")
                                ?? properties.GetValueOrDefault("datalabels.format")
                                ?? properties.GetValueOrDefault("labelformat");
                    if (!string.IsNullOrEmpty(dlNumFmt))
                    {
                        dl.NumberFormat = new CX.NumberFormat
                        {
                            FormatCode = dlNumFmt,
                            SourceLinked = false,
                        };
                    }
                    series.AppendChild(dl);
                }

                series.AppendChild(new CX.DataId { Val = (uint)si });

                // Chart-type specific layoutPr (histogram binning, treemap label
                // layout, boxWhisker stats, etc.). Pareto's clusteredColumn
                // series must NOT have binning — the data is categorical
                // (strDim categories), not continuous numeric for histogram bins.
                if (normalized != "pareto")
                {
                    var layoutPr = BuildLayoutProperties(normalized, properties, seriesData[si].values.Length);
                    if (layoutPr != null)
                        series.AppendChild(layoutPr);
                }

                // Pareto clusteredColumn series: explicit axisId binding to
                // the primary value axis (id=1), matching MSO's structure.
                if (normalized == "pareto")
                {
                    const string cxAxNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
                    var barAxisId = new OpenXmlUnknownElement("cx", "axisId", cxAxNs);
                    barAxisId.SetAttribute(new OpenXmlAttribute("val", "", "1"));
                    series.AppendChild(barAxisId);
                }

                plotAreaRegion.AppendChild(series);
        }

        // Pareto: append the paretoLine overlay series (derives from series 0
        // via ownerIdx="0", auto-computes cumulative %; bound to the secondary
        // percentage axis id=2). Matches MSO's on-the-wire structure.
        if (normalized == "pareto")
        {
            const string cxParetoNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
            var paretoLine = new CX.Series
            {
                LayoutId = new EnumValue<CX.SeriesLayout>(CX.SeriesLayout.ParetoLine),
                OwnerIdx = 0,
            };
            var axisIdEl = new OpenXmlUnknownElement("cx", "axisId", cxParetoNs);
            axisIdEl.SetAttribute(new OpenXmlAttribute("val", "", "2"));
            paretoLine.AppendChild(axisIdEl);
            plotAreaRegion.AppendChild(paretoLine);
        }

        plotArea.AppendChild(plotAreaRegion);

        // Axes for chart types that need them (histogram / boxWhisker / pareto).
        // Funnel/treemap/sunburst are axis-less. Pareto gets 3 axes: cat(0),
        // primary val(1) for bars, secondary percentage(2) for the cumulative line.
        if (normalized is "boxwhisker" or "histogram" or "pareto")
        {
            plotArea.AppendChild(BuildCategoryAxis(id: 0, chartType: normalized, properties));
            plotArea.AppendChild(BuildValueAxis(id: 1, properties));

            if (normalized == "pareto")
            {
                // Secondary percentage axis for the cumulative line (0-100%).
                // Uses raw elements for cx:units since the SDK doesn't expose
                // a typed CX.Units class.
                const string cxAxisNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
                var pctAxis = new CX.Axis { Id = 2 };
                pctAxis.AppendChild(new CX.ValueAxisScaling { Max = "1", Min = "0" });
                var unitsEl = new OpenXmlUnknownElement("cx", "units", cxAxisNs);
                unitsEl.SetAttribute(new OpenXmlAttribute("unit", "", "percentage"));
                pctAxis.AppendChild(unitsEl);
                pctAxis.AppendChild(new CX.TickLabels());
                plotArea.AppendChild(pctAxis);
            }
        }

        // Plot area fill / border — optional background styling
        // (CONSISTENCY(chart-area-fill)). Must be appended AFTER all axes
        // per CT_PlotArea schema sequence:
        //   plotSurface? → plotAreaRegion → axis* → spPr? → extLst?
        var plotAreaFill = properties.GetValueOrDefault("plotareafill")
                        ?? properties.GetValueOrDefault("plotfill");
        if (!string.IsNullOrEmpty(plotAreaFill))
            ApplyCxAreaFill(plotArea, plotAreaFill);

        var plotAreaBorder = properties.GetValueOrDefault("plotarea.border")
                          ?? properties.GetValueOrDefault("plotborder");
        if (!string.IsNullOrEmpty(plotAreaBorder))
            ApplyCxAreaBorder(plotArea, plotAreaBorder);

        chart.AppendChild(plotArea);

        // Legend (optional, appears AFTER plotArea per cx:chart schema order).
        // BuildLegend reads legend.overlay / legendfont from properties too.
        if (properties.TryGetValue("legend", out var legendPos) &&
            !string.IsNullOrEmpty(legendPos) &&
            !legendPos.Equals("none", StringComparison.OrdinalIgnoreCase) &&
            !legendPos.Equals("false", StringComparison.OrdinalIgnoreCase) &&
            !legendPos.Equals("off", StringComparison.OrdinalIgnoreCase))
        {
            chart.AppendChild(BuildLegend(legendPos, properties));
        }

        chartSpace.AppendChild(chart);

        // Chart area fill / border — attached to cx:chartSpace's own spPr.
        // This is the outermost background; tests should verify Excel
        // accepts it (the cx schema technically does not list spPr as a
        // chartSpace child but the SDK tolerates it; real Excel silently
        // ignores it rather than rejecting, so we still emit it for
        // round-trip Set() compatibility).
        var chartAreaFill = properties.GetValueOrDefault("chartareafill")
                         ?? properties.GetValueOrDefault("chartfill");
        if (!string.IsNullOrEmpty(chartAreaFill))
            ApplyCxAreaFill(chartSpace, chartAreaFill);

        var chartAreaBorder = properties.GetValueOrDefault("chartarea.border")
                           ?? properties.GetValueOrDefault("chartborder");
        if (!string.IsNullOrEmpty(chartAreaBorder))
            ApplyCxAreaBorder(chartSpace, chartAreaBorder);

        return chartSpace;
    }

    private static CX.ChartTitle BuildChartTitle(string title, Dictionary<string, string>? properties = null)
    {
        var rPr = new Drawing.RunProperties { Language = "en-US" };
        // Delegate style parsing to the shared helper so cChart and cxChart
        // stay in vocabulary lockstep. See
        // ChartHelper.ApplyRunStyleProperties.
        if (properties != null)
        {
            ChartHelper.ApplyRunStyleProperties(rPr, properties, keyPrefix: "title");

            // title.shadow is a separate knob — ApplyRunStyleProperties covers
            // color/size/bold/font only (see its doc-comment). Same format as
            // regular cChart: "COLOR-BLUR-ANGLE-DIST-OPACITY".
            var titleShadow = properties.GetValueOrDefault("title.shadow")
                           ?? properties.GetValueOrDefault("titleshadow");
            if (!string.IsNullOrEmpty(titleShadow))
                ApplyRunEffectShadow(rPr, titleShadow);
        }

        var chartTitle = new CX.ChartTitle();
        chartTitle.AppendChild(new CX.Text(
            new CX.RichTextBody(
                new Drawing.BodyProperties(),
                new Drawing.Paragraph(
                    new Drawing.Run(
                        rPr,
                        new Drawing.Text(title))))));
        return chartTitle;
    }

    private static CX.AxisTitle BuildAxisTitle(string title, Dictionary<string, string>? properties = null)
    {
        var rPr = new Drawing.RunProperties { Language = "en-US" };
        if (properties != null)
            ChartHelper.ApplyRunStyleProperties(rPr, properties, keyPrefix: "axisTitle");

        return new CX.AxisTitle(
            new CX.Text(
                new CX.RichTextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            rPr,
                            new Drawing.Text(title))))));
    }

    /// <summary>
    /// Wrap a shared `a:defRPr` (built from a compound `"size:color:fontname"`
    /// spec by <see cref="ChartHelper.BuildDefaultRunPropertiesFromCompoundSpec"/>)
    /// in a <see cref="CX.TxPrTextBody"/>. Only the outer container differs
    /// from the regular-cChart path (<see cref="C.TextProperties"/>).
    /// </summary>
    private static CX.TxPrTextBody? BuildAxisTickLabelStyle(string compoundSpec)
    {
        if (string.IsNullOrEmpty(compoundSpec)) return null;
        var defRp = ChartHelper.BuildDefaultRunPropertiesFromCompoundSpec(compoundSpec);
        return new CX.TxPrTextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.ParagraphProperties(defRp)));
    }

    /// <summary>
    /// Build a <see cref="CX.ShapeProperties"/> containing a solid-fill outline
    /// for coloring gridlines. Mirrors the regular-chart `gridline.color` knob.
    /// </summary>
    private static CX.ShapeProperties? BuildGridlineShapeProperties(string color)
    {
        if (string.IsNullOrEmpty(color)) return null;
        var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(color);
        var outline = new Drawing.Outline();
        outline.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb }));
        return new CX.ShapeProperties(outline);
    }

    private static CX.Legend BuildLegend(string posSpec, Dictionary<string, string>? properties = null)
    {
        var legend = new CX.Legend
        {
            Align = CX.PosAlign.Ctr,
            Overlay = false,
        };
        // CONSISTENCY(strict-enums / R34-1): unknown legend tokens used to
        // silently fall through to right; mirror cChart's strict validation.
        // Note: cx:legend's SidePos has no topRight — fall back to top with
        // a clear note rather than rejecting, since topRight is a valid
        // value for the regular cChart variant and users may pass it through.
        // CONSISTENCY(legend-separator-normalize): mirror SetterHelpers — dash
        // and underscore separators are equivalent (top-right == top_right).
        var posSpecNorm = (posSpec ?? string.Empty).ToLowerInvariant().Replace("-", "").Replace("_", "");
        legend.Pos = posSpecNorm switch
        {
            "top" or "t" or "topright" or "tr" => CX.SidePos.T,
            "bottom" or "b" => CX.SidePos.B,
            "left" or "l"   => CX.SidePos.L,
            "right" or "r"  => CX.SidePos.R,
            _ => throw new ArgumentException(
                $"Invalid legend position '{posSpec}'. " +
                "Valid: none, top, bottom, left, right, topRight " +
                "(or use 'none'/'false' to hide the legend)."),
        };

        if (properties != null)
        {
            // Optional overlay flag — matches regular cChart's `legend.overlay`.
            var overlay = properties.GetValueOrDefault("legend.overlay")
                       ?? properties.GetValueOrDefault("legendoverlay");
            if (!string.IsNullOrEmpty(overlay))
                legend.Overlay = ParseHelpers.IsTruthy(overlay);

            // Compound font styling — "size:color:fontname", same form as
            // regular cChart's `legendfont`. Wraps an a:defRPr in cx:txPr.
            var legendFont = properties.GetValueOrDefault("legendfont")
                          ?? properties.GetValueOrDefault("legend.font");
            if (!string.IsNullOrEmpty(legendFont))
            {
                var txPr = BuildAxisTickLabelStyle(legendFont);
                if (txPr != null) legend.AppendChild(txPr);
            }
        }

        return legend;
    }

    // ==================== Shared cx:spPr / effect helpers ====================
    //
    // These helpers mirror the regular-cChart versions in
    // ChartHelper.SetterHelpers.cs (ApplyAxisLine, BuildOutlineElement,
    // DrawingEffectsHelper.BuildOuterShadow) but target cx:spPr containers
    // instead of c:spPr / c:ChartShapeProperties.
    //
    // They are used by BOTH the Add path (ChartExBuilder.cs BuildExtended...)
    // and the Set path (ChartExBuilder.Setter.cs HandleSetKey), so each knob
    // creates the same OOXML tree regardless of whether it was set at Add
    // time or via a later Set call.

    /// <summary>
    /// Apply an a:outerShdw effect to a Drawing.RunProperties (used for
    /// `title.shadow`). Reuses the shared DrawingEffectsHelper format:
    /// "COLOR-BLUR-ANGLE-DIST-OPACITY" or "none" to clear.
    /// </summary>
    private static void ApplyRunEffectShadow(Drawing.RunProperties rPr, string value)
    {
        rPr.RemoveAllChildren<Drawing.EffectList>();
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase)) return;
        var effects = new Drawing.EffectList();
        effects.AppendChild(DrawingEffectsHelper.BuildOuterShadow(
            value, DrawingEffectsHelper.BuildRgbColor));
        rPr.AppendChild(effects);
    }

    /// <summary>
    /// Apply an a:ln outline to a cx:axis's own cx:spPr. Same vocabulary as
    /// ChartHelper.SetterHelpers.cs:ApplyAxisLine — "color" / "color:width" /
    /// "color:width:dash" / "none".
    /// </summary>
    private static void ApplyCxAxisLine(CX.Axis axis, string value)
    {
        var spPr = axis.GetFirstChild<CX.ShapeProperties>();
        if (spPr == null)
        {
            spPr = new CX.ShapeProperties();
            // cx:spPr comes after tickLabels but before txPr in the cx:axis
            // schema (catScaling → title → gridlines → tickLabels → numFmt
            // → spPr → txPr → extLst).
            var existingTxPr = axis.GetFirstChild<CX.TxPrTextBody>();
            if (existingTxPr != null) axis.InsertBefore(spPr, existingTxPr);
            else axis.AppendChild(spPr);
        }
        spPr.RemoveAllChildren<Drawing.Outline>();
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var noFillOutline = new Drawing.Outline();
            noFillOutline.AppendChild(new Drawing.NoFill());
            spPr.PrependChild(noFillOutline);
            return;
        }
        spPr.PrependChild(ChartHelper.BuildOutlineElement(value));
    }

    /// <summary>
    /// Apply an a:outerShdw (inside a:effectLst) to a cx:series's own cx:spPr.
    /// Preserves any existing solidFill so the series keeps its color.
    /// </summary>
    private static void ApplyCxSeriesShadow(CX.Series series, string value)
    {
        var spPr = series.GetFirstChild<CX.ShapeProperties>();
        if (spPr == null)
        {
            spPr = new CX.ShapeProperties();
            // spPr goes right after cx:tx per cx:series schema.
            var tx = series.GetFirstChild<CX.Text>();
            if (tx != null) tx.InsertAfterSelf(spPr);
            else series.PrependChild(spPr);
        }
        // Remove any existing effectList so repeated Sets don't stack.
        spPr.RemoveAllChildren<Drawing.EffectList>();
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase)) return;
        var effects = new Drawing.EffectList();
        effects.AppendChild(DrawingEffectsHelper.BuildOuterShadow(
            value, DrawingEffectsHelper.BuildRgbColor));
        spPr.AppendChild(effects);
    }

    /// <summary>
    /// Apply a solid background fill to a cx:plotArea or cx:chartSpace via
    /// its own cx:spPr child. Accepts "none" to clear.
    /// </summary>
    private static void ApplyCxAreaFill(OpenXmlCompositeElement container, string value)
    {
        var spPr = container.GetFirstChild<CX.ShapeProperties>();
        if (spPr == null)
        {
            spPr = new CX.ShapeProperties();
            container.AppendChild(spPr);
        }
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            spPr.PrependChild(new Drawing.NoFill());
            return;
        }
        var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
        spPr.PrependChild(new Drawing.SolidFill(
            new Drawing.RgbColorModelHex { Val = rgb }));
    }

    /// <summary>
    /// Apply an a:ln outline border to a cx:plotArea or cx:chartSpace via its
    /// own cx:spPr child. Shares the "color / color:width / color:width:dash"
    /// vocabulary with ChartHelper.BuildOutlineElement.
    /// </summary>
    private static void ApplyCxAreaBorder(OpenXmlCompositeElement container, string value)
    {
        var spPr = container.GetFirstChild<CX.ShapeProperties>();
        if (spPr == null)
        {
            spPr = new CX.ShapeProperties();
            container.AppendChild(spPr);
        }
        spPr.RemoveAllChildren<Drawing.Outline>();
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var noFillOutline = new Drawing.Outline();
            noFillOutline.AppendChild(new Drawing.NoFill());
            spPr.AppendChild(noFillOutline);
            return;
        }
        spPr.AppendChild(ChartHelper.BuildOutlineElement(value));
    }

    // Build the category axis (X axis for histogram / boxWhisker). Schema
    // order of Axis children: catScaling → title → majorGridlines →
    // tickLabels → ... (only the ones we emit are listed).
    private static CX.Axis BuildCategoryAxis(uint id, string chartType, Dictionary<string, string> properties)
    {
        var axis = new CX.Axis { Id = id };

        // CONSISTENCY(chart-axis-visibility): apply @hidden from axis.visible
        // / cataxis.visible / axis.delete props. See ApplyAxisHiddenFromProps
        // for the precedence rules.
        ApplyAxisHiddenFromProps(axis, properties, catOnly: true, valOnly: false);

        // catScaling is required. histogram defaults gapWidth="0" (bars touch)
        // because that's what real Excel emits and it's what users expect.
        var catScaling = new CX.CategoryAxisScaling();
        var gapWidth = properties.GetValueOrDefault("gapWidth");
        if (string.IsNullOrEmpty(gapWidth) && chartType == "histogram")
            gapWidth = "0";
        if (!string.IsNullOrEmpty(gapWidth))
            catScaling.GapWidth = gapWidth;
        axis.AppendChild(catScaling);

        if (properties.TryGetValue("xAxisTitle", out var xTitle) && !string.IsNullOrEmpty(xTitle))
            axis.AppendChild(BuildAxisTitle(xTitle, properties));

        // Category-axis major gridlines are off by default in Excel; opt-in.
        if (IsTruthyProp(properties, "xGridlines", defaultValue: false))
        {
            var gl = new CX.MajorGridlinesGridlines();
            // CONSISTENCY(chart-text-style): category-axis gridline color uses
            // `xGridlineColor` to distinguish from value-axis `gridlineColor`.
            var xglColor = properties.GetValueOrDefault("xGridlineColor")
                        ?? properties.GetValueOrDefault("xGridline.color");
            if (!string.IsNullOrEmpty(xglColor))
                gl.ShapeProperties = BuildGridlineShapeProperties(xglColor);
            axis.AppendChild(gl);
        }

        // Tick labels (bin range labels like "[100, 200]") are ON by default
        // to match real Excel output. Opt out with tickLabels=false. Note
        // that cx:tickLabels itself is an EMPTY element per CT_TickLabels —
        // label text styling lives on the axis's own cx:txPr sibling (below),
        // NOT inside tickLabels. Nesting txPr under tickLabels produces
        // schema-invalid XML that Excel silently "repairs".
        if (IsTruthyProp(properties, "tickLabels", defaultValue: true))
            axis.AppendChild(new CX.TickLabels());

        // CONSISTENCY(chart-text-style): axis-level cx:txPr styles tick
        // labels AND axis title text, matching what ApplyAxisTextProperties
        // does for regular cChart. Compound form `axisfont=size:color:fontname`.
        // Must be AFTER tickLabels per CT_Axis schema sequence
        // (catScaling → title → gridlines → tickLabels → numFmt → spPr → txPr).
        var axisFont = properties.GetValueOrDefault("axisfont")
                    ?? properties.GetValueOrDefault("axis.font");
        if (!string.IsNullOrEmpty(axisFont))
        {
            var tickTxPr = BuildAxisTickLabelStyle(axisFont);
            if (tickTxPr != null) axis.AppendChild(tickTxPr);
        }

        // CONSISTENCY(chart-axis-line): optional category-axis spine outline.
        // cataxis.line takes precedence over the shared axis.line.
        var catAxisLine = properties.GetValueOrDefault("cataxisline")
                       ?? properties.GetValueOrDefault("cataxis.line")
                       ?? properties.GetValueOrDefault("axisline")
                       ?? properties.GetValueOrDefault("axis.line");
        if (!string.IsNullOrEmpty(catAxisLine))
            ApplyCxAxisLine(axis, catAxisLine);

        return axis;
    }

    private static CX.Axis BuildValueAxis(uint id, Dictionary<string, string> properties)
    {
        var axis = new CX.Axis { Id = id };

        // CONSISTENCY(chart-axis-visibility): axis.visible / axis.delete are
        // mutually exclusive aliases for the same knob. valaxis.visible is
        // the value-axis-only variant (matches ChartHelper.Setter.cs:817).
        ApplyAxisHiddenFromProps(axis, properties, catOnly: false, valOnly: true);

        // CONSISTENCY(chart-axis-scaling): parse axismin/axismax/majorunit/
        // minorunit at Build time so newly created charts already have them.
        var valScaling = new CX.ValueAxisScaling();
        ApplyValueAxisScalingFromProps(valScaling, properties);
        axis.AppendChild(valScaling);

        if (properties.TryGetValue("yAxisTitle", out var yTitle) && !string.IsNullOrEmpty(yTitle))
            axis.AppendChild(BuildAxisTitle(yTitle, properties));

        // Value-axis gridlines are ON by default — matches Excel's histogram
        // and column charts out of the box.
        if (IsTruthyProp(properties, "gridlines", defaultValue: true))
        {
            var gl = new CX.MajorGridlinesGridlines();
            var glColor = properties.GetValueOrDefault("gridlineColor")
                       ?? properties.GetValueOrDefault("gridline.color");
            if (!string.IsNullOrEmpty(glColor))
                gl.ShapeProperties = BuildGridlineShapeProperties(glColor);
            axis.AppendChild(gl);
        }

        if (IsTruthyProp(properties, "tickLabels", defaultValue: true))
            axis.AppendChild(new CX.TickLabels());

        // cx:txPr must come after tickLabels per CT_Axis schema. See the
        // CONSISTENCY(chart-text-style) note in BuildCategoryAxis above.
        var axisFont = properties.GetValueOrDefault("axisfont")
                    ?? properties.GetValueOrDefault("axis.font");
        if (!string.IsNullOrEmpty(axisFont))
        {
            var tickTxPr = BuildAxisTickLabelStyle(axisFont);
            if (tickTxPr != null) axis.AppendChild(tickTxPr);
        }

        // CONSISTENCY(chart-axis-line): optional value-axis spine outline.
        // Accepts "color", "color:width", "color:width:dash", or "none".
        // ApplyCxAxisLine handles placement within the cx:axis schema.
        var valAxisLine = properties.GetValueOrDefault("valaxisline")
                       ?? properties.GetValueOrDefault("valaxis.line")
                       ?? properties.GetValueOrDefault("axisline")
                       ?? properties.GetValueOrDefault("axis.line");
        if (!string.IsNullOrEmpty(valAxisLine))
            ApplyCxAxisLine(axis, valAxisLine);

        return axis;
    }

    /// <summary>
    /// Apply CX.Axis.Hidden from the three-way prop set: axis.visible /
    /// axisvisible / axis.delete (both axes), cataxis.visible /
    /// cataxisvisible (category-only), valaxis.visible / valaxisvisible
    /// (value-only). The caller passes catOnly/valOnly flags indicating
    /// which specific axis is being built; the shared prop still applies
    /// universally. Matches ChartHelper.Setter.cs:795.
    /// </summary>
    private static void ApplyAxisHiddenFromProps(
        CX.Axis axis, Dictionary<string, string> properties, bool catOnly, bool valOnly)
    {
        // Universal axis.visible / axis.delete first (if present).
        var universalVisible = properties.GetValueOrDefault("axis.visible")
                            ?? properties.GetValueOrDefault("axisvisible");
        if (!string.IsNullOrEmpty(universalVisible))
            axis.Hidden = !ParseHelpers.IsTruthy(universalVisible);

        var universalDelete = properties.GetValueOrDefault("axis.delete");
        if (!string.IsNullOrEmpty(universalDelete))
            axis.Hidden = ParseHelpers.IsTruthy(universalDelete);

        // Axis-specific override (takes precedence over the universal form).
        if (catOnly)
        {
            var cv = properties.GetValueOrDefault("cataxis.visible")
                  ?? properties.GetValueOrDefault("cataxisvisible");
            if (!string.IsNullOrEmpty(cv)) axis.Hidden = !ParseHelpers.IsTruthy(cv);
        }
        if (valOnly)
        {
            var vv = properties.GetValueOrDefault("valaxis.visible")
                  ?? properties.GetValueOrDefault("valaxisvisible");
            if (!string.IsNullOrEmpty(vv)) axis.Hidden = !ParseHelpers.IsTruthy(vv);
        }
    }

    /// <summary>
    /// Copy axismin / axismax / majorunit / minorunit from properties onto
    /// a <see cref="CX.ValueAxisScaling"/>. These are string-typed attributes
    /// in cx namespace (unlike c:scaling which uses typed doubles), but we
    /// still round-trip through <see cref="ParseHelpers.SafeParseDouble"/>
    /// so NaN/Infinity are rejected.
    /// </summary>
    private static void ApplyValueAxisScalingFromProps(
        CX.ValueAxisScaling scaling, Dictionary<string, string> properties)
    {
        string? FormatIfPresent(string keyA, string? keyB)
        {
            var v = properties.GetValueOrDefault(keyA);
            if (string.IsNullOrEmpty(v) && keyB != null) v = properties.GetValueOrDefault(keyB);
            if (string.IsNullOrEmpty(v)) return null;
            var d = ParseHelpers.SafeParseDouble(v, keyA);
            return d.ToString("G", CultureInfo.InvariantCulture);
        }

        var min = FormatIfPresent("axismin", "min");
        if (min != null) scaling.Min = min;

        var max = FormatIfPresent("axismax", "max");
        if (max != null) scaling.Max = max;

        var maj = FormatIfPresent("majorunit", null);
        if (maj != null) scaling.MajorUnit = maj;

        var mnr = FormatIfPresent("minorunit", null);
        if (mnr != null) scaling.MinorUnit = mnr;
    }

    private static bool IsTruthyProp(Dictionary<string, string> properties, string key, bool defaultValue)
    {
        if (!properties.TryGetValue(key, out var v) || string.IsNullOrEmpty(v))
            return defaultValue;
        return !(v.Equals("false", StringComparison.OrdinalIgnoreCase)
              || v.Equals("off", StringComparison.OrdinalIgnoreCase)
              || v == "0"
              || v.Equals("no", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Build a single cx:data block for one boxWhisker group.
    /// Includes a strDim type="cat" with the group name repeated once per data
    /// point so the X axis shows the group label. The strDim is per-data-block
    /// (not shared across series), so each group stays at its own X position.
    /// </summary>
    private static CX.Data BuildBoxWhiskerGroupDataBlock(uint id, double[] values, string groupName)
    {
        var data = new CX.Data { Id = id };

        // strDim provides the X-axis label for this group.
        // Repeat the group name once per data point (required: ptCount must equal numDim ptCount).
        var strDim = new CX.StringDimension { Type = CX.StringDimensionType.Cat };
        var strLvl = new CX.StringLevel { PtCount = (uint)values.Length };
        for (int i = 0; i < values.Length; i++)
            strLvl.AppendChild(new CX.ChartStringValue(groupName) { Index = (uint)i });
        strDim.AppendChild(strLvl);
        data.AppendChild(strDim);

        var numDim = new CX.NumericDimension { Type = CX.NumericDimensionType.Val };
        var numLvl = new CX.NumericLevel { PtCount = (uint)values.Length, FormatCode = "General" };
        for (int i = 0; i < values.Length; i++)
            numLvl.AppendChild(new CX.NumericValue(values[i].ToString("G", CultureInfo.InvariantCulture)) { Idx = (uint)i });
        numDim.AppendChild(numLvl);
        data.AppendChild(numDim);

        return data;
    }

    private static CX.Data BuildDataBlock(uint id, string chartType, string[]? categories, double[] values)
    {
        var data = new CX.Data { Id = id };

        // String dimension for categories (if provided). Pareto is included
        // because both of its series (clusteredColumn + paretoLine) share
        // the same sorted category labels — unlike histogram which auto-bins
        // numeric data and has no explicit categories.
        if (categories != null && chartType is "funnel" or "treemap" or "sunburst" or "boxwhisker" or "pareto")
        {
            var strDim = new CX.StringDimension { Type = CX.StringDimensionType.Cat };

            // boxWhisker: each data block carries ONE group label but N values.
            // strDim.PtCount must equal numDim.PtCount — Excel requires them to
            // match or it collapses all series onto the same X position.
            // Repeat the single label N times (once per data point) so the
            // counts align. funnel/treemap/sunburst keep their original 1:1 mapping.
            bool repeatSingle = chartType == "boxwhisker" && categories.Length == 1;
            int ptCount = repeatSingle ? values.Length : categories.Length;

            var strLvl = new CX.StringLevel { PtCount = (uint)ptCount };
            for (int i = 0; i < ptCount; i++)
            {
                string cat = repeatSingle ? categories[0] : categories[i];
                strLvl.AppendChild(new CX.ChartStringValue(cat) { Index = (uint)i });
            }
            strDim.AppendChild(strLvl);
            data.AppendChild(strDim);
        }

        // Numeric dimension
        var numType = chartType is "treemap" or "sunburst"
            ? CX.NumericDimensionType.Size
            : CX.NumericDimensionType.Val;
        var numDim = new CX.NumericDimension { Type = numType };
        var numLvl = new CX.NumericLevel { PtCount = (uint)values.Length, FormatCode = "General" };
        for (int i = 0; i < values.Length; i++)
            numLvl.AppendChild(new CX.NumericValue(values[i].ToString("G")) { Idx = (uint)i });
        numDim.AppendChild(numLvl);
        data.AppendChild(numDim);

        return data;
    }

    private static CX.SeriesLayoutProperties? BuildLayoutProperties(
        string chartType, Dictionary<string, string> properties, int valueCount)
    {
        switch (chartType)
        {
            case "treemap":
            {
                var lp = new CX.SeriesLayoutProperties();
                var parentLayout = properties.GetValueOrDefault("parentLabelLayout") ?? "overlapping";
                lp.AppendChild(new CX.ParentLabelLayout
                {
                    ParentLabelLayoutVal = parentLayout.ToLowerInvariant() switch
                    {
                        "none" => CX.ParentLabelLayoutVal.None,
                        "banner" => CX.ParentLabelLayoutVal.Banner,
                        _ => CX.ParentLabelLayoutVal.Overlapping
                    }
                });
                return lp;
            }
            case "boxwhisker":
            {
                var lp = new CX.SeriesLayoutProperties();
                lp.AppendChild(new CX.SeriesElementVisibilities
                {
                    MeanLine = false, MeanMarker = true,
                    Nonoutliers = false, Outliers = true
                });
                var method = properties.GetValueOrDefault("quartileMethod") ?? "exclusive";
                lp.AppendChild(new CX.Statistics
                {
                    QuartileMethod = method.ToLowerInvariant() switch
                    {
                        "inclusive" => CX.QuartileMethod.Inclusive,
                        _ => CX.QuartileMethod.Exclusive
                    }
                });
                return lp;
            }
            case "histogram":
            {
                // cx:layoutPr > cx:binning (empty for auto-bin; child cx:binCount
                // OR cx:binSize for explicit bin count/width). `cx:aggregation`
                // is for Pareto charts and causes Excel to render the whole
                // dataset as a single bar.
                //
                // NOTE: the Open XML SDK models cx:binCount as a leaf text
                // element (BinCountXsdunsignedInt → `<cx:binCount>5</cx:binCount>`),
                // but real Excel writes it as an empty element with a `val`
                // attribute (`<cx:binCount val="5"/>`). SDK's form is schema-
                // valid per the generated type metadata but Excel rejects the
                // whole file with "We found a problem with some content"
                // and deletes the drawing. Same applies to cx:binSize. Work
                // around by appending a raw OpenXmlUnknownElement carrying
                // the correct form.
                const string cxNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
                var lp = new CX.SeriesLayoutProperties();
                var binning = new CX.Binning();

                // intervalClosed: "r" (default, bins are (a,b]) or "l" (bins are [a,b))
                var intervalClosed = properties.GetValueOrDefault("intervalClosed") ?? "r";
                binning.IntervalClosed = intervalClosed.ToLowerInvariant() switch
                {
                    "l" => CX.IntervalClosedSide.L,
                    _   => CX.IntervalClosedSide.R,
                };

                // underflow / overflow: cut-off values for outlier bins
                if (properties.TryGetValue("underflowBin", out var underflow))
                    binning.Underflow = underflow;
                if (properties.TryGetValue("overflowBin", out var overflow))
                    binning.Overflow = overflow;

                // binCount (explicit count) XOR binSize (explicit width). If
                // both are given, binCount wins (it's the more common knob).
                if (properties.TryGetValue("binCount", out var binCountStr) &&
                    uint.TryParse(binCountStr, out var binCount))
                {
                    var binCountEl = new OpenXmlUnknownElement("cx", "binCount", cxNs);
                    binCountEl.SetAttribute(new OpenXmlAttribute("val", "", binCount.ToString()));
                    binning.AppendChild(binCountEl);
                }
                else if (properties.TryGetValue("binSize", out var binSizeStr) &&
                         double.TryParse(binSizeStr, System.Globalization.NumberStyles.Float,
                             System.Globalization.CultureInfo.InvariantCulture, out var binSize))
                {
                    var binSizeEl = new OpenXmlUnknownElement("cx", "binSize", cxNs);
                    binSizeEl.SetAttribute(new OpenXmlAttribute("val", "",
                        binSize.ToString("G", System.Globalization.CultureInfo.InvariantCulture)));
                    binning.AppendChild(binSizeEl);
                }

                lp.AppendChild(binning);
                return lp;
            }
            default:
                return null;
        }
    }

    private static CX.SeriesLayout ParseSeriesLayout(string layoutId)
    {
        return layoutId switch
        {
            "funnel" => CX.SeriesLayout.Funnel,
            "treemap" => CX.SeriesLayout.Treemap,
            "sunburst" => CX.SeriesLayout.Sunburst,
            "boxWhisker" => CX.SeriesLayout.BoxWhisker,
            "clusteredColumn" => CX.SeriesLayout.ClusteredColumn,
            "paretoLine" => CX.SeriesLayout.ParetoLine,
            "regionMap" => CX.SeriesLayout.RegionMap,
            _ => CX.SeriesLayout.Funnel
        };
    }

    /// <summary>
    /// Detect if a cx:chartSpace contains an extended chart type and return the type name.
    /// Also handles MSO-authored Pareto files which may contain both a clusteredColumn
    /// and a paretoLine series — if any series has paretoLine layout, it's a pareto.
    /// </summary>
    internal static string? DetectExtendedChartType(CX.ChartSpace chartSpace)
    {
        var allSeries = chartSpace.Descendants<CX.Series>().ToList();
        if (allSeries.Count == 0) return null;

        // Pareto: any paretoLine series ⇒ the whole chart is a pareto.
        // Handles both OfficeCli-authored (single paretoLine series) and
        // MSO-authored (clusteredColumn + paretoLine pair) forms.
        if (allSeries.Any(s => s.LayoutId?.InnerText == "paretoLine"))
            return "pareto";

        var layoutId = allSeries[0].LayoutId?.InnerText;
        if (layoutId == null) return null;
        return layoutId switch
        {
            "funnel" => "funnel",
            "treemap" => "treemap",
            "sunburst" => "sunburst",
            "boxWhisker" => "boxWhisker",
            "clusteredColumn" => "histogram",
            "regionMap" => "regionMap",
            _ => layoutId
        };
    }

    /// <summary>
    /// Transform a user's single-series Pareto input into the 2-series form
    /// that Excel's cx:chart pareto uses internally. The first user series
    /// is sorted descending (biggest first); cumulative percentages are
    /// computed on the sorted order and returned as the second series.
    /// If the user supplies multiple series, extras are silently ignored —
    /// pareto is inherently univariate.
    /// </summary>
    /// <summary>
    /// Pre-sort the user's single series descending for Pareto. Returns a
    /// single series (the sorted values); the cumulative-% paretoLine
    /// series is appended in BuildExtendedChartSpace via ownerIdx=0
    /// (Excel auto-computes cumulative from the bar data).
    /// </summary>
    private static (string[]? categories, List<(string name, double[] values)> seriesData)
        PreparePareto(string[]? categories, List<(string name, double[] values)> seriesData)
    {
        if (seriesData.Count == 0)
            return (categories, seriesData);

        var (srcName, srcValues) = seriesData[0];
        int n = srcValues.Length;
        if (n == 0)
            return (categories, seriesData);

        var cats = (categories != null && categories.Length == n)
            ? categories
            : Enumerable.Range(1, n).Select(i => i.ToString(CultureInfo.InvariantCulture)).ToArray();

        // Sort by value descending; stable for equal values.
        var indices = Enumerable.Range(0, n).OrderByDescending(i => srcValues[i]).ToArray();
        var sortedCats = indices.Select(i => cats[i]).ToArray();
        var sortedVals = indices.Select(i => srcValues[i]).ToArray();

        var barsName = string.IsNullOrEmpty(srcName) ? "Value" : srcName;
        return (sortedCats, new List<(string, double[])>
        {
            (barsName, sortedVals),
        });
    }
}
