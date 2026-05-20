// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal static partial class ChartHelper
{
    internal static List<string> SetChartProperties(ChartPart chartPart, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var chartSpace = chartPart.ChartSpace;
        var chart = chartSpace?.GetFirstChild<C.Chart>();
        if (chart == null) { unsupported.AddRange(properties.Keys); return unsupported; }

        // R24-3: expand combined "legend.layout=x:N,y:N,w:N,h:N" (and the same
        // form for plotArea/title/trendlineLabel/displayUnitsLabel) into the
        // individual {prefix}.x/y/w/h keys consumed by the dispatch table
        // below. Without this, the combined form was silently accepted by
        // the lenient prefix validator but never emitted any <c:layout>.
        ExpandCombinedLayoutKeys(properties);

        // Process structural properties (legend, title) before styling properties (legendFont, titleFont)
        // to ensure the parent element exists before styling is applied.
        static int PropOrder(string k)
        {
            var lower = k.ToLowerInvariant();
            if (lower is "preset" or "style.preset" or "theme") return 0;
            if (lower is "title" or "legend" or "datalabels" or "labels") return 1;
            return 2;
        }
        var ordered = properties.OrderBy(kv => PropOrder(kv.Key));
        foreach (var (key, value) in ordered)
        {
            switch (key.ToLowerInvariant())
            {
                case "preset" or "style.preset" or "theme":
                {
                    var presetProps = ChartPresets.GetPreset(value);
                    if (presetProps == null)
                        throw new ArgumentException(
                            $"Unknown chart preset '{value}'. Available: {string.Join(", ", ChartPresets.PresetNames)}.");
                    // Recursively apply preset properties
                    var presetUnsupported = SetChartProperties(chartPart, presetProps);
                    // Silently skip title.* properties when chart has no title —
                    // presets include title styling but charts may legitimately have no title
                    var hasTitle = chart.GetFirstChild<C.Title>() != null;
                    if (!hasTitle)
                        presetUnsupported.RemoveAll(k => k.StartsWith("title.", StringComparison.OrdinalIgnoreCase)
                            || (k.StartsWith("title", StringComparison.OrdinalIgnoreCase) && k.Length > 5));
                    unsupported.AddRange(presetUnsupported);
                    break;
                }

                case "title":
                    chart.RemoveAllChildren<C.Title>();
                    if (!string.IsNullOrEmpty(value) && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // CONSISTENCY(autoTitleDeleted-paired): setting a title back
                        // must clear any autoTitleDeleted=1 that an earlier `title=`
                        // removal (or autoTitleDeleted=true) left behind — otherwise
                        // the new title element is present but suppressed at render.
                        chart.RemoveAllChildren<C.AutoTitleDeleted>();
                        chart.PrependChild(BuildChartTitle(value));
                    }
                    break;

                case "title.font" or "titlefont":
                case "title.size" or "titlesize":
                case "title.color" or "titlecolor":
                case "title.bold" or "titlebold":
                case "title.glow" or "titleglow":
                case "title.shadow" or "titleshadow":
                {
                    var ctitle = chart.GetFirstChild<C.Title>();
                    if (ctitle == null) { unsupported.Add(key); break; }
                    foreach (var run in ctitle.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        var normalizedKey = key.Replace("title.", "").Replace("title", "").ToLowerInvariant();
                        switch (normalizedKey)
                        {
                            case "font":
                                rPr.RemoveAllChildren<Drawing.LatinFont>();
                                rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                                rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                                rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                                break;
                            case "size":
                                var sizeStr = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                                    ? value[..^2] : value;
                                rPr.FontSize = (int)Math.Round(ParseHelpers.SafeParseDouble(sizeStr, "title.size") * 100);
                                break;
                            case "color":
                            {
                                rPr.RemoveAllChildren<Drawing.SolidFill>();
                                var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                                DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                                    new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb }));
                                break;
                            }
                            case "bold":
                                rPr.Bold = ParseHelpers.IsTruthy(value);
                                break;
                            case "glow":
                                DrawingEffectsHelper.ApplyTextEffect<Drawing.Glow>(run, value,
                                    () => DrawingEffectsHelper.BuildGlow(value, DrawingEffectsHelper.BuildRgbColor));
                                break;
                            case "shadow":
                                DrawingEffectsHelper.ApplyTextEffect<Drawing.OuterShadow>(run, value,
                                    () => DrawingEffectsHelper.BuildOuterShadow(value, DrawingEffectsHelper.BuildRgbColor));
                                break;
                        }
                        // Also update DefaultRunProperties for consistency
                        var defRp = ctitle.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
                        if (defRp != null)
                        {
                            switch (normalizedKey)
                            {
                                case "size": defRp.FontSize = rPr.FontSize; break;
                                case "bold": defRp.Bold = rPr.Bold; break;
                            }
                        }
                    }
                    break;
                }

                case "legendfont" or "legend.font":
                {
                    // Format: "size:color:fontname" e.g. "10:CCCCCC:Helvetica Neue"
                    var legend = chart.GetFirstChild<C.Legend>();
                    if (legend == null) { unsupported.Add(key); break; }
                    legend.RemoveAllChildren<C.TextProperties>();
                    var parts = value.Split(':');
                    var fontSize = parts.Length > 0 && int.TryParse(parts[0], out var fs) ? fs * 100 : 1000;
                    var color = parts.Length > 1 ? parts[1] : null;
                    var fontName = parts.Length > 2 ? parts[2] : null;
                    var defRp = new Drawing.DefaultRunProperties { FontSize = fontSize };
                    if (!string.IsNullOrEmpty(color))
                    {
                        var sf = new Drawing.SolidFill();
                        sf.AppendChild(BuildChartColorElement(color));
                        defRp.AppendChild(sf);
                    }
                    if (!string.IsNullOrEmpty(fontName))
                    {
                        defRp.AppendChild(new Drawing.LatinFont { Typeface = fontName });
                        defRp.AppendChild(new Drawing.EastAsianFont { Typeface = fontName });
                    }
                    legend.AppendChild(new C.TextProperties(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(new Drawing.ParagraphProperties(defRp))
                    ));
                    break;
                }

                case "legend":
                    chart.RemoveAllChildren<C.Legend>();
                    if (!value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
                        !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // CONSISTENCY(strict-enums / R34-1): unknown legend
                        // positions used to silently fall through to "bottom",
                        // producing a contradictory "Updated: legend=hidden"
                        // success message while the file actually carried
                        // legend=bottom. Reject up front with the valid set
                        // so users see typos at Set time.
                        var pos = ParseLegendPosition(value);
                        var plotVisOnly = chart.GetFirstChild<C.PlotVisibleOnly>();
                        var insertBefore = plotVisOnly as OpenXmlElement ?? chart.LastChild;
                        chart.InsertBefore(new C.Legend(
                            new C.LegendPosition { Val = pos },
                            new C.Overlay { Val = false }
                        ), insertBefore);
                    }
                    break;

                case "datalabels" or "labels":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var chartTypeEl in plotArea2.ChildElements
                        .Where(e => e.LocalName.Contains("Chart") || e.LocalName.Contains("chart")))
                    {
                        chartTypeEl.RemoveAllChildren<C.DataLabels>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var dl = new C.DataLabels();
                            // Normalize friendly aliases: seriesName→series, categoryName→category,
                            // percentage→percent. Keeps the dataLabels vocabulary consistent with
                            // the dotted datalabels.show* setter family (see CL15-derived cases below).
                            var partsRaw = value.ToLowerInvariant().Split(',').Select(s => s.Trim()).ToList();
                            for (int pi = 0; pi < partsRaw.Count; pi++)
                            {
                                partsRaw[pi] = partsRaw[pi] switch
                                {
                                    "seriesname" => "series",
                                    "categoryname" => "category",
                                    "percentage" => "percent",
                                    "valuelabel" or "values" => "value",
                                    _ => partsRaw[pi]
                                };
                            }
                            var parts = partsRaw.ToHashSet();
                            // Position values (outsideEnd, center, insideEnd, insideBase, top, bottom, left, right)
                            // implicitly enable showVal when used as the dataLabels value
                            var positionValues = new HashSet<string> { "outsideend", "center", "insideend", "insidebase",
                                "top", "bottom", "left", "right", "bestfit", "t", "b", "l", "r", "outend", "ctr" };
                            var isPositionValue = parts.Any(p => positionValues.Contains(p));
                            var showVal = parts.Contains("value") || parts.Contains("true") || parts.Contains("all") || isPositionValue;
                            dl.AppendChild(new C.ShowLegendKey { Val = false });
                            dl.AppendChild(new C.ShowValue { Val = showVal });
                            dl.AppendChild(new C.ShowCategoryName { Val = parts.Contains("category") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowSeriesName { Val = parts.Contains("series") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowPercent { Val = parts.Contains("percent") || parts.Contains("all") });
                            // If a position value was given, apply it as dLblPos
                            if (isPositionValue)
                            {
                                var posVal = parts.First(p => positionValues.Contains(p));
                                var dLblPos = posVal switch
                                {
                                    "outsideend" or "outend" => C.DataLabelPositionValues.OutsideEnd,
                                    "insideend" => C.DataLabelPositionValues.InsideEnd,
                                    "insidebase" => C.DataLabelPositionValues.InsideBase,
                                    "center" or "ctr" => C.DataLabelPositionValues.Center,
                                    "top" or "t" => C.DataLabelPositionValues.Top,
                                    "bottom" or "b" => C.DataLabelPositionValues.Bottom,
                                    "left" or "l" => C.DataLabelPositionValues.Left,
                                    "right" or "r" => C.DataLabelPositionValues.Right,
                                    "bestfit" => C.DataLabelPositionValues.BestFit,
                                    _ => C.DataLabelPositionValues.OutsideEnd
                                };
                                dl.AppendChild(new C.DataLabelPosition { Val = dLblPos });
                            }
                            // Insert dLbls before dropLines/hiLowLines/upDownBars/gapWidth/overlap/
                            // showMarker/holeSize/firstSliceAngle/axId per schema order. CT_StockChart
                            // and CT_LineChart both place dLbls before dropLines/hiLowLines/upDownBars;
                            // anchoring only on axId would land dLbls after hiLowLines (validator emits
                            // "unexpected child element 'dLbls' ... expected 'axId'").
                            var dlInsertBefore = chartTypeEl.GetFirstChild<C.DropLines>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.HighLowLines>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.UpDownBars>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.GapWidth>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.Overlap>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.ShowMarker>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.HoleSize>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.FirstSliceAngle>() as OpenXmlElement
                                ?? chartTypeEl.GetFirstChild<C.AxisId>();
                            if (dlInsertBefore != null)
                                chartTypeEl.InsertBefore(dl, dlInsertBefore);
                            else
                                chartTypeEl.AppendChild(dl);
                        }
                    }
                    break;
                }

                case "labelpos" or "labelposition":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }

                    // dLblPos is NOT supported by doughnut, area, radar, or stock charts —
                    // CT_DLbls for these chart groups omits dLblPos. Report unsupported so
                    // the caller learns the request didn't land instead of seeing a silent
                    // success ("Updated labelPos=...") with no on-disk change.
                    if (plotArea2.GetFirstChild<C.DoughnutChart>() != null
                        || plotArea2.GetFirstChild<C.AreaChart>() != null
                        || plotArea2.GetFirstChild<C.Area3DChart>() != null
                        || plotArea2.GetFirstChild<C.RadarChart>() != null
                        || plotArea2.GetFirstChild<C.StockChart>() != null)
                    { unsupported.Add(key); break; }

                    // Combo charts (bar+line in same plot area) have incompatible dLblPos
                    // value sets — bar supports inEnd/inBase/outEnd but not t/b/l/r, while
                    // line supports t/b/l/r but not inEnd/inBase/outEnd. Only 'ctr' is
                    // universally valid. Skip entirely for combo charts.
                    var chartGroupCount = plotArea2.ChildElements.Count(
                        e => e is C.BarChart or C.Bar3DChart or C.LineChart or C.Line3DChart
                            or C.ScatterChart or C.BubbleChart);
                    if (chartGroupCount > 1) break;

                    // Pie only supports: bestFit, center, insideEnd, insideBase
                    var isPie = plotArea2.GetFirstChild<C.PieChart>() != null
                        || plotArea2.GetFirstChild<C.Pie3DChart>() != null;

                    // Stacked bar/column/line/area series: ST_DLblPosBar restricts to
                    // {ctr, inBase, inEnd}. Mac PowerPoint reports the file as corrupt
                    // when given outEnd/t/b/l/r/bestFit on a stacked series, even though
                    // OpenXmlValidator schema check passes (the constraint is a
                    // simpleType union, not a structural rule).
                    static bool IsStackedGrouping(EnumValue<C.BarGroupingValues>? g) =>
                        g != null && (g.Value == C.BarGroupingValues.Stacked
                                      || g.Value == C.BarGroupingValues.PercentStacked);
                    static bool IsStackedLineGrouping(EnumValue<C.GroupingValues>? g) =>
                        g != null && (g.Value == C.GroupingValues.Stacked
                                      || g.Value == C.GroupingValues.PercentStacked);
                    var isStacked =
                        plotArea2.Elements<C.BarChart>().Any(c =>
                            IsStackedGrouping(c.GetFirstChild<C.BarGrouping>()?.Val))
                        || plotArea2.Elements<C.Bar3DChart>().Any(c =>
                            IsStackedGrouping(c.GetFirstChild<C.BarGrouping>()?.Val))
                        || plotArea2.Elements<C.LineChart>().Any(c =>
                            IsStackedLineGrouping(c.GetFirstChild<C.Grouping>()?.Val))
                        || plotArea2.Elements<C.Line3DChart>().Any(c =>
                            IsStackedLineGrouping(c.GetFirstChild<C.Grouping>()?.Val));
                        // AreaChart/Area3DChart are not checked here: the
                        // dLblPos handler early-exits for area charts above
                        // (line 256-259), so any area-stacked check below
                        // would be unreachable dead code.

                    // OOXML ST_DLblPosPie restricts pie/pie3D to {bestFit, ctr, inEnd, inBase}.
                    // outEnd/t/b/l/r are not legal here — reject up front instead of
                    // silently remapping to BestFit and reporting "Updated labelPos=...".
                    if (isPie)
                    {
                        var lc = value.ToLowerInvariant();
                        var pieAllowed = lc is "bestfit" or "best" or "auto"
                            or "center" or "ctr"
                            or "insideend" or "inend" or "inside"
                            or "insidebase" or "inbase" or "base";
                        if (!pieAllowed)
                            throw new ArgumentException(
                                $"Invalid labelPos '{value}' for pie chart: ST_DLblPosPie allows only bestFit, ctr, inEnd, inBase.");
                    }
                    var dlblPos = value.ToLowerInvariant() switch
                    {
                        "center" or "ctr" => C.DataLabelPositionValues.Center,
                        "insideend" or "inend" or "inside" => C.DataLabelPositionValues.InsideEnd,
                        "insidebase" or "inbase" or "base" => C.DataLabelPositionValues.InsideBase,
                        "outsideend" or "outend" or "outside" => isStacked
                            ? C.DataLabelPositionValues.InsideEnd
                            : C.DataLabelPositionValues.OutsideEnd,
                        "bestfit" or "best" or "auto" => isStacked
                            ? C.DataLabelPositionValues.Center
                            : C.DataLabelPositionValues.BestFit,
                        "top" or "t" => isStacked
                            ? C.DataLabelPositionValues.InsideEnd
                            : C.DataLabelPositionValues.Top,
                        "bottom" or "b" => isStacked
                            ? C.DataLabelPositionValues.InsideBase
                            : C.DataLabelPositionValues.Bottom,
                        "left" or "l" => isStacked
                            ? C.DataLabelPositionValues.InsideEnd
                            : C.DataLabelPositionValues.Left,
                        "right" or "r" => isStacked
                            ? C.DataLabelPositionValues.InsideEnd
                            : C.DataLabelPositionValues.Right,
                        // Schema enum: {ctr, inBase, inEnd, outEnd, t, b, l, r, bestFit}
                        // plus the long-form aliases handled above. Anything else
                        // is a typo or fuzz garbage — reject up front rather than
                        // silently map to BestFit/OutsideEnd and bury the bug.
                        _ => throw new ArgumentException(
                            $"Invalid labelPos '{value}': expected one of ctr, inBase, inEnd, outEnd, t, b, l, r, bestFit.")
                    };
                    var existingLabels = plotArea2.Descendants<C.DataLabels>().ToList();
                    if (existingLabels.Count == 0)
                    {
                        // Bootstrap charts often lack a c:dLbls element entirely.
                        // Without one, labelPos has nowhere to land and Get sees
                        // nothing — schema declares labelPos get:true so we must
                        // materialize the parent. Attach to the first chart-group
                        // (barChart/lineChart/pieChart/scatterChart/etc.).
                        var chartGroup = plotArea2.ChildElements.OfType<OpenXmlCompositeElement>()
                            .FirstOrDefault(e => e is C.BarChart or C.Bar3DChart
                                or C.LineChart or C.Line3DChart or C.PieChart or C.Pie3DChart
                                or C.ScatterChart or C.BubbleChart);
                        if (chartGroup != null)
                        {
                            var dLbls = new C.DataLabels();
                            // c:dLbls schema requires showLegendKey..showBubbleSize
                            // be present in canonical order; defaults are false.
                            dLbls.AppendChild(new C.ShowLegendKey { Val = false });
                            dLbls.AppendChild(new C.ShowValue { Val = false });
                            dLbls.AppendChild(new C.ShowCategoryName { Val = false });
                            dLbls.AppendChild(new C.ShowSeriesName { Val = false });
                            dLbls.AppendChild(new C.ShowPercent { Val = false });
                            dLbls.AppendChild(new C.ShowBubbleSize { Val = false });
                            chartGroup.PrependChild(dLbls);
                            existingLabels = new List<C.DataLabels> { dLbls };
                        }
                    }
                    foreach (var dl in existingLabels)
                    {
                        dl.RemoveAllChildren<C.DataLabelPosition>();
                        dl.PrependChild(new C.DataLabelPosition { Val = dlblPos });
                    }
                    break;
                }

                case "labelfont":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var dl in plotArea2.Descendants<C.DataLabels>())
                    {
                        dl.RemoveAllChildren<C.TextProperties>();
                        var tp = BuildLabelTextProperties(value);
                        dl.PrependChild(tp);
                    }
                    break;
                }

                case "axisfont" or "axis.font":
                {
                    // Format: "size:color:fontname" e.g. "10:8B949E:Helvetica Neue"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var axis in plotArea2.Elements<C.CategoryAxis>())
                        ApplyAxisTextProperties(axis, value);
                    foreach (var axis in plotArea2.Elements<C.ValueAxis>())
                        ApplyAxisTextProperties(axis, value);
                    foreach (var axis in plotArea2.Elements<C.DateAxis>())
                        ApplyAxisTextProperties(axis, value);
                    break;
                }

                // R15-4: tick-label rotation. Degrees (-90..90). Emits a
                // <c:txPr> with <a:bodyPr rot="deg*60000"/> on the target
                // axis so Excel rotates the tick labels on open.
                case "labelrotation":
                case "xaxis.labelrotation":
                case "xaxislabelrotation":
                case "valaxis.labelrotation":
                case "valaxislabelrotation":
                case "yaxis.labelrotation":
                case "yaxislabelrotation":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var deg))
                    { unsupported.Add(key); break; }
                    var rotAttrVal = ((int)(deg * 60000)).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    var lowerKey = key.ToLowerInvariant();
                    var targetCat = lowerKey is "labelrotation" or "xaxis.labelrotation" or "xaxislabelrotation";
                    var targetVal = lowerKey is "labelrotation" or "valaxis.labelrotation" or "valaxislabelrotation"
                        or "yaxis.labelrotation" or "yaxislabelrotation";
                    if (targetCat)
                    {
                        foreach (var axis in plotArea2.Elements<C.CategoryAxis>())
                            ApplyAxisLabelRotation(axis, rotAttrVal);
                        foreach (var axis in plotArea2.Elements<C.DateAxis>())
                            ApplyAxisLabelRotation(axis, rotAttrVal);
                    }
                    if (targetVal)
                    {
                        foreach (var axis in plotArea2.Elements<C.ValueAxis>())
                            ApplyAxisLabelRotation(axis, rotAttrVal);
                    }
                    break;
                }

                case "colors":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var colorList = value.Split(',').Select(c => c.Trim()).ToArray();

                    // Pie and doughnut charts use VaryColors with dPt elements per data point.
                    // Color per-series is meaningless (only 1 series); color each data point instead.
                    var isPieOrDoughnut = plotArea2.GetFirstChild<C.PieChart>() != null
                        || plotArea2.GetFirstChild<C.DoughnutChart>() != null;
                    if (isPieOrDoughnut)
                    {
                        var ser = plotArea2.Descendants<OpenXmlCompositeElement>()
                            .FirstOrDefault(e => e.LocalName == "ser");
                        if (ser != null)
                        {
                            // Remove existing dPt elements then re-add with new colors
                            var existing = ser.Elements<C.DataPoint>().ToList();
                            foreach (var dp in existing) dp.Remove();

                            for (int ci = 0; ci < colorList.Length; ci++)
                            {
                                var dPt = new C.DataPoint();
                                dPt.AppendChild(new C.Index { Val = (uint)ci });
                                dPt.AppendChild(new C.InvertIfNegative { Val = false });
                                var spPr = new C.ChartShapeProperties();
                                if (colorList[ci].Equals("none", StringComparison.OrdinalIgnoreCase))
                                {
                                    spPr.AppendChild(new Drawing.NoFill());
                                }
                                else
                                {
                                    var solidFill = new Drawing.SolidFill();
                                    solidFill.AppendChild(BuildChartColorElement(colorList[ci]));
                                    spPr.AppendChild(solidFill);
                                }
                                dPt.AppendChild(spPr);

                                // Insert dPt before cat/val data — after Order/SerText/spPr header elements
                                var insertBefore = ser.Elements<C.CategoryAxisData>().FirstOrDefault()
                                    ?? (OpenXmlElement?)ser.Elements<C.Values>().FirstOrDefault()
                                    ?? ser.Elements<C.Explosion>().FirstOrDefault();
                                if (insertBefore != null)
                                    ser.InsertBefore(dPt, insertBefore);
                                else
                                    ser.AppendChild(dPt);
                            }
                        }
                        break;
                    }

                    var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                        .Where(e => e.LocalName == "ser").ToList();
                    for (int ci = 0; ci < Math.Min(colorList.Length, allSer.Count); ci++)
                        ApplySeriesColor(allSer[ci], colorList[ci]);
                    break;
                }

                case "axistitle" or "vtitle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.Title>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var insertAfter = (OpenXmlElement?)valAxis.GetFirstChild<C.MinorGridlines>()
                            ?? (OpenXmlElement?)valAxis.GetFirstChild<C.MajorGridlines>()
                            ?? valAxis.GetFirstChild<C.AxisPosition>();
                        if (insertAfter != null) valAxis.InsertAfter(BuildChartTitle(value), insertAfter);
                    }
                    break;
                }

                case "cattitle" or "htitle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var catAxis = plotArea2?.GetFirstChild<C.CategoryAxis>();
                    if (catAxis == null) { unsupported.Add(key); break; }
                    catAxis.RemoveAllChildren<C.Title>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var insertAfter = (OpenXmlElement?)catAxis.GetFirstChild<C.MinorGridlines>()
                            ?? (OpenXmlElement?)catAxis.GetFirstChild<C.MajorGridlines>()
                            ?? catAxis.GetFirstChild<C.AxisPosition>();
                        if (insertAfter != null) catAxis.InsertAfter(BuildChartTitle(value), insertAfter);
                    }
                    break;
                }

                case "axismin" or "min":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAxis?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.MinAxisValue>();
                    scaling.AppendChild(new C.MinAxisValue { Val = ParseHelpers.SafeParseDouble(value, "axismin") });
                    break;
                }

                case "axismax" or "max":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAxis?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.MaxAxisValue>();
                    var maxEl = new C.MaxAxisValue { Val = ParseHelpers.SafeParseDouble(value, "axismax") };
                    // Schema order: logBase?, orientation, max?, min? — insert max after orientation
                    var orient = scaling.GetFirstChild<C.Orientation>();
                    if (orient != null) orient.InsertAfterSelf(maxEl);
                    else scaling.PrependChild(maxEl);
                    break;
                }

                case "majorunit":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    var mu = ParseHelpers.SafeParseDouble(value, "majorunit");
                    // OOXML ST_AxisUnit: positive double. 0 or negative would
                    // make Excel refuse to draw any tick on the axis (or, in
                    // older builds, freeze the chart). Reject up front instead
                    // of writing garbage that opens to a blank plot area.
                    if (!(mu > 0))
                        throw new ArgumentException($"Invalid majorUnit '{value}': must be a positive number (OOXML ST_AxisUnit > 0).");
                    valAxis.RemoveAllChildren<C.MajorUnit>();
                    InsertValAxChildInOrder(valAxis, new C.MajorUnit { Val = mu });
                    break;
                }

                case "minorunit":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    var nu = ParseHelpers.SafeParseDouble(value, "minorunit");
                    if (!(nu > 0))
                        throw new ArgumentException($"Invalid minorUnit '{value}': must be a positive number (OOXML ST_AxisUnit > 0).");
                    valAxis.RemoveAllChildren<C.MinorUnit>();
                    InsertValAxChildInOrder(valAxis, new C.MinorUnit { Val = nu });
                    break;
                }

                case "axisnumfmt" or "axisnumberformat":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.NumberingFormat>();
                    var nf = new C.NumberingFormat { FormatCode = value, SourceLinked = false };
                    // Schema order: ...title, numFmt, majorTickMark... — insert before majorTickMark
                    var nfInsertBefore = valAxis.GetFirstChild<C.MajorTickMark>();
                    if (nfInsertBefore != null) valAxis.InsertBefore(nf, nfInsertBefore);
                    else valAxis.AppendChild(nf);
                    break;
                }

                case "categories":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var newCats = value.Split(',').Select(c => c.Trim()).ToArray();
                    foreach (var catData in plotArea2.Descendants<C.CategoryAxisData>())
                    {
                        catData.RemoveAllChildren();
                        catData.AppendChild(BuildCategoryData(newCats).FirstChild!.CloneNode(true));
                    }
                    break;
                }

                case "data":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var newSeries = ParseSeriesData(new Dictionary<string, string> { ["data"] = value });
                    UpdateSeriesData(plotArea2, newSeries);
                    break;
                }

                // ---- #2 Gridline styles ----
                case "gridlines" or "majorgridlines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.MajorGridlines>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                        !value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        var gl = new C.MajorGridlines();
                        if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                            gl.AppendChild(BuildLineShapeProperties(value));
                        valAxis.InsertAfter(gl, valAxis.GetFirstChild<C.AxisPosition>());
                    }
                    break;
                }

                case "minorgridlines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.MinorGridlines>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                        !value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        var gl = new C.MinorGridlines();
                        if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                            gl.AppendChild(BuildLineShapeProperties(value));
                        var afterEl = (OpenXmlElement?)valAxis.GetFirstChild<C.MajorGridlines>()
                            ?? valAxis.GetFirstChild<C.AxisPosition>();
                        if (afterEl != null) valAxis.InsertAfter(gl, afterEl);
                    }
                    break;
                }

                // R24 — dotted subkeys mirroring Reader's gridlineColor /
                // gridlineWidth / gridlineDash (and minor* variants). The
                // existing compound "gridlines=color:width:dash" replaces the
                // whole spPr; these subkey paths preserve siblings so users
                // (and dump→replay) can tweak one attribute at a time. Schema
                // already declares get/set: true for these.
                case "gridlinecolor" or "majorgridlinecolor":
                {
                    var gl = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.MajorGridlines>();
                    if (gl == null) { unsupported.Add(key); break; }
                    SetGridlineColor(gl, value);
                    break;
                }
                case "gridlinewidth" or "majorgridlinewidth":
                {
                    var gl = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.MajorGridlines>();
                    if (gl == null) { unsupported.Add(key); break; }
                    if (!SetGridlineWidth(gl, value)) { unsupported.Add(key); }
                    break;
                }
                case "gridlinedash" or "majorgridlinedash":
                {
                    var gl = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.MajorGridlines>();
                    if (gl == null) { unsupported.Add(key); break; }
                    SetGridlineDash(gl, value);
                    break;
                }
                case "minorgridlinecolor":
                {
                    var gl = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.MinorGridlines>();
                    if (gl == null) { unsupported.Add(key); break; }
                    SetGridlineColor(gl, value);
                    break;
                }
                case "minorgridlinewidth":
                {
                    var gl = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.MinorGridlines>();
                    if (gl == null) { unsupported.Add(key); break; }
                    if (!SetGridlineWidth(gl, value)) { unsupported.Add(key); }
                    break;
                }
                case "minorgridlinedash":
                {
                    var gl = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.MinorGridlines>();
                    if (gl == null) { unsupported.Add(key); break; }
                    SetGridlineDash(gl, value);
                    break;
                }

                case "plotareafill" or "plotfill":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    plotArea2.RemoveAllChildren<C.ShapeProperties>();
                    var spPr = new C.ShapeProperties();
                    spPr.AppendChild(BuildFillElement(value));
                    var extLst = plotArea2.GetFirstChild<C.ExtensionList>();
                    if (extLst != null)
                        plotArea2.InsertBefore(spPr, extLst);
                    else
                        plotArea2.AppendChild(spPr);
                    break;
                }

                case "chartareafill" or "chartfill":
                {
                    // After round-trip, SDK may deserialize ChartShapeProperties as ShapeProperties
                    var cSpPr = chartSpace!.GetFirstChild<C.ChartShapeProperties>()
                        ?? (OpenXmlCompositeElement?)chartSpace.GetFirstChild<C.ShapeProperties>();
                    if (cSpPr == null) { cSpPr = new C.ShapeProperties(); chartSpace.InsertAfter(cSpPr, chart); }
                    // Replace fill but keep outline
                    cSpPr.RemoveAllChildren<Drawing.SolidFill>();
                    cSpPr.RemoveAllChildren<Drawing.GradientFill>();
                    cSpPr.RemoveAllChildren<Drawing.NoFill>();
                    cSpPr.PrependChild(BuildFillElement(value));
                    break;
                }

                // ---- #3 Per-series styling ----
                case "linewidth":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var widthEmu = (int)(ParseHelpers.SafeParseDouble(value, "linewidth") * 12700);
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                        ApplySeriesLineWidth(ser, widthEmu);
                    break;
                }

                case "linedash" or "dash":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                        ApplySeriesLineDash(ser, value);
                    break;
                }

                case "marker" or "markers":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        // Schema gate: CT_BarSer / CT_AreaSer / CT_PieSer / CT_BubbleSer
                        // / CT_SurfaceSer have no `c:marker` child. Emitting one
                        // produces a schema-invalid file (Sch_InvalidElementContent...)
                        // that PowerPoint reports as corrupt. Only line/scatter/radar
                        // series accept markers.
                        if (ser is not (C.LineChartSeries or C.ScatterChartSeries or C.RadarChartSeries))
                            continue;
                        ApplySeriesMarker(ser, value);
                    }
                    break;
                }

                case "markersize":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var mSize = ParseHelpers.SafeParseByte(value, "markersize");
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        if (ser is not (C.LineChartSeries or C.ScatterChartSeries or C.RadarChartSeries))
                            continue;
                        var marker = ser.GetFirstChild<C.Marker>();
                        if (marker == null)
                        {
                            // CONSISTENCY(chart/series-schema-order): marker must
                            // precede cat/val/xVal/yVal in CT_LineSer/CT_ScatterSer/
                            // CT_RadarSer. AppendChild lands marker at the tail,
                            // which PowerPoint silently renders but OpenXmlValidator
                            // rejects ('unexpected child element marker').
                            marker = new C.Marker();
                            InsertSeriesChildInOrder(ser, marker);
                        }
                        marker.RemoveAllChildren<C.Size>();
                        marker.AppendChild(new C.Size { Val = mSize });
                    }
                    break;
                }

                case "markercolor":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        if (ser is not (C.LineChartSeries or C.ScatterChartSeries or C.RadarChartSeries))
                            continue;
                        // Reuse the per-series dotted-property handler so
                        // symbol/size are preserved and schema-order insertion
                        // stays in one place.
                        HandleSeriesDottedProperty(ser, "markercolor", value);
                    }
                    break;
                }

                // ---- #4 Chart style ID ----
                case "style" or "styleid":
                {
                    chartSpace!.RemoveAllChildren<C.Style>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var styleVal = ParseHelpers.SafeParseInt(value, "style");
                        if (styleVal < 1 || styleVal > 48)
                            throw new ArgumentException($"Invalid style: '{value}'. Valid range is 1-48.");
                        chartSpace.InsertBefore(new C.Style { Val = (byte)styleVal }, chart);
                    }
                    break;
                }

                // ---- #5 Fill transparency ----
                case "transparency" or "opacity" or "alpha":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var alphaPercent = ParseHelpers.SafeParseDouble(value, key);
                    // If key is "transparency", convert to opacity (e.g. 30% transparency = 70% opacity)
                    if (key.Equals("transparency", StringComparison.OrdinalIgnoreCase))
                        alphaPercent = 100.0 - alphaPercent;
                    var alphaVal = (int)(alphaPercent * 1000); // OOXML uses 1/1000th percent
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                        ApplySeriesAlpha(ser, alphaVal);
                    break;
                }

                // ---- #6 Gradient fill ----
                // CONSISTENCY(gradient-fill-alias): accept `gradientFill=` as an
                // alias for `gradient=` so chart vocabulary matches shape/textbox
                // (ExcelHandler.Add.cs line 1931 / Set.cs line 727 use
                // BuildShapeGradientFill keyed on `gradientFill`).
                case "gradient" or "gradientfill":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    // Format: "color1-color2" or "color1-color2-color3" with optional ":angle"
                    // e.g. "FF0000-0000FF" or "FF0000-00FF00-0000FF:90"
                    var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                        .Where(e => e.LocalName == "ser").ToList();
                    // BUG-R41-B5: a chart with no series (empty/blank chart) used to silently
                    // succeed because the for-loop simply ran 0 iterations — the caller saw
                    // "Updated" while the underlying XML was untouched. Report unsupported
                    // instead so the operator gets a clear signal.
                    if (allSer.Count == 0) { unsupported.Add(key); break; }

                    // R24 — accept boolean forms. "false"/"none" removes the
                    // GradientFill from every series (back to solid). "true"
                    // is the degraded fallback when the dump emitter couldn't
                    // resolve a spec (e.g. theme-color-only stops); fade each
                    // series's solid color to white so dump→replay produces
                    // something visually similar instead of being rejected.
                    if (value.Equals("false", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var ser in allSer)
                        {
                            var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                            spPr?.RemoveAllChildren<Drawing.GradientFill>();
                        }
                        break;
                    }
                    if (value.Equals("true", StringComparison.OrdinalIgnoreCase))
                    {
                        for (int si = 0; si < allSer.Count; si++)
                        {
                            var spPr = allSer[si].GetFirstChild<C.ChartShapeProperties>();
                            var solid = spPr?.GetFirstChild<Drawing.SolidFill>();
                            var baseColor = ReadColorFromFill(solid)?.TrimStart('#')
                                ?? DefaultSeriesColors[si % DefaultSeriesColors.Length];
                            // Fan-out: preserveExisting so per-series gradient set before
                            // chart-level gradient= is not overwritten. Mirrors 2778017a.
                            ApplySeriesGradient(allSer[si], $"{baseColor}-FFFFFF", preserveExisting: true);
                        }
                        break;
                    }
                    for (int si = 0; si < allSer.Count; si++)
                        // Fan-out: preserveExisting so per-series gradient set before
                        // chart-level gradient= is not overwritten. Mirrors 2778017a.
                        ApplySeriesGradient(allSer[si], value, preserveExisting: true);
                    break;
                }

                case "gradients":
                {
                    // Per-series gradients: "FF0000-0000FF,00FF00-FFFF00" (comma-separated, one per series)
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var gradList = value.Split(';').Select(g => g.Trim()).ToArray();
                    var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                        .Where(e => e.LocalName == "ser").ToList();
                    // BUG-R41-B5: same silent-success-on-empty-chart bug as `gradient`.
                    if (allSer.Count == 0) { unsupported.Add(key); break; }
                    for (int si = 0; si < Math.Min(gradList.Length, allSer.Count); si++)
                        ApplySeriesGradient(allSer[si], gradList[si]);
                    break;
                }

                case "view3d" or "camera" or "perspective":
                {
                    // Format: "rotX,rotY,perspective" e.g. "15,20,30" or just "20" for perspective.
                    // Reject named-key form (e.g. "rotX=20,rotY=30") — would silently parse as 0,0,0.
                    if (value.Contains('='))
                    {
                        unsupported.Add(key);
                        break;
                    }
                    // Reject on 2D chart types: PowerPoint accepts the <c:view3D>
                    // tag in OOXML but only renders 3D perspective when the
                    // chartType element is itself 3D (bar3D/column3D/line3D/
                    // pie3D/area3D/surface3D). Silently writing it on a 2D
                    // chart looks fine in Get but renders flat in real PPT.
                    // Hint: switch chartType to a *3D variant.
                    var v3dPlotAreaProbe = chart.GetFirstChild<C.PlotArea>();
                    bool is3DChartType = v3dPlotAreaProbe != null && (
                        v3dPlotAreaProbe.GetFirstChild<C.Bar3DChart>() != null
                        || v3dPlotAreaProbe.GetFirstChild<C.Line3DChart>() != null
                        || v3dPlotAreaProbe.GetFirstChild<C.Area3DChart>() != null
                        || v3dPlotAreaProbe.GetFirstChild<C.Pie3DChart>() != null
                        || v3dPlotAreaProbe.GetFirstChild<C.Surface3DChart>() != null);
                    if (!is3DChartType)
                    {
                        unsupported.Add(key);
                        break;
                    }
                    var v3dParts = value.Split(',');
                    chart.RemoveAllChildren<C.View3D>();
                    var view3d = new C.View3D();
                    if (v3dParts.Length == 1)
                    {
                        // Single value → perspective only (per documented behavior).
                        if (!int.TryParse(v3dParts[0], out var p))
                        {
                            unsupported.Add(key);
                            break;
                        }
                        view3d.AppendChild(new C.Perspective { Val = (byte)p });
                    }
                    else
                    {
                        // Empty slot = field absent in source — do not emit (else
                        // dump→replay introduces phantom rotX=0/rotY=0 children).
                        if (v3dParts.Length >= 1 && !string.IsNullOrWhiteSpace(v3dParts[0])
                            && int.TryParse(v3dParts[0], out var rx))
                            view3d.AppendChild(new C.RotateX { Val = (sbyte)rx });
                        if (v3dParts.Length >= 2 && !string.IsNullOrWhiteSpace(v3dParts[1])
                            && int.TryParse(v3dParts[1], out var ry))
                            view3d.AppendChild(new C.RotateY { Val = (ushort)ry });
                        if (v3dParts.Length >= 3 && !string.IsNullOrWhiteSpace(v3dParts[2])
                            && int.TryParse(v3dParts[2], out var persp))
                            view3d.AppendChild(new C.Perspective { Val = (byte)persp });
                    }
                    // Schema order: title, autoTitleDeleted, pivotFmts, view3D, ..., plotArea
                    var v3dPlotArea = chart.GetFirstChild<C.PlotArea>();
                    if (v3dPlotArea != null) chart.InsertBefore(view3d, v3dPlotArea);
                    else chart.AppendChild(view3d);
                    break;
                }

                case "areafill" or "area.fill":
                {
                    // Apply gradient fill to area chart series. Format: "color1-color2[:angle]"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        var spPr = GetOrCreateSeriesShapeProperties(ser);
                        spPr.RemoveAllChildren<Drawing.SolidFill>();
                        spPr.RemoveAllChildren<Drawing.GradientFill>();
                        // BUG-R33-B3: areafill=none kept appending a fresh
                        // a:noFill on every call, ending up with duplicates
                        // that PowerPoint rejects. Sweep the existing NoFill
                        // before prepending the new fill element.
                        spPr.RemoveAllChildren<Drawing.NoFill>();
                        spPr.PrependChild(BuildFillElement(value));
                    }
                    break;
                }

                // ---- Series visual effects ----
                case "series.shadow" or "seriesshadow":
                {
                    // Apply shadow to all series bars. Format same as shape shadow: "COLOR-BLUR-ANGLE-DIST-OPACITY"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        var spPr = GetOrCreateSeriesShapeProperties(ser);
                        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? new Drawing.EffectList();
                        if (effectList.Parent == null)
                        {
                            // DrawingML spPr schema: ..., ln, effectLst, ... — insert after Outline if present
                            var ln = spPr.GetFirstChild<Drawing.Outline>();
                            if (ln != null) ln.InsertAfterSelf(effectList);
                            else spPr.AppendChild(effectList);
                        }
                        effectList.RemoveAllChildren<Drawing.OuterShadow>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            effectList.AppendChild(DrawingEffectsHelper.BuildOuterShadow(value, BuildChartColorElement));
                    }
                    break;
                }

                case "series.outline" or "seriesoutline":
                {
                    // Apply outline to all series bars. Format: "COLOR" or "COLOR:WIDTH" or "COLOR:WIDTH:DASH"
                    // Also accepts "-" separator for backward compat: "COLOR-WIDTH"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var outParts = value.Contains(':') ? value.Split(':') : value.Split('-');
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        var spPr = GetOrCreateSeriesShapeProperties(ser);
                        spPr.RemoveAllChildren<Drawing.Outline>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var widthPt = outParts.Length > 1 && double.TryParse(outParts[1], System.Globalization.CultureInfo.InvariantCulture, out var w) ? w : 0.5;
                            var outline = new Drawing.Outline { Width = (int)(widthPt * 12700) };
                            var sf = new Drawing.SolidFill();
                            sf.AppendChild(BuildChartColorElement(outParts[0]));
                            outline.AppendChild(sf);
                            if (outParts.Length > 2 && !string.IsNullOrEmpty(outParts[2]))
                                outline.AppendChild(new Drawing.PresetDash { Val = ParseDashStyle(outParts[2]) });
                            // Insert ln before effectLst per DrawingML schema order
                            var effLst = spPr.GetFirstChild<Drawing.EffectList>();
                            if (effLst != null) spPr.InsertBefore(outline, effLst);
                            else spPr.AppendChild(outline);
                        }
                    }
                    break;
                }

                case "gapwidth" or "gap":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    if (!int.TryParse(value, out var gw)) throw new ArgumentException($"Invalid gapWidth: '{value}'. Expected integer (0-500).");
                    bool gapUpdated = false;
                    foreach (var gapEl in plotArea2.Descendants<C.GapWidth>())
                    {
                        gapEl.Val = (ushort)gw;
                        gapUpdated = true;
                    }
                    if (!gapUpdated)
                    {
                        // No existing GapWidth — create one per bar/column chart element.
                        // This occurs when RebuildComboChart (applied via deferred
                        // comboTypes= prop) replaces the original barChart (which had
                        // a GapWidth seeded by BuildBarChart) with freshly constructed
                        // barChart elements that have no GapWidth child. The `foreach
                        // Descendants` above then finds nothing and the gapwidth round-
                        // trips as lost. Mirror the `overlap` upsert pattern.
                        foreach (var barChartEl in plotArea2.Elements<OpenXmlCompositeElement>()
                            .Where(e => e.LocalName == "barChart" || e.LocalName == "bar3DChart"))
                        {
                            // Insert before the first AxisId — mirrors BuildBarChart's
                            // schema order: [barDirection, barGrouping, varyColors,
                            // ser*, gapWidth, overlap?, axisId+].
                            var axisIdEl = barChartEl.GetFirstChild<C.AxisId>();
                            if (axisIdEl != null)
                                axisIdEl.InsertBeforeSelf(new C.GapWidth { Val = (ushort)gw });
                            else
                                barChartEl.AppendChild(new C.GapWidth { Val = (ushort)gw });
                        }
                    }
                    break;
                }

                case "overlap":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    if (!int.TryParse(value, out var ov)) throw new ArgumentException($"Invalid overlap: '{value}'. Expected integer (-100 to 100).");
                    if (ov < -100 || ov > 100) throw new ArgumentException($"Invalid overlap: '{value}'. Valid range is -100 to 100.");
                    foreach (var barChart in plotArea2.Elements<OpenXmlCompositeElement>().Where(e => e.LocalName.Contains("barChart") || e.LocalName.Contains("BarChart")))
                    {
                        var overlapEl = barChart.GetFirstChild<C.Overlap>();
                        if (overlapEl != null) overlapEl.Val = (sbyte)ov;
                        else
                        {
                            var gapEl = barChart.GetFirstChild<C.GapWidth>();
                            if (gapEl != null) gapEl.InsertAfterSelf(new C.Overlap { Val = (sbyte)ov });
                            else barChart.AppendChild(new C.Overlap { Val = (sbyte)ov });
                        }
                    }
                    break;
                }

                // ---- #7 Secondary axis ----
                case "secondaryaxis" or "secondary":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    // R24 — bare "true"/"false" support so the older dump emit
                    // shape (which lost which-series-on-secondary) still
                    // round-trips. "true" routes every non-first series to
                    // the secondary axis (the most common author intent);
                    // "false"/"none" is a no-op when the chart isn't already
                    // split (and a structural rebuild back to single-axis is
                    // out of scope here).
                    HashSet<int> secondaryIndices;
                    if (value.Equals("true", StringComparison.OrdinalIgnoreCase))
                    {
                        var totalSeries = plotArea2.Descendants<OpenXmlCompositeElement>()
                            .Count(e => e.LocalName == "ser");
                        secondaryIndices = new HashSet<int>(Enumerable.Range(2, Math.Max(0, totalSeries - 1)));
                    }
                    else if (value.Equals("false", StringComparison.OrdinalIgnoreCase)
                          || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // No-op: a "no secondary axis" state is the default;
                        // demoting an already-split chart back to single-axis
                        // would require a full rebuild path that doesn't yet
                        // exist. Skip rather than corrupt.
                        break;
                    }
                    else
                    {
                        // value = series indices on secondary axis, e.g. "2,3" (1-based)
                        secondaryIndices = value.Split(',')
                            .Select(s => int.TryParse(s.Trim(), out var v) ? v : -1)
                            .Where(v => v > 0).ToHashSet();
                    }
                    ApplySecondaryAxis(plotArea2, secondaryIndices);
                    break;
                }

                case "plotarea.x" or "plotarea.y" or "plotarea.w" or "plotarea.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal)
                        || !double.IsFinite(layoutVal))
                    { unsupported.Add(key); break; }
                    var plotArea3 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea3 == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(plotArea3, key.Split('.')[1].ToLowerInvariant(), layoutVal, isPlotArea: true);
                    break;
                }

                case "title.x" or "title.y" or "title.w" or "title.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal)
                        || !double.IsFinite(layoutVal))
                    { unsupported.Add(key); break; }
                    var titleEl = chart.GetFirstChild<C.Title>();
                    if (titleEl == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(titleEl, key.Split('.')[1].ToLowerInvariant(), layoutVal);
                    break;
                }

                case "legend.x" or "legend.y" or "legend.w" or "legend.h":
                {
                    // Reject NaN/Infinity — double.TryParse accepts "NaN"/"Infinity"
                    // and the resulting <c:x val="NaN"/> XML breaks Excel.
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal)
                        || !double.IsFinite(layoutVal))
                    { unsupported.Add(key); break; }
                    var legendEl = chart.GetFirstChild<C.Legend>();
                    if (legendEl == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(legendEl, key.Split('.')[1].ToLowerInvariant(), layoutVal);
                    break;
                }

                case "trendlinelabel.x" or "trendlinelabel.y" or "trendlinelabel.w" or "trendlinelabel.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal)
                        || !double.IsFinite(layoutVal))
                    { unsupported.Add(key); break; }
                    var plotArea4 = chart.GetFirstChild<C.PlotArea>();
                    var trendlineLbl = plotArea4?.Descendants<C.TrendlineLabel>().FirstOrDefault();
                    if (trendlineLbl == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(trendlineLbl, key.Split('.')[1].ToLowerInvariant(), layoutVal);
                    break;
                }

                case "displayunitslabel.x" or "displayunitslabel.y" or "displayunitslabel.w" or "displayunitslabel.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal)
                        || !double.IsFinite(layoutVal))
                    { unsupported.Add(key); break; }
                    var dispUnitsLbl = chart.Descendants<C.DisplayUnitsLabel>().FirstOrDefault();
                    if (dispUnitsLbl == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(dispUnitsLbl, key.Split('.')[1].ToLowerInvariant(), layoutVal);
                    break;
                }

                // ==================== Axis Properties ====================

                case "axisvisible" or "axis.visible" or "axis.delete":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var hide = key.Contains("delete") ? ParseHelpers.IsTruthy(value) : !ParseHelpers.IsTruthy(value);
                    foreach (var ax in plotArea2.Elements<C.ValueAxis>())
                    { ax.RemoveAllChildren<C.Delete>(); ax.InsertAfter(new C.Delete { Val = hide }, ax.GetFirstChild<C.Scaling>()); }
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    { ax.RemoveAllChildren<C.Delete>(); ax.InsertAfter(new C.Delete { Val = hide }, ax.GetFirstChild<C.Scaling>()); }
                    break;
                }

                case "cataxisvisible" or "cataxis.visible":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var catAx = plotArea2?.GetFirstChild<C.CategoryAxis>();
                    if (catAx == null) { unsupported.Add(key); break; }
                    catAx.RemoveAllChildren<C.Delete>();
                    catAx.InsertAfter(new C.Delete { Val = !ParseHelpers.IsTruthy(value) }, catAx.GetFirstChild<C.Scaling>());
                    break;
                }

                case "valaxisvisible" or "valaxis.visible":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAx == null) { unsupported.Add(key); break; }
                    valAx.RemoveAllChildren<C.Delete>();
                    valAx.InsertAfter(new C.Delete { Val = !ParseHelpers.IsTruthy(value) }, valAx.GetFirstChild<C.Scaling>());
                    break;
                }

                case "majortickmark" or "majortick":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var tickVal = ParseTickMark(value);
                    foreach (var ax in plotArea2.Elements<C.ValueAxis>())
                    { ax.RemoveAllChildren<C.MajorTickMark>(); InsertAxisChildInOrder(ax, new C.MajorTickMark { Val = tickVal }); }
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    { ax.RemoveAllChildren<C.MajorTickMark>(); InsertAxisChildInOrder(ax, new C.MajorTickMark { Val = tickVal }); }
                    break;
                }

                case "minortickmark" or "minortick":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var tickVal = ParseTickMark(value);
                    foreach (var ax in plotArea2.Elements<C.ValueAxis>())
                    { ax.RemoveAllChildren<C.MinorTickMark>(); InsertAxisChildInOrder(ax, new C.MinorTickMark { Val = tickVal }); }
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    { ax.RemoveAllChildren<C.MinorTickMark>(); InsertAxisChildInOrder(ax, new C.MinorTickMark { Val = tickVal }); }
                    break;
                }

                case "ticklabelpos" or "ticklabelposition":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var tlPos = value.ToLowerInvariant() switch
                    {
                        "none" => C.TickLabelPositionValues.None,
                        "high" or "top" => C.TickLabelPositionValues.High,
                        "low" or "bottom" => C.TickLabelPositionValues.Low,
                        _ => C.TickLabelPositionValues.NextTo
                    };
                    foreach (var ax in plotArea2.Elements<C.ValueAxis>())
                    { ax.RemoveAllChildren<C.TickLabelPosition>(); InsertAxisChildInOrder(ax, new C.TickLabelPosition { Val = tlPos }); }
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    { ax.RemoveAllChildren<C.TickLabelPosition>(); InsertAxisChildInOrder(ax, new C.TickLabelPosition { Val = tlPos }); }
                    break;
                }

                case "axisposition" or "axispos":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var axPos = value.ToLowerInvariant() switch
                    {
                        "top" or "t" => C.AxisPositionValues.Top,
                        "bottom" or "b" => C.AxisPositionValues.Bottom,
                        "left" or "l" => C.AxisPositionValues.Left,
                        "right" or "r" => C.AxisPositionValues.Right,
                        _ => C.AxisPositionValues.Bottom
                    };
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    {
                        ax.RemoveAllChildren<C.AxisPosition>();
                        // CONSISTENCY(chart/axis-schema-order): axPos must sit
                        // immediately after delete in the CT_*Ax prefix; an
                        // AppendChild lands it at the tail and PowerPoint silently
                        // honors a stale axPos while OpenXmlValidator rejects the
                        // file with 'unexpected child element majorTickMark'.
                        InsertAxisChildInOrder(ax, new C.AxisPosition { Val = axPos });
                    }
                    break;
                }

                case "crosses":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAx == null) { unsupported.Add(key); break; }
                    valAx.RemoveAllChildren<C.Crosses>();
                    valAx.RemoveAllChildren<C.CrossesAt>();
                    var crossVal = value.ToLowerInvariant() switch
                    {
                        "max" => C.CrossesValues.Maximum,
                        "min" => C.CrossesValues.Minimum,
                        _ => C.CrossesValues.AutoZero
                    };
                    // CONSISTENCY(chart/crosses-schema-order): CT_ValAx requires
                    // crossAx → crosses → crossBetween. BuildValueAxis emits
                    // CrossBetween last; AppendChild here would land after it and
                    // PowerPoint silently rejects the file. Insert before CrossBetween.
                    var newCrosses = new C.Crosses { Val = crossVal };
                    var cbBefore = valAx.GetFirstChild<C.CrossBetween>();
                    if (cbBefore != null) valAx.InsertBefore(newCrosses, cbBefore);
                    else valAx.AppendChild(newCrosses);
                    break;
                }

                case "crossesat":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAx == null) { unsupported.Add(key); break; }
                    valAx.RemoveAllChildren<C.Crosses>();
                    valAx.RemoveAllChildren<C.CrossesAt>();
                    // CONSISTENCY(chart/crosses-schema-order): same as crosses above.
                    var newCrossesAt = new C.CrossesAt { Val = ParseHelpers.SafeParseDouble(value, "crossesAt") };
                    var cbBefore2 = valAx.GetFirstChild<C.CrossBetween>();
                    if (cbBefore2 != null) valAx.InsertBefore(newCrossesAt, cbBefore2);
                    else valAx.AppendChild(newCrossesAt);
                    break;
                }

                case "crossbetween":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAx == null) { unsupported.Add(key); break; }
                    valAx.RemoveAllChildren<C.CrossBetween>();
                    var cbVal = value.ToLowerInvariant() switch
                    {
                        "midcat" or "midpoint" => C.CrossBetweenValues.MidpointCategory,
                        _ => C.CrossBetweenValues.Between
                    };
                    // CT_ValAx schema: crossAx, crosses?, crossesAt?, crossBetween?,
                    // majorUnit?, minorUnit?, dispUnits?, extLst?. AppendChild lands
                    // it after majorUnit / minorUnit which the validator rejects.
                    var cb = new C.CrossBetween { Val = cbVal };
                    var cbAnchor = valAx.GetFirstChild<C.CrossesAt>() as OpenXmlElement
                        ?? valAx.GetFirstChild<C.Crosses>() as OpenXmlElement
                        ?? valAx.GetFirstChild<C.CrossingAxis>() as OpenXmlElement;
                    if (cbAnchor != null) cbAnchor.InsertAfterSelf(cb);
                    else valAx.AppendChild(cb);
                    break;
                }

                case "axisorientation" or "axisreverse":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAx?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.Orientation>();
                    var orient = (ParseHelpers.IsValidBooleanString(value) && ParseHelpers.IsTruthy(value)) || value.Equals("maxmin", StringComparison.OrdinalIgnoreCase)
                        ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax;
                    scaling.PrependChild(new C.Orientation { Val = orient });
                    break;
                }

                case "logbase" or "logscale" or "yaxisscale":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAx?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.LogBase>();
                    // DEFERRED(xlsx/chart-logscale) CL23: accept `logScale=true`
                    // as shorthand for logBase=10 (Excel's default log base).
                    // `false`/`none` removes the log scale. `logBase=<n>` still
                    // accepts an explicit numeric base via the same key.
                    // R19-2: also accept `yAxisScale=log` / `yAxisScale=linear`
                    // as a verb-style alias. `log` == shorthand for logBase=10,
                    // `linear`/`none` removes the log scale.
                    if (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                        value.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                        value.Equals("log", StringComparison.OrdinalIgnoreCase) ||
                        value == "1")
                    {
                        scaling.PrependChild(new C.LogBase { Val = 10d });
                    }
                    else if (!value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                             !value.Equals("linear", StringComparison.OrdinalIgnoreCase) &&
                             !value.Equals("false", StringComparison.OrdinalIgnoreCase) &&
                             !value.Equals("no", StringComparison.OrdinalIgnoreCase) &&
                             value != "0")
                    {
                        var logVal = ParseHelpers.SafeParseDouble(value, "logBase");
                        // OOXML ST_LogBase: numeric range [2.0, 1000.0]. Values
                        // outside this band produce an unreadable .xlsx (Excel
                        // rewrites the chart back to linear on open and drops
                        // the user's intent silently). Reject up front so the
                        // caller sees the clamp rather than ghost-rewriting.
                        // Truthy/falsy shorthands (true/yes/log/1, false/no/
                        // none/linear/0) are intercepted earlier and don't
                        // reach this branch.
                        // ST_LogBase: minInclusive=2.0, maxExclusive=1000.0.
                        // logBase=1000 itself is rejected (matches Excel's clamp); logBase=2 is the lower bound.
                        if (logVal < 2.0 || logVal >= 1000.0)
                            throw new ArgumentException($"Invalid logBase '{value}': must be in the OOXML range [2, 1000) (ST_LogBase).");
                        scaling.PrependChild(new C.LogBase { Val = logVal });
                    }
                    break;
                }

                case "dispunits" or "displayunits":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAx == null) { unsupported.Add(key); break; }
                    valAx.RemoveAllChildren<C.DisplayUnits>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var builtInVal = value.ToLowerInvariant() switch
                        {
                            "hundreds" => C.BuiltInUnitValues.Hundreds,
                            "thousands" => C.BuiltInUnitValues.Thousands,
                            "tenthousands" or "10000" => C.BuiltInUnitValues.TenThousands,
                            "hundredthousands" or "100000" => C.BuiltInUnitValues.HundredThousands,
                            "millions" => C.BuiltInUnitValues.Millions,
                            "tenmillions" or "10000000" => C.BuiltInUnitValues.TenMillions,
                            "hundredmillions" or "100000000" => C.BuiltInUnitValues.HundredMillions,
                            "billions" => C.BuiltInUnitValues.Billions,
                            "trillions" => C.BuiltInUnitValues.Trillions,
                            _ => throw new ArgumentException(
                                $"Invalid dispUnits '{value}'. Valid values: hundreds, thousands, tenThousands, hundredThousands, millions, tenMillions, hundredMillions, billions, trillions.")
                        };
                        var du = new C.DisplayUnits();
                        du.AppendChild(new C.BuiltInUnit { Val = builtInVal });
                        du.AppendChild(new C.DisplayUnitsLabel());
                        // CONSISTENCY(chart/valAx-schema-order): dispUnits is the
                        // last optional in CT_ValAx (before extLst). AppendChild
                        // is safe only when nothing later already exists; if a
                        // following setter pre-emitted nothing, fine — but if a
                        // later minorUnit Set lands, the helper looks for
                        // dispUnits as the InsertBefore anchor. Route this Set
                        // through the helper to guarantee a stable anchor and
                        // make the order independent of Set sequencing.
                        InsertValAxChildInOrder(valAx, du);
                    }
                    break;
                }

                case "labeloffset":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var catAx = plotArea2?.GetFirstChild<C.CategoryAxis>();
                    if (catAx == null) { unsupported.Add(key); break; }
                    catAx.RemoveAllChildren<C.LabelOffset>();
                    catAx.AppendChild(new C.LabelOffset { Val = (ushort)ParseHelpers.SafeParseInt(value, "labelOffset") });
                    break;
                }

                case "ticklabelskip" or "tickskip":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var catAx = plotArea2?.GetFirstChild<C.CategoryAxis>();
                    if (catAx == null) { unsupported.Add(key); break; }
                    catAx.RemoveAllChildren<C.TickLabelSkip>();
                    catAx.AppendChild(new C.TickLabelSkip { Val = ParseHelpers.SafeParseInt(value, "tickLabelSkip") });
                    break;
                }

                // ==================== Chart-Level Properties ====================

                case "smooth":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var smoothVal = ParseHelpers.IsTruthy(value);
                    bool smoothApplied = false;
                    // Chart-level smooth on LineChart — insert before axId per CT_LineChart schema
                    foreach (var lc in plotArea2.Elements<C.LineChart>())
                    {
                        lc.RemoveAllChildren<C.Smooth>();
                        InsertLineChartChildInOrder(lc, new C.Smooth { Val = smoothVal });
                        smoothApplied = true;
                    }
                    // Also set per-series smooth for line and scatter series
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        if (ser.Parent is C.LineChart or C.ScatterChart)
                        {
                            ser.RemoveAllChildren<C.Smooth>();
                            InsertSeriesChildInOrder(ser, new C.Smooth { Val = smoothVal });
                            smoothApplied = true;
                        }
                    }
                    // BUG-FIX(B5): smooth has no effect on area/bar/column/pie/etc.
                    // Surface as UNSUPPORTED so the caller doesn't think it took.
                    if (!smoothApplied) unsupported.Add(key);
                    break;
                }

                case "showmarker" or "showmarkers":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var showVal = ParseHelpers.IsTruthy(value);
                    foreach (var lc in plotArea2.Elements<C.LineChart>())
                    { lc.RemoveAllChildren<C.ShowMarker>(); InsertLineChartChildInOrder(lc, new C.ShowMarker { Val = showVal }); }
                    // For scatter charts, set per-series marker symbol to none when hiding markers
                    if (!showVal)
                    {
                        foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>()
                            .Where(e => e.LocalName == "ser" && e.Parent is C.ScatterChart))
                        {
                            ser.RemoveAllChildren<C.Marker>();
                            InsertSeriesChildInOrder(ser, new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.None }));
                        }
                    }
                    break;
                }

                case "scatterstyle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var sc = plotArea2?.GetFirstChild<C.ScatterChart>();
                    if (sc == null) { unsupported.Add(key); break; }
                    sc.RemoveAllChildren<C.ScatterStyle>();
                    var ssVal = value.ToLowerInvariant() switch
                    {
                        "line" or "lineonly" => C.ScatterStyleValues.Line,
                        "linemarker" => C.ScatterStyleValues.LineMarker,
                        "marker" or "markeronly" => C.ScatterStyleValues.Marker,
                        "smooth" or "smoothline" => C.ScatterStyleValues.Smooth,
                        "smoothmarker" => C.ScatterStyleValues.SmoothMarker,
                        _ => throw new ArgumentException(
                            $"Invalid scatterStyle '{value}'. Valid values: line, lineMarker, marker, smooth, smoothMarker.")
                    };
                    sc.PrependChild(new C.ScatterStyle { Val = ssVal });
                    break;
                }

                case "varycolors":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var varyVal = ParseHelpers.IsTruthy(value);
                    foreach (var ct in plotArea2.ChildElements
                        .Where(e => e.LocalName.Contains("Chart") || e.LocalName.Contains("chart"))
                        .OfType<OpenXmlCompositeElement>())
                    {
                        ct.RemoveAllChildren<C.VaryColors>();
                        // ECMA-376: in every chart-type element (CT_BarChart,
                        // CT_LineChart, CT_PieChart, CT_AreaChart, ...) varyColors
                        // sits between barDir/grouping (when present) and ser*.
                        // PrependChild lands it before barDir which the validator
                        // rejects with "unexpected child element 'varyColors';
                        // expected: barDir".
                        var vc = new C.VaryColors { Val = varyVal };
                        // Schema: <chartType> element prefix is one of:
                        //   barChart: barDir, grouping, varyColors, ...
                        //   lineChart/areaChart/etc: grouping, varyColors, ...
                        //   radarChart: radarStyle, varyColors, ...
                        //   scatterChart: scatterStyle, varyColors, ...
                        //   pieChart/bubbleChart/doughnutChart: varyColors, ...
                        // Anchor varyColors after whichever predecessor element is present.
                        var anchor = ct.GetFirstChild<C.Grouping>() as OpenXmlElement
                            ?? ct.GetFirstChild<C.BarGrouping>() as OpenXmlElement
                            ?? ct.GetFirstChild<C.BarDirection>() as OpenXmlElement
                            ?? ct.GetFirstChild<C.RadarStyle>() as OpenXmlElement
                            ?? ct.GetFirstChild<C.ScatterStyle>() as OpenXmlElement;
                        if (anchor != null) anchor.InsertAfterSelf(vc);
                        else ct.PrependChild(vc);
                    }
                    break;
                }

                case "dispblanksas" or "blanksas":
                {
                    // CONSISTENCY(strict-enum): reject unknown enum values
                    // instead of silently falling back to Gap. Mirrors R10
                    // conditionalformatting / R11 cf-Add behavior so user
                    // typos surface immediately rather than producing a
                    // silently-different chart.
                    chart.RemoveAllChildren<C.DisplayBlanksAs>();
                    var dbVal = value.ToLowerInvariant() switch
                    {
                        "zero" => C.DisplayBlanksAsValues.Zero,
                        "span" or "connect" => C.DisplayBlanksAsValues.Span,
                        "gap" => C.DisplayBlanksAsValues.Gap,
                        _ => throw new ArgumentException(
                            $"Invalid dispBlanksAs value '{value}'. Allowed: gap, zero, span (alias: connect).")
                    };
                    chart.AppendChild(new C.DisplayBlanksAs { Val = dbVal });
                    break;
                }

                case "roundedcorners":
                {
                    chartSpace!.RemoveAllChildren<C.RoundedCorners>();
                    chartSpace.PrependChild(new C.RoundedCorners { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                case "autotitledeleted":
                {
                    // ECMA-376: <c:autoTitleDeleted> must immediately follow
                    // <c:title> in the c:chart sequence. AppendChild puts it at
                    // the end (after plotArea/legend/plotVisOnly/dispBlanksAs)
                    // which OpenXmlValidator rejects.
                    chart.RemoveAllChildren<C.AutoTitleDeleted>();
                    var atd = new C.AutoTitleDeleted { Val = ParseHelpers.IsTruthy(value) };
                    var atdTitle = chart.GetFirstChild<C.Title>();
                    if (atdTitle != null) chart.InsertAfter(atd, atdTitle);
                    else chart.PrependChild(atd);
                    break;
                }

                case "plotvisonly" or "plotvisibleonly":
                {
                    chart.RemoveAllChildren<C.PlotVisibleOnly>();
                    // CONSISTENCY(chart/chartSpace-schema-order): plotVisOnly must
                    // precede dispBlanksAs/showDLblsOverMax/extLst in CT_Chart.
                    // AppendChild lands it after dispBlanksAs (emitted by the
                    // chart builder), which PowerPoint silently honors but
                    // OpenXmlValidator rejects with 'unexpected child element
                    // plotVisOnly'.
                    InsertChartChildInOrder(chart, new C.PlotVisibleOnly { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                // ==================== Series-Level Properties ====================

                case "trendline":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    // R28-B2: chart-level trendline accepts a semicolon-joined
                    // multi-spec list (e.g. "linear;exp") so dump→replay can
                    // restore series that carried multiple trendlines.
                    var specs = !value.Equals("none", StringComparison.OrdinalIgnoreCase) && value.Contains(';')
                        ? value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        : new[] { value };
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        ser.RemoveAllChildren<C.Trendline>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            foreach (var spec in specs)
                            {
                                var tl = BuildTrendline(spec);
                                InsertSeriesChildInOrder(ser, tl);
                            }
                        }
                    }
                    break;
                }

                case "invertifneg" or "invertifnegative":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var inv = ParseHelpers.IsTruthy(value);
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        ser.RemoveAllChildren<C.InvertIfNegative>();
                        InsertSeriesChildInOrder(ser, new C.InvertIfNegative { Val = inv });
                    }
                    break;
                }

                case "explosion" or "explode":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    // CONSISTENCY(pie-only-prop): c:explosion lives on CT_PieSer
                    // (pie / pie3D / ofPie / doughnut). On bar/line/area the
                    // element is ignored by Excel — make the failure mode loud
                    // so callers see it, matching firstSliceAngle's existing
                    // unsupported-on-non-pie behavior below.
                    var pieOnly = plotArea2.GetFirstChild<C.PieChart>() != null
                                  || plotArea2.GetFirstChild<C.Pie3DChart>() != null
                                  || plotArea2.GetFirstChild<C.OfPieChart>() != null
                                  || plotArea2.GetFirstChild<C.DoughnutChart>() != null;
                    if (!pieOnly) { unsupported.Add(key); break; }
                    var expInt = ParseHelpers.SafeParseInt(value, "explosion");
                    // CT_DLblPercent/CT_UnsignedInt: explosion is non-negative
                    // and reads as a percentage. >100 is technically legal in
                    // OOXML but Excel UI caps at 100; clamp to [0,100] so a
                    // negative cast doesn't underflow to ~4 billion.
                    if (expInt < 0 || expInt > 100)
                        throw new ArgumentException($"Invalid explosion '{value}': must be in [0, 100] (percent).");
                    var expVal = (uint)expInt;
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        ser.RemoveAllChildren<C.Explosion>();
                        if (expVal > 0) InsertSeriesChildInOrder(ser, new C.Explosion { Val = expVal });
                    }
                    break;
                }

                case "errbars" or "errorbars":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        ser.RemoveAllChildren<C.ErrorBars>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase)
                            && SeriesSupportsErrorBars(ser))
                            InsertSeriesChildInOrder(ser, BuildErrorBars(value));
                    }
                    break;
                }

                // CL23 — errBars.direction / errBarDirection controls <c:errBarType val="plus|minus|both"/>.
                // Applied to any existing errBars on all series. If none exist yet, silently no-op
                // (consistency with other per-series options that require the parent prop to be set first).
                case "errbars.direction" or "errbardirection":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var dirVal = value.Trim().ToLowerInvariant() switch
                    {
                        "plus" => C.ErrorBarValues.Plus,
                        "minus" => C.ErrorBarValues.Minus,
                        "both" or "" => C.ErrorBarValues.Both,
                        _ => throw new ArgumentException(
                            $"Invalid errBarDirection '{value}'. Use: plus, minus, both.")
                    };
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        foreach (var eb in ser.Elements<C.ErrorBars>())
                        {
                            eb.RemoveAllChildren<C.ErrorBarType>();
                            // Schema order in CT_ErrBars: errDir, errBarType, errValType, noEndCap, plus, minus, val, spPr
                            var dir = eb.GetFirstChild<C.ErrorDirection>();
                            var newType = new C.ErrorBarType { Val = dirVal };
                            if (dir != null) dir.InsertAfterSelf(newType);
                            else eb.PrependChild(newType);
                        }
                    }
                    break;
                }

                // CL23 — chart-level trendline.* fan-out. Applies the sub-property to every
                // series' existing trendline. Use `series{N}.trendline.{prop}` for per-series.
                case "trendline.label" or "trendline.forecastforward" or "trendline.forecastbackward"
                    or "trendline.order" or "trendline.period"
                    or "trendline.intercept" or "trendline.displayequation" or "trendline.displayrsquared":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var subKey = key.ToLowerInvariant()["trendline.".Length..] switch
                    {
                        "label" => "name",
                        "forecastforward" => "forward",
                        "forecastbackward" => "backward",
                        "order" => "order",
                        "period" => "period",
                        "intercept" => "intercept",
                        "displayequation" => "dispeq",
                        "displayrsquared" => "disprsqr",
                        var s => s
                    };
                    // fuzz-TL01/TL02: validate value before fan-out so invalid
                    // input fails fast even when no series carries a trendline
                    // (otherwise the loop body never runs and bad input is
                    // silently accepted).
                    ValidateTrendlineOptionValue(subKey, value, key);
                    var trendlineTargets = plotArea2.Descendants<OpenXmlCompositeElement>()
                        .Where(e => e.LocalName == "ser")
                        .SelectMany(s => s.Elements<C.Trendline>())
                        .ToList();
                    if (trendlineTargets.Count == 0)
                    {
                        throw new InvalidOperationException(
                            $"{key}: chart has no trendlines to update. " +
                            "Add a trendline first via `series{N}.trendline=linear` (or similar).");
                    }
                    foreach (var tl in trendlineTargets)
                        ApplyTrendlineOptions(tl, subKey, value);
                    break;
                }

                // CL15 — showLeaderLines on pie/doughnut. Alias of datalabels.showleaderlines.
                case "showleaderlines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var dl in plotArea2.Descendants<C.DataLabels>())
                    {
                        dl.RemoveAllChildren<C.ShowLeaderLines>();
                        dl.AppendChild(new C.ShowLeaderLines { Val = show });
                    }
                    break;
                }

                // ==================== DataLabel Enhancements ====================

                case "datalabels.separator" or "labelseparator":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var dl in plotArea2.Descendants<C.DataLabels>())
                    {
                        dl.RemoveAllChildren<C.Separator>();
                        var sep = value.Replace("\\n", "\n");
                        dl.AppendChild(new C.Separator { Text = sep });
                    }
                    break;
                }

                case "datalabels.numfmt" or "labelnumfmt":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var dl in plotArea2.Descendants<C.DataLabels>())
                    {
                        dl.RemoveAllChildren<C.NumberingFormat>();
                        dl.PrependChild(new C.NumberingFormat { FormatCode = value, SourceLinked = false });
                    }
                    break;
                }

                case "datalabels.showleaderlines" or "leaderlines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var dl in plotArea2.Descendants<C.DataLabels>())
                    {
                        dl.RemoveAllChildren<C.ShowLeaderLines>();
                        dl.AppendChild(new C.ShowLeaderLines { Val = show });
                    }
                    break;
                }

                case "datalabels.showbubblesize":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var dl in plotArea2.Descendants<C.DataLabels>())
                    {
                        dl.RemoveAllChildren<C.ShowBubbleSize>();
                        dl.AppendChild(new C.ShowBubbleSize { Val = ParseHelpers.IsTruthy(value) });
                    }
                    break;
                }

                // CleanupE1 — dotted subkeys for toggling individual show* flags on existing
                // dataLabels. Useful for pie charts where `datalabels.showpercent=true` should
                // emit `<c:showPercent val="1"/>` rather than raw values.
                // CONSISTENCY(chart-datalabels-toggle): R28-B1 — accept top-level
                // showValue/showPercent/showCatName/showSerName/showLegendKey
                // aliases (in addition to the dotted datalabels.* form). Pie
                // charts especially want `showPercent=true` as the natural prop.
                case "datalabels.showvalue" or "datalabels.showval"
                    or "showvalue" or "showval":
                {
                    if (!EnsureDataLabelsForShowToggle(chart, key, unsupported, out var dls)) break;
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var dl in dls)
                    {
                        dl.RemoveAllChildren<C.ShowValue>();
                        dl.AppendChild(new C.ShowValue { Val = show });
                    }
                    break;
                }

                case "datalabels.showpercent" or "datalabels.showpct"
                    or "showpercent" or "showpct":
                {
                    if (!EnsureDataLabelsForShowToggle(chart, key, unsupported, out var dls)) break;
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var dl in dls)
                    {
                        dl.RemoveAllChildren<C.ShowPercent>();
                        dl.AppendChild(new C.ShowPercent { Val = show });
                    }
                    break;
                }

                case "datalabels.showcatname" or "datalabels.showcategoryname" or "datalabels.showcategory"
                    or "showcatname" or "showcategoryname" or "showcategory":
                {
                    if (!EnsureDataLabelsForShowToggle(chart, key, unsupported, out var dls)) break;
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var dl in dls)
                    {
                        dl.RemoveAllChildren<C.ShowCategoryName>();
                        dl.AppendChild(new C.ShowCategoryName { Val = show });
                    }
                    break;
                }

                case "datalabels.showsername" or "datalabels.showseriesname" or "datalabels.showseries"
                    or "showsername" or "showseriesname" or "showseries":
                {
                    if (!EnsureDataLabelsForShowToggle(chart, key, unsupported, out var dls)) break;
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var dl in dls)
                    {
                        dl.RemoveAllChildren<C.ShowSeriesName>();
                        dl.AppendChild(new C.ShowSeriesName { Val = show });
                    }
                    break;
                }

                case "datalabels.showlegendkey" or "showlegendkey":
                {
                    if (!EnsureDataLabelsForShowToggle(chart, key, unsupported, out var dls)) break;
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var dl in dls)
                    {
                        dl.RemoveAllChildren<C.ShowLegendKey>();
                        dl.AppendChild(new C.ShowLegendKey { Val = show });
                    }
                    break;
                }

                // ==================== Border / Outline ====================

                case "plotarea.border" or "plotborder":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var spPr = plotArea2.GetFirstChild<C.ShapeProperties>();
                    if (spPr == null) { spPr = new C.ShapeProperties(); plotArea2.AppendChild(spPr); }
                    spPr.RemoveAllChildren<Drawing.Outline>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        spPr.AppendChild(BuildOutlineElement(value));
                    break;
                }

                case "chartarea.border" or "chartborder":
                {
                    var cSpPr = chartSpace!.GetFirstChild<C.ChartShapeProperties>()
                        ?? (OpenXmlCompositeElement?)chartSpace.GetFirstChild<C.ShapeProperties>();
                    if (cSpPr == null) { cSpPr = new C.ShapeProperties(); chartSpace.InsertAfter(cSpPr, chart); }
                    cSpPr.RemoveAllChildren<Drawing.Outline>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        cSpPr.AppendChild(BuildOutlineElement(value));
                    break;
                }

                // ==================== Data Table ====================

                case "datatable":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    plotArea2.RemoveAllChildren<C.DataTable>();
                    if (ParseHelpers.IsTruthy(value))
                    {
                        var dt = new C.DataTable();
                        dt.AppendChild(new C.ShowHorizontalBorder { Val = true });
                        dt.AppendChild(new C.ShowVerticalBorder { Val = true });
                        dt.AppendChild(new C.ShowOutlineBorder { Val = true });
                        dt.AppendChild(new C.ShowKeys { Val = true });
                        plotArea2.AppendChild(dt);
                    }
                    break;
                }

                case "datatable.showhorzborder":
                {
                    var dt = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.DataTable>();
                    if (dt == null) { unsupported.Add(key); break; }
                    dt.RemoveAllChildren<C.ShowHorizontalBorder>();
                    dt.AppendChild(new C.ShowHorizontalBorder { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                case "datatable.showvertborder":
                {
                    var dt = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.DataTable>();
                    if (dt == null) { unsupported.Add(key); break; }
                    dt.RemoveAllChildren<C.ShowVerticalBorder>();
                    dt.AppendChild(new C.ShowVerticalBorder { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                case "datatable.showoutline":
                {
                    var dt = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.DataTable>();
                    if (dt == null) { unsupported.Add(key); break; }
                    dt.RemoveAllChildren<C.ShowOutlineBorder>();
                    dt.AppendChild(new C.ShowOutlineBorder { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                case "datatable.showkeys":
                {
                    var dt = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.DataTable>();
                    if (dt == null) { unsupported.Add(key); break; }
                    dt.RemoveAllChildren<C.ShowKeys>();
                    dt.AppendChild(new C.ShowKeys { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                // ==================== Chart-Type-Specific ====================

                case "firstsliceangle" or "sliceangle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var pie = plotArea2?.GetFirstChild<C.PieChart>();
                    if (pie == null) { unsupported.Add(key); break; }
                    var angInt = ParseHelpers.SafeParseInt(value, "firstSliceAngle");
                    // CT_FirstSliceAng: minInclusive=0, maxInclusive=360.
                    // Negative input would underflow on the ushort cast and
                    // write 65000+, which Excel rewrites silently on open.
                    if (angInt < 0 || angInt > 360)
                        throw new ArgumentException($"Invalid firstSliceAngle '{value}': must be in [0, 360] (degrees).");
                    pie.RemoveAllChildren<C.FirstSliceAngle>();
                    pie.AppendChild(new C.FirstSliceAngle { Val = (ushort)angInt });
                    break;
                }

                case "holesize":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var doughnut = plotArea2?.GetFirstChild<C.DoughnutChart>();
                    if (doughnut == null) { unsupported.Add(key); break; }
                    doughnut.RemoveAllChildren<C.HoleSize>();
                    doughnut.AppendChild(new C.HoleSize { Val = (byte)ParseHelpers.SafeParseInt(value, "holeSize") });
                    break;
                }

                case "radarstyle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var radar = plotArea2?.GetFirstChild<C.RadarChart>();
                    if (radar == null) { unsupported.Add(key); break; }
                    radar.RemoveAllChildren<C.RadarStyle>();
                    var rsVal = value.ToLowerInvariant() switch
                    {
                        "filled" or "fill" => C.RadarStyleValues.Filled,
                        "marker" => C.RadarStyleValues.Marker,
                        "standard" or "line" => C.RadarStyleValues.Standard,
                        _ => throw new ArgumentException(
                            $"Invalid radarStyle '{value}'. Valid values: standard, filled, marker.")
                    };
                    radar.PrependChild(new C.RadarStyle { Val = rsVal });
                    break;
                }

                case "bubblescale":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var bubble = plotArea2?.GetFirstChild<C.BubbleChart>();
                    if (bubble == null) { unsupported.Add(key); break; }
                    bubble.RemoveAllChildren<C.BubbleScale>();
                    InsertBubbleChartChildInOrder(bubble, new C.BubbleScale { Val = (uint)ParseHelpers.SafeParseInt(value, "bubbleScale") });
                    break;
                }

                case "shownegbubbles":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var bubble = plotArea2?.GetFirstChild<C.BubbleChart>();
                    if (bubble == null) { unsupported.Add(key); break; }
                    bubble.RemoveAllChildren<C.ShowNegativeBubbles>();
                    InsertBubbleChartChildInOrder(bubble, new C.ShowNegativeBubbles { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                case "sizerepresents":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var bubble = plotArea2?.GetFirstChild<C.BubbleChart>();
                    if (bubble == null) { unsupported.Add(key); break; }
                    bubble.RemoveAllChildren<C.SizeRepresents>();
                    var srVal = value.ToLowerInvariant() switch
                    {
                        "width" or "w" => C.SizeRepresentsValues.Width,
                        _ => C.SizeRepresentsValues.Area
                    };
                    InsertBubbleChartChildInOrder(bubble, new C.SizeRepresents { Val = srVal });
                    break;
                }

                case "gapdepth":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var target3d = plotArea2?.GetFirstChild<C.Bar3DChart>() as OpenXmlCompositeElement
                        ?? plotArea2?.GetFirstChild<C.Line3DChart>() as OpenXmlCompositeElement
                        ?? plotArea2?.GetFirstChild<C.Area3DChart>() as OpenXmlCompositeElement;
                    if (target3d == null) { unsupported.Add(key); break; }
                    target3d.RemoveAllChildren<C.GapDepth>();
                    InsertBar3DChartChildInOrder(target3d, new C.GapDepth { Val = (ushort)ParseHelpers.SafeParseInt(value, "gapDepth") });
                    break;
                }

                case "shape" or "barshape":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var bar3d = plotArea2?.GetFirstChild<C.Bar3DChart>();
                    if (bar3d == null) { unsupported.Add(key); break; }
                    bar3d.RemoveAllChildren<C.Shape>();
                    var shapeVal = value.ToLowerInvariant() switch
                    {
                        "box" or "cuboid" => C.ShapeValues.Box,
                        "cone" => C.ShapeValues.Cone,
                        "conetomax" => C.ShapeValues.ConeToMax,
                        "cylinder" => C.ShapeValues.Cylinder,
                        "pyramid" => C.ShapeValues.Pyramid,
                        "pyramidtomax" => C.ShapeValues.PyramidToMaximum,
                        _ => throw new ArgumentException(
                            $"Invalid bar shape '{value}'. Valid values: box, cone, coneToMax, cylinder, pyramid, pyramidToMax.")
                    };
                    InsertBar3DChartChildInOrder(bar3d, new C.Shape { Val = shapeVal });
                    break;
                }

                case "droplines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var lc = plotArea2?.GetFirstChild<C.LineChart>();
                    if (lc == null) { unsupported.Add(key); break; }
                    lc.RemoveAllChildren<C.DropLines>();
                    // "false"/"none" remove the overlay; both must skip the
                    // build path. The bool check (IsTruthy) used to gate this,
                    // but a falsy bool like "false" still slipped past
                    // !value.Equals("none") and reached BuildLineShapeProperties,
                    // which threw on the non-spec string.
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase)) break;
                    if ((ParseHelpers.IsValidBooleanString(value) && ParseHelpers.IsTruthy(value)) || !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var dl = new C.DropLines();
                        if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                            dl.AppendChild(BuildLineShapeProperties(value));
                        InsertLineChartChildInOrder(lc, dl);
                    }
                    break;
                }

                case "hilowlines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var lc = plotArea2?.GetFirstChild<C.LineChart>();
                    if (lc == null) { unsupported.Add(key); break; }
                    lc.RemoveAllChildren<C.HighLowLines>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase)) break;
                    if ((ParseHelpers.IsValidBooleanString(value) && ParseHelpers.IsTruthy(value)) || !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var hl = new C.HighLowLines();
                        if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                            hl.AppendChild(BuildLineShapeProperties(value));
                        InsertLineChartChildInOrder(lc, hl);
                    }
                    break;
                }

                case "updownbars":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var lc = plotArea2?.GetFirstChild<C.LineChart>();
                    if (lc == null) { unsupported.Add(key); break; }
                    lc.RemoveAllChildren<C.UpDownBars>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase)) break;
                    // Accept three input shapes: bool ("true"), bare numeric
                    // gapWidth ("150"), and the compound "gap:up:down" — the
                    // Reader emits the compound form, but users can also pass
                    // just the gap width when colors should default.
                    if (value.Contains(':')
                        || (ParseHelpers.IsValidBooleanString(value) && ParseHelpers.IsTruthy(value))
                        || ushort.TryParse(value, out _))
                    {
                        var udb = new C.UpDownBars();
                        ushort gapWidth = 150;
                        string? upColor = null, downColor = null;
                        if (value.Contains(':'))
                        {
                            var udbParts = value.Split(':');
                            if (udbParts.Length >= 1 && ushort.TryParse(udbParts[0], out var gw)) gapWidth = gw;
                            if (udbParts.Length >= 2 && !string.IsNullOrEmpty(udbParts[1])) upColor = udbParts[1];
                            if (udbParts.Length >= 3 && !string.IsNullOrEmpty(udbParts[2])) downColor = udbParts[2];
                        }
                        else if (ushort.TryParse(value, out var bareGw))
                        {
                            gapWidth = bareGw;
                        }
                        udb.AppendChild(new C.GapWidth { Val = gapWidth });
                        var upBars = new C.UpBars();
                        if (upColor != null)
                        {
                            var upSpPr = new C.ChartShapeProperties();
                            var upFill = new Drawing.SolidFill();
                            upFill.AppendChild(BuildChartColorElement(upColor));
                            upSpPr.AppendChild(upFill);
                            upBars.AppendChild(upSpPr);
                        }
                        udb.AppendChild(upBars);
                        var downBars = new C.DownBars();
                        if (downColor != null)
                        {
                            var downSpPr = new C.ChartShapeProperties();
                            var downFill = new Drawing.SolidFill();
                            downFill.AppendChild(BuildChartColorElement(downColor));
                            downSpPr.AppendChild(downFill);
                            downBars.AppendChild(downSpPr);
                        }
                        udb.AppendChild(downBars);
                        InsertLineChartChildInOrder(lc, udb);
                    }
                    break;
                }

                case "serlines" or "serieslines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var show = ParseHelpers.IsTruthy(value);
                    foreach (var barChart in plotArea2.Elements<C.BarChart>())
                    {
                        barChart.RemoveAllChildren<C.SeriesLines>();
                        if (show)
                        {
                            // CT_BarChart schema: ..., gapWidth?, overlap?, serLines?, axId+.
                            // serLines appended after axId is silently dropped by PowerPoint
                            // and flagged by the OOXML validator. Insert before first axId.
                            var sl = new C.SeriesLines();
                            var axIdAnchor = barChart.GetFirstChild<C.AxisId>();
                            if (axIdAnchor != null) barChart.InsertBefore(sl, axIdAnchor);
                            else barChart.AppendChild(sl);
                        }
                    }
                    break;
                }

                // ==================== Axis Line Styling ====================

                case "axisline" or "axis.line":
                {
                    // Style the axis spine line. Format: "color" or "color:width" or "color:width:dash" or "none"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ax in plotArea2.Elements<C.ValueAxis>())
                        ApplyAxisLine(ax, value);
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                        ApplyAxisLine(ax, value);
                    break;
                }

                case "cataxisline" or "cataxis.line":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var catAx = plotArea2?.GetFirstChild<C.CategoryAxis>();
                    if (catAx == null) { unsupported.Add(key); break; }
                    ApplyAxisLine(catAx, value);
                    break;
                }

                case "valaxisline" or "valaxis.line":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAx == null) { unsupported.Add(key); break; }
                    ApplyAxisLine(valAx, value);
                    break;
                }

                // R24 — dotted subkeys mirroring Reader's emit (valAxisLine.color,
                // catAxisLine.width, plotArea.border.dash, chartArea.border.color,
                // …). The existing compound forms above replace the whole outline;
                // these mutate a single attribute and keep siblings intact, so
                // dump→replay can round-trip an OOXML outline that was authored
                // attribute-by-attribute.
                case "valaxisline.color" or "valaxisline.width" or "valaxisline.dash":
                {
                    var ax = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.ValueAxis>();
                    if (ax == null) { unsupported.Add(key); break; }
                    if (!MutateAxisLineAttr(ax, key.Substring("valaxisline.".Length), value))
                        unsupported.Add(key);
                    break;
                }
                case "cataxisline.color" or "cataxisline.width" or "cataxisline.dash":
                {
                    var ax = chart.GetFirstChild<C.PlotArea>()?.GetFirstChild<C.CategoryAxis>();
                    if (ax == null) { unsupported.Add(key); break; }
                    if (!MutateAxisLineAttr(ax, key.Substring("cataxisline.".Length), value))
                        unsupported.Add(key);
                    break;
                }
                case "plotarea.border.color" or "plotarea.border.width" or "plotarea.border.dash":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var spPr = plotArea2.GetFirstChild<C.ShapeProperties>();
                    if (spPr == null) { spPr = new C.ShapeProperties(); plotArea2.AppendChild(spPr); }
                    if (!MutateOutlineAttr(spPr, key.Substring("plotarea.border.".Length), value))
                        unsupported.Add(key);
                    break;
                }
                case "chartarea.border.color" or "chartarea.border.width" or "chartarea.border.dash":
                {
                    var cSpPr = chartSpace!.GetFirstChild<C.ChartShapeProperties>()
                        ?? (OpenXmlCompositeElement?)chartSpace.GetFirstChild<C.ShapeProperties>();
                    if (cSpPr == null) { cSpPr = new C.ShapeProperties(); chartSpace.InsertAfter(cSpPr, chart); }
                    if (!MutateOutlineAttr(cSpPr, key.Substring("chartarea.border.".Length), value))
                        unsupported.Add(key);
                    break;
                }

                // ==================== Advanced Features ====================

                case "referenceline" or "refline" or "targetline":
                {
                    // Format: "value" or "value:color" or "value:color:label" or
                    // "value:color:label:dash". Multiple lines = semicolon-joined
                    // (matches Reader output format).
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                        if (plotArea2 != null)
                            RemoveExistingReferenceLines(plotArea2);
                        break;
                    }
                    var specs = value.Split(';', StringSplitOptions.RemoveEmptyEntries);
                    // Remove existing once, then add each spec without further
                    // sweeps so multi-line input accumulates instead of leaving
                    // only the last one (which broke dump→replay of charts with
                    // 2+ reference lines).
                    var plotAreaRl = chart.GetFirstChild<C.PlotArea>();
                    if (plotAreaRl != null)
                        RemoveExistingReferenceLines(plotAreaRl);
                    foreach (var spec in specs)
                        AddReferenceLine(chart, spec.Trim(), removeExisting: false);
                    break;
                }

                case "colorrule" or "colorRule" or "conditionalcolor":
                {
                    // Format: "threshold:belowColor:aboveColor" e.g. "0:FF0000:00AA00"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    ApplyColorRule(plotArea2, value);
                    break;
                }

                case "combotypes" or "combo.types":
                {
                    // Format: "column,column,line,area" — per-series chart type
                    RebuildComboChart(chart, value);
                    break;
                }

                // ==================== Legend Enhancements ====================

                case "legend.overlay" or "legendoverlay":
                {
                    var legendEl = chart.GetFirstChild<C.Legend>();
                    if (legendEl == null) { unsupported.Add(key); break; }
                    legendEl.RemoveAllChildren<C.Overlay>();
                    legendEl.AppendChild(new C.Overlay { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                // CONSISTENCY(rtl-cascade): chart-level reading direction.
                // Stamps rtl="1" on chartSpace c:txPr → a:lstStyle a:lvl1pPr
                // so default chart text bodies (axis labels, data labels)
                // render right-to-left for Arabic / Hebrew. Mirrors the
                // direction surface on shapes/textboxes.
                case "direction" or "rtl":
                {
                    bool rtlOn = value.ToLowerInvariant() switch
                    {
                        "rtl" or "righttoleft" or "right-to-left" or "true" or "1" => true,
                        "ltr" or "lefttoright" or "left-to-right" or "false" or "0" or "" => false,
                        _ => throw new ArgumentException(
                            $"Invalid direction value: '{value}'. Valid values: rtl, ltr (also accepts true/false, 1/0, righttoleft/lefttoright, right-to-left/left-to-right; case-insensitive).")
                    };
                    var txPr = chartSpace!.GetFirstChild<C.TextProperties>();
                    if (txPr == null)
                    {
                        txPr = new C.TextProperties(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(new Drawing.EndParagraphRunProperties { Language = "en-US" }));
                        chartSpace.AppendChild(txPr);
                    }
                    var lstStyle = txPr.GetFirstChild<Drawing.ListStyle>()
                        ?? txPr.AppendChild(new Drawing.ListStyle());
                    var lvl1 = lstStyle.GetFirstChild<Drawing.Level1ParagraphProperties>();
                    if (lvl1 == null)
                    {
                        lvl1 = new Drawing.Level1ParagraphProperties();
                        lstStyle.AppendChild(lvl1);
                    }
                    lvl1.RightToLeft = rtlOn;

                    // CONSISTENCY(rtl-cascade): axis-level c:txPr overrides
                    // chartSpace c:txPr in OOXML, so direction must propagate
                    // into every per-axis (catAx/valAx/serAx/dateAx) and
                    // dLbls c:txPr that exists. Without this, Arabic axis
                    // labels render LTR even when chart direction=rtl is set.
                    static void StampLvl1Rtl(C.TextProperties tp, bool on)
                    {
                        var ls = tp.GetFirstChild<Drawing.ListStyle>()
                            ?? tp.AppendChild(new Drawing.ListStyle());
                        var l1 = ls.GetFirstChild<Drawing.Level1ParagraphProperties>();
                        if (l1 == null)
                        {
                            l1 = new Drawing.Level1ParagraphProperties();
                            ls.AppendChild(l1);
                        }
                        l1.RightToLeft = on;
                    }
                    var plotAreaRtl = chart.GetFirstChild<C.PlotArea>();
                    if (plotAreaRtl != null)
                    {
                        foreach (var axisTxPr in plotAreaRtl.Descendants<C.TextProperties>().ToList())
                            StampLvl1Rtl(axisTxPr, rtlOn);
                    }
                    // Legend is a *sibling* of plotArea (direct child of c:chart),
                    // not a descendant — walk its c:txPr explicitly.
                    var legendRtl = chart.GetFirstChild<C.Legend>();
                    if (legendRtl != null)
                    {
                        foreach (var legTxPr in legendRtl.Descendants<C.TextProperties>().ToList())
                            StampLvl1Rtl(legTxPr, rtlOn);
                    }
                    // Chart-level c:dLbls (sibling of plotArea on certain chart types).
                    var chartDLblsRtl = chart.GetFirstChild<C.DataLabels>();
                    if (chartDLblsRtl != null)
                    {
                        foreach (var dlTxPr in chartDLblsRtl.Descendants<C.TextProperties>().ToList())
                            StampLvl1Rtl(dlTxPr, rtlOn);
                    }
                    // Title rich text: walk c:title/c:tx/c:rich a:lstStyle a:lvl1pPr.
                    var titleEl = chart.GetFirstChild<C.Title>();
                    var titleRich = titleEl?.ChartText?.RichText;
                    if (titleRich != null)
                    {
                        var tLst = titleRich.GetFirstChild<Drawing.ListStyle>()
                            ?? titleRich.AppendChild(new Drawing.ListStyle());
                        var tLvl1 = tLst.GetFirstChild<Drawing.Level1ParagraphProperties>();
                        if (tLvl1 == null)
                        {
                            tLvl1 = new Drawing.Level1ParagraphProperties();
                            tLst.AppendChild(tLvl1);
                        }
                        tLvl1.RightToLeft = rtlOn;
                    }
                    break;
                }

                default:
                    // dataLabel{N}.{x|y|w|h} — individual data label layout (1-based point index, first series)
                    if (TryParseDataLabelLayoutKey(key, out var dlPointIdx, out var dlProp))
                    {
                        if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var dlLayoutVal))
                        { unsupported.Add(key); break; }
                        var plotArea5 = chart.GetFirstChild<C.PlotArea>();
                        var firstSer = plotArea5?.Descendants<OpenXmlCompositeElement>()
                            .FirstOrDefault(e => e.LocalName == "ser");
                        if (firstSer == null) { unsupported.Add(key); break; }
                        var dLbls = firstSer.GetFirstChild<C.DataLabels>();
                        if (dLbls == null)
                        {
                            // Create minimal DataLabels container with ShowValue=true
                            dLbls = new C.DataLabels();
                            dLbls.AppendChild(new C.ShowLegendKey { Val = false });
                            dLbls.AppendChild(new C.ShowValue { Val = true });
                            dLbls.AppendChild(new C.ShowCategoryName { Val = false });
                            dLbls.AppendChild(new C.ShowSeriesName { Val = false });
                            dLbls.AppendChild(new C.ShowPercent { Val = false });
                            InsertSeriesChildInOrder(firstSer, dLbls);
                        }
                        // Find or create individual dLbl for the point index (0-based in OOXML)
                        var ooxmlIdx = (uint)(dlPointIdx - 1);
                        var dLbl = dLbls.Elements<C.DataLabel>()
                            .FirstOrDefault(dl => dl.Index?.Val?.Value == ooxmlIdx);
                        if (dLbl == null)
                        {
                            dLbl = new C.DataLabel();
                            dLbl.Index = new C.Index { Val = ooxmlIdx };
                            // Insert dLbl before the show* elements (dLbl comes before showLegendKey per schema)
                            var insertBefore = dLbls.GetFirstChild<C.ShowLegendKey>() as OpenXmlElement
                                ?? dLbls.GetFirstChild<C.ShowValue>()
                                ?? dLbls.FirstChild;
                            if (insertBefore != null)
                                dLbls.InsertBefore(dLbl, insertBefore);
                            else
                                dLbls.AppendChild(dLbl);
                        }
                        SetManualLayoutProperty(dLbl, dlProp, dlLayoutVal);
                        break;
                    }
                    // Per-series dotted keys: series{N}.smooth, series{N}.trendline, series{N}.point{M}.color, etc.
                    if (TryParseSeriesDottedKey(key, out var sIdx, out var sProp))
                    {
                        var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                        if (plotArea2 == null) { unsupported.Add(key); break; }
                        var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                            .Where(e => e.LocalName == "ser").ToList();
                        if (sIdx < 1 || sIdx > allSer.Count) { unsupported.Add(key); break; }
                        var ser = allSer[sIdx - 1];
                        if (!HandleSeriesDottedProperty(ser, sProp, value))
                            unsupported.Add(key);
                        break;
                    }
                    // dataLabel{N}.delete / dataLabel{N}.pos
                    if (TryParseDataLabelDottedKey(key, out var dlIdx2, out var dlProp2))
                    {
                        var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                        var firstSer2 = plotArea2?.Descendants<OpenXmlCompositeElement>()
                            .FirstOrDefault(e => e.LocalName == "ser");
                        if (firstSer2 == null) { unsupported.Add(key); break; }
                        HandleDataLabelDottedProperty(firstSer2, dlIdx2, dlProp2, value);
                        break;
                    }
                    // legendEntry{N}.delete
                    if (TryParseLegendEntryKey(key, out var leIdx))
                    {
                        var legendEl = chart.GetFirstChild<C.Legend>();
                        if (legendEl == null) { unsupported.Add(key); break; }
                        var existingEntry = legendEl.Elements<C.LegendEntry>()
                            .FirstOrDefault(le => le.Index?.Val?.Value == (uint)(leIdx - 1));
                        if (existingEntry != null) existingEntry.Remove();
                        if (ParseHelpers.IsTruthy(value))
                        {
                            var le = new C.LegendEntry();
                            le.AppendChild(new C.Index { Val = (uint)(leIdx - 1) });
                            le.AppendChild(new C.Delete { Val = true });
                            // CT_Legend schema order: legendPos, legendEntry+, layout, overlay, spPr, txPr
                            // Insert after legendPos (or at start if no legendPos), before overlay/layout
                            var legendPos2 = legendEl.GetFirstChild<C.LegendPosition>();
                            if (legendPos2 != null)
                                legendPos2.InsertAfterSelf(le);
                            else
                                legendEl.PrependChild(le);
                        }
                        break;
                    }
                    // Legacy: series{N} = "Name:1,2,3" (numeric data update)
                    if (key.StartsWith("series", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(key[6..], out var seriesIdx))
                    {
                        var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                        if (plotArea2 == null) { unsupported.Add(key); break; }
                        var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                            .Where(e => e.LocalName == "ser").ToList();
                        if (seriesIdx < 1 || seriesIdx > allSer.Count) { unsupported.Add(key); break; }
                        var ser = allSer[seriesIdx - 1];

                        var colonIdx = value.IndexOf(':');
                        double[] vals;
                        if (colonIdx >= 0)
                        {
                            var sName = value[..colonIdx].Trim();
                            vals = ParseSeriesValues(value[(colonIdx + 1)..], value[..colonIdx].Trim());
                            var serText = ser.GetFirstChild<C.SeriesText>();
                            if (serText != null)
                            {
                                serText.RemoveAllChildren();
                                serText.AppendChild(new C.NumericValue(sName));
                            }
                        }
                        else
                        {
                            vals = ParseSeriesValues(value, "series data");
                        }

                        var valEl = ser.GetFirstChild<C.Values>();
                        if (valEl != null)
                        {
                            valEl.RemoveAllChildren();
                            var builtVals = BuildValues(vals);
                            foreach (var child in builtVals.ChildElements.ToList())
                                valEl.AppendChild(child.CloneNode(true));
                        }
                        var yValEl = ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal");
                        if (yValEl != null)
                        {
                            yValEl.RemoveAllChildren();
                            var numLit = new C.NumberLiteral(
                                new C.FormatCode("General"),
                                new C.PointCount { Val = (uint)vals.Length });
                            for (int vi = 0; vi < vals.Length; vi++)
                                numLit.AppendChild(new C.NumericPoint(new C.NumericValue(vals[vi].ToString("G"))) { Index = (uint)vi });
                            yValEl.AppendChild(numLit);
                        }
                    }
                    else
                    {
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid chart props: title, legend, dataLabels, labelPos, labelFont, " +
                              "axisFont, axisTitle, catTitle, axisMin, axisMax, majorUnit, minorUnit, axisNumFmt, " +
                              "axisVisible, majorTickMark, minorTickMark, tickLabelPos, crosses, crossBetween, " +
                              "axisOrientation, logBase, dispUnits, gridlines, minorGridlines, " +
                              "plotFill, chartFill, plotArea.border, chartArea.border, " +
                              "colors, gradient, lineWidth, lineDash, marker, markerSize, transparency, " +
                              "smooth, showMarker, scatterStyle, varyColors, dispBlanksAs, dataTable, " +
                              "trendline, errBars, explosion, invertIfNeg, gapWidth, overlap, secondaryAxis, " +
                              "firstSliceAngle, holeSize, radarStyle, bubbleScale, shape, " +
                              "roundedCorners, legend.overlay, view3d, categories, data, " +
                              "plotArea.x/y/w/h, title.x/y/w/h, legend.x/y/w/h, " +
                              "series{N}=Name:1,2,3, series{N}.smooth/trendline/color/point{M}.color)"
                            : key);
                    }
                    break;
            }
        }

        chartSpace!.Save();
        return unsupported;
    }

    // ==================== #1 Data Label Helpers ====================

    /// <summary>
    /// Build text properties for data labels: "size:color:bold" e.g. "10:FF0000:true" or just "10"
    /// </summary>
    private static C.TextProperties BuildLabelTextProperties(string spec)
    {
        var parts = spec.Split(':');
        var fontSize = parts.Length > 0 && int.TryParse(parts[0], out var fs) ? fs * 100 : 1000;
        var color = parts.Length > 1 ? parts[1] : null;
        var bold = parts.Length > 2 && parts[2].Equals("true", StringComparison.OrdinalIgnoreCase);

        var defRp = new Drawing.DefaultRunProperties { FontSize = fontSize, Bold = bold };
        if (!string.IsNullOrEmpty(color))
        {
            var solidFill = new Drawing.SolidFill();
            solidFill.AppendChild(BuildChartColorElement(color));
            defRp.AppendChild(solidFill);
        }

        return new C.TextProperties(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.ParagraphProperties(defRp))
        );
    }

    // ==================== #2 Gridline / Shape Property Helpers ====================

    /// <summary>
    /// Build shape properties for gridlines/outlines. Format: "color" or "color:widthPt" or "color:widthPt:dash"
    /// e.g. "CCCCCC", "CCCCCC:0.5", "CCCCCC:1:dash"
    /// </summary>
    private static C.ChartShapeProperties BuildLineShapeProperties(string spec)
    {
        var parts = spec.Split(':');
        var color = parts[0].Trim();
        var widthPt = parts.Length > 1 && double.TryParse(parts[1], System.Globalization.CultureInfo.InvariantCulture, out var w) ? w : 0.5;
        var dash = parts.Length > 2 ? parts[2].Trim() : null;

        var outline = new Drawing.Outline { Width = (int)(widthPt * 12700) };
        var solidFill = new Drawing.SolidFill();
        solidFill.AppendChild(BuildChartColorElement(color));
        outline.AppendChild(solidFill);

        if (!string.IsNullOrEmpty(dash))
        {
            var dashVal = ParseDashStyle(dash);
            outline.AppendChild(new Drawing.PresetDash { Val = dashVal });
        }

        var spPr = new C.ChartShapeProperties();
        spPr.AppendChild(outline);
        return spPr;
    }

    /// <summary>
    /// Get or create the <c:spPr>/<a:ln> outline element on a gridline so a
    /// single attribute (color / width / dash) can be replaced without
    /// touching its siblings. Mirrors `<c:spPr><a:ln>…</a:ln></c:spPr>`.
    /// </summary>
    private static Drawing.Outline GetOrCreateGridlineOutline(OpenXmlCompositeElement gridlines)
    {
        var spPr = gridlines.GetFirstChild<C.ChartShapeProperties>();
        if (spPr == null)
        {
            spPr = new C.ChartShapeProperties();
            gridlines.AppendChild(spPr);
        }
        var outline = spPr.GetFirstChild<Drawing.Outline>();
        if (outline == null)
        {
            outline = new Drawing.Outline();
            spPr.AppendChild(outline);
        }
        return outline;
    }

    private static void SetGridlineColor(OpenXmlCompositeElement gridlines, string color)
    {
        var outline = GetOrCreateGridlineOutline(gridlines);
        outline.RemoveAllChildren<Drawing.SolidFill>();
        outline.RemoveAllChildren<Drawing.NoFill>();
        var fill = new Drawing.SolidFill();
        fill.AppendChild(BuildChartColorElement(color));
        // Outline schema: fill children precede PresetDash.
        var dash = outline.GetFirstChild<Drawing.PresetDash>();
        if (dash != null) outline.InsertBefore(fill, dash);
        else outline.PrependChild(fill);
    }

    private static bool SetGridlineWidth(OpenXmlCompositeElement gridlines, string value)
    {
        if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var widthPt))
            return false;
        var outline = GetOrCreateGridlineOutline(gridlines);
        outline.Width = (int)(widthPt * 12700);
        return true;
    }

    private static void SetGridlineDash(OpenXmlCompositeElement gridlines, string dash)
    {
        var outline = GetOrCreateGridlineOutline(gridlines);
        outline.RemoveAllChildren<Drawing.PresetDash>();
        outline.AppendChild(new Drawing.PresetDash { Val = ParseDashStyle(dash) });
    }

    /// <summary>
    /// Mutate one of color/width/dash on the <a:ln> child of an axis's spPr.
    /// Preserves the other two attributes. Returns false if attr is unknown.
    /// </summary>
    private static bool MutateAxisLineAttr(OpenXmlCompositeElement axis, string attr, string value)
    {
        var spPr = axis.GetFirstChild<C.ChartShapeProperties>();
        if (spPr == null)
        {
            spPr = new C.ChartShapeProperties();
            var tlPos = axis.GetFirstChild<C.TickLabelPosition>();
            if (tlPos != null) tlPos.InsertAfterSelf(spPr);
            else axis.AppendChild(spPr);
        }
        return MutateOutlineAttr(spPr, attr, value);
    }

    /// <summary>
    /// Mutate one of color/width/dash on the <a:ln> child of a spPr, preserving
    /// the other two. Shared by axisLine.* / plotArea.border.* / chartArea.border.*.
    /// </summary>
    private static bool MutateOutlineAttr(OpenXmlCompositeElement spPr, string attr, string value)
    {
        var outline = spPr.GetFirstChild<Drawing.Outline>();
        if (outline == null)
        {
            outline = new Drawing.Outline();
            spPr.AppendChild(outline);
        }
        // Drop NoFill if present — we're populating a real attribute now.
        outline.RemoveAllChildren<Drawing.NoFill>();

        switch (attr.ToLowerInvariant())
        {
            case "color":
            {
                outline.RemoveAllChildren<Drawing.SolidFill>();
                var fill = new Drawing.SolidFill();
                fill.AppendChild(BuildChartColorElement(value));
                var dashEl = outline.GetFirstChild<Drawing.PresetDash>();
                if (dashEl != null) outline.InsertBefore(fill, dashEl);
                else outline.PrependChild(fill);
                return true;
            }
            case "width":
            {
                if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var widthPt))
                    return false;
                outline.Width = (int)(widthPt * 12700);
                return true;
            }
            case "dash":
            {
                outline.RemoveAllChildren<Drawing.PresetDash>();
                outline.AppendChild(new Drawing.PresetDash { Val = ParseDashStyle(value) });
                return true;
            }
            default:
                return false;
        }
    }

    private static Drawing.PresetLineDashValues ParseDashStyle(string dash)
    {
        // CONSISTENCY(ooxml-dash-aliases): accept both the legacy snake_case
        // form (sysdash_dot) AND the OOXML-native camelCase form
        // (sysDashDot / lgDashDot / lgDashDotDot / sysDashDotDot) — the
        // Reader emits the camelCase form to mirror the schema spelling,
        // so dump→replay would otherwise hit the `_ => Solid` fallback.
        return dash.ToLowerInvariant() switch
        {
            "solid" => Drawing.PresetLineDashValues.Solid,
            "dot" => Drawing.PresetLineDashValues.Dot,
            "sysdot" => Drawing.PresetLineDashValues.SystemDot,
            "dash" => Drawing.PresetLineDashValues.Dash,
            "sysdash" => Drawing.PresetLineDashValues.SystemDash,
            "dashdot" => Drawing.PresetLineDashValues.DashDot,
            // dashDotDot has no native CT_PresetLineDashValues enum member —
            // ECMA-376 (DrawingML) defines only dash/dashDot plus the sys*/lg*
            // dot-dot variants. Tolerate the natural extrapolation as an
            // explicit alias for sysDashDotDot (the closest visual match)
            // rather than silently falling through to Solid. Documented in
            // schemas/help/_shared/chart-series.json lineDash and pptx
            // shape/connector lineDash so users know Get readback will be
            // sysDashDotDot, not dashDotDot.
            "dashdotdot" => Drawing.PresetLineDashValues.SystemDashDotDot,
            "sysdashdot" or "sysdash_dot" => Drawing.PresetLineDashValues.SystemDashDot,
            "sysdashdotdot" or "sysdash_dot_dot" => Drawing.PresetLineDashValues.SystemDashDotDot,
            "longdash" or "lgdash" => Drawing.PresetLineDashValues.LargeDash,
            "longdashdot" or "lgdashdot" => Drawing.PresetLineDashValues.LargeDashDot,
            "longdashdotdot" or "lgdashdotdot" => Drawing.PresetLineDashValues.LargeDashDotDot,
            _ => Drawing.PresetLineDashValues.Solid
        };
    }

    // ==================== #3 Per-Series Style Helpers ====================

    private static C.ChartShapeProperties GetOrCreateSeriesShapeProperties(OpenXmlCompositeElement series)
    {
        var spPr = series.GetFirstChild<C.ChartShapeProperties>();
        if (spPr != null) return spPr;
        spPr = new C.ChartShapeProperties();
        var serText = series.GetFirstChild<C.SeriesText>();
        if (serText != null) serText.InsertAfterSelf(spPr);
        else series.PrependChild(spPr);
        return spPr;
    }

    internal static void ApplySeriesLineWidth(OpenXmlCompositeElement series, int widthEmu)
    {
        var spPr = GetOrCreateSeriesShapeProperties(series);
        var outline = spPr.GetFirstChild<Drawing.Outline>();
        if (outline == null) { outline = new Drawing.Outline(); spPr.AppendChild(outline); }
        // BuildScatterChart pre-seeds NoFill for marker-only series; we must
        // drop it before assigning a real width — see outlineColor case.
        outline.RemoveAllChildren<Drawing.NoFill>();
        outline.Width = widthEmu;
    }

    internal static void ApplySeriesLineDash(OpenXmlCompositeElement series, string dashStyle)
    {
        var spPr = GetOrCreateSeriesShapeProperties(series);
        var outline = spPr.GetFirstChild<Drawing.Outline>();
        if (outline == null) { outline = new Drawing.Outline(); spPr.AppendChild(outline); }
        outline.RemoveAllChildren<Drawing.NoFill>();
        outline.RemoveAllChildren<Drawing.PresetDash>();
        outline.AppendChild(new Drawing.PresetDash { Val = ParseDashStyle(dashStyle) });
    }

    internal static void ApplySeriesMarker(OpenXmlCompositeElement series, string markerSpec)
    {
        // Format: "style" or "style:size" or "style:size:color", e.g. "circle", "diamond:8", "square:6:FF0000"
        var parts = markerSpec.Split(':');
        var style = parts[0].Trim().ToLowerInvariant() switch
        {
            "circle" => C.MarkerStyleValues.Circle,
            "diamond" => C.MarkerStyleValues.Diamond,
            "square" => C.MarkerStyleValues.Square,
            "triangle" => C.MarkerStyleValues.Triangle,
            "star" => C.MarkerStyleValues.Star,
            "x" => C.MarkerStyleValues.X,
            "plus" => C.MarkerStyleValues.Plus,
            "dash" => C.MarkerStyleValues.Dash,
            "dot" => C.MarkerStyleValues.Dot,
            "none" => C.MarkerStyleValues.None,
            "auto" => C.MarkerStyleValues.Auto,
            _ => C.MarkerStyleValues.Circle
        };

        // Snapshot existing per-series marker children so a fan-out (e.g.
        // `marker=circle` after `markerSize=10`) does not blow away
        // previously-set size/spPr/extLst. Spec parts override snapshots.
        var existing = series.GetFirstChild<C.Marker>();
        var existingSize = existing?.GetFirstChild<C.Size>()?.CloneNode(true) as C.Size;
        var existingSpPr = existing?.GetFirstChild<C.ChartShapeProperties>()?.CloneNode(true) as C.ChartShapeProperties;
        var existingExtLst = existing?.GetFirstChild<C.ExtensionList>()?.CloneNode(true) as C.ExtensionList;

        series.RemoveAllChildren<C.Marker>();
        var marker = new C.Marker();
        marker.AppendChild(new C.Symbol { Val = style });
        if (parts.Length > 1 && byte.TryParse(parts[1], out var size))
            marker.AppendChild(new C.Size { Val = size });
        else if (existingSize != null)
            marker.AppendChild(existingSize);
        if (parts.Length > 2)
        {
            var mSpPr = new C.ChartShapeProperties();
            var fill = new Drawing.SolidFill();
            fill.AppendChild(BuildChartColorElement(parts[2]));
            mSpPr.AppendChild(fill);
            marker.AppendChild(mSpPr);
        }
        else if (existingSpPr != null)
            marker.AppendChild(existingSpPr);
        if (existingExtLst != null)
            marker.AppendChild(existingExtLst);

        // Insert marker before data references (xVal, yVal, cat, val, bubbleSize)
        // to satisfy schema order for all chart types including scatter/bubble.
        var markerInsertBefore = (OpenXmlElement?)series.Elements().FirstOrDefault(e =>
            e.LocalName is "xVal" or "yVal" or "cat" or "val" or "bubbleSize"
                or "smooth" or "extLst")
            ?? series.Elements().FirstOrDefault(e => e.LocalName == "trendline");
        if (markerInsertBefore != null) series.InsertBefore(marker, markerInsertBefore);
        else series.AppendChild(marker);
    }

    // ==================== #5 Transparency Helper ====================

    internal static void ApplySeriesAlpha(OpenXmlCompositeElement series, int alphaVal)
    {
        var spPr = GetOrCreateSeriesShapeProperties(series);
        var solidFill = spPr.GetFirstChild<Drawing.SolidFill>();
        if (solidFill == null) return;

        var colorEl = solidFill.FirstChild;
        if (colorEl == null) return;
        // Remove existing alpha
        foreach (var existing in colorEl.Elements<Drawing.Alpha>().ToList())
            existing.Remove();
        colorEl.AppendChild(new Drawing.Alpha { Val = alphaVal });
    }

    // ==================== #6 Gradient Fill Helper ====================

    internal static void ApplySeriesGradient(OpenXmlCompositeElement series, string gradientSpec,
        bool preserveExisting = false)
    {
        // Format: "color1-color2" or "color1-color2-color3" optionally ":angle"
        // e.g. "FF0000-0000FF", "FF0000-00FF00-0000FF:90"
        //
        // preserveExisting=true: fan-out path — if this series already has a
        // per-series GradientFill (set before the chart-level gradient= key),
        // skip so the per-series value wins. Mirrors the ApplySeriesMarker
        // snapshot/restore pattern from 2778017a.
        if (preserveExisting)
        {
            var existingSpPr = series.GetFirstChild<C.ChartShapeProperties>();
            if (existingSpPr?.GetFirstChild<Drawing.GradientFill>() != null)
                return;
        }
        var anglePart = 0;
        var colorsPart = gradientSpec;
        var colonIdx = gradientSpec.LastIndexOf(':');
        if (colonIdx > 0 && int.TryParse(gradientSpec[(colonIdx + 1)..], out var angle))
        {
            anglePart = angle;
            colorsPart = gradientSpec[..colonIdx];
        }

        var colors = colorsPart.Split('-').Select(c => c.Trim()).Where(c => c.Length > 0).ToArray();
        if (colors.Length == 0) return;
        // R28-B4: tolerate a 1-stop spec (the Reader emits one when the
        // source had a single GradientStop) by duplicating the color so
        // the resulting gradient is well-formed (≥2 stops) and visually
        // equivalent to a solid fill. Matches the BuildGradientFill
        // duplicate-on-empty fallback.
        if (colors.Length == 1) colors = new[] { colors[0], colors[0] };

        var gradFill = new Drawing.GradientFill();
        var gsLst = new Drawing.GradientStopList();

        for (int i = 0; i < colors.Length; i++)
        {
            var pos = colors.Length == 1 ? 0 : (int)(i * 100000.0 / (colors.Length - 1));
            var gs = new Drawing.GradientStop { Position = pos };
            gs.AppendChild(BuildChartColorElement(colors[i]));
            gsLst.AppendChild(gs);
        }
        gradFill.AppendChild(gsLst);
        gradFill.AppendChild(new Drawing.LinearGradientFill
        {
            Angle = anglePart * 60000, // degrees to 60000ths
            Scaled = true
        });

        var spPr = GetOrCreateSeriesShapeProperties(series);
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        // Insert gradient before outline
        var outlineEl = spPr.GetFirstChild<Drawing.Outline>();
        if (outlineEl != null) spPr.InsertBefore(gradFill, outlineEl);
        else spPr.PrependChild(gradFill);
    }

    // ==================== #7 Secondary Axis Helper ====================

    /// <summary>
    /// Try to parse a key like "datalabel1.x", "dataLabel2.h" into point index and property.
    /// Returns true if the key matches the pattern.
    /// </summary>
    private static bool TryParseDataLabelLayoutKey(string key, out int pointIndex, out string prop)
    {
        pointIndex = 0;
        prop = "";
        var lower = key.ToLowerInvariant();
        if (!lower.StartsWith("datalabel")) return false;
        var rest = lower["datalabel".Length..]; // e.g. "1.x"
        var dotIdx = rest.IndexOf('.');
        if (dotIdx <= 0) return false;
        if (!int.TryParse(rest[..dotIdx], out pointIndex) || pointIndex < 1) return false;
        prop = rest[(dotIdx + 1)..];
        return prop is "x" or "y" or "w" or "h";
    }

    internal static void ApplySecondaryAxis(C.PlotArea plotArea, HashSet<int> secondarySeriesIndices)
    {
        // Find existing axis IDs
        var existingAxes = plotArea.Elements<C.ValueAxis>().ToList();
        var existingCatAxes = plotArea.Elements<C.CategoryAxis>().ToList();

        uint primaryCatAxisId = existingCatAxes.FirstOrDefault()?.GetFirstChild<C.AxisId>()?.Val?.Value ?? 1u;
        uint primaryValAxisId = existingAxes.FirstOrDefault()?.GetFirstChild<C.AxisId>()?.Val?.Value ?? 2u;
        uint secondaryCatAxisId = 3u;
        uint secondaryValAxisId = 4u;

        // Collect series that should be on secondary axis
        var allChartTypes = plotArea.ChildElements
            .Where(e => e.LocalName.Contains("Chart") || e.LocalName.Contains("chart"))
            .OfType<OpenXmlCompositeElement>().ToList();

        var seriesToMove = new List<OpenXmlElement>();
        int globalIdx = 0;
        foreach (var ct in allChartTypes)
        {
            foreach (var ser in ct.ChildElements.Where(e => e.LocalName == "ser").ToList())
            {
                globalIdx++;
                if (secondarySeriesIndices.Contains(globalIdx))
                    seriesToMove.Add(ser);
            }
        }

        if (seriesToMove.Count == 0) return;

        // Detect type of first moved series' parent chart
        var sourceChartType = seriesToMove[0].Parent;
        if (sourceChartType == null) return;

        // Reject 3D source charts. Excel itself greys out the secondary-axis
        // option on 3D charts because a 3D plotArea has one shared camera /
        // perspective and cannot host a sibling 2D chart element. Previously
        // the code below would match `bar3DChart` / `line3DChart` /
        // `area3DChart` against the StartsWith("bar"/"line"/"area") branches
        // and create a 2D sibling chart, which produced a plotArea mixing
        // 3D + 2D chart types and made Excel crash on open. Match Excel UI:
        // refuse the operation with a clear error.
        var sourceLocalName = sourceChartType.LocalName;
        if (sourceLocalName.Contains("3D", StringComparison.Ordinal))
        {
            throw new ArgumentException(
                $"Invalid secondaryaxis: source chart is 3D ({sourceLocalName}). " +
                "Excel does not support a secondary axis on 3D charts because a 3D " +
                "plot area cannot coexist with a second chart type. Convert to the 2D " +
                "variant first (e.g. column3d -> column) before applying secondaryaxis.");
        }

        // Create a new chart element of the same type for secondary axis.
        // Must match the source's series schema — moving a CT_ScatterSer
        // (xVal/yVal) into a c:lineChart group produces a schema-invalid
        // file because CT_LineSer has no xVal child.
        OpenXmlCompositeElement secondaryChart;
        var localName = sourceLocalName;
        if (localName.StartsWith("line", StringComparison.OrdinalIgnoreCase))
        {
            secondaryChart = new C.LineChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false }
            );
        }
        else if (localName.StartsWith("bar", StringComparison.OrdinalIgnoreCase))
        {
            var origDir = sourceChartType.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column;
            secondaryChart = new C.BarChart(
                new C.BarDirection { Val = origDir },
                new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                new C.VaryColors { Val = false }
            );
        }
        else if (localName.StartsWith("area", StringComparison.OrdinalIgnoreCase))
        {
            secondaryChart = new C.AreaChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false }
            );
        }
        else if (localName.StartsWith("scatter", StringComparison.OrdinalIgnoreCase))
        {
            var origStyle = sourceChartType.GetFirstChild<C.ScatterStyle>()?.Val?.Value
                            ?? C.ScatterStyleValues.LineMarker;
            secondaryChart = new C.ScatterChart(
                new C.ScatterStyle { Val = origStyle },
                new C.VaryColors { Val = false }
            );
        }
        else if (localName.StartsWith("bubble", StringComparison.OrdinalIgnoreCase))
        {
            secondaryChart = new C.BubbleChart(
                new C.VaryColors { Val = false }
            );
        }
        else if (localName.StartsWith("radar", StringComparison.OrdinalIgnoreCase))
        {
            var origStyle = sourceChartType.GetFirstChild<C.RadarStyle>()?.Val?.Value
                            ?? C.RadarStyleValues.Standard;
            secondaryChart = new C.RadarChart(
                new C.RadarStyle { Val = origStyle },
                new C.VaryColors { Val = false }
            );
        }
        else
        {
            // pie / doughnut / surface / stock / etc. — no meaningful concept
            // of a secondary value axis (pie is a single-axis chart; surface/
            // stock have rigid axis layouts). Reject loudly instead of writing
            // a schema-invalid line chart with the wrong series schema.
            throw new ArgumentException(
                $"secondaryaxis: source chart type '{sourceLocalName}' does not "
                + "support a secondary axis. Supported: line, bar, column, "
                + "area, scatter, bubble, radar.");
        }

        // Move series to secondary chart
        foreach (var ser in seriesToMove)
        {
            ser.Remove();
            secondaryChart.AppendChild(ser.CloneNode(true));
        }

        secondaryChart.AppendChild(new C.AxisId { Val = secondaryCatAxisId });
        secondaryChart.AppendChild(new C.AxisId { Val = secondaryValAxisId });

        // Insert secondary chart into plot area (before axes)
        var firstAxis = plotArea.Elements<C.CategoryAxis>().FirstOrDefault() as OpenXmlElement
            ?? plotArea.Elements<C.ValueAxis>().FirstOrDefault();
        if (firstAxis != null)
            plotArea.InsertBefore(secondaryChart, firstAxis);
        else
            plotArea.AppendChild(secondaryChart);

        // Remove existing secondary axes if any
        foreach (var ax in plotArea.Elements<C.CategoryAxis>()
            .Where(a => a.GetFirstChild<C.AxisId>()?.Val?.Value == secondaryCatAxisId).ToList())
            ax.Remove();
        foreach (var ax in plotArea.Elements<C.ValueAxis>()
            .Where(a => a.GetFirstChild<C.AxisId>()?.Val?.Value == secondaryValAxisId).ToList())
            ax.Remove();

        // Add secondary category axis (hidden) — insert after existing axes
        var secCatAxis = new C.CategoryAxis(
            new C.AxisId { Val = secondaryCatAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = true }, // hidden
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.None },
            new C.CrossingAxis { Val = secondaryValAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero }
        );

        // Add secondary value axis (visible, on the right)
        var secValAxis = BuildValueAxis(secondaryValAxisId, secondaryCatAxisId, C.AxisPositionValues.Right);
        secValAxis.RemoveAllChildren<C.MajorGridlines>(); // secondary axis typically has no gridlines

        // Bind secondary Y axis to the right edge by crossing the (hidden) secondary
        // category axis at its maximum. Without this, Excel ignores axPos="r" and
        // renders both Y axes on the left edge — BuildValueAxis defaults crosses to
        // autoZero, which is correct for the primary axis but wrong here.
        foreach (var c in secValAxis.Elements<C.Crosses>().ToList()) c.Remove();
        foreach (var c in secValAxis.Elements<C.CrossesAt>().ToList()) c.Remove();
        // Schema order in CT_ValAx: crossAx → crosses → crossBetween. BuildValueAxis
        // already emitted CrossBetween last, so a plain AppendChild here would place
        // the new Crosses *after* CrossBetween — schema-illegal and rejected by
        // Excel/PowerPoint. Insert before CrossBetween (or fall back to AppendChild
        // if the axis somehow has no CrossBetween).
        var newCrosses = new C.Crosses { Val = C.CrossesValues.Maximum };
        var crossBetween = secValAxis.GetFirstChild<C.CrossBetween>();
        if (crossBetween != null)
            secValAxis.InsertBefore(newCrosses, crossBetween);
        else
            secValAxis.AppendChild(newCrosses);

        // Insert after the last existing axis to maintain schema order
        var lastAxis = plotArea.Elements<C.ValueAxis>().LastOrDefault() as OpenXmlElement
            ?? plotArea.Elements<C.CategoryAxis>().LastOrDefault() as OpenXmlElement;
        if (lastAxis != null)
        {
            lastAxis.InsertAfterSelf(secCatAxis);
            secCatAxis.InsertAfterSelf(secValAxis);
        }
        else
        {
            plotArea.AppendChild(secCatAxis);
            plotArea.AppendChild(secValAxis);
        }
    }

    /// <summary>
    /// Returns a sort order for chart properties to ensure structural properties
    /// (legend, title) are processed before their styling counterparts (legendFont, title.color).
    /// </summary>
    private static int GetPropertyOrder(string key)
    {
        var k = key.ToLowerInvariant();
        // Presets first (they recursively call SetChartProperties)
        if (k is "preset" or "style.preset" or "theme") return 0;
        // Structural: create/position legend and title before styling them
        if (k == "legend") return 1;
        if (k == "title") return 1;
        // Styling of legend/title after structural
        if (k.StartsWith("legend")) return 2;
        if (k.StartsWith("title")) return 2;
        // Everything else at default priority
        return 5;
    }

    // R24-3: in-place expand keys of the form "{prefix}.layout" with value
    // "x:N,y:N,w:N,h:N" (any subset, any order) into individual {prefix}.x,
    // {prefix}.y, {prefix}.w, {prefix}.h entries. Existing individual keys
    // are not overwritten, so callers can still override one component.
    // Recognized prefixes match the dispatch table above.
    private static readonly string[] _layoutPrefixes =
    {
        "legend", "plotarea", "title",
        "trendlinelabel", "displayunitslabel",
    };

    internal static void ExpandCombinedLayoutKeys(Dictionary<string, string> properties)
    {
        // Find all "*.layout" keys (case-insensitive) up front so we can
        // mutate the dict while iterating.
        var layoutKeys = properties.Keys
            .Where(k => k.EndsWith(".layout", StringComparison.OrdinalIgnoreCase))
            .ToList();
        foreach (var key in layoutKeys)
        {
            var prefix = key[..^".layout".Length];
            if (!_layoutPrefixes.Contains(prefix.ToLowerInvariant())) continue;
            var raw = properties[key];
            if (string.IsNullOrWhiteSpace(raw)) { properties.Remove(key); continue; }
            // value: "x:0.1,y:0.5,w:0.2,h:0.4" — comma-separated k:v pairs,
            // or positional CSV "0.1,0.2,0.3,0.4" (exactly 4 → x,y,w,h).
            // CONSISTENCY(layout-csv): bt-2/fuzz-LL01 — positional CSV is the
            // user-friendly form; reject ambiguous arities so silent-success
            // bugs cannot recur.
            var parts = raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var hasColon = parts.Any(p => p.Contains(':'));
            if (!hasColon)
            {
                if (parts.Length != 4)
                    throw new ArgumentException(
                        $"{key}: positional CSV layout requires exactly 4 values (x,y,w,h); got {parts.Length}. " +
                        $"Use named form '{key}=x:N,y:N,w:N,h:N' for partial layouts.");
                var dims = new[] { "x", "y", "w", "h" };
                for (int i = 0; i < 4; i++)
                {
                    var expandedKey = $"{prefix}.{dims[i]}";
                    if (!properties.ContainsKey(expandedKey))
                        properties[expandedKey] = parts[i];
                }
            }
            else
            {
                foreach (var part in parts)
                {
                    var colonIdx = part.IndexOf(':');
                    if (colonIdx <= 0) continue;
                    var dim = part[..colonIdx].Trim().ToLowerInvariant();
                    var val = part[(colonIdx + 1)..].Trim();
                    if (dim is "x" or "y" or "w" or "h")
                    {
                        var expandedKey = $"{prefix}.{dim}";
                        if (!properties.ContainsKey(expandedKey))
                            properties[expandedKey] = val;
                    }
                }
            }
            properties.Remove(key);
        }
    }

    // fuzz-TL01/TL02: parse-validate a trendline.* sub-property value the same
    // way ApplyTrendlineOptions would, but without mutating any element. Used
    // by the chart-level fan-out so unrecognized values are rejected even when
    // the chart has no trendline to apply them to.
    private static void ValidateTrendlineOptionValue(string subKey, string value, string fullKey)
    {
        switch (subKey)
        {
            case "name" or "label":
                break; // any string is valid
            case "forward" or "forecastforward"
                or "backward" or "forecastbackward"
                or "intercept":
                ParseHelpers.SafeParseDouble(value, fullKey);
                break;
            case "order" or "period":
                ParseHelpers.SafeParseInt(value, fullKey);
                break;
            case "disprsqr" or "rsquared" or "r2" or "displayrsquared"
                or "dispeq" or "equation" or "displayequation":
                var v = (value ?? "").Trim().ToLowerInvariant();
                if (v is not ("true" or "false" or "1" or "0" or "yes" or "no" or "on" or "off"))
                    throw new ArgumentException(
                        $"{fullKey}: expected boolean (true/false/1/0/yes/no/on/off), got '{value}'.");
                break;
        }
    }

    // R8-3: previously the dotted show* / top-level show* setters only flipped
    // existing <c:dLbls> containers. On a chart whose data labels had been
    // cleared (datalabels=none, or new charts emitted without dLbls), the
    // Descendants<DataLabels> enumeration returned nothing and the operation
    // succeeded silently with no XML change. Caller saw success=true and an
    // unchanged chart — surprise round-trip behaviour. Enable-by-show*
    // semantics expect us to materialise a minimal container when one is
    // missing; collect (and seed if needed) the DataLabels for each chartType
    // in the PlotArea.
    private static bool EnsureDataLabelsForShowToggle(
        C.Chart chart, string key, List<string> unsupported, out List<C.DataLabels> dataLabels)
    {
        dataLabels = new List<C.DataLabels>();
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) { unsupported.Add(key); return false; }
        var chartTypes = plotArea.ChildElements
            .Where(e => e.LocalName.Contains("Chart") || e.LocalName.Contains("chart"))
            .ToList();
        if (chartTypes.Count == 0) { unsupported.Add(key); return false; }
        foreach (var chartTypeEl in chartTypes)
        {
            var existing = chartTypeEl.Elements<C.DataLabels>().FirstOrDefault();
            if (existing != null) { dataLabels.Add(existing); continue; }
            // Schema-order insertion mirrors the "datalabels" full-replace path:
            // dLbls precedes dropLines/hiLowLines/upDownBars/gapWidth/overlap/
            // showMarker/holeSize/firstSliceAngle/axId.
            // Schema order on CT_DLbls: showLegendKey, showVal, showCatName,
            // showSerName, showPercent, showBubbleSize. Match the existing
            // datalabels=full-replace seeding above.
            var dl = new C.DataLabels();
            dl.AppendChild(new C.ShowLegendKey { Val = false });
            dl.AppendChild(new C.ShowValue { Val = false });
            dl.AppendChild(new C.ShowCategoryName { Val = false });
            dl.AppendChild(new C.ShowSeriesName { Val = false });
            dl.AppendChild(new C.ShowPercent { Val = false });
            var insertBefore = chartTypeEl.GetFirstChild<C.DropLines>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.HighLowLines>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.UpDownBars>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.GapWidth>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.Overlap>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.ShowMarker>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.HoleSize>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.FirstSliceAngle>() as OpenXmlElement
                ?? chartTypeEl.GetFirstChild<C.AxisId>();
            if (insertBefore != null) chartTypeEl.InsertBefore(dl, insertBefore);
            else chartTypeEl.AppendChild(dl);
            dataLabels.Add(dl);
        }
        return dataLabels.Count > 0;
    }
}
