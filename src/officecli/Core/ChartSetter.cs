// Copyright 2025 OfficeCli (officecli.ai)
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
                        chart.PrependChild(BuildChartTitle(value));
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
                        var pos = value.ToLowerInvariant() switch
                        {
                            "top" or "t" => C.LegendPositionValues.Top,
                            "left" or "l" => C.LegendPositionValues.Left,
                            "right" or "r" => C.LegendPositionValues.Right,
                            _ => C.LegendPositionValues.Bottom
                        };
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
                            var parts = value.ToLowerInvariant().Split(',').Select(s => s.Trim()).ToHashSet();
                            dl.AppendChild(new C.ShowLegendKey { Val = false });
                            dl.AppendChild(new C.ShowValue { Val = parts.Contains("value") || parts.Contains("true") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowCategoryName { Val = parts.Contains("category") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowSeriesName { Val = parts.Contains("series") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowPercent { Val = parts.Contains("percent") || parts.Contains("all") });
                            // Insert dLbls before gapWidth/overlap/showMarker/holeSize/axId per schema order
                            var dlInsertBefore = chartTypeEl.GetFirstChild<C.GapWidth>() as OpenXmlElement
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

                    // Doughnut does NOT support dLblPos at all — skip entirely
                    if (plotArea2.GetFirstChild<C.DoughnutChart>() != null) break;

                    // Pie only supports: bestFit, center, insideEnd, insideBase
                    var isPie = plotArea2.GetFirstChild<C.PieChart>() != null
                        || plotArea2.GetFirstChild<C.Pie3DChart>() != null;

                    var dlblPos = value.ToLowerInvariant() switch
                    {
                        "center" or "ctr" => C.DataLabelPositionValues.Center,
                        "insideend" or "inside" => C.DataLabelPositionValues.InsideEnd,
                        "insidebase" or "base" => C.DataLabelPositionValues.InsideBase,
                        "outsideend" or "outside" => isPie
                            ? C.DataLabelPositionValues.BestFit
                            : C.DataLabelPositionValues.OutsideEnd,
                        "bestfit" or "best" or "auto" => C.DataLabelPositionValues.BestFit,
                        "top" or "t" => isPie
                            ? C.DataLabelPositionValues.BestFit
                            : C.DataLabelPositionValues.Top,
                        "bottom" or "b" => isPie
                            ? C.DataLabelPositionValues.BestFit
                            : C.DataLabelPositionValues.Bottom,
                        "left" or "l" => isPie
                            ? C.DataLabelPositionValues.BestFit
                            : C.DataLabelPositionValues.Left,
                        "right" or "r" => isPie
                            ? C.DataLabelPositionValues.BestFit
                            : C.DataLabelPositionValues.Right,
                        _ => isPie
                            ? C.DataLabelPositionValues.BestFit
                            : C.DataLabelPositionValues.OutsideEnd
                    };
                    foreach (var dl in plotArea2.Descendants<C.DataLabels>())
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

                case "colors":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var colorList = value.Split(',').Select(c => c.Trim()).ToArray();
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
                    valAxis.RemoveAllChildren<C.MajorUnit>();
                    valAxis.AppendChild(new C.MajorUnit { Val = ParseHelpers.SafeParseDouble(value, "majorunit") });
                    break;
                }

                case "minorunit":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.MinorUnit>();
                    valAxis.AppendChild(new C.MinorUnit { Val = ParseHelpers.SafeParseDouble(value, "minorunit") });
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
                    chartSpace!.RemoveAllChildren<C.ChartShapeProperties>();
                    var spPr = new C.ChartShapeProperties();
                    spPr.AppendChild(BuildFillElement(value));
                    chartSpace.InsertAfter(spPr, chart);
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
                        ApplySeriesMarker(ser, value);
                    break;
                }

                case "markersize":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var mSize = ParseHelpers.SafeParseByte(value, "markersize");
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        var marker = ser.GetFirstChild<C.Marker>();
                        if (marker == null) { marker = new C.Marker(); ser.AppendChild(marker); }
                        marker.RemoveAllChildren<C.Size>();
                        marker.AppendChild(new C.Size { Val = mSize });
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
                case "gradient":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    // Format: "color1-color2" or "color1-color2-color3" with optional ":angle"
                    // e.g. "FF0000-0000FF" or "FF0000-00FF00-0000FF:90"
                    var allSer = plotArea2.Descendants<OpenXmlCompositeElement>()
                        .Where(e => e.LocalName == "ser").ToList();
                    for (int si = 0; si < allSer.Count; si++)
                        ApplySeriesGradient(allSer[si], value);
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
                    for (int si = 0; si < Math.Min(gradList.Length, allSer.Count); si++)
                        ApplySeriesGradient(allSer[si], gradList[si]);
                    break;
                }

                case "view3d" or "camera" or "perspective":
                {
                    // Format: "rotX,rotY,perspective" e.g. "15,20,30" or just "20" for perspective
                    var v3dParts = value.Split(',');
                    chart.RemoveAllChildren<C.View3D>();
                    var view3d = new C.View3D();
                    if (v3dParts.Length >= 1 && int.TryParse(v3dParts[0], out var rx))
                        view3d.AppendChild(new C.RotateX { Val = (sbyte)rx });
                    if (v3dParts.Length >= 2 && int.TryParse(v3dParts[1], out var ry))
                        view3d.AppendChild(new C.RotateY { Val = (ushort)ry });
                    if (v3dParts.Length >= 3 && int.TryParse(v3dParts[2], out var persp))
                        view3d.AppendChild(new C.Perspective { Val = (byte)persp });
                    else if (v3dParts.Length == 1 && int.TryParse(v3dParts[0], out var p))
                        view3d.AppendChild(new C.Perspective { Val = (byte)p });
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
                        var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                        if (spPr == null) { spPr = new C.ChartShapeProperties(); ser.AppendChild(spPr); }
                        spPr.RemoveAllChildren<Drawing.SolidFill>();
                        spPr.RemoveAllChildren<Drawing.GradientFill>();
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
                        var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                        if (spPr == null) { spPr = new C.ChartShapeProperties(); ser.AppendChild(spPr); }
                        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? new Drawing.EffectList();
                        if (effectList.Parent == null) spPr.AppendChild(effectList);
                        effectList.RemoveAllChildren<Drawing.OuterShadow>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            effectList.AppendChild(DrawingEffectsHelper.BuildOuterShadow(value, BuildChartColorElement));
                    }
                    break;
                }

                case "series.outline" or "seriesoutline":
                {
                    // Apply outline to all series bars. Format: "COLOR" or "COLOR-WIDTH" e.g. "FFFFFF-1"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var outParts = value.Split('-');
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        var spPr = ser.GetFirstChild<C.ChartShapeProperties>();
                        if (spPr == null) { spPr = new C.ChartShapeProperties(); ser.AppendChild(spPr); }
                        spPr.RemoveAllChildren<Drawing.Outline>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var widthPt = outParts.Length > 1 && double.TryParse(outParts[1], System.Globalization.CultureInfo.InvariantCulture, out var w) ? w : 0.5;
                            var outline = new Drawing.Outline { Width = (int)(widthPt * 12700) };
                            var sf = new Drawing.SolidFill();
                            sf.AppendChild(BuildChartColorElement(outParts[0]));
                            outline.AppendChild(sf);
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
                    foreach (var gapEl in plotArea2.Descendants<C.GapWidth>())
                        gapEl.Val = (ushort)gw;
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
                    // value = series indices on secondary axis, e.g. "2,3" (1-based)
                    var secondaryIndices = value.Split(',')
                        .Select(s => int.TryParse(s.Trim(), out var v) ? v : -1)
                        .Where(v => v > 0).ToHashSet();
                    ApplySecondaryAxis(plotArea2, secondaryIndices);
                    break;
                }

                case "plotarea.x" or "plotarea.y" or "plotarea.w" or "plotarea.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal))
                    { unsupported.Add(key); break; }
                    var plotArea3 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea3 == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(plotArea3, key.Split('.')[1].ToLowerInvariant(), layoutVal, isPlotArea: true);
                    break;
                }

                case "title.x" or "title.y" or "title.w" or "title.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal))
                    { unsupported.Add(key); break; }
                    var titleEl = chart.GetFirstChild<C.Title>();
                    if (titleEl == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(titleEl, key.Split('.')[1].ToLowerInvariant(), layoutVal);
                    break;
                }

                case "legend.x" or "legend.y" or "legend.w" or "legend.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal))
                    { unsupported.Add(key); break; }
                    var legendEl = chart.GetFirstChild<C.Legend>();
                    if (legendEl == null) { unsupported.Add(key); break; }
                    SetManualLayoutProperty(legendEl, key.Split('.')[1].ToLowerInvariant(), layoutVal);
                    break;
                }

                case "trendlinelabel.x" or "trendlinelabel.y" or "trendlinelabel.w" or "trendlinelabel.h":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal))
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
                        System.Globalization.CultureInfo.InvariantCulture, out var layoutVal))
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
                    { ax.RemoveAllChildren<C.MajorTickMark>(); ax.AppendChild(new C.MajorTickMark { Val = tickVal }); }
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    { ax.RemoveAllChildren<C.MajorTickMark>(); ax.AppendChild(new C.MajorTickMark { Val = tickVal }); }
                    break;
                }

                case "minortickmark" or "minortick":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var tickVal = ParseTickMark(value);
                    foreach (var ax in plotArea2.Elements<C.ValueAxis>())
                    { ax.RemoveAllChildren<C.MinorTickMark>(); ax.AppendChild(new C.MinorTickMark { Val = tickVal }); }
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    { ax.RemoveAllChildren<C.MinorTickMark>(); ax.AppendChild(new C.MinorTickMark { Val = tickVal }); }
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
                    { ax.RemoveAllChildren<C.TickLabelPosition>(); ax.AppendChild(new C.TickLabelPosition { Val = tlPos }); }
                    foreach (var ax in plotArea2.Elements<C.CategoryAxis>())
                    { ax.RemoveAllChildren<C.TickLabelPosition>(); ax.AppendChild(new C.TickLabelPosition { Val = tlPos }); }
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
                    { ax.RemoveAllChildren<C.AxisPosition>(); ax.AppendChild(new C.AxisPosition { Val = axPos }); }
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
                    valAx.AppendChild(new C.Crosses { Val = crossVal });
                    break;
                }

                case "crossesat":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAx == null) { unsupported.Add(key); break; }
                    valAx.RemoveAllChildren<C.Crosses>();
                    valAx.RemoveAllChildren<C.CrossesAt>();
                    valAx.AppendChild(new C.CrossesAt { Val = ParseHelpers.SafeParseDouble(value, "crossesAt") });
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
                    valAx.AppendChild(new C.CrossBetween { Val = cbVal });
                    break;
                }

                case "axisorientation" or "axisreverse":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAx?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.Orientation>();
                    var orient = ParseHelpers.IsTruthy(value) || value.Equals("maxmin", StringComparison.OrdinalIgnoreCase)
                        ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax;
                    scaling.PrependChild(new C.Orientation { Val = orient });
                    break;
                }

                case "logbase" or "logscale":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAx = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAx?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.LogBase>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                        !value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        var logVal = ParseHelpers.SafeParseDouble(value, "logBase");
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
                        valAx.AppendChild(du);
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
                    // Chart-level smooth on LineChart
                    foreach (var lc in plotArea2.Elements<C.LineChart>())
                    { lc.RemoveAllChildren<C.Smooth>(); lc.AppendChild(new C.Smooth { Val = smoothVal }); }
                    // Also set per-series smooth for line and scatter series
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        if (ser.Parent is C.LineChart or C.ScatterChart)
                        {
                            ser.RemoveAllChildren<C.Smooth>();
                            ser.AppendChild(new C.Smooth { Val = smoothVal });
                        }
                    }
                    break;
                }

                case "showmarker" or "showmarkers":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var showVal = ParseHelpers.IsTruthy(value);
                    foreach (var lc in plotArea2.Elements<C.LineChart>())
                    { lc.RemoveAllChildren<C.ShowMarker>(); lc.AppendChild(new C.ShowMarker { Val = showVal }); }
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
                        ct.PrependChild(new C.VaryColors { Val = varyVal });
                    }
                    break;
                }

                case "dispblanksas" or "blanksas":
                {
                    chart.RemoveAllChildren<C.DisplayBlanksAs>();
                    var dbVal = value.ToLowerInvariant() switch
                    {
                        "zero" => C.DisplayBlanksAsValues.Zero,
                        "span" or "connect" => C.DisplayBlanksAsValues.Span,
                        _ => C.DisplayBlanksAsValues.Gap
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
                    chart.RemoveAllChildren<C.AutoTitleDeleted>();
                    chart.AppendChild(new C.AutoTitleDeleted { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                case "plotvisonly" or "plotvisibleonly":
                {
                    chart.RemoveAllChildren<C.PlotVisibleOnly>();
                    chart.AppendChild(new C.PlotVisibleOnly { Val = ParseHelpers.IsTruthy(value) });
                    break;
                }

                // ==================== Series-Level Properties ====================

                case "trendline":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        ser.RemoveAllChildren<C.Trendline>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var tl = BuildTrendline(value);
                            // CT_*Ser schema order: ..., dLbls, trendline, errBars, cat, val, ...
                            // Insert trendline before <c:cat> or <c:val> if present
                            var catEl = ser.GetFirstChild<C.CategoryAxisData>();
                            var valEl = ser.GetFirstChild<C.Values>();
                            var anchor = (OpenXmlElement?)catEl ?? valEl;
                            if (anchor != null)
                                anchor.InsertBeforeSelf(tl);
                            else
                                ser.AppendChild(tl);
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
                        ser.AppendChild(new C.InvertIfNegative { Val = inv });
                    }
                    break;
                }

                case "explosion" or "explode":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var expVal = (uint)ParseHelpers.SafeParseInt(value, "explosion");
                    foreach (var ser in plotArea2.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser"))
                    {
                        ser.RemoveAllChildren<C.Explosion>();
                        if (expVal > 0) ser.AppendChild(new C.Explosion { Val = expVal });
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
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            ser.AppendChild(BuildErrorBars(value));
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
                    var cSpPr = chartSpace!.GetFirstChild<C.ChartShapeProperties>();
                    if (cSpPr == null) { cSpPr = new C.ChartShapeProperties(); chartSpace.InsertAfter(cSpPr, chart); }
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
                    pie.RemoveAllChildren<C.FirstSliceAngle>();
                    pie.AppendChild(new C.FirstSliceAngle { Val = (ushort)ParseHelpers.SafeParseInt(value, "firstSliceAngle") });
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
                    var bsNode = new C.BubbleScale { Val = (uint)ParseHelpers.SafeParseInt(value, "bubbleScale") };
                    var bsAxId = bubble.GetFirstChild<C.AxisId>();
                    if (bsAxId != null) bubble.InsertBefore(bsNode, bsAxId);
                    else bubble.AppendChild(bsNode);
                    break;
                }

                case "shownegbubbles":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var bubble = plotArea2?.GetFirstChild<C.BubbleChart>();
                    if (bubble == null) { unsupported.Add(key); break; }
                    bubble.RemoveAllChildren<C.ShowNegativeBubbles>();
                    bubble.AppendChild(new C.ShowNegativeBubbles { Val = ParseHelpers.IsTruthy(value) });
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
                    bubble.AppendChild(new C.SizeRepresents { Val = srVal });
                    break;
                }

                case "gapdepth":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var bar3d = plotArea2?.GetFirstChild<C.Bar3DChart>();
                    if (bar3d == null) { unsupported.Add(key); break; }
                    bar3d.RemoveAllChildren<C.GapDepth>();
                    bar3d.AppendChild(new C.GapDepth { Val = (ushort)ParseHelpers.SafeParseInt(value, "gapDepth") });
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
                    bar3d.AppendChild(new C.Shape { Val = shapeVal });
                    break;
                }

                case "droplines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var lc = plotArea2?.GetFirstChild<C.LineChart>();
                    if (lc == null) { unsupported.Add(key); break; }
                    lc.RemoveAllChildren<C.DropLines>();
                    if (ParseHelpers.IsTruthy(value) || !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var dl = new C.DropLines();
                        if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                            dl.AppendChild(BuildLineShapeProperties(value));
                        lc.AppendChild(dl);
                    }
                    break;
                }

                case "hilowlines":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var lc = plotArea2?.GetFirstChild<C.LineChart>();
                    if (lc == null) { unsupported.Add(key); break; }
                    lc.RemoveAllChildren<C.HighLowLines>();
                    if (ParseHelpers.IsTruthy(value) || !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var hl = new C.HighLowLines();
                        if (!value.Equals("true", StringComparison.OrdinalIgnoreCase))
                            hl.AppendChild(BuildLineShapeProperties(value));
                        lc.AppendChild(hl);
                    }
                    break;
                }

                case "updownbars":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var lc = plotArea2?.GetFirstChild<C.LineChart>();
                    if (lc == null) { unsupported.Add(key); break; }
                    lc.RemoveAllChildren<C.UpDownBars>();
                    if (ParseHelpers.IsTruthy(value))
                    {
                        var udb = new C.UpDownBars();
                        udb.AppendChild(new C.GapWidth { Val = 150 });
                        udb.AppendChild(new C.UpBars());
                        udb.AppendChild(new C.DownBars());
                        lc.AppendChild(udb);
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
                        if (show) barChart.AppendChild(new C.SeriesLines());
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

                // ==================== Advanced Features ====================

                case "referenceline" or "refline" or "targetline":
                {
                    // Format: "value" or "value:color" or "value:color:label" or "value:color:label:dash"
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                        if (plotArea2 != null)
                            RemoveExistingReferenceLines(plotArea2);
                        break;
                    }
                    AddReferenceLine(chart, value);
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
                            firstSer.AppendChild(dLbls);
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
                        HandleSeriesDottedProperty(ser, sProp, value);
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

    private static Drawing.PresetLineDashValues ParseDashStyle(string dash)
    {
        return dash.ToLowerInvariant() switch
        {
            "solid" => Drawing.PresetLineDashValues.Solid,
            "dot" or "sysdot" => Drawing.PresetLineDashValues.SystemDot,
            "dash" or "sysdash" => Drawing.PresetLineDashValues.SystemDash,
            "dashdot" or "sysdash_dot" => Drawing.PresetLineDashValues.SystemDashDot,
            "longdash" => Drawing.PresetLineDashValues.LargeDash,
            "longdashdot" => Drawing.PresetLineDashValues.LargeDashDot,
            "longdashdotdot" => Drawing.PresetLineDashValues.LargeDashDotDot,
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
        outline.Width = widthEmu;
    }

    internal static void ApplySeriesLineDash(OpenXmlCompositeElement series, string dashStyle)
    {
        var spPr = GetOrCreateSeriesShapeProperties(series);
        var outline = spPr.GetFirstChild<Drawing.Outline>();
        if (outline == null) { outline = new Drawing.Outline(); spPr.AppendChild(outline); }
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
            _ => C.MarkerStyleValues.Circle
        };

        series.RemoveAllChildren<C.Marker>();
        var marker = new C.Marker();
        marker.AppendChild(new C.Symbol { Val = style });
        if (parts.Length > 1 && byte.TryParse(parts[1], out var size))
            marker.AppendChild(new C.Size { Val = size });
        if (parts.Length > 2)
        {
            var mSpPr = new C.ChartShapeProperties();
            var fill = new Drawing.SolidFill();
            fill.AppendChild(BuildChartColorElement(parts[2]));
            mSpPr.AppendChild(fill);
            marker.AppendChild(mSpPr);
        }

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

    internal static void ApplySeriesGradient(OpenXmlCompositeElement series, string gradientSpec)
    {
        // Format: "color1-color2" or "color1-color2-color3" optionally ":angle"
        // e.g. "FF0000-0000FF", "FF0000-00FF00-0000FF:90"
        var anglePart = 0;
        var colorsPart = gradientSpec;
        var colonIdx = gradientSpec.LastIndexOf(':');
        if (colonIdx > 0 && int.TryParse(gradientSpec[(colonIdx + 1)..], out var angle))
        {
            anglePart = angle;
            colorsPart = gradientSpec[..colonIdx];
        }

        var colors = colorsPart.Split('-').Select(c => c.Trim()).ToArray();
        if (colors.Length < 2) return;

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

        // Create a new chart element of the same type for secondary axis
        OpenXmlCompositeElement secondaryChart;
        var localName = sourceChartType.LocalName;
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
        else
        {
            // Default to line for secondary axis
            secondaryChart = new C.LineChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false }
            );
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
}
