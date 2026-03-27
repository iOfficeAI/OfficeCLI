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

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
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
                        chartSpace.InsertBefore(new C.Style { Val = (byte)ParseHelpers.SafeParseInt(value, "style") }, chart);
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
                    chart.PrependChild(view3d);
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
                            spPr.AppendChild(outline);
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
                        unsupported.Add(key);
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

        // Insert marker after spPr or seriesText
        var afterEl = (OpenXmlElement?)series.GetFirstChild<C.ChartShapeProperties>()
            ?? series.GetFirstChild<C.SeriesText>();
        if (afterEl != null) afterEl.InsertAfterSelf(marker);
        else series.PrependChild(marker);
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
}
