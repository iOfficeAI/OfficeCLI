// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal static partial class ChartHelper
{
    // ==================== Chart Readback ====================

    internal static void ReadChartProperties(C.Chart chart, DocumentNode node, int depth)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        var chartType = DetectChartType(plotArea);
        if (chartType != null) node.Format["chartType"] = chartType;

        var titleEl = chart.GetFirstChild<C.Title>();
        var titleText = titleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (titleText != null) node.Format["title"] = titleText;

        // Title formatting: font, size, color, bold from RunProperties
        if (titleEl != null)
        {
            var titleRun = titleEl.Descendants<Drawing.Run>().FirstOrDefault();
            var titleRp = titleRun?.RunProperties;
            if (titleRp != null)
            {
                var titleFont = titleRp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
                if (titleFont != null) node.Format["title.font"] = titleFont;
                if (titleRp.FontSize?.HasValue == true)
                    node.Format["title.size"] = $"{titleRp.FontSize.Value / 100.0:0.##}pt";
                var titleFill = titleRp.GetFirstChild<Drawing.SolidFill>();
                if (titleFill != null)
                {
                    var tColor = ReadColorFromFill(titleFill);
                    if (tColor != null) node.Format["title.color"] = tColor;
                }
                if (titleRp.Bold?.HasValue == true)
                    node.Format["title.bold"] = titleRp.Bold.Value ? "true" : "false";
            }
        }

        var legend = chart.GetFirstChild<C.Legend>();
        if (legend != null)
        {
            var posRaw = legend.GetFirstChild<C.LegendPosition>()?.Val?.HasValue == true
                ? legend.GetFirstChild<C.LegendPosition>()!.Val!.InnerText : "b";
            node.Format["legend"] = posRaw switch
            {
                "b" => "bottom",
                "t" => "top",
                "l" => "left",
                "r" => "right",
                "tr" => "topRight",
                _ => posRaw
            };
        }

        var dataLabels = plotArea.Descendants<C.DataLabels>().FirstOrDefault();
        if (dataLabels != null)
        {
            var parts = new List<string>();
            if (dataLabels.GetFirstChild<C.ShowValue>()?.Val?.Value == true) parts.Add("value");
            if (dataLabels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true) parts.Add("category");
            if (dataLabels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true) parts.Add("series");
            if (dataLabels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true) parts.Add("percent");
            if (parts.Count > 0) node.Format["dataLabels"] = string.Join(",", parts);
            var dlPos = dataLabels.GetFirstChild<C.DataLabelPosition>()?.Val;
            if (dlPos?.HasValue == true) node.Format["labelPos"] = dlPos.InnerText;
        }

        // Chart style
        var style = chart.Parent?.GetFirstChild<C.Style>();
        if (style?.Val?.HasValue == true) node.Format["style"] = (int)style.Val.Value;

        // ManualLayout readback: plotArea, title, legend, trendlineLabel, displayUnitsLabel
        ReadManualLayout(plotArea, node, "plotArea");
        if (titleEl != null) ReadManualLayout(titleEl, node, "title");
        if (legend != null) ReadManualLayout(legend, node, "legend");
        var trendlineLbl = plotArea.Descendants<C.TrendlineLabel>().FirstOrDefault();
        if (trendlineLbl != null) ReadManualLayout(trendlineLbl, node, "trendlineLabel");
        var dispUnitsLbl = chart.Descendants<C.DisplayUnitsLabel>().FirstOrDefault();
        if (dispUnitsLbl != null) ReadManualLayout(dispUnitsLbl, node, "displayUnitsLabel");

        // Individual data label (dLbl) layout readback — first series
        var firstSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .FirstOrDefault(e => e.LocalName == "ser");
        var dLbls = firstSer?.GetFirstChild<C.DataLabels>();
        if (dLbls != null)
        {
            foreach (var dLbl in dLbls.Elements<C.DataLabel>())
            {
                var idx = dLbl.Index?.Val?.Value;
                if (idx == null) continue;
                ReadManualLayout(dLbl, node, $"dataLabel{idx.Value + 1}");
            }
        }

        // Plot area fill (plotArea uses C.ShapeProperties, not C.ChartShapeProperties)
        var plotSpPr = plotArea.GetFirstChild<C.ShapeProperties>();
        var plotFill = plotSpPr?.GetFirstChild<Drawing.SolidFill>();
        if (plotFill != null)
        {
            var pColor = ReadColorFromFill(plotFill);
            if (pColor != null) node.Format["plotFill"] = pColor;
        }

        // Chart area fill (ChartSpace > spPr, NOT PlotArea)
        // Note: The SDK serializes ChartShapeProperties but deserializes it as C.ShapeProperties
        // after round-trip. Check both types, plus in-memory ChartShapeProperties.
        {
            Drawing.SolidFill? chartAreaFill = null;
            var csSpPr = chart.Parent?.GetFirstChild<C.ShapeProperties>();
            if (csSpPr != null)
                chartAreaFill = csSpPr.GetFirstChild<Drawing.SolidFill>();
            if (chartAreaFill == null)
            {
                var csCSpPr = chart.Parent?.GetFirstChild<C.ChartShapeProperties>();
                if (csCSpPr != null)
                    chartAreaFill = csCSpPr.GetFirstChild<Drawing.SolidFill>();
            }
            if (chartAreaFill != null)
            {
                var cColor = ReadColorFromFill(chartAreaFill);
                if (cColor != null) node.Format["chartFill"] = cColor;
            }
        }

        // Gridlines: "true" for presence, detail in gridlineColor/gridlineWidth/gridlineDash
        var valAxisForGrid = plotArea.GetFirstChild<C.ValueAxis>();
        var majorGL = valAxisForGrid?.GetFirstChild<C.MajorGridlines>();
        if (majorGL != null)
        {
            node.Format["gridlines"] = "true";
            ReadGridlineDetail(majorGL, node, "gridline");
        }
        var minorGL = valAxisForGrid?.GetFirstChild<C.MinorGridlines>();
        if (minorGL != null)
        {
            node.Format["minorGridlines"] = "true";
            ReadGridlineDetail(minorGL, node, "minorGridline");
        }

        // GapWidth / Overlap from bar/column chart
        var barChart = plotArea.GetFirstChild<C.BarChart>();
        var gapWidthEl = barChart?.GetFirstChild<C.GapWidth>();
        if (gapWidthEl?.Val?.HasValue == true) node.Format["gapwidth"] = gapWidthEl.Val.Value.ToString();
        var overlapEl = barChart?.GetFirstChild<C.Overlap>();
        if (overlapEl?.Val?.HasValue == true) node.Format["overlap"] = overlapEl.Val.Value.ToString();

        // Legend font (TextProperties on Legend element)
        if (legend != null)
        {
            var legendTp = legend.GetFirstChild<C.TextProperties>();
            if (legendTp != null)
            {
                var legendFontStr = ReadFontSpec(legendTp);
                if (legendFontStr != null) node.Format["legendFont"] = legendFontStr;
            }
        }

        // Axis font (TextProperties on value axis)
        var valAxisTp = valAxisForGrid?.GetFirstChild<C.TextProperties>();
        if (valAxisTp != null)
        {
            var axisFontStr = ReadFontSpec(valAxisTp);
            if (axisFontStr != null) node.Format["axisFont"] = axisFontStr;
        }

        // Secondary axis
        var valAxes = plotArea.Elements<C.ValueAxis>().ToList();
        if (valAxes.Count > 1) node.Format["secondaryAxis"] = "true";

        // Axis titles
        var valAxis = plotArea.GetFirstChild<C.ValueAxis>();
        var valAxisTitle = valAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (valAxisTitle != null) node.Format["axisTitle"] = valAxisTitle;

        var catAxis = plotArea.GetFirstChild<C.CategoryAxis>();
        var catAxisTitle = catAxis?.GetFirstChild<C.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (catAxisTitle != null) node.Format["catTitle"] = catAxisTitle;

        // Axis scale
        var scaling = valAxis?.GetFirstChild<C.Scaling>();
        var minVal = scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
        if (minVal != null) node.Format["axisMin"] = minVal;
        var maxVal = scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
        if (maxVal != null) node.Format["axisMax"] = maxVal;

        var majorUnit = valAxis?.GetFirstChild<C.MajorUnit>()?.Val?.Value;
        if (majorUnit != null) node.Format["majorUnit"] = majorUnit;
        var minorUnit = valAxis?.GetFirstChild<C.MinorUnit>()?.Val?.Value;
        if (minorUnit != null) node.Format["minorUnit"] = minorUnit;

        var axisNumFmt = valAxis?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
        if (axisNumFmt != null && axisNumFmt != "General") node.Format["axisNumFmt"] = axisNumFmt;

        var seriesCount = CountSeries(plotArea);
        node.Format["seriesCount"] = seriesCount;

        var cats = ReadCategories(plotArea);
        if (cats != null) node.Format["categories"] = string.Join(",", cats);

        var catsRef = ReadCategoriesRef(plotArea);
        if (catsRef != null) node.Format["categoriesRef"] = catsRef;

        if (depth > 0)
        {
            var seriesList = ReadAllSeries(plotArea);
            for (int i = 0; i < seriesList.Count; i++)
            {
                var (sName, sValues) = seriesList[i];
                var seriesNode = new DocumentNode
                {
                    Path = $"{node.Path}/series[{i + 1}]",
                    Type = "series",
                    Text = sName
                };
                seriesNode.Format["name"] = sName;
                seriesNode.Format["values"] = string.Join(",", sValues.Select(v => v.ToString("G")));

                var serEl = plotArea.Descendants<OpenXmlCompositeElement>()
                    .Where(e => e.LocalName == "ser").ElementAtOrDefault(i);

                // Cell reference formulas (for series with NumberReference/StringReference)
                if (serEl != null)
                {
                    var valRef = ReadFormulaRef(serEl.GetFirstChild<C.Values>());
                    if (valRef != null) seriesNode.Format["valuesRef"] = valRef;
                    var catRef = ReadFormulaRef(serEl.GetFirstChild<C.CategoryAxisData>());
                    if (catRef != null) seriesNode.Format["categoriesRef"] = catRef;
                }

                var serSpPr = serEl?.GetFirstChild<C.ChartShapeProperties>();
                var serColor = serSpPr?.GetFirstChild<Drawing.SolidFill>();
                if (serColor != null)
                {
                    var colorVal = ReadColorFromFill(serColor);
                    if (colorVal != null) seriesNode.Format["color"] = colorVal;
                    // Alpha/transparency
                    var alphaEl = serColor.Descendants<Drawing.Alpha>().FirstOrDefault();
                    if (alphaEl?.Val?.HasValue == true)
                        seriesNode.Format["alpha"] = alphaEl.Val.Value;
                }
                // Gradient
                var gradFill = serSpPr?.GetFirstChild<Drawing.GradientFill>();
                if (gradFill != null) seriesNode.Format["gradient"] = "true";
                // Line width
                var outline = serSpPr?.GetFirstChild<Drawing.Outline>();
                if (outline?.Width?.HasValue == true)
                    seriesNode.Format["lineWidth"] = Math.Round(outline.Width.Value / 12700.0, 2);
                // Line dash
                var prstDash = outline?.GetFirstChild<Drawing.PresetDash>();
                if (prstDash?.Val?.HasValue == true)
                    seriesNode.Format["lineDash"] = prstDash.Val.InnerText;
                // Outline color
                var outlineFill = outline?.GetFirstChild<Drawing.SolidFill>();
                if (outlineFill != null)
                {
                    var outColor = ReadColorFromFill(outlineFill);
                    if (outColor != null) seriesNode.Format["outlineColor"] = outColor;
                }
                // Shadow (from EffectList)
                var effectList = serSpPr?.GetFirstChild<Drawing.EffectList>();
                var outerShadow = effectList?.GetFirstChild<Drawing.OuterShadow>();
                if (outerShadow != null) seriesNode.Format["shadow"] = "true";
                // Marker
                var marker = serEl?.GetFirstChild<C.Marker>();
                var markerSymbol = marker?.GetFirstChild<C.Symbol>()?.Val;
                if (markerSymbol?.HasValue == true)
                    seriesNode.Format["marker"] = markerSymbol.InnerText;
                var markerSize = marker?.GetFirstChild<C.Size>()?.Val;
                if (markerSize?.HasValue == true)
                    seriesNode.Format["markerSize"] = (int)markerSize.Value;
                node.Children.Add(seriesNode);
            }
            node.ChildCount = seriesList.Count;
        }
        else
        {
            node.ChildCount = seriesCount;
        }
    }

    internal static string? DetectChartType(C.PlotArea plotArea)
    {
        var chartTypeCount = plotArea.ChildElements
            .Count(e => e is C.BarChart or C.LineChart or C.PieChart or C.AreaChart
                or C.ScatterChart or C.DoughnutChart or C.Bar3DChart or C.Line3DChart or C.Pie3DChart
                or C.BubbleChart or C.RadarChart or C.StockChart);
        if (chartTypeCount > 1) return "combo";

        if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart bar)
        {
            var dir = bar.GetFirstChild<C.BarDirection>()?.Val?.Value;
            var grp = bar.GetFirstChild<C.BarGrouping>()?.Val?.InnerText;
            var prefix = dir == C.BarDirectionValues.Bar ? "bar" : "column";
            if (grp == "stacked") return $"{prefix}_stacked";
            if (grp == "percentStacked") return $"{prefix}_percentStacked";
            return prefix;
        }
        if (plotArea.GetFirstChild<C.LineChart>() != null) return "line";
        if (plotArea.GetFirstChild<C.PieChart>() != null) return "pie";
        if (plotArea.GetFirstChild<C.DoughnutChart>() != null) return "doughnut";
        if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart area)
        {
            var areaGrp = area.GetFirstChild<C.Grouping>()?.Val?.InnerText;
            if (areaGrp == "stacked") return "area_stacked";
            if (areaGrp == "percentStacked") return "area_percentStacked";
            return "area";
        }
        if (plotArea.GetFirstChild<C.ScatterChart>() != null) return "scatter";
        if (plotArea.GetFirstChild<C.BubbleChart>() != null) return "bubble";
        if (plotArea.GetFirstChild<C.RadarChart>() != null) return "radar";
        if (plotArea.GetFirstChild<C.StockChart>() != null) return "stock";
        if (plotArea.GetFirstChild<C.Bar3DChart>() is C.Bar3DChart bar3d)
        {
            var dir3d = bar3d.GetFirstChild<C.BarDirection>()?.Val?.Value;
            var grp3d = bar3d.GetFirstChild<C.BarGrouping>()?.Val?.InnerText;
            var prefix3d = dir3d == C.BarDirectionValues.Bar ? "bar" : "column";
            var suffix3d = grp3d == "stacked" ? "_stacked"
                : grp3d == "percentStacked" ? "_percentStacked"
                : "";
            return $"{prefix3d}3d{suffix3d}";
        }
        if (plotArea.GetFirstChild<C.Line3DChart>() != null) return "line3d";
        if (plotArea.GetFirstChild<C.Pie3DChart>() != null) return "pie3d";
        return null;
    }

    internal static int CountSeries(C.PlotArea plotArea)
    {
        return plotArea.Descendants<C.Index>()
            .Count(idx => idx.Parent?.LocalName == "ser");
    }

    internal static string[]? ReadCategories(C.PlotArea plotArea)
    {
        var catData = plotArea.Descendants<C.CategoryAxisData>().FirstOrDefault();
        if (catData == null) return null;

        var strLit = catData.GetFirstChild<C.StringLiteral>();
        if (strLit != null)
        {
            return strLit.Elements<C.StringPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        var strRef = catData.GetFirstChild<C.StringReference>();
        var strCache = strRef?.GetFirstChild<C.StringCache>();
        if (strCache != null)
        {
            return strCache.Elements<C.StringPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => p.GetFirstChild<C.NumericValue>()?.Text ?? "")
                .ToArray();
        }

        // StringReference without cache — return null (data lives in cells)
        // The formula is read separately via ReadFormulaRef
        return null;
    }

    /// <summary>
    /// Read the categories formula reference from the first CategoryAxisData element.
    /// Returns null if no reference found (literal categories).
    /// </summary>
    internal static string? ReadCategoriesRef(C.PlotArea plotArea)
    {
        var catData = plotArea.Descendants<C.CategoryAxisData>().FirstOrDefault();
        return ReadFormulaRef(catData);
    }

    internal static List<(string name, double[] values)> ReadAllSeries(C.PlotArea plotArea)
    {
        var result = new List<(string name, double[] values)>();

        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent != null &&
                (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
        {
            var serText = ser.GetFirstChild<C.SeriesText>();
            var name = serText?.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? "?";

            var values = ReadNumericData(ser.GetFirstChild<C.Values>())
                ?? ReadNumericData(ser.Elements<OpenXmlCompositeElement>()
                    .FirstOrDefault(e => e.LocalName == "yVal"))
                ?? Array.Empty<double>();

            result.Add((name, values));
        }

        return result;
    }

    internal static double[]? ReadNumericData(OpenXmlCompositeElement? valElement)
    {
        if (valElement == null) return null;

        var numLit = valElement.GetFirstChild<C.NumberLiteral>();
        if (numLit != null)
        {
            return numLit.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => double.TryParse(p.GetFirstChild<C.NumericValue>()?.Text, out var v) ? v : 0)
                .ToArray();
        }

        var numRef = valElement.GetFirstChild<C.NumberReference>();
        var numCache = numRef?.GetFirstChild<C.NumberingCache>();
        if (numCache != null)
        {
            return numCache.Elements<C.NumericPoint>()
                .OrderBy(p => p.Index?.Value ?? 0)
                .Select(p => double.TryParse(p.GetFirstChild<C.NumericValue>()?.Text, out var v) ? v : 0)
                .ToArray();
        }

        // NumberReference without cache — return empty array (data lives in cells)
        if (numRef != null) return Array.Empty<double>();

        return null;
    }

    /// <summary>
    /// Read the formula string from a NumberReference or StringReference inside a Values/CategoryAxisData element.
    /// Returns null if no reference found.
    /// </summary>
    internal static string? ReadFormulaRef(OpenXmlCompositeElement? element)
    {
        if (element == null) return null;
        var numRef = element.GetFirstChild<C.NumberReference>();
        if (numRef != null) return numRef.GetFirstChild<C.Formula>()?.Text;
        var strRef = element.GetFirstChild<C.StringReference>();
        if (strRef != null) return strRef.GetFirstChild<C.Formula>()?.Text;
        return null;
    }

    internal static string? ReadColorFromFill(Drawing.SolidFill? solidFill)
    {
        if (solidFill == null) return null;
        var rgb = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null) return ParseHelpers.FormatHexColor(rgb);
        var scheme = solidFill.GetFirstChild<Drawing.SchemeColor>()?.Val;
        if (scheme?.HasValue == true) return scheme.InnerText;
        return null;
    }

    /// <summary>
    /// Read gridline detail into separate format keys: {prefix}Color, {prefix}Width, {prefix}Dash.
    /// </summary>
    private static void ReadGridlineDetail(OpenXmlCompositeElement gridlines, DocumentNode node, string prefix)
    {
        var spPr = gridlines.GetFirstChild<C.ChartShapeProperties>();
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
        if (outline == null) return;

        var fill = outline.GetFirstChild<Drawing.SolidFill>();
        var color = ReadColorFromFill(fill);
        if (color != null) node.Format[$"{prefix}Color"] = color;

        if (outline.Width?.HasValue == true)
            node.Format[$"{prefix}Width"] = Math.Round(outline.Width.Value / 12700.0, 2);

        var dash = outline.GetFirstChild<Drawing.PresetDash>()?.Val;
        if (dash?.HasValue == true)
            node.Format[$"{prefix}Dash"] = dash.InnerText!;
    }

    /// <summary>
    /// Read font spec from TextProperties: returns "SIZE:COLOR:FONTNAME" format or null.
    /// </summary>
    private static string? ReadFontSpec(C.TextProperties textProperties)
    {
        var defRp = textProperties.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
        if (defRp == null) return null;

        var parts = new List<string>();
        if (defRp.FontSize?.HasValue == true)
            parts.Add((defRp.FontSize.Value / 100.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture));
        else
            parts.Add("");

        var fill = defRp.GetFirstChild<Drawing.SolidFill>();
        var color = ReadColorFromFill(fill);
        parts.Add(color?.TrimStart('#') ?? "");

        var font = defRp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
        if (font != null)
            parts.Add(font);

        var result = string.Join(":", parts).TrimEnd(':');
        return string.IsNullOrEmpty(result) ? null : result;
    }

    // ==================== Chart Set ====================

    internal static void UpdateSeriesData(C.PlotArea plotArea, List<(string name, double[] values)> newData)
    {
        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        for (int i = 0; i < Math.Min(newData.Count, allSer.Count); i++)
        {
            var ser = allSer[i];
            var (sName, sVals) = newData[i];

            var serText = ser.GetFirstChild<C.SeriesText>();
            if (serText != null)
            {
                serText.RemoveAllChildren();
                serText.AppendChild(new C.NumericValue(sName));
            }

            var valEl = ser.GetFirstChild<C.Values>();
            if (valEl != null)
            {
                valEl.RemoveAllChildren();
                var builtVals = BuildValues(sVals);
                foreach (var child in builtVals.ChildElements.ToList())
                    valEl.AppendChild(child.CloneNode(true));
            }
        }
    }
}
