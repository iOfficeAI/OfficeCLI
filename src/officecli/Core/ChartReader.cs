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

        var legend = chart.GetFirstChild<C.Legend>();
        if (legend != null)
        {
            var pos = legend.GetFirstChild<C.LegendPosition>()?.Val?.HasValue == true
                ? legend.GetFirstChild<C.LegendPosition>()!.Val!.InnerText : "b";
            node.Format["legend"] = pos;
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
        if (style?.Val?.HasValue == true) node.Format["style"] = style.Val.Value;

        // Plot area fill (plotArea uses C.ShapeProperties, not C.ChartShapeProperties)
        var plotSpPr = plotArea.GetFirstChild<C.ShapeProperties>();
        var plotFill = plotSpPr?.GetFirstChild<Drawing.SolidFill>();
        if (plotFill != null)
        {
            var pColor = ReadColorFromFill(plotFill);
            if (pColor != null) node.Format["plotFill"] = pColor;
        }

        // Gridlines
        var valAxisForGrid = plotArea.GetFirstChild<C.ValueAxis>();
        if (valAxisForGrid?.GetFirstChild<C.MajorGridlines>() != null) node.Format["gridlines"] = "true";
        if (valAxisForGrid?.GetFirstChild<C.MinorGridlines>() != null) node.Format["minorGridlines"] = "true";

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
        if (plotArea.GetFirstChild<C.AreaChart>() != null) return "area";
        if (plotArea.GetFirstChild<C.ScatterChart>() != null) return "scatter";
        if (plotArea.GetFirstChild<C.BubbleChart>() != null) return "bubble";
        if (plotArea.GetFirstChild<C.RadarChart>() != null) return "radar";
        if (plotArea.GetFirstChild<C.StockChart>() != null) return "stock";
        if (plotArea.GetFirstChild<C.Bar3DChart>() != null) return "bar3d";
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

        return null;
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
