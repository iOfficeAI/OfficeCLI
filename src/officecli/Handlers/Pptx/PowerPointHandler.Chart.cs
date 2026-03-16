// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Chart Helpers ====================

    /// <summary>
    /// Parse chart type string to determine chart kind and variant.
    /// </summary>
    private static (string kind, bool is3D, bool stacked, bool percentStacked) ParseChartType(string chartType)
    {
        var ct = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
        var is3D = ct.EndsWith("3d") || ct.Contains("3d");
        ct = ct.Replace("3d", "");

        var stacked = ct.Contains("stacked") && !ct.Contains("percent");
        var percentStacked = ct.Contains("percentstacked") || ct.Contains("pstacked");
        ct = ct.Replace("percentstacked", "").Replace("pstacked", "").Replace("stacked", "");

        var kind = ct switch
        {
            "bar" => "bar",
            "column" or "col" => "column",
            "line" => "line",
            "pie" => "pie",
            "doughnut" or "donut" => "doughnut",
            "area" => "area",
            "scatter" or "xy" => "scatter",
            _ => ct
        };

        return (kind, is3D, stacked, percentStacked);
    }

    /// <summary>
    /// Parse series data from properties.
    /// Supports:
    ///   data = "Series1:1,2,3;Series2:4,5,6"
    ///   series1 = "Revenue:100,200,300"
    ///   series2 = "Cost:80,150,250"
    /// </summary>
    private static List<(string name, double[] values)> ParseSeriesData(Dictionary<string, string> properties)
    {
        var result = new List<(string name, double[] values)>();

        // Try compact "data" format first
        if (properties.TryGetValue("data", out var dataStr))
        {
            foreach (var seriesPart in dataStr.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var colonIdx = seriesPart.IndexOf(':');
                if (colonIdx < 0) continue;
                var name = seriesPart[..colonIdx].Trim();
                var vals = seriesPart[(colonIdx + 1)..].Split(',')
                    .Select(v => double.Parse(v.Trim())).ToArray();
                result.Add((name, vals));
            }
            return result;
        }

        // Try numbered series: series1, series2, ...
        for (int i = 1; i <= 20; i++)
        {
            if (!properties.TryGetValue($"series{i}", out var seriesStr)) break;
            var colonIdx = seriesStr.IndexOf(':');
            if (colonIdx < 0)
            {
                // No name, use "Series N"
                var vals = seriesStr.Split(',').Select(v => double.Parse(v.Trim())).ToArray();
                result.Add(($"Series {i}", vals));
            }
            else
            {
                var name = seriesStr[..colonIdx].Trim();
                var vals = seriesStr[(colonIdx + 1)..].Split(',')
                    .Select(v => double.Parse(v.Trim())).ToArray();
                result.Add((name, vals));
            }
        }

        return result;
    }

    /// <summary>
    /// Parse categories from properties. Returns null if not specified.
    /// </summary>
    private static string[]? ParseCategories(Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("categories", out var catStr)) return null;
        return catStr.Split(',').Select(c => c.Trim()).ToArray();
    }

    /// <summary>
    /// Build a complete ChartSpace from chart type, categories, and series data.
    /// </summary>
    private static C.ChartSpace BuildChartSpace(
        string chartType,
        string? title,
        string[]? categories,
        List<(string name, double[] values)> seriesData,
        Dictionary<string, string> properties)
    {
        var (kind, is3D, stacked, percentStacked) = ParseChartType(chartType);

        var chartSpace = new C.ChartSpace();
        var chart = new C.Chart();

        // Title
        if (!string.IsNullOrEmpty(title))
            chart.AppendChild(BuildChartTitle(title));

        // Auto-generate categories if not provided
        if (categories == null && seriesData.Count > 0)
        {
            var maxLen = seriesData.Max(s => s.values.Length);
            categories = Enumerable.Range(1, maxLen).Select(i => i.ToString()).ToArray();
        }

        // PlotArea
        var plotArea = new C.PlotArea(new C.Layout());

        uint catAxisId = 1;
        uint valAxisId = 2;

        // Build chart type element with series
        OpenXmlCompositeElement chartElement;
        bool needsAxes = true;

        switch (kind)
        {
            case "bar":
                chartElement = BuildBarChart(C.BarDirectionValues.Bar, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId);
                break;
            case "column":
                chartElement = BuildBarChart(C.BarDirectionValues.Column, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId);
                break;
            case "line":
                chartElement = BuildLineChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId);
                break;
            case "area":
                chartElement = BuildAreaChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId);
                break;
            case "pie":
                chartElement = BuildPieChart(categories, seriesData);
                needsAxes = false;
                break;
            case "doughnut":
                chartElement = BuildDoughnutChart(categories, seriesData);
                needsAxes = false;
                break;
            case "scatter":
                chartElement = BuildScatterChart(categories, seriesData, catAxisId, valAxisId);
                break;
            default:
                // Default to column chart
                chartElement = BuildBarChart(C.BarDirectionValues.Column, false, false,
                    categories, seriesData, catAxisId, valAxisId);
                break;
        }

        plotArea.AppendChild(chartElement);

        // Add axes for chart types that need them
        if (needsAxes)
        {
            if (kind == "scatter")
            {
                // Scatter uses two value axes
                plotArea.AppendChild(BuildValueAxis(catAxisId, valAxisId, C.AxisPositionValues.Bottom));
                plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));
            }
            else
            {
                plotArea.AppendChild(BuildCategoryAxis(catAxisId, valAxisId));
                plotArea.AppendChild(BuildValueAxis(valAxisId, catAxisId, C.AxisPositionValues.Left));
            }
        }

        chart.AppendChild(plotArea);

        // Legend
        var showLegend = properties.GetValueOrDefault("legend", "true");
        if (!showLegend.Equals("false", StringComparison.OrdinalIgnoreCase) &&
            !showLegend.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var legendPos = showLegend.ToLowerInvariant() switch
            {
                "top" or "t" => C.LegendPositionValues.Top,
                "left" or "l" => C.LegendPositionValues.Left,
                "right" or "r" => C.LegendPositionValues.Right,
                "bottom" or "b" => C.LegendPositionValues.Bottom,
                _ => C.LegendPositionValues.Bottom
            };
            chart.AppendChild(new C.Legend(
                new C.LegendPosition { Val = legendPos },
                new C.Overlay { Val = false }
            ));
        }

        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chart.AppendChild(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        chartSpace.AppendChild(chart);
        return chartSpace;
    }

    // ==================== Chart Type Builders ====================

    private static C.BarChart BuildBarChart(
        C.BarDirectionValues direction, bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        var grouping = percentStacked ? C.BarGroupingValues.PercentStacked
            : stacked ? C.BarGroupingValues.Stacked
            : C.BarGroupingValues.Clustered;

        var barChart = new C.BarChart(
            new C.BarDirection { Val = direction },
            new C.BarGrouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            barChart.AppendChild(BuildBarSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values));
        }

        barChart.AppendChild(new C.GapWidth { Val = (ushort)150 });
        if (stacked || percentStacked)
            barChart.AppendChild(new C.Overlap { Val = 100 });
        barChart.AppendChild(new C.AxisId { Val = catAxisId });
        barChart.AppendChild(new C.AxisId { Val = valAxisId });
        return barChart;
    }

    private static C.LineChart BuildLineChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        var grouping = percentStacked ? C.GroupingValues.PercentStacked
            : stacked ? C.GroupingValues.Stacked
            : C.GroupingValues.Standard;

        var lineChart = new C.LineChart(
            new C.Grouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            lineChart.AppendChild(BuildLineSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values));
        }

        lineChart.AppendChild(new C.ShowMarker { Val = true });
        lineChart.AppendChild(new C.AxisId { Val = catAxisId });
        lineChart.AppendChild(new C.AxisId { Val = valAxisId });
        return lineChart;
    }

    private static C.AreaChart BuildAreaChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        var grouping = percentStacked ? C.GroupingValues.PercentStacked
            : stacked ? C.GroupingValues.Stacked
            : C.GroupingValues.Standard;

        var areaChart = new C.AreaChart(
            new C.Grouping { Val = grouping },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            areaChart.AppendChild(BuildAreaSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values));
        }

        areaChart.AppendChild(new C.AxisId { Val = catAxisId });
        areaChart.AppendChild(new C.AxisId { Val = valAxisId });
        return areaChart;
    }

    private static C.PieChart BuildPieChart(
        string[]? categories, List<(string name, double[] values)> seriesData)
    {
        var pieChart = new C.PieChart(new C.VaryColors { Val = true });
        // Pie charts typically use only the first series
        if (seriesData.Count > 0)
            pieChart.AppendChild(BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values));
        return pieChart;
    }

    private static C.DoughnutChart BuildDoughnutChart(
        string[]? categories, List<(string name, double[] values)> seriesData)
    {
        var chart = new C.DoughnutChart(new C.VaryColors { Val = true });
        if (seriesData.Count > 0)
            chart.AppendChild(BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values));
        chart.AppendChild(new C.HoleSize { Val = 50 });
        return chart;
    }

    private static C.ScatterChart BuildScatterChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        var scatterChart = new C.ScatterChart(
            new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
            new C.VaryColors { Val = false }
        );

        // For scatter: categories are X values, series values are Y values
        double[]? xValues = null;
        if (categories != null)
            xValues = categories.Select(c => double.TryParse(c, out var v) ? v : 0).ToArray();

        for (int i = 0; i < seriesData.Count; i++)
        {
            scatterChart.AppendChild(BuildScatterSeries((uint)i, seriesData[i].name,
                xValues, seriesData[i].values));
        }

        scatterChart.AppendChild(new C.AxisId { Val = catAxisId });
        scatterChart.AppendChild(new C.AxisId { Val = valAxisId });
        return scatterChart;
    }

    // ==================== Series Builders ====================

    private static C.BarChartSeries BuildBarSeries(uint idx, string name,
        string[]? categories, double[] values)
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    private static C.LineChartSeries BuildLineSeries(uint idx, string name,
        string[]? categories, double[] values)
    {
        var series = new C.LineChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    private static C.AreaChartSeries BuildAreaSeries(uint idx, string name,
        string[]? categories, double[] values)
    {
        var series = new C.AreaChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    private static C.PieChartSeries BuildPieSeries(uint idx, string name,
        string[]? categories, double[] values)
    {
        var series = new C.PieChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    private static C.ScatterChartSeries BuildScatterSeries(uint idx, string name,
        double[]? xValues, double[] yValues)
    {
        var series = new C.ScatterChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );

        // XValues
        if (xValues != null)
        {
            var xLit = new C.NumberLiteral(new C.PointCount { Val = (uint)xValues.Length });
            for (int i = 0; i < xValues.Length; i++)
                xLit.AppendChild(new C.NumericPoint(new C.NumericValue(xValues[i].ToString("G"))) { Index = (uint)i });
            series.AppendChild(new C.XValues(xLit));
        }

        // YValues
        var yLit = new C.NumberLiteral(new C.PointCount { Val = (uint)yValues.Length });
        for (int i = 0; i < yValues.Length; i++)
            yLit.AppendChild(new C.NumericPoint(new C.NumericValue(yValues[i].ToString("G"))) { Index = (uint)i });
        series.AppendChild(new C.YValues(yLit));

        return series;
    }

    // ==================== Data Builders ====================

    private static C.CategoryAxisData BuildCategoryData(string[] categories)
    {
        var strLit = new C.StringLiteral(new C.PointCount { Val = (uint)categories.Length });
        for (int i = 0; i < categories.Length; i++)
            strLit.AppendChild(new C.StringPoint(new C.NumericValue(categories[i])) { Index = (uint)i });
        return new C.CategoryAxisData(strLit);
    }

    private static C.Values BuildValues(double[] values)
    {
        var numLit = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Length }
        );
        for (int i = 0; i < values.Length; i++)
            numLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[i].ToString("G"))) { Index = (uint)i });
        return new C.Values(numLit);
    }

    // ==================== Axis Builders ====================

    private static C.CategoryAxis BuildCategoryAxis(uint axisId, uint crossAxisId)
    {
        return new C.CategoryAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.MajorTickMark { Val = C.TickMarkValues.Outside },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 }
        );
    }

    private static C.ValueAxis BuildValueAxis(uint axisId, uint crossAxisId, C.AxisPositionValues position)
    {
        return new C.ValueAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = position },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark { Val = C.TickMarkValues.Outside },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between }
        );
    }

    // ==================== Title Builder ====================

    private static C.Title BuildChartTitle(string titleText)
    {
        return new C.Title(
            new C.ChartText(
                new C.RichText(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties(
                            new Drawing.DefaultRunProperties { FontSize = 1400, Bold = true }
                        ),
                        new Drawing.Run(
                            new Drawing.RunProperties { Language = "zh-CN", FontSize = 1400, Bold = true },
                            new Drawing.Text(titleText)
                        )
                    )
                )
            ),
            new C.Overlay { Val = false }
        );
    }

    // ==================== Chart GraphicFrame Builder ====================

    /// <summary>
    /// Create a GraphicFrame embedding a chart and add it to the slide's shape tree.
    /// </summary>
    private static GraphicFrame BuildChartGraphicFrame(
        SlidePart slidePart, ChartPart chartPart, uint shapeId, string name,
        long x, long y, long cx, long cy)
    {
        var relId = slidePart.GetIdOfPart(chartPart);

        var graphicFrame = new GraphicFrame();
        graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
            new NonVisualDrawingProperties { Id = shapeId, Name = name },
            new NonVisualGraphicFrameDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()
        );
        graphicFrame.Transform = new Transform(
            new Drawing.Offset { X = x, Y = y },
            new Drawing.Extents { Cx = cx, Cy = cy }
        );

        var chartRef = new C.ChartReference { Id = relId };
        graphicFrame.AppendChild(new Drawing.Graphic(
            new Drawing.GraphicData(chartRef)
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
            }
        ));

        return graphicFrame;
    }

    // ==================== Chart Readback ====================

    /// <summary>
    /// Build a DocumentNode from a chart GraphicFrame.
    /// </summary>
    private static DocumentNode ChartToNode(GraphicFrame gf, SlidePart slidePart, int slideNum, int chartIdx, int depth)
    {
        var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Chart";

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/chart[{chartIdx}]",
            Type = "chart",
            Preview = name
        };

        node.Format["name"] = name;

        // Position
        var offset = gf.Transform?.Offset;
        if (offset != null)
        {
            if (offset.X is not null) node.Format["x"] = FormatEmu(offset.X!);
            if (offset.Y is not null) node.Format["y"] = FormatEmu(offset.Y!);
        }
        var extents = gf.Transform?.Extents;
        if (extents != null)
        {
            if (extents.Cx is not null) node.Format["width"] = FormatEmu(extents.Cx!);
            if (extents.Cy is not null) node.Format["height"] = FormatEmu(extents.Cy!);
        }

        // Read chart data from ChartPart
        var chartRef = gf.Descendants<C.ChartReference>().FirstOrDefault();
        if (chartRef?.Id?.Value != null)
        {
            try
            {
                var chartPart = (ChartPart)slidePart.GetPartById(chartRef.Id.Value);
                var chartSpace = chartPart.ChartSpace;
                var chart = chartSpace?.GetFirstChild<C.Chart>();
                if (chart != null)
                    ReadChartProperties(chart, node, depth);
            }
            catch { }
        }

        return node;
    }

    /// <summary>
    /// Read chart properties into a DocumentNode's Format dictionary.
    /// </summary>
    private static void ReadChartProperties(C.Chart chart, DocumentNode node, int depth)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return;

        // Detect chart type
        var chartType = DetectChartType(plotArea);
        if (chartType != null) node.Format["chartType"] = chartType;

        // Title
        var titleEl = chart.GetFirstChild<C.Title>();
        var titleText = titleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        if (titleText != null) node.Format["title"] = titleText;

        // Legend
        var legend = chart.GetFirstChild<C.Legend>();
        if (legend != null)
        {
            var pos = legend.GetFirstChild<C.LegendPosition>()?.Val?.HasValue == true
                ? legend.GetFirstChild<C.LegendPosition>()!.Val!.InnerText : "b";
            node.Format["legend"] = pos;
        }

        // Series count
        var seriesCount = CountSeries(plotArea);
        node.Format["seriesCount"] = seriesCount;

        // Categories
        var cats = ReadCategories(plotArea);
        if (cats != null) node.Format["categories"] = string.Join(",", cats);

        // Series data at depth > 0
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
                node.Children.Add(seriesNode);
            }
            node.ChildCount = seriesList.Count;
        }
        else
        {
            node.ChildCount = seriesCount;
        }
    }

    private static string? DetectChartType(C.PlotArea plotArea)
    {
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
        if (plotArea.GetFirstChild<C.Bar3DChart>() != null) return "bar3d";
        if (plotArea.GetFirstChild<C.Line3DChart>() != null) return "line3d";
        if (plotArea.GetFirstChild<C.Pie3DChart>() != null) return "pie3d";
        return null;
    }

    private static int CountSeries(C.PlotArea plotArea)
    {
        // Count series across all chart type elements
        return plotArea.Descendants<C.Index>()
            .Count(idx => idx.Parent?.LocalName == "ser");
    }

    private static string[]? ReadCategories(C.PlotArea plotArea)
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

    private static List<(string name, double[] values)> ReadAllSeries(C.PlotArea plotArea)
    {
        var result = new List<(string name, double[] values)>();

        // Gather all series elements regardless of chart type
        foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent != null &&
                (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
        {
            // Series name
            var serText = ser.GetFirstChild<C.SeriesText>();
            var name = serText?.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? "?";

            // Values — try Values (bar/line/pie/area) or YValues (scatter)
            var values = ReadNumericData(ser.GetFirstChild<C.Values>())
                ?? ReadNumericData(ser.Elements<OpenXmlCompositeElement>()
                    .FirstOrDefault(e => e.LocalName == "yVal"))
                ?? Array.Empty<double>();

            result.Add((name, values));
        }

        return result;
    }

    private static double[]? ReadNumericData(OpenXmlCompositeElement? valElement)
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

    // ==================== Chart Set ====================

    /// <summary>
    /// Modify chart properties. Supports: title, legend.
    /// </summary>
    private static List<string> SetChartProperties(ChartPart chartPart, Dictionary<string, string> properties)
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

                default:
                    unsupported.Add(key);
                    break;
            }
        }

        chartSpace!.Save();
        return unsupported;
    }
}
