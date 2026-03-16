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

        var colors = ParseSeriesColors(properties);

        switch (kind)
        {
            case "bar":
                chartElement = BuildBarChart(C.BarDirectionValues.Bar, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "column":
                chartElement = BuildBarChart(C.BarDirectionValues.Column, stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "line":
                chartElement = BuildLineChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "area":
                chartElement = BuildAreaChart(stacked, percentStacked,
                    categories, seriesData, catAxisId, valAxisId, colors);
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
            case "combo":
            {
                // Combo: first N series as columns, rest as lines
                int splitAt = 1;
                if (properties.TryGetValue("combosplit", out var splitStr))
                    splitAt = int.Parse(splitStr);
                splitAt = Math.Min(splitAt, seriesData.Count);

                var barData = seriesData.Take(splitAt).ToList();
                var lineData = seriesData.Skip(splitAt).ToList();

                var comboBar = new C.BarChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Column },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false }
                );
                for (int ci = 0; ci < barData.Count; ci++)
                {
                    var clr = colors != null && ci < colors.Length ? colors[ci] : DefaultSeriesColors[ci % DefaultSeriesColors.Length];
                    comboBar.AppendChild(BuildBarSeries((uint)ci, barData[ci].name, categories, barData[ci].values, clr));
                }
                comboBar.AppendChild(new C.AxisId { Val = catAxisId });
                comboBar.AppendChild(new C.AxisId { Val = valAxisId });
                plotArea.AppendChild(comboBar);

                if (lineData.Count > 0)
                {
                    var comboLine = new C.LineChart(
                        new C.Grouping { Val = C.GroupingValues.Standard },
                        new C.VaryColors { Val = false }
                    );
                    for (int ci = 0; ci < lineData.Count; ci++)
                    {
                        var sIdx = (uint)(splitAt + ci);
                        var cIdx = splitAt + ci;
                        var clr = colors != null && cIdx < colors.Length ? colors[cIdx] : DefaultSeriesColors[cIdx % DefaultSeriesColors.Length];
                        comboLine.AppendChild(BuildLineSeries(sIdx, lineData[ci].name, categories, lineData[ci].values, clr));
                    }
                    comboLine.AppendChild(new C.ShowMarker { Val = true });
                    comboLine.AppendChild(new C.AxisId { Val = catAxisId });
                    comboLine.AppendChild(new C.AxisId { Val = valAxisId });
                    plotArea.AppendChild(comboLine);
                }
                chartElement = null; // already added
                break;
            }
            default:
                // Default to column chart
                chartElement = BuildBarChart(C.BarDirectionValues.Column, false, false,
                    categories, seriesData, catAxisId, valAxisId, colors);
                break;
        }

        if (chartElement != null)
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
        uint catAxisId, uint valAxisId, string[]? colors = null)
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
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            barChart.AppendChild(BuildBarSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
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
        uint catAxisId, uint valAxisId, string[]? colors = null)
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
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            lineChart.AppendChild(BuildLineSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
        }

        lineChart.AppendChild(new C.ShowMarker { Val = true });
        lineChart.AppendChild(new C.AxisId { Val = catAxisId });
        lineChart.AppendChild(new C.AxisId { Val = valAxisId });
        return lineChart;
    }

    private static C.AreaChart BuildAreaChart(
        bool stacked, bool percentStacked,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
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
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            areaChart.AppendChild(BuildAreaSeries((uint)i, seriesData[i].name,
                categories, seriesData[i].values, color));
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

    // ==================== Default Series Colors ====================

    private static readonly string[] DefaultSeriesColors =
    {
        "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
        "264478", "9B4A22", "636363", "BF8F00", "3A75A8", "4E8538"
    };

    /// <summary>
    /// Apply a color to a series via ChartShapeProperties.
    /// </summary>
    private static void ApplySeriesColor(OpenXmlCompositeElement series, string color)
    {
        series.RemoveAllChildren<C.ChartShapeProperties>();
        var spPr = new C.ChartShapeProperties();
        spPr.AppendChild(BuildSolidFill(color));
        // Insert after SeriesText (schema order: idx, order, tx, spPr, ...)
        var serText = series.GetFirstChild<C.SeriesText>();
        if (serText != null)
            serText.InsertAfterSelf(spPr);
        else
            series.PrependChild(spPr);
    }

    /// <summary>
    /// Parse series colors from properties. Format: "colors=FF0000,00FF00,0000FF" or "color1=FF0000"
    /// </summary>
    private static string[]? ParseSeriesColors(Dictionary<string, string> properties)
    {
        if (properties.TryGetValue("colors", out var colorsStr))
            return colorsStr.Split(',').Select(c => c.Trim()).ToArray();
        return null;
    }

    // ==================== Series Builders ====================

    private static C.BarChartSeries BuildBarSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    private static C.LineChartSeries BuildLineSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.LineChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    private static C.AreaChartSeries BuildAreaSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.AreaChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    private static C.PieChartSeries BuildPieSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.PieChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(new C.NumericValue(name)) { Index = 0 }
            ))
        );
        if (color != null) ApplySeriesColor(series, color);
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

        // Data labels
        var dataLabels = plotArea.Descendants<C.DataLabels>().FirstOrDefault();
        if (dataLabels != null)
        {
            var parts = new List<string>();
            if (dataLabels.GetFirstChild<C.ShowValue>()?.Val?.Value == true) parts.Add("value");
            if (dataLabels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true) parts.Add("category");
            if (dataLabels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true) parts.Add("series");
            if (dataLabels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true) parts.Add("percent");
            if (parts.Count > 0) node.Format["dataLabels"] = string.Join(",", parts);
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
                // Series color
                var serEl = plotArea.Descendants<OpenXmlCompositeElement>()
                    .Where(e => e.LocalName == "ser").ElementAtOrDefault(i);
                var serColor = serEl?.GetFirstChild<C.ChartShapeProperties>()
                    ?.GetFirstChild<Drawing.SolidFill>();
                if (serColor != null)
                {
                    var colorVal = ReadColorFromFill(serColor);
                    if (colorVal != null) seriesNode.Format["color"] = colorVal;
                }
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
        // Combo detection: multiple chart type elements
        var chartTypeCount = plotArea.ChildElements
            .Count(e => e is C.BarChart or C.LineChart or C.PieChart or C.AreaChart
                or C.ScatterChart or C.DoughnutChart or C.Bar3DChart or C.Line3DChart or C.Pie3DChart);
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
    /// Update series data in a plot area, matching by index.
    /// </summary>
    private static void UpdateSeriesData(C.PlotArea plotArea, List<(string name, double[] values)> newData)
    {
        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        for (int i = 0; i < Math.Min(newData.Count, allSer.Count); i++)
        {
            var ser = allSer[i];
            var (sName, sVals) = newData[i];

            // Update name
            var serText = ser.GetFirstChild<C.SeriesText>();
            if (serText != null)
            {
                serText.RemoveAllChildren();
                serText.AppendChild(new C.StringLiteral(
                    new C.PointCount { Val = 1 },
                    new C.StringPoint(new C.NumericValue(sName)) { Index = 0 }
                ));
            }

            // Update values
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

                case "datalabels" or "labels":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    // Apply to each chart type element's DataLabels
                    foreach (var chartTypeEl in plotArea2.ChildElements
                        .Where(e => e.LocalName.Contains("Chart") || e.LocalName.Contains("chart")))
                    {
                        chartTypeEl.RemoveAllChildren<C.DataLabels>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var dl = new C.DataLabels();
                            var parts = value.ToLowerInvariant().Split(',').Select(s => s.Trim()).ToHashSet();
                            dl.AppendChild(new C.ShowValue { Val = parts.Contains("value") || parts.Contains("true") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowCategoryName { Val = parts.Contains("category") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowSeriesName { Val = parts.Contains("series") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowPercent { Val = parts.Contains("percent") || parts.Contains("all") });
                            dl.AppendChild(new C.ShowLegendKey { Val = false });
                            chartTypeEl.AppendChild(dl);
                        }
                    }
                    break;
                }

                case "colors":
                {
                    // Set series colors: "FF0000,00FF00,0000FF"
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
                        valAxis.InsertAfter(BuildChartTitle(value), valAxis.GetFirstChild<C.Scaling>());
                    break;
                }

                case "cattitle" or "htitle":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var catAxis = plotArea2?.GetFirstChild<C.CategoryAxis>();
                    if (catAxis == null) { unsupported.Add(key); break; }
                    catAxis.RemoveAllChildren<C.Title>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        catAxis.InsertAfter(BuildChartTitle(value), catAxis.GetFirstChild<C.Scaling>());
                    break;
                }

                case "axismin" or "min":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAxis?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.MinAxisValue>();
                    scaling.AppendChild(new C.MinAxisValue { Val = double.Parse(value) });
                    break;
                }

                case "axismax" or "max":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    var scaling = valAxis?.GetFirstChild<C.Scaling>();
                    if (scaling == null) { unsupported.Add(key); break; }
                    scaling.RemoveAllChildren<C.MaxAxisValue>();
                    scaling.AppendChild(new C.MaxAxisValue { Val = double.Parse(value) });
                    break;
                }

                case "majorunit":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.MajorUnit>();
                    valAxis.AppendChild(new C.MajorUnit { Val = double.Parse(value) });
                    break;
                }

                case "minorunit":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.MinorUnit>();
                    valAxis.AppendChild(new C.MinorUnit { Val = double.Parse(value) });
                    break;
                }

                case "axisnumfmt" or "axisnumberformat":
                {
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    var valAxis = plotArea2?.GetFirstChild<C.ValueAxis>();
                    if (valAxis == null) { unsupported.Add(key); break; }
                    valAxis.RemoveAllChildren<C.NumberingFormat>();
                    valAxis.AppendChild(new C.NumberingFormat { FormatCode = value, SourceLinked = false });
                    break;
                }

                case "categories":
                {
                    // Update categories across all series
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
                    // Replace series data: "S1:1,2,3;S2:4,5,6"
                    var plotArea2 = chart.GetFirstChild<C.PlotArea>();
                    if (plotArea2 == null) { unsupported.Add(key); break; }
                    var newSeries = ParseSeriesData(new Dictionary<string, string> { ["data"] = value });
                    UpdateSeriesData(plotArea2, newSeries);
                    break;
                }

                default:
                    // Try numbered series: series1="Name:1,2,3"
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
                            vals = value[(colonIdx + 1)..].Split(',').Select(v => double.Parse(v.Trim())).ToArray();
                            // Update series name
                            var serText = ser.GetFirstChild<C.SeriesText>();
                            if (serText != null)
                            {
                                serText.RemoveAllChildren();
                                serText.AppendChild(new C.StringLiteral(
                                    new C.PointCount { Val = 1 },
                                    new C.StringPoint(new C.NumericValue(sName)) { Index = 0 }
                                ));
                            }
                        }
                        else
                        {
                            vals = value.Split(',').Select(v => double.Parse(v.Trim())).ToArray();
                        }

                        // Update values
                        var valEl = ser.GetFirstChild<C.Values>();
                        if (valEl != null)
                        {
                            valEl.RemoveAllChildren();
                            var builtVals = BuildValues(vals);
                            foreach (var child in builtVals.ChildElements.ToList())
                                valEl.AppendChild(child.CloneNode(true));
                        }
                        // Also try YValues for scatter charts
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
}
