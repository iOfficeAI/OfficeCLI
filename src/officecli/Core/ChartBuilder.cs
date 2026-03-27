// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal static partial class ChartHelper
{
    // ==================== Build ChartSpace ====================

    internal static C.ChartSpace BuildChartSpace(
        string chartType,
        string? title,
        string[]? categories,
        List<(string name, double[] values)> seriesData,
        Dictionary<string, string> properties)
    {
        var (kind, is3D, stacked, percentStacked) = ParseChartType(chartType);

        var chartSpace = new C.ChartSpace();
        var chart = new C.Chart();

        if (!string.IsNullOrEmpty(title))
            chart.AppendChild(BuildChartTitle(title));

        var originalCategories = categories;
        if (categories == null && seriesData.Count > 0)
        {
            var maxLen = seriesData.Max(s => s.values.Length);
            categories = Enumerable.Range(1, maxLen).Select(i => i.ToString()).ToArray();
        }

        var plotArea = new C.PlotArea(new C.Layout());
        uint catAxisId = 1;
        uint valAxisId = 2;

        OpenXmlCompositeElement? chartElement;
        bool needsAxes = true;

        var colors = ParseSeriesColors(properties);

        switch (kind)
        {
            case "bar" when is3D:
            case "column" when is3D:
            {
                var dir3dAuto = kind == "bar" ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column;
                var bar3dAuto = new C.Bar3DChart(
                    new C.BarDirection { Val = dir3dAuto },
                    new C.BarGrouping { Val = stacked ? C.BarGroupingValues.Stacked
                        : percentStacked ? C.BarGroupingValues.PercentStacked
                        : C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false }
                );
                for (int si = 0; si < seriesData.Count; si++)
                {
                    var s = BuildBarSeries((uint)si, seriesData[si].name, categories, seriesData[si].values,
                        colors != null && si < colors.Length ? colors[si] : null);
                    bar3dAuto.AppendChild(s);
                }
                bar3dAuto.AppendChild(new C.GapWidth { Val = 150 });
                bar3dAuto.AppendChild(new C.AxisId { Val = catAxisId });
                bar3dAuto.AppendChild(new C.AxisId { Val = valAxisId });
                chartElement = bar3dAuto;
                break;
            }
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
            case "bubble":
                chartElement = BuildBubbleChart(categories, seriesData, catAxisId, valAxisId, colors);
                break;
            case "radar":
            {
                var radarStyle = properties.GetValueOrDefault("radarStyle", "marker");
                chartElement = BuildRadarChart(radarStyle, categories, seriesData, catAxisId, valAxisId, colors);
                break;
            }
            // Note: column3d/bar3d are handled by "column when is3D" / "bar when is3D" above
            case "stock":
                chartElement = BuildStockChart(categories, seriesData, catAxisId, valAxisId);
                needsAxes = true;
                break;
            case "waterfall":
            {
                // Waterfall chart via stacked bar simulation
                double[] wfValues;
                string[]? wfCategories = categories;

                if (seriesData.Count > 1 && seriesData.All(s => s.values.Length == 1))
                {
                    // User passed per-category name:value format (e.g. "Start:1000,Revenue:500,Expense:-200,Net:1300")
                    // Flatten: use series names as categories, combine all single values into one array
                    if (originalCategories == null)
                        wfCategories = seriesData.Select(s => s.name).ToArray();
                    wfValues = seriesData.Select(s => s.values[0]).ToArray();
                }
                else
                {
                    wfValues = seriesData.Count > 0 ? seriesData[0].values : Array.Empty<double>();
                }

                var incColor = properties.GetValueOrDefault("increaseColor", null);
                var decColor = properties.GetValueOrDefault("decreaseColor", null);
                var totColor = properties.GetValueOrDefault("totalColor", null);
                var wfChartSpace = BuildWaterfallChart(title, wfCategories, wfValues,
                    incColor, decColor, totColor, properties);
                return wfChartSpace;
            }
            case "combo":
            {
                int splitAt = 1;
                if (properties.TryGetValue("combosplit", out var splitStr))
                    splitAt = ParseHelpers.SafeParseInt(splitStr, "combosplit");
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
                chartElement = null;
                break;
            }
            default:
                throw new ArgumentException(
                    $"Unknown chart type: '{kind}'. Supported: column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo, waterfall. " +
                    "Add 'stacked' or 'percentstacked' suffix for variants (e.g. columnstacked).");
        }

        if (chartElement != null)
            plotArea.AppendChild(chartElement);

        if (needsAxes)
        {
            if (kind == "scatter")
            {
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

        // Apply cell references for dotted syntax (series1.values=Sheet1!B2:B13)
        ApplySeriesReferences(plotArea, properties);

        return chartSpace;
    }

    /// <summary>
    /// Replace literal Values/CategoryAxisData with NumberReference/StringReference
    /// when dotted syntax cell references are used.
    /// </summary>
    private static void ApplySeriesReferences(C.PlotArea plotArea, Dictionary<string, string> properties)
    {
        var extSeries = ParseSeriesDataExtended(properties);
        if (extSeries == null || extSeries.Count == 0) return;
        if (!extSeries.Any(s => s.ValuesRef != null || s.CategoriesRef != null))
        {
            // Also check top-level categories ref
            var topCatRef = ParseCategoriesRef(properties);
            if (topCatRef == null) return;
        }

        var allSer = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();

        // Top-level categories reference applies to all series
        var topCategoriesRef = ParseCategoriesRef(properties);

        for (int i = 0; i < Math.Min(extSeries.Count, allSer.Count); i++)
        {
            var info = extSeries[i];
            var ser = allSer[i];

            // Replace Values with NumberReference
            if (!string.IsNullOrEmpty(info.ValuesRef))
            {
                var valEl = ser.GetFirstChild<C.Values>();
                if (valEl != null)
                {
                    valEl.RemoveAllChildren();
                    var numRef = new C.NumberReference(new C.Formula(info.ValuesRef));
                    valEl.AppendChild(numRef);
                }
            }

            // Replace CategoryAxisData with StringReference
            var catRef = info.CategoriesRef ?? topCategoriesRef;
            if (!string.IsNullOrEmpty(catRef))
            {
                var catEl = ser.GetFirstChild<C.CategoryAxisData>();
                if (catEl != null)
                {
                    catEl.RemoveAllChildren();
                    var strRef = new C.StringReference(new C.Formula(catRef));
                    catEl.AppendChild(strRef);
                }
                else
                {
                    // Insert CategoryAxisData before Values
                    var valEl = ser.GetFirstChild<C.Values>();
                    var newCat = new C.CategoryAxisData(new C.StringReference(new C.Formula(catRef)));
                    if (valEl != null)
                        valEl.InsertBeforeSelf(newCat);
                    else
                        ser.AppendChild(newCat);
                }
            }
        }
    }

    /// <summary>
    /// Keys that BuildChartSpace doesn't handle directly but SetChartProperties does.
    /// After saving ChartSpace to a ChartPart, call SetChartProperties with these to apply them.
    /// </summary>
    internal static readonly HashSet<string> DeferredAddKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "datalabels", "labels", "labelpos", "labelposition", "labelfont",
        "axistitle", "vtitle", "cattitle", "htitle",
        "axismin", "min", "axismax", "max",
        "majorunit", "minorunit",
        "axisnumfmt", "axisnumberformat",
        "gridlines", "majorgridlines", "minorgridlines",
        "plotareafill", "plotfill", "chartareafill", "chartfill",
        "linewidth", "linedash", "dash", "marker", "markers", "markersize",
        "style", "styleid",
        "transparency", "opacity", "alpha",
        "gradient", "gradients",
        "trendline",
        "secondaryaxis", "secondary",
        "referenceline", "refline", "targetline",
        "colorrule", "conditionalcolor",
        "combotypes", "combo.types",
        "preset", "style.preset", "theme",
        "view3d", "camera", "perspective",
        "holesize", "firstsliceangle", "sliceangle",
        "axisvisible", "axis.visible", "axis.delete",
        "majortickmark", "majortick", "minortickmark", "minortick",
        "ticklabelpos", "ticklabelposition",
        "smooth", "showmarker", "showmarkers",
        "scatterstyle", "radarstyle", "varycolors",
        "dispblanksas", "blanksas", "roundedcorners",
        "datatable", "legend.overlay", "legendoverlay",
        "plotarea.border", "plotborder", "chartarea.border", "chartborder",
        "gapwidth", "gap", "overlap",
        "axisline", "axis.line", "cataxisline", "valaxisline",
        "explosion", "explode", "invertifneg", "invertifnegative",
        "errbars", "errorbars", "series.shadow", "seriesshadow",
        "series.outline", "seriesoutline",
        "bubblescale", "shownegbubbles", "sizerepresents",
        "gapdepth", "shape", "barshape",
        "droplines", "hilowlines", "updownbars", "serlines", "serieslines",
        "axisorientation", "axisreverse", "logbase", "logscale",
        "dispunits", "displayunits", "labeloffset", "ticklabelskip", "tickskip",
        "axisposition", "axispos", "crosses", "crossesat", "crossbetween",
        "plotvisonly", "plotvisibleonly", "autotitledeleted",
        "datalabels.separator", "labelseparator",
        "datalabels.numfmt", "labelnumfmt",
        "datalabels.showleaderlines", "leaderlines",
        "datalabels.showbubblesize",
        "axisfont", "axis.font", "legendfont", "legend.font"
    };

    // ==================== Chart Type Builders ====================

    internal static C.BarChart BuildBarChart(
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

    internal static C.LineChart BuildLineChart(
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

    internal static C.AreaChart BuildAreaChart(
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

    internal static C.PieChart BuildPieChart(
        string[]? categories, List<(string name, double[] values)> seriesData)
    {
        var pieChart = new C.PieChart(new C.VaryColors { Val = true });
        if (seriesData.Count > 0)
            pieChart.AppendChild(BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values));
        return pieChart;
    }

    internal static C.DoughnutChart BuildDoughnutChart(
        string[]? categories, List<(string name, double[] values)> seriesData)
    {
        var chart = new C.DoughnutChart(new C.VaryColors { Val = true });
        if (seriesData.Count > 0)
            chart.AppendChild(BuildPieSeries(0, seriesData[0].name,
                categories, seriesData[0].values));
        chart.AppendChild(new C.HoleSize { Val = 50 });
        return chart;
    }

    internal static C.ScatterChart BuildScatterChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        var scatterChart = new C.ScatterChart(
            new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
            new C.VaryColors { Val = false }
        );

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

    // ==================== Bubble Chart ====================

    internal static C.BubbleChart BuildBubbleChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
    {
        var bubbleChart = new C.BubbleChart(new C.VaryColors { Val = false });

        double[]? xValues = null;
        if (categories != null)
            xValues = categories.Select(c => double.TryParse(c, out var v) ? v : 0).ToArray();

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            var (name, values) = seriesData[i];
            var series = new C.BubbleChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                new C.SeriesText(new C.NumericValue(name))
            );
            ApplySeriesColor(series, color);

            if (xValues != null)
            {
                var xLit = new C.NumberLiteral(new C.PointCount { Val = (uint)xValues.Length });
                for (int j = 0; j < xValues.Length; j++)
                    xLit.AppendChild(new C.NumericPoint(new C.NumericValue(xValues[j].ToString("G"))) { Index = (uint)j });
                series.AppendChild(new C.XValues(xLit));
            }

            var yLit = new C.NumberLiteral(new C.PointCount { Val = (uint)values.Length });
            for (int j = 0; j < values.Length; j++)
                yLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[j].ToString("G"))) { Index = (uint)j });
            series.AppendChild(new C.YValues(yLit));

            // Bubble sizes — use the values as sizes by default, or a third series if provided
            var sizeLit = new C.NumberLiteral(new C.PointCount { Val = (uint)values.Length });
            for (int j = 0; j < values.Length; j++)
                sizeLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[j].ToString("G"))) { Index = (uint)j });
            series.AppendChild(new C.BubbleSize(sizeLit));

            bubbleChart.AppendChild(series);
        }

        bubbleChart.AppendChild(new C.AxisId { Val = catAxisId });
        bubbleChart.AppendChild(new C.AxisId { Val = valAxisId });
        return bubbleChart;
    }

    // ==================== Radar Chart ====================

    internal static C.RadarChart BuildRadarChart(
        string radarStyle,
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId, string[]? colors = null)
    {
        var style = radarStyle.ToLowerInvariant() switch
        {
            "filled" or "fill" => C.RadarStyleValues.Filled,
            "marker" => C.RadarStyleValues.Marker,
            _ => C.RadarStyleValues.Standard
        };

        var radarChart = new C.RadarChart(
            new C.RadarStyle { Val = style },
            new C.VaryColors { Val = false }
        );

        for (int i = 0; i < seriesData.Count; i++)
        {
            var color = colors != null && i < colors.Length ? colors[i] : DefaultSeriesColors[i % DefaultSeriesColors.Length];
            var series = new C.RadarChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                new C.SeriesText(new C.NumericValue(seriesData[i].name))
            );
            ApplySeriesColor(series, color);
            if (categories != null) series.AppendChild(BuildCategoryData(categories));
            series.AppendChild(BuildValues(seriesData[i].values));
            radarChart.AppendChild(series);
        }

        radarChart.AppendChild(new C.AxisId { Val = catAxisId });
        radarChart.AppendChild(new C.AxisId { Val = valAxisId });
        return radarChart;
    }

    // ==================== Stock Chart ====================

    internal static C.StockChart BuildStockChart(
        string[]? categories, List<(string name, double[] values)> seriesData,
        uint catAxisId, uint valAxisId)
    {
        // Stock chart expects series in High-Low-Close order (minimum 3 series)
        // or Open-High-Low-Close order (4 series)
        var stockChart = new C.StockChart();

        for (int i = 0; i < seriesData.Count; i++)
        {
            var series = new C.LineChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                new C.SeriesText(new C.NumericValue(seriesData[i].name))
            );
            if (categories != null) series.AppendChild(BuildCategoryData(categories));
            series.AppendChild(BuildValues(seriesData[i].values));
            stockChart.AppendChild(series);
        }

        stockChart.AppendChild(new C.AxisId { Val = catAxisId });
        stockChart.AppendChild(new C.AxisId { Val = valAxisId });
        return stockChart;
    }

    // ==================== Default Series Colors ====================

    internal static readonly string[] DefaultSeriesColors =
    {
        "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
        "264478", "9B4A22", "636363", "BF8F00", "3A75A8", "4E8538"
    };

    // ==================== Series Color ====================

    internal static void ApplySeriesColor(OpenXmlCompositeElement series, string color)
    {
        series.RemoveAllChildren<C.ChartShapeProperties>();
        var spPr = new C.ChartShapeProperties();
        var solidFill = new Drawing.SolidFill();
        solidFill.AppendChild(BuildChartColorElement(color));
        spPr.AppendChild(solidFill);
        var serText = series.GetFirstChild<C.SeriesText>();
        if (serText != null)
            serText.InsertAfterSelf(spPr);
        else
            series.PrependChild(spPr);
    }

    /// <summary>
    /// Build a fill element: solid if single color, gradient if contains '-'.
    /// Gradient format: "color1-color2[:angle]" or "color1-color2-color3[:angle]"
    /// </summary>
    private static OpenXmlElement BuildFillElement(string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
            return new Drawing.NoFill();

        // Check if it's a gradient (contains - but not a single hex with alpha like 80FF0000)
        var colonIdx = value.LastIndexOf(':');
        var colorPart = colonIdx > 6 ? value[..colonIdx] : value;
        if (colorPart.Contains('-') && colorPart.Split('-').Length >= 2 && colorPart.Split('-')[0].Length <= 8)
        {
            // Gradient: reuse ApplySeriesGradient logic
            var anglePart = 0;
            if (colonIdx > 0 && int.TryParse(value[(colonIdx + 1)..], out var angle))
                anglePart = angle;
            else
                colonIdx = -1;

            var colors = (colonIdx > 0 ? value[..colonIdx] : value).Split('-').Select(c => c.Trim()).ToArray();
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
            gradFill.AppendChild(new Drawing.LinearGradientFill { Angle = anglePart * 60000, Scaled = true });
            return gradFill;
        }

        // Solid fill
        var solidFill = new Drawing.SolidFill();
        solidFill.AppendChild(BuildChartColorElement(value));
        return solidFill;
    }

    /// <summary>
    /// Apply text properties (font, size, color) to all axis labels.
    /// Format: "size:color:fontname" e.g. "10:8B949E:Helvetica Neue" or "10:CCCCCC"
    /// </summary>
    internal static void ApplyAxisTextProperties(OpenXmlCompositeElement axis, string value)
    {
        axis.RemoveAllChildren<C.TextProperties>();
        var parts = value.Split(':');
        var fontSize = parts.Length > 0 && int.TryParse(parts[0], out var fs) ? fs * 100 : 1000;
        var color = parts.Length > 1 ? parts[1] : null;
        var fontName = parts.Length > 2 ? parts[2] : null;

        var defRp = new Drawing.DefaultRunProperties { FontSize = fontSize };
        if (!string.IsNullOrEmpty(color))
        {
            var solidFill = new Drawing.SolidFill();
            solidFill.AppendChild(BuildChartColorElement(color));
            defRp.AppendChild(solidFill);
        }
        if (!string.IsNullOrEmpty(fontName))
        {
            defRp.AppendChild(new Drawing.LatinFont { Typeface = fontName });
            defRp.AppendChild(new Drawing.EastAsianFont { Typeface = fontName });
        }

        var tp = new C.TextProperties(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.ParagraphProperties(defRp))
        );

        // Insert before C.CrossingAxis or at end
        var crossAxis = axis.GetFirstChild<C.CrossingAxis>();
        if (crossAxis != null)
            axis.InsertBefore(tp, crossAxis);
        else
            axis.AppendChild(tp);
    }

    /// <summary>
    /// Build a color element supporting both hex RGB and scheme color names.
    /// </summary>
    private static OpenXmlElement BuildChartColorElement(string value)
    {
        var schemeColor = value.ToLowerInvariant().TrimStart('#') switch
        {
            "accent1" => Drawing.SchemeColorValues.Accent1,
            "accent2" => Drawing.SchemeColorValues.Accent2,
            "accent3" => Drawing.SchemeColorValues.Accent3,
            "accent4" => Drawing.SchemeColorValues.Accent4,
            "accent5" => Drawing.SchemeColorValues.Accent5,
            "accent6" => Drawing.SchemeColorValues.Accent6,
            "dk1" or "dark1" => Drawing.SchemeColorValues.Dark1,
            "dk2" or "dark2" => Drawing.SchemeColorValues.Dark2,
            "lt1" or "light1" => Drawing.SchemeColorValues.Light1,
            "lt2" or "light2" => Drawing.SchemeColorValues.Light2,
            _ => (Drawing.SchemeColorValues?)null
        };
        if (schemeColor.HasValue)
            return new Drawing.SchemeColor { Val = schemeColor.Value };
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(value);
        var el = new Drawing.RgbColorModelHex { Val = rgb };
        if (alpha.HasValue) el.AppendChild(new Drawing.Alpha { Val = alpha.Value });
        return el;
    }

    // ==================== Series Builders ====================

    internal static C.BarChartSeries BuildBarSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.LineChartSeries BuildLineSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.LineChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.AreaChartSeries BuildAreaSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.AreaChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.PieChartSeries BuildPieSeries(uint idx, string name,
        string[]? categories, double[] values, string? color = null)
    {
        var series = new C.PieChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );
        if (color != null) ApplySeriesColor(series, color);
        if (categories != null) series.AppendChild(BuildCategoryData(categories));
        series.AppendChild(BuildValues(values));
        return series;
    }

    internal static C.ScatterChartSeries BuildScatterSeries(uint idx, string name,
        double[]? xValues, double[] yValues)
    {
        var series = new C.ScatterChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.NumericValue(name))
        );

        if (xValues != null)
        {
            var xLit = new C.NumberLiteral(new C.PointCount { Val = (uint)xValues.Length });
            for (int i = 0; i < xValues.Length; i++)
                xLit.AppendChild(new C.NumericPoint(new C.NumericValue(xValues[i].ToString("G"))) { Index = (uint)i });
            series.AppendChild(new C.XValues(xLit));
        }

        var yLit = new C.NumberLiteral(new C.PointCount { Val = (uint)yValues.Length });
        for (int i = 0; i < yValues.Length; i++)
            yLit.AppendChild(new C.NumericPoint(new C.NumericValue(yValues[i].ToString("G"))) { Index = (uint)i });
        series.AppendChild(new C.YValues(yLit));

        return series;
    }

    // ==================== Data Builders ====================

    internal static C.CategoryAxisData BuildCategoryData(string[] categories)
    {
        var strLit = new C.StringLiteral(new C.PointCount { Val = (uint)categories.Length });
        for (int i = 0; i < categories.Length; i++)
            strLit.AppendChild(new C.StringPoint(new C.NumericValue(categories[i])) { Index = (uint)i });
        return new C.CategoryAxisData(strLit);
    }

    internal static C.Values BuildValues(double[] values)
    {
        var numLit = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Length }
        );
        for (int i = 0; i < values.Length; i++)
            numLit.AppendChild(new C.NumericPoint(new C.NumericValue(values[i].ToString("G"))) { Index = (uint)i });
        return new C.Values(numLit);
    }

    /// <summary>
    /// Build a Values element with a NumberReference (cell range formula, no cache).
    /// </summary>
    internal static C.Values BuildValuesRef(string formula)
    {
        var numRef = new C.NumberReference(new C.Formula(formula));
        return new C.Values(numRef);
    }

    /// <summary>
    /// Build a CategoryAxisData element with a StringReference (cell range formula, no cache).
    /// </summary>
    internal static C.CategoryAxisData BuildCategoryDataRef(string formula)
    {
        var strRef = new C.StringReference(new C.Formula(formula));
        return new C.CategoryAxisData(strRef);
    }

    // ==================== Axis Builders ====================

    internal static C.CategoryAxis BuildCategoryAxis(uint axisId, uint crossAxisId)
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

    internal static C.ValueAxis BuildValueAxis(uint axisId, uint crossAxisId, C.AxisPositionValues position)
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

    internal static C.Title BuildChartTitle(string titleText)
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
                            new Drawing.RunProperties { Language = "en-US", FontSize = 1400, Bold = true },
                            new Drawing.Text(titleText)
                        )
                    )
                )
            ),
            new C.Overlay { Val = false }
        );
    }

}
