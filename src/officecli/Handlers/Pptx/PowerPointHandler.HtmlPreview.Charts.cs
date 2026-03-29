// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Chart Rendering ====================

    // Default chart colors matching PowerPoint Office theme accent colors
    private static readonly string[] ChartColors = [
        "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
        "#264478", "#9E480E", "#636363", "#997300", "#255E91", "#43682B"
    ];

    // Chart styling — set per-chart in RenderChart, used by all sub-render methods
    private string _chartValueColor = "#D0D8E0";   // data value labels
    private string _chartCatColor = "#C8D0D8";      // category axis labels
    private string _chartAxisColor = "#B0B8C0";     // value axis labels
    private string _chartGridColor = "#333";        // gridlines
    private string _chartAxisLineColor = "#555";    // axis lines
    private int _chartValFontPx = 9;               // value axis label font size (from OOXML or default)
    private int _chartCatFontPx = 9;               // category axis label font size (from OOXML or default)

    private void RenderChart(StringBuilder sb, GraphicFrame gf, SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        // p:xfrm contains a:off and a:ext
        var pxfrm = gf.GetFirstChild<DocumentFormat.OpenXml.Presentation.Transform>();
        var off = pxfrm?.GetFirstChild<Drawing.Offset>();
        var ext = pxfrm?.GetFirstChild<Drawing.Extents>();
        if (off == null || ext == null) return;

        var x = Units.EmuToPt(off.X?.Value ?? 0);
        var y = Units.EmuToPt(off.Y?.Value ?? 0);
        var w = Units.EmuToPt(ext.Cx?.Value ?? 0);
        var h = Units.EmuToPt(ext.Cy?.Value ?? 0);

        // Read chart data — find c:chart element with r:id
        var chartEl = gf.Descendants().FirstOrDefault(e => e.LocalName == "chart" && e.NamespaceUri.Contains("chart"));
        var rId = chartEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri.Contains("relationships")).Value;
        if (rId == null) return;

        DocumentFormat.OpenXml.Drawing.Charts.Chart? chart;
        DocumentFormat.OpenXml.Drawing.Charts.PlotArea? plotArea;
        try
        {
            var chartPart = (ChartPart)slidePart.GetPartById(rId);
            chart = chartPart.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            plotArea = chart?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>();
            if (plotArea == null) return;
        }
        catch { return; }

        var chartType = ChartHelper.DetectChartType(plotArea) ?? "bar";
        var categories = ChartHelper.ReadCategories(plotArea) ?? [];
        var seriesList = ChartHelper.ReadAllSeries(plotArea);
        if (seriesList.Count == 0) return;

        // Read series colors
        var seriesColors = new List<string>();
        var serElements = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();
        for (int i = 0; i < seriesList.Count; i++)
        {
            var serEl = i < serElements.Count ? serElements[i] : null;
            var spPr = serEl?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>();
            var fill = spPr?.GetFirstChild<Drawing.SolidFill>();
            var rgb = fill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            seriesColors.Add(rgb != null ? $"#{rgb}" : ChartColors[i % ChartColors.Length]);
        }

        // Derive text color from theme: use tx1 or dk1 (with #), fallback to light gray
        var chartTextColor = themeColors.TryGetValue("tx1", out var tx1) ? $"#{tx1}"
            : themeColors.TryGetValue("dk1", out var dk1) ? $"#{dk1}" : "#D0D8E0";
        var chartLabelColor = chartTextColor;
        var chartAxisColor = chartTextColor;

        // Set instance fields for sub-render methods to use theme-derived colors
        _chartValueColor = chartTextColor;
        _chartCatColor = chartTextColor;
        _chartAxisColor = chartTextColor;
        // Derive gridline/axis line colors: dim version of text color
        var isDarkText = IsColorDark(chartTextColor.TrimStart('#'));
        _chartGridColor = isDarkText ? "#ccc" : "#333";
        _chartAxisLineColor = isDarkText ? "#aaa" : "#555";

        // Title
        var chartTitle = chart?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
        var titleText = chartTitle?.Descendants<Drawing.Text>().FirstOrDefault()?.Text ?? "";
        var titleFontSize = chartTitle?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
        var titleSizeCss = titleFontSize?.HasValue == true ? $"{titleFontSize.Value / 100.0:0.##}pt" : "8pt";

        // Check if dataLabels are enabled
        var dataLabels = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
        var showValues = dataLabels?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowValue>()?.Val?.Value == true
            || dataLabels?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowCategoryName>()?.Val?.Value == true
            || dataLabels?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowPercent>()?.Val?.Value == true;

        // Plot/chart fill — only direct children, not series fills
        var plotSpPr = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShapeProperties>();
        var plotFillColor = plotSpPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        var chartSpPr = chart?.Parent?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>();
        var chartFillColor = chartSpPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;

        // Axis titles
        var valAxis = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>();
        var valAxisTitle = valAxis?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        var catAxis = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>();
        var catAxisTitle = catAxis?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;

        // Read explicit axis parameters from OOXML (override auto-calculation when present)
        var valScaling = valAxis?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Scaling>();
        double? ooxmlAxisMax = null, ooxmlAxisMin = null, ooxmlMajorUnit = null;
        var maxEl = valScaling?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.MaxAxisValue>();
        if (maxEl?.Val?.HasValue == true) ooxmlAxisMax = maxEl.Val.Value;
        var minEl = valScaling?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.MinAxisValue>();
        if (minEl?.Val?.HasValue == true) ooxmlAxisMin = minEl.Val.Value;
        var majorUnitEl = valAxis?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.MajorUnit>();
        if (majorUnitEl?.Val?.HasValue == true) ooxmlMajorUnit = majorUnitEl.Val.Value;

        // Read gapWidth from bar/column chart
        var gapWidthEl = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.GapWidth>().FirstOrDefault();
        int? ooxmlGapWidth = gapWidthEl?.Val?.HasValue == true ? (int)gapWidthEl.Val.Value : null;

        // Read axis label font sizes from OOXML
        var valAxisFontSize = valAxis?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
        var catAxisFontSize = catAxis?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
        int valLabelPx = valAxisFontSize?.HasValue == true ? (int)(valAxisFontSize.Value / 100.0 * 96 / 72) : 9;
        int catLabelPx = catAxisFontSize?.HasValue == true ? (int)(catAxisFontSize.Value / 100.0 * 96 / 72) : 9;
        _chartValFontPx = valLabelPx;
        _chartCatFontPx = catLabelPx;

        // Shared SVG renderer for 2D charts (shared with Excel)
        var svgRenderer = new OfficeCli.Core.ChartSvgRenderer
        {
            ValueColor = _chartValueColor, CatColor = _chartCatColor,
            AxisColor = _chartAxisColor, GridColor = _chartGridColor,
            AxisLineColor = _chartAxisLineColor, ValFontPx = _chartValFontPx, CatFontPx = _chartCatFontPx
        };

        // Container with optional chart background
        var bgStyle = chartFillColor != null ? $"background:#{chartFillColor};" : "background:transparent;";
        sb.AppendLine($"    <div class=\"shape\" style=\"left:{x}pt;top:{y}pt;width:{w}pt;height:{h}pt;{bgStyle}display:flex;flex-direction:column;overflow:hidden\">");

        // Title
        if (!string.IsNullOrEmpty(titleText))
            sb.AppendLine($"      <div style=\"text-align:center;font-size:{titleSizeCss};font-weight:bold;padding:4px;flex-shrink:0;color:{chartTextColor}\">{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(titleText)}</div>");

        // SVG chart area — proportional to actual shape dimensions
        var widthEmu = ext.Cx?.Value ?? 3600000;
        var heightEmu = ext.Cy?.Value ?? 2520000;
        var svgW = (int)(widthEmu / 10000.0); // scale down to reasonable SVG units
        var svgH = (int)(heightEmu / 10000.0);
        var titleH = string.IsNullOrEmpty(titleText) ? 0 : 20;
        var chartSvgH = svgH - titleH;
        // Try to read manual layout from OOXML plotArea
        var plotAreaLayout = plotArea?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Layout>();
        var manualLayout = plotAreaLayout?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ManualLayout>();
        int marginTop, marginRight, marginBottom, marginLeft;
        if (manualLayout != null)
        {
            // ManualLayout x/y/w/h are fractions of the chart area (0.0 - 1.0)
            var mlX = manualLayout.Left?.Val?.Value ?? 0.0;
            var mlY = manualLayout.Top?.Val?.Value ?? 0.0;
            var mlW = manualLayout.Width?.Val?.Value ?? 1.0;
            var mlH = manualLayout.Height?.Val?.Value ?? 1.0;
            marginLeft = (int)(mlX * svgW);
            marginTop = (int)(mlY * chartSvgH);
            marginRight = (int)((1.0 - mlX - mlW) * svgW);
            marginBottom = (int)((1.0 - mlY - mlH) * chartSvgH);
            if (marginLeft < 5) marginLeft = 5;
            if (marginRight < 5) marginRight = 5;
            if (marginTop < 5) marginTop = 5;
            if (marginBottom < 5) marginBottom = 5;
        }
        else
        {
            marginTop = 10; marginRight = 15; marginBottom = 25; marginLeft = 40;
        }
        var margin = new { top = marginTop, right = marginRight, bottom = marginBottom, left = marginLeft };
        var plotW = svgW - margin.left - margin.right;
        var plotH = chartSvgH - margin.top - margin.bottom;

        var is3D = chartType.Contains("3d");

        // Show legend by default for multi-series or pie/doughnut charts.
        // Only hide if the OOXML chart explicitly has <c:legend> with <c:delete val="1"/>.
        var legendEl = chart?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Legend>();
        var isPieOrDoughnut = chartType.Contains("pie") || chartType.Contains("doughnut");
        bool hasLegend;
        if (legendEl != null)
        {
            var deleteEl = legendEl.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Delete>();
            hasLegend = deleteEl?.Val?.Value != true;
        }
        else
        {
            hasLegend = seriesList.Count > 1 || isPieOrDoughnut;
        }
        sb.AppendLine($"      <svg viewBox=\"0 0 {svgW} {chartSvgH}\" style=\"width:100%;flex:1;min-height:0\" preserveAspectRatio=\"xMidYMin meet\">");

        // Plot area background
        if (plotFillColor != null)
            sb.AppendLine($"        <rect x=\"{margin.left}\" y=\"{margin.top}\" width=\"{plotW}\" height=\"{plotH}\" fill=\"#{plotFillColor}\"/>");

        if (is3D && (chartType.Contains("pie") || chartType.Contains("doughnut")))
        {
            RenderPie3DSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH);
        }
        else if (is3D && (chartType.Contains("column") || chartType.Contains("bar")))
        {
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            var is3DStacked = chartType.Contains("stacked") || chartType.Contains("Stacked");
            if (is3DStacked)
            {
                // 3D stacked bars: fall through to 2D stacked renderer for correct stacking
                var isPercent = chartType.Contains("percent") || chartType.Contains("Percent");
                svgRenderer.RenderBarChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH, isHorizontal, true, isPercent);
            }
            else
            {
                RenderBar3DSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH, isHorizontal);
            }
        }
        else if (is3D && chartType.Contains("line"))
        {
            // 3D line: render with depth shadows
            RenderLine3DSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType.Contains("pie") || chartType.Contains("doughnut"))
        {
            var isDoughnut = chartType.Contains("doughnut");
            var holeSize = 0.0;
            if (isDoughnut)
            {
                var holeSizeEl = plotArea!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.HoleSize>().FirstOrDefault();
                holeSize = (holeSizeEl?.Val?.Value ?? 50) / 100.0;
            }
            svgRenderer.RenderPieChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH, holeSize, showValues);
        }
        else if (chartType.Contains("area"))
        {
            var areaStacked = chartType.Contains("stacked") || chartType.Contains("Stacked");
            var areaW = plotW - (int)(plotW * 0.03);
            svgRenderer.RenderAreaChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, areaW, plotH, areaStacked);
        }
        else if (chartType == "combo")
        {
            svgRenderer.RenderComboChartSvg(sb, plotArea!, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType.Contains("radar"))
        {
            svgRenderer.RenderRadarChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH, catLabelPx);
        }
        else if (chartType == "bubble")
        {
            svgRenderer.RenderBubbleChartSvg(sb, plotArea!, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType == "stock")
        {
            svgRenderer.RenderStockChartSvg(sb, plotArea!, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType.Contains("line") || chartType == "scatter")
        {
            var lineW = plotW - (int)(plotW * 0.03);
            svgRenderer.RenderLineChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, lineW, plotH, showValues);
        }
        else
        {
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            var isStacked = chartType.Contains("stacked") || chartType.Contains("Stacked");
            var isPercent = chartType.Contains("percent") || chartType.Contains("Percent");
            svgRenderer.RenderBarChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH, isHorizontal, isStacked, isPercent,
                ooxmlAxisMax, ooxmlAxisMin, ooxmlMajorUnit, ooxmlGapWidth, valLabelPx, catLabelPx);
        }

        // Axis titles inside SVG
        if (!string.IsNullOrEmpty(valAxisTitle))
            sb.AppendLine($"        <text x=\"10\" y=\"{chartSvgH / 2}\" fill=\"{chartAxisColor}\" font-size=\"{_chartValFontPx}\" text-anchor=\"middle\" dominant-baseline=\"middle\" transform=\"rotate(-90,10,{chartSvgH / 2})\">{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(valAxisTitle)}</text>");
        if (!string.IsNullOrEmpty(catAxisTitle))
            sb.AppendLine($"        <text x=\"{svgW / 2}\" y=\"{chartSvgH - 2}\" fill=\"{chartAxisColor}\" font-size=\"{_chartValFontPx}\" text-anchor=\"middle\">{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(catAxisTitle)}</text>");

        sb.AppendLine("      </svg>");

        // Legend — render when the OOXML chart contains a <c:legend> element
        var legendFontSize = legendEl?.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
        var legendSizeCss = legendFontSize?.HasValue == true ? $"{legendFontSize.Value / 100.0:0.##}pt" : "8pt";
        if (hasLegend)
        {
            sb.Append($"      <div class=\"chart-legend\" style=\"display:flex;flex-shrink:0;justify-content:center;gap:16px;padding:4px 0;font-size:{legendSizeCss};color:{chartLabelColor}\">");
            if (isPieOrDoughnut && categories.Length > 0)
            {
                for (int i = 0; i < categories.Length; i++)
                {
                    var color = i < seriesColors.Count ? seriesColors[i] : ChartColors[i % ChartColors.Length];
                    sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><span style=\"display:inline-block;width:12px;height:12px;background:{color};border-radius:1px\"></span>{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(categories[i])}</span>");
                }
            }
            else
            {
                for (int i = 0; i < seriesList.Count; i++)
                {
                    sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><span style=\"display:inline-block;width:12px;height:12px;background:{seriesColors[i]};border-radius:1px\"></span>{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(seriesList[i].name)}</span>");
                }
            }
            sb.AppendLine("</div>");
        }

        sb.AppendLine("    </div>");
    }


    // ==================== 3D Chart Helpers ====================

    /// <summary>Darken or lighten a hex color by a factor (0.0-2.0, 1.0=unchanged)</summary>
    private static string AdjustColor(string hexColor, double factor)
    {
        var hex = hexColor.TrimStart('#');
        if (hex.Length < 6) return hexColor;
        var r = (int)Math.Clamp(int.Parse(hex[..2], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        var g = (int)Math.Clamp(int.Parse(hex[2..4], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        var b = (int)Math.Clamp(int.Parse(hex[4..6], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    // 3D isometric offsets (simulating ~30° viewing angle)
    private const double Depth3D = 12; // pixel depth for 3D extrusion
    private const double DxIso = 8;    // horizontal offset for depth
    private const double DyIso = -6;   // vertical offset for depth (negative = upward)

    private void RenderBar3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool horizontal)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var (maxVal, _, _) = OfficeCli.Core.ChartSvgRenderer.ComputeNiceAxis(allValues.Max());
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;

        if (horizontal)
        {
            var hLabelMargin = 50;
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;
            var groupH = (double)ph / Math.Max(catCount, 1);
            var barH = groupH * 0.5 / serCount;
            var gap = groupH * 0.2;

            // Gridlines
            for (int t = 1; t <= 4; t++)
            {
                var gx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{_chartGridColor}\" stroke-width=\"0.5\" stroke-dasharray=\"none\"/>");
            }
            // Axis lines
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"{_chartAxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"{_chartAxisLineColor}\" stroke-width=\"1\"/>");

            for (int s = 0; s < serCount; s++)
            {
                var color = colors[s % colors.Count];
                var sideColor = AdjustColor(color, 0.7);
                var topColor = AdjustColor(color, 1.3);
                for (int c = 0; c < series[s].values.Length && c < catCount; c++)
                {
                    var val = series[s].values[c];
                    var barW = (val / maxVal) * plotPw;
                    var bx = plotOx;
                    var by = oy + c * groupH + gap + s * barH;
                    sb.AppendLine($"        <polygon points=\"{bx:0.#},{by:0.#} {bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + DxIso:0.#},{by + DyIso:0.#}\" fill=\"{topColor}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{color}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <polygon points=\"{bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + barW + DxIso:0.#},{by + barH + DyIso:0.#} {bx + barW:0.#},{by + barH:0.#}\" fill=\"{sideColor}\" opacity=\"0.9\"/>");
                    // Value label
                    var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                    sb.AppendLine($"        <text x=\"{bx + barW + DxIso + 4:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"{_chartValueColor}\" font-size=\"7\" text-anchor=\"start\" dominant-baseline=\"middle\">{vlabel}</text>");
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var ly = oy + c * groupH + groupH / 2;
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"{_chartCatColor}\" font-size=\"{_chartCatFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var tx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{_chartAxisColor}\" font-size=\"{_chartValFontPx}\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.5 / serCount;
            var gap = groupW * 0.2;

            // Gridlines
            for (int t = 1; t <= 4; t++)
            {
                var gy = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{_chartGridColor}\" stroke-width=\"0.5\" stroke-dasharray=\"none\"/>");
            }
            // Axis lines
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{_chartAxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{_chartAxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                for (int s = 0; s < serCount; s++)
                {
                    if (c >= series[s].values.Length) continue;
                    var val = series[s].values[c];
                    var color = colors[s % colors.Count];
                    var sideColor = AdjustColor(color, 0.65);
                    var topColor = AdjustColor(color, 1.25);
                    var barH = (val / maxVal) * ph;
                    var bx = ox + c * groupW + gap + s * barW;
                    var by = oy + ph - barH;

                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{color}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <polygon points=\"{bx:0.#},{by:0.#} {bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + DxIso:0.#},{by + DyIso:0.#}\" fill=\"{topColor}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <polygon points=\"{bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + barW + DxIso:0.#},{oy + ph + DyIso:0.#} {bx + barW:0.#},{oy + ph:0.#}\" fill=\"{sideColor}\" opacity=\"0.9\"/>");
                    // Value label above top face
                    var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                    sb.AppendLine($"        <text x=\"{bx + barW / 2 + DxIso / 2:0.#}\" y=\"{by + DyIso - 3:0.#}\" fill=\"{_chartValueColor}\" font-size=\"7\" text-anchor=\"middle\">{vlabel}</text>");
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{_chartCatColor}\" font-size=\"{_chartCatFontPx}\" text-anchor=\"middle\">{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var ty = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{_chartAxisColor}\" font-size=\"{_chartValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    private void RenderPie3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH)
    {
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var rx = Math.Min(svgW, svgH) * 0.35;   // horizontal radius
        var ry = rx * 0.55;                       // vertical radius (elliptical for 3D tilt)
        var depth = rx * 0.15;                    // extrusion depth
        var startAngle = -Math.PI / 2;

        // Render extrusion sides first (back to front)
        // Sort slices by midpoint angle for correct z-ordering of sides
        var slices = new List<(int idx, double start, double end, string color)>();
        var angle = startAngle;
        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var color = i < colors.Count ? colors[i] : ChartColors[i % ChartColors.Length];
            slices.Add((i, angle, angle + sliceAngle, color));
            angle += sliceAngle;
        }

        // Draw side extrusions for slices that face the viewer (bottom half)
        foreach (var (idx, start, end, color) in slices)
        {
            var sideColor = AdjustColor(color, 0.6);
            // Only draw sides for the visible portion (angles where sin > 0, i.e. bottom)
            var visStart = Math.Max(start, 0);
            var visEnd = Math.Min(end, Math.PI);
            if (start < Math.PI && end > 0)
            {
                var clampedStart = Math.Max(start, -0.01); // slightly past top to avoid gaps
                var clampedEnd = Math.Min(end, Math.PI + 0.01);
                // Build side path: outer arc at bottom, lines down, inner arc at top+depth
                var steps = Math.Max(8, (int)((clampedEnd - clampedStart) / 0.1));
                var pathPoints = new StringBuilder();
                pathPoints.Append($"M {cx + rx * Math.Cos(clampedStart):0.#},{cy + ry * Math.Sin(clampedStart):0.#} ");
                for (int step = 0; step <= steps; step++)
                {
                    var a = clampedStart + (clampedEnd - clampedStart) * step / steps;
                    pathPoints.Append($"L {cx + rx * Math.Cos(a):0.#},{cy + ry * Math.Sin(a):0.#} ");
                }
                for (int step = steps; step >= 0; step--)
                {
                    var a = clampedStart + (clampedEnd - clampedStart) * step / steps;
                    pathPoints.Append($"L {cx + rx * Math.Cos(a):0.#},{cy + ry * Math.Sin(a) + depth:0.#} ");
                }
                pathPoints.Append("Z");
                sb.AppendLine($"        <path d=\"{pathPoints}\" fill=\"{sideColor}\" opacity=\"0.9\"/>");
            }
        }

        // Draw top elliptical slices
        startAngle = -Math.PI / 2;
        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : ChartColors[i % ChartColors.Length];

            if (values.Length == 1)
            {
                sb.AppendLine($"        <ellipse cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" rx=\"{rx:0.#}\" ry=\"{ry:0.#}\" fill=\"{color}\" opacity=\"0.9\"/>");
            }
            else
            {
                var x1 = cx + rx * Math.Cos(startAngle);
                var y1 = cy + ry * Math.Sin(startAngle);
                var x2 = cx + rx * Math.Cos(endAngle);
                var y2 = cy + ry * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {cx:0.#},{cy:0.#} L {x1:0.#},{y1:0.#} A {rx:0.#},{ry:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"0.9\"/>");
            }

            // Label
            var midAngle = startAngle + sliceAngle / 2;
            var lx = cx + rx * 0.55 * Math.Cos(midAngle);
            var ly = cy + ry * 0.55 * Math.Sin(midAngle);
            var label = i < categories.Length ? categories[i] : "";
            if (!string.IsNullOrEmpty(label))
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"white\" font-size=\"9\" text-anchor=\"middle\" dominant-baseline=\"middle\">{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(label)}</text>");

            startAngle = endAngle;
        }
    }

    private void RenderLine3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var (maxVal, _, _) = OfficeCli.Core.ChartSvgRenderer.ComputeNiceAxis(allValues.Max());
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{_chartAxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{_chartAxisLineColor}\" stroke-width=\"1\"/>");

        // Render series back to front
        for (int s = series.Count - 1; s >= 0; s--)
        {
            var color = colors[s % colors.Count];
            var shadowColor = AdjustColor(color, 0.5);
            var points = new List<(double x, double y)>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / maxVal) * ph;
                points.Add((px, py));
            }
            if (points.Count > 1)
            {
                // Draw "ribbon" — a filled area between the line and its offset
                var ribbon = new StringBuilder();
                ribbon.Append("M ");
                for (int p = 0; p < points.Count; p++)
                    ribbon.Append($"{points[p].x:0.#},{points[p].y:0.#} L ");
                for (int p = points.Count - 1; p >= 0; p--)
                    ribbon.Append($"{points[p].x + DxIso:0.#},{points[p].y + DyIso:0.#} L ");
                ribbon.Length -= 2; // remove trailing " L"
                ribbon.Append(" Z");
                sb.AppendLine($"        <path d=\"{ribbon}\" fill=\"{shadowColor}\" opacity=\"0.4\"/>");

                // Main line
                var linePoints = string.Join(" ", points.Select(p => $"{p.x:0.#},{p.y:0.#}"));
                sb.AppendLine($"        <polyline points=\"{linePoints}\" fill=\"none\" stroke=\"{color}\" stroke-width=\"2.5\"/>");
                foreach (var pt in points)
                    sb.AppendLine($"        <circle cx=\"{pt.x:0.#}\" cy=\"{pt.y:0.#}\" r=\"3\" fill=\"{color}\"/>");
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{_chartCatColor}\" font-size=\"{_chartCatFontPx}\" text-anchor=\"middle\">{OfficeCli.Core.ChartSvgRenderer.HtmlEncode(label)}</text>");
        }
    }

    /// <summary>
    /// Compute a "nice" axis scale with ~10-15% headroom above the data max.
    /// Returns (niceMax, tickStep, nTicks).
    /// </summary>
}
