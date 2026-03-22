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

    private static readonly string[] ChartColors = [
        "#E74C3C", "#3498DB", "#2ECC71", "#F39C12", "#9B59B6", "#1ABC9C",
        "#E67E22", "#34495E", "#E91E63", "#00BCD4", "#8BC34A", "#FF9800"
    ];

    private void RenderChart(StringBuilder sb, GraphicFrame gf, SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        // p:xfrm contains a:off and a:ext
        var pxfrm = gf.GetFirstChild<DocumentFormat.OpenXml.Presentation.Transform>();
        var off = pxfrm?.GetFirstChild<Drawing.Offset>();
        var ext = pxfrm?.GetFirstChild<Drawing.Extents>();
        if (off == null || ext == null) return;

        var x = EmuToCm(off.X?.Value ?? 0);
        var y = EmuToCm(off.Y?.Value ?? 0);
        var w = EmuToCm(ext.Cx?.Value ?? 0);
        var h = EmuToCm(ext.Cy?.Value ?? 0);

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

        // Title
        var titleText = chart?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>()
            ?.Descendants<Drawing.Text>().FirstOrDefault()?.Text ?? "";

        // Check if dataLabels are enabled
        var dataLabels = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
        var showValues = dataLabels?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowValue>()?.Val?.Value == true
            || dataLabels?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowCategoryName>()?.Val?.Value == true
            || dataLabels?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowPercent>()?.Val?.Value == true;

        // Plot/chart fill
        var plotSpPr = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShapeProperties>();
        var plotFillColor = plotSpPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        var chartSpPr = chart?.Parent?.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().FirstOrDefault();
        var chartFillColor = chartSpPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;

        // Axis titles
        var valAxis = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>();
        var valAxisTitle = valAxis?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
        var catAxis = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>();
        var catAxisTitle = catAxis?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>()?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;

        // Container with optional chart background
        var bgStyle = chartFillColor != null ? $"background:#{chartFillColor};" : "background:rgba(255,255,255,0.05);";
        sb.AppendLine($"    <div class=\"shape\" style=\"left:{x}cm;top:{y}cm;width:{w}cm;height:{h}cm;{bgStyle}\">");

        // Title
        if (!string.IsNullOrEmpty(titleText))
            sb.AppendLine($"      <div style=\"text-align:center;font-size:11px;font-weight:bold;padding:4px;color:#ccc\">{HtmlEncode(titleText)}</div>");

        // SVG chart area — proportional to actual shape dimensions
        var widthEmu = ext.Cx?.Value ?? 3600000;
        var heightEmu = ext.Cy?.Value ?? 2520000;
        var svgW = (int)(widthEmu / 10000.0); // scale down to reasonable SVG units
        var svgH = (int)(heightEmu / 10000.0);
        var titleH = string.IsNullOrEmpty(titleText) ? 0 : 20;
        var chartSvgH = svgH - titleH;
        var margin = new { top = 10, right = 15, bottom = 25, left = 40 };
        var plotW = svgW - margin.left - margin.right;
        var plotH = chartSvgH - margin.top - margin.bottom;

        var is3D = chartType.Contains("3d");

        sb.AppendLine($"      <svg viewBox=\"0 0 {svgW} {chartSvgH}\" style=\"width:100%;height:calc(100% - {titleH + 4}px)\" preserveAspectRatio=\"xMidYMin meet\">");

        // Plot area background
        if (plotFillColor != null)
            sb.AppendLine($"        <rect x=\"{margin.left}\" y=\"{margin.top}\" width=\"{plotW}\" height=\"{plotH}\" fill=\"#{plotFillColor}\" opacity=\"0.3\"/>");

        if (is3D && (chartType.Contains("pie") || chartType.Contains("doughnut")))
        {
            RenderPie3DSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH);
        }
        else if (is3D && (chartType.Contains("column") || chartType.Contains("bar")))
        {
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            RenderBar3DSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH, isHorizontal);
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
                var holeSizeEl = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.HoleSize>().FirstOrDefault();
                holeSize = (holeSizeEl?.Val?.Value ?? 50) / 100.0;
            }
            RenderPieChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH, holeSize);
        }
        else if (chartType.Contains("area"))
        {
            RenderAreaChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType == "combo")
        {
            RenderComboChartSvg(sb, plotArea, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType.Contains("radar"))
        {
            RenderRadarChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH);
        }
        else if (chartType == "bubble")
        {
            RenderBubbleChartSvg(sb, plotArea, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType == "stock")
        {
            RenderStockChartSvg(sb, plotArea, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType.Contains("line") || chartType == "scatter")
        {
            RenderLineChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else
        {
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            var isStacked = chartType.Contains("stacked") || chartType.Contains("Stacked");
            var isPercent = chartType.Contains("percent") || chartType.Contains("Percent");
            RenderBarChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH, isHorizontal, isStacked, isPercent);
        }

        // Axis titles inside SVG
        if (!string.IsNullOrEmpty(valAxisTitle))
            sb.AppendLine($"        <text x=\"10\" y=\"{chartSvgH / 2}\" fill=\"#888\" font-size=\"8\" text-anchor=\"middle\" dominant-baseline=\"middle\" transform=\"rotate(-90,10,{chartSvgH / 2})\">{HtmlEncode(valAxisTitle)}</text>");
        if (!string.IsNullOrEmpty(catAxisTitle))
            sb.AppendLine($"        <text x=\"{svgW / 2}\" y=\"{chartSvgH - 2}\" fill=\"#888\" font-size=\"8\" text-anchor=\"middle\">{HtmlEncode(catAxisTitle)}</text>");

        sb.AppendLine("      </svg>");

        // Legend
        if (seriesList.Count > 1)
        {
            sb.Append("      <div style=\"display:flex;justify-content:center;gap:8px;font-size:8px;color:#aaa;padding:2px\">");
            for (int i = 0; i < seriesList.Count; i++)
            {
                sb.Append($"<span><span style=\"display:inline-block;width:8px;height:8px;background:{seriesColors[i]};margin-right:2px;border-radius:1px\"></span>{HtmlEncode(seriesList[i].name)}</span>");
            }
            sb.AppendLine("</div>");
        }

        sb.AppendLine("    </div>");
    }

    private static void RenderBarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph,
        bool horizontal, bool stacked = false, bool percentStacked = false)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;
        if (percentStacked) stacked = true;

        double maxVal;
        if (percentStacked)
        {
            maxVal = 100;
        }
        else if (stacked)
        {
            maxVal = 0;
            for (int c = 0; c < catCount; c++)
            {
                var sum = series.Sum(s => c < s.values.Length ? s.values[c] : 0);
                if (sum > maxVal) maxVal = sum;
            }
        }
        else
        {
            maxVal = allValues.Max();
        }
        if (maxVal <= 0) maxVal = 1;

        if (horizontal)
        {
            var hLabelMargin = 50;
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;
            var groupH = (double)ph / Math.Max(catCount, 1);
            var barH = stacked ? groupH * 0.7 : groupH * 0.7 / serCount;
            var gap = groupH * 0.15;

            // Gridlines
            for (int t = 1; t <= 4; t++)
            {
                var gx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"#333\" stroke-width=\"0.5\" stroke-dasharray=\"3,3\"/>");
            }

            // Axis lines
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

            // Bars + value labels
            for (int c = 0; c < catCount; c++)
            {
                double stackX = 0;
                var catSum = percentStacked ? series.Sum(s => c < s.values.Length ? s.values[c] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = c < series[s].values.Length ? series[s].values[c] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    var barW = (val / maxVal) * plotPw;
                    if (stacked)
                    {
                        var bx = plotOx + (stackX / maxVal) * plotPw;
                        var by = oy + c * groupH + gap;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        if (barW > 20)
                        {
                            var vlabel = percentStacked ? $"{val:0}%" : (rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}");
                            sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"white\" font-size=\"7\" text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                        }
                        stackX += val;
                    }
                    else
                    {
                        var bx = plotOx;
                        var by = oy + c * groupH + gap + s * barH;
                        sb.AppendLine($"        <rect x=\"{bx}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        var vlabel = rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}";
                        sb.AppendLine($"        <text x=\"{bx + barW + 4:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"#aaa\" font-size=\"7\" text-anchor=\"start\" dominant-baseline=\"middle\">{vlabel}</text>");
                    }
                }
            }

            // Category labels
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var ly = oy + c * groupH + groupH / 2;
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"#999\" font-size=\"9\" text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }

            // Value axis labels
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var tx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#777\" font-size=\"8\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = stacked ? groupW * 0.7 : groupW * 0.7 / serCount;
            var gap = groupW * 0.15;

            // Gridlines
            for (int t = 1; t <= 4; t++)
            {
                var gy = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"#333\" stroke-width=\"0.5\" stroke-dasharray=\"3,3\"/>");
            }

            // Axis lines
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

            // Bars + value labels
            for (int c = 0; c < catCount; c++)
            {
                double stackY = 0;
                var catSum = percentStacked ? series.Sum(s => c < s.values.Length ? s.values[c] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = c < series[s].values.Length ? series[s].values[c] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    var barH = (val / maxVal) * ph;
                    if (stacked)
                    {
                        var bx = ox + c * groupW + gap;
                        var by = oy + ph - (stackY / maxVal) * ph - barH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        if (barH > 12)
                        {
                            var vlabel = percentStacked ? $"{val:0}%" : (rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}");
                            sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"white\" font-size=\"7\" text-anchor=\"middle\" dominant-baseline=\"middle\">{vlabel}</text>");
                        }
                        stackY += val;
                    }
                    else
                    {
                        var bx = ox + c * groupW + gap + s * barW;
                        var by = oy + ph - barH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        var vlabel = rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}";
                        sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by - 3:0.#}\" fill=\"#aaa\" font-size=\"7\" text-anchor=\"middle\">{vlabel}</text>");
                    }
                }
            }

            // Category labels
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }

            // Value axis labels
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var ty = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    private static void RenderLineChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Gridlines
        for (int t = 1; t <= 4; t++)
        {
            var gy = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"#333\" stroke-width=\"0.5\" stroke-dasharray=\"3,3\"/>");
        }

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / maxVal) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
                // Dots + value labels
                for (int p = 0; p < points.Count; p++)
                {
                    var parts = points[p].Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s]}\"/>");
                    var val = series[s].values[p];
                    var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                    sb.AppendLine($"        <text x=\"{parts[0]}\" y=\"{double.Parse(parts[1]) - 6:0.#}\" fill=\"#aaa\" font-size=\"7\" text-anchor=\"middle\">{vlabel}</text>");
                }
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    private static void RenderPieChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, double holeRatio = 0.0)
    {
        // Use first series values
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.35;
        var innerR = r * holeRatio;
        var startAngle = -Math.PI / 2;

        // Render all slices first
        var labels = new List<(double x, double y, string text, string anchor, string fill)>();
        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : ChartColors[i % ChartColors.Length];

            if (values.Length == 1 && holeRatio <= 0)
            {
                sb.AppendLine($"        <circle cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" r=\"{r:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            else if (holeRatio > 0)
            {
                var ox1 = cx + r * Math.Cos(startAngle);
                var oy1 = cy + r * Math.Sin(startAngle);
                var ox2 = cx + r * Math.Cos(endAngle);
                var oy2 = cy + r * Math.Sin(endAngle);
                var ix1 = cx + innerR * Math.Cos(endAngle);
                var iy1 = cy + innerR * Math.Sin(endAngle);
                var ix2 = cx + innerR * Math.Cos(startAngle);
                var iy2 = cy + innerR * Math.Sin(startAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {ox1:0.#},{oy1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {ox2:0.#},{oy2:0.#} L {ix1:0.#},{iy1:0.#} A {innerR:0.#},{innerR:0.#} 0 {largeArc},0 {ix2:0.#},{iy2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            else
            {
                var x1 = cx + r * Math.Cos(startAngle);
                var y1 = cy + r * Math.Sin(startAngle);
                var x2 = cx + r * Math.Cos(endAngle);
                var y2 = cy + r * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {cx:0.#},{cy:0.#} L {x1:0.#},{y1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }

            // Collect label — small slices get labels outside, large slices inside
            var midAngle = startAngle + sliceAngle / 2;
            var label = i < categories.Length ? categories[i] : "";
            if (!string.IsNullOrEmpty(label))
            {
                var slicePct = values[i] / total;
                bool outside = slicePct < 0.08; // < 8% of total → label outside
                double labelR;
                if (holeRatio > 0)
                    labelR = outside ? r + 12 : (r + innerR) / 2;
                else
                    labelR = outside ? r + 12 : r * 0.55;
                var lx = cx + labelR * Math.Cos(midAngle);
                var ly = cy + labelR * Math.Sin(midAngle);
                var anchor = outside ? (Math.Cos(midAngle) >= 0 ? "start" : "end") : "middle";
                var fill = outside ? "#ccc" : "white";
                labels.Add((lx, ly, label, anchor, fill));
            }

            startAngle = endAngle;
        }

        // Render labels on top of all slices
        foreach (var (lx, ly, label, anchor, fill) in labels)
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"{fill}\" font-size=\"9\" font-weight=\"bold\" text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
    }

    private static void RenderAreaChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Gridlines
        for (int t = 1; t <= 4; t++)
        {
            var gy = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"#333\" stroke-width=\"0.5\" stroke-dasharray=\"3,3\"/>");
        }

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        // Render in reverse order so first series is on top
        for (int s = series.Count - 1; s >= 0; s--)
        {
            var points = new List<(double x, double y, double val)>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / maxVal) * ph;
                points.Add((px, py, series[s].values[c]));
            }
            if (points.Count > 0)
            {
                var ptStr = string.Join(" ", points.Select(p => $"{p.x:0.#},{p.y:0.#}"));
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastX = ox + (catCount > 1 ? (double)pw * (series[s].values.Length - 1) / (catCount - 1) : pw / 2.0);
                var polygonPoints = $"{firstX:0.#},{oy + ph} {ptStr} {lastX:0.#},{oy + ph}";
                sb.AppendLine($"        <polygon points=\"{polygonPoints}\" fill=\"{colors[s]}\" opacity=\"0.4\"/>");
                sb.AppendLine($"        <polyline points=\"{ptStr}\" fill=\"none\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
                // Value labels on top series only (first series = last rendered)
                if (s == 0)
                {
                    foreach (var p in points)
                    {
                        var vlabel = p.val % 1 == 0 ? $"{(int)p.val}" : $"{p.val:0.#}";
                        sb.AppendLine($"        <text x=\"{p.x:0.#}\" y=\"{p.y - 6:0.#}\" fill=\"#aaa\" font-size=\"7\" text-anchor=\"middle\">{vlabel}</text>");
                    }
                }
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Value axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    private void RenderComboChartSvg(StringBuilder sb, DocumentFormat.OpenXml.Drawing.Charts.PlotArea plotArea,
        List<(string name, double[] values)> seriesList, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        // Combo: detect series type from parent chart element
        var barIndices = new HashSet<int>();
        var lineIndices = new HashSet<int>();
        var areaIndices = new HashSet<int>();
        var idx = 0;
        foreach (var chartEl in plotArea.ChildElements)
        {
            var serElements = chartEl.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser").ToList();
            if (serElements.Count == 0) continue;
            var localName = chartEl.LocalName.ToLowerInvariant();
            var isBar = localName.Contains("bar");
            var isArea = localName.Contains("area");
            foreach (var _ in serElements)
            {
                if (isBar) barIndices.Add(idx);
                else if (isArea) areaIndices.Add(idx);
                else lineIndices.Add(idx);
                idx++;
            }
        }

        var allValues = seriesList.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, seriesList.Max(s => s.values.Length));

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        // Bar series
        var barSeries = barIndices.Where(i => i < seriesList.Count).ToList();
        var barCount = barSeries.Count;
        if (barCount > 0)
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.7 / barCount;
            var gap = groupW * 0.15;
            for (int bi = 0; bi < barCount; bi++)
            {
                var s = barSeries[bi];
                for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
                {
                    var val = seriesList[s].values[c];
                    var barH = (val / maxVal) * ph;
                    var bx = ox + c * groupW + gap + bi * barW;
                    var by = oy + ph - barH;
                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                }
            }
        }

        // Area series (render before lines so lines appear on top)
        foreach (var s in areaIndices.Where(i => i < seriesList.Count))
        {
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (seriesList[s].values[c] / maxVal) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastX = ox + (catCount > 1 ? (double)pw * (seriesList[s].values.Length - 1) / (catCount - 1) : pw / 2.0);
                var polygonPoints = $"{firstX:0.#},{oy + ph} {string.Join(" ", points)} {lastX:0.#},{oy + ph}";
                sb.AppendLine($"        <polygon points=\"{polygonPoints}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.3\"/>");
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2\"/>");
            }
        }

        // Line series
        foreach (var s in lineIndices.Where(i => i < seriesList.Count))
        {
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (seriesList[s].values[c] / maxVal) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2.5\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s % colors.Count]}\"/>");
                }
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (double)pw * c / Math.Max(catCount, 1) + (double)pw / Math.Max(catCount, 1) / 2;
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Value axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    private static void RenderRadarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH)
    {
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        if (catCount < 3) return;
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.35;

        // Grid lines (3 rings)
        for (int ring = 1; ring <= 3; ring++)
        {
            var rr = r * ring / 3;
            var gridPoints = new List<string>();
            for (int c = 0; c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                gridPoints.Add($"{cx + rr * Math.Cos(angle):0.#},{cy + rr * Math.Sin(angle):0.#}");
            }
            sb.AppendLine($"        <polygon points=\"{string.Join(" ", gridPoints)}\" fill=\"none\" stroke=\"#444\" stroke-width=\"0.5\"/>");
        }

        // Axis lines
        for (int c = 0; c < catCount; c++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            var ax = cx + r * Math.Cos(angle);
            var ay = cy + r * Math.Sin(angle);
            sb.AppendLine($"        <line x1=\"{cx:0.#}\" y1=\"{cy:0.#}\" x2=\"{ax:0.#}\" y2=\"{ay:0.#}\" stroke=\"#444\" stroke-width=\"0.5\"/>");
        }

        // Data series
        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                var val = series[s].values[c] / maxVal * r;
                points.Add($"{cx + val * Math.Cos(angle):0.#},{cy + val * Math.Sin(angle):0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", points)}\" fill=\"{colors[s]}\" fill-opacity=\"0.2\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s]}\"/>");
                }
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            var lx = cx + (r + 15) * Math.Cos(angle);
            var ly = cy + (r + 15) * Math.Sin(angle);
            var anchor = Math.Abs(Math.Cos(angle)) < 0.1 ? "middle" : (Math.Cos(angle) > 0 ? "start" : "end");
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#999\" font-size=\"9\" text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    private static void RenderBubbleChartSvg(StringBuilder sb,
        DocumentFormat.OpenXml.Drawing.Charts.PlotArea plotArea,
        List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        // Read X, Y, and bubble size from each series in the BubbleChart
        var bubbleSeries = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent?.LocalName == "bubbleChart").ToList();

        var allX = new List<double>();
        var allY = new List<double>();
        var allSize = new List<double>();
        var seriesData = new List<(double[] x, double[] y, double[] size)>();

        for (int s = 0; s < bubbleSeries.Count; s++)
        {
            var ser = bubbleSeries[s];
            var xVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "xVal")) ?? [];
            var yVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal")) ?? [];
            var sizeVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "bubbleSize")) ?? yVals;
            seriesData.Add((xVals, yVals, sizeVals));
            allX.AddRange(xVals);
            allY.AddRange(yVals);
            allSize.AddRange(sizeVals);
        }

        // Fallback if no bubble series found
        if (seriesData.Count == 0)
        {
            // Use regular series data as Y, index as X
            foreach (var s in series)
            {
                var xVals = Enumerable.Range(0, s.values.Length).Select(i => (double)i).ToArray();
                seriesData.Add((xVals, s.values, s.values));
                allX.AddRange(xVals);
                allY.AddRange(s.values);
                allSize.AddRange(s.values);
            }
        }

        if (allY.Count == 0) return;
        var minX = allX.Count > 0 ? allX.Min() : 0;
        var maxX = allX.Count > 0 ? allX.Max() : 1;
        if (maxX <= minX) maxX = minX + 1;
        var minY = allY.Min();
        var maxY = allY.Max();
        if (maxY <= minY) maxY = minY + 1;
        var maxSize = allSize.Count > 0 ? allSize.Max() : 1;
        if (maxSize <= 0) maxSize = 1;
        var maxRadius = Math.Min(pw, ph) * 0.08;

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        for (int s = 0; s < seriesData.Count; s++)
        {
            var (xVals, yVals, sizeVals) = seriesData[s];
            var count = Math.Min(xVals.Length, yVals.Length);
            for (int i = 0; i < count; i++)
            {
                var bx = ox + ((xVals[i] - minX) / (maxX - minX)) * pw;
                var by = oy + ph - ((yVals[i] - minY) / (maxY - minY)) * ph;
                var sz = i < sizeVals.Length ? sizeVals[i] : yVals[i];
                var r = (sz / maxSize) * maxRadius + 4;
                sb.AppendLine($"        <circle cx=\"{bx:0.#}\" cy=\"{by:0.#}\" r=\"{r:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.5\"/>");
            }
        }

        // X axis labels (5 ticks)
        for (int t = 0; t <= 4; t++)
        {
            var val = minX + (maxX - minX) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var tx = ox + (double)pw * t / 4;
            sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"8\" text-anchor=\"middle\">{label}</text>");
        }

        // Y axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = minY + (maxY - minY) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    private static void RenderStockChartSvg(StringBuilder sb,
        DocumentFormat.OpenXml.Drawing.Charts.PlotArea plotArea,
        List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        var minVal = allValues.Min();
        if (maxVal <= minVal) { maxVal = minVal + 1; }
        var range = maxVal - minVal;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Read up/down bar colors from StockChart
        var upColor = "#2ECC71";
        var downColor = "#E74C3C";
        var stockChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.StockChart>();
        if (stockChart != null)
        {
            var upBars = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "upBars");
            var upFill = upBars?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (upFill != null) upColor = $"#{upFill}";
            var downBars = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "downBars");
            var downFill = downBars?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (downFill != null) downColor = $"#{downFill}";
        }

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        var groupW = (double)pw / Math.Max(catCount, 1);

        if (series.Count >= 4)
        {
            // OHLC: Open, High, Low, Close
            for (int c = 0; c < catCount; c++)
            {
                var open = c < series[0].values.Length ? series[0].values[c] : 0;
                var high = c < series[1].values.Length ? series[1].values[c] : 0;
                var low = c < series[2].values.Length ? series[2].values[c] : 0;
                var close = c < series[3].values.Length ? series[3].values[c] : 0;
                var cx = ox + c * groupW + groupW / 2;
                var yHigh = oy + ph - ((high - minVal) / range) * ph;
                var yLow = oy + ph - ((low - minVal) / range) * ph;
                var yOpen = oy + ph - ((open - minVal) / range) * ph;
                var yClose = oy + ph - ((close - minVal) / range) * ph;
                var color = close >= open ? upColor : downColor;
                var barW = groupW * 0.5;

                // High-Low line
                sb.AppendLine($"        <line x1=\"{cx:0.#}\" y1=\"{yHigh:0.#}\" x2=\"{cx:0.#}\" y2=\"{yLow:0.#}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
                // Open-Close body
                var bodyTop = Math.Min(yOpen, yClose);
                var bodyH = Math.Abs(yOpen - yClose);
                if (bodyH < 1) bodyH = 1;
                sb.AppendLine($"        <rect x=\"{cx - barW / 2:0.#}\" y=\"{bodyTop:0.#}\" width=\"{barW:0.#}\" height=\"{bodyH:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
        }
        else
        {
            // Fallback: render as line chart
            RenderLineChartSvg(sb, series, categories, colors, ox, oy, pw, ph);
            return;
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + c * groupW + groupW / 2;
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Value axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = minVal + range * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
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

    private static void RenderBar3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool horizontal)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;

        if (horizontal)
        {
            var hLabelMargin = 50;
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;
            var groupH = (double)ph / Math.Max(catCount, 1);
            var barH = groupH * 0.6 / serCount;
            var gap = groupH * 0.2;

            // Gridlines
            for (int t = 1; t <= 4; t++)
            {
                var gx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"#333\" stroke-width=\"0.5\" stroke-dasharray=\"3,3\"/>");
            }
            // Axis lines
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

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
                    sb.AppendLine($"        <text x=\"{bx + barW + DxIso + 4:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"#aaa\" font-size=\"7\" text-anchor=\"start\" dominant-baseline=\"middle\">{vlabel}</text>");
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var ly = oy + c * groupH + groupH / 2;
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"#999\" font-size=\"9\" text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var tx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#777\" font-size=\"8\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.6 / serCount;
            var gap = groupW * 0.2;

            // Gridlines
            for (int t = 1; t <= 4; t++)
            {
                var gy = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"#333\" stroke-width=\"0.5\" stroke-dasharray=\"3,3\"/>");
            }
            // Axis lines
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

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
                    sb.AppendLine($"        <text x=\"{bx + barW / 2 + DxIso / 2:0.#}\" y=\"{by + DyIso - 3:0.#}\" fill=\"#aaa\" font-size=\"7\" text-anchor=\"middle\">{vlabel}</text>");
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var ty = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    private static void RenderPie3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
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
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"white\" font-size=\"9\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");

            startAngle = endAngle;
        }
    }

    private static void RenderLine3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

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
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
    }
}
