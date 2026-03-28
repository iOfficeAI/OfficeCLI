// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Chart Rendering ====================

    private void RenderChartHtml(StringBuilder sb, Drawing drawing, OpenXmlElement chartRef)
    {
        var relId = chartRef.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value;
        if (relId == null) return;

        try
        {
            var chartPart = _doc.MainDocumentPart?.GetPartById(relId) as DocumentFormat.OpenXml.Packaging.ChartPart;
            if (chartPart?.ChartSpace == null) return;

            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            int svgW = extent?.Cx?.Value > 0 ? (int)(extent.Cx.Value / 9525) : 500;
            int svgH = extent?.Cy?.Value > 0 ? (int)(extent.Cy.Value / 9525) : 300;

            // Use the shared ChartSvgRenderer
            var chartSpace = chartPart.ChartSpace;
            var chart = chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            if (chart == null) return;

            var plotArea = chart.PlotArea;
            if (plotArea == null) return;

            // Extract chart data using ChartHelper
            var chartType = Core.ChartHelper.DetectChartType(plotArea) ?? "column";
            var categories = Core.ChartHelper.ReadCategories(plotArea) ?? [];
            var seriesList = Core.ChartHelper.ReadAllSeries(plotArea);
            if (seriesList.Count == 0) return;

            // Get title
            var title = chart.Title;
            string? titleText = null;
            if (title != null)
            {
                var titleRuns = title.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>()
                    .SelectMany(ct => ct.Descendants<A.Run>())
                    .Select(r => r.GetFirstChild<A.Text>()?.Text)
                    .Where(t => t != null);
                titleText = string.Join("", titleRuns);
            }

            // Read series colors: collect all ser elements from the specific chart type element
            // (barChart/lineChart/pieChart etc.) to match order with ChartHelper.ReadAllSeries
            var chartTypeEl = plotArea.Elements().FirstOrDefault(e =>
                e.LocalName is "barChart" or "bar3DChart" or "lineChart" or "line3DChart"
                    or "pieChart" or "pie3DChart" or "doughnutChart" or "areaChart" or "area3DChart"
                    or "scatterChart" or "radarChart" or "bubbleChart" or "ofPieChart");
            var serElements = chartTypeEl?.Elements().Where(e => e.LocalName == "ser").ToList() ?? [];
            var colors = new List<string>();
            for (int si = 0; si < seriesList.Count; si++)
            {
                string? seriesColor = null;
                if (si < serElements.Count)
                {
                    // Look for solidFill in the series' spPr
                    var spPr = serElements[si].Elements().FirstOrDefault(e => e.LocalName == "spPr");
                    var solidFill = spPr?.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
                    if (solidFill != null)
                    {
                        var srgb = solidFill.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
                        seriesColor = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                        if (seriesColor != null) seriesColor = $"#{seriesColor}";
                    }
                }
                colors.Add(seriesColor ?? Core.ChartSvgRenderer.DefaultColors[si % Core.ChartSvgRenderer.DefaultColors.Length]);
            }

            // Render SVG chart (use dark label colors for white background)
            var renderer = new Core.ChartSvgRenderer
            {
                CatColor = "#333333",
                AxisColor = "#555555",
                ValueColor = "#444444"
            };

            sb.Append($"<div style=\"margin:0.5em 0;text-align:center\">");
            if (!string.IsNullOrEmpty(titleText))
                sb.Append($"<div style=\"font-weight:bold;margin-bottom:4px\">{HtmlEncode(titleText)}</div>");

            sb.Append($"<svg width=\"{svgW}\" height=\"{svgH}\" xmlns=\"http://www.w3.org/2000/svg\" style=\"background:white\">");

            int margin = 40;
            int plotW = svgW - margin * 2;
            int plotH = svgH - margin * 2;
            var seriesColors = colors;

            switch (chartType)
            {
                case "bar":
                    renderer.RenderBarChartSvg(sb, seriesList, categories, seriesColors, margin, margin, plotW, plotH, true, true, false);
                    break;
                case "column":
                    renderer.RenderBarChartSvg(sb, seriesList, categories, seriesColors, margin, margin, plotW, plotH, false, true, false);
                    break;
                case "line":
                    renderer.RenderLineChartSvg(sb, seriesList, categories, seriesColors, margin, margin, plotW, plotH, false);
                    break;
                case "pie":
                case "doughnut":
                    renderer.RenderPieChartSvg(sb, seriesList, categories, seriesColors, svgW, svgH, chartType == "doughnut" ? 50 : 0, false);
                    break;
                case "area":
                    renderer.RenderAreaChartSvg(sb, seriesList, categories, seriesColors, margin, margin, plotW, plotH, false);
                    break;
                case "scatter":
                    // Scatter rendered as line chart with markers (closest available approximation)
                    renderer.RenderLineChartSvg(sb, seriesList, categories, seriesColors, margin, margin, plotW, plotH, true);
                    break;
                case "radar":
                    renderer.RenderRadarChartSvg(sb, seriesList, categories, seriesColors, svgW, svgH, 30);
                    break;
                default:
                    // Fallback: render as column chart
                    renderer.RenderBarChartSvg(sb, seriesList, categories, seriesColors, margin, margin, plotW, plotH, false, true, false);
                    break;
            }

            sb.Append("</svg>");

            // Render legend if multiple series
            if (seriesList.Count > 1)
            {
                sb.Append("<div style=\"display:flex;justify-content:center;gap:16px;margin-top:4px;font-size:9pt\">");
                for (int li = 0; li < seriesList.Count; li++)
                {
                    var lColor = li < seriesColors.Count ? seriesColors[li] : "#999";
                    sb.Append($"<span><span style=\"display:inline-block;width:12px;height:12px;background:{lColor};margin-right:4px;vertical-align:middle\"></span>{HtmlEncode(seriesList[li].name)}</span>");
                }
                sb.Append("</div>");
            }

            sb.Append("</div>");
        }
        catch (Exception ex)
        {
            sb.Append($"<div style=\"padding:1em;color:#999;text-align:center\">[Chart: {HtmlEncode(ex.Message)}]</div>");
        }
    }
}
