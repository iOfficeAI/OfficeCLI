// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

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
            // cx:chart (extended) path — different part type, different extractor.
            var anyPart = _doc.MainDocumentPart?.GetPartById(relId);
            if (anyPart is ExtendedChartPart extPart)
            {
                RenderChartExHtml(sb, drawing, extPart);
                return;
            }

            var chartPart = anyPart as ChartPart;
            if (chartPart?.ChartSpace == null) return;

            var chart = chartPart.ChartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            if (chart == null) return;
            var plotArea = chart.PlotArea;
            if (plotArea == null) return;

            // Extract all chart metadata via shared helper
            var info = ChartSvgRenderer.ExtractChartInfo(plotArea, chart);
            if (info.Series.Count == 0) return;

            // Chart dimensions from drawing extent
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            int svgW = extent?.Cx?.Value > 0 ? (int)(extent.Cx.Value / 9525) : 500;
            int svgH = extent?.Cy?.Value > 0 ? (int)(extent.Cy.Value / 9525) : 300;

            // Renderer — use chart XML colors if available, else reasonable defaults
            var renderer = new ChartSvgRenderer
            {
                ThemeAccentColors = ChartSvgRenderer.BuildThemeAccentColors(GetThemeColors()),
                CatColor = info.CatFontColor != null ? $"#{info.CatFontColor}" : "#333333",
                AxisColor = info.ValFontColor != null ? $"#{info.ValFontColor}" : "#555555",
                ValueColor = info.ValFontColor != null ? $"#{info.ValFontColor}" : "#444444",
                GridColor = info.GridlineColor != null ? $"#{info.GridlineColor}" : "#ddd",
                AxisLineColor = info.AxisLineColor != null ? $"#{info.AxisLineColor}" : "#999",
                ValFontPx = info.ValFontPx,
                CatFontPx = info.CatFontPx
            };

            var titleH = string.IsNullOrEmpty(info.Title) ? 0 : 24;
            var legendH = info.HasLegend ? 24 : 0;
            var chartSvgH = svgH - titleH - legendH;

            sb.Append($"<div style=\"margin:0.5em 0;text-align:center\">");
            if (!string.IsNullOrEmpty(info.Title))
                sb.Append($"<div style=\"font-weight:bold;margin-bottom:4px;font-size:{info.TitleFontSize}\">{HtmlEncode(info.Title)}</div>");

            var bgStyle = info.ChartFillColor != null ? $"background:#{info.ChartFillColor};" : "background:white;";
            sb.Append($"<svg width=\"{svgW}\" height=\"{chartSvgH}\" xmlns=\"http://www.w3.org/2000/svg\" style=\"{bgStyle}\">");

            renderer.RenderChartSvgContent(sb, info, svgW, chartSvgH);

            sb.Append("</svg>");

            renderer.RenderLegendHtml(sb, info, "#333");

            sb.Append("</div>");
        }
        catch (Exception ex)
        {
            sb.Append($"<div style=\"padding:1em;color:#999;text-align:center\">[Chart: {HtmlEncode(ex.Message)}]</div>");
        }
    }

    /// <summary>
    /// Render a cx:chart (Office 2016 extended chart — histogram, funnel,
    /// treemap, sunburst, boxWhisker) inside a Word document. Mirrors the
    /// regular-chart path in <see cref="RenderChartHtml"/>, but uses
    /// <see cref="ChartSvgRenderer.ExtractCxChartInfo"/> and skips the
    /// a:plotArea extraction (cx has its own PlotArea shape).
    /// </summary>
    private void RenderChartExHtml(StringBuilder sb, Drawing drawing, ExtendedChartPart extPart)
    {
        try
        {
            var chart = extPart.ChartSpace?
                .GetFirstChild<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Chart>();
            if (chart == null) return;

            var info = ChartSvgRenderer.ExtractCxChartInfo(chart);
            if (info.Series.Count == 0) return;

            // Chart dimensions from the drawing extent, same as regular charts.
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            int svgW = extent?.Cx?.Value > 0 ? (int)(extent.Cx.Value / 9525) : 500;
            int svgH = extent?.Cy?.Value > 0 ? (int)(extent.Cy.Value / 9525) : 300;

            var renderer = new ChartSvgRenderer
            {
                ThemeAccentColors = ChartSvgRenderer.BuildThemeAccentColors(GetThemeColors()),
                CatColor = info.CatFontColor != null ? $"#{info.CatFontColor}" : "#333333",
                AxisColor = info.ValFontColor != null ? $"#{info.ValFontColor}" : "#555555",
                ValueColor = info.ValFontColor != null ? $"#{info.ValFontColor}" : "#444444",
                GridColor = info.GridlineColor != null ? $"#{info.GridlineColor}" : "#ddd",
                AxisLineColor = info.AxisLineColor != null ? $"#{info.AxisLineColor}" : "#999",
                ValFontPx = info.ValFontPx,
                CatFontPx = info.CatFontPx,
            };

            var titleH = string.IsNullOrEmpty(info.Title) ? 0 : 24;
            var chartSvgH = svgH - titleH;
            if (chartSvgH < 80) return;

            sb.Append("<div style=\"margin:0.5em 0;text-align:center\">");
            if (!string.IsNullOrEmpty(info.Title))
                sb.Append($"<div style=\"font-weight:bold;margin-bottom:4px;font-size:{info.TitleFontSize}\">{HtmlEncode(info.Title)}</div>");
            sb.Append($"<svg width=\"{svgW}\" height=\"{chartSvgH}\" xmlns=\"http://www.w3.org/2000/svg\" style=\"background:white;\">");
            renderer.RenderChartSvgContent(sb, info, svgW, chartSvgH);
            sb.Append("</svg>");
            sb.Append("</div>");
        }
        catch (Exception ex)
        {
            sb.Append($"<div style=\"padding:1em;color:#999;text-align:center\">[cxChart: {HtmlEncode(ex.Message)}]</div>");
        }
    }
}
