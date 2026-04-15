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

    // Chart text color — set per-chart, also used by SvgPreview
    private string _chartValueColor = "#D0D8E0";

    private void RenderChart(StringBuilder sb, GraphicFrame gf, SlidePart slidePart, Dictionary<string, string> themeColors, string? dataPath = null)
    {
        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        // Position and size from p:xfrm
        var pxfrm = gf.GetFirstChild<DocumentFormat.OpenXml.Presentation.Transform>();
        var off = pxfrm?.GetFirstChild<Drawing.Offset>();
        var ext = pxfrm?.GetFirstChild<Drawing.Extents>();
        if (off == null || ext == null) return;

        var x = Units.EmuToPt(off.X?.Value ?? 0);
        var y = Units.EmuToPt(off.Y?.Value ?? 0);
        var w = Units.EmuToPt(ext.Cx?.Value ?? 0);
        var h = Units.EmuToPt(ext.Cy?.Value ?? 0);

        // Get chart part
        var chartEl = gf.Descendants().FirstOrDefault(e => e.LocalName == "chart" && e.NamespaceUri.Contains("chart"));
        var rId = chartEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri.Contains("relationships")).Value;
        if (rId == null) return;

        DocumentFormat.OpenXml.Drawing.Charts.Chart? chart;
        DocumentFormat.OpenXml.Drawing.Charts.PlotArea? plotArea;
        ChartSvgRenderer.ChartInfo info;
        try
        {
            var anyPart = slidePart.GetPartById(rId);
            // cx:chart (extended) path — branch early, extract via ExtractCxChartInfo,
            // skip the regular c:PlotArea pipeline since cx uses its own layout.
            if (anyPart is ExtendedChartPart extPart)
            {
                var cxChart = extPart.ChartSpace?
                    .GetFirstChild<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Chart>();
                if (cxChart == null) return;
                info = ChartSvgRenderer.ExtractCxChartInfo(cxChart);
                chart = null;
                plotArea = null;
            }
            else if (anyPart is ChartPart chartPart)
            {
                chart = chartPart.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                plotArea = chart?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>();
                if (plotArea == null) return;
                info = ChartSvgRenderer.ExtractChartInfo(plotArea, chart);
            }
            else return;
        }
        catch { return; }

        if (info.Series.Count == 0) return;

        // Derive text color from theme
        var chartTextColor = themeColors.TryGetValue("tx1", out var tx1) ? $"#{tx1}"
            : themeColors.TryGetValue("dk1", out var dk1) ? $"#{dk1}" : "#D0D8E0";
        _chartValueColor = chartTextColor;
        var isDarkText = IsColorDark(chartTextColor.TrimStart('#'));

        // Create renderer with theme-derived colors
        var renderer = new ChartSvgRenderer
        {
            ThemeAccentColors = ChartSvgRenderer.BuildThemeAccentColors(themeColors),
            ValueColor = chartTextColor,
            CatColor = chartTextColor,
            AxisColor = chartTextColor,
            GridColor = info.GridlineColor != null ? $"#{info.GridlineColor}" : (isDarkText ? "#ccc" : "#333"),
            AxisLineColor = info.AxisLineColor != null ? $"#{info.AxisLineColor}" : (isDarkText ? "#aaa" : "#555"),
            ValFontPx = info.ValFontPx,
            CatFontPx = info.CatFontPx
        };

        // SVG dimensions (scale EMU to reasonable SVG units)
        var widthEmu = ext.Cx?.Value ?? 3600000;
        var heightEmu = ext.Cy?.Value ?? 2520000;
        var svgW = (int)(widthEmu / 10000.0);
        var svgH = (int)(heightEmu / 10000.0);
        var titleH = string.IsNullOrEmpty(info.Title) ? 0 : 20;
        var chartSvgH = svgH - titleH;

        // Manual layout margins — only regular c:chart has a ManualLayout.
        var plotAreaLayout = plotArea?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Layout>();
        var manualLayout = plotAreaLayout?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ManualLayout>();
        int marginTop, marginRight, marginBottom, marginLeft;
        if (manualLayout != null)
        {
            var mlX = manualLayout.Left?.Val?.Value ?? 0.0;
            var mlY = manualLayout.Top?.Val?.Value ?? 0.0;
            var mlW = manualLayout.Width?.Val?.Value ?? 1.0;
            var mlH = manualLayout.Height?.Val?.Value ?? 1.0;
            marginLeft = Math.Max((int)(mlX * svgW), 5);
            marginTop = Math.Max((int)(mlY * chartSvgH), 5);
            marginRight = Math.Max((int)((1.0 - mlX - mlW) * svgW), 5);
            marginBottom = Math.Max((int)((1.0 - mlY - mlH) * chartSvgH), 5);
        }
        else
        {
            marginTop = 10; marginRight = 15; marginBottom = 25; marginLeft = 40;
        }

        // Container with chart background
        var bgStyle = info.ChartFillColor != null ? $"background:#{info.ChartFillColor};" : "background:transparent;";
        sb.AppendLine($"    <div class=\"shape\"{dataPathAttr} style=\"left:{x}pt;top:{y}pt;width:{w}pt;height:{h}pt;{bgStyle}display:flex;flex-direction:column;overflow:hidden\">");

        // Title
        if (!string.IsNullOrEmpty(info.Title))
            sb.AppendLine($"      <div style=\"text-align:center;font-size:{info.TitleFontSize};font-weight:bold;padding:4px;flex-shrink:0;color:{chartTextColor}\">{ChartSvgRenderer.HtmlEncode(info.Title)}</div>");

        sb.AppendLine($"      <svg viewBox=\"0 0 {svgW} {chartSvgH}\" style=\"width:100%;flex:1;min-height:0\" preserveAspectRatio=\"xMidYMin meet\">");

        renderer.RenderChartSvgContent(sb, info, svgW, chartSvgH, marginLeft, marginTop, marginRight, marginBottom);

        sb.AppendLine("      </svg>");

        renderer.RenderLegendHtml(sb, info, chartTextColor);

        sb.AppendLine("    </div>");
    }
}
