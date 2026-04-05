// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Render all charts in a worksheet as SVG elements, respecting anchor positions.
    /// Charts with overlapping row ranges are placed side-by-side using flex layout.
    /// </summary>
    private void RenderSheetCharts(StringBuilder sb, WorksheetPart worksheetPart)
    {
        var charts = CollectSheetCharts(worksheetPart);
        foreach (var (_, _, html) in charts)
            sb.Append(html);
    }

    /// <summary>
    /// Pre-render all charts and return them with their anchor row positions.
    /// Charts with overlapping row ranges are grouped into flex rows.
    /// </summary>
    private List<(int fromRow, int toRow, string html)> CollectSheetCharts(WorksheetPart worksheetPart)
    {
        var result = new List<(int fromRow, int toRow, string html)>();
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing == null) return result;

        // Find all graphic frames that contain chart references
        var chartFrames = drawingsPart.WorksheetDrawing
            .Descendants<XDR.GraphicFrame>()
            .Where(gf => gf.Descendants<C.ChartReference>().Any())
            .ToList();

        if (chartFrames.Count == 0) return result;

        // Read anchor positions and group charts into rows (overlapping row ranges = same row)
        var chartAnchors = chartFrames.Select(gf =>
        {
            var anchor = gf.Parent as XDR.TwoCellAnchor;
            int fromRow = 0, toRow = 0, fromCol = 0;
            if (anchor?.FromMarker != null && anchor?.ToMarker != null)
            {
                int.TryParse(anchor.FromMarker.RowId?.Text, out fromRow);
                int.TryParse(anchor.ToMarker.RowId?.Text, out toRow);
                int.TryParse(anchor.FromMarker.ColumnId?.Text, out fromCol);
            }
            return (gf, fromRow, toRow, fromCol);
        }).OrderBy(x => x.fromRow).ThenBy(x => x.fromCol).ToList();

        // Group into rows: charts whose row ranges overlap go in the same flex row
        var groups = new List<(int fromRow, int toRow, List<XDR.GraphicFrame> frames)>();
        int currentRowEnd = -1;
        List<XDR.GraphicFrame>? currentGroup = null;
        int currentFromRow = 0;
        foreach (var (gf, fromRow, toRow, _) in chartAnchors)
        {
            if (currentGroup == null || fromRow >= currentRowEnd)
            {
                currentGroup = new List<XDR.GraphicFrame>();
                currentFromRow = fromRow;
                currentRowEnd = toRow;
                groups.Add((fromRow, toRow, currentGroup));
            }
            else
            {
                currentRowEnd = Math.Max(currentRowEnd, toRow);
                // Update toRow in the group
                groups[^1] = (groups[^1].fromRow, currentRowEnd, currentGroup);
            }
            currentGroup.Add(gf);
        }

        foreach (var (fromRow, toRow, frames) in groups)
        {
            var chartSb = new StringBuilder();
            if (frames.Count > 1)
            {
                chartSb.AppendLine("<div style=\"display:flex;gap:16px;flex-wrap:wrap\">");
                foreach (var gf in frames)
                    RenderExcelChart(chartSb, gf, drawingsPart, worksheetPart);
                chartSb.AppendLine("</div>");
            }
            else
            {
                RenderExcelChart(chartSb, frames[0], drawingsPart, worksheetPart);
            }
            result.Add((fromRow, toRow, chartSb.ToString()));
        }

        return result;
    }

    private void RenderExcelChart(StringBuilder sb, XDR.GraphicFrame gf,
        DrawingsPart drawingsPart, WorksheetPart worksheetPart)
    {
        // 1. Get chart reference and load ChartPart
        var chartRef = gf.Descendants<C.ChartReference>().FirstOrDefault();
        if (chartRef?.Id?.Value == null) return;

        C.Chart? chart;
        C.PlotArea? plotArea;
        try
        {
            var chartPart = (ChartPart)drawingsPart.GetPartById(chartRef.Id.Value);
            chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
            plotArea = chart?.GetFirstChild<C.PlotArea>();
            if (plotArea == null) return;
        }
        catch { return; }

        // 2. Read chart data using shared ChartHelper
        var chartType = ChartHelper.DetectChartType(plotArea) ?? "bar";
        var categories = ChartHelper.ReadCategories(plotArea) ?? [];
        var seriesList = ChartHelper.ReadAllSeries(plotArea);
        if (seriesList.Count == 0) return;

        // 2b. Resolve series names from cell references when strCache is missing
        if (seriesList.Any(s => s.name == "?"))
        {
            var nameSerEls = plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser" && e.Parent != null &&
                    (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart")))
                .ToList();
            for (int i = 0; i < seriesList.Count && i < nameSerEls.Count; i++)
            {
                if (seriesList[i].name != "?") continue;
                var strRef = nameSerEls[i].GetFirstChild<C.SeriesText>()
                    ?.Descendants<C.StringReference>().FirstOrDefault();
                var formula = strRef?.GetFirstChild<C.Formula>()?.Text;
                if (!string.IsNullOrEmpty(formula))
                {
                    var resolved = ReadCellRangeAsStrings(formula);
                    if (resolved != null && resolved.Length > 0)
                        seriesList[i] = (resolved[0], seriesList[i].values);
                }
            }
        }

        // 2c. Resolve cell references when cache is missing (chart references other sheets)
        var needsCatResolve = categories.Length == 0;
        var needsValResolve = seriesList.All(s => s.values.Length == 0);
        if (needsCatResolve || needsValResolve)
        {
            ResolveChartDataFromCells(plotArea, ref categories, ref seriesList, needsCatResolve, needsValResolve);
            if (seriesList.All(s => s.values.Length == 0)) return;
        }

        // 3. Extract all chart metadata via shared helper
        var info = ChartSvgRenderer.ExtractChartInfo(plotArea, chart);
        // Override with locally-resolved data (Excel cell resolution may have updated categories/series)
        info.ChartType = chartType;
        info.Categories = categories;
        info.Series = seriesList;
        if (info.Series.Count == 0) return;
        // Ensure colors match series count (ExtractChartInfo may have extracted for a different count)
        while (info.Colors.Count < info.Series.Count)
            info.Colors.Add(ChartSvgRenderer.DefaultColors[info.Colors.Count % ChartSvgRenderer.DefaultColors.Length]);
        if (info.Colors.Count > info.Series.Count && !info.ChartType.Contains("pie") && !info.ChartType.Contains("doughnut"))
            info.Colors = info.Colors.Take(info.Series.Count).ToList();

        // 4. Estimate chart dimensions from TwoCellAnchor using actual column widths
        var colWidths = GetColumnWidths(GetSheet(worksheetPart));
        var (widthPt, heightPt) = EstimateChartSize(gf, colWidths);

        // 5. Create renderer — colors from OOXML with Excel-appropriate fallbacks
        var renderer = new ChartSvgRenderer
        {
            ValueColor = info.ValFontColor ?? "#333",
            CatColor = info.CatFontColor ?? "#555",
            AxisColor = info.ValFontColor ?? "#666",
            GridColor = info.GridlineColor ?? "#ddd",
            AxisLineColor = info.AxisLineColor ?? "#999",
            ValFontPx = info.ValFontPx,
            CatFontPx = info.CatFontPx
        };

        // 6. Build SVG
        var svgW = Math.Max(widthPt, 225);
        var svgH = Math.Max(heightPt, 150);
        // Title/legend height from actual font sizes
        var titleFontPt = 10.0;
        if (!string.IsNullOrEmpty(info.TitleFontSize) && double.TryParse(info.TitleFontSize.Replace("pt", ""), out var tfp))
            titleFontPt = tfp;
        var titleH = string.IsNullOrEmpty(info.Title) ? 0 : (int)(titleFontPt * 1.6 + 8);
        var legendFontPt = 8.0;
        if (!string.IsNullOrEmpty(info.LegendFontSize) && double.TryParse(info.LegendFontSize.Replace("pt", ""), out var lfp))
            legendFontPt = lfp;
        var legendH = info.HasLegend ? (int)(legendFontPt * 1.6 + 12) : 0;
        var chartSvgH = svgH - titleH - legendH;
        if (chartSvgH < 80) return;

        var bgStyle = info.ChartFillColor != null ? $"background:#{info.ChartFillColor};" : "";
        // Use estimated width as max-width, but allow stretching to fill parent (e.g. colspan td)
        sb.AppendLine($"<div class=\"chart-container\" style=\"max-width:max({svgW}pt,100%);flex:1;min-width:200pt;{bgStyle}\">");

        var titleColor = info.TitleFontColor ?? "#333";
        if (!string.IsNullOrEmpty(info.Title))
            sb.AppendLine($"  <div style=\"text-align:center;font-size:{info.TitleFontSize};font-weight:bold;padding:6px 0;color:{titleColor}\">{HtmlEncode(info.Title)}</div>");

        sb.AppendLine($"  <svg viewBox=\"0 0 {svgW} {chartSvgH}\" style=\"width:100%;height:auto\" preserveAspectRatio=\"xMidYMin meet\">");

        renderer.RenderChartSvgContent(sb, info, svgW, chartSvgH);

        sb.AppendLine("  </svg>");

        var legendColor = info.LegendFontColor ?? "#555";
        renderer.RenderLegendHtml(sb, info, legendColor);

        sb.AppendLine("</div>");
    }

    /// <summary>
    /// Estimate chart size from the TwoCellAnchor parent, using actual column widths when available.
    /// </summary>
    private static (int widthPt, int heightPt) EstimateChartSize(XDR.GraphicFrame gf,
        Dictionary<int, double>? colWidths = null)
    {
        var anchor = gf.Parent as XDR.TwoCellAnchor;
        if (anchor == null) return (450, 263);

        var from = anchor.FromMarker;
        var to = anchor.ToMarker;
        if (from == null || to == null) return (450, 263);

        var fromCol = int.TryParse(from.ColumnId?.Text, out var fc) ? fc : 0;
        var toCol = int.TryParse(to.ColumnId?.Text, out var tc) ? tc : 0;
        var fromRow = int.TryParse(from.RowId?.Text, out var fr) ? fr : 0;
        var toRow = int.TryParse(to.RowId?.Text, out var tr) ? tr : 0;

        var fromColOff = long.TryParse(from.ColumnOffset?.Text, out var fco) ? fco : 0;
        var toColOff = long.TryParse(to.ColumnOffset?.Text, out var tco) ? tco : 0;
        var fromRowOff = long.TryParse(from.RowOffset?.Text, out var fro) ? fro : 0;
        var toRowOff = long.TryParse(to.RowOffset?.Text, out var tro) ? tro : 0;

        // Sum actual column widths; fall back to 48pt for columns without explicit width
        double totalWidth = 0;
        for (int c = fromCol + 1; c <= toCol; c++)
            totalWidth += (colWidths != null && colWidths.TryGetValue(c, out var w)) ? w : 48.0;
        totalWidth += (toColOff - fromColOff) / 12700.0;

        // Default row height ~15pt; offsets in EMU (1pt = 12700 EMU)
        double totalHeight = (toRow - fromRow) * 15.0 + (toRowOff - fromRowOff) / 12700.0;

        return ((int)Math.Max(totalWidth, 225), (int)Math.Max(totalHeight, 150));
    }

    /// <summary>
    /// Resolve chart data from actual cells when the chart XML has no cache.
    /// Parses formula references like "'Income Statement'!$B$10:$D$10" and reads cell values.
    /// </summary>
    private void ResolveChartDataFromCells(C.PlotArea plotArea,
        ref string[] categories, ref List<(string name, double[] values)> seriesList,
        bool resolveCats, bool resolveVals)
    {
        if (resolveCats)
        {
            var catRef = ChartHelper.ReadCategoriesRef(plotArea);
            if (catRef != null)
            {
                var resolved = ReadCellRangeAsStrings(catRef);
                if (resolved != null) categories = resolved;
            }
        }

        if (resolveVals)
        {
            var newSeries = new List<(string name, double[] values)>();
            foreach (var ser in plotArea.Descendants<OpenXmlCompositeElement>()
                .Where(e => e.LocalName == "ser" && e.Parent != null &&
                    (e.Parent.LocalName.Contains("Chart") || e.Parent.LocalName.Contains("chart"))))
            {
                var serText = ser.GetFirstChild<C.SeriesText>();
                var name = serText?.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? "?";

                var valRef = ChartHelper.ReadFormulaRef(ser.GetFirstChild<C.Values>())
                    ?? ChartHelper.ReadFormulaRef(ser.Elements<OpenXmlCompositeElement>()
                        .FirstOrDefault(e => e.LocalName == "yVal"));

                double[] values = [];
                if (valRef != null)
                {
                    var resolved = ReadCellRangeAsDoubles(valRef);
                    if (resolved != null) values = resolved;
                }

                newSeries.Add((name, values));
            }
            if (newSeries.Count > 0) seriesList = newSeries;
        }
    }

    /// <summary>
    /// Parse a cell range reference like "'Sheet Name'!$B$1:$D$1" and return cell values as strings.
    /// </summary>
    private string[]? ReadCellRangeAsStrings(string formula)
    {
        var (sheetData, startCol, startRow, endCol, endRow) = ParseCellRangeFormula(formula);
        if (sheetData == null) return null;

        var results = new List<string>();
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startCol; c <= endCol; c++)
            {
                var cellRef = GetColumnLetter(c) + r;
                var cell = sheetData.Descendants<Cell>()
                    .FirstOrDefault(cl => cl.CellReference?.Value == cellRef);
                results.Add(cell != null ? GetCellDisplayValue(cell) : "");
            }
        }
        return results.ToArray();
    }

    /// <summary>
    /// Parse a cell range reference and return cell values as doubles.
    /// Uses FormulaEvaluator with cross-sheet support.
    /// </summary>
    private double[]? ReadCellRangeAsDoubles(string formula)
    {
        var (sheetData, startCol, startRow, endCol, endRow) = ParseCellRangeFormula(formula);
        if (sheetData == null) return null;

        var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        var results = new List<double>();
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startCol; c <= endCol; c++)
            {
                var cellRef = GetColumnLetter(c) + r;
                var cell = sheetData.Descendants<Cell>()
                    .FirstOrDefault(cl => cl.CellReference?.Value == cellRef);

                double val = 0;
                if (cell != null)
                {
                    var raw = cell.CellValue?.Text;
                    if (!string.IsNullOrEmpty(raw) && double.TryParse(raw,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var v))
                    {
                        val = v;
                    }
                    else if (cell.CellFormula?.Text != null)
                    {
                        val = evaluator.TryEvaluate(cell.CellFormula.Text) ?? 0;
                    }
                }
                results.Add(val);
            }
        }
        return results.ToArray();
    }

    /// <summary>
    /// Parse "'Sheet Name'!$B$1:$D$1" into (SheetData, startCol, startRow, endCol, endRow).
    /// </summary>
    private (SheetData? sheetData, int startCol, int startRow, int endCol, int endRow) ParseCellRangeFormula(string formula)
    {
        // Pattern: optional 'SheetName'! or SheetName! prefix, then cell range like $B$1:$D$1 or B1:D1
        var match = Regex.Match(formula, @"^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$");
        if (!match.Success) return (null, 0, 0, 0, 0);

        var sheetName = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
        var startColStr = match.Groups[3].Value;
        var startRow = int.Parse(match.Groups[4].Value);
        var endColStr = match.Groups[5].Success ? match.Groups[5].Value : startColStr;
        var endRow = match.Groups[6].Success ? int.Parse(match.Groups[6].Value) : startRow;

        var startCol = ColumnLetterToIndex(startColStr);
        var endCol = ColumnLetterToIndex(endColStr);

        // Find the worksheet by name
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return (null, 0, 0, 0, 0);

        var sheet = workbookPart.Workbook?.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet?.Id?.Value == null) return (null, 0, 0, 0, 0);

        try
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
            return (sheetData, startCol, startRow, endCol, endRow);
        }
        catch { return (null, 0, 0, 0, 0); }
    }

    private static int ColumnLetterToIndex(string col)
    {
        int result = 0;
        foreach (var c in col)
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    private static string GetColumnLetter(int colIndex)
    {
        var result = "";
        while (colIndex > 0)
        {
            colIndex--;
            result = (char)('A' + colIndex % 26) + result;
            colIndex /= 26;
        }
        return result;
    }
}
