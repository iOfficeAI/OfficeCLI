// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for chart and cell-run paths. Mechanically
// extracted from the original god-method Set(); each helper owns one
// path-pattern's full handling.
public partial class ExcelHandler
{
    private List<string> SetChartAxisByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var caChartIdx = int.Parse(m.Groups[1].Value);
        var caRole = m.Groups[2].Value;
        var caDrawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("No charts in this sheet");
        var caAllCharts = GetExcelCharts(caDrawingsPart);
        if (caChartIdx < 1 || caChartIdx > caAllCharts.Count)
            throw new ArgumentException($"Chart {caChartIdx} not found (total: {caAllCharts.Count})");
        var caChartInfo = caAllCharts[caChartIdx - 1];
        if (caChartInfo.IsExtended || caChartInfo.StandardPart == null)
            throw new ArgumentException("Axis Set not supported on extended charts.");
        var axUnsupported = ChartHelper.SetAxisProperties(
            caChartInfo.StandardPart, caRole, properties);
        caChartInfo.StandardPart.ChartSpace?.Save();
        return axUnsupported;
    }

    private List<string> SetChartByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var chartIdx = int.Parse(m.Groups[1].Value);
        var drawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("No charts in this sheet");
        var excelCharts = GetExcelCharts(drawingsPart);
        if (chartIdx < 1 || chartIdx > excelCharts.Count)
            throw new ArgumentException($"Chart {chartIdx} not found (total: {excelCharts.Count})");
        var chartInfo = excelCharts[chartIdx - 1];

        // If series sub-path, prefix all properties with series{N}. for ChartSetter
        var chartProps = properties;
        var isSeriesPath = m.Groups[2].Success;
        if (isSeriesPath)
        {
            var seriesIdx = int.Parse(m.Groups[2].Value);
            chartProps = new Dictionary<string, string>();
            foreach (var (key, value) in properties)
                chartProps[$"series{seriesIdx}.{key}"] = value;
        }

        // Chart-level position/size Set — TwoCellAnchor mutation. Skip for series
        // sub-paths (series don't have their own position). Accepts x/y/width/height
        // in the same units as OLE Set and chart Add.
        // CONSISTENCY(chart-position-set): mirrors PPTX path so users learn one
        // vocabulary for all three doc types. Excel mutates a TwoCellAnchor instead
        // of a GraphicFrame Transform because xlsx charts are cell-anchored.
        if (!isSeriesPath)
        {
            var positionUnsupported = ApplyChartPositionSet(
                drawingsPart, chartIdx, chartProps);
            foreach (var k in new[] { "x", "y", "width", "height" })
            {
                var matched = chartProps.Keys
                    .FirstOrDefault(key => key.Equals(k, StringComparison.OrdinalIgnoreCase));
                if (matched != null && !positionUnsupported.Contains(matched))
                    chartProps.Remove(matched);
            }
        }

        if (chartInfo.StandardPart != null)
        {
            var unsup = ChartHelper.SetChartProperties(chartInfo.StandardPart, chartProps);
            chartInfo.StandardPart.ChartSpace?.Save();
            return unsup;
        }
        else if (chartInfo.ExtendedPart != null)
        {
            // cx:chart — delegates to ChartExBuilder.SetChartProperties.
            return ChartExBuilder.SetChartProperties(chartInfo.ExtendedPart, chartProps);
        }
        else
        {
            return chartProps.Keys.ToList();
        }
    }

    private List<string> SetCellRunByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var runCellRef = m.Groups[1].Value.ToUpperInvariant();
        var runIdx = int.Parse(m.Groups[2].Value);

        var runSheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet data not found");
        var runCell = FindOrCreateCell(runSheetData, runCellRef);

        if (runCell.DataType?.Value != CellValues.SharedString ||
            !int.TryParse(runCell.CellValue?.Text, out var sstIdx))
            throw new ArgumentException($"Cell {runCellRef} is not a rich text cell");

        var sstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var ssi = sstPart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx)
            ?? throw new ArgumentException($"SharedString entry {sstIdx} not found");

        var runs = ssi.Elements<Run>().ToList();
        if (runIdx < 1 || runIdx > runs.Count)
            throw new ArgumentException($"Run index {runIdx} out of range (1-{runs.Count})");

        var run = runs[runIdx - 1];
        var rProps = run.RunProperties ?? run.PrependChild(new RunProperties());

        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text" or "value":
                    var textEl = run.GetFirstChild<Text>();
                    if (textEl != null) textEl.Text = value;
                    else run.AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                    break;
                case "bold":
                    rProps.RemoveAllChildren<Bold>();
                    if (ParseHelpers.IsTruthy(value)) rProps.InsertAt(new Bold(), 0);
                    break;
                case "italic":
                    rProps.RemoveAllChildren<Italic>();
                    if (ParseHelpers.IsTruthy(value)) rProps.AppendChild(new Italic());
                    break;
                case "strike":
                    rProps.RemoveAllChildren<Strike>();
                    if (ParseHelpers.IsTruthy(value)) rProps.AppendChild(new Strike());
                    break;
                case "underline":
                    rProps.RemoveAllChildren<Underline>();
                    if (!string.IsNullOrEmpty(value) && value != "false" && value != "none")
                    {
                        var ul = new Underline();
                        if (value.ToLowerInvariant() == "double") ul.Val = UnderlineValues.Double;
                        rProps.AppendChild(ul);
                    }
                    break;
                case "superscript":
                    rProps.RemoveAllChildren<VerticalTextAlignment>();
                    if (ParseHelpers.IsTruthy(value))
                        rProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript });
                    break;
                case "subscript":
                    rProps.RemoveAllChildren<VerticalTextAlignment>();
                    if (ParseHelpers.IsTruthy(value))
                        rProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript });
                    break;
                case "size":
                    rProps.RemoveAllChildren<FontSize>();
                    rProps.AppendChild(new FontSize { Val = ParseHelpers.ParseFontSize(value) });
                    break;
                case "color":
                    rProps.RemoveAllChildren<Color>();
                    rProps.AppendChild(new Color { Rgb = ParseHelpers.NormalizeArgbColor(value) });
                    break;
                case "font":
                    rProps.RemoveAllChildren<RunFont>();
                    rProps.AppendChild(new RunFont { Val = value });
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ReorderRunProperties(rProps);
        sstPart!.SharedStringTable!.Save();
        SaveWorksheet(worksheet);
        return unsupported;
    }
}
