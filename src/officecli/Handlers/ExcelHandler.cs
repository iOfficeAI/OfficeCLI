// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler : IDocumentHandler
{
    private readonly SpreadsheetDocument _doc;
    private readonly string _filePath;
    private readonly HashSet<string> _initialSheetNames;
    public int LastFindMatchCount { get; internal set; }

    public ExcelHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        try
        {
            _doc = SpreadsheetDocument.Open(filePath, editable);
            // Force early validation: access WorkbookPart to catch corrupt packages now
            _ = _doc.WorkbookPart?.Workbook;
            // Capture initial sheet names to detect duplicate additions
            _initialSheetNames = new HashSet<string>(
                GetWorkbook().GetFirstChild<Sheets>()?.Elements<Sheet>()
                    .Select(s => s.Name?.Value ?? "")
                    .Where(n => !string.IsNullOrEmpty(n)) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
        }
        catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException ex)
        {
            throw new InvalidOperationException(
                $"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex);
        }
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return "(empty)";

        if (partPath == "/" || partPath == "/workbook")
            return workbookPart.Workbook?.OuterXml ?? "(empty)";

        if (partPath == "/styles")
        {
            var styleManager = new ExcelStyleManager(workbookPart);
            return styleManager.EnsureStylesPart().Stylesheet!.OuterXml;
        }

        if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            return sst?.SharedStringTable?.OuterXml ?? "(no shared strings)";
        }

        // Drawing part: /SheetName/drawing
        var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
        if (drawingMatch.Success)
        {
            var drawSheetName = drawingMatch.Groups[1].Value;
            var drawWs = FindWorksheet(drawSheetName)
                ?? throw SheetNotFoundException(drawSheetName);
            var dp = drawWs.DrawingsPart
                ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
            return dp.WorksheetDrawing!.OuterXml;
        }

        // Chart part: /SheetName/chart[N] or /chart[N]
        var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartSheetName = chartMatch.Groups[1].Value;
            var chartIdx = int.Parse(chartMatch.Groups[2].Value);
            var chartWs = FindWorksheet(chartSheetName)
                ?? throw SheetNotFoundException(chartSheetName);
            var chartPart = GetChartPart(chartWs, chartIdx);
            return chartPart.ChartSpace!.OuterXml;
        }

        // Global chart: /chart[N] — searches all sheets
        var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
        if (globalChartMatch.Success)
        {
            var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
            var chartPart = GetGlobalChartPart(chartIdx);
            return chartPart.ChartSpace!.OuterXml;
        }

        // Try as sheet name
        var sheetName = partPath.TrimStart('/');
        var worksheet = FindWorksheet(sheetName);
        if (worksheet != null)
        {
            if (startRow.HasValue || endRow.HasValue || cols != null)
                return RawSheetWithFilter(worksheet, startRow, endRow, cols);
            return GetSheet(worksheet).OuterXml;
        }

        // /SheetName/<relId> fallback — resolve a worksheet relationship by id
        // (covers OLE embed parts, image parts, etc. that have no named path).
        // Open XML SDK generates relIds like "rId12" or "Rff3244f593f8481a";
        // accept both forms (any non-slash token starting with R/r).
        var relIdMatch = Regex.Match(partPath, @"^/([^/]+)/([Rr][A-Za-z0-9]+)$");
        if (relIdMatch.Success)
        {
            var relSheetName = relIdMatch.Groups[1].Value;
            var relId = relIdMatch.Groups[2].Value;
            var relWs = FindWorksheet(relSheetName)
                ?? throw SheetNotFoundException(relSheetName);
            try
            {
                var part = relWs.GetPartById(relId);
                if (part != null)
                {
                    var ct = part.ContentType ?? "";
                    bool isText = ct.Contains("xml", StringComparison.OrdinalIgnoreCase)
                        || ct.StartsWith("text/", StringComparison.OrdinalIgnoreCase);
                    using var partStream = part.GetStream();
                    if (isText)
                    {
                        using var reader = new StreamReader(partStream);
                        return reader.ReadToEnd();
                    }
                    long size = 0;
                    try { size = partStream.Length; } catch { /* non-seekable */ }
                    return $"(binary part: {ct}, {size} bytes)";
                }
            }
            catch (KeyNotFoundException)
            {
                // fall through to the unknown-part error
            }
            catch (ArgumentException)
            {
                // fall through to the unknown-part error
            }
        }

        throw new ArgumentException($"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /<SheetName>, /<SheetName>/drawing, /<SheetName>/chart[N], /chart[N], /<SheetName>/<relId>");
    }

    private static string RawSheetWithFilter(WorksheetPart worksheetPart, int? startRow, int? endRow, HashSet<string>? cols)
    {
        var worksheet = GetSheet(worksheetPart);
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
            return worksheet.OuterXml;

        var cloned = (Worksheet)worksheet.CloneNode(true);
        var clonedSheetData = cloned.GetFirstChild<SheetData>()!;
        clonedSheetData.RemoveAllChildren();

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowNum = (int)row.RowIndex!.Value;
            if (startRow.HasValue && rowNum < startRow.Value) continue;
            if (endRow.HasValue && rowNum > endRow.Value) break;

            if (cols != null)
            {
                var filteredRow = (Row)row.CloneNode(false);
                filteredRow.RowIndex = row.RowIndex;
                foreach (var cell in row.Elements<Cell>())
                {
                    var colName = ParseCellReference(cell.CellReference?.Value ?? "A1").Column;
                    if (cols.Contains(colName))
                        filteredRow.AppendChild(cell.CloneNode(true));
                }
                clonedSheetData.AppendChild(filteredRow);
            }
            else
            {
                clonedSheetData.AppendChild(row.CloneNode(true));
            }
        }

        return cloned.OuterXml;
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        OpenXmlPartRootElement rootElement;
        if (partPath is "/" or "/workbook")
        {
            rootElement = workbookPart.Workbook
                ?? throw new InvalidOperationException("No workbook");
        }
        else if (partPath == "/styles")
        {
            var styleManager = new ExcelStyleManager(workbookPart);
            rootElement = styleManager.EnsureStylesPart().Stylesheet!;
        }
        else if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                ?? throw new InvalidOperationException("No shared strings");
            rootElement = sst.SharedStringTable!;
        }
        else
        {
            // Drawing part: /SheetName/drawing
            var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
            if (drawingMatch.Success)
            {
                var drawSheetName = drawingMatch.Groups[1].Value;
                var drawWs = FindWorksheet(drawSheetName)
                    ?? throw SheetNotFoundException(drawSheetName);
                var dp = drawWs.DrawingsPart
                    ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
                rootElement = dp.WorksheetDrawing!;
            }
            else
            {
            // Chart part: /SheetName/chart[N] or /chart[N]
            var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
            if (chartMatch.Success)
            {
                var chartSheetName = chartMatch.Groups[1].Value;
                var chartIdx = int.Parse(chartMatch.Groups[2].Value);
                var chartWs = FindWorksheet(chartSheetName)
                    ?? throw SheetNotFoundException(chartSheetName);
                var chartPart = GetChartPart(chartWs, chartIdx);
                rootElement = chartPart.ChartSpace!;
            }
            else
            {
                var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
                if (globalChartMatch.Success)
                {
                    var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
                    var chartPart = GetGlobalChartPart(chartIdx);
                    rootElement = chartPart.ChartSpace!;
                }
                else
                {
                    // Try as sheet name
                    var sheetName = partPath.TrimStart('/');
                    var worksheet = FindWorksheet(sheetName)
                        ?? throw new ArgumentException($"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /<SheetName>, /<SheetName>/chart[N], /chart[N]");
                    rootElement = GetSheet(worksheet);
                }
            }
            }
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        Console.WriteLine($"raw-set: {affected} element(s) affected");
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose() => _doc.Dispose();

}
