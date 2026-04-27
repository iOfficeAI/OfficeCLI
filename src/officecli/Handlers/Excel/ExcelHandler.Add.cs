// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using SpreadsheetDrawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // Normalize to case-insensitive lookup so camelCase keys (e.g. minColor) match lowercase lookups
        if (properties != null && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        parentPath = NormalizeExcelPath(parentPath);
        parentPath = ResolveSheetIndexInPath(parentPath);
        switch (type.ToLowerInvariant())
        {
            case "sheet":
                return AddSheet(parentPath, type, position, properties);

            case "row":
                return AddRow(parentPath, type, position, properties);

            case "cell":
                return AddCell(parentPath, type, position, properties);

            case "namedrange" or "definedname":
                return AddNamedRange(parentPath, type, position, properties);

            case "comment" or "note":
                return AddComment(parentPath, type, position, properties);

            case "validation":
            case "datavalidation":
                return AddValidation(parentPath, type, position, properties);

            case "autofilter":
                return AddAutoFilter(parentPath, type, position, properties);

            case "cf":
                return AddCf(parentPath, type, position, properties);

            case "databar":
            case "conditionalformatting":
                return AddDataBar(parentPath, type, position, properties);

            case "colorscale":
                return AddColorScale(parentPath, type, position, properties);

            case "iconset":
                return AddIconSet(parentPath, type, position, properties);

            case "formulacf":
                return AddFormulaCf(parentPath, type, position, properties);

            case "cellis":
                return AddCellIs(parentPath, type, position, properties);

            case "ole":
            case "oleobject":
            case "object":
            case "embed":
                return AddOle(parentPath, type, position, properties);

            case "picture":
            case "image":
            case "img":
                return AddPicture(parentPath, type, position, properties);

            case "shape" or "textbox":
                return AddShape(parentPath, type, position, properties);

            case "table" or "listobject":
                return AddTable(parentPath, type, position, properties);

            case "chart":
                return AddChart(parentPath, type, position, properties);

            case "pivottable" or "pivot":
                return AddPivotTable(parentPath, type, position, properties);

            case "slicer":
                return AddSlicer(parentPath, type, position, properties);

            case "col" or "column":
                return AddCol(parentPath, type, position, properties);

            case "pagebreak":
                return AddPageBreak(parentPath, type, position, properties);

            case "rowbreak":
                return AddRowBreak(parentPath, type, position, properties);

            case "colbreak":
                return AddColBreak(parentPath, type, position, properties);

            case "run":
                return AddRun(parentPath, type, position, properties);

            case "topn":
            case "aboveaverage":
            case "uniquevalues":
            case "duplicatevalues":
            case "containstext":
            case "dateoccurring":
            case "cfextended":
                return AddCfExtended(parentPath, type, position, properties);

            case "sparkline":
                return AddSparkline(parentPath, type, position, properties);

            default:
                return AddDefault(parentPath, type, position, properties);
        }
    }

    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
        {
            // Move (reorder) the sheet within the workbook.
            // CONSISTENCY(move-anchor): mirrors PowerPointHandler.Move slide reorder —
            // supports --index / --after /Sheet2 / --before /Sheet3.
            var workbook = GetWorkbook();
            var sheets = workbook.GetFirstChild<Sheets>()
                ?? throw new InvalidOperationException("Workbook has no sheets element");
            var sheetEl = sheets.Elements<Sheet>().FirstOrDefault(s =>
                string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException($"Sheet not found: {sheetName}");

            // Resolve after/before anchor BEFORE removing sheetEl.
            static string ExtractAnchorSheetName(string raw) =>
                (raw.StartsWith("/") ? raw[1..] : raw).Split('/', 2)[0];

            Sheet? afterAnchor = null, beforeAnchor = null;
            if (position?.After != null)
            {
                var anchorName = ExtractAnchorSheetName(position.After);
                afterAnchor = sheets.Elements<Sheet>().FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, anchorName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"After anchor not found: {position.After}");
            }
            else if (position?.Before != null)
            {
                var anchorName = ExtractAnchorSheetName(position.Before);
                beforeAnchor = sheets.Elements<Sheet>().FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, anchorName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"Before anchor not found: {position.Before}");
            }
            else if (index == null)
            {
                throw new ArgumentException("One of --index, --after, or --before is required when moving a sheet");
            }

            sheetEl.Remove();

            if (afterAnchor != null)
            {
                afterAnchor.InsertAfterSelf(sheetEl);
            }
            else if (beforeAnchor != null)
            {
                beforeAnchor.InsertBeforeSelf(sheetEl);
            }
            else
            {
                var targetIndex = index!.Value;
                var sheetList = sheets.Elements<Sheet>().ToList();
                if (targetIndex >= 0 && targetIndex < sheetList.Count)
                    sheetList[targetIndex].InsertBeforeSelf(sheetEl);
                else
                    sheets.AppendChild(sheetEl);
            }
            workbook.Save();
            return $"/{sheetName}";
        }

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Determine target
        string effectiveParentPath;
        SheetData targetSheetData;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            effectiveParentPath = $"/{sheetName}";
            targetSheetData = sheetData;
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
            var tgtWorksheet = FindWorksheet(tgtSegments[0])
                ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
            targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
                ?? throw new ArgumentException("Target sheet has no data");
        }

        // Find and move the row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = int.Parse(rowMatch.Groups[1].Value);
            // Try ordinal lookup first (Nth row element), then fall back to RowIndex
            var allRows = sheetData.Elements<Row>().ToList();
            var row = (rowIdx >= 1 && rowIdx <= allRows.Count ? allRows[rowIdx - 1] : null)
                ?? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            row.Remove();

            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                    rows[index.Value].InsertBeforeSelf(row);
                else
                    targetSheetData.AppendChild(row);
            }
            else
            {
                targetSheetData.AppendChild(row);
            }

            SaveWorksheet(worksheet);
            var rowIndex = row.RowIndex?.Value ?? (uint)(targetSheetData.Elements<Row>().ToList().IndexOf(row) + 1);
            return $"{effectiveParentPath}/row[{rowIndex}]";
        }

        throw new ArgumentException($"Move not supported for: {elementRef}. Supported: row[N]");
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        // Parse both paths: /SheetName/row[N]
        var seg1 = path1.TrimStart('/').Split('/', 2);
        var seg2 = path2.TrimStart('/').Split('/', 2);
        if (seg1.Length < 2 || seg2.Length < 2)
            throw new ArgumentException("Swap requires element paths (e.g. /Sheet1/row[1])");
        if (seg1[0] != seg2[0])
            throw new ArgumentException("Cannot swap elements across different sheets");

        var sheetName = seg1[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        var rowMatch1 = Regex.Match(seg1[1], @"^row\[(\d+)\]$");
        var rowMatch2 = Regex.Match(seg2[1], @"^row\[(\d+)\]$");
        if (!rowMatch1.Success || !rowMatch2.Success)
            throw new ArgumentException("Swap only supports row[N] elements in Excel");

        var allRows = sheetData.Elements<Row>().ToList();
        var idx1 = int.Parse(rowMatch1.Groups[1].Value);
        var idx2 = int.Parse(rowMatch2.Groups[1].Value);
        var row1 = (idx1 >= 1 && idx1 <= allRows.Count ? allRows[idx1 - 1] : null)
            ?? throw new ArgumentException($"Row {idx1} not found");
        var row2 = (idx2 >= 1 && idx2 <= allRows.Count ? allRows[idx2 - 1] : null)
            ?? throw new ArgumentException($"Row {idx2} not found");

        // Swap RowIndex values and cell references
        var rowIndex1 = row1.RowIndex?.Value ?? (uint)idx1;
        var rowIndex2 = row2.RowIndex?.Value ?? (uint)idx2;
        row1.RowIndex = new DocumentFormat.OpenXml.UInt32Value(rowIndex2);
        row2.RowIndex = new DocumentFormat.OpenXml.UInt32Value(rowIndex1);

        // Update cell references (e.g. A1→A3, B1→B3)
        foreach (var cell in row1.Elements<Cell>())
        {
            if (cell.CellReference?.Value != null)
            {
                var colRef = Regex.Match(cell.CellReference.Value, @"^([A-Z]+)").Groups[1].Value;
                cell.CellReference = $"{colRef}{rowIndex2}";
            }
        }
        foreach (var cell in row2.Elements<Cell>())
        {
            if (cell.CellReference?.Value != null)
            {
                var colRef = Regex.Match(cell.CellReference.Value, @"^([A-Z]+)").Groups[1].Value;
                cell.CellReference = $"{colRef}{rowIndex1}";
            }
        }

        PowerPointHandler.SwapXmlElements(row1, row2);
        SaveWorksheet(worksheet);

        return ($"/{sheetName}/row[{rowIndex2}]", $"/{sheetName}/row[{rowIndex1}]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
            throw new ArgumentException("Cannot copy an entire sheet with --from. Use add --type sheet instead.");

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Find target
        var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
        var tgtWorksheet = FindWorksheet(tgtSegments[0])
            ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
        var targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Target sheet has no data");

        // Copy row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            var clone = (Row)row.CloneNode(true);

            // R8-1: CloneNode preserves the source row's RowIndex and every
            // cell's CellReference (e.g. "A1","B1"). Without rewriting these,
            // the new row collides with the source (Excel shows one row at
            // rowIdx, A2 appears empty) or is silently ignored. Compute the
            // new rowIndex from the target sheet and rewrite all cell refs.
            uint newRowIndex;
            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                {
                    newRowIndex = rows[index.Value].RowIndex?.Value ?? (uint)(index.Value + 1);
                    // Shift existing rows at/after this position down by 1
                    ShiftRowsDown(tgtWorksheet, (int)newRowIndex);
                    // Re-fetch sheetData (ShiftRowsDown may reorder)
                    targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()!;
                    var afterRow = targetSheetData.Elements<Row>()
                        .LastOrDefault(r => (r.RowIndex?.Value ?? 0) < newRowIndex);
                    if (afterRow != null) afterRow.InsertAfterSelf(clone);
                    else targetSheetData.InsertAt(clone, 0);
                }
                else
                {
                    newRowIndex = (targetSheetData.Elements<Row>()
                        .LastOrDefault()?.RowIndex?.Value ?? 0u) + 1;
                    targetSheetData.AppendChild(clone);
                }
            }
            else
            {
                newRowIndex = (targetSheetData.Elements<Row>()
                    .LastOrDefault()?.RowIndex?.Value ?? 0u) + 1;
                targetSheetData.AppendChild(clone);
            }

            clone.RowIndex = newRowIndex;
            foreach (var c in clone.Elements<Cell>())
            {
                var oldRef = c.CellReference?.Value;
                if (string.IsNullOrEmpty(oldRef)) continue;
                var m = Regex.Match(oldRef, @"^([A-Z]+)\d+$", RegexOptions.IgnoreCase);
                if (m.Success)
                    c.CellReference = $"{m.Groups[1].Value.ToUpperInvariant()}{newRowIndex}";
            }

            SaveWorksheet(tgtWorksheet);
            return $"{targetParentPath}/row[{newRowIndex}]";
        }

        throw new ArgumentException($"Copy not supported for: {elementRef}. Supported: row[N]");
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a worksheet's DrawingsPart
                var sheetName = parentPartPath.TrimStart('/');
                var worksheetPart = FindWorksheet(sheetName)
                    ?? throw new ArgumentException(
                        $"Sheet not found: {sheetName}. Chart must be added under a sheet: add-part <file> /<SheetName> --type chart");

                var drawingsPart = worksheetPart.DrawingsPart
                    ?? worksheetPart.AddNewPart<DrawingsPart>();

                // Initialize DrawingsPart if new
                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing =
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                    drawingsPart.WorksheetDrawing.Save();

                    // Link DrawingsPart to worksheet if not already linked
                    if (GetSheet(worksheetPart).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
                        GetSheet(worksheetPart).Append(
                            new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                        SaveWorksheet(worksheetPart);
                    }
                }

                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                var relId = drawingsPart.GetIdOfPart(chartPart);

                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = drawingsPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/{sheetName}/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }
}
