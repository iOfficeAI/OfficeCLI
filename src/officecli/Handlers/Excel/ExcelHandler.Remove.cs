// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string? Remove(string path)
    {
        path = NormalizeExcelPath(path);
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        // Handle /namedrange[N] or /namedrange[Name] before sheet lookup
        var namedRangeRemoveMatch = Regex.Match(sheetName, @"^namedrange\[(.+?)\]$", RegexOptions.IgnoreCase);
        if (namedRangeRemoveMatch.Success)
        {
            var selector = namedRangeRemoveMatch.Groups[1].Value;
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames == null)
                throw new ArgumentException("No named ranges found in workbook");

            var allDefs = definedNames.Elements<DefinedName>().ToList();
            DefinedName? dn = null;

            if (int.TryParse(selector, out var dnIndex))
            {
                if (dnIndex < 1 || dnIndex > allDefs.Count)
                    throw new ArgumentException($"Named range index {dnIndex} out of range (1-{allDefs.Count})");
                dn = allDefs[dnIndex - 1];
            }
            else
            {
                dn = allDefs.FirstOrDefault(d =>
                    d.Name?.Value?.Equals(selector, StringComparison.OrdinalIgnoreCase) == true);
                if (dn == null)
                    throw new ArgumentException($"Named range '{selector}' not found");
            }

            dn.Remove();
            if (!definedNames.HasChildren) definedNames.Remove();
            workbook.Save();
            return null;
        }

        if (segments.Length == 1)
        {
            // Remove entire sheet
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var sheets = GetWorkbook().GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
            if (sheet == null)
                throw SheetNotFoundException(sheetName);

            var sheetCount = sheets!.Elements<Sheet>().Count();
            if (sheetCount <= 1)
                throw new InvalidOperationException($"Cannot remove the last sheet. A workbook must contain at least one sheet.");

            var relId = sheet.Id?.Value;
            sheet.Remove();
            if (relId != null)
                workbookPart.DeletePart(workbookPart.GetPartById(relId));

            // Clean up named ranges referencing the deleted sheet
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                var toRemove = definedNames.Elements<DefinedName>()
                    .Where(dn => dn.Text?.Contains(sheetName + "!", StringComparison.OrdinalIgnoreCase) == true)
                    .ToList();
                foreach (var dn in toRemove) dn.Remove();
                if (!definedNames.HasChildren) definedNames.Remove();
            }

            // Fix ActiveTab to prevent workbook corruption when deleting the last tab
            var remainingCount = sheets!.Elements<Sheet>().Count();
            var bookViews = workbook.GetFirstChild<BookViews>();
            if (bookViews != null)
            {
                foreach (var bv in bookViews.Elements<WorkbookView>())
                {
                    if (bv.ActiveTab?.Value >= (uint)remainingCount)
                        bv.ActiveTab = (uint)Math.Max(0, remainingCount - 1);
                }
            }

            workbook.Save();
            return null;
        }

        var cellRef = segments[1];
        var worksheet = FindWorksheet(sheetName)
            ?? throw SheetNotFoundException(sheetName);
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // row[N] — true shift delete
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = int.Parse(rowMatch.Groups[1].Value);
            sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx)
                ?.Remove();
            var affected = CollectFormulaCellsAffectedByRowDelete(worksheet, rowIdx);
            ShiftRowsUp(worksheet, rowIdx);
            DeleteCalcChainIfPresent();
            SaveWorksheet(worksheet);
            return FormatFormulaWarning(affected);
        }

        // col[X] — true shift delete
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colName = colMatch.Groups[1].Value.ToUpperInvariant();
            var deletedColIdx = ColumnNameToIndex(colName);
            var affected = CollectFormulaCellsAffectedByColDelete(worksheet, deletedColIdx);
            ShiftColumnsLeft(worksheet, colName);
            DeleteCalcChainIfPresent();
            SaveWorksheet(worksheet);
            return FormatFormulaWarning(affected);
        }

        // sparkline[N] — remove sparkline group
        var sparklineRemoveMatch = Regex.Match(cellRef, @"^sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (sparklineRemoveMatch.Success)
        {
            var spkIdx = int.Parse(sparklineRemoveMatch.Groups[1].Value);
            var spkGroup = GetSparklineGroup(worksheet, spkIdx)
                ?? throw new ArgumentException($"Sparkline[{spkIdx}] not found in sheet '{sheetName}'");
            var spkGroups = spkGroup.Parent!;
            spkGroup.Remove();
            // If no more sparkline groups, clean up empty extension
            if (!spkGroups.HasChildren)
            {
                var spkExt = spkGroups.Parent;
                spkGroups.Remove();
                if (spkExt != null && !spkExt.HasChildren)
                {
                    var extList = spkExt.Parent;
                    spkExt.Remove();
                    if (extList != null && !extList.HasChildren)
                        extList.Remove();
                }
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // rowbreak[N] / colbreak[N]
        var rbRemoveMatch = Regex.Match(cellRef, @"^rowbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (rbRemoveMatch.Success)
        {
            var rbIdx = int.Parse(rbRemoveMatch.Groups[1].Value);
            var rowBreaks = GetSheet(worksheet).GetFirstChild<RowBreaks>();
            var breaks = rowBreaks?.Elements<Break>().ToList() ?? new();
            if (rbIdx >= 1 && rbIdx <= breaks.Count)
            {
                breaks[rbIdx - 1].Remove();
                if (rowBreaks != null)
                {
                    rowBreaks.Count = (uint)rowBreaks.Elements<Break>().Count();
                    rowBreaks.ManualBreakCount = rowBreaks.Count;
                    if (rowBreaks.Count == 0) rowBreaks.Remove();
                }
            }
            SaveWorksheet(worksheet);
            return null;
        }
        var cbRemoveMatch = Regex.Match(cellRef, @"^colbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cbRemoveMatch.Success)
        {
            var cbIdx = int.Parse(cbRemoveMatch.Groups[1].Value);
            var colBreaks = GetSheet(worksheet).GetFirstChild<ColumnBreaks>();
            var breaks = colBreaks?.Elements<Break>().ToList() ?? new();
            if (cbIdx >= 1 && cbIdx <= breaks.Count)
            {
                breaks[cbIdx - 1].Remove();
                if (colBreaks != null)
                {
                    colBreaks.Count = (uint)colBreaks.Elements<Break>().Count();
                    colBreaks.ManualBreakCount = colBreaks.Count;
                    if (colBreaks.Count == 0) colBreaks.Remove();
                }
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // shape[N] — remove shape anchor from DrawingsPart
        var shapeRemoveMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (shapeRemoveMatch.Success)
        {
            var shpIdx = int.Parse(shapeRemoveMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/shapes");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/shapes");
            var shpAnchors = wsDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
                .Where(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().Any())
                .ToList();
            if (shpIdx < 1 || shpIdx > shpAnchors.Count)
                throw new ArgumentException($"Shape index {shpIdx} out of range (1..{shpAnchors.Count})");
            shpAnchors[shpIdx - 1].Remove();
            wsDrawing.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // picture[N] — remove picture anchor from DrawingsPart
        var picRemoveMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (picRemoveMatch.Success)
        {
            var picIdx = int.Parse(picRemoveMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/pictures");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/pictures");
            var picAnchors = wsDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
                .Where(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().Any())
                .ToList();
            if (picIdx < 1 || picIdx > picAnchors.Count)
                throw new ArgumentException($"Picture index {picIdx} out of range (1..{picAnchors.Count})");
            // Remove associated image part to avoid storage bloat
            var pic = picAnchors[picIdx - 1].Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().First();
            var blipFill = pic.BlipFill?.Blip?.Embed?.Value;
            picAnchors[picIdx - 1].Remove();
            if (blipFill != null)
            {
                try { drawingsPart.DeletePart(drawingsPart.GetPartById(blipFill)); } catch { }
            }
            wsDrawing.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // chart[N] — remove chart anchor from DrawingsPart
        var chartRemoveMatch = Regex.Match(cellRef, @"^chart\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (chartRemoveMatch.Success)
        {
            var chartIdx = int.Parse(chartRemoveMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/charts");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/charts");
            var chartAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
                .Where(a => a.Descendants<C.ChartReference>().Any())
                .ToList();
            if (chartIdx < 1 || chartIdx > chartAnchors.Count)
                throw new ArgumentException($"Chart index {chartIdx} out of range (1..{chartAnchors.Count})");
            var anchor = chartAnchors[chartIdx - 1];
            var chartRef = anchor.Descendants<C.ChartReference>().First();
            var relId = chartRef.Id?.Value;
            anchor.Remove();
            if (relId != null)
            {
                try { drawingsPart.DeletePart(drawingsPart.GetPartById(relId)); } catch { }
            }
            wsDrawing.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // table[N] — remove table (ListObject) from worksheet
        var tableRemoveMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableRemoveMatch.Success)
        {
            var tblIdx = int.Parse(tableRemoveMatch.Groups[1].Value);
            var tableParts = worksheet.TableDefinitionParts.ToList();
            if (tblIdx < 1 || tblIdx > tableParts.Count)
                throw new ArgumentException($"Table index {tblIdx} out of range (1..{tableParts.Count})");
            var tablePart = tableParts[tblIdx - 1];
            worksheet.DeletePart(tablePart);
            // Also remove the tablePart reference from the TableParts element
            var tblParts = worksheet.Worksheet.GetFirstChild<TableParts>();
            if (tblParts != null)
            {
                var tblPartEntries = tblParts.Elements<TablePart>().ToList();
                if (tblIdx <= tblPartEntries.Count)
                    tblPartEntries[tblIdx - 1].Remove();
                tblParts.Count = (uint)tblParts.Elements<TablePart>().Count();
                if (tblParts.Count == 0)
                    tblParts.Remove();
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // comment[N] — remove comment from WorksheetCommentsPart
        var commentRemoveMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentRemoveMatch.Success)
        {
            var cmtIdx = int.Parse(commentRemoveMatch.Groups[1].Value);
            var commentsPart = worksheet.WorksheetCommentsPart;
            if (commentsPart?.Comments == null)
                throw new ArgumentException($"No comments found in sheet");
            var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
            var comments = cmtList?.Elements<Comment>().ToList() ?? new();
            if (cmtIdx < 1 || cmtIdx > comments.Count)
                throw new ArgumentException($"Comment index {cmtIdx} out of range (1..{comments.Count})");
            comments[cmtIdx - 1].Remove();
            if (cmtList != null && !cmtList.HasChildren)
            {
                worksheet.DeletePart(commentsPart);
                // Clean up VmlDrawingPart only if it contains no non-comment shapes (e.g. form controls)
                var vmlPart = worksheet.VmlDrawingParts.FirstOrDefault();
                if (vmlPart != null)
                {
                    bool hasNonCommentShapes = false;
                    try
                    {
                        using var stream = vmlPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
                        var vmlDoc = System.Xml.Linq.XDocument.Load(stream);
                        var vNs = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:vml";
                        var xNs = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:office:excel";
                        var shapes = vmlDoc.Descendants(vNs + "shape").ToList();
                        hasNonCommentShapes = shapes.Any(s =>
                        {
                            var clientData = s.Element(xNs + "ClientData");
                            return clientData == null ||
                                   clientData.Attribute("ObjectType")?.Value != "Note";
                        });
                    }
                    catch { }

                    if (!hasNonCommentShapes)
                    {
                        worksheet.DeletePart(vmlPart);
                        var legacyDrawing = GetSheet(worksheet).Elements<LegacyDrawing>().FirstOrDefault();
                        legacyDrawing?.Remove();
                    }
                    else
                    {
                        // Remove only comment shapes from VML, keep form controls
                        try
                        {
                            using var stream = vmlPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
                            var vmlDoc = System.Xml.Linq.XDocument.Load(stream);
                            var vNs2 = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:vml";
                            var xNs2 = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:office:excel";
                            var commentShapes = vmlDoc.Descendants(vNs2 + "shape")
                                .Where(s =>
                                {
                                    var cd = s.Element(xNs2 + "ClientData");
                                    return cd != null && cd.Attribute("ObjectType")?.Value == "Note";
                                }).ToList();
                            foreach (var cs in commentShapes) cs.Remove();
                            stream.SetLength(0);
                            vmlDoc.Save(stream);
                        }
                        catch { }
                    }
                }
            }
            else
            {
                commentsPart.Comments.Save();
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // validation[N] — remove data validation
        var validationRemoveMatch = Regex.Match(cellRef, @"^validation\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationRemoveMatch.Success)
        {
            var dvIdx = int.Parse(validationRemoveMatch.Groups[1].Value);
            var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>();
            if (dvs == null)
                throw new ArgumentException("No data validations found in sheet");
            var dvList = dvs.Elements<DataValidation>().ToList();
            if (dvIdx < 1 || dvIdx > dvList.Count)
                throw new ArgumentException($"Validation index {dvIdx} out of range (1..{dvList.Count})");
            dvList[dvIdx - 1].Remove();
            if (!dvs.HasChildren)
                dvs.Remove();
            else
                dvs.Count = (uint)dvs.Elements<DataValidation>().Count();
            SaveWorksheet(worksheet);
            return null;
        }

        // cf[N] — remove conditional formatting
        var cfRemoveMatch = Regex.Match(cellRef, @"^cf\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cfRemoveMatch.Success)
        {
            var cfIdx = int.Parse(cfRemoveMatch.Groups[1].Value);
            var ws = GetSheet(worksheet);
            var cfElements = ws.Elements<ConditionalFormatting>().ToList();
            if (cfIdx < 1 || cfIdx > cfElements.Count)
                throw new ArgumentException($"Conditional formatting index {cfIdx} out of range (1..{cfElements.Count})");
            cfElements[cfIdx - 1].Remove();
            SaveWorksheet(worksheet);
            return null;
        }

        // autofilter — remove AutoFilter from worksheet
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            var ws = GetSheet(worksheet);
            var autoFilter = ws.GetFirstChild<AutoFilter>();
            if (autoFilter != null)
            {
                autoFilter.Remove();
                SaveWorksheet(worksheet);
            }
            return null;
        }

        // run[N] — remove individual run from rich text cell
        var runRemoveMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runRemoveMatch.Success)
        {
            var runCellRef = runRemoveMatch.Groups[1].Value.ToUpperInvariant();
            var runIdx = int.Parse(runRemoveMatch.Groups[2].Value);

            var runCell = FindCell(sheetData, runCellRef)
                ?? throw new ArgumentException($"Cell {runCellRef} not found");

            if (runCell.DataType?.Value != CellValues.SharedString ||
                !int.TryParse(runCell.CellValue?.Text, out var sstIdx))
                throw new ArgumentException($"Cell {runCellRef} is not a rich text cell");

            var sstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi = sstPart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx);
            if (ssi == null) throw new ArgumentException($"SharedString entry {sstIdx} not found");

            var runs = ssi.Elements<Run>().ToList();
            if (runIdx < 1 || runIdx > runs.Count)
                throw new ArgumentException($"Run index {runIdx} out of range (1-{runs.Count})");

            runs[runIdx - 1].Remove();

            // Convert back to plain text if appropriate
            var remainingRuns = ssi.Elements<Run>().ToList();
            if (remainingRuns.Count == 0)
            {
                // All runs removed — set empty plain text to avoid orphaned SSI
                ssi.RemoveAllChildren<Text>();
                ssi.AppendChild(new Text("") { Space = SpaceProcessingModeValues.Preserve });
            }
            else if (remainingRuns.Count == 1)
            {
                var lastRun = remainingRuns[0];
                var rProps = lastRun.RunProperties;
                bool hasFormatting = rProps != null && rProps.HasChildren;
                if (!hasFormatting)
                {
                    var plainText = lastRun.GetFirstChild<Text>()?.Text ?? "";
                    lastRun.Remove();
                    ssi.RemoveAllChildren<Text>();
                    ssi.AppendChild(new Text(plainText) { Space = SpaceProcessingModeValues.Preserve });
                }
            }

            sstPart!.SharedStringTable!.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // Single cell
        var cell = FindCell(sheetData, cellRef)
            ?? throw new ArgumentException($"Cell {cellRef} not found");
        cell.Remove();
        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return null;
    }

    // ==================== Row/Column insert shift ====================

    /// <summary>
    /// Shift all rows >= insertRow down by 1 to make room for a new row insert.
    /// Mirrors ShiftRowsUp but in the opposite direction.
    /// </summary>
    internal void ShiftRowsDown(WorksheetPart worksheet, int insertRow)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();

        if (sheetData != null)
        {
            // Process in reverse order to avoid collision
            foreach (var row in sheetData.Elements<Row>().OrderByDescending(r => r.RowIndex?.Value ?? 0).ToList())
            {
                var rowIdx = (int)(row.RowIndex?.Value ?? 0);
                if (rowIdx < insertRow) continue;

                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value != null)
                    {
                        var (col, _) = ParseCellReference(cell.CellReference.Value);
                        cell.CellReference = $"{col}{rowIdx + 1}";
                    }
                }
                row.RowIndex = (uint)(rowIdx + 1);
            }
        }

        // Merge cells
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>().ToList())
            {
                var shifted = ShiftRowInRefDown(mc.Reference?.Value, insertRow);
                if (shifted != null) mc.Reference = shifted;
            }
        }

        // Conditional formatting sqref
        foreach (var cf in ws.Elements<ConditionalFormatting>().ToList())
        {
            if (cf.SequenceOfReferences?.HasValue != true) continue;
            var newRefs = cf.SequenceOfReferences.Items
                .Select(r => ShiftRowInRefDown(r.Value, insertRow) ?? r.Value).ToList();
            cf.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
        }

        // Data validations sqref
        var dvs = ws.GetFirstChild<DataValidations>();
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements<DataValidation>().ToList())
            {
                if (dv.SequenceOfReferences?.HasValue != true) continue;
                var newRefs = dv.SequenceOfReferences.Items
                    .Select(r => ShiftRowInRefDown(r.Value, insertRow) ?? r.Value).ToList();
                dv.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
            }
        }

        // AutoFilter
        var af = ws.GetFirstChild<AutoFilter>();
        if (af?.Reference?.Value != null)
        {
            var shifted = ShiftRowInRefDown(af.Reference.Value, insertRow);
            if (shifted != null) af.Reference = shifted;
        }

        // Named ranges
        ShiftNamedRangeRowsDown(worksheet, insertRow);
    }

    /// <summary>
    /// Shift all columns >= insertColIdx right by 1 to make room for a new column insert.
    /// </summary>
    internal void ShiftColumnsRight(WorksheetPart worksheet, int insertColIdx)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();

        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    if (cell.CellReference?.Value == null) continue;
                    var (col, rowIdx) = ParseCellReference(cell.CellReference.Value);
                    var colIdx = ColumnNameToIndex(col);
                    if (colIdx >= insertColIdx)
                        cell.CellReference = $"{IndexToColumnName(colIdx + 1)}{rowIdx}";
                }
            }
        }

        // Column width/style definitions
        var columns = ws.GetFirstChild<Columns>();
        if (columns != null)
        {
            foreach (var col in columns.Elements<Column>().OrderByDescending(c => c.Min?.Value ?? 0).ToList())
            {
                var min = (int)(col.Min?.Value ?? 0);
                var max = (int)(col.Max?.Value ?? 0);
                if (min >= insertColIdx)
                {
                    col.Min = (uint)(min + 1);
                    col.Max = (uint)(max + 1);
                }
                else if (max >= insertColIdx)
                {
                    col.Max = (uint)(max + 1);
                }
            }
        }

        // Merge cells
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>().ToList())
            {
                var shifted = ShiftColInRefRight(mc.Reference?.Value, insertColIdx);
                if (shifted != null) mc.Reference = shifted;
            }
        }

        // Named ranges
        ShiftNamedRangeColsRight(worksheet, insertColIdx);
    }

    private static string? ShiftRowInRefDown(string? refStr, int insertRow)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                shifted.Add(row >= insertRow ? $"{col}{row + 1}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    private static string? ShiftColInRefRight(string? refStr, int insertColIdx)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                var colIdx = ColumnNameToIndex(col);
                shifted.Add(colIdx >= insertColIdx ? $"{IndexToColumnName(colIdx + 1)}{row}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    private void ShiftNamedRangeRowsDown(WorksheetPart worksheet, int insertRow)
    {
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name;
        if (string.IsNullOrEmpty(sheetName)) return;
        var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
        if (definedNames == null) return;
        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (dn.Text == null) continue;
            dn.Text = Regex.Replace(dn.Text,
                $@"(?<={Regex.Escape(sheetName)}!\$?[A-Z]+\$?)(\d+)",
                m =>
                {
                    var row = int.Parse(m.Value);
                    return row >= insertRow ? (row + 1).ToString() : m.Value;
                },
                RegexOptions.IgnoreCase);
        }
        GetWorkbook().Save();
    }

    private void ShiftNamedRangeColsRight(WorksheetPart worksheet, int insertColIdx)
    {
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name;
        if (string.IsNullOrEmpty(sheetName)) return;
        var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
        if (definedNames == null) return;
        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (dn.Text == null) continue;
            dn.Text = Regex.Replace(dn.Text,
                $@"(?<={Regex.Escape(sheetName)}!)\$?([A-Z]+)\$?(\d+)",
                m =>
                {
                    var col = m.Groups[1].Value.ToUpperInvariant();
                    var row = m.Groups[2].Value;
                    var colIdx = ColumnNameToIndex(col);
                    if (colIdx < insertColIdx) return m.Value;
                    var dollar1 = m.Value.StartsWith("$") ? "$" : "";
                    var dollar2 = m.Value.Contains("$" + col + "$") ? "$" : "";
                    return $"{dollar1}{IndexToColumnName(colIdx + 1)}{dollar2}{row}";
                },
                RegexOptions.IgnoreCase);
        }
        GetWorkbook().Save();
    }

    // ==================== Row shift ====================

    private void ShiftRowsUp(WorksheetPart worksheet, int deletedRow)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();

        // 1. Shift all rows after the deleted row: update RowIndex + all CellReferences
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>().ToList())
            {
                var rowIdx = (int)(row.RowIndex?.Value ?? 0);
                if (rowIdx <= deletedRow) continue;

                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value != null)
                    {
                        var (col, _) = ParseCellReference(cell.CellReference.Value);
                        cell.CellReference = $"{col}{rowIdx - 1}";
                    }
                }
                row.RowIndex = (uint)(rowIdx - 1);
            }
        }

        // 2. Merge cells
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>().ToList())
            {
                var shifted = ShiftRowInRef(mc.Reference?.Value, deletedRow);
                if (shifted == null) mc.Remove();
                else mc.Reference = shifted;
            }
            if (!mergeCells.HasChildren) mergeCells.Remove();
        }

        // 3. Conditional formatting sqref
        foreach (var cf in ws.Elements<ConditionalFormatting>().ToList())
        {
            if (cf.SequenceOfReferences?.HasValue != true) continue;
            var newRefs = cf.SequenceOfReferences.Items
                .Select(r => ShiftRowInRef(r.Value, deletedRow))
                .OfType<string>().ToList();
            if (newRefs.Count == 0) cf.Remove();
            else cf.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
        }

        // 4. Data validations sqref
        var dvs = ws.GetFirstChild<DataValidations>();
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements<DataValidation>().ToList())
            {
                if (dv.SequenceOfReferences?.HasValue != true) continue;
                var newRefs = dv.SequenceOfReferences.Items
                    .Select(r => ShiftRowInRef(r.Value, deletedRow))
                    .OfType<string>().ToList();
                if (newRefs.Count == 0) dv.Remove();
                else dv.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
            }
            if (!dvs.HasChildren) dvs.Remove();
        }

        // 5. AutoFilter
        var af = ws.GetFirstChild<AutoFilter>();
        if (af?.Reference?.Value != null)
        {
            var shifted = ShiftRowInRef(af.Reference.Value, deletedRow);
            if (shifted != null) af.Reference = shifted;
            else af.Remove();
        }

        // 6. Named ranges (workbook-level)
        ShiftNamedRangeRows(worksheet, deletedRow);
    }

    // ==================== Column shift ====================

    private void ShiftColumnsLeft(WorksheetPart worksheet, string deletedColName)
    {
        var ws = GetSheet(worksheet);
        var deletedColIdx = ColumnNameToIndex(deletedColName);
        var sheetData = ws.GetFirstChild<SheetData>();

        // 1. Remove cells in deleted column, shift remaining cell references left
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    if (cell.CellReference?.Value == null) continue;
                    var (col, rowIdx) = ParseCellReference(cell.CellReference.Value);
                    var colIdx = ColumnNameToIndex(col);

                    if (colIdx == deletedColIdx)
                        cell.Remove();
                    else if (colIdx > deletedColIdx)
                        cell.CellReference = $"{IndexToColumnName(colIdx - 1)}{rowIdx}";
                }
            }
        }

        // 2. Column width/style definitions
        var columns = ws.GetFirstChild<Columns>();
        if (columns != null)
        {
            foreach (var col in columns.Elements<Column>().ToList())
            {
                var min = (int)(col.Min?.Value ?? 0);
                var max = (int)(col.Max?.Value ?? 0);

                if (min == deletedColIdx && max == deletedColIdx)
                {
                    col.Remove();
                }
                else if (min > deletedColIdx)
                {
                    col.Min = (uint)(min - 1);
                    col.Max = (uint)(max - 1);
                }
                else if (max >= deletedColIdx)
                {
                    col.Max = (uint)(max - 1);
                }
            }
            if (!columns.HasChildren) columns.Remove();
        }

        // 3. Merge cells
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>().ToList())
            {
                var shifted = ShiftColInRef(mc.Reference?.Value, deletedColIdx);
                if (shifted == null) mc.Remove();
                else mc.Reference = shifted;
            }
            if (!mergeCells.HasChildren) mergeCells.Remove();
        }

        // 4. Conditional formatting sqref
        foreach (var cf in ws.Elements<ConditionalFormatting>().ToList())
        {
            if (cf.SequenceOfReferences?.HasValue != true) continue;
            var newRefs = cf.SequenceOfReferences.Items
                .Select(r => ShiftColInRef(r.Value, deletedColIdx))
                .OfType<string>().ToList();
            if (newRefs.Count == 0) cf.Remove();
            else cf.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
        }

        // 5. Data validations sqref
        var dvs = ws.GetFirstChild<DataValidations>();
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements<DataValidation>().ToList())
            {
                if (dv.SequenceOfReferences?.HasValue != true) continue;
                var newRefs = dv.SequenceOfReferences.Items
                    .Select(r => ShiftColInRef(r.Value, deletedColIdx))
                    .OfType<string>().ToList();
                if (newRefs.Count == 0) dv.Remove();
                else dv.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
            }
            if (!dvs.HasChildren) dvs.Remove();
        }

        // 6. AutoFilter
        var af = ws.GetFirstChild<AutoFilter>();
        if (af?.Reference?.Value != null)
        {
            var shifted = ShiftColInRef(af.Reference.Value, deletedColIdx);
            if (shifted != null) af.Reference = shifted;
            else af.Remove();
        }

        // 7. Named ranges
        ShiftNamedRangeCols(worksheet, deletedColIdx);
    }

    // ==================== Shift helpers ====================

    /// <summary>
    /// Shift row numbers in a cell/range reference after a row deletion.
    /// Returns null if the reference sits exactly on the deleted row (should be removed).
    /// For ranges: if either endpoint is on the deleted row the range is removed;
    /// endpoints after the deleted row are decremented by 1.
    /// </summary>
    private static string? ShiftRowInRef(string? refStr, int deletedRow)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                if (row == deletedRow) return null;
                shifted.Add(row > deletedRow ? $"{col}{row - 1}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    /// <summary>
    /// Shift column letters in a cell/range reference after a column deletion.
    /// Returns null if the reference sits exactly on the deleted column.
    /// </summary>
    private static string? ShiftColInRef(string? refStr, int deletedColIdx)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                var colIdx = ColumnNameToIndex(col);
                if (colIdx == deletedColIdx) return null;
                shifted.Add(colIdx > deletedColIdx ? $"{IndexToColumnName(colIdx - 1)}{row}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    /// <summary>
    /// Update workbook-level named ranges after a row deletion.
    /// Handles both relative (A1) and absolute ($A$1) references.
    /// Row numbers in named range formula text are shifted via regex replacement.
    /// </summary>
    private void ShiftNamedRangeRows(WorksheetPart worksheet, int deletedRow)
    {
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name;
        if (string.IsNullOrEmpty(sheetName)) return;

        var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
        if (definedNames == null) return;

        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (dn.Text == null) continue;
            dn.Text = ShiftRowNumbersInText(dn.Text, sheetName, deletedRow);
        }
        GetWorkbook().Save();
    }

    /// <summary>
    /// Update workbook-level named ranges after a column deletion.
    /// </summary>
    private void ShiftNamedRangeCols(WorksheetPart worksheet, int deletedColIdx)
    {
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name;
        if (string.IsNullOrEmpty(sheetName)) return;

        var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
        if (definedNames == null) return;

        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (dn.Text == null) continue;
            dn.Text = ShiftColLettersInText(dn.Text, sheetName, deletedColIdx);
        }
        GetWorkbook().Save();
    }

    // ==================== Formula impact detection ====================

    private record FormulaImpact(string CellRef, bool IsRefError);

    /// <summary>
    /// Find all surviving cells with formulas that reference the deleted row (→ #REF!) or rows after it (→ shifted).
    /// </summary>
    private List<FormulaImpact> CollectFormulaCellsAffectedByRowDelete(WorksheetPart worksheet, int deletedRow)
    {
        var affected = new List<FormulaImpact>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return affected;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var formula = cell.CellFormula?.Text;
                if (string.IsNullOrEmpty(formula)) continue;

                bool refError = FormulaReferencesExactRow(formula, deletedRow);
                bool shifted = !refError && FormulaReferencesRowAbove(formula, deletedRow);

                if (refError || shifted)
                    affected.Add(new FormulaImpact(cell.CellReference?.Value ?? "?", refError));
            }
        }
        return affected;
    }

    private static bool FormulaReferencesExactRow(string formula, int row)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?[A-Z]+\$?(\d+)", RegexOptions.IgnoreCase))
        {
            if (int.TryParse(m.Groups[1].Value, out var r) && r == row)
                return true;
        }
        return false;
    }

    private static bool FormulaReferencesRowAbove(string formula, int deletedRow)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?[A-Z]+\$?(\d+)", RegexOptions.IgnoreCase))
        {
            if (int.TryParse(m.Groups[1].Value, out var row) && row > deletedRow)
                return true;
        }
        return false;
    }

    /// <summary>
    /// Find all surviving cells with formulas that reference the deleted column (→ #REF!) or columns after it (→ shifted).
    /// </summary>
    private List<FormulaImpact> CollectFormulaCellsAffectedByColDelete(WorksheetPart worksheet, int deletedColIdx)
    {
        var affected = new List<FormulaImpact>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return affected;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var formula = cell.CellFormula?.Text;
                if (string.IsNullOrEmpty(formula)) continue;

                bool refError = FormulaReferencesExactCol(formula, deletedColIdx);
                bool shifted = !refError && FormulaReferencesColAbove(formula, deletedColIdx);

                if (refError || shifted)
                    affected.Add(new FormulaImpact(cell.CellReference?.Value ?? "?", refError));
            }
        }
        return affected;
    }

    private static bool FormulaReferencesExactCol(string formula, int colIdx)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?([A-Z]+)\$?\d+", RegexOptions.IgnoreCase))
        {
            if (ColumnNameToIndex(m.Groups[1].Value.ToUpperInvariant()) == colIdx)
                return true;
        }
        return false;
    }

    private static bool FormulaReferencesColAbove(string formula, int deletedColIdx)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?([A-Z]+)\$?\d+", RegexOptions.IgnoreCase))
        {
            if (ColumnNameToIndex(m.Groups[1].Value.ToUpperInvariant()) > deletedColIdx)
                return true;
        }
        return false;
    }

    private static string? FormatFormulaWarning(List<FormulaImpact> affected)
    {
        if (affected.Count == 0) return null;

        var refErrors = affected.Where(a => a.IsRefError).Select(a => a.CellRef).ToList();
        var shifted = affected.Where(a => !a.IsRefError).Select(a => a.CellRef).ToList();

        var parts = new List<string>();
        if (refErrors.Count > 0)
            parts.Add($"{refErrors.Count} cell(s) will become #REF!: {string.Join(", ", refErrors)}");
        if (shifted.Count > 0)
            parts.Add($"{shifted.Count} cell(s) reference shifted rows/cols (formula text unchanged): {string.Join(", ", shifted)}");

        return $"Warning: {affected.Count} formula cell(s) affected — {string.Join("; ", parts)}";
    }

    /// <summary>
    /// In a formula/reference string like "Sheet1!$A$3:$B$5", decrement row numbers > deletedRow.
    /// Only touches references that belong to the given sheet.
    /// </summary>
    private static string ShiftRowNumbersInText(string text, string sheetName, int deletedRow)
    {
        // Match: optional sheet prefix (Sheet1! or 'Sheet 1'!), optional $, column letters, optional $, row number
        return Regex.Replace(text,
            $@"(?<={Regex.Escape(sheetName)}!\$?[A-Z]+\$?)(\d+)",
            m =>
            {
                var row = int.Parse(m.Value);
                return row > deletedRow ? (row - 1).ToString() : m.Value;
            },
            RegexOptions.IgnoreCase);
    }

    /// <summary>
    /// In a formula/reference string like "Sheet1!$B$1:$D$5", shift column letters > deletedColIdx left by one.
    /// Only touches references that belong to the given sheet.
    /// </summary>
    private static string ShiftColLettersInText(string text, string sheetName, int deletedColIdx)
    {
        return Regex.Replace(text,
            $@"(?<={Regex.Escape(sheetName)}!)\$?([A-Z]+)\$?(\d+)",
            m =>
            {
                var col = m.Groups[1].Value.ToUpperInvariant();
                var row = m.Groups[2].Value;
                var colIdx = ColumnNameToIndex(col);
                if (colIdx <= deletedColIdx) return m.Value;
                var dollar1 = m.Value.StartsWith("$") ? "$" : "";
                var dollar2 = m.Value.Contains("$" + col + "$") ? "$" : "";
                return $"{dollar1}{IndexToColumnName(colIdx - 1)}{dollar2}{row}";
            },
            RegexOptions.IgnoreCase);
    }
}
