// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using ThreadedCmt = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Column Set (width, hidden) ====================

    private List<string> SetColumn(WorksheetPart worksheet, string colName, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);
        var colIdx = (uint)ColumnNameToIndex(colName);
        // Excel's column space tops out at XFD (16384). Anything past that
        // can't render in Excel; reject at the handler boundary so callers
        // get an invalid_value rather than a quietly-corrupt OOXML file.
        if (colIdx < 1 || colIdx > 16384)
            throw new ArgumentException(
                $"Invalid column '{colName}'. Column index {colIdx} is out of range; valid range is A-XFD (1-16384).");

        var columns = ws.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData != null)
                ws.InsertBefore(columns, sheetData);
            else
                ws.AppendChild(columns);
        }

        // Find existing column definition or create one
        var col = columns.Elements<Column>()
            .FirstOrDefault(c => c.Min?.Value <= colIdx && c.Max?.Value >= colIdx);
        // A multi-column <col min max> range (e.g. a shared width for A:C)
        // must be split before mutating, otherwise a set targeting one column
        // bleeds the property onto every column in the range.
        if (col != null && (col.Min!.Value < colIdx || col.Max!.Value > colIdx))
        {
            var rangeMin = col.Min.Value;
            var rangeMax = col.Max!.Value;
            if (rangeMin < colIdx)
            {
                var left = (Column)col.CloneNode(true);
                left.Min = rangeMin;
                left.Max = colIdx - 1;
                col.InsertBeforeSelf(left);
            }
            if (rangeMax > colIdx)
            {
                var right = (Column)col.CloneNode(true);
                right.Min = colIdx + 1;
                right.Max = rangeMax;
                col.InsertAfterSelf(right);
            }
            col.Min = colIdx;
            col.Max = colIdx;
        }
        if (col == null)
        {
            // Leave width/customWidth unset on implicit creation — `add column`
            // produces a bare <col> and stamping the 8.43 default here made a
            // set-only column (hidden=, outline=) round-trip non-idempotent:
            // the replayed file gained width="8.43" customWidth="1" the source
            // never had. The width case below stamps both when actually set.
            col = new Column { Min = colIdx, Max = colIdx };
            var afterCol = columns.Elements<Column>().LastOrDefault(c => (c.Min?.Value ?? 0) < colIdx);
            if (afterCol != null)
                afterCol.InsertAfterSelf(col);
            else
                columns.PrependChild(col);
        }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "width":
                    col.Width = ParseColWidthChars(value);
                    col.CustomWidth = true;
                    break;
                case "hidden":
                    col.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "outline" or "outlinelevel" or "group":
                    // DEFERRED(xlsx/row-height-validation) RC2: Excel outline level max is 7.
                    if (!byte.TryParse(value, out var colOutline) || colOutline > 7)
                        throw new ArgumentException($"Invalid 'outline' value: '{value}'. Expected an integer 0-7 (outline/group level).");
                    col.OutlineLevel = colOutline;
                    break;
                case "collapsed":
                    col.Collapsed = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "autofit":
                    if (ParseHelpers.IsTruthy(value))
                    {
                        var autoFitWidth = CalculateAutoFitWidth(worksheet, colName);
                        col.Width = autoFitWidth;
                        col.CustomWidth = true;
                    }
                    break;
                case "numfmt" or "numberformat" or "format":
                {
                    // A column number format is applied via a style reference
                    // (<col s="N">), NOT a raw attribute — CT_Col has no numFmt
                    // attribute, so writing one (the old default-case fallback)
                    // produced schema-invalid XML that failed `validate`.
                    // Register the format in the stylesheet and point the
                    // column's style at the resulting cellXf.
                    var colWorkbookPart = _doc.WorkbookPart
                        ?? throw new InvalidOperationException("Workbook not found");
                    var colStyleManager = new ExcelStyleManager(colWorkbookPart);
                    var tempCell = new Cell { StyleIndex = col.Style };
                    col.Style = colStyleManager.ApplyStyle(
                        tempCell,
                        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["numberformat"] = value },
                        unsupported);
                    _dirtyStylesheet = true;
                    break;
                }
                default:
                    // Long-tail Column attribute (CT_Col attrs beyond width/
                    // hidden/outlineLevel/collapsed/customWidth — e.g. style,
                    // bestFit, phonetic). Set as raw OOXML attribute. Symmetric
                    // with the column Get reader which now uses
                    // FillUnknownAttrProps for unrecognized attrs. Preserve
                    // original case (OOXML attribute names are case-sensitive).
                    col.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", key, "", value));
                    break;
            }
        }

        SaveWorksheet(worksheet);
        return unsupported;
    }

    // ==================== Column Auto-Fit ====================

    private double CalculateAutoFitWidth(WorksheetPart worksheet, string colName)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        var colIdx = ColumnNameToIndex(colName);
        double maxLen = 0;

        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value;
                    if (cellRef == null) continue;
                    var (cellCol, _) = ParseCellReference(cellRef);
                    if (ColumnNameToIndex(cellCol) != colIdx) continue;

                    var text = GetCellDisplayValue(cell);
                    var textWidth = ParseHelpers.EstimateTextWidthInChars(text);
                    if (textWidth > maxLen)
                        maxLen = textWidth;
                }
            }
        }

        // Approximate width: characters * 1.1 + 2 for padding, minimum 8
        return Math.Max(maxLen * 1.1 + 2, 8);
    }

    private void AutoFitAllColumns(WorksheetPart worksheet)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        // Collect all used column indices
        var usedColumns = new HashSet<int>();
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;
                var (cellCol, _) = ParseCellReference(cellRef);
                usedColumns.Add(ColumnNameToIndex(cellCol));
            }
        }

        if (usedColumns.Count == 0) return;

        var columns = ws.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            ws.InsertBefore(columns, sheetData);
        }

        foreach (var colIdx in usedColumns.OrderBy(c => c))
        {
            var colName = IndexToColumnName(colIdx);
            var width = CalculateAutoFitWidth(worksheet, colName);
            var uColIdx = (uint)colIdx;

            var col = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min?.Value <= uColIdx && c.Max?.Value >= uColIdx);
            if (col == null)
            {
                col = new Column { Min = uColIdx, Max = uColIdx, Width = width, CustomWidth = true };
                var afterCol = columns.Elements<Column>().LastOrDefault(c => (c.Min?.Value ?? 0) < uColIdx);
                if (afterCol != null)
                    afterCol.InsertAfterSelf(col);
                else
                    columns.PrependChild(col);
            }
            else
            {
                col.Width = width;
                col.CustomWidth = true;
            }
        }

        SaveWorksheet(worksheet);
    }

    /// <summary>
    /// Heuristic row height (no rendering engine): per cell, the wrapped line
    /// count is ceil(display width / effective column width) summed over the
    /// cell's explicit newline segments; the row height is max-lines × 15pt
    /// line pitch (the 11pt Calibri default), clamped to Excel's 15–409pt
    /// range. Run AFTER column autofit — the wrap estimate depends on the
    /// final column widths. Excel's own autofit stays available in-app for
    /// font-exact sizing.
    /// </summary>
    private double CalculateAutoFitHeight(WorksheetPart worksheet, int rowIdx)
    {
        const double LinePitchPt = 15.0;
        const double DefaultColWidthChars = 8.43;

        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        var row = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null) return LinePitchPt;

        double EffectiveWidth(int colIdx)
        {
            var cols = ws.GetFirstChild<Columns>();
            var col = cols?.Elements<Column>()
                .FirstOrDefault(c => c.Min?.Value <= colIdx && c.Max?.Value >= colIdx);
            return col?.Width?.Value is double w && w > 0 ? w : DefaultColWidthChars;
        }

        var maxLines = 1;
        foreach (var cell in row.Elements<Cell>())
        {
            var cellRef = cell.CellReference?.Value;
            if (cellRef == null) continue;
            var (cellCol, _) = ParseCellReference(cellRef);
            var text = GetCellDisplayValue(cell);
            if (string.IsNullOrEmpty(text)) continue;
            var widthChars = EffectiveWidth(ColumnNameToIndex(cellCol));
            int lines = 0;
            foreach (var segment in text.Split('\n'))
            {
                var segWidth = ParseHelpers.EstimateTextWidthInChars(segment);
                lines += Math.Max(1, (int)Math.Ceiling(segWidth / widthChars));
            }
            if (lines > maxLines) maxLines = lines;
        }

        return Math.Clamp(maxLines * LinePitchPt, 15, 409);
    }

    private void AutoFitAllRows(WorksheetPart worksheet)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        foreach (var row in sheetData.Elements<Row>())
        {
            if (row.RowIndex == null || !row.Elements<Cell>().Any()) continue;
            row.Height = CalculateAutoFitHeight(worksheet, (int)row.RowIndex.Value);
            row.CustomHeight = true;
        }
        SaveWorksheet(worksheet);
    }

    // ==================== Row Set (height, hidden) ====================

    // Set a manual page break's position / span / manual flag, keeping the
    // parent's manualBreakCount in sync. Page breaks are otherwise add/remove
    // markers; this gives the queried /Sheet/{row,col}break[N] a Set target so
    // query→set round-trips (e.g. reposition via row=N / col=N).
    private List<string> SetPageBreak(WorksheetPart worksheet, bool isRow, int index, Dictionary<string, string> properties)
    {
        var ws = GetSheet(worksheet);
        var rowBreaks = isRow ? ws.GetFirstChild<RowBreaks>() : null;
        var colBreaks = isRow ? null : ws.GetFirstChild<ColumnBreaks>();
        var breaks = (isRow ? rowBreaks?.Elements<Break>() : colBreaks?.Elements<Break>())?.ToList()
            ?? new List<Break>();
        var kind = isRow ? "rowbreak" : "colbreak";
        if (index < 1 || index > breaks.Count)
            throw new ArgumentException($"{kind}[{index}] not found (sheet has {breaks.Count} {(isRow ? "row" : "column")} break(s)).");

        var brk = breaks[index - 1];
        var unsupported = new List<string>();
        bool changed = false;
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "row" when isRow:
                case "col" when !isRow:
                case "index":
                    brk.Id = uint.Parse(value); changed = true; break;
                case "manual":
                    brk.ManualPageBreak = ParseHelpers.IsTruthySafe(value); changed = true; break;
                case "max": brk.Max = uint.Parse(value); changed = true; break;
                case "min": brk.Min = uint.Parse(value); changed = true; break;
                default: unsupported.Add(key); break;
            }
        }

        if (changed)
        {
            var manCount = (uint)breaks.Count(b => b.ManualPageBreak?.Value == true);
            if (isRow) rowBreaks!.ManualBreakCount = manCount;
            else colBreaks!.ManualBreakCount = manCount;
            SaveWorksheet(worksheet);
        }
        return unsupported;
    }

    // Genuine CT_Row long-tail attributes that the row Get reader round-trips.
    // Anything outside height/hidden/outlineLevel/collapsed (handled explicitly)
    // and this set is rejected rather than silently written, so a typo or a
    // column-name that binds to no table surfaces as unsupported_property.
    private static bool IsLongTailRowAttribute(string key) => key.ToLowerInvariant() switch
    {
        "spans" or "style" or "s" or "customformat" or "ph"
            or "thicktop" or "thickbot" or "customheight" => true,
        _ => false,
    };

    // Every key SetRow interprets as a row property — used by the
    // column-shadow collision check: a bare key in this set that ALSO resolves
    // as a table column is ambiguous. Single source: the query-side
    // RowAttributeKeys (Get-emitted props) plus the Set-only input aliases and
    // the long-tail CT_Row passthrough attrs. The set is deliberately WIDER
    // than the query side's: query can only filter on keys Get emits, so a
    // bare `row[style=…]` is unambiguous there (must be a column), while Set
    // can write style/group/… and therefore must treat them as colliding.
    private static bool IsRowSetAttributeKey(string key)
        => RowAttributeKeys.Contains(key)
            || key.ToLowerInvariant() is "rowheight" or "outline" or "group"
            || IsLongTailRowAttribute(key);

    private List<string> SetRow(WorksheetPart worksheet, uint rowIdx, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        // Excel's row space tops out at 1048576 (2^20). Reject anything past
        // that at the handler boundary so callers get an invalid_value
        // rather than a quietly-corrupt OOXML file.
        if (rowIdx < 1 || rowIdx > 1048576)
            throw new ArgumentException(
                $"Invalid row index {rowIdx}. Valid row range is 1-1048576.");
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
            throw new ArgumentException("Sheet has no data");

        // Use the shared find-or-create helper (same one FindOrCreateCell uses)
        // so the _rowIndex cache stays consistent and both paths share one <Row>
        // element — avoids duplicate <row r="N"> when set row[N] is followed by
        // set cell-in-row-N.
        var row = FindOrCreateRow(sheetData, rowIdx);

        foreach (var (rawKey, value) in properties)
        {
            // CONSISTENCY(row-col-collision): same escape vocabulary as the
            // query side (`row[@height…]` / `row[col.Salary…]`) — `@key`
            // forces the ROW PROPERTY, `col.key` forces the table COLUMN. A
            // bare key that names BOTH (a table column literally called
            // "Height"/"Group"/"Style" over this row) is ambiguous and throws
            // instead of silently picking one: before this check the switch
            // cases below always won for height/hidden/outline/group/
            // collapsed (so `--prop Height=180` re-sized the row instead of
            // writing the cell), while long-tail attrs silently lost to the
            // column — both directions were silent wrong-writes.
            bool forcedAttr = rawKey.StartsWith("@", StringComparison.Ordinal);
            var key = forcedAttr ? rawKey[1..] : rawKey;
            bool forcedCol = !forcedAttr
                && System.Text.RegularExpressions.Regex.IsMatch(key, @"^col(?:umn)?\.", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (!forcedAttr && !forcedCol && IsRowSetAttributeKey(key)
                && TryResolveRowColumnCell(worksheet, rowIdx, key, out _))
                throw new Core.CliException(
                    $"set row[{rowIdx}] --prop {key}=… is ambiguous: '{key}' names both a row property and a table column. " +
                    $"Use --prop col.{key}=… to write the column's cell, or --prop @{key}=… for the row property.")
                {
                    Code = "invalid_selector",
                    Suggestion = $"col.{key}={value} (table column) or @{key}={value} (row property)",
                    ValidValues = new[] { $"col.{key}", $"@{key}" },
                };

            switch (forcedCol ? null : key.ToLowerInvariant())
            {
                case "height" or "rowheight":
                    row.Height = ParseRowHeightPoints(value);
                    row.CustomHeight = true;
                    break;
                case "hidden":
                    row.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "outline" or "outlinelevel" or "group":
                    // DEFERRED(xlsx/row-height-validation) RC2: Excel outline level max is 7.
                    if (!byte.TryParse(value, out var outlineVal) || outlineVal > 7)
                        throw new ArgumentException($"Invalid 'outline' value: '{value}'. Expected an integer 0-7 (outline/group level).");
                    row.OutlineLevel = outlineVal;
                    break;
                case "collapsed":
                    row.Collapsed = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "autofit":
                    if (ParseHelpers.IsTruthy(value))
                    {
                        row.Height = CalculateAutoFitHeight(worksheet, (int)rowIdx);
                        row.CustomHeight = true;
                    }
                    break;
                default:
                    // A non-row-attribute key. First try it as a TABLE COLUMN on
                    // this row — `set /Sheet/row[N] --prop <colName>=<val>` writes
                    // the cell in that column, closing the query→edit loop
                    // symmetrically with `query row[col op val]`. Only fall back
                    // to a raw <row> XML attribute for genuine long-tail CT_Row
                    // attrs (spans/style/ph/thickTop/thickBot/customFormat), which
                    // the row Get reader round-trips. An unknown key that is
                    // neither is reported as unsupported — NEVER silently written
                    // (the old bug stamped `年龄="99"` onto <row> and reported
                    // success while the data never changed).
                    if (!forcedAttr && TryResolveRowColumnCell(worksheet, rowIdx, key, out var colCellRef))
                    {
                        var colCell = FindOrCreateCell(sheetData, colCellRef);
                        unsupported.AddRange(ApplyCellProperties(
                            colCell, colCellRef, worksheet, new() { ["value"] = value }));
                        PruneEmptyCell(colCell);
                    }
                    else if (!forcedCol && IsLongTailRowAttribute(key))
                    {
                        row.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", key, "", value));
                    }
                    else
                    {
                        // Point at the real headers (薪水 → 薪资) when a table
                        // covers this row; fall back to the bare key otherwise.
                        unsupported.Add(DescribeRowColumnsHint(worksheet, rowIdx, key) ?? key);
                    }
                    break;
            }
        }

        SaveWorksheet(worksheet);
        return unsupported;
    }

    // ==================== AutoFilter Set ====================

    private List<string> SetAutoFilter(WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "range":
                {
                    var autoFilter = ws.GetFirstChild<AutoFilter>();
                    if (autoFilter == null)
                    {
                        autoFilter = new AutoFilter();
                        // AutoFilter goes after SheetData (after MergeCells if present)
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        var sheetData = ws.GetFirstChild<SheetData>();
                        if (mergeCells != null)
                            mergeCells.InsertAfterSelf(autoFilter);
                        else if (sheetData != null)
                            sheetData.InsertAfterSelf(autoFilter);
                        else
                            ws.AppendChild(autoFilter);
                    }
                    autoFilter.Reference = value.ToUpperInvariant();
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        SaveWorksheet(worksheet);
        return unsupported;
    }
}
