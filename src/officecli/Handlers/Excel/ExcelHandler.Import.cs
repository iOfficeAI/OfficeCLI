// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Import CSV/TSV data into a worksheet starting at the given cell.
    /// </summary>
    /// <param name="parentPath">Sheet path, e.g. "/Sheet1"</param>
    /// <param name="csvContent">Raw CSV/TSV string content</param>
    /// <param name="delimiter">Field delimiter: ',' for CSV, '\t' for TSV</param>
    /// <param name="hasHeader">If true, set AutoFilter, freeze pane and print titles on first row</param>
    /// <param name="startCell">Starting cell reference, e.g. "A1"</param>
    /// <param name="decimalSeparator">Decimal mark the SOURCE uses: '.' (default)
    /// or ',' for the de-DE / ru-RU spelling, where '.' becomes the thousands
    /// group. Declared by the caller — never detected (see Core.NumericText).</param>
    /// <returns>Summary of rows/cols imported</returns>
    public string Import(string parentPath, string csvContent, char delimiter, bool hasHeader, string startCell,
        char decimalSeparator = '.')
    {
        parentPath = NormalizeExcelPath(parentPath);
        parentPath = ResolveSheetIndexInPath(parentPath);
        var pathSegments = parentPath.TrimStart('/').Split('/', 2);
        var sheetName = pathSegments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        // A cell-qualified target (/Sheet1/D3) is the same address every other
        // xlsx verb takes, so it names the top-left landing cell. It used to be
        // cut off after the sheet segment without a word, and the matrix landed
        // on A1 over whatever was there. Anything that is not a single cell —
        // a range, a row/col locator, a typo — is refused rather than guessed.
        if (pathSegments.Length > 1 && pathSegments[1].Length > 0)
        {
            var cellSeg = pathSegments[1].Replace("$", "").ToUpperInvariant();
            if (!System.Text.RegularExpressions.Regex.IsMatch(cellSeg, @"^[A-Z]{1,3}[1-9]\d*$"))
                throw new Core.CliException(
                    $"import target '{parentPath}' must be a sheet (/{sheetName}) or a single cell (/{sheetName}/A1); '{pathSegments[1]}' is neither.")
                { Code = "invalid_path", Suggestion = $"Use /{sheetName}/{{top-left cell}} or --start-cell." };
            if (!string.Equals(startCell, "A1", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(startCell, cellSeg, StringComparison.OrdinalIgnoreCase))
                throw new Core.CliException(
                    $"import target '{parentPath}' names cell {cellSeg} but --start-cell says {startCell}; pass one or make them agree.")
                { Code = "invalid_argument" };
            startCell = cellSeg;
        }

        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>()
            ?? ws.AppendChild(new SheetData());

        // Parse start cell
        var (startCol, startRow) = ParseCellReference(startCell.ToUpperInvariant());
        var startColIdx = ColumnNameToIndex(startCol);

        // Excel's `sep=X` first line is a declaration, not a row. Strip it
        // whatever the caller chose as the delimiter: left in place it becomes a
        // junk first row that shifts every real row down one, so hasHeader
        // freezes and auto-filters the wrong line. Which separator to USE is the
        // caller's call (an explicit --delimiter still wins) — see
        // Core.CsvSepDeclaration.
        if (Core.CsvSepDeclaration.TryRead(csvContent, out _, out var withoutSepLine))
            csvContent = withoutSepLine;

        // Parse CSV
        var rows = ParseCsv(csvContent, delimiter);
        if (rows.Count == 0)
            return "No data to import";

        int maxCols = 0;
        for (int r = 0; r < rows.Count; r++)
            if (rows[r].Count > maxCols) maxCols = rows[r].Count;

        // DOS-hardening: reject imports that exceed Excel's sheet dimensions
        // BEFORE writing anything. Without this an over-sized CSV (e.g. >XFD
        // columns or >1048576 rows) spun indefinitely instead of erroring.
        const int ExcelMaxRow = 1048576;
        const int ExcelMaxCol = 16384; // XFD (ColumnNameToIndex is 1-based)
        long endRowReq = (long)startRow + rows.Count - 1;
        if (endRowReq > ExcelMaxRow)
            throw new ArgumentException(
                $"Import exceeds Excel's row limit: data would reach row {endRowReq} " +
                $"(maximum {ExcelMaxRow}). Reduce the CSV or change the start cell.");
        long endColIdx = (long)startColIdx + maxCols - 1;
        if (endColIdx > ExcelMaxCol)
            throw new ArgumentException(
                $"Import exceeds Excel's column limit: data would reach column {endColIdx} " +
                $"(maximum {ExcelMaxCol} / XFD). Reduce the CSV width or change the start cell.");

        // BUG-R11-import-dup-row BUG-11: import previously always appended a
        // brand-new <row r="N">, producing duplicate row entries when the
        // target rows already existed (Excel auto-repaired by keeping the
        // first one, silently losing imported data). Upsert by RowIndex —
        // reuse an existing row, otherwise insert a new one in sorted position.
        //
        // PERF(dos-hardening): the previous implementation re-scanned the whole
        // SheetData (LINQ FirstOrDefault) for every imported row AND every cell,
        // making a bulk import O(rows*cells * existing) — a 100k-row CSV took
        // 9+ minutes. Pre-index existing rows once and walk them with an
        // ascending cursor for sorted insertion; build a per-row cell index
        // only when reusing a pre-existing row. Bulk-append into a fresh sheet
        // is now linear.
        var existingRows = sheetData.Elements<Row>()
            .Where(rr => rr.RowIndex?.Value != null)
            .OrderBy(rr => rr.RowIndex!.Value)
            .ToList();
        var rowByIndex = new Dictionary<uint, Row>();
        foreach (var er in existingRows)
            rowByIndex[er.RowIndex!.Value] = er;
        int exCursor = 0; // points at the first existing row with index > last processed
        var importedFormulaCells = new List<Cell>();
        // Lazily created the first time an ISO date is imported, so a detected
        // date cell gets a date number format (matching Set/Add) instead of
        // displaying its raw serial number.
        Core.ExcelStyleManager? styleManager = null;

        for (int r = 0; r < rows.Count; r++)
        {
            var fields = rows[r];
            var rowIdx = (uint)(startRow + r);

            Dictionary<string, Cell>? cellByRef = null;
            if (rowByIndex.TryGetValue(rowIdx, out var row))
            {
                // Reused row may already hold cells — index them once for upsert.
                cellByRef = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
                foreach (var existingCell in row.Elements<Cell>())
                    if (existingCell.CellReference?.Value is { } cr)
                        cellByRef[cr] = existingCell;
            }
            else if (fields.All(string.IsNullOrEmpty))
            {
                // All-empty row on a row that doesn't exist yet: nothing to
                // write. Materializing it would add phantom <row>/<c> elements
                // absent from the imported data (dump's gap-bridge "," rows
                // land here), inflating UsedRange and non-empty iteration.
                continue;
            }
            else
            {
                row = new Row { RowIndex = rowIdx };
                // Advance the cursor past existing rows that sort before rowIdx,
                // then insert before the first existing row that sorts after it
                // (or append when none remain). O(existing) total across all rows.
                while (exCursor < existingRows.Count
                       && existingRows[exCursor].RowIndex!.Value < rowIdx)
                    exCursor++;
                if (exCursor < existingRows.Count)
                    sheetData.InsertBefore(row, existingRows[exCursor]);
                else
                    sheetData.Append(row);
            }

            for (int c = 0; c < fields.Count; c++)
            {
                var colIdx = startColIdx + c;
                var cellRef = $"{IndexToColumnName(colIdx)}{rowIdx}".ToUpperInvariant();
                Cell? cell = null;
                cellByRef?.TryGetValue(cellRef, out cell);
                if (cell == null)
                {
                    // Empty field, no pre-existing cell: skip. Creating an
                    // empty <c> here would fabricate cells the source never
                    // had; when the cell DOES exist, the empty field keeps
                    // its clear-the-value semantics below.
                    if (string.IsNullOrEmpty(fields[c])) continue;
                    cell = new Cell { CellReference = cellRef };
                    row.Append(cell);
                }
                else
                {
                    cell.CellFormula = null;
                    cell.CellValue = null;
                    cell.DataType = null;
                }
                if (SetCellValueWithTypeDetection(cell, fields[c], IsWorkbookDate1904(), decimalSeparator,
                        keepAsText: CellCarriesTextFormat(cell)))
                {
                    // Date cell — apply a date number format so it shows as a
                    // date, not the raw serial. Mirrors Set/Add (numFmt yyyy-mm-dd).
                    styleManager ??= new Core.ExcelStyleManager(_doc.WorkbookPart!);
                    cell.StyleIndex = styleManager.ApplyStyle(cell,
                        new Dictionary<string, string> { ["numberformat"] = "yyyy-mm-dd" });
                }
                if (cell.CellFormula != null && cell.CellValue == null)
                    importedFormulaCells.Add(cell);
            }
        }

        // CONSISTENCY(cell-formula-cache): `set formula=` evaluates and caches
        // the result (t= type + <v>), but the import path used to write a bare
        // <f> — a dump→import replay of a formula cell silently dropped the
        // cached value and its Error/Boolean type, so Get on the replayed file
        // disagreed with the source. Evaluate once after the whole block is in
        // (handles forward references within the imported range).
        if (importedFormulaCells.Count > 0)
        {
            var importEvaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
            foreach (var fc in importedFormulaCells)
            {
                var fText = fc.CellFormula!.Text ?? "";
                // A formula referencing a sheet that does not exist yet (dump
                // replay imports a sheet before the sheet its formula points at)
                // would evaluate to #REF! here and cache that wrong value. Leave
                // the cache empty instead; the persist-time RefreshStaleFormulaCaches
                // sweep fills it once every referenced sheet's data exists. Mirrors
                // the sweep's own missing-sheet guard.
                if (FormulaReferencesMissingSheet(fText)) continue;
                WriteFormulaResultToCell(fc, importEvaluator.TryEvaluateFull(fText));
            }
            EnsureFullCalcOnLoad();
        }

        InvalidateRowIndex(sheetData);

        // --header: set AutoFilter on data range and freeze pane below first row
        if (hasHeader && rows.Count > 0)
        {
            var endCol = IndexToColumnName(startColIdx + maxCols - 1);
            var endRow = startRow + rows.Count - 1;
            var filterRange = $"{startCol}{startRow}:{endCol}{endRow}";

            // Set AutoFilter
            var autoFilter = ws.GetFirstChild<AutoFilter>();
            if (autoFilter == null)
            {
                autoFilter = new AutoFilter();
                var mergeCells = ws.GetFirstChild<MergeCells>();
                var sd = ws.GetFirstChild<SheetData>();
                if (mergeCells != null)
                    mergeCells.InsertAfterSelf(autoFilter);
                else if (sd != null)
                    sd.InsertAfterSelf(autoFilter);
                else
                    ws.AppendChild(autoFilter);
            }
            autoFilter.Reference = filterRange;

            // Set freeze pane below first row
            var sheetViews = ws.GetFirstChild<SheetViews>();
            if (sheetViews == null)
            {
                sheetViews = new SheetViews();
                InsertSheetViewsInSchemaOrder(ws, sheetViews);
            }
            var sheetView = sheetViews.GetFirstChild<SheetView>();
            if (sheetView == null)
            {
                sheetView = new SheetView { WorkbookViewId = 0 };
                sheetViews.AppendChild(sheetView);
            }

            var existingPane = sheetView.GetFirstChild<Pane>();
            existingPane?.Remove();

            var freezeRow = startRow; // freeze after the header row
            var freezeCell = $"{startCol}{freezeRow + 1}";
            var pane = new Pane
            {
                VerticalSplit = freezeRow,
                TopLeftCell = freezeCell,
                State = PaneStateValues.Frozen,
                ActivePane = PaneValues.BottomLeft
            };
            sheetView.InsertAt(pane, 0);

            // Repeat the header row at the top of every printed page
            // (_xlnm.Print_Titles, the name `set printTitleRows` writes). A
            // sheet that already declares print titles keeps them.
            var workbook = GetWorkbook();
            var allSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
            var sheetIdx = allSheets?.FindIndex(s =>
                s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
            bool hasPrintTitles = workbook.GetFirstChild<DefinedNames>()?.Elements<DefinedName>()
                .Any(d => d.Name == "_xlnm.Print_Titles" && d.LocalSheetId?.Value == (uint)sheetIdx) == true;
            if (sheetIdx >= 0 && !hasPrintTitles)
            {
                var sheetRef = Core.ModernFunctionQualifier.QuoteSheetNameForRef(allSheets![sheetIdx].Name!.Value!);
                GetOrCreateDefinedNames(workbook).AppendChild(
                    new DefinedName($"{sheetRef}!${startRow}:${startRow}") { Name = "_xlnm.Print_Titles", LocalSheetId = (uint)sheetIdx });
                workbook.Save();
            }
        }

        // Mark the document modified so Dispose flushes it. Without this, an
        // `import` with no explicit Save (the CLI path) hits the !Modified
        // byte-preserving discard branch in Dispose and the imported rows are
        // silently dropped — success reported, disk unchanged. Same gap that
        // affected Move/Swap/CopyFrom.
        Modified = true;
        SaveWorksheet(worksheet);
        return $"Imported {rows.Count} rows x {maxCols} cols into /{sheetName} starting at {startCell.ToUpperInvariant()}";
    }

    /// <summary>
    /// Set a cell's value with automatic type detection.
    /// Order: number -> date (ISO) -> boolean -> formula -> string
    /// </summary>
    /// <returns>true when the value was stored as a DATE (serial number needing
    /// a date number format); false for every other type.</returns>
    /// <summary>
    /// True when the (pre-existing) cell is formatted as Text — numFmtId 49 /
    /// format code "@". Excel keeps whatever is typed into such a cell as a
    /// literal string, formulas and numbers included; a pre-formatted import
    /// target must get the same treatment.
    /// </summary>
    private bool CellCarriesTextFormat(Cell cell)
    {
        if (cell.StyleIndex == null) return false;
        var (numFmtId, code) = ExcelDataFormatter.GetCellFormat(cell, _doc.WorkbookPart);
        return numFmtId == 49 || (code != null && code.Trim() == "@");
    }

    /// <summary>
    /// The same rule the `set` path applies before treating a digit string as a
    /// number: an identifier-shaped literal — leading zero (007, 01234) or more
    /// digits than a double can carry (>15) — is text. Storing it numeric drops
    /// the zeros / rounds the tail in every consumer, and nothing can bring
    /// them back; the reverse (type=number on a text cell) is always available.
    /// </summary>
    private static bool LooksLikeIdentifierNotNumber(string value)
        => value.Length > 1 && value.All(char.IsDigit)
           && (value[0] == '0' || value.Length > 15);

    private static bool SetCellValueWithTypeDetection(Cell cell, string value, bool date1904,
        char decimalSeparator = '.', bool keepAsText = false)
    {
        // Empty
        if (string.IsNullOrEmpty(value))
        {
            cell.CellValue = null;
            cell.DataType = null;
            return false;
        }

        // R13-1: enforce Excel's 32767-char per-cell limit at the CSV/TSV
        // import path too, so bulk imports fail fast instead of producing a
        // file Excel refuses to open.
        EnsureCellValueLength(value, cell.CellReference?.Value);

        // Text-formatted target, or a literal that only reads correctly as
        // text: store verbatim, no formula/number/date detection.
        if (keepAsText || LooksLikeIdentifierNotNumber(value))
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            return false;
        }

        // Formula: starts with =
        if (value.StartsWith('='))
        {
            // Same A1-only guard as Add/Set: verbatim R1C1 text in <f> makes
            // real Excel refuse the file.
            ValidateFormulaCellRefs(value);
            cell.CellFormula = new CellFormula(OfficeCli.Core.PivotTableHelper.SanitizeXmlText(OfficeCli.Core.ModernFunctionQualifier.Qualify(value[1..])));
            cell.CellValue = null;
            cell.DataType = null;
            return false;
        }

        // Number (integer or decimal). Which spellings count is Core.NumericText's
        // call: by default "1,5" is NOT a number (AllowThousands would read it as
        // 15, silently 10x-ing every decimal-comma value in a de-DE / ru-RU CSV);
        // under a declared decimal comma it is rewritten to 1.5 instead.
        string? numericText = decimalSeparator == ','
            ? (Core.NumericText.TryRewriteCommaDecimal(value, out var rewritten) ? rewritten : null)
            : (HasValidThousandsGrouping(value) ? value : null);
        if (numericText != null
            && double.TryParse(numericText, NumberStyles.Any, CultureInfo.InvariantCulture, out var numVal)
            && double.IsFinite(numVal)) // "Infinity"/"NaN" parse but have no OOXML numeric form — fall through to string
        {
            // Preserve the literal digits when the input is already a plain
            // canonical numeric literal. Round-tripping through double
            // silently rounds >15-16 significant digits (e.g. 18-digit IDs
            // stored as numbers by openpyxl-authored files), so dump→replay
            // would corrupt them. Normalization is kept for the non-canonical
            // spellings double.TryParse accepts (whitespace padding,
            // thousands separators, "Infinity"/"NaN").
            cell.CellValue = new CellValue(NormalizeNumericCellText(numericText, numVal));
            cell.DataType = null; // numeric is default
            return false;
        }

        // Date: ISO 8601 formats (yyyy-MM-dd, yyyy-MM-ddTHH:mm:ss, etc.)
        if (TryParseIsoDate(value, out var dateVal))
        {
            // Excel stores dates as OLE Automation date numbers
            cell.CellValue = new CellValue(ExcelDataFormatter.ToExcelSerial(dateVal, date1904).ToString(CultureInfo.InvariantCulture));
            cell.DataType = null; // numeric
            return true; // caller applies a date number format
        }

        // Boolean: TRUE/FALSE (case-insensitive)
        if (value.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
        {
            cell.CellValue = new CellValue("1");
            cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
            return false;
        }
        if (value.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            cell.CellValue = new CellValue("0");
            cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
            return false;
        }

        // String (fallback)
        cell.CellValue = new CellValue(value);
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
        return false;
    }

    private static bool TryParseIsoDate(string value, out DateTime result)
    {
        // Try common ISO date formats
        string[] formats =
        [
            "yyyy-MM-dd",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fff",
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yyyy-MM-dd HH:mm:ss"
        ];
        return DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture,
            DateTimeStyles.None, out result);
    }

    /// <summary>
    /// Parse CSV/TSV content into a list of rows, each containing field values.
    /// Handles quoted fields, embedded delimiters, escaped quotes (""), and newlines within quotes.
    /// UTF-8 with optional BOM.
    /// </summary>
    internal static List<List<string>> ParseCsv(string content, char delimiter)
    {
        var rows = new List<List<string>>();
        if (string.IsNullOrEmpty(content))
            return rows;

        // Strip BOM if present
        if (content.Length > 0 && content[0] == '\uFEFF')
            content = content[1..];

        var currentRow = new List<string>();
        var field = new StringBuilder();
        bool inQuotes = false;
        int i = 0;

        while (i < content.Length)
        {
            char c = content[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    // Check for escaped quote ""
                    if (i + 1 < content.Length && content[i + 1] == '"')
                    {
                        field.Append('"');
                        i += 2;
                    }
                    else
                    {
                        // End of quoted field
                        inQuotes = false;
                        i++;
                    }
                }
                else
                {
                    field.Append(c);
                    i++;
                }
            }
            else
            {
                if (c == '"' && field.Length == 0)
                {
                    // Start of quoted field
                    inQuotes = true;
                    i++;
                }
                else if (c == delimiter)
                {
                    currentRow.Add(field.ToString());
                    field.Clear();
                    i++;
                }
                else if (c == '\r')
                {
                    // End of row
                    currentRow.Add(field.ToString());
                    field.Clear();
                    // Keep blank lines as single-empty-field rows: dropping
                    // them shifted every subsequent line up (r4 landed on row
                    // 2). The import loop skips materializing all-empty rows,
                    // so no phantom <row> elements are written — only the row
                    // cursor advances, preserving source line positions.
                    rows.Add(currentRow);
                    currentRow = new List<string>();
                    i++;
                    if (i < content.Length && content[i] == '\n')
                        i++; // skip \n after \r
                }
                else if (c == '\n')
                {
                    // End of row
                    currentRow.Add(field.ToString());
                    field.Clear();
                    // Blank lines preserved — see the \r branch above.
                    rows.Add(currentRow);
                    currentRow = new List<string>();
                    i++;
                }
                else
                {
                    field.Append(c);
                    i++;
                }
            }
        }

        // Last field/row
        if (field.Length > 0 || currentRow.Count > 0)
        {
            currentRow.Add(field.ToString());
            if (currentRow.Count > 0 && !(currentRow.Count == 1 && currentRow[0] == ""))
                rows.Add(currentRow);
        }

        return rows;
    }
}
