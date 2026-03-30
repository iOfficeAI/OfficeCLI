// Copyright 2025 OfficeCli (officecli.ai)
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
    /// <param name="hasHeader">If true, set AutoFilter and freeze pane on first row</param>
    /// <param name="startCell">Starting cell reference, e.g. "A1"</param>
    /// <returns>Summary of rows/cols imported</returns>
    public string Import(string parentPath, string csvContent, char delimiter, bool hasHeader, string startCell)
    {
        parentPath = NormalizeExcelPath(parentPath);
        parentPath = ResolveSheetIndexInPath(parentPath);
        var sheetName = parentPath.TrimStart('/').Split('/', 2)[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>()
            ?? ws.AppendChild(new SheetData());

        // Parse start cell
        var (startCol, startRow) = ParseCellReference(startCell.ToUpperInvariant());
        var startColIdx = ColumnNameToIndex(startCol);

        // Parse CSV
        var rows = ParseCsv(csvContent, delimiter);
        if (rows.Count == 0)
            return "No data to import";

        int maxCols = 0;
        for (int r = 0; r < rows.Count; r++)
        {
            var fields = rows[r];
            if (fields.Count > maxCols) maxCols = fields.Count;
            var rowIdx = startRow + r;

            for (int c = 0; c < fields.Count; c++)
            {
                var colIdx = startColIdx + c;
                var cellRef = $"{IndexToColumnName(colIdx)}{rowIdx}";
                var cell = FindOrCreateCell(sheetData, cellRef);
                SetCellValueWithTypeDetection(cell, fields[c]);
            }
        }

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
                ws.InsertAt(sheetViews, 0);
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
        }

        SaveWorksheet(worksheet);
        return $"Imported {rows.Count} rows x {maxCols} cols into /{sheetName} starting at {startCell.ToUpperInvariant()}";
    }

    /// <summary>
    /// Set a cell's value with automatic type detection.
    /// Order: number -> date (ISO) -> boolean -> formula -> string
    /// </summary>
    private static void SetCellValueWithTypeDetection(Cell cell, string value)
    {
        // Empty
        if (string.IsNullOrEmpty(value))
        {
            cell.CellValue = null;
            cell.DataType = null;
            return;
        }

        // Formula: starts with =
        if (value.StartsWith('='))
        {
            cell.CellFormula = new CellFormula(value[1..]);
            cell.CellValue = null;
            cell.DataType = null;
            return;
        }

        // Number (integer or decimal)
        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var numVal))
        {
            cell.CellValue = new CellValue(numVal.ToString(CultureInfo.InvariantCulture));
            cell.DataType = null; // numeric is default
            return;
        }

        // Date: ISO 8601 formats (yyyy-MM-dd, yyyy-MM-ddTHH:mm:ss, etc.)
        if (TryParseIsoDate(value, out var dateVal))
        {
            // Excel stores dates as OLE Automation date numbers
            cell.CellValue = new CellValue(dateVal.ToOADate().ToString(CultureInfo.InvariantCulture));
            cell.DataType = null; // numeric
            return;
        }

        // Boolean: TRUE/FALSE (case-insensitive)
        if (value.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
        {
            cell.CellValue = new CellValue("1");
            cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
            return;
        }
        if (value.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            cell.CellValue = new CellValue("0");
            cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
            return;
        }

        // String (fallback)
        cell.CellValue = new CellValue(value);
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
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
                    if (currentRow.Count > 0 && !(currentRow.Count == 1 && currentRow[0] == ""))
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
                    if (currentRow.Count > 0 && !(currentRow.Count == 1 && currentRow[0] == ""))
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
