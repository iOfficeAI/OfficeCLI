// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Selector ====================

    private record CellSelector(string? Sheet, string? Column, string? ValueEquals, string? ValueNotEquals,
        string? ValueContains, bool? HasFormula, bool? IsEmpty, string? TypeEquals, string? TypeNotEquals,
        Dictionary<string, string>? FormatEquals = null, Dictionary<string, string>? FormatNotEquals = null);

    private CellSelector ParseCellSelector(string selector)
    {
        string? sheet = null;
        string? column = null;
        string? valueEquals = null;
        string? valueNotEquals = null;
        string? valueContains = null;
        bool? hasFormula = null;
        bool? isEmpty = null;
        string? typeEquals = null;
        string? typeNotEquals = null;

        // Normalize path-style selectors: "/Sheet1/cell[...]" → "Sheet1!cell[...]"
        if (selector.StartsWith('/'))
        {
            var slashIdx = selector.IndexOf('/', 1);
            if (slashIdx > 0)
            {
                sheet = selector[1..slashIdx];
                selector = selector[(slashIdx + 1)..];
            }
            else
            {
                // Just "/cell" — strip leading slash
                selector = selector[1..];
            }
        }

        // Check for sheet prefix: Sheet1!cell[...]
        // Only treat '!' as sheet separator if NOT part of '!=' operator
        var exclMatch = Regex.Match(selector, @"^(.+?)!(?!=)");
        if (exclMatch.Success)
        {
            sheet = exclMatch.Groups[1].Value;
            selector = selector[(exclMatch.Length)..];
        }

        // Parse element and attributes: cell[attr=value]
        var match = Regex.Match(selector, @"^(\w+)?(.*)$");
        var element = match.Groups[1].Value;

        // Column filter: e.g., "B" or "AB" — but NOT known element types like "row"
        if (element.Length <= 3 && Regex.IsMatch(element, @"^[A-Z]+$", RegexOptions.IgnoreCase)
            && element.ToLowerInvariant() is not ("row" or "cell" or "col"))
        {
            column = element.ToUpperInvariant();
        }

        // Parse attributes (\\?! handles zsh escaping \! as !)
        Dictionary<string, string>? formatEquals = null;
        Dictionary<string, string>? formatNotEquals = null;
        foreach (Match attrMatch in Regex.Matches(selector, @"\[([\w.]+)(\\?!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value.Replace("\\", "");
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "value" when op == "=": valueEquals = val; break;
                case "value" when op == "!=": valueNotEquals = val; break;
                case "type" when op == "=": typeEquals = val; break;
                case "type" when op == "!=": typeNotEquals = val; break;
                case "formula": hasFormula = val.ToLowerInvariant() != "false"; break;
                case "empty": isEmpty = val.ToLowerInvariant() != "false"; break;
                default:
                    if (op == "=")
                    {
                        formatEquals ??= new Dictionary<string, string>();
                        formatEquals[attrMatch.Groups[1].Value] = val;
                    }
                    else if (op == "!=")
                    {
                        formatNotEquals ??= new Dictionary<string, string>();
                        formatNotEquals[attrMatch.Groups[1].Value] = val;
                    }
                    break;
            }
        }

        // :contains() pseudo-selector
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) valueContains = containsMatch.Groups[1].Value;

        // Shorthand: "cell:text" → treat as :contains(text)
        if (valueContains == null)
        {
            var shorthandMatch = Regex.Match(selector, @"^(?:\w+)?:(?!contains|empty|has)(.+)$");
            if (shorthandMatch.Success) valueContains = shorthandMatch.Groups[1].Value;
        }

        // :empty pseudo-selector
        if (selector.Contains(":empty")) isEmpty = true;

        // :has(formula) pseudo-selector
        if (selector.Contains(":has(formula)")) hasFormula = true;

        return new CellSelector(sheet, column, valueEquals, valueNotEquals, valueContains, hasFormula, isEmpty, typeEquals, typeNotEquals, formatEquals, formatNotEquals);
    }

    private bool MatchesCellSelector(Cell cell, string sheetName, CellSelector selector)
    {
        // Column filter
        if (selector.Column != null)
        {
            var cellRef = cell.CellReference?.Value ?? "";
            var (colName, _) = ParseCellReference(cellRef);
            if (!colName.Equals(selector.Column, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        var value = GetCellDisplayValue(cell);

        // Value filters
        if (selector.ValueEquals != null && !value.Equals(selector.ValueEquals, StringComparison.OrdinalIgnoreCase))
            return false;
        if (selector.ValueNotEquals != null && value.Equals(selector.ValueNotEquals, StringComparison.OrdinalIgnoreCase))
            return false;
        if (selector.ValueContains != null && !value.Contains(selector.ValueContains, StringComparison.OrdinalIgnoreCase))
            return false;

        // Formula filter
        if (selector.HasFormula == true && cell.CellFormula == null)
            return false;
        if (selector.HasFormula == false && cell.CellFormula != null)
            return false;

        // Empty filter
        if (selector.IsEmpty == true && !string.IsNullOrEmpty(value))
            return false;
        if (selector.IsEmpty == false && string.IsNullOrEmpty(value))
            return false;

        // Type filter (use friendly names matching CellToNode output)
        if (selector.TypeEquals != null || selector.TypeNotEquals != null)
        {
            var type = GetCellTypeName(cell);
            if (selector.TypeEquals != null && !type.Equals(selector.TypeEquals, StringComparison.OrdinalIgnoreCase))
                return false;
            if (selector.TypeNotEquals != null && type.Equals(selector.TypeNotEquals, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    private static string GetCellTypeName(Cell cell)
    {
        if (cell.DataType?.HasValue != true) return "Number";
        var dt = cell.DataType.Value;
        if (dt == CellValues.String) return "String";
        if (dt == CellValues.SharedString) return "SharedString";
        if (dt == CellValues.Boolean) return "Boolean";
        if (dt == CellValues.Error) return "Error";
        if (dt == CellValues.InlineString) return "InlineString";
        if (dt == CellValues.Date) return "Date";
        return "Number";
    }

    private static bool MatchesFormatAttributes(DocumentNode node, CellSelector selector)
    {
        if (selector.FormatEquals != null)
        {
            foreach (var (key, expected) in selector.FormatEquals)
            {
                var matchedKey = node.Format.Keys.FirstOrDefault(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase));
                if (matchedKey == null) return false;
                var actual = node.Format[matchedKey]?.ToString() ?? "";
                if (!ColorNormalizedEquals(actual, expected))
                    return false;
            }
        }
        if (selector.FormatNotEquals != null)
        {
            foreach (var (key, expected) in selector.FormatNotEquals)
            {
                var matchedKey = node.Format.Keys.FirstOrDefault(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase));
                var actual = matchedKey != null ? (node.Format[matchedKey]?.ToString() ?? "") : "";
                if (ColorNormalizedEquals(actual, expected))
                    return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Compare two strings with color-aware normalization: "#FF0000" matches "FF0000".
    /// </summary>
    private static bool ColorNormalizedEquals(string a, string b)
    {
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return true;
        return string.Equals(a.TrimStart('#'), b.TrimStart('#'), StringComparison.OrdinalIgnoreCase);
    }

    // ==================== Cell Reference Utils ====================

    private static (string Column, int Row) ParseCellReference(string cellRef)
    {
        var match = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success)
            throw new ArgumentException($"Invalid cell reference: '{cellRef}'. Expected format like 'A1', 'B2', 'XFD1048576'.");
        var col = match.Groups[1].Value.ToUpperInvariant();
        var row = int.Parse(match.Groups[2].Value);
        if (row < 1 || row > 1048576)
            throw new ArgumentException($"Row {row} in cell reference '{cellRef}' is out of range. Valid range: 1-1048576.");
        var colIdx = ColumnNameToIndex(col);
        if (colIdx < 1 || colIdx > 16384)
            throw new ArgumentException($"Column '{col}' in cell reference '{cellRef}' is out of range. Valid range: A-XFD (1-16384).");
        return (col, row);
    }

    private static int ColumnNameToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
        {
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    private static string IndexToColumnName(int index)
    {
        var result = "";
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        return result;
    }

    private static DocumentFormat.OpenXml.Packaging.ChartPart GetChartPart(WorksheetPart worksheetPart, int index)
    {
        var drawingsPart = worksheetPart.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/charts");
        var chartParts = drawingsPart.ChartParts.ToList();
        if (index < 1 || index > chartParts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{chartParts.Count})");
        return chartParts[index - 1];
    }

    private DocumentFormat.OpenXml.Packaging.ChartPart GetGlobalChartPart(int index)
    {
        var allCharts = new List<DocumentFormat.OpenXml.Packaging.ChartPart>();
        foreach (var (_, worksheetPart) in GetWorksheets())
        {
            if (worksheetPart.DrawingsPart != null)
                allCharts.AddRange(worksheetPart.DrawingsPart.ChartParts);
        }
        if (allCharts.Count == 0)
            throw new ArgumentException("No charts found in workbook");
        if (index < 1 || index > allCharts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{allCharts.Count})");
        return allCharts[index - 1];
    }
}
