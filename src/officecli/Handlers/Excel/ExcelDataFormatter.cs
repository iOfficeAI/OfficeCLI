// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

/// <summary>
/// Applies Excel number format codes to raw cell values, producing display strings.
/// Mirrors Apache POI's DataFormatter — raw double + numFmtId + formatCode → display string.
/// </summary>
internal static class ExcelDataFormatter
{
    // Built-in Excel number format IDs that are date/time formats (ECMA-376 18.8.30)
    private static readonly HashSet<uint> BuiltInDateFormatIds = new()
        { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 };

    // Built-in format codes by ID
    private static readonly Dictionary<uint, string> BuiltInFormats = new()
    {
        [0]  = "General",
        [1]  = "0",
        [2]  = "0.00",
        [3]  = "#,##0",
        [4]  = "#,##0.00",
        [9]  = "0%",
        [10] = "0.00%",
        [11] = "0.00E+00",
        [12] = "# ?/?",
        [13] = "# ??/??",
        [14] = "m/d/yy",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "m/d/yy h:mm",
        [37] = "#,##0 ;(#,##0)",
        [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)",
        [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mmss.0",
        [48] = "##0.0E+0",
        [49] = "@",
    };

    // Regex to detect date tokens in a format code (after stripping quoted strings and brackets)
    private static readonly Regex DateTokenRegex = new(@"[yYdD]|(?<![a-zA-Z])m(?![a-zA-Z])|mm+", RegexOptions.Compiled);

    // Regex to detect time tokens (h/s) — when present alongside date, output includes time
    private static readonly Regex TimeTokenRegex = new(@"[hHsS]", RegexOptions.Compiled);

    // Strip color codes [Red], [Blue], etc. and locale codes [$xxx-yyy]
    private static readonly Regex BracketCodeRegex = new(@"\[[^\]]*\]", RegexOptions.Compiled);

    /// <summary>
    /// Format a raw numeric cell value using its number format.
    /// Returns null if no formatting is needed (raw value is fine as-is).
    /// </summary>
    public static string? TryFormat(double value, uint numFmtId, string? customFormatCode)
    {
        var formatCode = customFormatCode ?? (BuiltInFormats.TryGetValue(numFmtId, out var b) ? b : null);

        if (IsDateFormat(numFmtId, formatCode))
            return FormatDate(value, formatCode);

        if (IsPercentFormat(formatCode))
            return FormatPercent(value, formatCode!);

        return null; // let caller fall back to raw value
    }

    /// <summary>
    /// Look up a cell's numFmtId and custom format code from the workbook stylesheet.
    /// Returns (0, null) if no style is applied.
    /// </summary>
    public static (uint numFmtId, string? formatCode) GetCellFormat(Cell cell, WorkbookPart? wbPart)
    {
        if (wbPart?.WorkbookStylesPart?.Stylesheet == null)
            return (0, null);

        var styleIndex = cell.StyleIndex?.Value ?? 0;
        var cellFormats = wbPart.WorkbookStylesPart.Stylesheet.CellFormats;
        if (cellFormats == null) return (0, null);

        var xfList = cellFormats.Elements<CellFormat>().ToList();
        if (styleIndex >= (uint)xfList.Count) return (0, null);

        var xf = xfList[(int)styleIndex];
        var numFmtId = xf.NumberFormatId?.Value ?? 0;
        if (numFmtId == 0) return (0, null);

        // Look up custom format code if not built-in
        string? formatCode = null;
        var numFmts = wbPart.WorkbookStylesPart.Stylesheet.NumberingFormats;
        if (numFmts != null)
        {
            formatCode = numFmts.Elements<NumberingFormat>()
                .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId)
                ?.FormatCode?.Value;
        }

        return (numFmtId, formatCode);
    }

    private static bool IsDateFormat(uint numFmtId, string? formatCode)
    {
        if (BuiltInDateFormatIds.Contains(numFmtId)) return true;
        if (formatCode == null) return false;

        // Strip quoted strings and bracket codes before scanning for date tokens
        var stripped = Regex.Replace(formatCode, "\"[^\"]*\"", "");
        stripped = BracketCodeRegex.Replace(stripped, "");

        return DateTokenRegex.IsMatch(stripped);
    }

    private static bool IsPercentFormat(string? formatCode)
    {
        if (formatCode == null) return false;
        var stripped = Regex.Replace(formatCode, "\"[^\"]*\"", "");
        return stripped.Contains('%');
    }

    private static string FormatDate(double value, string? formatCode)
    {
        try
        {
            var dt = DateTime.FromOADate(value);

            // Detect whether time component is significant
            bool hasTime = false;
            if (formatCode != null)
            {
                var stripped = Regex.Replace(formatCode, "\"[^\"]*\"", "");
                stripped = BracketCodeRegex.Replace(stripped, "");
                hasTime = TimeTokenRegex.IsMatch(stripped);
            }

            if (hasTime)
            {
                // If fractional seconds are zero, omit them
                return dt.Second == 0 && dt.Millisecond == 0
                    ? dt.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture)
                    : dt.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
            }

            return dt.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        }
        catch
        {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private static string FormatPercent(double value, string formatCode)
    {
        // Count decimal places from format code (e.g. "0.00%" → 2)
        var match = Regex.Match(formatCode, @"0\.(0+)%");
        int decimals = match.Success ? match.Groups[1].Value.Length : 0;
        return (value * 100).ToString($"F{decimals}", System.Globalization.CultureInfo.InvariantCulture) + "%";
    }
}
