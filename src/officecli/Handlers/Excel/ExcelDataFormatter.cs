// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

/// <summary>
/// Applies Excel number format codes to raw cell values, producing display strings.
/// raw double + numFmtId + formatCode → display string.
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
    /// Resolve a built-in numFmtId to its canonical format-code string
    /// (ECMA-376 Part 1, 18.8.30). Returns null for ids that carry no
    /// implied code (a truly-custom id whose code lives in the styles part,
    /// or an unknown id). Callers use this to surface a built-in number
    /// format whose code is NOT stored in <numFmt> — mirrors the cell reader.
    /// </summary>
    public static string? ResolveBuiltInFormatCode(uint numFmtId) =>
        BuiltInFormats.TryGetValue(numFmtId, out var code) ? code : null;

    /// <summary>
    /// Convert a DateTime to an Excel 1900-system serial. .NET's ToOADate
    /// runs one day AHEAD of Excel for dates before 1900-03-01 (Excel keeps
    /// Lotus 1-2-3's fictitious 1900-02-29 as serial 60); writing the raw
    /// OADate stored 1900-02-28 as serial 60, which Excel reads back as
    /// 02-29. Mirrors the read-side adjustment in FormatDate.
    /// </summary>
    public static double ToExcelSerial(DateTime dt)
    {
        var oa = dt.ToOADate();
        return dt < new DateTime(1900, 3, 1) && oa >= 2 ? oa - 1 : oa;
    }

    /// <summary>
    /// Convert a DateTime to an Excel serial for the workbook's active date
    /// system. In a 1904-system workbook (Date1904=true) the stored serial must
    /// be 1462 days smaller than the 1900-system serial — this is the exact
    /// inverse of the read-side +1462 shift in <see cref="TryFormat"/>. Writing
    /// the 1900 serial into a 1904 workbook made Get read the date back 4 years
    /// later. Callers that turn a date string into a cell value must pass the
    /// workbook flag so the write and read sides agree.
    /// </summary>
    public static double ToExcelSerial(DateTime dt, bool date1904)
        => date1904 ? ToExcelSerial(dt) - 1462 : ToExcelSerial(dt);

    /// <summary>
    /// Shared date-format-code detector for Get and the HTML preview (which
    /// previously kept a forked contains-check that misclassified Y0.00 as a
    /// date). A format mixing date tokens with digit placeholders is not a
    /// date; fractional seconds (ss.00) are the one exception.
    /// </summary>
    public static bool LooksLikeDateFormatCode(string formatCode)
        => IsDateFormat(0, formatCode);

    /// <summary>
    /// Excel-1900-system serial → DateTime, applying the Lotus leap-bug
    /// alignment (serial 60 is the fictitious 1900-02-29 — mapped to
    /// 02-28's DateTime here since .NET has no such date, callers wanting
    /// the exact string should special-case 60; serials 1-59 run one day
    /// ahead of the OADate scale). Shared with the HTML preview.
    /// </summary>
    public static DateTime FromExcelSerial(double value)
    {
        if (value >= 60 && value < 61)
            // Fictitious 1900-02-29: no DateTime exists; approximate with
            // 02-28 (+ time fraction). Callers that render a date string
            // should special-case the day (see FormatDate / the HTML branch).
            return new DateTime(1900, 2, 28).AddDays(value - 60);
        if (value >= 1 && value < 60) value += 1;
        return DateTime.FromOADate(value);
    }

    /// <summary>
    /// Format a raw numeric cell value using its number format.
    /// Returns null if no formatting is needed (raw value is fine as-is).
    /// </summary>
    public static string? TryFormat(double value, uint numFmtId, string? customFormatCode, bool date1904 = false)
    {
        var formatCode = customFormatCode ?? (BuiltInFormats.TryGetValue(numFmtId, out var b) ? b : null);

        // Multi-section formats (pos;neg;zero;text): the DETECTORS below must
        // see only the section that applies to THIS value. Whole-string
        // scanning misclassified m/d/yy;0% — the digit in the negative
        // section defeated the date check and the % matched the percent
        // check, so a positive date serial rendered as a huge percentage.
        // Formatting itself still receives the full code where the formatter
        // has its own section handling (FormatPercent's sign-magnitude rule).
        var detectCode = formatCode;
        if (detectCode != null && detectCode.Contains(';'))
        {
            var sections = detectCode.Split(';');
            detectCode = value switch
            {
                > 0 => sections[0],
                < 0 => sections.Length > 1 ? sections[1] : sections[0],
                _ => sections.Length > 2 ? sections[2] : sections[0],
            };
            if (string.IsNullOrWhiteSpace(detectCode)) return null;
        }

        // Elapsed-time accumulator formats ([h]:mm:ss, [mm]:ss, [s]) count
        // total units instead of wrapping into a calendar time — 1.5 under
        // [h]:mm:ss is 36:00:00, not "1899-12-31 12:00". IsDateFormat strips
        // bracket codes before scanning, so these used to fall through to
        // the calendar-date path and render a nonsense wrapped date.
        if (detectCode != null
            && Regex.IsMatch(detectCode, @"\[(h+|m+|s+)\]", RegexOptions.IgnoreCase))
            return FormatElapsed(value, detectCode);

        if (IsDateFormat(numFmtId, detectCode))
            // 1904 date system: serials count from 1904-01-01, which is
            // exactly 1462 days after the 1900 system's 1899-12-30 epoch
            // that FromOADate assumes. Without the shift a date1904
            // workbook's cells read back four years early.
            return FormatDate(date1904 ? value + 1462 : value, detectCode);

        if (IsPercentFormat(detectCode))
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

        // A date/time format never mixes date tokens with numeric
        // placeholders (0/#) — the only exception is fractional seconds
        // (ss.00), normalized away first. Without this, a garbage custom
        // format like Y0.00 tripped the date heuristic and Get displayed the
        // numeric value 5000 as the date string 1913-09-08.
        var noSecondsFraction = Regex.Replace(stripped, @"s\.0+", "s", RegexOptions.IgnoreCase);
        if (noSecondsFraction.IndexOfAny(new[] { '0', '#' }) >= 0)
            return false;

        return DateTokenRegex.IsMatch(stripped);
    }

    private static bool IsPercentFormat(string? formatCode)
    {
        if (formatCode == null) return false;
        var stripped = Regex.Replace(formatCode, "\"[^\"]*\"", "");
        return stripped.Contains('%');
    }

    // Serial days → accumulated h/m/s per the first bracketed token; the
    // tokens after it decide whether minute/second parts are appended.
    private static string FormatElapsed(double value, string formatCode)
    {
        var negative = value < 0;
        var totalSeconds = (long)Math.Round(Math.Abs(value) * 86400);
        var acc = Regex.Match(formatCode, @"\[(h+|m+|s+)\]", RegexOptions.IgnoreCase);
        var unit = char.ToLowerInvariant(acc.Groups[1].Value[0]);
        var rest = formatCode[(acc.Index + acc.Length)..];
        var restStripped = Regex.Replace(rest, "\"[^\"]*\"", "");
        string text;
        switch (unit)
        {
            case 'h':
            {
                var h = totalSeconds / 3600;
                var mm = totalSeconds % 3600 / 60;
                var ss = totalSeconds % 60;
                var hasMin = restStripped.Contains('m', StringComparison.OrdinalIgnoreCase);
                var hasSec = restStripped.Contains('s', StringComparison.OrdinalIgnoreCase);
                text = hasMin && hasSec ? $"{h}:{mm:00}:{ss:00}"
                    : hasMin ? $"{h}:{mm:00}"
                    : h.ToString();
                break;
            }
            case 'm':
            {
                var mTotal = totalSeconds / 60;
                var ss = totalSeconds % 60;
                text = restStripped.Contains('s', StringComparison.OrdinalIgnoreCase)
                    ? $"{mTotal}:{ss:00}"
                    : mTotal.ToString();
                break;
            }
            default:
                text = totalSeconds.ToString();
                break;
        }
        return negative ? "-" + text : text;
    }

    private static string FormatDate(double value, string? formatCode)
    {
        try
        {
            // Excel's 1900 date system deliberately keeps Lotus 1-2-3's leap
            // bug: serial 60 is the fictitious 1900-02-29, and serials 1-59
            // (Jan/Feb 1900) run one day AHEAD of the OADate scale FromOADate
            // uses. Serials 61+ agree exactly.
            if (value >= 60 && value < 61)
                return "1900-02-29";
            if (value >= 1 && value < 60)
                value += 1;
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
        // Multi-section percent formats (positive;negative;zero). Excel picks the
        // section by the value's sign, and that section's literal characters
        // (parentheses, etc.) define the negative representation — there is no
        // automatic '-' when an explicit negative section is present. The color
        // code ([Red] etc.) is orthogonal and stripped here (it never affects the
        // text). Without an explicit negative section, Excel reuses the positive
        // section for negatives with a leading '-'.
        if (formatCode.Contains(';'))
        {
            var sections = formatCode.Split(';');
            if (value < 0 && sections.Length >= 2)
                // Negative section owns its sign/literals; format the magnitude
                // through it (no auto-minus).
                return FormatPercentSection(Math.Abs(value), sections[1]);
            if (value == 0 && sections.Length >= 3)
                return FormatPercentSection(value, sections[2]);
            return FormatPercentSection(value, sections[0]);
        }
        return FormatPercentSection(value, formatCode);
    }

    // Format a single percent section: emit the percentage with the section's
    // decimal count, preserving any literal characters (e.g. parentheses) written
    // in the section around the numeric placeholder. [Color] codes are dropped.
    private static string FormatPercentSection(double value, string section)
    {
        section = BracketCodeRegex.Replace(section, "").Trim();
        // Count decimal places from the section (e.g. "0.00%" → 2)
        var match = Regex.Match(section, @"0\.(0+)%");
        int decimals = match.Success ? match.Groups[1].Value.Length : 0;
        // A ',' in the digit placeholder run requests thousands grouping
        // (#,##0.00% renders 123,455.55%); F-format never groups.
        var numFormat = section.Contains(',') ? $"N{decimals}" : $"F{decimals}";
        // Excel rounds half away from zero; pre-round so the .NET half-to-even
        // ToString has no midpoint left to resolve the wrong way.
        var pct = Math.Round(value * 100, Math.Max(0, Math.Min(15, decimals)), MidpointRounding.AwayFromZero);
        var num = pct.ToString(numFormat, System.Globalization.CultureInfo.InvariantCulture) + "%";
        // Re-attach the section's literal prefix/suffix around the digit run by
        // replacing the numeric placeholder token (the '#'/'0'/'.'/'%' span) with
        // the rendered number. This carries '(' ... ')' accounting wrappers.
        var placeholder = Regex.Match(section, @"[#0][#0,]*(?:\.[#0]+)?%");
        if (placeholder.Success)
        {
            var prefix = section[..placeholder.Index];
            var suffix = section[(placeholder.Index + placeholder.Length)..];
            return prefix + num + suffix;
        }
        return num;
    }
}
