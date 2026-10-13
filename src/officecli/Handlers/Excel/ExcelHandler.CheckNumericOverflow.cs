// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // Excel's General format is width-adaptive: it can drop decimal places or
    // switch to scientific notation before it renders ###. The HTML formatter
    // has no column-width input, so comparing its General string against a
    // narrow column would create false positives. This scanner intentionally
    // limits itself to explicit, understood numeric/date formats whose display
    // width is stable. Unknown formats are skipped for the same reason.

    private const double NumericFitCellPaddingPt = 3.0;
    private const double NumericFitGlyphAdvance = 0.62;

    private sealed record NumericOverflowFinding(
        string Path,
        string Message,
        string Context,
        string Suggestion);

    private readonly record struct NumericMergeRange(
        int StartColumn,
        int StartRow,
        int EndColumn,
        int EndRow);

    private List<NumericOverflowFinding> CheckAllNumericOverflow(int? limit = null)
    {
        var findings = new List<NumericOverflowFinding>();
        if (limit is <= 0) return findings;
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        if (stylesheet == null) return findings;
        var renderStyles = new RenderStyleArrays(stylesheet);
        if (renderStyles.CellFormats.Length == 0) return findings;

        foreach (var (sheetName, part) in GetWorksheets(_doc))
        {
            // A hidden sheet has no delivered visual surface.
            if (IsSheetHidden(sheetName)) continue;
            var ws = part.Worksheet;
            if (ws == null) continue;
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            bool showFormulas = SheetShowsFormulas(ws);
            var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
            // BuildMergeMap is intentionally capped at row 5000 / column 200
            // for HTML DOM rendering. An issues scan covers the full worksheet,
            // so use compact, untruncated ranges and keep only the ranges active
            // for the current row.
            var mergeRanges = GetNumericMergeRanges(ws);
            var activeMerges = new List<NumericMergeRange>();
            int nextMerge = 0;
            var colWidths = GetColumnWidths(ws);
            var hiddenCols = GetHiddenColumns(ws);
            var columnStyles = GetColumnStyleIndexes(ws);
            var sheetFmtPr = ws.GetFirstChild<SheetFormatProperties>();
            double defaultColWidthPt = sheetFmtPr?.DefaultColumnWidth?.Value != null
                ? sheetFmtPr.DefaultColumnWidth.Value * ColWidthCharToPt
                : ExcelDefaultColWidthPt;

            foreach (var row in sheetData.Elements<Row>()
                         .OrderBy(row => row.RowIndex?.Value ?? uint.MaxValue))
            {
                int rowIndex = (int)(row.RowIndex?.Value ?? 0);
                if (rowIndex <= 0) continue;
                activeMerges.RemoveAll(range => range.EndRow < rowIndex);
                while (nextMerge < mergeRanges.Count
                    && mergeRanges[nextMerge].StartRow <= rowIndex)
                {
                    var range = mergeRanges[nextMerge++];
                    if (range.EndRow >= rowIndex) activeMerges.Add(range);
                }
                if (row.Hidden?.Value == true || row.Height?.Value == 0) continue;

                foreach (var cell in row.Elements<Cell>())
                {
                    // Show Formulas changes only formula cells. Ordinary numeric
                    // and date cells on the same sheet still render normally.
                    if (showFormulas && cell.CellFormula != null) continue;
                    var cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    var (startCol, cellRow) = ParseCellReference(cellRef);
                    int startColIdx = ColumnNameToIndex(startCol);
                    var mergeRange = activeMerges.FirstOrDefault(range =>
                        cellRow >= range.StartRow && cellRow <= range.EndRow
                        && startColIdx >= range.StartColumn && startColIdx <= range.EndColumn);
                    bool isMerged = mergeRange.StartColumn != 0;
                    if (isMerged
                        && (startColIdx != mergeRange.StartColumn || cellRow != mergeRange.StartRow))
                        continue;
                    int colSpan = isMerged
                        ? mergeRange.EndColumn - mergeRange.StartColumn + 1
                        : 1;
                    if (colSpan <= 0 || hiddenCols.Contains(startColIdx)) continue;

                    double totalWidthPt = 0;
                    for (int colIdx = startColIdx; colIdx < startColIdx + colSpan; colIdx++)
                    {
                        if (hiddenCols.Contains(colIdx)) continue;
                        totalWidthPt += colWidths.TryGetValue(colIdx, out var widthPt)
                            ? widthPt
                            : defaultColWidthPt;
                    }
                    if (!double.IsFinite(totalWidthPt)
                        || totalWidthPt <= NumericFitCellPaddingPt)
                        continue;

                    uint styleIndex = ResolveNumericFitStyleIndex(
                        cell, row, startColIdx, columnStyles);
                    if (!TryGetNumericFitStyle(styleIndex, stylesheet, renderStyles,
                            out var fontSizePt, out var shrinkToFit))
                        continue;
                    if (shrinkToFit) continue;

                    if (!TryGetStableNumericDisplay(
                            cell, styleIndex, stylesheet, renderStyles, evaluator,
                            out var displayValue, out var formatCode))
                        continue;

                    double requiredPt = displayValue.Length * fontSizePt * NumericFitGlyphAdvance;
                    double usablePt = totalWidthPt - NumericFitCellPaddingPt;
                    // The glyph model is deliberately approximate. Require more
                    // than one point of overflow so rounding noise at the fit
                    // boundary does not become an audit finding.
                    if (requiredPt <= usablePt + 1.0) continue;

                    double requiredWidthChars = (requiredPt + NumericFitCellPaddingPt) / ColWidthCharToPt;
                    double currentWidthChars = totalWidthPt / ColWidthCharToPt;
                    int suggestedWidth = Math.Max(1, (int)Math.Ceiling(requiredWidthChars));
                    string widthTarget = isMerged
                        ? $"merged range {cellRef}:{IndexToColumnName(startColIdx + colSpan - 1)}{row.RowIndex?.Value ?? 0}"
                        : $"column {startCol}";

                    findings.Add(new NumericOverflowFinding(
                        $"/{sheetName}/{cellRef}",
                        $"numeric overflow: '{displayValue}' at {fontSizePt:F1}pt needs {requiredWidthChars:F1} width, {widthTarget} is {currentWidthChars:F2}",
                        $"format=\"{formatCode}\" required={requiredPt:F1}pt available={usablePt:F1}pt",
                        suggestedWidth <= 255
                            ? $"suggest.width={suggestedWidth}; widen column {startCol} to at least {suggestedWidth}"
                            : "required width exceeds Excel's 255-character column limit; use shrinkToFit or a shorter number format"));
                    if (limit.HasValue && findings.Count >= limit.Value) return findings;
                }
            }
        }

        return findings;
    }

    private static bool SheetShowsFormulas(Worksheet ws)
        => ws.GetFirstChild<SheetViews>()?.Elements<SheetView>()
            .Any(view => view.ShowFormulas?.Value == true) == true;

    private static List<NumericMergeRange> GetNumericMergeRanges(Worksheet ws)
    {
        var result = new List<NumericMergeRange>();
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells == null) return result;

        foreach (var mergeCell in mergeCells.Elements<MergeCell>())
        {
            var reference = mergeCell.Reference?.Value?.Replace("$", "");
            if (string.IsNullOrWhiteSpace(reference)) continue;
            var parts = reference.Split(':');
            if (parts.Length != 2) continue;
            var (startColumnName, startRow) = ParseCellReference(parts[0]);
            var (endColumnName, endRow) = ParseCellReference(parts[1]);
            int startColumn = ColumnNameToIndex(startColumnName);
            int endColumn = ColumnNameToIndex(endColumnName);
            if (startColumn > endColumn) (startColumn, endColumn) = (endColumn, startColumn);
            if (startRow > endRow) (startRow, endRow) = (endRow, startRow);
            result.Add(new NumericMergeRange(startColumn, startRow, endColumn, endRow));
        }

        return result
            .OrderBy(range => range.StartRow)
            .ThenBy(range => range.StartColumn)
            .ToList();
    }

    private static HashSet<int> GetHiddenColumns(Worksheet ws)
    {
        var result = new HashSet<int>();
        var columns = ws.GetFirstChild<Columns>();
        if (columns == null) return result;

        foreach (var col in columns.Elements<Column>())
        {
            if (col.Hidden?.Value != true && col.Width?.Value != 0) continue;
            uint minRaw = col.Min?.Value ?? 1;
            uint maxRaw = col.Max?.Value ?? minRaw;
            int min = (int)Math.Min(16384u, Math.Max(1u, minRaw));
            int max = (int)Math.Min(16384u, Math.Max(1u, maxRaw));
            for (int colIdx = min; colIdx <= max; colIdx++) result.Add(colIdx);
        }
        return result;
    }

    private static Dictionary<int, uint> GetColumnStyleIndexes(Worksheet ws)
    {
        var result = new Dictionary<int, uint>();
        var columns = ws.GetFirstChild<Columns>();
        if (columns == null) return result;

        foreach (var col in columns.Elements<Column>())
        {
            if (col.Style?.Value is not uint styleIndex) continue;
            uint minRaw = col.Min?.Value ?? 1;
            uint maxRaw = col.Max?.Value ?? minRaw;
            int min = (int)Math.Min(16384u, Math.Max(1u, minRaw));
            int max = (int)Math.Min(16384u, Math.Max(1u, maxRaw));
            for (int colIdx = min; colIdx <= max; colIdx++) result[colIdx] = styleIndex;
        }
        return result;
    }

    private static uint ResolveNumericFitStyleIndex(
        Cell cell,
        Row row,
        int columnIndex,
        Dictionary<int, uint> columnStyles)
    {
        // Direct cell style wins even when explicitly set to 0 (General).
        if (cell.StyleIndex?.Value is uint cellStyle) return cellStyle;
        // OOXML gates a row's default style on customFormat=true.
        if (row.CustomFormat?.Value == true && row.StyleIndex?.Value is uint rowStyle)
            return rowStyle;
        return columnStyles.TryGetValue(columnIndex, out var columnStyle)
            ? columnStyle
            : 0;
    }

    private static bool TryGetNumericFitStyle(
        uint styleIndex,
        Stylesheet stylesheet,
        RenderStyleArrays renderStyles,
        out double fontSizePt,
        out bool shrinkToFit)
    {
        fontSizePt = 11.0;
        shrinkToFit = false;

        if (styleIndex >= (uint)renderStyles.CellFormats.Length)
            return false;

        var format = renderStyles.CellFormats[(int)styleIndex];
        var baseFormat = GetBaseCellStyleFormat(format, stylesheet);
        var alignment = format.Alignment;
        var baseAlignment = baseFormat?.Alignment;
        // If either level requests shrinking, skipping the finding is the
        // conservative choice. This prevents a base cellStyle's shrinkToFit
        // from being mistaken for an overflowing direct cellXf.
        shrinkToFit = alignment?.ShrinkToFit?.Value == true
            || baseAlignment?.ShrinkToFit?.Value == true;

        // Rotated/vertical text has different horizontal geometry. Skipping is
        // safer than treating the unrotated width estimate as authoritative.
        if ((alignment?.TextRotation?.Value ?? 0) != 0
            || (baseAlignment?.TextRotation?.Value ?? 0) != 0)
            return false;

        uint fontId = format.ApplyFont?.Value == false
            ? baseFormat?.FontId?.Value ?? format.FontId?.Value ?? 0
            : format.FontId?.Value ?? baseFormat?.FontId?.Value ?? 0;
        if (fontId < (uint)renderStyles.Fonts.Length)
            fontSizePt = renderStyles.Fonts[(int)fontId].FontSize?.Val?.Value ?? 11.0;
        return double.IsFinite(fontSizePt) && fontSizePt > 0;
    }

    private bool TryGetStableNumericDisplay(
        Cell cell,
        uint styleIndex,
        Stylesheet stylesheet,
        RenderStyleArrays renderStyles,
        Core.FormulaEvaluator evaluator,
        out string displayValue,
        out string formatCode)
    {
        displayValue = "";
        formatCode = "";

        if (!TryGetNumericValue(cell, evaluator, out var evaluatedValue)) return false;

        var resolvedFormat = ResolveNumericFitFormatCode(styleIndex, stylesheet, renderStyles);
        if (string.IsNullOrWhiteSpace(resolvedFormat)) return false;
        var activeFormat = SelectNumericFitFormatSection(evaluatedValue, resolvedFormat);
        if (string.IsNullOrWhiteSpace(activeFormat) || IsGeneralFormatSection(activeFormat))
            return false;
        bool isElapsedFormat = System.Text.RegularExpressions.Regex.IsMatch(
            activeFormat, @"\[(h+|m+|s+)\]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        bool isCalendarFormat = !isElapsedFormat
            && ExcelDataFormatter.LooksLikeDateFormatCode(activeFormat);
        if (!isElapsedFormat && !isCalendarFormat
            && !ContainsNumericPlaceholderOutsideQuotes(activeFormat))
            return false;

        // FormulaEvaluator mixes two date-epoch domains in 1904 workbooks:
        // DATE()/TODAY() return an absolute OADate, while a reference to a
        // stored cell returns the workbook's 1904 serial. Until that evaluator
        // contract is unified, any adjustment would be wrong for one family.
        if (isCalendarFormat && IsWorkbookDate1904() && cell.CellFormula != null)
            return false;

        // Reuse the HTML preview's canonical ApplyNumberFormat pipeline so the
        // finding describes the same formatted value OfficeCLI renders. Apply
        // the workbook's date epoch here: GetFormattedCellValue recovers a raw
        // serial after its first formatting pass and currently loses the 1904
        // adjustment, which would make an audit message name the wrong date.
        // ISO date cells already carry an absolute date and need no epoch shift.
        double displayValueNumber = evaluatedValue;
        DateTime isoDate = default;
        bool isIsoDateCell = cell.DataType?.Value == CellValues.Date
            && DateTime.TryParse(cell.CellValue?.Text, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind, out isoDate);
        if (isCalendarFormat)
        {
            if (isIsoDateCell) displayValueNumber = isoDate.ToOADate();
            else if (IsWorkbookDate1904()) displayValueNumber += 1462.0;
            // A negative date/time serial renders ### regardless of column
            // width, so a suggest.width finding would prescribe a non-fix.
            if (displayValueNumber < 0) return false;
        }
        else if (isElapsedFormat && evaluatedValue < 0 && !IsWorkbookDate1904())
        {
            // The 1900 date system renders negative elapsed/time values as
            // #### regardless of column width. A width suggestion cannot fix
            // that display, while the 1904 system can render negative time.
            return false;
        }
        else if (isElapsedFormat && isIsoDateCell)
        {
            // An ISO absolute date formatted as an elapsed duration has no
            // unambiguous serial-domain interpretation; do not guess.
            return false;
        }
        displayValue = ApplyNumberFormat(displayValueNumber, resolvedFormat);

        if (string.IsNullOrEmpty(displayValue) || IsExcelErrorText(displayValue)) return false;
        formatCode = resolvedFormat;
        return true;
    }

    private static string SelectNumericFitFormatSection(double value, string formatCode)
    {
        if (!formatCode.Contains(';')) return formatCode.Trim();
        var sections = formatCode.Split(';');
        if (System.Text.RegularExpressions.Regex.IsMatch(formatCode, @"\[[<>=]=?\d"))
            return SelectConditionalSection(value, sections);
        if (value < 0 && sections.Length >= 2) return sections[1].Trim();
        if (value == 0 && sections.Length >= 3) return sections[2].Trim();
        return sections[0].Trim();
    }

    private static bool IsGeneralFormatSection(string formatSection)
    {
        // Modifiers and literals do not make General width-stable:
        // [Red]General and General "units" still adapt to the column width.
        var semantic = System.Text.RegularExpressions.Regex.Replace(
            formatSection, "\"[^\"]*\"", "");
        semantic = System.Text.RegularExpressions.Regex.Replace(semantic, @"\[[^\]]*\]", "");
        semantic = System.Text.RegularExpressions.Regex.Replace(semantic, @"_.|\*.", "");
        semantic = System.Text.RegularExpressions.Regex.Replace(semantic, @"\\.", "");
        return semantic.Trim().Equals("General", StringComparison.OrdinalIgnoreCase);
    }

    private static CellFormat? GetBaseCellStyleFormat(
        CellFormat format,
        Stylesheet stylesheet)
    {
        if (format.FormatId?.Value is not uint baseIndex || baseIndex > int.MaxValue) return null;
        return stylesheet.CellStyleFormats?.Elements<CellFormat>()
            .ElementAtOrDefault((int)baseIndex);
    }

    private static string? ResolveNumericFitFormatCode(
        uint styleIndex,
        Stylesheet stylesheet,
        RenderStyleArrays renderStyles)
    {
        if (styleIndex >= (uint)renderStyles.CellFormats.Length) return null;
        var format = renderStyles.CellFormats[(int)styleIndex];
        var baseFormat = GetBaseCellStyleFormat(format, stylesheet);
        uint directId = format.NumberFormatId?.Value ?? 0;
        uint baseId = baseFormat?.NumberFormatId?.Value ?? 0;
        uint numberFormatId;
        if (format.ApplyNumberFormat?.Value == true || baseFormat == null)
        {
            numberFormatId = directId;
        }
        else if (format.ApplyNumberFormat?.Value == false)
        {
            numberFormatId = baseId;
        }
        else if (directId != 0 && baseId != 0 && directId != baseId)
        {
            // Ambiguous inheritance is not a safe basis for an audit finding.
            return null;
        }
        else
        {
            numberFormatId = directId != 0 ? directId : baseId;
        }
        if (numberFormatId == 0) return null; // General

        var customFormat = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId)
            ?.FormatCode?.Value;
        return customFormat ?? ExcelDataFormatter.ResolveBuiltInFormatCode(numberFormatId);
    }

    private bool TryGetNumericValue(
        Cell cell,
        Core.FormulaEvaluator evaluator,
        out double numericValue)
    {
        numericValue = 0;
        if (cell.CellFormula?.Text is { } formula)
        {
            // This is the issues scan's second evaluation funnel — the numeric-fit
            // and general-precision families both reach the evaluator here — and it
            // is not the content scan's: it calls TryEvaluateFull, not
            // EvaluateForReport. The workbook in #456 drives both of them into the
            // same exponential blowup, so both must spend the same budget (see
            // ViewIssuesFormulaBudget). Returning false leaves the cell unassessed,
            // which is what an exhausted scan does everywhere: fewer findings, no
            // wrong ones.
            if (ViewIssuesFormulaBudgetExhausted()) return false;
            var result = evaluator.TryEvaluateFull(formula);
            if (result != null)
            {
                if (!result.IsNumeric) return false;
                numericValue = result.NumericValue!.Value;
                return double.IsFinite(numericValue);
            }
        }

        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.SharedString
            || dataType == CellValues.InlineString
            || dataType == CellValues.String
            || dataType == CellValues.Boolean
            || dataType == CellValues.Error)
            return false;

        if (double.TryParse(cell.CellValue?.Text, NumberStyles.Any,
                CultureInfo.InvariantCulture, out numericValue))
            return double.IsFinite(numericValue);

        // ISO date cells (t="d") are uncommon but still numeric/date display
        // candidates. Their stable visible value is handled by the shared
        // formatter; this conversion is only for candidate classification.
        if (dataType == CellValues.Date
            && DateTime.TryParse(cell.CellValue?.Text, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind, out var date))
        {
            numericValue = date.ToOADate();
            return true;
        }
        return false;
    }

    private static bool ContainsNumericPlaceholderOutsideQuotes(string formatCode)
    {
        bool inQuote = false;
        int bracketDepth = 0;
        for (int i = 0; i < formatCode.Length; i++)
        {
            char ch = formatCode[i];
            if (ch == '\\' && i + 1 < formatCode.Length) { i++; continue; }
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (inQuote) continue;
            if (ch == '[') { bracketDepth++; continue; }
            if (ch == ']' && bracketDepth > 0) { bracketDepth--; continue; }
            if (bracketDepth == 0 && ch is '0' or '#' or '?') return true;
        }
        return false;
    }
}
