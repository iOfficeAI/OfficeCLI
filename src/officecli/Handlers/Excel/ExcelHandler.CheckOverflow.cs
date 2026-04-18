// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // CONSISTENCY(text-overflow-check): mirrors PowerPointHandler.CheckShapeTextOverflow.
    // Narrow scope vs PPT: only flags wrapText cells where row height is fixed too small
    // (merged cells, or non-merged cells with explicit customHeight). Skips overflow-right
    // on non-wrapText cells — that is Excel's normal rendering, not a bug.

    /// <summary>
    /// Scan every sheet for cells whose wrapped text cannot fit inside the visible
    /// row-height budget. Returns (path, message) pairs suitable for the `check`
    /// command output. Mirrors PowerPointHandler's CheckShapeTextOverflow pattern.
    /// </summary>
    public List<(string Path, string Message)> CheckAllCellOverflow()
    {
        var issues = new List<(string, string)>();
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;

        foreach (var (sheetName, part) in GetWorksheets(_doc))
        {
            var ws = part.Worksheet;
            if (ws == null) continue;
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            var ctx = BuildOverflowContext(ws, sheetData);
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;
                    var msg = EvaluateCellOverflow(cell, cellRef, stylesheet, ctx);
                    if (msg != null) issues.Add(($"/{sheetName}/{cellRef}", msg));
                }
            }
        }
        return issues;
    }

    /// <summary>
    /// Check overflow on a single cell identified by a DOM path like "/SheetName/A16"
    /// or Excel notation "SheetName!A16". Returns warning or null.
    /// Used by `add`/`set` command dispatchers to warn inline after edits.
    /// </summary>
    public string? CheckCellOverflow(string path)
    {
        if (string.IsNullOrEmpty(path)) return null;

        // Accept "/Sheet/A1", "Sheet!A1", or bare "A1" (falls back to first sheet).
        string? sheetName = null;
        string cellRef = path;
        var slashIdx = -1;
        if (path.StartsWith('/'))
        {
            slashIdx = path.IndexOf('/', 1);
            if (slashIdx > 0)
            {
                sheetName = path[1..slashIdx];
                cellRef = path[(slashIdx + 1)..];
            }
        }
        else
        {
            var excl = path.IndexOf('!');
            if (excl > 0)
            {
                sheetName = path[..excl];
                cellRef = path[(excl + 1)..];
            }
        }

        // Bail if the remainder isn't a plain cell ref (e.g. "A16" — reject "row[1]" etc.)
        if (!Regex.IsMatch(cellRef, @"^[A-Za-z]+\d+$")) return null;
        cellRef = cellRef.ToUpperInvariant();

        var worksheets = GetWorksheets(_doc);
        if (worksheets.Count == 0) return null;
        var resolved = sheetName != null
            ? worksheets.FirstOrDefault(w => w.Name.Equals(ResolveSheetName(sheetName!), StringComparison.OrdinalIgnoreCase))
            : worksheets[0];
        if (resolved.Part == null) return null;

        var ws = resolved.Part.Worksheet;
        var sheetData = ws?.GetFirstChild<SheetData>();
        if (ws == null || sheetData == null) return null;

        var (startCol, startRow) = ParseCellReference(cellRef);
        var cell = sheetData.Elements<Row>()
            .FirstOrDefault(r => (int)(r.RowIndex?.Value ?? 0) == startRow)
            ?.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (cell == null) return null;

        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        var ctx = BuildOverflowContext(ws, sheetData);
        return EvaluateCellOverflow(cell, cellRef, stylesheet, ctx);
    }

    private record OverflowContext(
        Dictionary<string, MergeInfo> MergeMap,
        Dictionary<int, double> ColWidths,
        Dictionary<int, (double Height, bool Custom)> RowHeights,
        double DefaultRowHeightPt,
        double DefaultColWidthPt);

    private OverflowContext BuildOverflowContext(Worksheet ws, SheetData sheetData)
    {
        var rowHeights = new Dictionary<int, (double Height, bool Custom)>();
        foreach (var row in sheetData.Elements<Row>())
        {
            int rIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rIdx == 0 || row.Height?.Value == null) continue;
            rowHeights[rIdx] = (row.Height.Value, row.CustomHeight?.Value == true);
        }
        var sheetFmtPr = ws.GetFirstChild<SheetFormatProperties>();
        double defaultRowHeightPt = sheetFmtPr?.DefaultRowHeight?.Value ?? 15.0;
        double defaultColWidthPt = sheetFmtPr?.DefaultColumnWidth?.Value != null
            ? sheetFmtPr.DefaultColumnWidth.Value * 7.0017 * 0.75
            : 8.43 * 7.0017 * 0.75;
        return new OverflowContext(BuildMergeMap(ws), GetColumnWidths(ws), rowHeights,
            defaultRowHeightPt, defaultColWidthPt);
    }

    private string? EvaluateCellOverflow(Cell cell, string cellRef, Stylesheet? stylesheet, OverflowContext ctx)
    {
        bool isMerged = ctx.MergeMap.TryGetValue(cellRef, out var mInfo);
        if (isMerged && !mInfo.IsAnchor) return null;

        if (!TryGetCellAlignmentAndFont(cell, stylesheet, out var wrapText, out var fontSizePt))
            return null;
        if (!wrapText) return null;

        var text = GetCellDisplayValue(cell);
        if (string.IsNullOrEmpty(text)) return null;

        var (startCol, startRow) = ParseCellReference(cellRef);
        int startColIdx = ColumnNameToIndex(startCol);
        int rowSpan = isMerged ? mInfo.RowSpan : 1;
        int colSpan = isMerged ? mInfo.ColSpan : 1;

        // Non-merged cells with wrapText default to auto-fit — only flag when someone
        // explicitly pinned the row height (customHeight="1").
        if (!isMerged)
        {
            if (!ctx.RowHeights.TryGetValue(startRow, out var rh) || !rh.Custom)
                return null;
        }

        double usableWidth = 0;
        for (int c = startColIdx; c < startColIdx + colSpan; c++)
            usableWidth += ctx.ColWidths.TryGetValue(c, out var w) ? w : ctx.DefaultColWidthPt;
        usableWidth -= 6; // ~3pt side padding total

        double usableHeight = 0;
        for (int r = startRow; r < startRow + rowSpan; r++)
            usableHeight += ctx.RowHeights.TryGetValue(r, out var rh2) ? rh2.Height : ctx.DefaultRowHeightPt;
        usableHeight -= 4; // ~2pt top/bottom padding total

        if (usableWidth <= 0 || usableHeight <= 0) return null;

        double lineHeight = fontSizePt * 1.2;
        int totalLines = CountWrappedLines(text, fontSizePt, usableWidth);
        double needed = totalLines * lineHeight;
        // Require at least ~30% of one line to be clipped. 1-2pt differences are
        // rendering-metric noise and would drown real issues in false positives.
        if (needed - usableHeight < lineHeight * 0.3) return null;

        string mergeNote = isMerged
            ? $" (merged {cellRef}:{IndexToColumnName(startColIdx + colSpan - 1)}{startRow + rowSpan - 1})"
            : "";
        string suggest;
        if (isMerged)
        {
            double perRowPt = Math.Ceiling((needed + 4) / rowSpan / 5.0) * 5.0;
            suggest = $"suggest.rowHeight={perRowPt:F0}pt per row (Excel does not auto-fit merged rows)";
        }
        else
        {
            suggest = "suggest: clear customHeight to let Excel auto-fit";
        }
        return $"text overflow{mergeNote}: {totalLines} lines at {fontSizePt:F1}pt need {needed:F0}pt, usable {usableHeight:F0}pt. {suggest}";
    }

    private static int CountWrappedLines(string text, double fontSizePt, double usableWidthPt)
    {
        // Newline handling mirrors PowerPointHandler.CheckTextOverflow: both literal
        // and escaped "\n" split into separate paragraphs.
        var paragraphs = text.Replace("\\n", "\n").Split('\n');
        int total = 0;
        foreach (var segment in paragraphs)
        {
            if (segment.Length == 0) { total++; continue; }
            int lines = 1;
            double w = 0;
            foreach (char ch in segment)
            {
                double cw = ParseHelpers.IsCjkOrFullWidth(ch) ? fontSizePt : fontSizePt * 0.55;
                if (w + cw > usableWidthPt && w > 0) { lines++; w = cw; }
                else { w += cw; }
            }
            total += lines;
        }
        return total;
    }

    private static bool TryGetCellAlignmentAndFont(
        Cell cell, Stylesheet? stylesheet, out bool wrapText, out double fontSizePt)
    {
        wrapText = false;
        fontSizePt = 11.0; // Excel default body font
        if (stylesheet == null) return true;

        var styleIndex = (int)(cell.StyleIndex?.Value ?? 0);
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null) return true;
        var xfList = cellFormats.Elements<CellFormat>().ToList();
        if (styleIndex >= xfList.Count) return true;
        var xf = xfList[styleIndex];

        wrapText = xf.Alignment?.WrapText?.Value == true;

        var fonts = stylesheet.Fonts;
        if (fonts != null)
        {
            var fontId = (int)(xf.FontId?.Value ?? 0);
            var fontList = fonts.Elements<Font>().ToList();
            if (fontId < fontList.Count)
            {
                var sz = fontList[fontId].FontSize?.Val?.Value;
                if (sz.HasValue) fontSizePt = sz.Value;
            }
        }
        return true;
    }
}
