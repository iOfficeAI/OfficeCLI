// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Generate a self-contained HTML file that previews all sheets as spreadsheet tables.
    /// Supports cell formatting (font, fill, borders, alignment), merged cells,
    /// column widths, row heights, frozen panes, and sheet tab switching.
    /// </summary>
    public string ViewAsHtml()
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        var wbStylesPart = _doc.WorkbookPart?.WorkbookStylesPart;
        var stylesheet = wbStylesPart?.Stylesheet;

        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateExcelCss());
        sb.AppendLine("</style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        // File title
        sb.AppendLine($"<div class=\"file-title\">{HtmlEncode(Path.GetFileName(_filePath))}</div>");

        // Sheet content areas (tabs moved to bottom)
        sb.AppendLine("<div class=\"sheet-slider\">");
        for (int sheetIdx = 0; sheetIdx < sheets.Count; sheetIdx++)
        {
            var (sheetName, worksheetPart) = sheets[sheetIdx];
            var activeClass = sheetIdx == 0 ? " active" : "";
            // Check if sheet is RTL
            var sheetView = GetSheet(worksheetPart).GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
            var isRtl = sheetView?.RightToLeft?.Value == true;
            var dirAttr = isRtl ? " dir=\"rtl\"" : "";
            sb.AppendLine($"<div class=\"sheet-content{activeClass}\" data-sheet=\"{sheetIdx}\"{dirAttr}>");
            var charts = CollectSheetCharts(worksheetPart);
            RenderSheetTable(sb, sheetName, worksheetPart, stylesheet, charts);
            sb.AppendLine("</div>");
        }
        sb.AppendLine("</div>");

        // Sheet tabs at bottom (like real Excel)
        sb.AppendLine("<div class=\"sheet-tabs\" role=\"tablist\">");
        for (int i = 0; i < sheets.Count; i++)
        {
            var activeClass = i == 0 ? " active" : "";
            var tabColorStyle = "";
            var sheetProps = GetSheet(sheets[i].Part).GetFirstChild<SheetProperties>();
            var tabColorEl = sheetProps?.TabColor;
            if (tabColorEl?.Rgb?.Value != null)
            {
                var rgb = tabColorEl.Rgb.Value;
                if (rgb.Length > 6) rgb = rgb[^6..];
                tabColorStyle = $" style=\"--tab-color:#{rgb}\"";
            }
            sb.AppendLine($"  <div class=\"sheet-tab{activeClass}\"{tabColorStyle} data-sheet=\"{i}\" role=\"tab\" tabindex=\"0\" onclick=\"switchSheet({i})\" onkeydown=\"if(event.key==='Enter'||event.key===' ')switchSheet({i})\">{HtmlEncode(sheets[i].Name)}</div>");
        }
        sb.AppendLine("</div>");

        // Sheet switching JavaScript
        sb.AppendLine("<script>");
        sb.AppendLine(GenerateExcelJs());
        sb.AppendLine("</script>");

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        return sb.ToString();
    }

    /// <summary>
    /// Get the number of sheets (for watch notifications).
    /// </summary>
    public int GetSheetCount() => GetWorksheets().Count;

    /// <summary>Get the 0-based index of a sheet by name, or -1 if not found.</summary>
    public int GetSheetIndex(string sheetName)
    {
        var sheets = GetWorksheets();
        for (int i = 0; i < sheets.Count; i++)
            if (string.Equals(sheets[i].Name, sheetName, System.StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }

    // ==================== Sheet Rendering ====================

    private void RenderSheetTable(StringBuilder sb, string sheetName, WorksheetPart worksheetPart, Stylesheet? stylesheet,
        List<(int fromRow, int toRow, string html)>? charts = null)
    {
        var ws = GetSheet(worksheetPart);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            // Don't show "Empty sheet" if there are charts
            if (worksheetPart.DrawingsPart?.WorksheetDrawing == null)
                sb.AppendLine("<div class=\"empty-sheet\">Empty sheet</div>");
            return;
        }

        // Read default dimensions from sheetFormatPr
        var sheetFmtPr = ws.GetFirstChild<SheetFormatProperties>();
        var defaultColWidthPt = sheetFmtPr?.DefaultColumnWidth?.Value != null
            ? sheetFmtPr.DefaultColumnWidth.Value * 5.625 + 3.75 : 48.0;
        var defaultRowHeightPt = sheetFmtPr?.DefaultRowHeight?.Value ?? 15.0;

        // Read default font size from stylesheet
        var defaultFontPt = 11.0;
        if (stylesheet?.Fonts != null && stylesheet.Fonts.Elements<Font>().Any())
        {
            var defFont = stylesheet.Fonts.Elements<Font>().First();
            defaultFontPt = defFont.FontSize?.Val?.Value ?? 11.0;
        }

        // Create formula evaluator for this sheet to compute uncached formula values
        var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);

        // Collect merge info
        var mergeMap = BuildMergeMap(ws);

        // Build conditional formatting CSS overrides
        var cfMap = BuildConditionalFormatMap(ws, stylesheet, sheetData, _doc.WorkbookPart);

        // Collect column widths
        var colWidths = GetColumnWidths(ws);

        // Detect frozen panes
        var (frozenRows, frozenCols) = GetFrozenPanes(ws);

        // Compute cumulative left offsets for frozen columns (for sticky positioning)
        // Index 0 = row header width (30pt), index 1 = col 1 left offset, etc.
        var frozenLeftOffsets = new Dictionary<int, double>();
        if (frozenCols > 0)
        {
            double cumLeft = 30; // row header width in pt
            for (int fc = 1; fc <= frozenCols; fc++)
            {
                frozenLeftOffsets[fc] = cumLeft;
                cumLeft += colWidths.TryGetValue(fc, out var w) ? w : defaultColWidthPt;
            }
        }

        // Determine grid dimensions
        var rows = sheetData.Elements<Row>().ToList();
        int maxCol = 0;
        int maxRow = 0;
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx > maxRow) maxRow = rowIdx;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef != null)
                {
                    var (colName, _) = ParseCellReference(cellRef);
                    var colIdx = ColumnNameToIndex(colName);
                    if (colIdx > maxCol) maxCol = colIdx;
                }
            }
        }

        // Empty sheet (SheetData exists but no rows/cells)
        if (maxRow == 0 || maxCol == 0)
        {
            if (charts == null || charts.Count == 0)
            {
                if (worksheetPart.DrawingsPart?.WorksheetDrawing == null)
                    sb.AppendLine("<div class=\"empty-sheet\">Empty sheet</div>");
                return;
            }
            // Charts exist but no cell data — just render charts
            foreach (var (_, _, html) in charts)
                sb.Append(html);
            return;
        }

        // Extend maxRow to include chart anchor ranges so charts render at their position
        if (charts != null)
            foreach (var (_, toRow, _) in charts)
                if (toRow > maxRow) maxRow = toRow;

        // Limit rendering to reasonable size
        var actualRow = maxRow;
        var actualCol = maxCol;
        maxRow = Math.Min(maxRow, 5000);
        maxCol = Math.Min(maxCol, 200);
        var truncated = actualRow > maxRow || actualCol > maxCol;

        // Build cell lookup: (row, col) → Cell
        var cellMap = new Dictionary<(int row, int col), Cell>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx > maxRow) break;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;
                var (colName, _) = ParseCellReference(cellRef);
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx <= maxCol)
                    cellMap[(rowIdx, colIdx)] = cell;
            }
        }

        // Row height and hidden row lookup
        var rowHeights = new Dictionary<int, double>();
        var hiddenRows = new HashSet<int>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (row.CustomHeight?.Value == true && row.Height?.Value != null)
                rowHeights[rowIdx] = row.Height.Value;
            if (row.Hidden?.Value == true)
                hiddenRows.Add(rowIdx);
        }

        // Compute cumulative top offsets for frozen rows (for sticky positioning)
        // Includes thead height (~24pt for column headers)
        var frozenTopOffsets = new Dictionary<int, double>();
        if (frozenRows > 0)
        {
            double cumTop = 24; // approximate thead (column header) height
            for (int fr = 1; fr <= frozenRows; fr++)
            {
                frozenTopOffsets[fr] = cumTop;
                if (rowHeights.TryGetValue(fr, out var rh))
                    cumTop += rh;
                else
                {
                    // Estimate row height from max font size in the row's cells
                    double maxFontPt = defaultFontPt;
                    foreach (var cell in cellMap.Where(kv => kv.Key.row == fr).Select(kv => kv.Value))
                    {
                        var si = cell.StyleIndex?.Value ?? 0;
                        if (stylesheet?.CellFormats != null && si < (uint)stylesheet.CellFormats.Elements<CellFormat>().Count())
                        {
                            var xf = stylesheet.CellFormats.Elements<CellFormat>().ElementAt((int)si);
                            var fontId = xf.FontId?.Value ?? 0;
                            if (stylesheet.Fonts != null && fontId < (uint)stylesheet.Fonts.Elements<Font>().Count())
                            {
                                var font = stylesheet.Fonts.Elements<Font>().ElementAt((int)fontId);
                                var sz = font.FontSize?.Val?.Value ?? defaultFontPt;
                                if (sz > maxFontPt) maxFontPt = sz;
                            }
                        }
                    }
                    cumTop += maxFontPt * 1.4 + 4; // font height + padding
                }
            }
        }

        // Collect hidden columns
        var hiddenCols = new HashSet<int>();
        foreach (var (colIdx, widthPx) in colWidths)
        {
            if (widthPx <= 0) hiddenCols.Add(colIdx);
        }

        // Start table
        sb.AppendLine("<div class=\"table-wrapper\">");
        sb.AppendLine("<table>");
        sb.AppendLine($"<caption class=\"sr-only\">{HtmlEncode(sheetName)}</caption>");

        // Colgroup for column widths + header column (skip hidden columns to match td count)
        sb.Append("<colgroup><col class=\"row-header-col\">");
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue; // skip hidden cols — tds are also skipped
            var width = colWidths.TryGetValue(c, out var w) ? w : defaultColWidthPt;
            sb.Append($"<col style=\"width:{width:0.##}pt\">");
        }
        sb.AppendLine("</colgroup>");

        // Column header row
        sb.Append("<thead><tr><th class=\"corner-cell\"");
        if (frozenRows > 0 || frozenCols > 0) sb.Append(" style=\"position:sticky;top:0;left:0;z-index:4\"");
        sb.Append("></th>");
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue;
            var colName = IndexToColumnName(c);
            var isFrozenColHeader = frozenCols > 0 && c <= frozenCols;
            string stickyStyle;
            if (frozenRows > 0 && isFrozenColHeader)
            {
                var leftPt = frozenLeftOffsets.TryGetValue(c, out var lf) ? lf : 0;
                stickyStyle = $" style=\"position:sticky;top:0;left:{leftPt:0.##}pt;z-index:4\"";
            }
            else if (frozenRows > 0)
                stickyStyle = " style=\"position:sticky;top:0;z-index:3\"";
            else if (isFrozenColHeader)
            {
                var leftPt = frozenLeftOffsets.TryGetValue(c, out var lf2) ? lf2 : 0;
                stickyStyle = $" style=\"position:sticky;left:{leftPt:0.##}pt;z-index:3\"";
            }
            else
                stickyStyle = "";
            sb.Append($"<th class=\"col-header\"{stickyStyle}>{colName}</th>");
        }
        sb.AppendLine("</tr></thead>");

        // Build chart lookup: fromRow → (toRow, html) for inline insertion
        var chartAtRow = new Dictionary<int, (int toRow, string html)>();
        if (charts != null)
            foreach (var (fromRow, toRow, html) in charts)
                chartAtRow[fromRow] = (toRow, html);

        // Visible column count for chart colspan
        var visibleColCount = Enumerable.Range(1, maxCol).Count(c => !hiddenCols.Contains(c));

        // Data rows
        sb.AppendLine("<tbody>");
        for (int r = 1; r <= maxRow; r++)
        {
            // Insert chart at its anchor row position
            if (chartAtRow.TryGetValue(r, out var chartEntry))
            {
                sb.AppendLine($"<tr><td colspan=\"{visibleColCount + 1}\" style=\"padding:0;border:none\">");
                sb.Append(chartEntry.html);
                sb.AppendLine("</td></tr>");
                r = chartEntry.toRow - 1;
                continue;
            }
            if (charts != null && charts.Any(ch => r > ch.fromRow && r < ch.toRow)) continue;

            if (hiddenRows.Contains(r)) { sb.AppendLine("<tr style=\"display:none\"></tr>"); continue; }
            bool isRowFrozen = frozenRows > 0 && r <= frozenRows;
            var rowStyles = new List<string>();
            if (rowHeights.TryGetValue(r, out var rh)) rowStyles.Add($"height:{rh:0.##}pt");
            if (isRowFrozen) rowStyles.Add("background:#fff");
            var rowStyle = rowStyles.Count > 0 ? $" style=\"{string.Join(";", rowStyles)}\"" : "";
            var frozenAttr = isRowFrozen ? " data-frozen=\"1\"" : "";
            sb.Append($"<tr{rowStyle}{frozenAttr}>");

            // Row header
            string rowHeaderStyle;
            if (isRowFrozen)
                rowHeaderStyle = " style=\"position:sticky;top:0;left:0;z-index:3\"";
            else if (frozenCols > 0)
                rowHeaderStyle = " style=\"position:sticky;left:0;z-index:2\"";
            else
                rowHeaderStyle = "";
            sb.Append($"<th class=\"row-header\"{rowHeaderStyle}>{r}</th>");

            for (int c = 1; c <= maxCol; c++)
            {
                if (hiddenCols.Contains(c)) continue;
                // Check if this cell is hidden by a merge (non-anchor cell in merged range)
                var cellRef = $"{IndexToColumnName(c)}{r}";
                if (mergeMap.TryGetValue(cellRef, out var mergeInfo))
                {
                    if (!mergeInfo.IsAnchor) continue; // skip non-anchor cells

                    var cell = cellMap.TryGetValue((r, c), out var mc) ? mc : null;
                    var style = GetCellStyleCss(cell, stylesheet, frozenRows, frozenCols, r, c, frozenLeftOffsets, frozenTopOffsets, cfMap);
                    var value = cell != null ? GetFormattedCellValue(cell, stylesheet, evaluator) : "";
                    // Adjust colspan to exclude hidden columns within the merge range
                    var adjColSpan = mergeInfo.ColSpan;
                    if (adjColSpan > 1 && hiddenCols.Count > 0)
                    {
                        for (int hc = c + 1; hc < c + mergeInfo.ColSpan; hc++)
                            if (hiddenCols.Contains(hc)) adjColSpan--;
                    }
                    var spanAttrs = "";
                    if (adjColSpan > 1) spanAttrs += $" colspan=\"{adjColSpan}\"";
                    if (mergeInfo.RowSpan > 1) spanAttrs += $" rowspan=\"{mergeInfo.RowSpan}\"";

                    sb.Append($"<td{spanAttrs}{style}>{CellHtml(value)}</td>");
                }
                else
                {
                    var cell = cellMap.TryGetValue((r, c), out var nc) ? nc : null;
                    var style = GetCellStyleCss(cell, stylesheet, frozenRows, frozenCols, r, c, frozenLeftOffsets, frozenTopOffsets, cfMap);
                    var value = cell != null ? GetFormattedCellValue(cell, stylesheet, evaluator) : "";
                    sb.Append($"<td{style}>{CellHtml(value)}</td>");
                }
            }
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</tbody>");
        sb.AppendLine("</table>");
        // Truncation warning
        if (truncated)
            sb.AppendLine($"<div class=\"truncation-warning\">Showing {maxRow} of {actualRow} rows, {maxCol} of {actualCol} columns</div>");
        sb.AppendLine("</div>");
    }

    // ==================== Merge Map ====================

    private record struct MergeInfo(bool IsAnchor, int RowSpan, int ColSpan);

    private Dictionary<string, MergeInfo> BuildMergeMap(Worksheet ws)
    {
        var map = new Dictionary<string, MergeInfo>(StringComparer.OrdinalIgnoreCase);
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells == null) return map;

        foreach (var mc in mergeCells.Elements<MergeCell>())
        {
            var rangeRef = mc.Reference?.Value;
            if (string.IsNullOrEmpty(rangeRef) || !rangeRef.Contains(':')) continue;

            var parts = rangeRef.Split(':');
            var (startCol, startRow) = ParseCellReference(parts[0]);
            var (endCol, endRow) = ParseCellReference(parts[1]);
            var startColIdx = ColumnNameToIndex(startCol);
            var endColIdx = ColumnNameToIndex(endCol);
            // Clamp merge range to rendering limits to prevent memory explosion
            var clampedEndRow = Math.Min(endRow, 5000);
            var clampedEndCol = Math.Min(endColIdx, 200);
            var rowSpan = clampedEndRow - startRow + 1;
            var colSpan = clampedEndCol - startColIdx + 1;

            for (int r = startRow; r <= clampedEndRow; r++)
            {
                for (int ci = startColIdx; ci <= clampedEndCol; ci++)
                {
                    var cellRef = $"{IndexToColumnName(ci)}{r}";
                    bool isAnchor = (r == startRow && ci == startColIdx);
                    map[cellRef] = new MergeInfo(isAnchor, isAnchor ? rowSpan : 0, isAnchor ? colSpan : 0);
                }
            }
        }

        return map;
    }

    // ==================== Column Widths ====================

    private static Dictionary<int, double> GetColumnWidths(Worksheet ws)
    {
        var result = new Dictionary<int, double>();
        var columns = ws.GetFirstChild<Columns>();
        if (columns == null) return result;

        foreach (var col in columns.Elements<Column>())
        {
            if (col.Width?.Value == null) continue;
            var min = (int)(col.Min?.Value ?? 1u);
            var max = (int)(col.Max?.Value ?? (uint)min);
            // Hidden columns get width 0
            var widthPt = col.Hidden?.Value == true ? 0 : (col.Width.Value == 0 ? 0 : col.Width.Value * 5.625 + 3.75);
            for (int c = min; c <= max; c++)
                result[c] = widthPt;
        }

        return result;
    }

    // ==================== Frozen Panes ====================

    private static (int frozenRows, int frozenCols) GetFrozenPanes(Worksheet ws)
    {
        var sheetViews = ws.GetFirstChild<SheetViews>();
        var sheetView = sheetViews?.GetFirstChild<SheetView>();
        var pane = sheetView?.GetFirstChild<Pane>();
        if (pane == null) return (0, 0);

        // Only handle frozen panes (not split panes)
        if (pane.State?.Value != PaneStateValues.Frozen && pane.State?.Value != PaneStateValues.FrozenSplit)
            return (0, 0);

        var frozenRows = (int)(pane.VerticalSplit?.Value ?? 0);
        var frozenCols = (int)(pane.HorizontalSplit?.Value ?? 0);
        return (frozenRows, frozenCols);
    }

    // ==================== Conditional Formatting ====================

    /// <summary>
    /// Evaluate conditional formatting rules and return CSS overrides per cell.
    /// </summary>
    private Dictionary<string, string> BuildConditionalFormatMap(
        Worksheet ws, Stylesheet? stylesheet, SheetData sheetData, WorkbookPart? workbookPart)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (stylesheet == null) return result;

        var dxfs = stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().ToArray();
        if (dxfs == null || dxfs.Length == 0) return result;

        var cfElements = ws.Elements<ConditionalFormatting>().ToList();
        if (cfElements.Count == 0) return result;

        var evaluator = new Core.FormulaEvaluator(sheetData, workbookPart);

        foreach (var cf in cfElements)
        {
            var sqref = cf.SequenceOfReferences?.Items?.ToList();
            if (sqref == null || sqref.Count == 0) continue;

            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var dxfId = rule.FormatId?.Value;
                if (dxfId == null || dxfId >= dxfs.Length) continue;
                var dxf = dxfs[(int)dxfId];

                // Extract CSS from dxf
                var cssParts = new List<string>();
                var fill = dxf.Fill?.PatternFill;
                if (fill != null)
                {
                    var bgColor = fill.BackgroundColor?.Rgb?.Value ?? fill.ForegroundColor?.Rgb?.Value;
                    if (bgColor != null)
                    {
                        if (bgColor.Length > 6) bgColor = bgColor[^6..];
                        cssParts.Add($"background:#{bgColor}");
                    }
                }
                var font = dxf.Font;
                if (font != null)
                {
                    var fontColor = font.Color?.Rgb?.Value;
                    if (fontColor != null)
                    {
                        if (fontColor.Length > 6) fontColor = fontColor[^6..];
                        cssParts.Add($"color:#{fontColor}");
                    }
                }
                if (cssParts.Count == 0) continue;
                var cssOverride = string.Join(";", cssParts);

                // Expand sqref and evaluate each cell
                foreach (var rangeStr in sqref)
                {
                    var cells = ExpandSqref(rangeStr.Value ?? "");
                    foreach (var (cellRef, row, col) in cells)
                    {
                        if (result.ContainsKey(cellRef)) continue; // first matching rule wins

                        bool matches = EvaluateCfRule(rule, cellRef, row, col, sheetData, evaluator);
                        if (matches)
                            result[cellRef] = cssOverride;
                    }
                }
            }
        }
        return result;
    }

    /// <summary>Evaluate whether a conditional formatting rule matches a specific cell.</summary>
    private bool EvaluateCfRule(ConditionalFormattingRule rule, string cellRef, int row, int col,
        SheetData sheetData, Core.FormulaEvaluator evaluator)
    {
        var ruleType = rule.Type?.Value;

        // Get cell value for comparison
        double? cellValue = null;
        var cell = sheetData.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (cell != null)
        {
            if (double.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
                cellValue = v;
        }

        if (ruleType == ConditionalFormatValues.Expression)
        {
            // Formula-based rule: evaluate with cell reference adjustment
            var formula = rule.Elements<Formula>().FirstOrDefault()?.Text;
            if (string.IsNullOrEmpty(formula)) return false;

            // Adjust formula references relative to the first cell in sqref
            // The formula is written for the top-left cell; adjust for current cell
            var adjusted = AdjustCfFormula(formula, row, col, rule);
            var result = evaluator.TryEvaluateFull(adjusted);
            return result?.BoolValue == true || (result?.NumericValue != null && result.NumericValue != 0);
        }

        if (ruleType == ConditionalFormatValues.CellIs && cellValue.HasValue)
        {
            var op = rule.Operator?.Value;
            var f1 = rule.Elements<Formula>().FirstOrDefault()?.Text;
            var f2 = rule.Elements<Formula>().Skip(1).FirstOrDefault()?.Text;
            double? v1 = f1 != null ? evaluator.TryEvaluate(f1) ?? (double.TryParse(f1, out var p1) ? p1 : null) : null;
            double? v2 = f2 != null ? evaluator.TryEvaluate(f2) ?? (double.TryParse(f2, out var p2) ? p2 : null) : null;
            if (v1 == null) return false;
            if (op == ConditionalFormattingOperatorValues.GreaterThan) return cellValue > v1;
            if (op == ConditionalFormattingOperatorValues.LessThan) return cellValue < v1;
            if (op == ConditionalFormattingOperatorValues.GreaterThanOrEqual) return cellValue >= v1;
            if (op == ConditionalFormattingOperatorValues.LessThanOrEqual) return cellValue <= v1;
            if (op == ConditionalFormattingOperatorValues.Equal) return cellValue == v1;
            if (op == ConditionalFormattingOperatorValues.NotEqual) return cellValue != v1;
            if (op == ConditionalFormattingOperatorValues.Between) return v2.HasValue && cellValue >= v1 && cellValue <= v2;
            if (op == ConditionalFormattingOperatorValues.NotBetween) return v2.HasValue && (cellValue < v1 || cellValue > v2);
            return false;
        }

        return false;
    }

    /// <summary>Adjust a CF formula's cell references from the anchor cell to the target cell.</summary>
    private string AdjustCfFormula(string formula, int targetRow, int targetCol, ConditionalFormattingRule rule)
    {
        // Find the anchor cell from the parent ConditionalFormatting sqref
        var cf = rule.Parent as ConditionalFormatting;
        var sqref = cf?.SequenceOfReferences?.Items?.FirstOrDefault()?.Value;
        if (string.IsNullOrEmpty(sqref)) return formula;

        // Extract anchor from sqref (e.g. "E7:E21" → anchor is E7)
        var anchorRef = sqref.Contains(':') ? sqref.Split(':')[0] : sqref;
        var (anchorColName, anchorRow) = ParseCellReference(anchorRef);
        var anchorCol = ColumnNameToIndex(anchorColName);

        var rowDelta = targetRow - anchorRow;
        var colDelta = targetCol - anchorCol;
        if (rowDelta == 0 && colDelta == 0) return formula;

        // Replace cell references in formula, adjusting by delta
        return Regex.Replace(formula, @"(\$?)([A-Z]+)(\$?)(\d+)", m =>
        {
            var colAbsolute = m.Groups[1].Value == "$";
            var rowAbsolute = m.Groups[3].Value == "$";
            var refCol = ColumnNameToIndex(m.Groups[2].Value);
            var refRow = int.Parse(m.Groups[4].Value);

            var newCol = colAbsolute ? refCol : refCol + colDelta;
            var newRow = rowAbsolute ? refRow : refRow + rowDelta;
            if (newCol < 1) newCol = 1;
            if (newRow < 1) newRow = 1;
            return $"{(colAbsolute ? "$" : "")}{IndexToColumnName(newCol)}{(rowAbsolute ? "$" : "")}{newRow}";
        });
    }

    /// <summary>Expand a sqref string like "E7:E21" into individual cell references.</summary>
    private List<(string cellRef, int row, int col)> ExpandSqref(string sqref)
    {
        var result = new List<(string, int, int)>();
        foreach (var part in sqref.Split(' '))
        {
            if (part.Contains(':'))
            {
                var sides = part.Split(':');
                var (startColName, startRow) = ParseCellReference(sides[0]);
                var (endColName, endRow) = ParseCellReference(sides[1]);
                var startCol = ColumnNameToIndex(startColName);
                var endCol = ColumnNameToIndex(endColName);
                for (int r = startRow; r <= endRow; r++)
                    for (int c = startCol; c <= endCol; c++)
                        result.Add(($"{IndexToColumnName(c)}{r}", r, c));
            }
            else
            {
                var (colName, row) = ParseCellReference(part);
                result.Add((part, row, ColumnNameToIndex(colName)));
            }
        }
        return result;
    }

    // ==================== Cell Style to CSS ====================

    private string GetCellStyleCss(Cell? cell, Stylesheet? stylesheet, int frozenRows, int frozenCols, int row, int col,
        Dictionary<int, double>? frozenLeftOffsets = null, Dictionary<int, double>? frozenTopOffsets = null,
        Dictionary<string, string>? cfMap = null)
    {
        var styles = new List<string>();

        // Frozen pane sticky positioning
        bool isFrozenRow = frozenRows > 0 && row <= frozenRows;
        bool isFrozenCol = frozenCols > 0 && col <= frozenCols;
        // z-index layering: corner-cell=4, col-header=3, frozen-row+col=2, frozen-col=1
        var frozenLeft = frozenLeftOffsets?.TryGetValue(col, out var fl) == true ? fl : 0;
        var frozenTop = frozenTopOffsets?.TryGetValue(row, out var ft) == true ? ft : 0;
        if (isFrozenRow && isFrozenCol)
            styles.Add($"position:sticky;top:0;left:{frozenLeft:0.##}pt;z-index:2");
        else if (isFrozenRow)
            styles.Add("position:sticky;top:0;z-index:1");
        else if (isFrozenCol)
            styles.Add($"position:sticky;left:{frozenLeft:0.##}pt;z-index:1");

        if (cell == null || stylesheet == null)
        {
            // Frozen rows need opaque background so scrolling content doesn't show through
            if (isFrozenRow) styles.Add("background:#fff");
            return styles.Count > 0 ? $" style=\"{string.Join(";", styles)}\"" : "";
        }

        var styleIndex = cell.StyleIndex?.Value ?? 0;

        {
            var cellFormats = stylesheet.CellFormats;
            if (cellFormats != null && styleIndex < (uint)cellFormats.Elements<CellFormat>().Count())
            {
                var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
                BuildFontCss(xf, stylesheet, styles);
                BuildFillCss(xf, stylesheet, styles);
                BuildBorderCss(xf, stylesheet, styles);
                BuildAlignmentCss(xf, styles);
            }
        }

        // Conditional formatting overrides (background, color)
        var cfCellRef = $"{IndexToColumnName(col)}{row}";
        if (cfMap != null && cfMap.TryGetValue(cfCellRef, out var cfCss))
        {
            // CF overrides existing background/color — remove conflicting base styles
            foreach (var cfPart in cfCss.Split(';'))
            {
                var prop = cfPart.Split(':')[0].Trim();
                styles.RemoveAll(s => s.StartsWith(prop + ":"));
            }
            styles.Add(cfCss);
        }

        // Frozen rows need opaque background so scrolling content doesn't show through
        if (isFrozenRow && !styles.Any(s => s.StartsWith("background:")))
            styles.Add("background:#fff");

        return styles.Count > 0 ? $" style=\"{string.Join(";", styles)}\"" : "";
    }

    private static void BuildFontCss(CellFormat xf, Stylesheet stylesheet, List<string> styles)
    {
        var fontId = xf.FontId?.Value ?? 0;
        var fonts = stylesheet.Fonts;
        if (fonts == null || fontId >= (uint)fonts.Elements<Font>().Count()) return;

        var font = fonts.Elements<Font>().ElementAt((int)fontId);

        if (font.Bold != null && font.Bold.Val?.Value != false) styles.Add("font-weight:bold");
        if (font.Italic != null && font.Italic.Val?.Value != false) styles.Add("font-style:italic");
        if (font.Strike != null && font.Strike.Val?.Value != false) styles.Add("text-decoration:line-through");
        if (font.Underline != null)
        {
            var existing = styles.FindIndex(s => s.StartsWith("text-decoration:"));
            if (existing >= 0)
                styles[existing] = styles[existing] + " underline";
            else
                styles.Add("text-decoration:underline");
        }

        // Superscript/Subscript via VerticalTextAlignment
        var vertAlign = font.GetFirstChild<VerticalTextAlignment>();
        if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Superscript)
            styles.Add("vertical-align:super;font-size:smaller");
        else if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Subscript)
            styles.Add("vertical-align:sub;font-size:smaller");

        if (font.FontSize?.Val?.Value != null)
            styles.Add($"font-size:{font.FontSize.Val.Value:0.##}pt");

        if (font.FontName?.Val?.Value != null)
            styles.Add($"font-family:'{CssSanitize(font.FontName.Val.Value)}'");

        var color = ResolveFontColor(font);
        if (color != null) styles.Add($"color:{color}");
    }

    private static void BuildFillCss(CellFormat xf, Stylesheet stylesheet, List<string> styles)
    {
        var fillId = xf.FillId?.Value ?? 0;
        if (fillId <= 1) return; // 0=none, 1=gray125 pattern (default)

        var fills = stylesheet.Fills;
        if (fills == null || fillId >= (uint)fills.Elements<Fill>().Count()) return;

        var fill = fills.Elements<Fill>().ElementAt((int)fillId);

        // Gradient fill
        var gf = fill.GetFirstChild<GradientFill>();
        if (gf != null)
        {
            var stops = gf.Elements<GradientStop>().ToList();
            if (stops.Count >= 2)
            {
                var colors = stops
                    .Select(s => ResolveColorRgb(s.Color))
                    .Where(c => c != null)
                    .ToList();
                if (colors.Count >= 2)
                {
                    var deg = (int)(gf.Degree?.Value ?? 0);
                    styles.Add($"background:linear-gradient({deg}deg,{string.Join(",", colors)})");
                    return;
                }
            }
        }

        // Pattern fill
        var pf = fill.PatternFill;
        if (pf != null)
        {
            var bgColor = ResolveColorRgb(pf.ForegroundColor);
            if (bgColor != null) styles.Add($"background:{bgColor}");
        }
    }

    private static void BuildBorderCss(CellFormat xf, Stylesheet stylesheet, List<string> styles)
    {
        var borderId = xf.BorderId?.Value ?? 0;
        if (borderId == 0) return;

        var borders = stylesheet.Borders;
        if (borders == null || borderId >= (uint)borders.Elements<Border>().Count()) return;

        var border = borders.Elements<Border>().ElementAt((int)borderId);

        AddBorderSideCss(border.TopBorder, "top", styles);
        AddBorderSideCss(border.RightBorder, "right", styles);
        AddBorderSideCss(border.BottomBorder, "bottom", styles);
        AddBorderSideCss(border.LeftBorder, "left", styles);
    }

    private static void AddBorderSideCss(BorderPropertiesType? bp, string side, List<string> styles)
    {
        if (bp?.Style?.Value == null || bp.Style.Value == BorderStyleValues.None) return;

        var bsv = bp.Style.Value;
        var width = "1px";
        if (bsv == BorderStyleValues.Medium) width = "2px";
        else if (bsv == BorderStyleValues.Thick) width = "3px";
        else if (bsv == BorderStyleValues.Double) width = "3px";

        var cssStyle = "solid";
        if (bsv == BorderStyleValues.Dashed || bsv == BorderStyleValues.MediumDashed) cssStyle = "dashed";
        else if (bsv == BorderStyleValues.Dotted) cssStyle = "dotted";
        else if (bsv == BorderStyleValues.Double) cssStyle = "double";

        var color = ResolveColorRgb(bp.Color);
        color ??= "#000";

        styles.Add($"border-{side}:{width} {cssStyle} {color}");
    }

    private static void BuildAlignmentCss(CellFormat xf, List<string> styles)
    {
        var alignment = xf.Alignment;
        if (alignment == null) return;

        if (alignment.Horizontal?.HasValue == true)
        {
            var h = alignment.Horizontal.InnerText;
            var cssAlign = h switch
            {
                "center" => "center",
                "right" => "right",
                "left" => "left",
                "justify" => "justify",
                "fill" => "left",
                _ => null
            };
            if (cssAlign != null) styles.Add($"text-align:{cssAlign}");
        }

        if (alignment.Vertical?.HasValue == true)
        {
            var v = alignment.Vertical.InnerText;
            var cssVAlign = v switch
            {
                "top" => "top",
                "center" => "middle",
                "bottom" => "bottom",
                _ => null
            };
            if (cssVAlign != null) styles.Add($"vertical-align:{cssVAlign}");
        }

        if (alignment.WrapText?.Value == true)
            styles.Add("white-space:pre-wrap;word-wrap:break-word");

        if (alignment.TextRotation?.HasValue == true && alignment.TextRotation.Value != 0)
        {
            var rot = alignment.TextRotation.Value;
            if (rot == 255)
            {
                // 255 = stacked vertical text (each char on its own line)
                styles.Add("writing-mode:vertical-rl;text-orientation:upright;letter-spacing:-2px");
            }
            else
            {
                // Excel: 0-90 = counter-clockwise, 91-180 = clockwise (91=1°CW, 180=90°CW)
                // Excel: 1-90 = CCW (CSS negative), 91-180 = CW (CSS positive, 91=1°, 180=90°)
                int cssDeg = rot <= 90 ? -(int)rot : (int)rot - 90;
                styles.Add($"transform:rotate({cssDeg}deg);white-space:nowrap");
            }
        }

        if (alignment.Indent?.HasValue == true && alignment.Indent.Value > 0)
            styles.Add($"padding-left:{alignment.Indent.Value * 6}pt");

        // Reading order: 1=LTR, 2=RTL (for mixed-direction content)
        if (alignment.ReadingOrder?.HasValue == true)
        {
            var ro = alignment.ReadingOrder.Value;
            if (ro == 2) styles.Add("direction:rtl;unicode-bidi:embed");
            else if (ro == 1) styles.Add("direction:ltr;unicode-bidi:embed");
        }
    }

    // ==================== Color Resolution ====================

    private static string? ResolveFontColor(Font font)
    {
        if (font.Color?.Rgb?.Value != null)
        {
            var raw = font.Color.Rgb.Value;
            return FormatColorForCss(raw);
        }
        if (font.Color?.Theme?.Value != null)
        {
            // Theme 0=lt1 (usually white bg), 1=dk1 (usually black text)
            // For HTML preview, map common theme colors
            return font.Color.Theme.Value switch
            {
                0 => "#FFFFFF",
                1 => "#000000",
                _ => null // skip unresolved theme colors — will use default
            };
        }
        return null;
    }

    // Standard Excel indexed color palette (first 64 colors)
    private static readonly string[] IndexedColors = [
        "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
        "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
        "#800000","#008000","#000080","#808000","#800080","#008080","#C0C0C0","#808080",
        "#9999FF","#993366","#FFFFCC","#CCFFFF","#660066","#FF8080","#0066CC","#CCCCFF",
        "#000080","#FF00FF","#FFFF00","#00FFFF","#800080","#800000","#008080","#0000FF",
        "#00CCFF","#CCFFFF","#CCFFCC","#FFFF99","#99CCFF","#FF99CC","#CC99FF","#FFCC99",
        "#3366FF","#33CCCC","#99CC00","#FFCC00","#FF9900","#FF6600","#666699","#969696",
        "#003366","#339966","#003300","#333300","#993300","#993366","#333399","#333333"
    ];

    private static string? ResolveColorRgb(ColorType? color)
    {
        if (color?.Rgb?.Value != null)
            return FormatColorForCss(color.Rgb.Value);
        if (color?.Indexed?.Value != null)
        {
            var idx = (int)color.Indexed.Value;
            if (idx >= 0 && idx < IndexedColors.Length)
                return IndexedColors[idx];
            if (idx == 64) return null; // system foreground (context dependent)
            if (idx == 65) return null; // system background
        }
        if (color?.Theme?.Value != null)
        {
            return color.Theme.Value switch
            {
                0 => "#FFFFFF", // lt1
                1 => "#000000", // dk1
                2 => "#E7E6E6", // lt2
                3 => "#44546A", // dk2
                4 => "#4472C4", // accent1
                5 => "#ED7D31", // accent2
                6 => "#A5A5A5", // accent3
                7 => "#FFC000", // accent4
                8 => "#5B9BD5", // accent5
                9 => "#70AD47", // accent6
                _ => null
            };
        }
        return null;
    }

    private static string FormatColorForCss(string raw)
    {
        // ARGB "FFFF0000" → "#FF0000", or 6-char hex
        if (raw.Length == 8)
            return "#" + raw[2..];
        if (raw.Length == 6)
            return "#" + raw;
        return "#" + raw;
    }

    // ==================== Formatted Cell Value ====================

    /// <summary>
    /// Get cell display value with number formatting applied for HTML preview.
    /// Handles common formats: percentage, thousands separator, decimal places, dates.
    /// </summary>
    private string GetFormattedCellValue(Cell cell, Stylesheet? stylesheet, Core.FormulaEvaluator? evaluator = null)
    {
        var rawValue = GetCellDisplayValue(cell);

        // If the cell has a formula, always try to evaluate (cached values may be stale)
        if (cell.CellFormula?.Text != null && evaluator != null)
        {
            var result = evaluator.TryEvaluateFull(cell.CellFormula.Text);
            if (result != null)
            {
                if (result.IsError) return result.ErrorValue!;
                rawValue = result.ToCellValueText();
                if (result.IsString) return rawValue;
                if (result.IsBool) return result.BoolValue!.Value ? "TRUE" : "FALSE";
            }
            // If evaluation fails (null), fall through to use cached value / raw display
        }

        if (string.IsNullOrEmpty(rawValue)) return rawValue;

        // Boolean: convert 1/0 to TRUE/FALSE
        if (cell.DataType?.Value == CellValues.Boolean)
            return rawValue == "1" ? "TRUE" : "FALSE";

        // Only format numeric values (not strings, shared strings, etc.)
        if (cell.DataType?.Value == CellValues.SharedString ||
            cell.DataType?.Value == CellValues.InlineString ||
            cell.DataType?.Value == CellValues.String ||
            cell.DataType?.Value == CellValues.Error)
            return rawValue;

        if (!double.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
            return rawValue;

        // Clean up floating point artifacts for display (e.g. 25300000.000000004 → 25300000)
        var cleanVal = numVal;
        var rounded = Math.Round(numVal, 10);
        if (Math.Abs(rounded - Math.Round(rounded)) < 1e-9)
            cleanVal = Math.Round(rounded);
        rawValue = cleanVal == numVal ? rawValue
            : cleanVal.ToString(System.Globalization.CultureInfo.InvariantCulture);

        // Look up number format
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        if (styleIndex == 0 || stylesheet == null) return rawValue;

        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= (uint)cellFormats.Elements<CellFormat>().Count())
            return rawValue;

        var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var numFmtId = xf.NumberFormatId?.Value ?? 0;
        if (numFmtId == 0) return rawValue;

        // Resolve format code
        string? fmtCode = null;
        var customFmt = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
        if (customFmt?.FormatCode?.Value != null)
            fmtCode = customFmt.FormatCode.Value;
        else
            fmtCode = ResolveBuiltInFormat(numFmtId);

        if (fmtCode == null) return rawValue;

        return ApplyNumberFormat(numVal, fmtCode);
    }

    private static string? ResolveBuiltInFormat(uint numFmtId) => numFmtId switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        14 => "m/d/yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;(#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;(#,##0.00)",
        49 => "@",
        _ => null
    };

    private static string ApplyNumberFormat(double value, string fmtCode)
    {
        // Handle multi-section format codes: positive;negative;zero
        if (fmtCode.Contains(';'))
        {
            var sections = fmtCode.Split(';');
            if (value < 0 && sections.Length >= 2)
            {
                var negFmt = sections[1].Trim();
                // If format already handles negative (has parens or minus), don't add extra minus
                return ApplyNumberFormat(Math.Abs(value), negFmt);
            }
            if (value == 0 && sections.Length >= 3)
            {
                var zeroFmt = sections[2].Trim();
                // Quoted literal for zero section: "zero" → zero
                if (zeroFmt.StartsWith('"') && zeroFmt.EndsWith('"'))
                    return zeroFmt[1..^1];
                return ApplyNumberFormat(value, zeroFmt);
            }
            fmtCode = sections[0].Trim();
        }

        // Strip [Color] markers: [Red], [Blue], [Green], [Color N], etc.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[(Red|Blue|Green|Yellow|White|Black|Cyan|Magenta|Color\s*\d+)\]", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();

        // Strip [$...] locale/currency specifiers (e.g. [$-409], [$€-407], [$¥-411])
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[\$[^\]]*\]", "").Trim();

        // Strip Excel numfmt special characters:
        // _X = space placeholder, *X = fill character, \X = literal character escape
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"_.", "").Trim();
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\*.", "").Trim();
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\\(.)", "$1").Trim();

        // Strip condition markers: [>100], [<=0], etc.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[[<>=!]+\d+\.?\d*\]", "").Trim();

        // Handle parenthesis wrapping: ($#,##0.00) → prefix="(" suffix=")"
        if (fmtCode.StartsWith('(') && fmtCode.EndsWith(')'))
        {
            var inner = fmtCode[1..^1];
            return "(" + ApplyNumberFormat(value, inner) + ")";
        }

        var fmt = fmtCode.ToLowerInvariant();

        // Extract currency/text prefix and suffix (e.g. "$", "€", "¥", or quoted strings like "USD ")
        var prefix = "";
        var suffix = "";
        var cleanFmt = fmtCode;
        // Handle literal characters: $, ¥, €, £
        foreach (var sym in new[] { "$", "¥", "€", "£", "₹" })
        {
            if (cleanFmt.Contains(sym))
            {
                var idx = cleanFmt.IndexOf(sym);
                var hashIdx = cleanFmt.IndexOf('#');
                var zeroIdx = cleanFmt.IndexOf('0');
                var firstDigit = (hashIdx >= 0 && zeroIdx >= 0) ? Math.Min(hashIdx, zeroIdx)
                    : Math.Max(hashIdx, zeroIdx);
                if (firstDigit < 0 || idx <= firstDigit)
                    prefix = sym;
                else
                    suffix = sym;
                cleanFmt = cleanFmt.Replace(sym, "");
            }
        }
        // Handle quoted prefix/suffix: "USD "
        var quoteMatch = System.Text.RegularExpressions.Regex.Match(cleanFmt, "^\"([^\"]+)\"");
        if (quoteMatch.Success) { prefix += quoteMatch.Groups[1].Value; cleanFmt = cleanFmt[quoteMatch.Length..]; }
        var quoteSuffix = System.Text.RegularExpressions.Regex.Match(cleanFmt, "\"([^\"]+)\"$");
        if (quoteSuffix.Success) { suffix = quoteSuffix.Groups[1].Value + suffix; cleanFmt = cleanFmt[..^quoteSuffix.Length]; }

        // Handle +/- prefix in format (e.g. "+0.0%", "-#,##0")
        cleanFmt = cleanFmt.Trim();
        if (cleanFmt.StartsWith('+'))
        { prefix += "+"; cleanFmt = cleanFmt[1..]; }
        else if (cleanFmt.StartsWith('-'))
        { prefix += "-"; cleanFmt = cleanFmt[1..]; }

        var formatted = ApplyNumberFormatCore(value, cleanFmt.Trim());
        return prefix + formatted + suffix;
    }

    private static string ApplyNumberFormatCore(double value, string fmtCode)
    {
        var fmt = fmtCode.ToLowerInvariant();

        // Percentage formats
        if (fmt.Contains('%'))
        {
            var pctVal = value * 100;
            var decimals = CountDecimalPlaces(fmtCode);
            return pctVal.ToString($"F{decimals}") + "%";
        }

        // Elapsed time format: [h]:mm:ss or [mm]:ss (total hours/minutes, can exceed 24/60)
        var elapsedMatch = System.Text.RegularExpressions.Regex.Match(fmtCode, @"\[(h+)\]:?(mm)?:?(ss)?");
        if (elapsedMatch.Success)
        {
            var totalHours = (int)(value * 24);
            var totalMinutes = (int)(value * 24 * 60) % 60;
            var totalSeconds = (int)(value * 24 * 3600) % 60;
            var parts = new List<string> { totalHours.ToString() };
            if (elapsedMatch.Groups[2].Success) parts.Add(totalMinutes.ToString("D2"));
            if (elapsedMatch.Groups[3].Success) parts.Add(totalSeconds.ToString("D2"));
            return string.Join(":", parts);
        }

        // Date formats (serial number → DateTime)
        if (fmt.Contains('y') || fmt.Contains('m') || fmt.Contains('d') || fmt.Contains('h'))
        {
            try
            {
                var dt = DateTime.FromOADate(value);
                // Context-sensitive m/mm: after h → minute, otherwise → month
                // Strategy: mark minute 'm' as '\x01' placeholder, then convert remaining m→M
                var dotnetFmt = fmtCode
                    .Replace("AM/PM", "tt").Replace("am/pm", "tt");
                // Step 1: Replace h:mm and h:m patterns → mark minutes as placeholder
                dotnetFmt = System.Text.RegularExpressions.Regex.Replace(dotnetFmt, @"([hH]+)([:.])(mm?)", m =>
                    m.Groups[1].Value + m.Groups[2].Value + new string('\x01', m.Groups[3].Value.Length));
                // Also handle mm:ss (mm before ss is also minutes)
                dotnetFmt = System.Text.RegularExpressions.Regex.Replace(dotnetFmt, @"(mm?)([:.])(ss?)", m =>
                    new string('\x01', m.Groups[1].Value.Length) + m.Groups[2].Value + m.Groups[3].Value);
                // Step 2: Convert remaining m/mm to M/MM (month)
                dotnetFmt = dotnetFmt.Replace("mmmm", "MMMM").Replace("mmm", "MMM")
                    .Replace("mm", "MM").Replace("m", "M");
                // Step 3: Restore minute placeholders
                dotnetFmt = dotnetFmt.Replace("\x01\x01", "mm").Replace("\x01", "m");
                // Step 4: Other conversions
                // If AM/PM format (has 'tt'), use h (12h); otherwise use H (24h)
                if (!dotnetFmt.Contains("tt"))
                    dotnetFmt = dotnetFmt.Replace("hh", "HH").Replace("h", "H");
                dotnetFmt = dotnetFmt.Replace("dddd", "dddd").Replace("ddd", "ddd").Replace("dd", "dd");
                return dt.ToString(dotnetFmt, System.Globalization.CultureInfo.InvariantCulture);
            }
            catch { return value.ToString(); }
        }

        // Scientific notation
        if (fmt.Contains("e+") || fmt.Contains("e-"))
        {
            var decimals = CountDecimalPlaces(fmtCode);
            if (value == 0) return decimals > 0 ? $"0.{new string('0', decimals)}E+00" : "0E+00";
            var eIdx = fmt.IndexOf("e+", StringComparison.Ordinal);
            if (eIdx < 0) eIdx = fmt.IndexOf("e-", StringComparison.Ordinal);
            var expDigits = eIdx >= 0 ? fmtCode[(eIdx + 2)..].Count(c => c == '0') : 2;
            var exp = (int)Math.Floor(Math.Log10(Math.Abs(value)));
            var mantissa = value / Math.Pow(10, exp);
            var expStr = exp >= 0 ? $"+{exp.ToString().PadLeft(expDigits, '0')}" : $"-{Math.Abs(exp).ToString().PadLeft(expDigits, '0')}";
            return $"{mantissa.ToString($"F{decimals}")}E{expStr}";
        }

        // Trailing comma scaling: each trailing comma divides value by 1000
        // e.g. "#," = ÷1000, "#,," = ÷1000000, "#,##0," = thousands + ÷1000
        var trailingCommas = 0;
        var fmtTrimmed = fmtCode.TrimEnd();
        while (fmtTrimmed.EndsWith(',')) { trailingCommas++; fmtTrimmed = fmtTrimmed[..^1]; }
        if (trailingCommas > 0)
        {
            value /= Math.Pow(1000, trailingCommas);
            fmtCode = fmtTrimmed;
        }

        // Numeric with thousands separator and/or decimals
        bool hasThousands = fmtCode.Contains(',') && (fmtCode.Contains('#') || fmtCode.Contains('0'));
        var numDecimals = CountDecimalPlaces(fmtCode);

        if (hasThousands)
            return value.ToString($"N{numDecimals}", System.Globalization.CultureInfo.InvariantCulture);
        if (numDecimals > 0)
            return value.ToString($"F{numDecimals}");

        // @ = text format — return raw
        if (fmt == "@") return value.ToString();

        // Integer format "0"
        if (fmtCode.Trim() == "0") return ((long)Math.Round(value)).ToString();

        return value.ToString();
    }

    private static int CountDecimalPlaces(string fmtCode)
    {
        var dotIdx = fmtCode.IndexOf('.');
        if (dotIdx < 0) return 0;
        int count = 0;
        for (int i = dotIdx + 1; i < fmtCode.Length; i++)
        {
            if (fmtCode[i] == '0' || fmtCode[i] == '#') count++;
            else break;
        }
        return count;
    }

    // ==================== CSS ====================

    private static string GenerateExcelCss() => """
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { height: 100%; }
        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            background: #f0f0f0;
            color: #333;
            display: flex;
            flex-direction: column;
            min-height: 100vh;
        }
        .file-title {
            padding: 12px 20px;
            font-size: 14px;
            font-weight: 600;
            background: #217346;
            color: #fff;
        }
        .sheet-tabs {
            display: flex;
            background: #e0e0e0;
            border-top: 1px solid #ccc;
            overflow-x: auto;
            padding: 0 8px;
            flex-shrink: 0;
            position: sticky;
            bottom: 0;
            z-index: 10;
        }
        .sheet-tab {
            --tab-color: #e8e8e8;
            padding: 8px 16px;
            font-size: 12px;
            cursor: pointer;
            border: 1px solid #bbb;
            border-top: none;
            background: var(--tab-color);
            color: #fff;
            margin-bottom: 0;
            border-radius: 0 0 3px 3px;
            white-space: nowrap;
            user-select: none;
            position: relative;
            transition: background 0.15s, color 0.15s;
        }
        .sheet-tab[style*="--tab-color:#e8e8e8"], .sheet-tab:not([style*="--tab-color"]) {
            color: #333;
        }
        .sheet-tab:hover { opacity: 0.85; }
        .sheet-tab.active {
            background: linear-gradient(to bottom, #fff 60%, color-mix(in srgb, var(--tab-color) 30%, #fff)) !important;
            color: #333 !important;
            border-color: #aaa;
            border-bottom: 3px solid var(--tab-color);
            font-weight: 600;
        }
        .sheet-slider { flex: 1; position: relative; overflow: hidden; display: flex; flex-direction: column; min-height: 0; }
        .sheet-content { background: #fff; display: none; flex: 1; min-height: 0; }
        .sheet-content.active { display: flex; flex-direction: column; }
        .table-wrapper {
            flex: 1;
            overflow: auto;
            min-height: 0;
            background: #fff;
        }
        table {
            border-collapse: collapse;
            font-size: 11px;
            font-family: 'Calibri', 'Segoe UI', sans-serif;
            table-layout: fixed;
        }
        .row-header-col { width: 30pt; }
        th {
            background: #f8f8f8;
            border: 1px solid #e0e0e0;
            font-weight: normal;
            color: #666;
            font-size: 10px;
            text-align: center;
            padding: 2px 4px;
        }
        .corner-cell { background: #f0f0f0; z-index: 4; }
        .col-header {
            position: sticky;
            top: 0;
            z-index: 3;
            background: #f8f8f8;
            min-width: 50px;
        }
        .row-header {
            position: sticky;
            left: 0;
            z-index: 2;
            background: #f8f8f8;
            min-width: 40px;
        }
        td {
            border: 1px solid #e0e0e0;
            padding: 2px 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            vertical-align: bottom;
            max-width: 500px;
            word-break: break-all; /* CJK text wrapping support */
        }
        .empty-sheet {
            padding: 40px;
            text-align: center;
            color: #999;
            font-size: 14px;
        }
        /* Frozen pane visual separator */
        tr:nth-child(1) td { border-top-color: #e0e0e0; }
        /* Chart containers */
        .chart-container {
            margin: 16px auto;
            background: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            padding: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .chart-container svg { display: block; }
        /* Truncation warning */
        .truncation-warning {
            padding: 8px 16px;
            background: #FFF3CD;
            color: #856404;
            border: 1px solid #FFEEBA;
            font-size: 12px;
            text-align: center;
            margin: 4px 0;
        }
        /* Screen reader only */
        .sr-only { position:absolute; clip:rect(0 0 0 0); width:1px; height:1px; overflow:hidden; }
        /* Print styles */
        @media print {
            .file-title, .sheet-tabs { display: none !important; }
            .table-wrapper { max-height: none !important; overflow: visible !important; flex: none !important; }
            body { background: #fff !important; min-height: auto !important; }
            .sheet-content { display: block !important; flex: none !important; }
            td { max-width: none !important; white-space: normal !important; overflow: visible !important; }
        }
        """;

    // ==================== JavaScript ====================

    private static string GenerateExcelJs() => """
        function switchSheet(idx) {
            document.querySelectorAll('.sheet-tab').forEach(function(t) {
                t.classList.toggle('active', parseInt(t.getAttribute('data-sheet')) === idx);
            });
            document.querySelectorAll('.sheet-content').forEach(function(c) {
                c.classList.toggle('active', parseInt(c.getAttribute('data-sheet')) === idx);
            });
            window.scrollTo(0, 0);
        }
        // Fix frozen row sticky top values using actual rendered heights
        document.querySelectorAll('.table-wrapper table').forEach(function(table) {
            var thead = table.querySelector('thead');
            if (!thead) return;
            var theadH = thead.offsetHeight;
            var cumTop = theadH;
            var frozen = table.querySelectorAll('tr[data-frozen]');
            frozen.forEach(function(tr) {
                tr.querySelectorAll('th, td').forEach(function(cell) {
                    if (cell.style.position === 'sticky') cell.style.top = cumTop + 'px';
                });
                cumTop += tr.offsetHeight;
            });
        });
        """;

    // ==================== Utility ====================

    private static string HtmlEncode(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }

    /// <summary>HtmlEncode + convert newlines to br for cell display</summary>
    private static string CellHtml(string text)
    {
        var encoded = HtmlEncode(text);
        return encoded.Contains('\n') ? encoded.Replace("\n", "<br>") : encoded;
    }

    private static string CssSanitize(string value)
    {
        // Strip characters that could break CSS context
        return Regex.Replace(value, @"[;:{}()\\""']", "");
    }

}
