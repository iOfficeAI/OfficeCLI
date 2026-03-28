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
        sb.AppendLine("<html lang=\"en\">");
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
        for (int sheetIdx = 0; sheetIdx < sheets.Count; sheetIdx++)
        {
            var (sheetName, worksheetPart) = sheets[sheetIdx];
            var displayStyle = sheetIdx == 0 ? "" : " style=\"display:none\"";
            sb.AppendLine($"<div class=\"sheet-content\" data-sheet=\"{sheetIdx}\"{displayStyle}>");
            RenderSheetTable(sb, sheetName, worksheetPart, stylesheet);
            RenderSheetCharts(sb, worksheetPart);
            sb.AppendLine("</div>");
        }

        // Sheet tabs at bottom (like real Excel)
        sb.AppendLine("<div class=\"sheet-tabs\">");
        for (int i = 0; i < sheets.Count; i++)
        {
            var activeClass = i == 0 ? " active" : "";
            sb.AppendLine($"  <div class=\"sheet-tab{activeClass}\" data-sheet=\"{i}\" onclick=\"switchSheet({i})\">{HtmlEncode(sheets[i].Name)}</div>");
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

    // ==================== Sheet Rendering ====================

    private void RenderSheetTable(StringBuilder sb, string sheetName, WorksheetPart worksheetPart, Stylesheet? stylesheet)
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

        // Collect merge info
        var mergeMap = BuildMergeMap(ws);

        // Collect column widths
        var colWidths = GetColumnWidths(ws);

        // Detect frozen panes
        var (frozenRows, frozenCols) = GetFrozenPanes(ws);

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
            if (worksheetPart.DrawingsPart?.WorksheetDrawing == null)
                sb.AppendLine("<div class=\"empty-sheet\">Empty sheet</div>");
            return;
        }

        // Limit rendering to reasonable size
        maxRow = Math.Min(maxRow, 5000);
        maxCol = Math.Min(maxCol, 200);

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

        // Row height lookup
        var rowHeights = new Dictionary<int, double>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (row.CustomHeight?.Value == true && row.Height?.Value != null)
                rowHeights[rowIdx] = row.Height.Value;
        }

        // Start table
        sb.AppendLine("<div class=\"table-wrapper\">");
        sb.AppendLine("<table>");

        // Colgroup for column widths + header column
        sb.Append("<colgroup><col class=\"row-header-col\">");
        for (int c = 1; c <= maxCol; c++)
        {
            var width = colWidths.TryGetValue(c, out var w) ? w : 64.0; // default ~8.43 chars ≈ 64px
            sb.Append($"<col style=\"width:{width:0.#}px\">");
        }
        sb.AppendLine("</colgroup>");

        // Column header row
        sb.Append("<thead><tr><th class=\"corner-cell\"");
        if (frozenRows > 0 || frozenCols > 0) sb.Append(" style=\"position:sticky;top:0;left:0;z-index:4\"");
        sb.Append("></th>");
        for (int c = 1; c <= maxCol; c++)
        {
            var colName = IndexToColumnName(c);
            var stickyStyle = frozenRows > 0 ? " style=\"position:sticky;top:0;z-index:3\"" : "";
            sb.Append($"<th class=\"col-header\"{stickyStyle}>{colName}</th>");
        }
        sb.AppendLine("</tr></thead>");

        // Data rows
        sb.AppendLine("<tbody>");
        for (int r = 1; r <= maxRow; r++)
        {
            var rowH = rowHeights.TryGetValue(r, out var rh) ? $" style=\"height:{rh * 1.33:0.#}px\"" : "";
            sb.Append($"<tr{rowH}>");

            // Row header
            var rowHeaderSticky = frozenCols > 0 ? " style=\"position:sticky;left:0;z-index:2\"" : "";
            sb.Append($"<th class=\"row-header\"{rowHeaderSticky}>{r}</th>");

            for (int c = 1; c <= maxCol; c++)
            {
                // Check if this cell is hidden by a merge (non-anchor cell in merged range)
                var cellRef = $"{IndexToColumnName(c)}{r}";
                if (mergeMap.TryGetValue(cellRef, out var mergeInfo))
                {
                    if (!mergeInfo.IsAnchor) continue; // skip non-anchor cells

                    var cell = cellMap.TryGetValue((r, c), out var mc) ? mc : null;
                    var style = GetCellStyleCss(cell, stylesheet, frozenRows, frozenCols, r, c);
                    var value = cell != null ? GetFormattedCellValue(cell, stylesheet) : "";
                    var spanAttrs = "";
                    if (mergeInfo.ColSpan > 1) spanAttrs += $" colspan=\"{mergeInfo.ColSpan}\"";
                    if (mergeInfo.RowSpan > 1) spanAttrs += $" rowspan=\"{mergeInfo.RowSpan}\"";

                    sb.Append($"<td{spanAttrs}{style}>{CellHtml(value)}</td>");
                }
                else
                {
                    var cell = cellMap.TryGetValue((r, c), out var nc) ? nc : null;
                    var style = GetCellStyleCss(cell, stylesheet, frozenRows, frozenCols, r, c);
                    var value = cell != null ? GetFormattedCellValue(cell, stylesheet) : "";
                    sb.Append($"<td{style}>{CellHtml(value)}</td>");
                }
            }
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</tbody>");
        sb.AppendLine("</table>");
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
            var rowSpan = endRow - startRow + 1;
            var colSpan = endColIdx - startColIdx + 1;

            for (int r = startRow; r <= endRow; r++)
            {
                for (int ci = startColIdx; ci <= endColIdx; ci++)
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
            // Excel column width is in character units; convert to approximate pixels (1 char ≈ 7.5px + 5px padding)
            var widthPx = col.Width.Value == 0 ? 0 : col.Width.Value * 7.5 + 5;
            var min = (int)(col.Min?.Value ?? 1u);
            var max = (int)(col.Max?.Value ?? (uint)min);
            for (int c = min; c <= max; c++)
                result[c] = widthPx;
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

    // ==================== Cell Style to CSS ====================

    private string GetCellStyleCss(Cell? cell, Stylesheet? stylesheet, int frozenRows, int frozenCols, int row, int col)
    {
        var styles = new List<string>();

        // Frozen pane sticky positioning
        bool isFrozenRow = frozenRows > 0 && row <= frozenRows;
        bool isFrozenCol = frozenCols > 0 && col <= frozenCols;
        if (isFrozenRow && isFrozenCol)
            styles.Add("position:sticky;top:0;left:0;z-index:3");
        else if (isFrozenRow)
            styles.Add("position:sticky;top:0;z-index:2");
        else if (isFrozenCol)
            styles.Add("position:sticky;left:0;z-index:2");

        if (cell == null || stylesheet == null)
            return styles.Count > 0 ? $" style=\"{string.Join(";", styles)}\"" : "";

        var styleIndex = cell.StyleIndex?.Value ?? 0;
        if (styleIndex == 0 && styles.Count == 0) return "";

        if (styleIndex > 0)
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
            // Excel: 0-90 = counter-clockwise, 91-180 = clockwise (91=1°CW, 180=90°CW)
            int cssDeg = rot <= 90 ? -(int)rot : 90 - (int)rot;
            styles.Add($"writing-mode:vertical-lr;transform:rotate({cssDeg}deg)");
        }

        if (alignment.Indent?.HasValue == true && alignment.Indent.Value > 0)
            styles.Add($"padding-left:{alignment.Indent.Value * 8}px");
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

    private static string? ResolveColorRgb(ColorType? color)
    {
        if (color?.Rgb?.Value != null)
            return FormatColorForCss(color.Rgb.Value);
        if (color?.Theme?.Value != null)
        {
            return color.Theme.Value switch
            {
                0 => "#FFFFFF",
                1 => "#000000",
                4 => "#4472C4",
                5 => "#ED7D31",
                6 => "#A5A5A5",
                7 => "#FFC000",
                8 => "#5B9BD5",
                9 => "#70AD47",
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
    private string GetFormattedCellValue(Cell cell, Stylesheet? stylesheet)
    {
        var rawValue = GetCellDisplayValue(cell);
        if (string.IsNullOrEmpty(rawValue)) return rawValue;

        // Only format numeric values (not strings, shared strings, etc.)
        if (cell.DataType?.Value == CellValues.SharedString ||
            cell.DataType?.Value == CellValues.InlineString ||
            cell.DataType?.Value == CellValues.String ||
            cell.DataType?.Value == CellValues.Boolean ||
            cell.DataType?.Value == CellValues.Error)
            return rawValue;

        if (!double.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
            return rawValue;

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
        var fmt = fmtCode.ToLowerInvariant();

        // Percentage formats
        if (fmt.Contains('%'))
        {
            var pctVal = value * 100;
            var decimals = CountDecimalPlaces(fmtCode);
            return pctVal.ToString($"F{decimals}") + "%";
        }

        // Date formats (serial number → DateTime)
        if (fmt.Contains('y') || fmt.Contains('m') || fmt.Contains('d') || fmt.Contains('h'))
        {
            try
            {
                var dt = DateTime.FromOADate(value);
                var dotnetFmt = fmtCode
                    .Replace("yyyy", "yyyy").Replace("yy", "yy")
                    .Replace("mmmm", "MMMM").Replace("mmm", "MMM").Replace("mm", "MM").Replace("m", "M")
                    .Replace("dddd", "dddd").Replace("ddd", "ddd").Replace("dd", "dd").Replace("d", "d")
                    .Replace("hh", "HH").Replace("h", "H")
                    .Replace("ss", "ss").Replace("s", "s")
                    .Replace("AM/PM", "tt").Replace("am/pm", "tt");
                return dt.ToString(dotnetFmt, System.Globalization.CultureInfo.InvariantCulture);
            }
            catch { return value.ToString(); }
        }

        // Scientific notation
        if (fmt.Contains("e+") || fmt.Contains("e-"))
        {
            var decimals = CountDecimalPlaces(fmtCode);
            return value.ToString($"E{decimals}");
        }

        // Numeric with thousands separator and/or decimals
        bool hasThousands = fmtCode.Contains(',') && fmtCode.Contains('#');
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
        }
        .sheet-tab {
            padding: 8px 16px;
            font-size: 12px;
            cursor: pointer;
            border: 1px solid transparent;
            border-top: none;
            background: #e8e8e8;
            margin-bottom: 4px;
            border-radius: 0 0 4px 4px;
            white-space: nowrap;
            user-select: none;
        }
        .sheet-tab:hover { background: #f5f5f5; }
        .sheet-tab.active {
            background: #fff;
            border-color: #ccc;
            font-weight: 600;
        }
        .sheet-content { background: #fff; flex: 1; }
        .table-wrapper {
            overflow: auto;
            max-height: calc(100vh - 90px);
            background: #fff;
        }
        table {
            border-collapse: collapse;
            font-size: 11px;
            font-family: 'Calibri', 'Segoe UI', sans-serif;
            table-layout: fixed;
        }
        .row-header-col { width: 40px; }
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
            max-width: 300px;
            height: 20px;
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
        """;

    // ==================== JavaScript ====================

    private static string GenerateExcelJs() => """
        function switchSheet(idx) {
            document.querySelectorAll('.sheet-tab').forEach(function(t) {
                t.classList.toggle('active', parseInt(t.getAttribute('data-sheet')) === idx);
            });
            document.querySelectorAll('.sheet-content').forEach(function(c) {
                c.style.display = parseInt(c.getAttribute('data-sheet')) === idx ? '' : 'none';
            });
        }
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
