// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Table Rendering ====================

    private void RenderTableHtml(StringBuilder sb, Table table, string? dataPath = null)
    {
        // Check table-level borders to determine if this is a borderless layout table
        // First try direct table borders, then fall back to table style borders
        var tblPr = table.GetFirstChild<TableProperties>();
        var tblBorders = tblPr?.TableBorders;
        var styleId = tblPr?.TableStyle?.Val?.Value;
        if (tblBorders == null && styleId != null)
            tblBorders = ResolveTableStyleBorders(styleId);
        bool tableBordersNone = IsTableBorderless(tblBorders);

        // Parse tblLook bitmask for conditional formatting
        var tblLook = ParseTableLook(tblPr);

        // Resolve conditional formatting from table style
        var condFormats = styleId != null ? ResolveTableStyleConditionalFormats(styleId) : null;

        // Check for floating table (tblpPr = text wrapping)
        var tblpPr = tblPr?.GetFirstChild<TablePositionProperties>();
        var tableStyles = new List<string>();
        if (tblpPr != null)
        {
            // #2: Float the table with approximate positioning. Horizontal
            // anchor + tblpX/tblpY translated into float + margin. Coverage
            // is ~40% of Word's 2D flow (horzAnchor=margin + vertAnchor=text);
            // vertAnchor=page/margin would need absolute positioning which
            // doesn't interact with text flow.
            var hAnchor = tblpPr.HorizontalAnchor?.InnerText;
            var vAnchor = tblpPr.VerticalAnchor?.InnerText;
            var tblpX = tblpPr.TablePositionX?.Value ?? 0;
            var tblpY = tblpPr.TablePositionY?.Value ?? 0;
            var xAlign = tblpPr.TablePositionXAlignment?.InnerText;
            var floatDir = xAlign == "right" || (hAnchor == "page" && tblpX > 5000)
                ? "right"
                : xAlign == "left" ? "left" : "left";
            tableStyles.Add($"float:{floatDir}");
            // Margins from text distance (dist…FromText).
            var rightDist = tblpPr.RightFromText?.Value ?? 0;
            var bottomDist = tblpPr.BottomFromText?.Value ?? 0;
            var leftDist = tblpPr.LeftFromText?.Value ?? 0;
            var topDist = tblpPr.TopFromText?.Value ?? 0;
            // Fold tblpX into margin-left (or margin-right for float:right)
            // when the anchor is margin-relative so the column offset shows.
            var horzShiftPt = hAnchor == "margin" ? tblpX / 20.0 : 0;
            if (floatDir == "left")
            {
                var leftMargin = leftDist / 20.0 + horzShiftPt;
                if (leftMargin > 0) tableStyles.Add($"margin-left:{leftMargin:0.#}pt");
                if (rightDist > 0) tableStyles.Add($"margin-right:{rightDist / 20.0:0.#}pt");
            }
            else
            {
                var rightMargin = rightDist / 20.0 + horzShiftPt;
                if (rightMargin > 0) tableStyles.Add($"margin-right:{rightMargin:0.#}pt");
                if (leftDist > 0) tableStyles.Add($"margin-left:{leftDist / 20.0:0.#}pt");
            }
            // Vertical offset: only honor vertAnchor=text (default); other
            // anchors would need absolute positioning, which breaks text
            // flow and is better left to a future pass.
            var vertShiftPt = (vAnchor == null || vAnchor == "text") ? tblpY / 20.0 : 0;
            var topMargin = topDist / 20.0 + vertShiftPt;
            if (topMargin > 0) tableStyles.Add($"margin-top:{topMargin:0.#}pt");
            if (bottomDist > 0) tableStyles.Add($"margin-bottom:{bottomDist / 20.0:0.#}pt");
        }

        // Table horizontal alignment on page (jc = center/right)
        var tblJc = tblPr?.TableJustification?.Val?.InnerText;
        if (tblJc == "center")
            tableStyles.Add("margin-left:auto;margin-right:auto");
        else if (tblJc == "right")
            tableStyles.Add("margin-left:auto;margin-right:0");

        // Apply base table style rPr (font-size, color, alignment) to the <table>
        if (styleId != null)
        {
            var baseStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
            var baseRPr = baseStyle?.StyleRunProperties;
            if (baseRPr?.FontSize?.Val?.Value is string bsz && int.TryParse(bsz, out var bhp))
                tableStyles.Add($"font-size:{bhp / 2.0:0.##}pt");
            var baseColor = ResolveRunColor(baseRPr?.Color);
            if (baseColor != null) tableStyles.Add($"color:{baseColor}");
            var basePPr = baseStyle?.StyleParagraphProperties;
            if (basePPr?.Justification?.Val?.InnerText is string bjc)
            {
                var align = bjc switch { "center" => "center", "right" => "right", _ => (string?)null };
                if (align != null) tableStyles.Add($"text-align:{align}");
            }
        }

        // Table width: explicit tblW → use it; pct → percentage; otherwise sum gridCol widths
        var tblW = tblPr?.TableWidth;
        var tblWType = tblW?.Type?.InnerText;
        if (tblWType == "dxa" && int.TryParse(tblW!.Width?.Value, out var twW) && twW > 0)
        {
            tableStyles.Add($"width:{twW / 20.0:0.##}pt");
        }
        else if (tblWType == "pct" && int.TryParse(tblW!.Width?.Value, out var pctW) && pctW > 0)
        {
            // pct values are in 1/50th of a percent (5000 = 100%)
            tableStyles.Add($"width:{pctW / 50.0:0.##}%");
        }
        else
        {
            // No explicit tblW or type=auto: use gridCol sum as max-width (Word auto-fit behavior)
            // auto layout tables in Word shrink to content; max-width lets browser do the same
            var isFixed = tblPr?.TableLayout?.Type?.InnerText == "fixed";
            var grid = table.GetFirstChild<TableGrid>();
            var gridCols = grid?.Elements<GridColumn>().ToList();
            if (gridCols != null && gridCols.Count > 0)
            {
                int totalTwips = 0;
                bool allValid = true;
                foreach (var gc in gridCols)
                {
                    if (gc.Width?.Value is string gw && int.TryParse(gw, out var gwVal))
                        totalTwips += gwVal;
                    else
                        allValid = false;
                }
                if (allValid && totalTwips > 0)
                {
                    var prop = isFixed ? "width" : "max-width";
                    tableStyles.Add($"{prop}:{totalTwips / 20.0:0.##}pt");
                }
            }
            // else: no grid info — browser auto-fits to content
        }

        var tableClass = tableBordersNone ? "borderless" : "";
        var tableStyleAttr = tableStyles.Count > 0 ? $" style=\"{string.Join(";", tableStyles)}\"" : "";
        var dataPathAttr = !string.IsNullOrEmpty(dataPath) ? $" data-path=\"{dataPath}\"" : "";
        if (!string.IsNullOrEmpty(tableClass))
            sb.AppendLine($"<table class=\"{tableClass}\"{dataPathAttr}{tableStyleAttr}>");
        else
            sb.AppendLine($"<table{dataPathAttr}{tableStyleAttr}>");

        // Get column widths from grid
        // tblLayout=fixed → use fixed col widths; auto/missing → let browser auto-fit by content
        var isFixedLayout = tblPr?.TableLayout?.Type?.InnerText == "fixed";
        var tblGrid = table.GetFirstChild<TableGrid>();
        if (tblGrid != null)
        {
            sb.Append("<colgroup>");
            foreach (var col in tblGrid.Elements<GridColumn>())
            {
                var w = col.Width?.Value;
                if (w != null && isFixedLayout)
                {
                    var pt = double.Parse(w, System.Globalization.CultureInfo.InvariantCulture) / 20.0; // twips to pt
                    sb.Append($"<col style=\"width:{pt:0.##}pt\">");
                }
                else
                {
                    sb.Append("<col>");
                }
            }
            sb.AppendLine("</colgroup>");
        }

        var rows = table.Elements<TableRow>().ToList();
        var totalRows = rows.Count;
        var totalCols = tblGrid?.Elements<GridColumn>().Count() ?? rows.FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;

        for (int rowIdx = 0; rowIdx < totalRows; rowIdx++)
        {
            var row = rows[rowIdx];
            var isHeader = row.TableRowProperties?.GetFirstChild<TableHeader>() != null;
            // Row height
            var trHeight = row.TableRowProperties?.GetFirstChild<TableRowHeight>();
            var trStyle = "";
            if (trHeight?.Val?.Value is uint hVal && hVal > 0)
                trStyle = $" style=\"height:{hVal / 20.0:0.#}pt\"";
            // #7b00: mark tblHeader rows so the JS paginator can clone them
            // onto every continuation page when a long table spans pages.
            var hdrMarker = isHeader ? " data-tbl-header=\"1\"" : "";
            // Row data-path for goto/mark navigation. Skipped for nested tables
            // (dataPath is only set for top-level tables — see RenderTableHtml
            // call sites in HtmlPreview.cs:1906) because nested tables don't
            // have a stable /body/table[N] index.
            var rowDataPath = !string.IsNullOrEmpty(dataPath) ? $"{dataPath}/tr[{rowIdx + 1}]" : null;
            var rowDataPathAttr = rowDataPath != null ? $" data-path=\"{rowDataPath}\"" : "";
            sb.AppendLine(isHeader ? $"<tr class=\"header-row\"{hdrMarker}{rowDataPathAttr}{trStyle}>" : $"<tr{rowDataPathAttr}{trStyle}>");

            int colIdx = 0;
            foreach (var cell in row.Elements<TableCell>())
            {
                var tag = isHeader ? "th" : "td";
                var condTypes = GetConditionalTypes(tblLook, rowIdx, colIdx, totalRows, totalCols);
                var cellStyle = GetTableCellInlineCss(cell, tableBordersNone, tblBorders, condFormats, condTypes,
                    rowIdx, colIdx, totalRows, totalCols);

                // Check if conditional format overrides font-size (needs class for CSS override)
                bool hasTsf = cellStyle.Contains("__TSF__");
                cellStyle = cellStyle.Replace(";__TSF__", "").Replace("__TSF__", "");

                // Merge attributes
                var attrs = new StringBuilder();
                if (hasTsf) attrs.Append(" class=\"tsf\"");
                var gridSpan = cell.TableCellProperties?.GridSpan?.Val?.Value;
                if (gridSpan > 1) attrs.Append($" colspan=\"{gridSpan}\"");

                var vMerge = cell.TableCellProperties?.VerticalMerge;
                if (vMerge != null && vMerge.Val?.Value == MergedCellValues.Restart)
                {
                    // Count rowspan
                    var rowspan = CountRowSpan(table, row, cell);
                    if (rowspan > 1) attrs.Append($" rowspan=\"{rowspan}\"");
                }
                else if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
                {
                    colIdx += gridSpan ?? 1;
                    continue; // Skip merged continuation cells
                }

                if (!string.IsNullOrEmpty(cellStyle))
                    attrs.Append($" style=\"{cellStyle}\"");

                // Cell data-path uses the OOXML positional cell index (colIdx+1)
                // rather than the visual grid column, to match the handler's
                // /body/table[N]/tr[R]/tc[C] addressing.
                if (rowDataPath != null)
                    attrs.Append($" data-path=\"{rowDataPath}/tc[{colIdx + 1}]\"");

                sb.Append($"<{tag}{attrs}>");

                // Render cell content — every paragraph (including empty ones)
                // goes through the same path as body paragraphs: <div> wrapper
                // with inline pPr CSS plus an &nbsp; placeholder for empties so
                // the line box forms and renders the resolved line-height. Word
                // and LibreOffice apply pPr to empty paragraphs identically to
                // non-empty ones; the previous <br> fast-path silently dropped
                // line-height + spaceBefore + spaceAfter for empty cell paragraphs.
                foreach (var cellPara in cell.Elements<Paragraph>())
                {
                    var text = GetParagraphText(cellPara);
                    var runs = GetAllRuns(cellPara);
                    var pCss = GetParagraphInlineCss(cellPara);
                    sb.Append("<div");
                    if (!string.IsNullOrEmpty(pCss))
                        sb.Append($" style=\"{pCss}\"");
                    sb.Append(">");
                    bool hasVisibleContent = runs.Count > 0 || !string.IsNullOrWhiteSpace(text);
                    RenderParagraphContentHtml(sb, cellPara);
                    if (!hasVisibleContent) sb.Append("&nbsp;");
                    sb.Append("</div>");
                }

                // Render nested tables
                foreach (var nestedTable in cell.Elements<Table>())
                    RenderTableHtml(sb, nestedTable);

                sb.AppendLine($"</{tag}>");
                colIdx += gridSpan ?? 1;
            }

            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</table>");
    }

    private static bool IsTableBorderless(TableBorders? borders)
    {
        if (borders == null) return false;
        // Check if all borders are none/nil
        return IsBorderNone(borders.TopBorder)
            && IsBorderNone(borders.BottomBorder)
            && IsBorderNone(borders.LeftBorder)
            && IsBorderNone(borders.RightBorder)
            && IsBorderNone(borders.InsideHorizontalBorder)
            && IsBorderNone(borders.InsideVerticalBorder);
    }

    private static bool IsBorderNone(OpenXmlElement? border)
    {
        if (border == null) return true;
        var val = border.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return val is null or "nil" or "none";
    }

    /// <summary>Apply or clear a conditional format border edge.</summary>
    private void ApplyCondBorder(List<string> parts, OpenXmlElement? border, string cssProperty)
    {
        if (border == null) return;
        parts.RemoveAll(p => p.StartsWith(cssProperty + ":"));
        if (!IsBorderNone(border))
            RenderBorderCss(parts, border, cssProperty);
        // If val=nil/none, the RemoveAll already cleared it — border is removed
    }

    /// <summary>Resolve TableBorders from a table style (walking basedOn chain).</summary>
    private TableBorders? ResolveTableStyleBorders(string styleId)
    {
        var visited = new HashSet<string>();
        var currentId = styleId;
        while (currentId != null && visited.Add(currentId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentId);
            if (style == null) break;
            var borders = style.StyleTableProperties?.TableBorders;
            if (borders != null) return borders;
            currentId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    // ==================== Table Look / Conditional Formatting ====================

    [Flags]
    private enum TableLookFlags
    {
        None = 0,
        FirstRow = 0x0020,
        LastRow = 0x0040,
        FirstColumn = 0x0080,
        LastColumn = 0x0100,
        NoHBand = 0x0200,
        NoVBand = 0x0400,
    }

    /// <summary>Parse tblLook from table properties. Start from the legacy
    /// val hex bitmask (if present) and let each authored individual attr
    /// override only the bit it names — per ECMA-376 §17.7.6.7, individual
    /// attrs are independent overrides of val, not a full replacement.</summary>
    private static TableLookFlags ParseTableLook(TableProperties? tblPr)
    {
        var tblLook = tblPr?.GetFirstChild<TableLook>();
        if (tblLook == null) return TableLookFlags.None;

        var flags = TableLookFlags.None;
        var val = tblLook.Val?.Value;
        if (val != null && int.TryParse(val, System.Globalization.NumberStyles.HexNumber, null, out var hex))
            flags = (TableLookFlags)hex;

        // Each authored attr (regardless of true/false) overrides its bit.
        if (tblLook.FirstRow != null)
            flags = tblLook.FirstRow.Value == true ? flags | TableLookFlags.FirstRow : flags & ~TableLookFlags.FirstRow;
        if (tblLook.LastRow != null)
            flags = tblLook.LastRow.Value == true ? flags | TableLookFlags.LastRow : flags & ~TableLookFlags.LastRow;
        if (tblLook.FirstColumn != null)
            flags = tblLook.FirstColumn.Value == true ? flags | TableLookFlags.FirstColumn : flags & ~TableLookFlags.FirstColumn;
        if (tblLook.LastColumn != null)
            flags = tblLook.LastColumn.Value == true ? flags | TableLookFlags.LastColumn : flags & ~TableLookFlags.LastColumn;
        if (tblLook.NoHorizontalBand != null)
            flags = tblLook.NoHorizontalBand.Value == true ? flags | TableLookFlags.NoHBand : flags & ~TableLookFlags.NoHBand;
        if (tblLook.NoVerticalBand != null)
            flags = tblLook.NoVerticalBand.Value == true ? flags | TableLookFlags.NoVBand : flags & ~TableLookFlags.NoVBand;

        return flags;
    }

    /// <summary>Cached conditional format data from a table style.</summary>
    private class TableConditionalFormat
    {
        public Shading? Shading { get; set; }
        public TableCellBorders? Borders { get; set; }
        public RunPropertiesBaseStyle? RunProperties { get; set; }
    }

    /// <summary>Resolve all tblStylePr conditional formatting from a table style (walking basedOn chain).</summary>
    private Dictionary<string, TableConditionalFormat>? ResolveTableStyleConditionalFormats(string styleId)
    {
        var result = new Dictionary<string, TableConditionalFormat>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<string>();
        var currentId = styleId;

        // Walk basedOn chain, collecting conditional formats (child style overrides parent)
        var chainStyles = new List<Style>();
        while (currentId != null && visited.Add(currentId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentId);
            if (style == null) break;
            chainStyles.Add(style);
            currentId = style.BasedOn?.Val?.Value;
        }

        // Process in reverse (base first, derived last — derived wins)
        chainStyles.Reverse();
        foreach (var style in chainStyles)
        {
            foreach (var tsp in style.Elements<TableStyleProperties>())
            {
                var type = tsp.Type;
                if (type == null) continue;
                // Use the XML serialized value (e.g. "firstRow", "band1Horz") for consistent lookup
                var typeName = type.InnerText;

                var fmt = new TableConditionalFormat();
                // Try SDK-typed property first, then fall back to generic child lookup
                var tcPr = tsp.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>();
                if (tcPr != null)
                {
                    fmt.Shading = tcPr.GetFirstChild<Shading>();
                    fmt.Borders = tcPr.GetFirstChild<TableCellBorders>();
                }
                fmt.RunProperties = tsp.GetFirstChild<RunPropertiesBaseStyle>();

                if (typeName != null)
                    result[typeName] = fmt;
            }
        }

        return result.Count > 0 ? result : null;
    }

    /// <summary>Get the list of conditional format type names that apply to a cell at the given position.</summary>
    private static List<string> GetConditionalTypes(TableLookFlags look, int rowIdx, int colIdx, int totalRows, int totalCols)
    {
        var types = new List<string>();

        // Banded rows (applied first, lowest priority)
        if ((look & TableLookFlags.NoHBand) == 0)
        {
            // Banding skips first/last row if those flags are set
            int bandRowIdx = rowIdx;
            if ((look & TableLookFlags.FirstRow) != 0 && rowIdx > 0) bandRowIdx = rowIdx - 1;
            else if ((look & TableLookFlags.FirstRow) != 0 && rowIdx == 0) bandRowIdx = -1; // first row, skip banding

            if (bandRowIdx >= 0)
                types.Add(bandRowIdx % 2 == 0 ? "band1Horz" : "band2Horz");
        }

        // Banded columns
        if ((look & TableLookFlags.NoVBand) == 0)
        {
            int bandColIdx = colIdx;
            if ((look & TableLookFlags.FirstColumn) != 0 && colIdx > 0) bandColIdx = colIdx - 1;
            else if ((look & TableLookFlags.FirstColumn) != 0 && colIdx == 0) bandColIdx = -1;

            if (bandColIdx >= 0)
                types.Add(bandColIdx % 2 == 0 ? "band1Vert" : "band2Vert");
        }

        // First/last column (higher priority than banding)
        if ((look & TableLookFlags.FirstColumn) != 0 && colIdx == 0)
            types.Add("firstCol");
        if ((look & TableLookFlags.LastColumn) != 0 && colIdx == totalCols - 1)
            types.Add("lastCol");

        // First/last row (highest priority)
        if ((look & TableLookFlags.FirstRow) != 0 && rowIdx == 0)
            types.Add("firstRow");
        if ((look & TableLookFlags.LastRow) != 0 && rowIdx == totalRows - 1)
            types.Add("lastRow");

        return types;
    }

    /// <summary>Calculate the grid column index for a cell, accounting for gridSpan in preceding cells.</summary>
    private static int GetGridColumn(TableRow row, TableCell cell)
    {
        int gridCol = 0;
        foreach (var c in row.Elements<TableCell>())
        {
            if (c == cell) return gridCol;
            gridCol += c.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
        }
        return gridCol;
    }

    /// <summary>Find the cell at a given grid column in a row, accounting for gridSpan.</summary>
    private static TableCell? GetCellAtGridColumn(TableRow row, int targetGridCol)
    {
        int gridCol = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            if (gridCol == targetGridCol) return cell;
            gridCol += cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            if (gridCol > targetGridCol) return null; // target is inside a spanned cell
        }
        return null;
    }

    private static int CountRowSpan(Table table, TableRow startRow, TableCell startCell)
    {
        var rows = table.Elements<TableRow>().ToList();
        var startRowIdx = rows.IndexOf(startRow);
        if (startRowIdx < 0) return 1;

        // Use grid column position instead of cell index
        var gridCol = GetGridColumn(startRow, startCell);

        int span = 1;
        for (int i = startRowIdx + 1; i < rows.Count; i++)
        {
            var cell = GetCellAtGridColumn(rows[i], gridCol);
            if (cell == null) break;

            var vm = cell.TableCellProperties?.VerticalMerge;
            if (vm != null && (vm.Val == null || vm.Val.Value == MergedCellValues.Continue))
                span++;
            else
                break;
        }
        return span;
    }
}
