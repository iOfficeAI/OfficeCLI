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

    private void RenderTableHtml(StringBuilder sb, Table table)
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
            // Float the table; determine alignment from horzAnchor/tblpX
            var hAnchor = tblpPr.HorizontalAnchor?.InnerText;
            var tblpX = tblpPr.TablePositionX?.Value ?? 0;
            var floatDir = (hAnchor == "page" && tblpX > 5000) ? "right" : "left";
            tableStyles.Add($"float:{floatDir}");
            // Margins from text distance
            var rightDist = tblpPr.RightFromText?.Value ?? 0;
            var bottomDist = tblpPr.BottomFromText?.Value ?? 0;
            var leftDist = tblpPr.LeftFromText?.Value ?? 0;
            if (rightDist > 0) tableStyles.Add($"margin-right:{rightDist / 20.0:0.#}pt");
            if (bottomDist > 0) tableStyles.Add($"margin-bottom:{bottomDist / 20.0:0.#}pt");
            if (leftDist > 0) tableStyles.Add($"margin-left:{leftDist / 20.0:0.#}pt");
        }

        var tableClass = tableBordersNone ? "borderless" : "";
        var tableStyleAttr = tableStyles.Count > 0 ? $" style=\"{string.Join(";", tableStyles)}\"" : "";
        if (!string.IsNullOrEmpty(tableClass))
            sb.AppendLine($"<table class=\"{tableClass}\"{tableStyleAttr}>");
        else
            sb.AppendLine($"<table{tableStyleAttr}>");

        // Get column widths from grid
        var tblGrid = table.GetFirstChild<TableGrid>();
        if (tblGrid != null)
        {
            sb.Append("<colgroup>");
            foreach (var col in tblGrid.Elements<GridColumn>())
            {
                var w = col.Width?.Value;
                if (w != null)
                {
                    var px = (int)(double.Parse(w, System.Globalization.CultureInfo.InvariantCulture) / 1440.0 * 96); // twips to px
                    sb.Append($"<col style=\"width:{px}px\">");
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
            sb.AppendLine(isHeader ? "<tr class=\"header-row\">" : "<tr>");

            int colIdx = 0;
            foreach (var cell in row.Elements<TableCell>())
            {
                var tag = isHeader ? "th" : "td";
                var condTypes = GetConditionalTypes(tblLook, rowIdx, colIdx, totalRows, totalCols);
                var cellStyle = GetTableCellInlineCss(cell, tableBordersNone, tblBorders, condFormats, condTypes,
                    rowIdx, colIdx, totalRows, totalCols);

                // Merge attributes
                var attrs = new StringBuilder();
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

                sb.Append($"<{tag}{attrs}>");

                // Render cell content — use paragraph tags for multi-paragraph cells
                var cellParagraphs = cell.Elements<Paragraph>().ToList();
                for (int pi = 0; pi < cellParagraphs.Count; pi++)
                {
                    var cellPara = cellParagraphs[pi];
                    var text = GetParagraphText(cellPara);
                    var runs = GetAllRuns(cellPara);

                    if (runs.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        // empty cell paragraph — skip but preserve spacing between paragraphs
                        if (pi > 0 && pi < cellParagraphs.Count - 1)
                            sb.Append("<br>");
                    }
                    else
                    {
                        var pCss = GetParagraphInlineCss(cellPara);
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append($"<div style=\"{pCss}\">");
                        RenderParagraphContentHtml(sb, cellPara);
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append("</div>");
                        else if (pi < cellParagraphs.Count - 1)
                            sb.Append("<br>");
                    }
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

    /// <summary>Parse tblLook from table properties. Supports both val hex bitmask and individual attributes.</summary>
    private static TableLookFlags ParseTableLook(TableProperties? tblPr)
    {
        var tblLook = tblPr?.GetFirstChild<TableLook>();
        if (tblLook == null) return TableLookFlags.None;

        // Try val attribute (hex bitmask)
        var val = tblLook.Val?.Value;
        if (val != null && int.TryParse(val, System.Globalization.NumberStyles.HexNumber, null, out var hex))
            return (TableLookFlags)hex;

        // Fall back to individual boolean attributes
        var flags = TableLookFlags.None;
        if (tblLook.FirstRow?.Value == true) flags |= TableLookFlags.FirstRow;
        if (tblLook.LastRow?.Value == true) flags |= TableLookFlags.LastRow;
        if (tblLook.FirstColumn?.Value == true) flags |= TableLookFlags.FirstColumn;
        if (tblLook.LastColumn?.Value == true) flags |= TableLookFlags.LastColumn;
        if (tblLook.NoHorizontalBand?.Value == true) flags |= TableLookFlags.NoHBand;
        if (tblLook.NoVerticalBand?.Value == true) flags |= TableLookFlags.NoVBand;
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
