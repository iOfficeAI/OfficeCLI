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
        var tblBorders = table.GetFirstChild<TableProperties>()?.TableBorders;
        bool tableBordersNone = IsTableBorderless(tblBorders);

        var tableClass = tableBordersNone ? "borderless" : "";
        sb.AppendLine(string.IsNullOrEmpty(tableClass) ? "<table>" : $"<table class=\"{tableClass}\">");

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

        foreach (var row in table.Elements<TableRow>())
        {
            var isHeader = row.TableRowProperties?.GetFirstChild<TableHeader>() != null;
            sb.AppendLine(isHeader ? "<tr class=\"header-row\">" : "<tr>");

            foreach (var cell in row.Elements<TableCell>())
            {
                var tag = isHeader ? "th" : "td";
                var cellStyle = GetTableCellInlineCss(cell, tableBordersNone, tblBorders);

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
