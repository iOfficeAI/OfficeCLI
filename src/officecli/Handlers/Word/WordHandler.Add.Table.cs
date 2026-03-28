// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddTable(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var table = new Table();
        var tblProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            )
        );
        table.AppendChild(tblProps);

        // Apply border properties from Add parameters
        foreach (var (bk, bv) in properties)
        {
            if (bk.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                ApplyTableBorders(tblProps, bk, bv);
        }

        // Parse data if provided: "H1,H2;R1C1,R1C2;R2C1,R2C2" or CSV file path
        string[][]? tableData = null;
        if (properties.TryGetValue("data", out var dataStr))
        {
            if (File.Exists(dataStr))
                tableData = File.ReadAllLines(dataStr)
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(l => l.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
            else
                tableData = dataStr.Split(';')
                    .Select(r => r.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
        }

        int rows, cols;
        if (tableData != null)
        {
            rows = tableData.Length;
            cols = tableData.Max(r => r.Length);
        }
        else
        {
            rows = 1;
            if (properties.TryGetValue("rows", out var rowsStr))
            {
                if (!int.TryParse(rowsStr, out rows))
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Expected a positive integer.");
                if (rows <= 0)
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Must be a positive integer (> 0).");
            }
            cols = 1;
            if (properties.TryGetValue("cols", out var colsStr))
            {
                cols = ParseHelpers.SafeParseInt(colsStr, "cols");
                if (cols <= 0)
                    throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Must be a positive integer (> 0).");
            }
        }

        // Parse per-column widths: colWidths="3000,2000,5000"
        int[]? colWidthArr = null;
        if (properties.TryGetValue("colwidths", out var cwStr) || properties.TryGetValue("colWidths", out cwStr))
        {
            var parts = cwStr.Split(',');
            colWidthArr = new int[parts.Length];
            for (int ci = 0; ci < parts.Length; ci++)
            {
                if (!int.TryParse(parts[ci].Trim(), out colWidthArr[ci]))
                    throw new ArgumentException($"Invalid 'colwidths' value: '{parts[ci].Trim()}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
            }
        }

        // Add table grid
        var tblGrid = new TableGrid();
        for (int gc = 0; gc < cols; gc++)
        {
            var w = colWidthArr != null && gc < colWidthArr.Length ? colWidthArr[gc].ToString() : "2400";
            tblGrid.AppendChild(new GridColumn { Width = w });
        }
        table.AppendChild(tblGrid);

        // Apply table-level properties from Add parameters
        foreach (var (tk, tv) in properties)
        {
            var tkl = tk.ToLowerInvariant();
            if (tkl is "rows" or "cols" or "colwidths" || tkl.StartsWith("border")) continue;
            switch (tkl)
            {
                case "alignment":
                    tblProps.TableJustification = new TableJustification
                    {
                        Val = tv.ToLowerInvariant() switch
                        {
                            "center" => TableRowAlignmentValues.Center,
                            "right" => TableRowAlignmentValues.Right,
                            "left" => TableRowAlignmentValues.Left,
                            _ => throw new ArgumentException($"Invalid table alignment value: '{tv}'. Valid values: left, center, right.")
                        }
                    };
                    break;
                case "width":
                    if (tv.EndsWith('%'))
                    {
                        var pct = ParseHelpers.SafeParseInt(tv.TrimEnd('%'), "width") * 50;
                        tblProps.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                    }
                    else
                    {
                        tblProps.TableWidth = new TableWidth { Width = ParseHelpers.SafeParseUint(tv, "width").ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "indent":
                    tblProps.TableIndentation = new TableIndentation { Width = ParseHelpers.SafeParseInt(tv, "indent"), Type = TableWidthUnitValues.Dxa };
                    break;
                case "cellspacing":
                    tblProps.TableCellSpacing = new TableCellSpacing { Width = ParseHelpers.SafeParseUint(tv, "cellspacing").ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                case "layout":
                    tblProps.TableLayout = new TableLayout
                    {
                        Type = tv.ToLowerInvariant() == "fixed" ? TableLayoutValues.Fixed : TableLayoutValues.Autofit
                    };
                    break;
                case "padding":
                    var cm = tblProps.TableCellMarginDefault ?? tblProps.AppendChild(new TableCellMarginDefault());
                    var paddingVal = ParseHelpers.SafeParseInt(tv, "padding");
                    cm.TopMargin = new TopMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    cm.BottomMargin = new BottomMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    break;
                case "style":
                    tblProps.TableStyle = new TableStyle { Val = tv };
                    // Add TableLook so built-in styles apply banding correctly
                    tblProps.RemoveAllChildren<TableLook>();
                    tblProps.AppendChild(new TableLook { Val = "04A0" });
                    break;
            }
        }

        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
            {
                var cellText = tableData != null && r < tableData.Length && c < tableData[r].Length
                    ? tableData[r][c] : (properties.TryGetValue($"r{r + 1}c{c + 1}", out var rc) ? rc : "");
                var cellPara = new Paragraph(new ParagraphProperties(
                    new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }));
                if (!string.IsNullOrEmpty(cellText))
                    cellPara.AppendChild(new Run(new Text { Text = cellText, Space = SpaceProcessingModeValues.Preserve }));
                var cell = new TableCell(cellPara);
                if (colWidthArr != null && c < colWidthArr.Length)
                    cell.PrependChild(new TableCellProperties(new TableCellWidth { Width = colWidthArr[c].ToString(), Type = TableWidthUnitValues.Dxa }));
                row.AppendChild(cell);
            }
            table.AppendChild(row);
        }

        AppendToParent(parent, table);
        var tblCount = parent.Elements<Table>().Count();
        return $"{parentPath}/tbl[{tblCount}]";
    }

    private string AddRow(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not Table targetTable)
            throw new ArgumentException("Rows can only be added to a table: /body/tbl[N]");

        var existingCols = targetTable.Elements<TableGrid>().FirstOrDefault()
            ?.Elements<GridColumn>().Count() ?? 1;
        int newCols = existingCols;
        if (properties.TryGetValue("cols", out var colsVal))
            newCols = ParseHelpers.SafeParseInt(colsVal, "cols");

        var newRow = new TableRow();
        TableRowProperties? newRowProps = null;
        if (properties.TryGetValue("height", out var rowHeight))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeight), HeightType = HeightRuleValues.AtLeast });
        }
        if (properties.TryGetValue("height.exact", out var rowHeightExact))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.GetFirstChild<TableRowHeight>()?.Remove();
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeightExact), HeightType = HeightRuleValues.Exact });
        }
        if (properties.TryGetValue("header", out var headerVal) && IsTruthy(headerVal))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.AppendChild(new TableHeader());
        }

        for (int c = 0; c < newCols; c++)
        {
            var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
            var cellPara = new Paragraph();
            if (!string.IsNullOrEmpty(cellText))
                cellPara.AppendChild(new Run(new Text { Text = cellText, Space = SpaceProcessingModeValues.Preserve }));
            newRow.AppendChild(new TableCell(cellPara));
        }

        if (index.HasValue)
        {
            var existingRows = targetTable.Elements<TableRow>().ToList();
            if (index.Value < existingRows.Count)
                targetTable.InsertBefore(newRow, existingRows[index.Value]);
            else
                targetTable.AppendChild(newRow);
        }
        else
        {
            targetTable.AppendChild(newRow);
        }

        var rowIdx = targetTable.Elements<TableRow>().ToList().IndexOf(newRow) + 1;
        return $"{parentPath}/tr[{rowIdx}]";
    }

    private string AddCell(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not TableRow targetRow)
            throw new ArgumentException("Cells can only be added to a table row: /body/tbl[N]/tr[M]");

        var cellParagraph = new Paragraph();
        if (properties.TryGetValue("text", out var cellTxt))
            cellParagraph.AppendChild(new Run(new Text { Text = cellTxt, Space = SpaceProcessingModeValues.Preserve }));

        var newCell = new TableCell(cellParagraph);

        if (properties.TryGetValue("width", out var cellWidth))
            newCell.PrependChild(new TableCellProperties(new TableCellWidth { Width = cellWidth, Type = TableWidthUnitValues.Dxa }));

        if (index.HasValue)
        {
            var cells = targetRow.Elements<TableCell>().ToList();
            if (index.Value < cells.Count)
                targetRow.InsertBefore(newCell, cells[index.Value]);
            else
                targetRow.AppendChild(newCell);
        }
        else
        {
            targetRow.AppendChild(newCell);
        }

        var cellIdx = targetRow.Elements<TableCell>().ToList().IndexOf(newCell) + 1;
        return $"{parentPath}/tc[{cellIdx}]";
    }
}
