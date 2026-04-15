// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string AddTable(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var tblSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!tblSlideMatch.Success)
                    throw new ArgumentException("Tables must be added to a slide: /slide[N]");

                var tblSlideIdx = int.Parse(tblSlideMatch.Groups[1].Value);
                var tblSlideParts = GetSlideParts().ToList();
                if (tblSlideIdx < 1 || tblSlideIdx > tblSlideParts.Count)
                    throw new ArgumentException($"Slide {tblSlideIdx} not found (total: {tblSlideParts.Count})");

                var tblSlidePart = tblSlideParts[tblSlideIdx - 1];
                var tblShapeTree = GetSlide(tblSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                // Parse data if provided: "H1,H2;R1C1,R1C2;R2C1,R2C2" or CSV file/URL/data-URI
                string[][]? tableData = null;
                if (properties.TryGetValue("data", out var dataStr))
                {
                    if (OfficeCli.Core.FileSource.IsResolvable(dataStr))
                    {
                        // CSV file/URL/data-URI
                        tableData = OfficeCli.Core.FileSource.ResolveLines(dataStr)
                            .Where(l => !string.IsNullOrWhiteSpace(l))
                            .Select(l => l.Split(',').Select(c => c.Trim()).ToArray())
                            .ToArray();
                    }
                    else
                    {
                        // Inline: semicolons separate rows, commas separate cells
                        tableData = dataStr.Split(';')
                            .Select(r => r.Split(',').Select(c => c.Trim()).ToArray())
                            .ToArray();
                    }
                }

                int rows, cols;
                if (tableData != null)
                {
                    rows = tableData.Length;
                    cols = tableData.Max(r => r.Length);
                }
                else
                {
                    var rowsStr = properties.GetValueOrDefault("rows", "3");
                    var colsStr = properties.GetValueOrDefault("cols", "3");
                    if (!int.TryParse(rowsStr, out rows))
                        throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Expected a positive integer.");
                    if (!int.TryParse(colsStr, out cols))
                        throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Expected a positive integer.");
                }
                if (rows < 1 || cols < 1)
                    throw new ArgumentException("rows and cols must be >= 1");

                // Position & size
                long tblX = properties.TryGetValue("x", out var txStr) ? ParseEmu(txStr) : 457200; // ~1.27cm
                long tblY = properties.TryGetValue("y", out var tyStr) ? ParseEmu(tyStr) : 1600200; // ~4.44cm
                long tblCx = properties.TryGetValue("width", out var twStr) ? ParseEmu(twStr) : 8229600; // ~22.86cm
                long rowHeight;
                long tblCy;
                if (properties.TryGetValue("rowHeight", out var rhStr) || properties.TryGetValue("rowheight", out rhStr))
                {
                    rowHeight = ParseEmu(rhStr);
                    tblCy = properties.TryGetValue("height", out var thStr) ? ParseEmu(thStr) : rowHeight * rows;
                }
                else
                {
                    tblCy = properties.TryGetValue("height", out var thStr) ? ParseEmu(thStr) : (long)(rows * 370840); // ~1.03cm per row
                    rowHeight = tblCy / rows;
                }
                long colWidth = tblCx / cols;

                var tblId = GenerateUniqueShapeId(tblShapeTree);

                // Build GraphicFrame
                var graphicFrame = new GraphicFrame();
                graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = tblId, Name = properties.GetValueOrDefault("name", $"Table {tblShapeTree.Elements<GraphicFrame>().Count(gf => gf.Descendants<Drawing.Table>().Any()) + 1}") },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                graphicFrame.Transform = new Transform(
                    new Drawing.Offset { X = tblX, Y = tblY },
                    new Drawing.Extents { Cx = tblCx, Cy = tblCy }
                );

                // Build table
                var table = new Drawing.Table();
                var tblProps = new Drawing.TableProperties { FirstRow = true, BandRow = true };

                // Apply table style if specified
                if (properties.TryGetValue("style", out var tblStyleVal))
                {
                    var styleId = ResolveTableStyleId(tblStyleVal);
                    tblProps.AppendChild(new Drawing.TableStyleId(styleId));
                }

                table.Append(tblProps);

                var tableGrid = new Drawing.TableGrid();
                for (int c = 0; c < cols; c++)
                    tableGrid.Append(new Drawing.GridColumn { Width = colWidth });
                table.Append(tableGrid);

                // Parse optional fill colors for header/body rows
                string? headerFillColor = null;
                if (properties.TryGetValue("headerFill", out var hfVal) || properties.TryGetValue("headerfill", out hfVal))
                    headerFillColor = ParseHelpers.SanitizeColorForOoxml(hfVal).Rgb;
                string? bodyFillColor = null;
                if (properties.TryGetValue("bodyFill", out var bfVal) || properties.TryGetValue("bodyfill", out bfVal))
                    bodyFillColor = ParseHelpers.SanitizeColorForOoxml(bfVal).Rgb;

                for (int r = 0; r < rows; r++)
                {
                    var tableRow = new Drawing.TableRow { Height = rowHeight };
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = new Drawing.TableCell();
                        var cellText = tableData != null && r < tableData.Length && c < tableData[r].Length
                            ? tableData[r][c] : (properties.TryGetValue($"r{r + 1}c{c + 1}", out var rc) ? rc : "");
                        var cellPara = new Drawing.Paragraph();
                        if (!string.IsNullOrEmpty(cellText))
                            cellPara.Append(new Drawing.Run(
                                new Drawing.RunProperties { Language = "en-US" },
                                new Drawing.Text { Text = cellText }));
                        else
                            cellPara.Append(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                        cell.Append(new Drawing.TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            cellPara
                        ));
                        var tcPr = new Drawing.TableCellProperties();
                        // Apply row-level fill: headerFill for row 0, bodyFill for others
                        var rowFill = (r == 0 ? headerFillColor : bodyFillColor);
                        if (rowFill != null)
                            tcPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rowFill }));
                        cell.Append(tcPr);
                        tableRow.Append(cell);
                    }
                    table.Append(tableRow);
                }

                var graphic = new Drawing.Graphic(
                    new Drawing.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
                );
                graphicFrame.Append(graphic);
                InsertAtPosition(tblShapeTree, graphicFrame, index);
                GetSlide(tblSlidePart).Save();

                var tblCount = tblShapeTree.Elements<GraphicFrame>()
                    .Count(gf => gf.Descendants<Drawing.Table>().Any());
                return $"/slide[{tblSlideIdx}]/{BuildElementPathSegment("table", graphicFrame, tblCount)}";
    }


    private string AddRow(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Resolve parent table via logical path
                var rowLogical = ResolveLogicalPath(parentPath);
                if (!rowLogical.HasValue || rowLogical.Value.element is not Drawing.Table rowTable)
                    throw new ArgumentException("Rows can only be added to a table: /slide[N]/table[M]");

                var rowSlidePart = rowLogical.Value.slidePart;

                // Determine column count from existing grid
                var existingColCount = rowTable.Elements<Drawing.TableGrid>().FirstOrDefault()
                    ?.Elements<Drawing.GridColumn>().Count() ?? 1;
                int newColCount = existingColCount;
                if (properties.TryGetValue("cols", out var rcVal))
                {
                    if (!int.TryParse(rcVal, out newColCount))
                        throw new ArgumentException($"Invalid 'cols' value: '{rcVal}'. Expected a positive integer.");
                }

                // Row height: default from first existing row, or 370840 EMU (~1cm)
                long newRowHeight = properties.TryGetValue("height", out var rhVal)
                    ? ParseEmu(rhVal)
                    : rowTable.Elements<Drawing.TableRow>().FirstOrDefault()?.Height?.Value ?? 370840;

                var newTblRow = new Drawing.TableRow { Height = newRowHeight };
                for (int c = 0; c < newColCount; c++)
                {
                    var newTblCell = new Drawing.TableCell();
                    var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
                    var bodyProps = new Drawing.BodyProperties();
                    var listStyle = new Drawing.ListStyle();
                    var cellPara = new Drawing.Paragraph();
                    if (!string.IsNullOrEmpty(cellText))
                        cellPara.Append(new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US" },
                            new Drawing.Text { Text = cellText }));
                    else
                        cellPara.Append(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                    newTblCell.Append(new Drawing.TextBody(bodyProps, listStyle, cellPara));
                    newTblCell.Append(new Drawing.TableCellProperties());
                    newTblRow.Append(newTblCell);
                }

                if (index.HasValue)
                {
                    var existingRows = rowTable.Elements<Drawing.TableRow>().ToList();
                    if (index.Value < existingRows.Count)
                        rowTable.InsertBefore(newTblRow, existingRows[index.Value]);
                    else
                        rowTable.AppendChild(newTblRow);
                }
                else
                {
                    rowTable.AppendChild(newTblRow);
                }

                // Update GraphicFrame container height to match sum of all row heights
                var graphicFrame = rowTable.Ancestors<GraphicFrame>().FirstOrDefault();
                if (graphicFrame?.Transform?.Extents != null)
                {
                    long totalRowHeight = rowTable.Elements<Drawing.TableRow>()
                        .Sum(r => r.Height?.Value ?? 370840);
                    graphicFrame.Transform.Extents.Cy = totalRowHeight;
                }

                GetSlide(rowSlidePart).Save();
                var rowIdx = rowTable.Elements<Drawing.TableRow>().ToList().IndexOf(newTblRow) + 1;
                return $"{parentPath}/tr[{rowIdx}]";
    }


    private string AddColumn(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Resolve parent table via logical path
                var colLogical = ResolveLogicalPath(parentPath);
                if (!colLogical.HasValue || colLogical.Value.element is not Drawing.Table colTable)
                    throw new ArgumentException("Columns can only be added to a table: /slide[N]/table[M]");

                var colSlidePart = colLogical.Value.slidePart;

                // Determine column width: specified or average of existing columns
                var tableGrid = colTable.GetFirstChild<Drawing.TableGrid>()
                    ?? colTable.AppendChild(new Drawing.TableGrid());
                var existingGridCols = tableGrid.Elements<Drawing.GridColumn>().ToList();
                long colWidth = properties.TryGetValue("width", out var wVal)
                    ? ParseEmu(wVal)
                    : (existingGridCols.Count > 0
                        ? (long)existingGridCols.Average(gc => gc.Width?.Value ?? 914400)
                        : 914400); // default ~2.54cm

                // Create and insert the new grid column
                var newGridCol = new Drawing.GridColumn { Width = colWidth };
                if (index.HasValue && index.Value < existingGridCols.Count)
                    tableGrid.InsertBefore(newGridCol, existingGridCols[index.Value]);
                else
                    tableGrid.AppendChild(newGridCol);

                var insertIdx = tableGrid.Elements<Drawing.GridColumn>().ToList().IndexOf(newGridCol);

                // Cell text from property
                var cellText = properties.GetValueOrDefault("text", "");

                // For each row, insert a new cell at the same column index
                foreach (var row in colTable.Elements<Drawing.TableRow>())
                {
                    var newCell = new Drawing.TableCell();
                    var cPara = new Drawing.Paragraph();
                    if (!string.IsNullOrEmpty(cellText))
                        cPara.Append(new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US" },
                            new Drawing.Text { Text = cellText }));
                    else
                        cPara.Append(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                    newCell.Append(new Drawing.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        cPara));
                    newCell.Append(new Drawing.TableCellProperties());

                    var existingCells = row.Elements<Drawing.TableCell>().ToList();
                    if (insertIdx < existingCells.Count)
                        row.InsertBefore(newCell, existingCells[insertIdx]);
                    else
                        row.AppendChild(newCell);
                }

                // Update GraphicFrame container width to match sum of all column widths
                var graphicFrame = colTable.Ancestors<GraphicFrame>().FirstOrDefault();
                if (graphicFrame?.Transform?.Extents != null)
                {
                    long totalColWidth = tableGrid.Elements<Drawing.GridColumn>()
                        .Sum(gc => gc.Width?.Value ?? 914400);
                    graphicFrame.Transform.Extents.Cx = totalColWidth;
                }

                GetSlide(colSlidePart).Save();
                var colIdx = tableGrid.Elements<Drawing.GridColumn>().ToList().IndexOf(newGridCol) + 1;
                return $"{parentPath}/col[{colIdx}]";
    }


    private string AddCell(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Resolve parent row via logical path
                var cellLogical = ResolveLogicalPath(parentPath);
                if (!cellLogical.HasValue || cellLogical.Value.element is not Drawing.TableRow cellRow)
                    throw new ArgumentException("Cells can only be added to a table row: /slide[N]/table[M]/tr[R]");

                var cellSlidePart = cellLogical.Value.slidePart;

                var newCell = new Drawing.TableCell();
                var cBodyProps = new Drawing.BodyProperties();
                var cListStyle = new Drawing.ListStyle();
                var cPara = new Drawing.Paragraph();
                if (properties.TryGetValue("text", out var cText) && !string.IsNullOrEmpty(cText))
                    cPara.Append(new Drawing.Run(
                        new Drawing.RunProperties { Language = "en-US" },
                        new Drawing.Text { Text = cText }));
                else
                    cPara.Append(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                newCell.Append(new Drawing.TextBody(cBodyProps, cListStyle, cPara));
                newCell.Append(new Drawing.TableCellProperties());

                if (index.HasValue)
                {
                    var existingCells = cellRow.Elements<Drawing.TableCell>().ToList();
                    if (index.Value < existingCells.Count)
                        cellRow.InsertBefore(newCell, existingCells[index.Value]);
                    else
                        cellRow.AppendChild(newCell);
                }
                else
                {
                    cellRow.AppendChild(newCell);
                }

                GetSlide(cellSlidePart).Save();
                var cellIdx = cellRow.Elements<Drawing.TableCell>().ToList().IndexOf(newCell) + 1;
                return $"{parentPath}/tc[{cellIdx}]";
    }


}
