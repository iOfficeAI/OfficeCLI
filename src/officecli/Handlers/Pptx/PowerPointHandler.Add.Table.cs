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

                var rowsStr = properties.GetValueOrDefault("rows", "3");
                var colsStr = properties.GetValueOrDefault("cols", "3");
                if (!int.TryParse(rowsStr, out var rows))
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Expected a positive integer.");
                if (!int.TryParse(colsStr, out var cols))
                    throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Expected a positive integer.");
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

                var tblId = (uint)(tblShapeTree.ChildElements.Count + 2);

                // Build GraphicFrame
                var graphicFrame = new GraphicFrame();
                graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = tblId, Name = properties.GetValueOrDefault("name", $"Table {tblId}") },
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
                    var styleId = tblStyleVal.ToLowerInvariant() switch
                    {
                        "medium1" or "mediumstyle1" => "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
                        "medium2" or "mediumstyle2" => "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
                        "medium3" or "mediumstyle3" => "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
                        "medium4" or "mediumstyle4" => "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
                        "light1" or "lightstyle1" => "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
                        "light2" or "lightstyle2" => "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
                        "light3" or "lightstyle3" => "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
                        "dark1" or "darkstyle1" => "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
                        "dark2" or "darkstyle2" => "{125E5076-3810-47DD-B79F-674D7AD40C01}",
                        "none" => "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
                        _ when tblStyleVal.StartsWith("{") => tblStyleVal,
                        _ => tblStyleVal
                    };
                    tblProps.AppendChild(new Drawing.TableStyleId(styleId));
                }

                table.Append(tblProps);

                var tableGrid = new Drawing.TableGrid();
                for (int c = 0; c < cols; c++)
                    tableGrid.Append(new Drawing.GridColumn { Width = colWidth });
                table.Append(tableGrid);

                for (int r = 0; r < rows; r++)
                {
                    var tableRow = new Drawing.TableRow { Height = rowHeight };
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = new Drawing.TableCell();
                        cell.Append(new Drawing.TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(new Drawing.EndParagraphRunProperties { Language = "en-US" })
                        ));
                        cell.Append(new Drawing.TableCellProperties());
                        tableRow.Append(cell);
                    }
                    table.Append(tableRow);
                }

                var graphic = new Drawing.Graphic(
                    new Drawing.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
                );
                graphicFrame.Append(graphic);
                tblShapeTree.AppendChild(graphicFrame);
                GetSlide(tblSlidePart).Save();

                var tblCount = tblShapeTree.Elements<GraphicFrame>()
                    .Count(gf => gf.Descendants<Drawing.Table>().Any());
                return $"/slide[{tblSlideIdx}]/table[{tblCount}]";
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
                            new Drawing.Text(cellText)));
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

                GetSlide(rowSlidePart).Save();
                var rowIdx = rowTable.Elements<Drawing.TableRow>().ToList().IndexOf(newTblRow) + 1;
                return $"{parentPath}/tr[{rowIdx}]";
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
                        new Drawing.Text(cText)));
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
