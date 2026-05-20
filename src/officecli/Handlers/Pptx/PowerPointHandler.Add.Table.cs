// Copyright 2025 OfficeCLI (officecli.ai)
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
                    // Prefix "Invalid" so OutputFormatter maps this to
                    // invalid_value rather than the internal_error catch-all.
                    throw new ArgumentException($"Invalid table dimensions: rows={rows}, cols={cols}. Both must be >= 1.");

                // BUG-R6-D: enforce a practical upper bound on rows/cols so the
                // EMU height/width calculations stay safely within int32 (the
                // OOXML cy/cx attributes are int32). With the default rowHeight
                // of 370840 EMU, int.MaxValue / 370840 ≈ 5790. Cap rows/cols at
                // 5000 — well within OOXML practical limits and prevents the
                // negative-cy schema-invalid output that 99999 rows produced.
                const int MaxTableDim = 5000;
                if (rows > MaxTableDim)
                    throw new ArgumentException($"rows={rows} exceeds practical maximum ({MaxTableDim}); reduce rows or split into multiple tables.");
                if (cols > MaxTableDim)
                    throw new ArgumentException($"cols={cols} exceeds practical maximum ({MaxTableDim}); reduce cols or split into multiple tables.");

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

                var tblId = AcquireShapeId(tblShapeTree, properties);

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
                var tblProps = new Drawing.TableProperties();

                // tblLook props: read overrides from properties, with default firstRow/bandRow=true.
                static bool? ReadBoolProp(Dictionary<string, string> p, params string[] keys)
                {
                    foreach (var k in keys)
                        if (p.TryGetValue(k, out var v))
                            return IsTruthy(v);
                    return null;
                }
                tblProps.FirstRow    = ReadBoolProp(properties, "firstRow", "firstrow") ?? true;
                tblProps.BandRow     = ReadBoolProp(properties, "bandedRows", "bandedrows", "bandRow", "bandrow") ?? true;
                var lastRowProp      = ReadBoolProp(properties, "lastRow", "lastrow");
                if (lastRowProp.HasValue) tblProps.LastRow = lastRowProp.Value;
                var firstColProp     = ReadBoolProp(properties, "firstCol", "firstcol", "firstColumn", "firstcolumn");
                if (firstColProp.HasValue) tblProps.FirstColumn = firstColProp.Value;
                var lastColProp      = ReadBoolProp(properties, "lastCol", "lastcol", "lastColumn", "lastcolumn");
                if (lastColProp.HasValue) tblProps.LastColumn = lastColProp.Value;
                var bandColProp      = ReadBoolProp(properties, "bandedCols", "bandedcols", "bandCol", "bandcol", "bandColumn", "bandcolumn");
                if (bandColProp.HasValue) tblProps.BandColumn = bandColProp.Value;

                // Apply table style if specified
                if (properties.TryGetValue("style", out var tblStyleVal))
                {
                    var styleId = ResolveTableStyleId(tblStyleVal);
                    tblProps.AppendChild(new Drawing.TableStyleId(styleId));
                }

                table.Append(tblProps);

                // Optional explicit colWidths (semicolon- or comma-separated EMU/cm/pt values).
                long[]? explicitColWidths = null;
                if (properties.TryGetValue("colWidths", out var cwStr) || properties.TryGetValue("colwidths", out cwStr))
                {
                    var parts = cwStr.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    explicitColWidths = parts.Select(p => ParseEmu(p.Trim())).ToArray();
                }

                var tableGrid = new Drawing.TableGrid();
                for (int c = 0; c < cols; c++)
                {
                    var w = (explicitColWidths != null && c < explicitColWidths.Length) ? explicitColWidths[c] : colWidth;
                    tableGrid.Append(new Drawing.GridColumn { Width = w });
                }
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
                        XmlTextValidator.ValidateOrThrow(cellText, $"r{r + 1}c{c + 1}");
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

                // CONSISTENCY(add-set-parity): border-prefixed props on AddTable
                // delegate to the same fan-out used by Set. PPT OOXML has no
                // table-level border element — borders are per-cell lnL/lnR/lnT/lnB,
                // so border.all / border.top / etc. are applied to every cell.
                // border.horizontal / border.vertical mean inside row/column dividers.
                var tblBorderProps = properties
                    .Where(kv => kv.Key.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (tblBorderProps.Count > 0)
                    ApplyTableBorderFanOut(table, tblBorderProps);

                if (properties.TryGetValue("zorder", out var tblZ)
                    || properties.TryGetValue("z-order", out tblZ)
                    || properties.TryGetValue("order", out tblZ))
                    ApplyZOrder(tblSlidePart, graphicFrame, tblZ);

                GetSlide(tblSlidePart).Save();

                var tblCount = tblShapeTree.Elements<GraphicFrame>()
                    .Count(gf => gf.Descendants<Drawing.Table>().Any());
                return $"/slide[{tblSlideIdx}]/{BuildElementPathSegment("table", graphicFrame, tblCount)}";
    }


    // Apply table-level border properties by fan-out to per-cell lnL/lnR/lnT/lnB.
    // PPT OOXML has no table-level border element; "table border" is the union
    // of cell borders along the outer edges (and optionally inside dividers).
    //
    // Semantics:
    //   border / border.all              → every edge of every cell
    //   border.top                       → top of cells in row 1
    //   border.bottom                    → bottom of cells in last row
    //   border.left                      → left of cells in column 1
    //   border.right                     → right of cells in last column
    //   border.horizontal / border.insideH → bottom of rows 1..N-1 + top of rows 2..N
    //   border.vertical   / border.insideV → right of cols 1..M-1 + left of cols 2..M
    //   border.tl2br / border.tr2bl      → diagonals on every cell
    // Each can also use split form: border.top.width, border.left.color, etc.
    internal static void ApplyTableBorderFanOut(Drawing.Table table, Dictionary<string, string> borderProps)
    {
        var rows = table.Elements<Drawing.TableRow>().ToList();
        if (rows.Count == 0) return;
        int colCount = rows.Max(r => r.Elements<Drawing.TableCell>().Count());
        if (colCount == 0) return;

        foreach (var (rawKey, value) in borderProps)
        {
            var key = rawKey.ToLowerInvariant();

            bool isAll = key is "border" or "border.all";
            bool isTop = key.StartsWith("border.top");
            bool isBottom = key.StartsWith("border.bottom");
            bool isLeft = key.StartsWith("border.left");
            bool isRight = key.StartsWith("border.right");
            bool isInsideH = key.StartsWith("border.horizontal") || key.StartsWith("border.insideh");
            bool isInsideV = key.StartsWith("border.vertical")   || key.StartsWith("border.insidev");
            bool isDiag = key.StartsWith("border.tl2br") || key.StartsWith("border.tr2bl");

            // Split-form suffix preserved on cell-level key (e.g. ".width" / ".color" / ".dash").
            string splitSuffix = "";
            foreach (var s in new[] { ".width", ".color", ".dash" })
                if (key.EndsWith(s)) { splitSuffix = s; break; }

            void ApplyToCell(Drawing.TableCell cell, string edgeKey)
            {
                var cellKey = edgeKey + splitSuffix;
                SetTableCellProperties(cell, new Dictionary<string, string> { { cellKey, value } });
            }

            if (isAll)
            {
                foreach (var row in rows)
                    foreach (var cell in row.Elements<Drawing.TableCell>())
                        ApplyToCell(cell, "border.all");
                continue;
            }
            if (isDiag)
            {
                var diagEdge = key.StartsWith("border.tl2br") ? "border.tl2br" : "border.tr2bl";
                foreach (var row in rows)
                    foreach (var cell in row.Elements<Drawing.TableCell>())
                        ApplyToCell(cell, diagEdge);
                continue;
            }
            if (isTop)
            {
                foreach (var cell in rows[0].Elements<Drawing.TableCell>())
                    ApplyToCell(cell, "border.top");
                continue;
            }
            if (isBottom)
            {
                foreach (var cell in rows[^1].Elements<Drawing.TableCell>())
                    ApplyToCell(cell, "border.bottom");
                continue;
            }
            if (isLeft)
            {
                foreach (var row in rows)
                {
                    var firstCell = row.Elements<Drawing.TableCell>().FirstOrDefault();
                    if (firstCell != null) ApplyToCell(firstCell, "border.left");
                }
                continue;
            }
            if (isRight)
            {
                foreach (var row in rows)
                {
                    var lastCell = row.Elements<Drawing.TableCell>().LastOrDefault();
                    if (lastCell != null) ApplyToCell(lastCell, "border.right");
                }
                continue;
            }
            if (isInsideH)
            {
                // Apply to bottom of rows[0..N-2] and top of rows[1..N-1].
                for (int r = 0; r < rows.Count - 1; r++)
                {
                    foreach (var cell in rows[r].Elements<Drawing.TableCell>())
                        ApplyToCell(cell, "border.bottom");
                    foreach (var cell in rows[r + 1].Elements<Drawing.TableCell>())
                        ApplyToCell(cell, "border.top");
                }
                continue;
            }
            if (isInsideV)
            {
                foreach (var row in rows)
                {
                    var cells = row.Elements<Drawing.TableCell>().ToList();
                    for (int c = 0; c < cells.Count - 1; c++)
                    {
                        ApplyToCell(cells[c], "border.right");
                        ApplyToCell(cells[c + 1], "border.left");
                    }
                }
                continue;
            }
            // Unknown border.* key — ignore (Set table dispatch already validates).
        }
    }

    private string AddRow(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Resolve parent table via logical path
                var rowLogical = ResolveLogicalPath(parentPath);
                if (!rowLogical.HasValue || rowLogical.Value.element is not Drawing.Table rowTable)
                    throw new ArgumentException("Rows can only be added to a table: /slide[N]/table[M]");

                var rowSlidePart = rowLogical.Value.slidePart;

                // Row width is fixed by the parent <a:tblGrid>. OOXML requires
                // every <a:tr> to have <a:tc> count (gridSpan-summed) equal to
                // <a:tblGrid> column count — there is no "narrower row" concept.
                // We always emit `existingColCount` cells; user-supplied c1..cN
                // populate text and unspecified positions stay empty (legal).
                // An explicit `cols=` is only accepted when it matches the grid
                // (no-op) — any other value is a misuse caused by treating a
                // row as a width container, which OOXML doesn't model.
                var existingColCount = rowTable.Elements<Drawing.TableGrid>().FirstOrDefault()
                    ?.Elements<Drawing.GridColumn>().Count() ?? 1;
                if (properties.TryGetValue("cols", out var rcVal))
                {
                    if (!int.TryParse(rcVal, out var requested))
                        throw new ArgumentException($"Invalid 'cols' value: '{rcVal}'. Expected a positive integer.");
                    if (requested != existingColCount)
                        throw new ArgumentException(
                            $"cols={requested} does not match the table's grid ({existingColCount} columns). " +
                            "A row's cell count is fixed by <a:tblGrid> and cannot be narrower or wider. " +
                            "Omit --prop cols and leave unused c1..cN empty, or use --prop gridSpan=N on c1 " +
                            "(via set tr[i]/tc[1]) to merge cells into a wider span.");
                }
                int newColCount = existingColCount;

                // Row height: default from first existing row, or 370840 EMU (~1cm).
                // CONSISTENCY(positive-size): ST_TableCellSize disallows negatives.
                long newRowHeight = properties.TryGetValue("height", out var rhVal)
                    ? ParseEmu(rhVal)
                    : rowTable.Elements<Drawing.TableRow>().FirstOrDefault()?.Height?.Value ?? 370840;
                if (newRowHeight < 0)
                    throw new ArgumentException(
                        $"Invalid height '{rhVal}': table row height cannot be negative.");

                var newTblRow = new Drawing.TableRow { Height = newRowHeight };
                for (int c = 0; c < newColCount; c++)
                {
                    var newTblCell = new Drawing.TableCell();
                    var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
                    XmlTextValidator.ValidateOrThrow(cellText, $"c{c + 1}");
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
                // CONSISTENCY(positive-size): ST_PositiveSize2D disallows negatives.
                if (colWidth < 0)
                    throw new ArgumentException(
                        $"Invalid width '{wVal}': table column width cannot be negative.");

                // Create and insert the new grid column
                var newGridCol = new Drawing.GridColumn { Width = colWidth };
                if (index.HasValue && index.Value < existingGridCols.Count)
                    tableGrid.InsertBefore(newGridCol, existingGridCols[index.Value]);
                else
                    tableGrid.AppendChild(newGridCol);

                var insertIdx = tableGrid.Elements<Drawing.GridColumn>().ToList().IndexOf(newGridCol);

                // Cell text from property
                var cellText = properties.GetValueOrDefault("text", "");
                XmlTextValidator.ValidateOrThrow(cellText, "text");

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

                // Reject cell-append that would make the row wider than the
                // table's <a:tblGrid>. Real PowerPoint silently DROPS cells
                // beyond the gridCol count on render, so the user's content
                // would be lost. The user almost certainly meant "add column"
                // (which atomically grows tblGrid AND pads every sibling row);
                // surface that explicitly rather than silently corrupting.
                var parentTable = cellRow.Ancestors<Drawing.Table>().FirstOrDefault();
                var gridColCount = parentTable?.GetFirstChild<Drawing.TableGrid>()
                    ?.Elements<Drawing.GridColumn>().Count() ?? 0;
                var currentCellCount = cellRow.Elements<Drawing.TableCell>().Count();
                if (gridColCount > 0 && currentCellCount >= gridColCount)
                {
                    throw new ArgumentException(
                        $"Row already has {currentCellCount} cell(s); table grid has {gridColCount} column(s). " +
                        "Appending another cell would make the row wider than the grid and real PowerPoint " +
                        "would silently drop the orphan on render. Use `add /slide[N]/table[M] --type column` " +
                        "to grow the table rectangularly instead.");
                }

                var newCell = new Drawing.TableCell();
                var cBodyProps = new Drawing.BodyProperties();
                var cListStyle = new Drawing.ListStyle();
                var cPara = new Drawing.Paragraph();
                if (properties.TryGetValue("text", out var cText) && !string.IsNullOrEmpty(cText))
                {
                    XmlTextValidator.ValidateOrThrow(cText, "text");
                    cPara.Append(new Drawing.Run(
                        new Drawing.RunProperties { Language = "en-US" },
                        new Drawing.Text { Text = cText }));
                }
                else
                    cPara.Append(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                newCell.Append(new Drawing.TextBody(cBodyProps, cListStyle, cPara));
                newCell.Append(new Drawing.TableCellProperties());

                // CONSISTENCY(add-set-parity): fill / background applied at Add time
                // by delegating to SetTableCellProperties — same builder, same schema
                // ordering, no divergence between Add and Set.
                if (properties.TryGetValue("fill", out var cFill)
                    || properties.TryGetValue("background", out cFill))
                {
                    SetTableCellProperties(newCell, new Dictionary<string, string> { { "fill", cFill } });
                }

                // CONSISTENCY(add-set-parity): border-prefixed props on AddCell
                // delegate to SetTableCellProperties — same builder, same schema
                // ordering. Excludes border.horizontal/border.vertical which only
                // make sense at table level (inside-row / inside-column dividers).
                var addCellBorderProps = properties
                    .Where(kv => kv.Key.StartsWith("border", StringComparison.OrdinalIgnoreCase)
                        && !kv.Key.Equals("border.horizontal", StringComparison.OrdinalIgnoreCase)
                        && !kv.Key.Equals("border.vertical", StringComparison.OrdinalIgnoreCase)
                        && !kv.Key.Equals("border.insideh", StringComparison.OrdinalIgnoreCase)
                        && !kv.Key.Equals("border.insidev", StringComparison.OrdinalIgnoreCase)
                        && !kv.Key.Equals("border.insideH", StringComparison.OrdinalIgnoreCase)
                        && !kv.Key.Equals("border.insideV", StringComparison.OrdinalIgnoreCase))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (addCellBorderProps.Count > 0)
                    SetTableCellProperties(newCell, addCellBorderProps);

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
