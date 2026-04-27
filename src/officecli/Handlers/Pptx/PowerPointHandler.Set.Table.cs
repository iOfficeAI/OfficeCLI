// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for table paths. Mechanically extracted
// from the original god-method Set(); each helper owns one path-pattern's
// full handling. No behavior change.
public partial class PowerPointHandler
{
    private List<string> SetTableCellByPath(Match tblCellMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(tblCellMatch.Groups[1].Value);
        var tblIdx = int.Parse(tblCellMatch.Groups[2].Value);
        var rowIdx = int.Parse(tblCellMatch.Groups[3].Value);
        var cellIdx = int.Parse(tblCellMatch.Groups[4].Value);

        var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
        var tableRows = table.Elements<Drawing.TableRow>().ToList();
        if (rowIdx < 1 || rowIdx > tableRows.Count)
            throw new ArgumentException($"Row {rowIdx} not found (table has {tableRows.Count} rows)");
        var cells = tableRows[rowIdx - 1].Elements<Drawing.TableCell>().ToList();
        if (cellIdx < 1 || cellIdx > cells.Count)
            throw new ArgumentException($"Cell {cellIdx} not found (row has {cells.Count} cells)");

        var cell = cells[cellIdx - 1];
        // Clone cell for rollback on failure (atomic: no partial modifications)
        var cellBackup = cell.CloneNode(true);
        try
        {
            var unsupported = SetTableCellProperties(cell, properties);
            GetSlide(slidePart).Save();
            return unsupported;
        }
        catch
        {
            cell.Parent?.ReplaceChild(cellBackup, cell);
            throw;
        }
    }
    private List<string> SetTableRowByPath(Match tblRowMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(tblRowMatch.Groups[1].Value);
        var tblIdx = int.Parse(tblRowMatch.Groups[2].Value);
        var rowIdx = int.Parse(tblRowMatch.Groups[3].Value);

        var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
        var tableRows = table.Elements<Drawing.TableRow>().ToList();
        if (rowIdx < 1 || rowIdx > tableRows.Count)
            throw new ArgumentException($"Row {rowIdx} not found (table has {tableRows.Count} rows)");

        var row = tableRows[rowIdx - 1];
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "height":
                    row.Height = ParseEmu(value);
                    break;
                case "text":
                {
                    // Two behaviors based on presence of tab:
                    //  - No tab: broadcast the same text to all cells in the row
                    //  - Tab-delimited: distribute tokens across cells by position
                    //    ("X1\tX2\tX3" → tc[1]="X1", tc[2]="X2", tc[3]="X3")
                    // Extra tokens beyond cell count are dropped; cells beyond token
                    // count are left unchanged.
                    var rowCells = row.Elements<Drawing.TableCell>().ToList();
                    if (value.Contains('\t'))
                    {
                        var tokens = value.Split('\t');
                        for (int i = 0; i < rowCells.Count && i < tokens.Length; i++)
                            ReplaceCellText(rowCells[i], tokens[i]);
                    }
                    else
                    {
                        foreach (var c in rowCells)
                            ReplaceCellText(c, value);
                    }
                    break;
                }
                default:
                    // c1, c2, ... shorthand: set text of specific cell by index
                    if (key.Length >= 2 && key[0] == 'c' && int.TryParse(key.AsSpan(1), out var cIdx))
                    {
                        var rowCells = row.Elements<Drawing.TableCell>().ToList();
                        if (cIdx < 1 || cIdx > rowCells.Count)
                            throw new ArgumentException($"Cell c{cIdx} out of range (row has {rowCells.Count} cells)");
                        ReplaceCellText(rowCells[cIdx - 1], value);
                    }
                    else
                    {
                        // Apply to all cells in this row
                        var cellUnsup = new HashSet<string>();
                        foreach (var cell in row.Elements<Drawing.TableCell>())
                        {
                            var u = SetTableCellProperties(cell, new Dictionary<string, string> { { key, value } });
                            foreach (var k in u) cellUnsup.Add(k);
                        }
                        unsupported.AddRange(cellUnsup);
                    }
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }
    private List<string> SetTableByPath(Match tblMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(tblMatch.Groups[1].Value);
        var tblIdx = int.Parse(tblMatch.Groups[2].Value);

        var slideParts2 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts2.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts2.Count})");

        var slidePart = slideParts2[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var graphicFrames = shapeTree.Elements<GraphicFrame>()
            .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
        if (tblIdx < 1 || tblIdx > graphicFrames.Count)
            throw new ArgumentException($"Table {tblIdx} not found (total: {graphicFrames.Count})");

        var gf = graphicFrames[tblIdx - 1];
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "x" or "y" or "width" or "height":
                {
                    var xfrm = gf.Transform ?? (gf.Transform = new Transform());
                    TryApplyPositionSize(key.ToLowerInvariant(), value,
                        xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                        xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                    break;
                }
                case "name":
                    var nvPr = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                    if (nvPr != null) nvPr.Name = value;
                    break;
                case "tablestyle" or "style":
                {
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                            ?? table.PrependChild(new Drawing.TableProperties());
                        // Well-known style names → GUIDs
                        var styleId = ResolveTableStyleId(value);
                        tblPr.RemoveAllChildren<Drawing.TableStyleId>();
                        tblPr.AppendChild(new Drawing.TableStyleId(styleId));
                    }
                    break;
                }
                case "firstrow":
                case "lastrow":
                case "firstcol" or "firstcolumn":
                case "lastcol" or "lastcolumn":
                case "bandrow" or "bandedrows" or "bandrows":
                case "bandcol" or "bandedcols" or "bandcols":
                {
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                            ?? table.PrependChild(new Drawing.TableProperties());
                        var bv = IsTruthy(value);
                        switch (key.ToLowerInvariant())
                        {
                            case "firstrow": tblPr.FirstRow = bv; break;
                            case "lastrow": tblPr.LastRow = bv; break;
                            case "firstcol" or "firstcolumn": tblPr.FirstColumn = bv; break;
                            case "lastcol" or "lastcolumn": tblPr.LastColumn = bv; break;
                            case "bandrow" or "bandedrows" or "bandrows": tblPr.BandRow = bv; break;
                            case "bandcol" or "bandedcols" or "bandcols": tblPr.BandColumn = bv; break;
                        }
                    }
                    break;
                }
                case "colwidth" or "colwidths":
                {
                    // Set individual column widths: "3cm,5cm,3cm" or single value for all
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
                        if (gridCols != null && gridCols.Count > 0)
                        {
                            var widths = value.Split(',').Select(w => ParseEmu(w.Trim())).ToArray();
                            for (int ci = 0; ci < gridCols.Count; ci++)
                                gridCols[ci].Width = ci < widths.Length ? widths[ci] : widths[^1];
                        }
                    }
                    break;
                }
                case "autofit" or "autowidth":
                {
                    // Heuristic auto column width: measure max text length per column
                    if (!IsTruthy(value)) break;
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table == null) break;
                    var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
                    var tableRows = table.Elements<Drawing.TableRow>().ToList();
                    if (gridCols == null || gridCols.Count == 0 || tableRows.Count == 0) break;

                    var totalWidth = gridCols.Sum(gc => gc.Width?.Value ?? 0);
                    var colCount = gridCols.Count;
                    var maxLens = new int[colCount];
                    foreach (var row in tableRows)
                    {
                        var cells = row.Elements<Drawing.TableCell>().ToList();
                        for (int ci = 0; ci < Math.Min(cells.Count, colCount); ci++)
                        {
                            var text = cells[ci].TextBody?.InnerText ?? "";
                            maxLens[ci] = Math.Max(maxLens[ci], text.Length);
                        }
                    }
                    var totalLen = maxLens.Sum();
                    if (totalLen == 0) break;
                    // Minimum 10% per column, distribute rest by text length
                    var minWidth = totalWidth * 0.1 / colCount;
                    var distributable = totalWidth - minWidth * colCount;
                    for (int ci = 0; ci < colCount; ci++)
                        gridCols[ci].Width = (long)(minWidth + distributable * maxLens[ci] / totalLen);
                    break;
                }
                case "shadow":
                {
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                            ?? table.PrependChild(new Drawing.TableProperties());
                        var effectList = tblPr.GetFirstChild<Drawing.EffectList>();
                        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            effectList?.RemoveAllChildren<Drawing.OuterShadow>();
                            if (effectList?.ChildElements.Count == 0) effectList.Remove();
                        }
                        else
                        {
                            if (effectList == null) effectList = tblPr.AppendChild(new Drawing.EffectList());
                            effectList.RemoveAllChildren<Drawing.OuterShadow>();
                            var shadow = OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(value, BuildColorElement);
                            InsertEffectInOrder(effectList, shadow);
                        }
                    }
                    break;
                }
                case "glow":
                {
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                            ?? table.PrependChild(new Drawing.TableProperties());
                        var effectList = tblPr.GetFirstChild<Drawing.EffectList>();
                        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            effectList?.RemoveAllChildren<Drawing.Glow>();
                            if (effectList?.ChildElements.Count == 0) effectList.Remove();
                        }
                        else
                        {
                            if (effectList == null) effectList = tblPr.AppendChild(new Drawing.EffectList());
                            effectList.RemoveAllChildren<Drawing.Glow>();
                            var glow = OfficeCli.Core.DrawingEffectsHelper.BuildGlow(value, BuildColorElement);
                            InsertEffectInOrder(effectList, glow);
                        }
                    }
                    break;
                }
                case "bandcolor.odd" or "bandcolor.even":
                {
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var isOdd = key.ToLowerInvariant().EndsWith("odd");
                        var rows = table.Elements<Drawing.TableRow>().ToList();
                        for (int ri = 0; ri < rows.Count; ri++)
                        {
                            bool matchesOddEven = isOdd ? (ri % 2 == 0) : (ri % 2 == 1); // 0-based: odd rows are 0,2,4...
                            if (matchesOddEven)
                            {
                                foreach (var cell in rows[ri].Elements<Drawing.TableCell>())
                                    SetTableCellProperties(cell, new Dictionary<string, string> { { "fill", value } });
                            }
                        }
                    }
                    break;
                }
                case var k when k.StartsWith("border") || k is "text" or "bold" or "italic" or "size" or "font" or "color" or "underline" or "strike" or "valign" or "fill" or "baseline" or "charspacing" or "opacity" or "bevel" or "margin" or "padding" or "textdirection" or "wordwrap" or "linespacing" or "spacebefore" or "spaceafter":
                {
                    // Apply cell-level properties to all cells in the table
                    var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        foreach (var cell in table.Descendants<Drawing.TableCell>())
                        {
                            var u = SetTableCellProperties(cell, new Dictionary<string, string> { { key, value } });
                            foreach (var uk in u) { if (!unsupported.Contains(uk)) unsupported.Add(uk); }
                        }
                    }
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(gf, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid table props: x, y, width, height, name, style, firstRow, lastRow, firstCol, lastCol, bandedRows, bandedCols, colWidths)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

}
