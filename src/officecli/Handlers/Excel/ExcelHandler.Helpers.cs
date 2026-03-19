// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Private Helpers ====================

    private static Worksheet GetSheet(WorksheetPart part) =>
        part.Worksheet ?? throw new InvalidOperationException("Corrupt file: worksheet data missing");

    /// <summary>
    /// Save worksheet with automatic schema-order reorder.
    /// Must be used instead of ws.Save() to prevent element ordering violations.
    /// </summary>
    private static void SaveWorksheet(WorksheetPart part)
    {
        ReorderWorksheetChildren(GetSheet(part));
        GetSheet(part).Save();
    }

    /// <summary>
    /// Reorder worksheet children to match OpenXML schema sequence.
    /// Schema: sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData,
    ///   autoFilter, sortState, mergeCells, conditionalFormatting,
    ///   dataValidations, hyperlinks, printOptions, pageMargins, pageSetup,
    ///   headerFooter, drawing, legacyDrawing, tableParts, extLst
    /// </summary>
    private static void ReorderWorksheetChildren(Worksheet ws)
    {
        var order = new Dictionary<string, int>
        {
            ["sheetPr"] = 0, ["dimension"] = 1, ["sheetViews"] = 2, ["sheetFormatPr"] = 3,
            ["cols"] = 4, ["sheetData"] = 5, ["sheetCalcPr"] = 6, ["sheetProtection"] = 7,
            ["protectedRanges"] = 8, ["scenarios"] = 9, ["autoFilter"] = 10, ["sortState"] = 11,
            ["dataConsolidate"] = 12, ["customSheetViews"] = 13, ["mergeCells"] = 14,
            ["phoneticPr"] = 15, ["conditionalFormatting"] = 16, ["dataValidations"] = 17,
            ["hyperlinks"] = 18, ["printOptions"] = 19, ["pageMargins"] = 20,
            ["pageSetup"] = 21, ["headerFooter"] = 22, ["rowBreaks"] = 23, ["colBreaks"] = 24,
            ["drawing"] = 25, ["legacyDrawing"] = 26, ["tableParts"] = 27, ["extLst"] = 99
        };

        var children = ws.ChildElements.ToList();
        var sorted = children
            .OrderBy(c => order.TryGetValue(c.LocalName, out var idx) ? idx : 50)
            .ToList();

        bool needsReorder = false;
        for (int i = 0; i < children.Count; i++)
        {
            if (!ReferenceEquals(children[i], sorted[i]))
            {
                needsReorder = true;
                break;
            }
        }

        if (needsReorder)
        {
            foreach (var child in children) child.Remove();
            foreach (var child in sorted) ws.AppendChild(child);
        }
    }

    private Workbook GetWorkbook() =>
        _doc.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Corrupt file: workbook missing");

    private List<(string Name, WorksheetPart Part)> GetWorksheets()
    {
        var result = new List<(string, WorksheetPart)>();
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return result;

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return result;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var id = sheet.Id?.Value;
            if (id == null) continue;
            var part = (WorksheetPart)_doc.WorkbookPart!.GetPartById(id);
            result.Add((name, part));
        }

        return result;
    }

    private WorksheetPart? FindWorksheet(string sheetName)
    {
        foreach (var (name, part) in GetWorksheets())
        {
            if (name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                return part;
        }
        return null;
    }

    private string GetCellDisplayValue(Cell cell)
    {
        var value = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sst?.SharedStringTable != null && int.TryParse(value, out int idx))
            {
                var item = sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? value;
            }
        }

        // Formula cells without cached value: show the formula
        if (string.IsNullOrEmpty(value) && cell.CellFormula != null
            && !string.IsNullOrEmpty(cell.CellFormula.Text))
        {
            return $"={cell.CellFormula.Text}";
        }

        return value;
    }

    private List<DocumentNode> GetSheetChildNodes(string sheetName, SheetData sheetData, int depth, WorksheetPart? worksheetPart = null)
    {
        var children = new List<DocumentNode>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = row.RowIndex?.Value ?? 0;
            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{rowIdx}]",
                Type = "row",
                ChildCount = row.Elements<Cell>().Count()
            };
            if (row.Height?.Value != null)
                rowNode.Format["height"] = row.Height.Value;
            if (row.Hidden?.Value == true)
                rowNode.Format["hidden"] = true;

            if (depth > 0)
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    rowNode.Children.Add(CellToNode(sheetName, cell));
                }
            }

            children.Add(rowNode);
        }

        // Add chart children from DrawingsPart (following Apache POI pattern)
        if (worksheetPart?.DrawingsPart != null)
        {
            var chartParts = worksheetPart.DrawingsPart.ChartParts.ToList();
            for (int i = 0; i < chartParts.Count; i++)
            {
                var chart = chartParts[i].ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                var chartNode = new DocumentNode
                {
                    Path = $"/{sheetName}/chart[{i + 1}]",
                    Type = "chart"
                };
                if (chart != null)
                    ChartHelper.ReadChartProperties(chart, chartNode, 0);
                children.Add(chartNode);
            }
        }

        return children;
    }

    private DocumentNode CellToNode(string sheetName, Cell cell, WorksheetPart? part = null)
    {
        var cellRef = cell.CellReference?.Value ?? "?";
        var value = GetCellDisplayValue(cell);
        var formula = cell.CellFormula?.Text;
        string type;
        if (cell.DataType?.HasValue != true)
            type = "Number";
        else if (cell.DataType.Value == CellValues.String)
            type = "String";
        else if (cell.DataType.Value == CellValues.SharedString)
            type = "SharedString";
        else if (cell.DataType.Value == CellValues.Boolean)
            type = "Boolean";
        else if (cell.DataType.Value == CellValues.Error)
            type = "Error";
        else if (cell.DataType.Value == CellValues.InlineString)
            type = "InlineString";
        else if (cell.DataType.Value == CellValues.Date)
            type = "Date";
        else
            type = "Number";

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{cellRef}",
            Type = "cell",
            Text = value,
            Preview = cellRef
        };

        node.Format["type"] = type;
        if (formula != null) node.Format["formula"] = formula;
        if (string.IsNullOrEmpty(value)) node.Format["empty"] = true;

        // Hyperlink readback
        if (part != null)
        {
            var hyperlink = GetSheet(part).GetFirstChild<Hyperlinks>()?.Elements<Hyperlink>()
                .FirstOrDefault(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
            if (hyperlink?.Id?.Value != null)
            {
                try
                {
                    var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id.Value);
                    if (rel != null)
                    {
                        var linkStr = rel.Uri.OriginalString;
                        // Strip trailing slash added by Uri normalization for bare authority URLs
                        if (linkStr.EndsWith("/") && rel.Uri.IsAbsoluteUri && rel.Uri.AbsolutePath == "/")
                            linkStr = linkStr.TrimEnd('/');
                        node.Format["link"] = linkStr;
                    }
                }
                catch { }
            }

            // Border readback from stylesheet
            var styleIndex = cell.StyleIndex?.Value ?? 0;
            var wbStylesPart = _doc.WorkbookPart?.WorkbookStylesPart;
            if (wbStylesPart?.Stylesheet != null && styleIndex > 0)
            {
                var cellFormats = wbStylesPart.Stylesheet.CellFormats;
                if (cellFormats != null && styleIndex < (uint)cellFormats.Elements<CellFormat>().Count())
                {
                    var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
                    // Font readback
                    var fontId = xf.FontId?.Value ?? 0;
                    if (fontId > 0)
                    {
                        var fonts = wbStylesPart.Stylesheet.Fonts;
                        if (fonts != null && fontId < (uint)fonts.Elements<Font>().Count())
                        {
                            var font = fonts.Elements<Font>().ElementAt((int)fontId);
                            if (font.Bold != null) { node.Format["font.bold"] = true; node.Format["bold"] = true; }
                            if (font.Italic != null)
                            {
                                node.Format["font.italic"] = true;
                                node.Format["italic"] = true;
                            }
                            if (font.Strike != null) node.Format["font.strike"] = true;
                            if (font.Underline != null)
                                node.Format["font.underline"] = font.Underline.Val?.InnerText == "double" ? "double" : "single";
                            if (font.Color?.Rgb?.Value != null)
                            {
                                var rgbVal = font.Color.Rgb.Value;
                                // Strip ARGB alpha prefix (e.g. "FFFF0000" → "FF0000") for consistency with Word/PPTX
                                if (rgbVal.Length == 8 && rgbVal.StartsWith("FF", StringComparison.OrdinalIgnoreCase))
                                    rgbVal = rgbVal[2..];
                                node.Format["font.color"] = rgbVal;
                            }
                            if (font.FontSize?.Val?.Value != null)
                                node.Format["font.size"] = $"{font.FontSize.Val.Value:0.##}pt";
                            if (font.FontName?.Val?.Value != null) node.Format["font.name"] = font.FontName.Val.Value;
                        }
                    }

                    // Fill readback
                    var fillId = xf.FillId?.Value ?? 0;
                    if (fillId > 0)
                    {
                        var fills = wbStylesPart.Stylesheet.Fills;
                        if (fills != null && fillId < (uint)fills.Elements<Fill>().Count())
                        {
                            var fill = fills.Elements<Fill>().ElementAt((int)fillId);
                            // Check gradient fill first
                            var gf = fill.GetFirstChild<GradientFill>();
                            if (gf != null)
                            {
                                var stops = gf.Elements<GradientStop>().ToList();
                                if (stops.Count >= 2)
                                {
                                    var c1 = stops[0].Color?.Rgb?.Value ?? "";
                                    var c2 = stops[^1].Color?.Rgb?.Value ?? "";
                                    // Strip FF alpha prefix
                                    if (c1.Length == 8 && c1.StartsWith("FF", StringComparison.OrdinalIgnoreCase)) c1 = c1[2..];
                                    if (c2.Length == 8 && c2.StartsWith("FF", StringComparison.OrdinalIgnoreCase)) c2 = c2[2..];
                                    int deg = (int)(gf.Degree?.Value ?? 0);
                                    node.Format["fill"] = $"gradient;{c1};{c2};{deg}";
                                }
                            }
                            else
                            {
                                var pf = fill.PatternFill;
                                if (pf?.ForegroundColor?.Rgb?.Value != null)
                                {
                                    var fillRgb = pf.ForegroundColor.Rgb.Value;
                                    if (fillRgb.Length == 8 && fillRgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase))
                                        fillRgb = fillRgb[2..];
                                    node.Format["fill"] = fillRgb;
                                }
                            }
                        }
                    }

                    var borderId = xf.BorderId?.Value ?? 0;
                    if (borderId > 0)
                    {
                        var borders = wbStylesPart.Stylesheet.Borders;
                        if (borders != null && borderId < (uint)borders.Elements<Border>().Count())
                        {
                            var border = borders.Elements<Border>().ElementAt((int)borderId);
                            var sides = new (string name, BorderPropertiesType? bp)[] {
                                ("left", border.LeftBorder), ("right", border.RightBorder),
                                ("top", border.TopBorder), ("bottom", border.BottomBorder)
                            };
                            foreach (var (side, b) in sides)
                            {
                                if (b?.Style?.Value != null && b.Style.Value != BorderStyleValues.None)
                                {
                                    node.Format[$"border.{side}"] = b.Style.InnerText;
                                    if (b.Color?.Rgb?.Value != null)
                                    {
                                        var borderRgb = b.Color.Rgb.Value!;
                                        if (borderRgb.Length == 8 && borderRgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase))
                                            borderRgb = borderRgb[2..];
                                        node.Format[$"border.{side}.color"] = borderRgb;
                                    }
                                }
                            }
                            // Diagonal border readback
                            var diag = border.DiagonalBorder;
                            if (diag?.Style?.Value != null && diag.Style.Value != BorderStyleValues.None)
                            {
                                node.Format["border.diagonal"] = diag.Style.InnerText;
                                if (diag.Color?.Rgb?.Value != null)
                                {
                                    var diagRgb = diag.Color.Rgb.Value!;
                                    if (diagRgb.Length == 8 && diagRgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase))
                                        diagRgb = diagRgb[2..];
                                    node.Format["border.diagonal.color"] = diagRgb;
                                }
                            }
                            if (border.DiagonalUp?.Value == true)
                                node.Format["border.diagonalUp"] = true;
                            if (border.DiagonalDown?.Value == true)
                                node.Format["border.diagonalDown"] = true;
                        }
                    }

                    // Alignment + wrap readback (like POI XSSFCellStyle.getWrapText)
                    var alignment = xf.Alignment;
                    if (alignment != null)
                    {
                        if (alignment.WrapText?.Value == true)
                            node.Format["alignment.wrapText"] = true;
                        if (alignment.Horizontal?.HasValue == true)
                            node.Format["alignment.horizontal"] = alignment.Horizontal.InnerText;
                        if (alignment.Vertical?.HasValue == true)
                        {
                            node.Format["alignment.vertical"] = alignment.Vertical.InnerText;
                            node.Format["valign"] = alignment.Vertical.InnerText;
                        }
                    }

                    // Number format readback
                    var numFmtId = xf.NumberFormatId?.Value ?? 0;
                    if (numFmtId > 0)
                    {
                        node.Format["numFmtId"] = numFmtId;
                        var numFmts = wbStylesPart.Stylesheet.NumberingFormats;
                        var customFmt = numFmts?.Elements<NumberingFormat>()
                            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
                        object fmtVal;
                        if (customFmt?.FormatCode?.Value != null)
                            fmtVal = customFmt.FormatCode.Value;
                        else
                        {
                            // Resolve built-in number format IDs to their format strings
                            // See ECMA-376 Part 1, 18.8.30 (numFmt) for built-in IDs
                            fmtVal = numFmtId switch
                            {
                                1 => "0",
                                2 => "0.00",
                                3 => "#,##0",
                                4 => "#,##0.00",
                                9 => "0%",
                                10 => "0.00%",
                                11 => "0.00E+00",
                                12 => "# ?/?",
                                13 => "# ??/??",
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
                                38 => "#,##0 ;[Red](#,##0)",
                                39 => "#,##0.00;(#,##0.00)",
                                40 => "#,##0.00;[Red](#,##0.00)",
                                45 => "mm:ss",
                                46 => "[h]:mm:ss",
                                47 => "mmss.0",
                                48 => "##0.0E+0",
                                49 => "@",
                                _ => (object)numFmtId // fallback to ID for truly unknown formats
                            };
                        }
                        node.Format["numberformat"] = fmtVal;
                    }
                }
            }

            // Merge cell readback
            var mergeCells = GetSheet(part).GetFirstChild<MergeCells>();
            if (mergeCells != null)
            {
                var mergeCell = mergeCells.Elements<MergeCell>()
                    .FirstOrDefault(m => IsCellInMergeRange(cellRef, m.Reference?.Value));
                if (mergeCell != null)
                {
                    node.Format["merge"] = mergeCell.Reference?.Value ?? "";
                }
            }
        }

        return node;
    }

    private static bool IsCellInMergeRange(string cellRef, string? rangeRef)
    {
        if (string.IsNullOrEmpty(rangeRef) || !rangeRef.Contains(':')) return false;
        var parts = rangeRef.Split(':');
        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);
        var (cellCol, cellRow) = ParseCellReference(cellRef);

        var cellColIdx = ColumnNameToIndex(cellCol);
        return cellRow >= startRow && cellRow <= endRow
            && cellColIdx >= ColumnNameToIndex(startCol) && cellColIdx <= ColumnNameToIndex(endCol);
    }

    private DocumentNode GetCellRange(string sheetName, SheetData sheetData, string range, int depth)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid range: {range}");

        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{range}",
            Type = "range",
            Preview = range
        };

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = row.RowIndex?.Value ?? 0;
            if (rowIdx < startRow || rowIdx > endRow) continue;

            foreach (var cell in row.Elements<Cell>())
            {
                var (colName, _) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx < ColumnNameToIndex(startCol) || colIdx > ColumnNameToIndex(endCol)) continue;

                node.Children.Add(CellToNode(sheetName, cell));
            }
        }

        node.ChildCount = node.Children.Count;
        return node;
    }

    private static Cell? FindCell(SheetData sheetData, string cellRef)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                    return cell;
            }
        }
        return null;
    }

    private static Cell FindOrCreateCell(SheetData sheetData, string cellRef)
    {
        var (colName, rowIdx) = ParseCellReference(cellRef);

        // Find or create row
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIdx };
            // Insert in order
            var after = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < rowIdx);
            if (after != null)
                after.InsertAfterSelf(row);
            else
                sheetData.InsertAt(row, 0);
        }

        // Find or create cell
        var cell = row.Elements<Cell>().FirstOrDefault(c =>
            c.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef.ToUpperInvariant() };
            // Insert in column order
            var afterCell = row.Elements<Cell>().LastOrDefault(c =>
            {
                var (cn, _) = ParseCellReference(c.CellReference?.Value ?? "A1");
                return ColumnNameToIndex(cn) < ColumnNameToIndex(colName);
            });
            if (afterCell != null)
                afterCell.InsertAfterSelf(cell);
            else
                row.InsertAt(cell, 0);
        }

        return cell;
    }

    // ==================== Conditional Formatting Helpers ====================

    private static bool IsTruthy(string value) =>
        ParseHelpers.IsTruthy(value);

    private static IconSetValues ParseIconSetValues(string name)
    {
        return name.ToLowerInvariant() switch
        {
            "3arrows" => IconSetValues.ThreeArrows,
            "3arrowsgray" => IconSetValues.ThreeArrowsGray,
            "3flags" => IconSetValues.ThreeFlags,
            "3trafficlights1" => IconSetValues.ThreeTrafficLights1,
            "3trafficlights2" => IconSetValues.ThreeTrafficLights2,
            "3signs" => IconSetValues.ThreeSigns,
            "3symbols" => IconSetValues.ThreeSymbols,
            "3symbols2" => IconSetValues.ThreeSymbols2,
            "4arrows" => IconSetValues.FourArrows,
            "4arrowsgray" => IconSetValues.FourArrowsGray,
            "4rating" => IconSetValues.FourRating,
            "4redtoblack" => IconSetValues.FourRedToBlack,
            "4trafficlights" => IconSetValues.FourTrafficLights,
            "5arrows" => IconSetValues.FiveArrows,
            "5arrowsgray" => IconSetValues.FiveArrowsGray,
            "5rating" => IconSetValues.FiveRating,
            "5quarters" => IconSetValues.FiveQuarters,
            _ => throw new ArgumentException($"Unknown icon set name: '{name}'. Valid names: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 3Symbols2, 4Arrows, 4ArrowsGray, 4Rating, 4RedToBlack, 4TrafficLights, 5Arrows, 5ArrowsGray, 5Rating, 5Quarters")
        };
    }

    private static int GetIconCount(string name)
    {
        var lower = name.ToLowerInvariant();
        if (lower.StartsWith("5")) return 5;
        if (lower.StartsWith("4")) return 4;
        return 3;
    }

    // ==================== Data Validation Helpers ====================

    private DocumentNode TableToNode(string sheetName, WorksheetPart worksheetPart, int tableIndex, int depth)
    {
        var tableParts = worksheetPart.TableDefinitionParts.ToList();
        if (tableIndex < 1 || tableIndex > tableParts.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range (1..{tableParts.Count})");

        var tbl = tableParts[tableIndex - 1].Table
            ?? throw new ArgumentException($"Table {tableIndex} has no definition");

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/table[{tableIndex}]",
            Type = "table",
            Text = tbl.DisplayName?.Value ?? tbl.Name?.Value ?? $"Table{tableIndex}",
            Preview = $"{tbl.Name?.Value} ({tbl.Reference?.Value})"
        };

        node.Format["name"] = tbl.Name?.Value ?? "";
        node.Format["displayName"] = tbl.DisplayName?.Value ?? "";
        node.Format["ref"] = tbl.Reference?.Value ?? "";

        var styleInfo = tbl.GetFirstChild<TableStyleInfo>();
        if (styleInfo?.Name?.Value != null)
            node.Format["style"] = styleInfo.Name.Value;

        node.Format["headerRow"] = (tbl.HeaderRowCount?.Value ?? 1) != 0;
        node.Format["totalRow"] = tbl.TotalsRowShown?.Value ?? false;

        var tableColumns = tbl.GetFirstChild<TableColumns>();
        if (tableColumns != null)
        {
            var colNames = tableColumns.Elements<TableColumn>()
                .Select(c => c.Name?.Value ?? "").ToArray();
            node.Format["columns"] = string.Join(",", colNames);
            node.ChildCount = colNames.Length;
        }

        return node;
    }

    private DocumentNode CommentToNode(string sheetName, Comment comment, Comments comments, int index)
    {
        var reference = comment.Reference?.Value ?? "?";
        var text = comment.CommentText?.InnerText ?? "";
        var authorId = comment.AuthorId?.Value ?? 0;

        var authors = comments.GetFirstChild<Authors>();
        var authorName = authors?.Elements<Author>().ElementAtOrDefault((int)authorId)?.Text ?? "Unknown";

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/comment[{index}]",
            Type = "comment",
            Text = text,
            Preview = $"{reference}: {text}"
        };

        node.Format["ref"] = reference;
        node.Format["author"] = authorName;

        return node;
    }

    private static DocumentNode DataValidationToNode(string sheetName, DataValidation dv, int index)
    {
        var sqref = dv.SequenceOfReferences?.InnerText ?? "";
        var node = new DocumentNode
        {
            Path = $"/{sheetName}/validation[{index}]",
            Type = "validation",
            Text = sqref,
            Preview = $"validation[{index}] ({sqref})"
        };

        node.Format["sqref"] = sqref;

        if (dv.Type?.HasValue == true)
            node.Format["type"] = dv.Type.InnerText;
        if (dv.Operator?.HasValue == true)
            node.Format["operator"] = dv.Operator.InnerText;

        if (dv.Formula1 != null)
        {
            var f1 = dv.Formula1.Text ?? "";
            if (f1.StartsWith("\"") && f1.EndsWith("\""))
                f1 = f1[1..^1];
            node.Format["formula1"] = f1;
        }

        if (dv.Formula2 != null)
            node.Format["formula2"] = dv.Formula2.Text ?? "";

        if (dv.AllowBlank?.HasValue == true)
            node.Format["allowBlank"] = dv.AllowBlank.Value;
        if (dv.ShowErrorMessage?.HasValue == true)
            node.Format["showError"] = dv.ShowErrorMessage.Value;
        if (dv.ShowInputMessage?.HasValue == true)
            node.Format["showInput"] = dv.ShowInputMessage.Value;

        if (!string.IsNullOrEmpty(dv.ErrorTitle?.Value))
            node.Format["errorTitle"] = dv.ErrorTitle!.Value!;
        if (!string.IsNullOrEmpty(dv.Error?.Value))
            node.Format["error"] = dv.Error!.Value!;
        if (!string.IsNullOrEmpty(dv.PromptTitle?.Value))
            node.Format["promptTitle"] = dv.PromptTitle!.Value!;
        if (!string.IsNullOrEmpty(dv.Prompt?.Value))
            node.Format["prompt"] = dv.Prompt!.Value!;

        return node;
    }

    // ==================== Picture Helpers ====================

    private DocumentNode GetPictureNode(string sheetName, WorksheetPart worksheetPart, int index, string path)
    {
        var drawingsPart = worksheetPart.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/pictures");

        var wsDrawing = drawingsPart.WorksheetDrawing
            ?? throw new ArgumentException("Sheet has no drawings/pictures");

        var picAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Picture>().Any())
            .ToList();

        if (index < 1 || index > picAnchors.Count)
            throw new ArgumentException($"Picture index {index} out of range (1..{picAnchors.Count})");

        var anchor = picAnchors[index - 1];
        var picture = anchor.Descendants<XDR.Picture>().First();

        var node = new DocumentNode { Path = path, Type = "picture" };

        var nvProps = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
        if (nvProps != null)
        {
            if (!string.IsNullOrEmpty(nvProps.Description?.Value))
            {
                node.Format["alt"] = nvProps.Description.Value;
                node.Text = nvProps.Description.Value;
            }
            if (!string.IsNullOrEmpty(nvProps.Name?.Value))
                node.Format["name"] = nvProps.Name.Value;
        }

        var from = anchor.FromMarker;
        var to = anchor.ToMarker;
        if (from != null)
        {
            node.Format["x"] = from.ColumnId?.Text ?? "0";
            node.Format["y"] = from.RowId?.Text ?? "0";
        }
        if (to != null && from != null)
        {
            var fromCol = int.TryParse(from.ColumnId?.Text, out var fc) ? fc : 0;
            var toCol = int.TryParse(to.ColumnId?.Text, out var tc) ? tc : 0;
            var fromRow = int.TryParse(from.RowId?.Text, out var fr) ? fr : 0;
            var toRow = int.TryParse(to.RowId?.Text, out var tr2) ? tr2 : 0;
            node.Format["width"] = (toCol - fromCol).ToString();
            node.Format["height"] = (toRow - fromRow).ToString();
        }

        return node;
    }
}
