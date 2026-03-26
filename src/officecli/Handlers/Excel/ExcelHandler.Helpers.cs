// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Path Normalization ====================

    /// <summary>
    /// Normalize Excel-native path notation to DOM style.
    /// Sheet1!A1 → /Sheet1/A1
    /// Sheet1!A1:D10 → /Sheet1/A1:D10
    /// Sheet1!row[2] → /Sheet1/row[2]
    /// Sheet1!1:1 → /Sheet1/row[1]   (whole row)
    /// Sheet1!A:A → /Sheet1/col[A]   (whole column)
    /// Paths already starting with '/' are returned unchanged.
    /// </summary>
    internal static string NormalizeExcelPath(string path)
    {
        // Handle "/Sheet1!A1" — strip leading '/' when '!' is present so native notation is parsed correctly
        if (path.StartsWith('/') && path.Contains('!'))
            path = path[1..];
        if (path.StartsWith('/')) return path;
        var bang = path.IndexOf('!');
        if (bang > 0)
        {
            var sheet = path[..bang];
            var selector = path[(bang + 1)..];

            // Whole-row notation: "1:1" or "3:3"
            var wholeRow = System.Text.RegularExpressions.Regex.Match(selector, @"^(\d+):\1$");
            if (wholeRow.Success)
                return $"/{sheet}/row[{wholeRow.Groups[1].Value}]";

            // Whole-column notation: "A:A" or "AB:AB"
            var wholeCol = System.Text.RegularExpressions.Regex.Match(selector, @"^([A-Za-z]+):\1$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (wholeCol.Success)
                return $"/{sheet}/col[{wholeCol.Groups[1].Value.ToUpperInvariant()}]";

            return $"/{sheet}/{selector}";
        }
        return path;
    }

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
    /// Get a sparkline group by 1-based index from a worksheet's extension list.
    /// Returns null if not found.
    /// </summary>
    internal X14.SparklineGroup? GetSparklineGroup(WorksheetPart worksheet, int index)
    {
        var ws = GetSheet(worksheet);
        var extList = ws.GetFirstChild<WorksheetExtensionList>();
        if (extList == null) return null;

        var spkExt = extList.Elements<WorksheetExtension>()
            .FirstOrDefault(e => e.Uri == "{05C60535-1F16-4fd2-B633-E4A46CF9E463}");
        if (spkExt == null) return null;

        var spkGroups = spkExt.GetFirstChild<X14.SparklineGroups>();
        if (spkGroups == null) return null;

        var groups = spkGroups.Elements<X14.SparklineGroup>().ToList();
        if (index < 1 || index > groups.Count) return null;
        return groups[index - 1];
    }

    /// <summary>
    /// Build a DocumentNode for a sparkline group.
    /// </summary>
    internal static DocumentNode SparklineGroupToNode(string sheetName, X14.SparklineGroup spkGroup, int index)
    {
        var node = new DocumentNode
        {
            Path = $"/{sheetName}/sparkline[{index}]",
            Type = "sparkline"
        };

        // Type: default is line when attribute is absent
        string spkType;
        if (spkGroup.Type?.HasValue == true)
        {
            var tv = spkGroup.Type.Value;
            spkType = tv == X14.SparklineTypeValues.Column ? "column"
                : tv == X14.SparklineTypeValues.Stacked ? "stacked"
                : "line";
        }
        else
        {
            spkType = "line";
        }
        node.Format["type"] = spkType;

        // Color
        var colorRgb = spkGroup.SeriesColor?.Rgb?.Value;
        node.Format["color"] = colorRgb != null
            ? ParseHelpers.FormatHexColor(colorRgb)
            : "#4472C4";

        // Negative color
        var negColorRgb = spkGroup.NegativeColor?.Rgb?.Value;
        if (negColorRgb != null)
            node.Format["negativeColor"] = ParseHelpers.FormatHexColor(negColorRgb);

        // Boolean flags
        if (spkGroup.Markers?.Value == true) node.Format["markers"] = true;
        if (spkGroup.High?.Value == true) node.Format["highPoint"] = true;
        if (spkGroup.Low?.Value == true) node.Format["lowPoint"] = true;
        if (spkGroup.First?.Value == true) node.Format["firstPoint"] = true;
        if (spkGroup.Last?.Value == true) node.Format["lastPoint"] = true;
        if (spkGroup.Negative?.Value == true) node.Format["negative"] = true;

        // Line weight
        if (spkGroup.LineWeight?.HasValue == true)
            node.Format["lineWeight"] = spkGroup.LineWeight.Value;

        // Cell / range from first sparkline element
        var firstSparkline = spkGroup.GetFirstChild<X14.Sparklines>()?.GetFirstChild<X14.Sparkline>();
        if (firstSparkline != null)
        {
            var cell = firstSparkline.ReferenceSequence?.Text ?? "";
            node.Format["cell"] = cell;

            // Strip sheet prefix from range (Sheet1!A1:E1 → A1:E1)
            var formulaText = firstSparkline.Formula?.Text ?? "";
            var excl = formulaText.IndexOf('!');
            node.Format["range"] = excl >= 0 ? formulaText[(excl + 1)..] : formulaText;
        }

        return node;
    }

    /// <summary>
    /// Delete the calculation chain part if present.
    /// Excel will recalculate and recreate it on next open.
    /// This avoids stale calc chain references after cell/formula mutations.
    /// </summary>
    private void DeleteCalcChainIfPresent()
    {
        var calcChainPart = _doc.WorkbookPart?.CalculationChainPart;
        if (calcChainPart != null)
            _doc.WorkbookPart!.DeletePart(calcChainPart);
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

    private ArgumentException SheetNotFoundException(string sheetName)
    {
        var available = GetWorksheets().Select(w => w.Name).ToList();
        var availableStr = available.Count > 0
            ? string.Join(", ", available)
            : "(none)";
        return new ArgumentException(
            $"Sheet not found: \"{sheetName}\". Available sheets: [{availableStr}]. " +
            $"Use DOM path \"/{available.FirstOrDefault() ?? "SheetName"}/A1\" or Excel notation \"{available.FirstOrDefault() ?? "SheetName"}!A1\".");
    }

    private string GetCellDisplayValue(Cell cell)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? "";
        }

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

        // Formula cells: Text is always the cached value (or "" if none).
        // The formula itself is exposed via Format["formula"] in CellToNode.
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
                    rowNode.Children.Add(CellToNode(sheetName, cell, worksheetPart));
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

        // When a formula cell has no cached value, display the formula as text
        var displayText = value;
        if (string.IsNullOrEmpty(displayText) && formula != null)
            displayText = "=" + formula;

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{cellRef}",
            Type = "cell",
            Text = displayText,
            Preview = cellRef
        };

        node.Format["type"] = type;
        if (formula != null)
        {
            node.Format["formula"] = formula;
            // Expose cached value separately so callers know whether the formula has been evaluated
            var rawCached = cell.CellValue?.Text;
            if (!string.IsNullOrEmpty(rawCached))
                node.Format["cachedValue"] = rawCached;
            else
                node.Format["uncalculated"] = true;
        }
        // Array formula readback — keys match Set input
        if (cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array)
        {
            node.Format["arrayformula"] = true;
            if (cell.CellFormula.Reference?.Value != null)
                node.Format["arrayref"] = cell.CellFormula.Reference.Value;
        }
        if (string.IsNullOrEmpty(value) && formula == null) node.Format["empty"] = true;

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
                            if (font.Bold != null) { node.Format["bold"] = true; }
                            if (font.Italic != null)
                            {
                                node.Format["italic"] = true;
                            }
                            if (font.Strike != null) node.Format["font.strike"] = true;
                            if (font.Underline != null)
                                node.Format["font.underline"] = font.Underline.Val?.InnerText == "double" ? "double" : "single";
                            if (font.Color?.Rgb?.Value != null)
                                node.Format["font.color"] = ParseHelpers.FormatHexColor(font.Color.Rgb.Value);
                            else if (font.Color?.Theme?.Value != null)
                            {
                                var themeName = ParseHelpers.ExcelThemeIndexToName(font.Color.Theme.Value);
                                if (themeName != null) node.Format["font.color"] = themeName;
                            }
                            // vertAlign (superscript/subscript) readback — dual keys like bold/italic
                            var vertAlign = font.GetFirstChild<VerticalTextAlignment>();
                            if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Superscript)
                            {
                                node.Format["superscript"] = true;
                            }
                            else if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Subscript)
                            {
                                node.Format["subscript"] = true;
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
                                    var colorParts = string.Join(";", stops.Select(s => ParseHelpers.FormatHexColor(s.Color?.Rgb?.Value ?? "")));
                                    int deg = (int)(gf.Degree?.Value ?? 0);
                                    node.Format["fill"] = $"gradient;{colorParts};{deg}";
                                }
                            }
                            else
                            {
                                var pf = fill.PatternFill;
                                if (pf?.ForegroundColor?.Rgb?.Value != null)
                                    node.Format["fill"] = ParseHelpers.FormatHexColor(pf.ForegroundColor.Rgb.Value);
                                else if (pf?.ForegroundColor?.Theme?.Value != null)
                                {
                                    var themeName = ParseHelpers.ExcelThemeIndexToName(pf.ForegroundColor.Theme.Value);
                                    if (themeName != null) node.Format["fill"] = themeName;
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
                                        node.Format[$"border.{side}.color"] = ParseHelpers.FormatHexColor(b.Color.Rgb.Value!);
                                }
                            }
                            // Diagonal border readback
                            var diag = border.DiagonalBorder;
                            if (diag?.Style?.Value != null && diag.Style.Value != BorderStyleValues.None)
                            {
                                node.Format["border.diagonal"] = diag.Style.InnerText;
                                if (diag.Color?.Rgb?.Value != null)
                                    node.Format["border.diagonal.color"] = ParseHelpers.FormatHexColor(diag.Color.Rgb.Value!);
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
                        if (alignment.TextRotation?.HasValue == true && alignment.TextRotation.Value != 0)
                            node.Format["alignment.textRotation"] = alignment.TextRotation.Value.ToString();
                        if (alignment.Indent?.HasValue == true && alignment.Indent.Value > 0)
                            node.Format["alignment.indent"] = alignment.Indent.Value.ToString();
                        if (alignment.ShrinkToFit?.Value == true)
                            node.Format["alignment.shrinkToFit"] = true;
                    }

                    // Number format readback
                    var numFmtId = xf.NumberFormatId?.Value ?? 0;
                    if (numFmtId > 0)
                    {
                        node.Format["numFmtId"] = (int)numFmtId;
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
                                _ => (object)(int)numFmtId // fallback to ID for truly unknown formats
                            };
                        }
                        node.Format["numberformat"] = fmtVal;
                    }

                    // Protection readback — always output locked state when protection is set
                    var prot = xf.Protection;
                    if (xf.ApplyProtection?.Value == true && prot != null)
                    {
                        // Always output locked state so agent can see it
                        node.Format["locked"] = prot.Locked?.Value ?? true;
                        if (prot.Hidden?.Value == true)
                            node.Format["formulahidden"] = true;
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
                    var mergeRef = mergeCell.Reference?.Value ?? "";
                    node.Format["merge"] = mergeRef;
                    // Indicate if this cell is the top-left anchor of the merged range
                    if (mergeRef.Split(':')[0].Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                        node.Format["mergeAnchor"] = true;
                }
            }
        }

        // Rich text (SST runs) readback
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(cell.CellValue?.Text, out var sstIdx2))
        {
            var sst2 = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi2 = sst2?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx2);
            if (ssi2 != null)
            {
                var runs = ssi2.Elements<Run>().ToList();
                if (runs.Count > 0)
                {
                    node.Format["richtext"] = true;
                    node.ChildCount = runs.Count;
                    int runI = 1;
                    foreach (var run in runs)
                    {
                        node.Children.Add(RunToNode(run, $"/{sheetName}/{cellRef}/run[{runI}]"));
                        runI++;
                    }
                }
            }
        }

        return node;
    }

    private static DocumentNode RunToNode(Run run, string path)
    {
        var runNode = new DocumentNode { Path = path, Type = "run", Text = run.Text?.Text ?? "" };
        var rp = run.RunProperties;
        if (rp != null)
        {
            if (rp.GetFirstChild<Bold>() != null) runNode.Format["bold"] = true;
            if (rp.GetFirstChild<Italic>() != null) runNode.Format["italic"] = true;
            if (rp.GetFirstChild<Strike>() != null) runNode.Format["strike"] = true;
            var ul = rp.GetFirstChild<Underline>();
            if (ul != null) runNode.Format["underline"] = ul.Val?.InnerText == "double" ? "double" : "single";
            var va = rp.GetFirstChild<VerticalTextAlignment>();
            if (va?.Val?.Value == VerticalAlignmentRunValues.Superscript) runNode.Format["superscript"] = true;
            if (va?.Val?.Value == VerticalAlignmentRunValues.Subscript) runNode.Format["subscript"] = true;
            if (rp.GetFirstChild<FontSize>()?.Val?.Value != null)
                runNode.Format["size"] = $"{rp.GetFirstChild<FontSize>()!.Val!.Value:0.##}pt";
            if (rp.GetFirstChild<Color>()?.Rgb?.Value != null)
                runNode.Format["color"] = ParseHelpers.FormatHexColor(rp.GetFirstChild<Color>()!.Rgb!.Value!);
            if (rp.GetFirstChild<RunFont>()?.Val?.Value != null)
                runNode.Format["font"] = rp.GetFirstChild<RunFont>()!.Val!.Value!;
        }
        return runNode;
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

    private DocumentNode GetCellRange(string sheetName, SheetData sheetData, string range, int depth, WorksheetPart? part = null)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid range: {range}");

        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{range}",
            Type = "range",
            Preview = range
        };

        // Build lookup of existing cells so we can fill empty stubs for missing positions
        var existingCells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value != null)
                    existingCells[cell.CellReference.Value] = cell;
            }
        }

        // Enumerate every position in the range in row-major order,
        // materializing empty stubs for positions that have no cell element.
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startColIdx; c <= endColIdx; c++)
            {
                var cellRef = $"{IndexToColumnName(c)}{r}";
                if (existingCells.TryGetValue(cellRef, out var existingCell))
                    node.Children.Add(CellToNode(sheetName, existingCell, part));
                else
                    node.Children.Add(new DocumentNode
                    {
                        Path = $"/{sheetName}/{cellRef}",
                        Type = "cell",
                        Text = "",
                        Preview = cellRef,
                        Format = { ["type"] = "Number", ["empty"] = true }
                    });
            }
        }

        node.ChildCount = node.Children.Count;
        return node;
    }

    /// <summary>
    /// Parse a cell value for sorting: returns a tuple (rank, numVal, strVal) so that
    /// nulls/empties sort last, numbers sort before strings, and cross-type comparison never occurs.
    /// rank=0 for numbers, rank=1 for strings, rank=2 for empty/null.
    /// </summary>
    private static (int Rank, double NumVal, string StrVal) ParseSortValue(string value)
    {
        if (string.IsNullOrEmpty(value)) return (2, 0.0, "");
        if (double.TryParse(value, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var num))
            return (0, num, "");
        return (1, 0.0, value);
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
        node.Format["totalRow"] = (tbl.TotalsRowCount?.Value ?? 0) > 0 || (tbl.TotalsRowShown?.Value ?? false);

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

    private DocumentNode? GetPictureNode(string sheetName, WorksheetPart worksheetPart, int index, string path)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart == null) return null;

        var wsDrawing = drawingsPart.WorksheetDrawing;
        if (wsDrawing == null) return null;

        var picAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Picture>().Any())
            .ToList();

        if (index < 1 || index > picAnchors.Count)
            return null;

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

        ReadAnchorPosition(anchor, node);

        return node;
    }

    private DocumentNode? GetShapeNode(string sheetName, WorksheetPart worksheetPart, int index, string path)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart == null) return null;
        var wsDrawing = drawingsPart.WorksheetDrawing;
        if (wsDrawing == null) return null;

        var shpAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Shape>().Any()).ToList();

        if (index < 1 || index > shpAnchors.Count)
            return null;

        var anchor = shpAnchors[index - 1];
        var shape = anchor.Descendants<XDR.Shape>().First();

        var node = new DocumentNode { Path = path, Type = "shape" };

        // Name
        var nvProps = shape.NonVisualShapeProperties?.GetFirstChild<XDR.NonVisualDrawingProperties>();
        if (nvProps?.Name?.Value != null)
            node.Format["name"] = nvProps.Name.Value;

        // Text
        var textRuns = shape.TextBody?.Descendants<Drawing.Run>().ToList();
        if (textRuns != null && textRuns.Count > 0)
            node.Text = string.Join("", textRuns.Select(r => r.Text?.Text ?? ""));

        // Position/size
        ReadAnchorPosition(anchor, node);

        // Font properties from first run
        var firstRun = textRuns?.FirstOrDefault();
        var rPr = firstRun?.RunProperties;
        if (rPr != null)
        {
            if (rPr.FontSize?.HasValue == true)
                node.Format["size"] = $"{rPr.FontSize.Value / 100.0}pt";
            if (rPr.Bold?.HasValue == true && rPr.Bold.Value)
                node.Format["bold"] = true;
            if (rPr.Italic?.HasValue == true && rPr.Italic.Value)
                node.Format["italic"] = true;

            var solidFill = rPr.GetFirstChild<Drawing.SolidFill>();
            var colorHex = solidFill?.GetFirstChild<Drawing.RgbColorModelHex>();
            if (colorHex?.Val?.Value != null)
                node.Format["color"] = ParseHelpers.FormatHexColor(colorHex.Val.Value);
            else
            {
                var schemeClr = solidFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                if (schemeClr?.HasValue == true) node.Format["color"] = schemeClr.InnerText;
            }

            var latin = rPr.GetFirstChild<Drawing.LatinFont>();
            if (latin?.Typeface?.Value != null)
                node.Format["font"] = latin.Typeface.Value;
        }

        // Fill
        var spPr = shape.ShapeProperties;
        if (spPr?.GetFirstChild<Drawing.NoFill>() != null)
            node.Format["fill"] = "none";
        else
        {
            var shapeFill = spPr?.GetFirstChild<Drawing.SolidFill>();
            var fillColor = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>();
            if (fillColor?.Val?.Value != null)
                node.Format["fill"] = ParseHelpers.FormatHexColor(fillColor.Val.Value);
            else
            {
                var schemeClr = shapeFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                if (schemeClr?.HasValue == true) node.Format["fill"] = schemeClr.InnerText;
            }
        }

        // Effects — check shape-level then text-level
        var effectList = spPr?.GetFirstChild<Drawing.EffectList>();
        var textEffectList = (effectList == null || !effectList.HasChildren)
            ? rPr?.GetFirstChild<Drawing.EffectList>()
            : null;
        var activeEffects = effectList?.HasChildren == true ? effectList : textEffectList;
        if (activeEffects != null)
        {
            var shadow = activeEffects.GetFirstChild<Drawing.OuterShadow>();
            if (shadow != null)
            {
                var sColor = ParseHelpers.FormatHexColor(shadow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000");
                node.Format["shadow"] = sColor;
            }
            var glow = activeEffects.GetFirstChild<Drawing.Glow>();
            if (glow != null)
            {
                var gColor = ParseHelpers.FormatHexColor(glow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000");
                var gRadius = glow.Radius?.HasValue == true ? $"{glow.Radius.Value / 12700.0:0.##}" : "8";
                node.Format["glow"] = $"{gColor}-{gRadius}";
            }
        }

        return node;
    }

    // ==================== Shared Anchor Helpers ====================

    /// <summary>
    /// Set position/size properties (x, y, width, height) on a TwoCellAnchor.
    /// Returns true if the key was handled, false otherwise.
    /// </summary>
    private static bool TrySetAnchorPosition(XDR.TwoCellAnchor anchor, string key, string value)
    {
        switch (key)
        {
            case "x":
                if (anchor.FromMarker != null)
                {
                    var xVal = ParseHelpers.SafeParseInt(value, "x");
                    if (xVal < 0) throw new ArgumentException($"Invalid 'x' value: '{value}'. Column index must be >= 0.");
                    anchor.FromMarker.ColumnId!.Text = xVal.ToString();
                }
                return true;
            case "y":
                if (anchor.FromMarker != null)
                {
                    var yVal = ParseHelpers.SafeParseInt(value, "y");
                    if (yVal < 0) throw new ArgumentException($"Invalid 'y' value: '{value}'. Row index must be >= 0.");
                    anchor.FromMarker.RowId!.Text = yVal.ToString();
                }
                return true;
            case "width":
                if (anchor.FromMarker != null && anchor.ToMarker != null)
                {
                    var fromCol = int.TryParse(anchor.FromMarker.ColumnId?.Text, out var fc) ? fc : 0;
                    anchor.ToMarker.ColumnId!.Text = (fromCol + ParseHelpers.SafeParseInt(value, "width")).ToString();
                }
                return true;
            case "height":
                if (anchor.FromMarker != null && anchor.ToMarker != null)
                {
                    var fromRow = int.TryParse(anchor.FromMarker.RowId?.Text, out var fr) ? fr : 0;
                    anchor.ToMarker.RowId!.Text = (fromRow + ParseHelpers.SafeParseInt(value, "height")).ToString();
                }
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Read position/size from a TwoCellAnchor into a DocumentNode's Format dictionary.
    /// </summary>
    private static void ReadAnchorPosition(XDR.TwoCellAnchor anchor, DocumentNode node)
    {
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
    }

    /// <summary>
    /// Set rotation on a ShapeProperties element.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetRotation(XDR.ShapeProperties? spPr, string key, string value)
    {
        if (key is not ("rotation" or "rot")) return false;
        if (spPr == null) return true;

        var xfrm = spPr.GetFirstChild<Drawing.Transform2D>();
        if (xfrm == null)
        {
            xfrm = new Drawing.Transform2D(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 }
            );
            spPr.InsertAt(xfrm, 0);
        }
        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var degrees))
            throw new ArgumentException($"Invalid 'rotation' value: '{value}'. Expected a number in degrees (e.g. 45, -90, 180.5).");
        xfrm.Rotation = (int)(degrees * 60000);
        return true;
    }

    /// <summary>
    /// Apply shape-level effects (shadow, glow, reflection, softedge) on a ShapeProperties element.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetShapeEffect(XDR.ShapeProperties? spPr, string key, string value)
    {
        if (key is not ("shadow" or "glow" or "reflection" or "softedge")) return false;
        if (spPr == null) return true;

        var effectList = spPr.GetFirstChild<Drawing.EffectList>();
        var normalizedVal = value.Replace(':', '-');
        if (normalizedVal == "true") normalizedVal = key == "shadow" ? "000000" : key == "glow" ? "4472C4" : "half";

        if (normalizedVal.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            normalizedVal.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (effectList != null)
            {
                switch (key)
                {
                    case "shadow": effectList.RemoveAllChildren<Drawing.OuterShadow>(); break;
                    case "glow": effectList.RemoveAllChildren<Drawing.Glow>(); break;
                    case "reflection": effectList.RemoveAllChildren<Drawing.Reflection>(); break;
                    case "softedge": effectList.RemoveAllChildren<Drawing.SoftEdge>(); break;
                }
                if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            }
        }
        else
        {
            if (effectList == null) { effectList = new Drawing.EffectList(); spPr.AppendChild(effectList); }
            switch (key)
            {
                case "shadow":
                    effectList.RemoveAllChildren<Drawing.OuterShadow>();
                    effectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                    break;
                case "glow":
                    effectList.RemoveAllChildren<Drawing.Glow>();
                    effectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                    break;
                case "reflection":
                    effectList.RemoveAllChildren<Drawing.Reflection>();
                    effectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildReflection(normalizedVal));
                    break;
                case "softedge":
                    effectList.RemoveAllChildren<Drawing.SoftEdge>();
                    effectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildSoftEdge(normalizedVal));
                    break;
            }
        }
        return true;
    }

    /// <summary>
    /// Parse x, y, width, height from properties with given defaults. Used by both picture Add and shape Add.
    /// </summary>
    private static (int x, int y, int width, int height) ParseAnchorBounds(
        Dictionary<string, string> properties, string defX, string defY, string defW, string defH)
    {
        return (
            ParseHelpers.SafeParseInt(properties.GetValueOrDefault("x", defX) ?? defX, "x"),
            ParseHelpers.SafeParseInt(properties.GetValueOrDefault("y", defY) ?? defY, "y"),
            ParseHelpers.SafeParseInt(properties.GetValueOrDefault("width", defW) ?? defW, "width"),
            ParseHelpers.SafeParseInt(properties.GetValueOrDefault("height", defH) ?? defH, "height")
        );
    }

    /// <summary>
    /// Reorder RunProperties children to match CT_RPrElt schema order:
    /// b, i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color, rFont, family, charset, scheme
    /// </summary>
    private static void ReorderRunProperties(RunProperties rpr)
    {
        if (rpr == null || !rpr.HasChildren) return;
        var children = rpr.ChildElements.ToList();
        var ordered = children.OrderBy(c => GetRunPropertyOrder(c)).ToList();
        rpr.RemoveAllChildren();
        foreach (var child in ordered) rpr.AppendChild(child);
    }

    private static int GetRunPropertyOrder(DocumentFormat.OpenXml.OpenXmlElement element) => element switch
    {
        Bold => 0,
        Italic => 1,
        Strike => 2,
        Condense => 3,
        Extend => 4,
        Outline => 5,
        Shadow => 6,
        Underline => 7,
        VerticalTextAlignment => 8,
        FontSize => 9,
        Color => 10,
        RunFont => 11,
        FontFamily => 12,
        RunPropertyCharSet => 13,
        FontScheme => 14,
        _ => 99
    };
}
