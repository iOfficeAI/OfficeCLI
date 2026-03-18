// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
        {
            var node = new DocumentNode { Path = "/", Type = "workbook" };
            foreach (var (name, part) in GetWorksheets())
            {
                var sheetNode = new DocumentNode { Path = $"/{name}", Type = "sheet", Preview = name };
                var sheetData = GetSheet(part).GetFirstChild<SheetData>();
                sheetNode.ChildCount = sheetData?.Elements<Row>().Count() ?? 0;

                if (depth > 0 && sheetData != null)
                {
                    sheetNode.Children = GetSheetChildNodes(name, sheetData, depth);
                }

                node.Children.Add(sheetNode);
            }
            node.ChildCount = node.Children.Count;
            return node;
        }

        // Handle /namedrange[N] or /namedrange[Name]
        var namedRangeMatch = Regex.Match(path.TrimStart('/'), @"^namedrange\[(.+?)\]$", RegexOptions.IgnoreCase);
        if (namedRangeMatch.Success)
        {
            var selector = namedRangeMatch.Groups[1].Value;
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames == null)
                throw new ArgumentException("No named ranges found in workbook");

            var allDefs = definedNames.Elements<DefinedName>().ToList();
            DefinedName? dn = null;
            int dnIndex;

            if (int.TryParse(selector, out dnIndex))
            {
                if (dnIndex < 1 || dnIndex > allDefs.Count)
                    throw new ArgumentException($"Named range index {dnIndex} out of range (1-{allDefs.Count})");
                dn = allDefs[dnIndex - 1];
            }
            else
            {
                dn = allDefs.FirstOrDefault(d =>
                    d.Name?.Value?.Equals(selector, StringComparison.OrdinalIgnoreCase) == true);
                if (dn == null)
                    throw new ArgumentException($"Named range '{selector}' not found");
                dnIndex = allDefs.IndexOf(dn) + 1;
            }

            var nrNode = new DocumentNode
            {
                Path = $"/namedrange[{dnIndex}]",
                Type = "namedrange",
                Text = dn.Name?.Value ?? "",
                Preview = dn.InnerText
            };
            nrNode.Format["name"] = dn.Name?.Value ?? "";
            nrNode.Format["ref"] = dn.InnerText ?? "";
            if (dn.LocalSheetId?.HasValue == true)
            {
                var sheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                if (sheets != null && (int)dn.LocalSheetId.Value < sheets.Count)
                    nrNode.Format["scope"] = sheets[(int)dn.LocalSheetId.Value].Name?.Value ?? "";
            }
            if (!string.IsNullOrEmpty(dn.Comment?.Value))
                nrNode.Format["comment"] = dn.Comment.Value;

            return nrNode;
        }

        // Parse path: /SheetName or /SheetName/A1 or /SheetName/A1:D10
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetNameFromPath = segments[0];
        var worksheet = FindWorksheet(sheetNameFromPath);
        if (worksheet == null)
            throw new ArgumentException($"Sheet not found: {sheetNameFromPath}");

        var data = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (data == null)
            return new DocumentNode { Path = path, Type = "sheet", Preview = "(empty)" };

        if (segments.Length == 1)
        {
            // Return sheet overview
            var sheetNode = new DocumentNode
            {
                Path = path,
                Type = "sheet",
                Preview = sheetNameFromPath,
                ChildCount = data.Elements<Row>().Count()
            };

            // Include freeze pane info
            var ws = GetSheet(worksheet);
            var pane = ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Pane>();
            if (pane != null && pane.State?.Value == PaneStateValues.Frozen)
            {
                sheetNode.Format["freeze"] = pane.TopLeftCell?.Value ?? "";
            }

            // Include autofilter info
            var autoFilter = ws.GetFirstChild<AutoFilter>();
            if (autoFilter?.Reference?.Value != null)
            {
                sheetNode.Format["autofilter"] = autoFilter.Reference.Value;
            }

            if (depth > 0)
            {
                sheetNode.Children = GetSheetChildNodes(sheetNameFromPath, data, depth);
            }
            return sheetNode;
        }

        var cellRef = segments[1];

        // Validation path: /Sheet1/validation[N]
        var validationMatch = Regex.Match(cellRef, @"^validation\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationMatch.Success)
        {
            var dvIdx = int.Parse(validationMatch.Groups[1].Value);
            var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>();
            if (dvs == null)
                throw new ArgumentException("No data validations found in sheet");

            var dvList = dvs.Elements<DataValidation>().ToList();
            if (dvIdx < 1 || dvIdx > dvList.Count)
                throw new ArgumentException($"Validation index {dvIdx} out of range (1-{dvList.Count})");

            return DataValidationToNode(sheetNameFromPath, dvList[dvIdx - 1], dvIdx);
        }

        // Column path: /Sheet1/col[A]
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z0-9]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colValue = colMatch.Groups[1].Value;
            var colName = int.TryParse(colValue, out var numIdx) ? IndexToColumnName(numIdx) : colValue.ToUpperInvariant();
            var colIdx = (uint)ColumnNameToIndex(colName);
            var colNode = new DocumentNode { Path = path, Type = "column", Preview = colName };
            var columns = GetSheet(worksheet).GetFirstChild<Columns>();
            if (columns != null)
            {
                var col = columns.Elements<Column>().FirstOrDefault(c =>
                    c.Min?.Value <= colIdx && c.Max?.Value >= colIdx);
                if (col != null)
                {
                    if (col.Width?.Value != null) colNode.Format["width"] = col.Width.Value;
                    if (col.Hidden?.Value == true) colNode.Format["hidden"] = true;
                    if (col.CustomWidth?.Value == true) colNode.Format["customWidth"] = true;
                }
            }
            return colNode;
        }

        // Row path: /Sheet1/row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = data.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
            if (row == null)
                return new DocumentNode { Path = path, Type = "row", Preview = $"row {rowIdx}", Text = "(empty)" };
            var rowNode = new DocumentNode
            {
                Path = path, Type = "row", ChildCount = row.Elements<Cell>().Count()
            };
            if (row.Height?.Value != null) rowNode.Format["height"] = row.Height.Value;
            if (row.Hidden?.Value == true) rowNode.Format["hidden"] = true;
            if (depth > 0)
                foreach (var c in row.Elements<Cell>())
                    rowNode.Children.Add(CellToNode(sheetNameFromPath, c, worksheet));
            return rowNode;
        }

        // Conditional formatting path: /Sheet1/cf[N]
        var cfMatch = Regex.Match(cellRef, @"^cf\[(\d+)\]$");
        if (cfMatch.Success)
        {
            var cfIdx = int.Parse(cfMatch.Groups[1].Value);
            var cfElements = GetSheet(worksheet).Elements<ConditionalFormatting>().ToList();
            if (cfIdx < 1 || cfIdx > cfElements.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"CF {cfIdx} not found (total: {cfElements.Count})" };

            var cf = cfElements[cfIdx - 1];
            var cfNode = new DocumentNode { Path = path, Type = "conditionalFormatting" };
            cfNode.Format["sqref"] = cf.SequenceOfReferences?.InnerText ?? "";

            var rule = cf.Elements<ConditionalFormattingRule>().FirstOrDefault();
            if (rule != null)
            {
                if (rule.Type?.Value != null)
                    cfNode.Format["ruleType"] = rule.Type.InnerText;

                // DataBar
                var dataBar = rule.GetFirstChild<DataBar>();
                if (dataBar != null)
                {
                    cfNode.Format["cfType"] = "dataBar";
                    var dbColor = dataBar.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                    if (dbColor?.Rgb?.Value != null) cfNode.Format["color"] = dbColor.Rgb.Value;
                }

                // ColorScale
                var colorScale = rule.GetFirstChild<ColorScale>();
                if (colorScale != null)
                {
                    cfNode.Format["cfType"] = "colorScale";
                    var colors = colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                    if (colors.Count >= 2)
                    {
                        cfNode.Format["mincolor"] = colors[0].Rgb?.Value ?? "";
                        cfNode.Format["maxcolor"] = colors[^1].Rgb?.Value ?? "";
                        if (colors.Count >= 3)
                            cfNode.Format["midcolor"] = colors[1].Rgb?.Value ?? "";
                    }
                }

                // IconSet
                var iconSet = rule.GetFirstChild<IconSet>();
                if (iconSet != null)
                {
                    cfNode.Format["cfType"] = "iconSet";
                    if (iconSet.IconSetValue?.Value != null)
                        cfNode.Format["iconset"] = iconSet.IconSetValue.InnerText;
                    if (iconSet.ShowValue?.Value != null)
                        cfNode.Format["showvalue"] = iconSet.ShowValue.Value;
                    if (iconSet.Reverse?.Value == true)
                        cfNode.Format["reverse"] = true;
                }

                // Formula-based
                var formula = rule.GetFirstChild<Formula>();
                if (formula != null && rule.Type?.Value == ConditionalFormatValues.Expression)
                {
                    cfNode.Format["cfType"] = "formula";
                    cfNode.Format["formula"] = formula.Text ?? "";
                    if (rule.FormatId?.Value != null)
                        cfNode.Format["dxfId"] = rule.FormatId.Value;
                }
            }
            return cfNode;
        }

        // AutoFilter path: /Sheet1/autofilter
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            var af = GetSheet(worksheet).GetFirstChild<AutoFilter>();
            var afNode = new DocumentNode { Path = path, Type = "autofilter" };
            if (af?.Reference?.Value != null) afNode.Format["range"] = af.Reference.Value;
            return afNode;
        }

        // Chart path: /Sheet1/chart[N]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartIdx = int.Parse(chartMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart;
            if (drawingsPart == null)
                return new DocumentNode { Path = path, Type = "error", Text = "No charts in this sheet" };

            var chartParts = drawingsPart.ChartParts.ToList();
            if (chartIdx < 1 || chartIdx > chartParts.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"Chart {chartIdx} not found" };

            var chartPart = chartParts[chartIdx - 1];
            var chart = chartPart.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            var chartNode = new DocumentNode { Path = path, Type = "chart" };
            if (chart != null)
                ChartHelper.ReadChartProperties(chart, chartNode, depth);
            return chartNode;
        }

        // Comment path: /Sheet1/comment[N]
        var commentMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentMatch.Success)
        {
            var cmtIndex = int.Parse(commentMatch.Groups[1].Value);
            var commentsPart = worksheet.WorksheetCommentsPart;
            if (commentsPart?.Comments == null)
                throw new ArgumentException($"No comments found in sheet: {sheetNameFromPath}");

            var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
            var cmtElement = cmtList?.Elements<Comment>().ElementAtOrDefault(cmtIndex - 1)
                ?? throw new ArgumentException($"Comment [{cmtIndex}] not found");

            return CommentToNode(sheetNameFromPath, cmtElement, commentsPart.Comments, cmtIndex);
        }

        // Table path: /Sheet1/table[N]
        var tableMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableMatch.Success)
        {
            var tableIdx = int.Parse(tableMatch.Groups[1].Value);
            return TableToNode(sheetNameFromPath, worksheet, tableIdx, depth);
        }

        // Cell reference: A1 or range A1:D10
        // Check if it's a cell reference or a generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = System.Text.RegularExpressions.Regex.IsMatch(firstPart, @"^[A-Z]+\d+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        if (!isCellRef)
        {
            // Handle picture[N] path segment
            var picMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (picMatch.Success)
            {
                var picIndex = int.Parse(picMatch.Groups[1].Value);
                return GetPictureNode(sheetNameFromPath, worksheet, picIndex, path);
            }

            // Generic XML fallback: navigate worksheet XML tree
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {cellRef}" };
            return GenericXmlQuery.ElementToNode(target, path, depth);
        }

        if (cellRef.Contains(':'))
        {
            // Range
            return GetCellRange(sheetNameFromPath, data, cellRef, depth);
        }
        else
        {
            // Single cell
            var cell = FindCell(data, cellRef);
            if (cell == null)
                return new DocumentNode { Path = path, Type = "cell", Text = "(empty)", Preview = cellRef };
            return CellToNode(sheetNameFromPath, cell, worksheet);
        }
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();

        // Check if element type is known (Scheme A) or should fall back to generic XML (Scheme B)
        var elementMatch = Regex.Match(selector.Split('!').Last(), @"^(\w+)");
        var elementName = elementMatch.Success ? elementMatch.Groups[1].Value : "";
        bool isKnownType = string.IsNullOrEmpty(elementName)
            || elementName is "cell" or "row" or "sheet" or "validation" or "comment" or "note" or "table" or "listobject" or "chart"
            || (elementName.Length <= 3 && Regex.IsMatch(elementName, @"^[A-Z]+$", RegexOptions.IgnoreCase));
        if (!isKnownType)
        {
            // Scheme B: generic XML fallback
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var (_, worksheetPart) in GetWorksheets())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSheet(worksheetPart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        var parsed = ParseCellSelector(selector);

        // Handle validation queries
        if (elementName == "validation")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var dvs = GetSheet(worksheetPart).GetFirstChild<DataValidations>();
                if (dvs == null) continue;

                var dvList = dvs.Elements<DataValidation>().ToList();
                for (int i = 0; i < dvList.Count; i++)
                    results.Add(DataValidationToNode(sheetName, dvList[i], i + 1));
            }
            return results;
        }

        // Handle comment queries
        if (elementName is "comment" or "note")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var commentsPart = worksheetPart.WorksheetCommentsPart;
                if (commentsPart?.Comments == null) continue;

                var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
                if (cmtList == null) continue;

                var cmtElements = cmtList.Elements<Comment>().ToList();
                for (int i = 0; i < cmtElements.Count; i++)
                    results.Add(CommentToNode(sheetName, cmtElements[i], commentsPart.Comments, i + 1));
            }
            return results;
        }

        // Handle table queries
        if (elementName is "table" or "listobject")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var tableParts = worksheetPart.TableDefinitionParts.ToList();
                for (int i = 0; i < tableParts.Count; i++)
                    results.Add(TableToNode(sheetName, worksheetPart, i + 1, 0));
            }
            return results;
        }

        // Handle chart queries
        if (elementName == "chart")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null) continue;

                var chartParts = drawingsPart.ChartParts.ToList();
                for (int i = 0; i < chartParts.Count; i++)
                {
                    var chart = chartParts[i].ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                    var node = new DocumentNode { Path = $"/{sheetName}/chart[{i + 1}]", Type = "chart" };
                    if (chart != null)
                        ChartHelper.ReadChartProperties(chart, node, 0);

                    // Filter by contains text (match on title)
                    if (parsed.ValueContains != null)
                    {
                        var title = node.Format.TryGetValue("title", out var t) ? t?.ToString() : null;
                        if (title == null || !title.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        foreach (var (sheetName, worksheetPart) in GetWorksheets())
        {
            // If selector specifies a sheet, skip non-matching sheets
            if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                continue;

            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (MatchesCellSelector(cell, sheetName, parsed))
                    {
                        results.Add(CellToNode(sheetName, cell, worksheetPart));
                    }
                }
            }
        }

        return results;
    }
}
