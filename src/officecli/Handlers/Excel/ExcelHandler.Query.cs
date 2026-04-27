// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Path cannot be empty.");
        path = NormalizeExcelPath(path);
        path = ResolveSheetIndexInPath(path);
        if (path == "/")
        {
            var node = new DocumentNode { Path = "/", Type = "workbook" };

            // Core document properties
            var props = _doc.PackageProperties;
            if (props.Title != null) node.Format["title"] = props.Title;
            if (props.Creator != null) node.Format["author"] = props.Creator;
            if (props.Subject != null) node.Format["subject"] = props.Subject;
            if (props.Keywords != null) node.Format["keywords"] = props.Keywords;
            if (props.Description != null) node.Format["description"] = props.Description;
            if (props.Category != null) node.Format["category"] = props.Category;
            if (props.LastModifiedBy != null) node.Format["lastModifiedBy"] = props.LastModifiedBy;
            if (props.Revision != null) node.Format["revision"] = props.Revision;
            if (props.Created != null) node.Format["created"] = props.Created.Value.ToString("o");
            if (props.Modified != null) node.Format["modified"] = props.Modified.Value.ToString("o");

            foreach (var (name, part) in GetWorksheets())
            {
                var sheetNode = new DocumentNode { Path = $"/{name}", Type = "sheet", Preview = name };
                var sheetData = GetSheet(part).GetFirstChild<SheetData>();
                // R6-5: dedupe by RowIndex so a pivot placed on its own source
                // sheet doesn't double-count row children.
                var rowCount = sheetData?.Elements<Row>()
                    .Select(r => r.RowIndex?.Value ?? 0u)
                    .Where(i => i != 0)
                    .Distinct()
                    .Count() ?? 0;
                var chartCount = part.DrawingsPart != null ? CountExcelCharts(part.DrawingsPart) : 0;
                sheetNode.ChildCount = rowCount + chartCount;

                if (depth > 0 && sheetData != null)
                {
                    sheetNode.Children = GetSheetChildNodes(name, sheetData, depth, part);
                }

                node.Children.Add(sheetNode);
            }
            // Workbook-level settings
            PopulateWorkbookSettings(node);
            Core.ThemeHandler.PopulateTheme(_doc.WorkbookPart?.ThemePart, node);
            Core.ExtendedPropertiesHandler.PopulateExtendedProperties(_doc.ExtendedFilePropertiesPart, node);

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
                return null!;

            var allDefs = definedNames.Elements<DefinedName>().ToList();
            DefinedName? dn = null;
            int dnIndex;

            if (int.TryParse(selector, out dnIndex))
            {
                if (dnIndex < 1 || dnIndex > allDefs.Count)
                    return null!;
                dn = allDefs[dnIndex - 1];
            }
            else
            {
                dn = allDefs.FirstOrDefault(d =>
                    d.Name?.Value?.Equals(selector, StringComparison.OrdinalIgnoreCase) == true);
                if (dn == null)
                    return null!;
                dnIndex = allDefs.IndexOf(dn) + 1;
            }

            var nrNode = new DocumentNode
            {
                Path = $"/namedrange[{dnIndex}]",
                Type = "namedrange",
                Text = dn.InnerText ?? dn.Name?.Value ?? "",
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
            else
            {
                // Schema declares scope get=true; emit "workbook" for workbook-scope names.
                nrNode.Format["scope"] = "workbook";
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
            throw SheetNotFoundException(sheetNameFromPath);

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
                ChildCount = data.Elements<Row>().Select(r => r.RowIndex?.Value ?? 0u).Where(i => i != 0).Distinct().Count() + (worksheet.DrawingsPart != null ? CountExcelCharts(worksheet.DrawingsPart) : 0)
            };

            // Include freeze pane info
            var ws = GetSheet(worksheet);
            var pane = ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Pane>();
            if (pane != null && pane.State?.Value == PaneStateValues.Frozen)
            {
                sheetNode.Format["freeze"] = pane.TopLeftCell?.Value ?? "";
            }

            // Include zoom and view properties
            var sheetView = ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
            if (sheetView?.ZoomScale?.HasValue == true && sheetView.ZoomScale.Value != 100)
                sheetNode.Format["zoom"] = (int)sheetView.ZoomScale.Value;
            if (sheetView?.ShowGridLines != null && !sheetView.ShowGridLines.Value)
                sheetNode.Format["gridlines"] = false;
            if (sheetView?.ShowRowColHeaders != null && !sheetView.ShowRowColHeaders.Value)
                sheetNode.Format["headings"] = false;

            // Include tab color
            var tabColor = ws.GetFirstChild<SheetProperties>()?.GetFirstChild<TabColor>();
            if (tabColor?.Rgb?.HasValue == true)
                sheetNode.Format["tabColor"] = ParseHelpers.FormatHexColor(tabColor.Rgb.Value!);
            else if (tabColor?.Theme?.HasValue == true)
            {
                // CONSISTENCY(scheme-color): echo back the symbolic name
                // (e.g. "accent1") instead of the numeric theme index.
                var schemeName = ParseHelpers.ExcelThemeIndexToName(tabColor.Theme.Value);
                if (schemeName != null) sheetNode.Format["tabColor"] = schemeName;
            }

            // Include autofilter info
            var autoFilter = ws.GetFirstChild<AutoFilter>();
            if (autoFilter?.Reference?.Value != null)
            {
                sheetNode.Format["autoFilter"] = autoFilter.Reference.Value;
            }

            // Sheet-state (hidden / very hidden) readback — lives on the
            // workbook-level Sheet element, not on the Worksheet.
            var wbSheet = GetWorkbook().GetFirstChild<Sheets>()?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetNameFromPath, StringComparison.OrdinalIgnoreCase) == true);
            if (wbSheet?.State?.Value is { } sheetState
                && (sheetState == SheetStateValues.Hidden || sheetState == SheetStateValues.VeryHidden))
                sheetNode.Format["hidden"] = true;

            // Sheet protection readback
            var sheetProtection = ws.GetFirstChild<SheetProtection>();
            if (sheetProtection?.Sheet?.Value == true)
                sheetNode.Format["protect"] = true;

            // Print settings readback
            var pageSetup = ws.GetFirstChild<PageSetup>();
            if (pageSetup != null)
            {
                if (pageSetup.Orientation?.HasValue == true)
                    sheetNode.Format["orientation"] = pageSetup.Orientation.InnerText;
                if (pageSetup.PaperSize?.HasValue == true)
                    sheetNode.Format["paperSize"] = (int)pageSetup.PaperSize.Value;
                if (pageSetup.FitToWidth?.HasValue == true)
                    sheetNode.Format["fitToPage"] = $"{pageSetup.FitToWidth.Value}x{pageSetup.FitToHeight?.Value ?? 1}";
            }

            // Print area readback
            var workbook = GetWorkbook();
            var allSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
            var sheetIdx = allSheets?.FindIndex(s =>
                s.Name?.Value?.Equals(sheetNameFromPath, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
            var printAreaDn = workbook.GetFirstChild<DefinedNames>()?.Elements<DefinedName>()
                .FirstOrDefault(d => d.Name == "_xlnm.Print_Area" && d.LocalSheetId?.Value == (uint)sheetIdx);
            if (printAreaDn != null)
            {
                // Strip "SheetName!" prefix so Get output can round-trip to Set input
                var paText = printAreaDn.Text ?? "";
                var bangIdx = paText.IndexOf('!');
                if (bangIdx >= 0) paText = paText[(bangIdx + 1)..];
                sheetNode.Format["printArea"] = paText;
            }

            // Header/Footer readback
            var headerFooter = ws.GetFirstChild<HeaderFooter>();
            if (headerFooter?.OddHeader?.Text != null)
                sheetNode.Format["header"] = headerFooter.OddHeader.Text;
            if (headerFooter?.OddFooter?.Text != null)
                sheetNode.Format["footer"] = headerFooter.OddFooter.Text;

            // Sort state readback
            var sortState = ws.GetFirstChild<SortState>();
            if (sortState != null)
            {
                var sortConditions = sortState.Elements<SortCondition>().ToList();
                var sortDesc = string.Join(",", sortConditions.Select(sc =>
                {
                    var colRef = sc.Reference?.Value?.Split(':')[0] ?? "";
                    var colName = Regex.Match(colRef, @"^([A-Z]+)").Groups[1].Value;
                    var dir = sc.Descending?.Value == true ? "desc" : "asc";
                    return $"{colName}:{dir}";
                }));
                sheetNode.Format["sort"] = sortDesc;
            }

            // Page breaks readback
            var rowBreaks = ws.GetFirstChild<RowBreaks>();
            if (rowBreaks != null && rowBreaks.Elements<Break>().Any())
            {
                var breaks = rowBreaks.Elements<Break>().Select(b => b.Id?.Value.ToString() ?? "").ToList();
                sheetNode.Format["rowBreaks"] = string.Join(",", breaks);
            }
            var colBreaks = ws.GetFirstChild<ColumnBreaks>();
            if (colBreaks != null && colBreaks.Elements<Break>().Any())
            {
                var cbreaks = colBreaks.Elements<Break>().Select(b => b.Id?.Value.ToString() ?? "").ToList();
                sheetNode.Format["colBreaks"] = string.Join(",", cbreaks);
            }

            if (depth > 0)
            {
                sheetNode.Children = GetSheetChildNodes(sheetNameFromPath, data, depth, worksheet);
            }
            return sheetNode;
        }

        var cellRef = segments[1];

        // Page break path: /Sheet1/rowbreak[N] or /Sheet1/colbreak[N]
        var rbMatch = Regex.Match(cellRef, @"^rowbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (rbMatch.Success)
        {
            var rbIdx = int.Parse(rbMatch.Groups[1].Value);
            var rowBreaks = GetSheet(worksheet).GetFirstChild<RowBreaks>();
            var breaks = rowBreaks?.Elements<Break>().ToList() ?? new();
            if (rbIdx < 1 || rbIdx > breaks.Count)
                throw new ArgumentException($"Row break index {rbIdx} out of range (1-{breaks.Count})");
            var brk = breaks[rbIdx - 1];
            return new DocumentNode
            {
                Path = path, Type = "rowbreak",
                Format = { ["row"] = brk.Id?.Value ?? 0u, ["manual"] = brk.ManualPageBreak?.Value ?? false }
            };
        }
        var cbMatch = Regex.Match(cellRef, @"^colbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cbMatch.Success)
        {
            var cbIdx = int.Parse(cbMatch.Groups[1].Value);
            var colBreaks = GetSheet(worksheet).GetFirstChild<ColumnBreaks>();
            var breaks = colBreaks?.Elements<Break>().ToList() ?? new();
            if (cbIdx < 1 || cbIdx > breaks.Count)
                throw new ArgumentException($"Column break index {cbIdx} out of range (1-{breaks.Count})");
            var brk = breaks[cbIdx - 1];
            return new DocumentNode
            {
                Path = path, Type = "colbreak",
                Format = { ["col"] = (int)(brk.Id?.Value ?? 0u), ["manual"] = brk.ManualPageBreak?.Value ?? false }
            };
        }

        // Validation path: /Sheet1/validation[N]
        var validationMatch = Regex.Match(cellRef, @"^validation\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationMatch.Success)
        {
            var dvIdx = int.Parse(validationMatch.Groups[1].Value);
            var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>();
            if (dvs == null)
                return null!;

            var dvList = dvs.Elements<DataValidation>().ToList();
            if (dvIdx < 1 || dvIdx > dvList.Count)
                return null!;

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
                    if (col.OutlineLevel?.HasValue == true && col.OutlineLevel.Value > 0)
                        colNode.Format["outlineLevel"] = (int)col.OutlineLevel.Value;
                    if (col.Collapsed?.Value == true) colNode.Format["collapsed"] = true;
                    // Long-tail CT_Col attributes (style, bestFit, phonetic, ...).
                    // Symmetric with column Set's case-preserving SetAttribute fallback.
                    FillUnknownAttrProps(col, colNode, "", CuratedColAttrs);
                }
            }
            // Include cells in this column as children (non-empty rows only)
            if (depth > 0)
            {
                var eval = new Core.FormulaEvaluator(data, _doc.WorkbookPart);
                foreach (var row in data.Elements<Row>().OrderBy(r => r.RowIndex?.Value ?? 0))
                {
                    var cell = row.Elements<Cell>().FirstOrDefault(c =>
                    {
                        if (c.CellReference?.Value == null) return false;
                        var (cn, _) = ParseCellReference(c.CellReference.Value);
                        return cn.Equals(colName, StringComparison.OrdinalIgnoreCase);
                    });
                    if (cell != null)
                        colNode.Children.Add(CellToNode(sheetNameFromPath, cell, worksheet, eval));
                }
                colNode.ChildCount = colNode.Children.Count;
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
            if (row.OutlineLevel?.HasValue == true && row.OutlineLevel.Value > 0)
                rowNode.Format["outlineLevel"] = (int)row.OutlineLevel.Value;
            if (row.Collapsed?.Value == true) rowNode.Format["collapsed"] = true;
            // Long-tail CT_Row attributes (spans, style, ph, thickTop, thickBot,
            // customFormat, ...). Symmetric with row Set's case-preserving fallback.
            FillUnknownAttrProps(row, rowNode, "", CuratedRowAttrs);
            if (depth > 0)
            {
                var eval = new Core.FormulaEvaluator(data, _doc.WorkbookPart);
                foreach (var c in row.Elements<Cell>())
                    rowNode.Children.Add(CellToNode(sheetNameFromPath, c, worksheet, eval));
            }
            return rowNode;
        }

        // Conditional formatting path: /Sheet1/cf[N]
        var cfMatch = Regex.Match(cellRef, @"^cf\[(\d+)\]$");
        if (cfMatch.Success)
        {
            var cfIdx = int.Parse(cfMatch.Groups[1].Value);
            var cfElements = GetSheet(worksheet).Elements<ConditionalFormatting>().ToList();
            if (cfIdx < 1 || cfIdx > cfElements.Count)
                return null!;

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
                    if (dbColor?.Rgb?.Value != null)
                        cfNode.Format["color"] = ParseHelpers.FormatHexColor(dbColor.Rgb.Value);
                    else if (dbColor?.Theme?.Value != null)
                        cfNode.Format["color"] = $"theme{dbColor.Theme.Value}";
                }

                // ColorScale
                var colorScale = rule.GetFirstChild<ColorScale>();
                if (colorScale != null)
                {
                    cfNode.Format["cfType"] = "colorScale";
                    var colors = colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                    if (colors.Count >= 2)
                    {
                        var minRgb = colors[0].Rgb?.Value;
                        var maxRgb = colors[^1].Rgb?.Value;
                        if (!string.IsNullOrEmpty(minRgb))
                            cfNode.Format["mincolor"] = ParseHelpers.FormatHexColor(minRgb);
                        if (!string.IsNullOrEmpty(maxRgb))
                            cfNode.Format["maxcolor"] = ParseHelpers.FormatHexColor(maxRgb);
                        if (colors.Count >= 3)
                        {
                            var midRgb = colors[1].Rgb?.Value;
                            if (!string.IsNullOrEmpty(midRgb))
                                cfNode.Format["midcolor"] = ParseHelpers.FormatHexColor(midRgb);
                        }
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

                // Top/Bottom N
                if (rule.Type?.Value == ConditionalFormatValues.Top10)
                {
                    cfNode.Format["cfType"] = "topN";
                    if (rule.Rank?.HasValue == true) cfNode.Format["rank"] = rule.Rank.Value;
                    if (rule.Bottom?.Value == true) cfNode.Format["bottom"] = true;
                    if (rule.Percent?.Value == true) cfNode.Format["percent"] = true;
                    if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
                }

                // Above/Below Average
                if (rule.Type?.Value == ConditionalFormatValues.AboveAverage)
                {
                    cfNode.Format["cfType"] = "aboveAverage";
                    if (rule.AboveAverage?.HasValue == true) cfNode.Format["aboveAverage"] = rule.AboveAverage.Value;
                    if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
                }

                // Duplicate Values
                if (rule.Type?.Value == ConditionalFormatValues.DuplicateValues)
                {
                    cfNode.Format["cfType"] = "duplicateValues";
                    if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
                }

                // Unique Values
                if (rule.Type?.Value == ConditionalFormatValues.UniqueValues)
                {
                    cfNode.Format["cfType"] = "uniqueValues";
                    if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
                }

                // Contains Text
                if (rule.Type?.Value == ConditionalFormatValues.ContainsText)
                {
                    cfNode.Format["cfType"] = "containsText";
                    if (rule.Text?.HasValue == true) cfNode.Format["text"] = rule.Text.Value;
                    if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
                }

                // Time Period (date occurring)
                if (rule.Type?.Value == ConditionalFormatValues.TimePeriod)
                {
                    cfNode.Format["cfType"] = "timePeriod";
                    if (rule.TimePeriod?.HasValue == true) cfNode.Format["period"] = rule.TimePeriod.InnerText;
                    if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
                }

                // Resolve dxfId to actual fill/font colors from the stylesheet
                if (rule.FormatId?.Value != null)
                    PopulateCfNodeFromDxf(cfNode, (int)rule.FormatId.Value);
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

        // Chart axis-by-role sub-path: /Sheet1/chart[N]/axis[@role=ROLE].
        // Per schemas/help/pptx/chart-axis.json (shared contract).
        var chartAxisGetMatch = Regex.Match(cellRef,
            @"^chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisGetMatch.Success)
        {
            var caChartIdx = int.Parse(chartAxisGetMatch.Groups[1].Value);
            var caRole = chartAxisGetMatch.Groups[2].Value;
            var caDrawingsPart = worksheet.DrawingsPart;
            if (caDrawingsPart == null)
                throw new ArgumentException($"No charts found in sheet");
            var caAllCharts = GetExcelCharts(caDrawingsPart);
            if (caChartIdx < 1 || caChartIdx > caAllCharts.Count)
                throw new ArgumentException($"Chart index {caChartIdx} out of range (1-{caAllCharts.Count})");
            var caChartInfo = caAllCharts[caChartIdx - 1];
            if (caChartInfo.IsExtended || caChartInfo.StandardPart?.ChartSpace == null)
                throw new ArgumentException($"Axis not available on chart {caChartIdx}: extended charts not supported.");
            var axisNode = ChartHelper.BuildAxisNode(caChartInfo.StandardPart.ChartSpace, caRole, path);
            if (axisNode == null)
                throw new ArgumentException($"Axis with role '{caRole}' not found on chart {caChartIdx}.");
            return axisNode;
        }

        // Chart path: /Sheet1/chart[N] or /Sheet1/chart[N]/series[K]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success)
        {
            var chartIdx = int.Parse(chartMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart;
            if (drawingsPart == null)
                throw new ArgumentException($"No charts found in sheet");

            var allCharts = GetExcelCharts(drawingsPart);
            if (chartIdx < 1 || chartIdx > allCharts.Count)
                throw new ArgumentException($"Chart index {chartIdx} out of range (1-{allCharts.Count})");

            var chartInfo = allCharts[chartIdx - 1];
            var chartNode = new DocumentNode { Path = $"/{sheetNameFromPath}/chart[{chartIdx}]", Type = "chart" };

            // BUG-R11-04: chart Get used to skip the TwoCellAnchor even though
            // `add chart --prop anchor=B2:F7` and `set ... anchor=...` both
            // support it. Round-trip requires Get to surface the anchor range
            // in the same `B2:F7` grammar. CONSISTENCY(ole-width-units) —
            // mirrors the Add/Set accepted grammar.
            var chartAnchorRange = GetChartAnchorRange(drawingsPart, chartIdx);
            if (chartAnchorRange != null)
                chartNode.Format["anchor"] = chartAnchorRange;

            if (chartInfo.IsExtended)
            {
                var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace!;
                var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                if (cxType != null) chartNode.Format["chartType"] = cxType;
                // Title
                var cxTitle = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
                var cxTitleText = cxTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
                if (cxTitleText != null) chartNode.Format["title"] = cxTitleText;
                // Count series
                var cxSeries = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Series>().ToList();
                chartNode.Format["seriesCount"] = cxSeries.Count;
            }
            else
            {
                var chart = chartInfo.StandardPart!.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                if (chart != null)
                    ChartHelper.ReadChartProperties(chart, chartNode, chartMatch.Groups[2].Success ? 1 : depth);
            }

            // If series sub-path requested, extract the specific series child
            if (chartMatch.Groups[2].Success)
            {
                var seriesIdx = int.Parse(chartMatch.Groups[2].Value);
                var seriesChildren = chartNode.Children.Where(c => c.Type == "series").ToList();
                if (seriesIdx < 1 || seriesIdx > seriesChildren.Count)
                    throw new ArgumentException($"Series {seriesIdx} not found (total: {seriesChildren.Count})");
                var seriesNode = seriesChildren[seriesIdx - 1];
                seriesNode.Path = path;
                return seriesNode;
            }
            return chartNode;
        }

        // Pivot table path: /Sheet1/pivottable[N]
        var pivotMatch = Regex.Match(cellRef, @"^pivottable\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (pivotMatch.Success)
        {
            var ptIdx = int.Parse(pivotMatch.Groups[1].Value);
            var pivotParts = worksheet.PivotTableParts.ToList();
            if (ptIdx < 1 || ptIdx > pivotParts.Count)
                throw new ArgumentException($"PivotTable index {ptIdx} out of range (1-{pivotParts.Count})");

            var pivotPart = pivotParts[ptIdx - 1];
            var ptNode = new DocumentNode { Path = path, Type = "pivottable" };
            if (pivotPart.PivotTableDefinition != null)
                PivotTableHelper.ReadPivotTableProperties(pivotPart.PivotTableDefinition, ptNode, pivotPart);
            return ptNode;
        }

        // Slicer path: /Sheet1/slicer[N]
        var slicerMatch = Regex.Match(cellRef, @"^slicer\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (slicerMatch.Success)
        {
            var slIdx = int.Parse(slicerMatch.Groups[1].Value);
            if (!TryFindSlicerByIndex(worksheet, slIdx, out var slicerElem, out var slicerCache) || slicerElem == null)
                throw new ArgumentException($"slicer[{slIdx}] not found on sheet '{sheetNameFromPath}'");
            var slNode = new DocumentNode { Path = path, Type = "slicer" };
            ReadSlicerProperties(slicerElem, slicerCache, slNode);
            return slNode;
        }

        // OLE object path: /Sheet1/ole[N]
        // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
        var oleMatch = Regex.Match(cellRef, @"^(?:ole|oleobject|object|embed)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (oleMatch.Success)
        {
            var oleIdx = int.Parse(oleMatch.Groups[1].Value);
            var oleList = CollectOleNodesForSheet(sheetNameFromPath, worksheet);
            if (oleIdx < 1 || oleIdx > oleList.Count)
                throw new ArgumentException($"OLE object {oleIdx} not found at /{sheetNameFromPath} (available: {oleList.Count}).");
            return oleList[oleIdx - 1];
        }

        // Comment path: /Sheet1/comment[N]
        var commentMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentMatch.Success)
        {
            var cmtIndex = int.Parse(commentMatch.Groups[1].Value);
            var commentsPart = worksheet.WorksheetCommentsPart;
            if (commentsPart?.Comments == null)
                return null!;

            var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
            var cmtElement = cmtList?.Elements<Comment>().ElementAtOrDefault(cmtIndex - 1);
            if (cmtElement == null)
                return null!;

            return CommentToNode(sheetNameFromPath, cmtElement, commentsPart.Comments, cmtIndex);
        }

        // Table path: /Sheet1/table[N]
        var tableMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableMatch.Success)
        {
            var tableIdx = int.Parse(tableMatch.Groups[1].Value);
            return TableToNode(sheetNameFromPath, worksheet, tableIdx, depth);
        }

        // Table column path: /Sheet1/table[N]/columns[M] or /column[M]
        var tableColMatch = Regex.Match(cellRef,
            @"^table\[(\d+)\]/(?:columns|column)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableColMatch.Success)
        {
            var tIdx = int.Parse(tableColMatch.Groups[1].Value);
            var cIdx = int.Parse(tableColMatch.Groups[2].Value);
            var tParts = worksheet.TableDefinitionParts.ToList();
            if (tIdx < 1 || tIdx > tParts.Count)
                throw new ArgumentException($"Table index {tIdx} out of range (1..{tParts.Count})");
            var tbl = tParts[tIdx - 1].Table
                ?? throw new ArgumentException($"Table {tIdx} has no definition");
            var tCols = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>().ToList();
            if (tCols == null || cIdx < 1 || cIdx > tCols.Count)
                throw new ArgumentException($"Column index {cIdx} out of range (1..{tCols?.Count ?? 0})");
            var tCol = tCols[cIdx - 1];
            var tcNode = new DocumentNode
            {
                Path = $"/{sheetNameFromPath}/table[{tIdx}]/columns[{cIdx}]",
                Type = "tableColumn",
                Text = tCol.Name?.Value ?? ""
            };
            tcNode.Format["name"] = tCol.Name?.Value ?? "";
            if (tCol.Id?.Value != null) tcNode.Format["id"] = tCol.Id.Value;
            if (tCol.TotalsRowFunction?.Value != null)
                tcNode.Format["totalFunction"] = tCol.TotalsRowFunction.Value.ToString().ToLowerInvariant();
            if (tCol.TotalsRowLabel?.Value != null)
                tcNode.Format["totalLabel"] = tCol.TotalsRowLabel.Value;
            var ccf = tCol.CalculatedColumnFormula?.Text;
            if (!string.IsNullOrEmpty(ccf)) tcNode.Format["formula"] = ccf;
            return tcNode;
        }

        // Cell reference: A1 or range A1:D10
        // Check if it's a cell reference or a generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = System.Text.RegularExpressions.Regex.IsMatch(firstPart, @"^[A-Z]+\d+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        if (!isCellRef)
        {
            // Handle sparkline[N] path segment
            var spkMatch = Regex.Match(cellRef, @"^sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (spkMatch.Success)
            {
                var spkIndex = int.Parse(spkMatch.Groups[1].Value);
                var spkGroup = GetSparklineGroup(worksheet, spkIndex)
                    ?? throw new ArgumentException($"Sparkline[{spkIndex}] not found in sheet '{sheetNameFromPath}'");
                return SparklineGroupToNode(sheetNameFromPath, spkGroup, spkIndex);
            }

            // Handle picture[N] path segment
            var picMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (picMatch.Success)
            {
                var picIndex = int.Parse(picMatch.Groups[1].Value);
                return GetPictureNode(sheetNameFromPath, worksheet, picIndex, path)!;
            }

            // Handle shape[N] path segment
            var shpMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (shpMatch.Success)
            {
                var shpIndex = int.Parse(shpMatch.Groups[1].Value);
                return GetShapeNode(sheetNameFromPath, worksheet, shpIndex, path)!;
            }


            // If it looks like it could be a malformed cell reference (digits only, etc.), reject it
            if (Regex.IsMatch(cellRef, @"^\d+$"))
                throw new ArgumentException($"Invalid cell reference: '{cellRef}'. Expected format like 'A1', 'B2'.");

            // Generic XML fallback: navigate worksheet XML tree
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {cellRef}" };
            return GenericXmlQuery.ElementToNode(target, path, depth);
        }

        // Handle /SheetName/A1/run[N] (rich text run direct access)
        var runGetMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runGetMatch.Success)
        {
            var runCellRef = runGetMatch.Groups[1].Value.ToUpperInvariant();
            var runIdx = int.Parse(runGetMatch.Groups[2].Value);
            ParseCellReference(runCellRef);
            var runCell = FindCell(data, runCellRef);
            if (runCell == null)
                throw new ArgumentException($"Cell {runCellRef} not found");
            if (runCell.DataType?.Value != CellValues.SharedString ||
                !int.TryParse(runCell.CellValue?.Text, out var sstIdx))
                throw new ArgumentException($"Cell {runCellRef} is not a rich text cell");
            var sstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi = sstPart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx);
            if (ssi == null) throw new ArgumentException($"SharedString entry {sstIdx} not found");
            var runs = ssi.Elements<Run>().ToList();
            if (runIdx < 1 || runIdx > runs.Count)
                throw new ArgumentException($"Run index {runIdx} out of range (1-{runs.Count})");
            return RunToNode(runs[runIdx - 1], $"/{sheetNameFromPath}/{runCellRef}/run[{runIdx}]");
        }

        if (cellRef.Contains(':'))
        {
            // Range — validate both endpoints
            var rangeParts = cellRef.Split(':');
            ParseCellReference(rangeParts[0]);
            if (rangeParts.Length > 1) ParseCellReference(rangeParts[1]);
            return GetCellRange(sheetNameFromPath, data, cellRef, depth, worksheet);
        }
        else
        {
            // Single cell — validate cell reference
            ParseCellReference(cellRef);
            var cell = FindCell(data, cellRef);
            if (cell == null)
            {
                var emptyNode = new DocumentNode { Path = path, Type = "cell", Text = "(empty)", Preview = cellRef };
                // Still check merge status for empty cells — they may be part of a merged range
                if (worksheet != null)
                {
                    var mergeCells = GetSheet(worksheet).GetFirstChild<MergeCells>();
                    if (mergeCells != null)
                    {
                        var mergeCell = mergeCells.Elements<MergeCell>()
                            .FirstOrDefault(m => IsCellInMergeRange(cellRef, m.Reference?.Value));
                        if (mergeCell != null)
                        {
                            var mergeRef = mergeCell.Reference?.Value ?? "";
                            emptyNode.Format["merge"] = mergeRef;
                            if (mergeRef.Split(':')[0].Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                                emptyNode.Format["mergeAnchor"] = true;
                        }
                    }
                }
                return emptyNode;
            }
            return CellToNode(sheetNameFromPath, cell, worksheet);
        }
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();

        // Handle Excel-native direct cell ref: Sheet1!A1 or Sheet1!A1:D10
        var nativeCellRef = Regex.Match(selector, @"^([^/!]+)!([A-Z]+\d+(:[A-Z]+\d+)?)$", RegexOptions.IgnoreCase);
        if (nativeCellRef.Success)
            return [Get($"/{nativeCellRef.Groups[1].Value}/{nativeCellRef.Groups[2].Value}")];

        // CONSISTENCY(excel-sheet-separator-warn): Detect the PPT-style `>`
        // separator form (e.g. `Sheet1>ole`) that users familiar with the
        // PowerPoint query grammar may try against Excel. Excel uses `!`
        // (Sheet1!cell[...]) — the legacy spreadsheet separator — so a `>`
        // in the sheet-prefix slot will silently fall through to generic
        // XML and return an empty result. We emit a single stderr warning
        // pointing to the correct `!` form, then let the normal flow run.
        // Only fire when the prefix looks like a sheet name (no `/`) and
        // the suffix is a known Excel element type we would have handled.
        {
            var pptStyle = Regex.Match(selector, @"^([^/!>]+)>(\w+)");
            if (pptStyle.Success)
            {
                var suffixType = pptStyle.Groups[2].Value.ToLowerInvariant();
                if (suffixType is "ole" or "oleobject" or "object" or "embed" or "cell" or "row"
                    or "chart" or "pivottable" or "pivot" or "slicer" or "shape"
                    or "picture" or "table" or "listobject" or "comment" or "note"
                    or "validation" or "namedrange" or "definedname" or "media"
                    or "image" or "sparkline")
                {
                    Console.Error.WriteLine(
                        $"Warning: Excel uses '!' not '>' as sheet separator " +
                        $"(e.g. '{pptStyle.Groups[1].Value}!{suffixType}' not " +
                        $"'{pptStyle.Groups[1].Value}>{suffixType}').");
                }
            }
        }

        // Check if element type is known (Scheme A) or should fall back to generic XML (Scheme B)
        // Strip sheet prefix (Sheet1!cell[...]) but not != operator
        var selectorForType = Regex.Replace(selector, @"^.+?!(?!=)", "");
        var elementMatch = Regex.Match(selectorForType, @"^(\w+)");
        // Lowercase once so all downstream `elementName is "..."` dispatch is
        // case-insensitive. CONSISTENCY(query-case-insensitive): matches how
        // WordHandler.Query normalizes selector.element to lowercase.
        var elementName = elementMatch.Success ? elementMatch.Groups[1].Value.ToLowerInvariant() : "";
        bool isKnownType = string.IsNullOrEmpty(elementName)
            // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
            || elementName is "cell" or "row" or "col" or "column" or "sheet" or "validation" or "comment" or "note" or "table" or "listobject" or "chart" or "pivottable" or "pivot" or "slicer" or "shape" or "picture" or "sparkline" or "namedrange" or "definedname" or "media" or "image" or "ole" or "oleobject" or "object" or "embed" or "hyperlink"
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

                var allCharts = GetExcelCharts(drawingsPart);
                for (int i = 0; i < allCharts.Count; i++)
                {
                    var chartInfo = allCharts[i];
                    var node = new DocumentNode { Path = $"/{sheetName}/chart[{i + 1}]", Type = "chart" };

                    if (chartInfo.IsExtended)
                    {
                        var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace!;
                        var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                        if (cxType != null) node.Format["chartType"] = cxType;
                        var cxTitle = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
                        var cxTitleText = cxTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
                        if (cxTitleText != null) node.Format["title"] = cxTitleText;
                        var cxSeries = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Series>().ToList();
                        node.Format["seriesCount"] = cxSeries.Count;
                    }
                    else
                    {
                        var chart = chartInfo.StandardPart!.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                        if (chart != null)
                            ChartHelper.ReadChartProperties(chart, node, 0);
                    }

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

        // Handle sheet queries
        if (elementName == "sheet")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var sheetNode = new DocumentNode { Path = $"/{sheetName}", Type = "sheet", Preview = sheetName };
                var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
                var rowCount = sheetData?.Elements<Row>().Count() ?? 0;
                var chartCount = worksheetPart.DrawingsPart != null ? CountExcelCharts(worksheetPart.DrawingsPart) : 0;
                sheetNode.ChildCount = rowCount + chartCount;
                results.Add(sheetNode);
            }
            return results;
        }

        // Handle pivottable queries
        if (elementName == "pivottable" || elementName == "pivot")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var pivotParts = worksheetPart.PivotTableParts.ToList();
                for (int i = 0; i < pivotParts.Count; i++)
                {
                    var node = new DocumentNode { Path = $"/{sheetName}/pivottable[{i + 1}]", Type = "pivottable" };
                    var pivotDef = pivotParts[i].PivotTableDefinition;
                    if (pivotDef != null)
                        PivotTableHelper.ReadPivotTableProperties(pivotDef, node, pivotParts[i]);

                    if (parsed.ValueContains != null)
                    {
                        var name = node.Format.TryGetValue("name", out var n) ? n?.ToString() : null;
                        if (name == null || !name.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle slicer queries
        if (elementName == "slicer")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var slicersPart = worksheetPart.GetPartsOfType<SlicersPart>().FirstOrDefault();
                if (slicersPart?.Slicers == null) continue;

                var slicers = slicersPart.Slicers.Elements<X14.Slicer>().ToList();
                for (int i = 0; i < slicers.Count; i++)
                {
                    if (!TryFindSlicerByIndex(worksheetPart, i + 1, out var slElem, out var slCache) || slElem == null)
                        continue;
                    var node = new DocumentNode
                    {
                        Path = $"/{sheetName}/slicer[{i + 1}]",
                        Type = "slicer"
                    };
                    ReadSlicerProperties(slElem, slCache, node);

                    if (parsed.ValueContains != null)
                    {
                        var nm = node.Format.TryGetValue("name", out var n) ? n?.ToString() : null;
                        if (nm == null || !nm.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle sparkline queries
        if (elementName == "sparkline")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var ws = GetSheet(worksheetPart);
                var extList = ws.GetFirstChild<WorksheetExtensionList>();
                if (extList == null) continue;

                var spkExt = extList.Elements<WorksheetExtension>()
                    .FirstOrDefault(e => e.Uri == "{05C60535-1F16-4fd2-B633-E4A46CF9E463}");
                if (spkExt == null) continue;

                var spkGroups = spkExt.GetFirstChild<X14.SparklineGroups>();
                if (spkGroups == null) continue;

                var groups = spkGroups.Elements<X14.SparklineGroup>().ToList();
                for (int i = 0; i < groups.Count; i++)
                    results.Add(SparklineGroupToNode(sheetName, groups[i], i + 1));
            }
            return results;
        }

        // Handle shape queries
        if (elementName == "shape")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) continue;

                var shpAnchors = drawingsPart.WorksheetDrawing
                    .Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
                    .Where(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().Any())
                    .ToList();

                for (int i = 0; i < shpAnchors.Count; i++)
                {
                    var node = GetShapeNode(sheetName, worksheetPart, i + 1, $"/{sheetName}/shape[{i + 1}]");
                    if (node == null) continue;

                    if (parsed.ValueContains != null)
                    {
                        if (node.Text == null || !node.Text.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle OLE object queries. Excel stores OLE objects in two
        // parallel structures:
        //   1. <oleObjects> inside the worksheet (schema-typed OleObject
        //      elements with progId + shapeId + r:id)
        //   2. EmbeddedObjectParts/EmbeddedPackageParts on the WorksheetPart
        //      (the actual binary payloads, joined via rel id)
        // We enumerate (1) as the source of truth for path indexing and
        // join (2) for contentType/fileSize enrichment. Worksheets that
        // somehow have orphan parts without a matching oleObjects entry
        // are still surfaced from the parts side so nothing is missed.
        // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
        if (elementName is "ole" or "oleobject" or "object" or "embed")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var oleNodes = CollectOleNodesForSheet(sheetName, worksheetPart);
                foreach (var node in oleNodes)
                {
                    if (parsed.ValueContains != null)
                    {
                        var pid = node.Format.TryGetValue("progId", out var p) ? p?.ToString() : null;
                        if (pid == null || !pid.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle picture queries
        if (elementName == "picture")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) continue;

                var picAnchors = drawingsPart.WorksheetDrawing
                    .Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
                    .Where(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().Any())
                    .ToList();

                for (int i = 0; i < picAnchors.Count; i++)
                {
                    var node = GetPictureNode(sheetName, worksheetPart, i + 1, $"/{sheetName}/picture[{i + 1}]");
                    if (node == null) continue;

                    if (parsed.ValueContains != null)
                    {
                        var alt = node.Format.TryGetValue("alt", out var a) ? a?.ToString() : null;
                        if (alt == null || !alt.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle media/image queries
        if (elementName is "media" or "image")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) continue;

                var picAnchors = drawingsPart.WorksheetDrawing
                    .Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
                    .Where(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().Any())
                    .ToList();

                for (int i = 0; i < picAnchors.Count; i++)
                {
                    var node = GetPictureNode(sheetName, worksheetPart, i + 1, $"/{sheetName}/picture[{i + 1}]");
                    if (node == null) continue;

                    // Add content type from image part
                    var pic = picAnchors[i].Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().First();
                    var blip = pic.BlipFill?.Blip;
                    if (blip?.Embed?.Value != null)
                    {
                        var part = drawingsPart.GetPartById(blip.Embed.Value);
                        if (part != null)
                        {
                            node.Format["contentType"] = part.ContentType;
                            node.Format["fileSize"] = part.GetStream().Length;
                        }
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle column queries. OOXML stores columns as <col min=".." max="..">,
        // which can span a range of column indices. We expand spans into one
        // DocumentNode per concrete column so `/SheetName/col[X]` paths align
        // with the Get path format.
        if (elementName is "col" or "column")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var columns = GetSheet(worksheetPart).GetFirstChild<Columns>();
                if (columns == null) continue;

                foreach (var col in columns.Elements<Column>())
                {
                    var min = col.Min?.Value ?? 0u;
                    var max = col.Max?.Value ?? min;
                    if (min == 0) continue;
                    for (uint ci = min; ci <= max; ci++)
                    {
                        var colName = IndexToColumnName((int)ci);
                        var node = new DocumentNode
                        {
                            Path = $"/{sheetName}/col[{colName}]",
                            Type = "column",
                            Preview = colName
                        };
                        if (col.Width?.Value != null) node.Format["width"] = col.Width.Value;
                        if (col.Hidden?.Value == true) node.Format["hidden"] = true;
                        if (col.CustomWidth?.Value == true) node.Format["customWidth"] = true;
                        if (col.OutlineLevel?.HasValue == true && col.OutlineLevel.Value > 0)
                            node.Format["outlineLevel"] = (int)col.OutlineLevel.Value;
                        if (col.Collapsed?.Value == true) node.Format["collapsed"] = true;
                        if (MatchesFormatAttributes(node, parsed))
                            results.Add(node);
                    }
                }
            }
            return results;
        }

        // Handle hyperlink queries. In xlsx, hyperlinks are cell-level metadata
        // (worksheet <hyperlinks><hyperlink ref=".." r:id=".."/></hyperlinks>),
        // not standalone addressable elements. We surface them as discoverable
        // nodes whose Path points at the owning cell so the agent can Get/Set
        // the hyperlink via cell `link` / `tooltip` / `display` props.
        // CONSISTENCY(xlsx-hyperlink-cell-backed): Add/Set live on cells, not here.
        if (elementName is "hyperlink")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var ws = GetSheet(worksheetPart);
                var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                if (hyperlinksEl == null) continue;

                foreach (var hl in hyperlinksEl.Elements<Hyperlink>())
                {
                    var cellRef = hl.Reference?.Value ?? "";
                    var node = new DocumentNode
                    {
                        Path = string.IsNullOrEmpty(cellRef) ? $"/{sheetName}" : $"/{sheetName}/{cellRef}",
                        Type = "hyperlink",
                        Preview = cellRef
                    };
                    if (!string.IsNullOrEmpty(cellRef)) node.Format["ref"] = cellRef;
                    // Resolve external URL via relationship id
                    if (hl.Id?.Value != null)
                    {
                        try
                        {
                            var rel = worksheetPart.HyperlinkRelationships
                                .FirstOrDefault(r => r.Id == hl.Id.Value);
                            if (rel != null) node.Format["url"] = rel.Uri.ToString();
                        }
                        catch { }
                    }
                    if (hl.Location?.Value != null) node.Format["location"] = hl.Location.Value;
                    if (hl.Display?.Value != null) node.Format["display"] = hl.Display.Value;
                    if (hl.Tooltip?.Value != null) node.Format["tooltip"] = hl.Tooltip.Value;

                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle namedrange / definedname queries
        if (elementName is "namedrange" or "definedname")
        {
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                var allDefs = definedNames.Elements<DefinedName>().ToList();
                for (int i = 0; i < allDefs.Count; i++)
                {
                    var dn = allDefs[i];
                    var nrNode = new DocumentNode
                    {
                        Path = $"/namedrange[{i + 1}]",
                        Type = "namedrange",
                        Text = dn.InnerText ?? dn.Name?.Value ?? "",
                        Preview = dn.InnerText
                    };
                    if (dn.Name?.Value != null) nrNode.Format["name"] = dn.Name.Value;
                    nrNode.Format["ref"] = dn.InnerText ?? "";
                    if (dn.LocalSheetId?.HasValue == true)
                    {
                        var sheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                        if (sheets != null && (int)dn.LocalSheetId.Value < sheets.Count)
                            nrNode.Format["scope"] = sheets[(int)dn.LocalSheetId.Value].Name?.Value ?? "";
                    }
                    else
                    {
                        // Schema declares scope get=true; emit "workbook" for workbook-scope names.
                        nrNode.Format["scope"] = "workbook";
                    }
                    if (dn.Comment?.HasValue == true) nrNode.Format["comment"] = dn.Comment!.Value!;

                    if (parsed.ValueContains != null)
                    {
                        var name = dn.Name?.Value ?? "";
                        if (!name.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(nrNode);
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

            var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (MatchesCellSelector(cell, sheetName, parsed))
                    {
                        var node = CellToNode(sheetName, cell, worksheetPart, eval);
                        if (MatchesFormatAttributes(node, parsed))
                            results.Add(node);
                    }
                }
            }
        }

        return results;
    }

    // ==================== CF DXF resolution ====================

    /// <summary>
    /// Resolves a conditional formatting rule's dxfId to fill and font colors
    /// from the workbook stylesheet, and populates the DocumentNode accordingly.
    /// </summary>
    private void PopulateCfNodeFromDxf(DocumentNode cfNode, int dxfId)
    {
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        if (stylesheet == null) return;

        var dxfs = stylesheet.GetFirstChild<DifferentialFormats>();
        if (dxfs == null) return;

        var dxfList = dxfs.Elements<DifferentialFormat>().ToList();
        if (dxfId < 0 || dxfId >= dxfList.Count) return;

        var dxf = dxfList[dxfId];

        // Resolve fill color
        var fill = dxf.GetFirstChild<Fill>();
        if (fill != null)
        {
            var patternFill = fill.GetFirstChild<PatternFill>();
            if (patternFill != null)
            {
                var bgColor = patternFill.GetFirstChild<BackgroundColor>();
                if (bgColor?.Rgb?.Value != null)
                    cfNode.Format["fill"] = ParseHelpers.FormatHexColor(bgColor.Rgb.Value);
                else
                {
                    var fgColor = patternFill.GetFirstChild<ForegroundColor>();
                    if (fgColor?.Rgb?.Value != null)
                        cfNode.Format["fill"] = ParseHelpers.FormatHexColor(fgColor.Rgb.Value);
                }
            }
        }

        // Resolve font color
        var font = dxf.GetFirstChild<Font>();
        if (font != null)
        {
            var fontColor = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
            if (fontColor?.Rgb?.Value != null)
                cfNode.Format["font.color"] = ParseHelpers.FormatHexColor(fontColor.Rgb.Value);
        }
    }
}
