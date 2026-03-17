// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
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
            DefinedName? dn;

            if (int.TryParse(selector, out var dnIndex))
            {
                if (dnIndex < 1 || dnIndex > allDefs.Count)
                    throw new ArgumentException($"Named range index {dnIndex} out of range (1-{allDefs.Count})");
                dn = allDefs[dnIndex - 1];
            }
            else
            {
                dn = allDefs.FirstOrDefault(d =>
                    d.Name?.Value?.Equals(selector, StringComparison.OrdinalIgnoreCase) == true)
                    ?? throw new ArgumentException($"Named range '{selector}' not found");
            }

            var nrUnsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "ref": dn.Text = value; break;
                    case "name": dn.Name = value; break;
                    case "comment": dn.Comment = value; break;
                    default: nrUnsupported.Add(key); break;
                }
            }

            workbook.Save();
            return nrUnsupported;
        }

        // Parse path: /SheetName, /SheetName/A1, /SheetName/A1:D1, /SheetName/col[A], /SheetName/row[1], /SheetName/autofilter
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null)
            throw new ArgumentException($"Sheet not found: {sheetName}");

        // Sheet-level Set (path is just /SheetName)
        if (segments.Length < 2)
        {
            return SetSheetLevel(worksheet, properties);
        }

        var cellRef = segments[1];

        // Handle /SheetName/validation[N]
        var validationSetMatch = Regex.Match(cellRef, @"^validation\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationSetMatch.Success)
        {
            var dvIdx = int.Parse(validationSetMatch.Groups[1].Value);
            var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>();
            if (dvs == null)
                throw new ArgumentException("No data validations found in sheet");

            var dvList = dvs.Elements<DataValidation>().ToList();
            if (dvIdx < 1 || dvIdx > dvList.Count)
                throw new ArgumentException($"Validation index {dvIdx} out of range (1-{dvList.Count})");

            var dv = dvList[dvIdx - 1];
            var dvUnsupported = new List<string>();

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "sqref":
                        dv.SequenceOfReferences = new ListValue<StringValue>(
                            value.Split(' ').Select(s => new StringValue(s)));
                        break;
                    case "type":
                        dv.Type = value.ToLowerInvariant() switch
                        {
                            "list" => DataValidationValues.List,
                            "whole" => DataValidationValues.Whole,
                            "decimal" => DataValidationValues.Decimal,
                            "date" => DataValidationValues.Date,
                            "time" => DataValidationValues.Time,
                            "textlength" => DataValidationValues.TextLength,
                            "custom" => DataValidationValues.Custom,
                            _ => throw new ArgumentException($"Unknown validation type: {value}")
                        };
                        break;
                    case "formula1":
                        if (dv.Type?.Value == DataValidationValues.List && !value.StartsWith("\""))
                            dv.Formula1 = new Formula1($"\"{value}\"");
                        else
                            dv.Formula1 = new Formula1(value);
                        break;
                    case "formula2":
                        dv.Formula2 = new Formula2(value);
                        break;
                    case "allowblank":
                        dv.AllowBlank = IsTruthy(value);
                        break;
                    case "showerror":
                        dv.ShowErrorMessage = IsTruthy(value);
                        break;
                    case "errortitle":
                        dv.ErrorTitle = value;
                        break;
                    case "error":
                        dv.Error = value;
                        break;
                    case "showinput":
                        dv.ShowInputMessage = IsTruthy(value);
                        break;
                    case "prompttitle":
                        dv.PromptTitle = value;
                        break;
                    case "prompt":
                        dv.Prompt = value;
                        break;
                    default:
                        dvUnsupported.Add(key);
                        break;
                }
            }

            SaveWorksheet(worksheet);
            return dvUnsupported;
        }

        // Handle /SheetName/picture[N]
        var picSetMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (picSetMatch.Success)
        {
            var picIdx = int.Parse(picSetMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/pictures");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/pictures");

            var picAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
                .Where(a => a.Descendants<XDR.Picture>().Any()).ToList();
            if (picIdx < 1 || picIdx > picAnchors.Count)
                throw new ArgumentException($"Picture index {picIdx} out of range (1..{picAnchors.Count})");

            var anchor = picAnchors[picIdx - 1];
            var picUnsupported = new List<string>();

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "x":
                        if (anchor.FromMarker != null)
                            anchor.FromMarker.ColumnId!.Text = value;
                        break;
                    case "y":
                        if (anchor.FromMarker != null)
                            anchor.FromMarker.RowId!.Text = value;
                        break;
                    case "width":
                        if (anchor.FromMarker != null && anchor.ToMarker != null)
                        {
                            var fromCol = int.TryParse(anchor.FromMarker.ColumnId?.Text, out var fc) ? fc : 0;
                            anchor.ToMarker.ColumnId!.Text = (fromCol + int.Parse(value)).ToString();
                        }
                        break;
                    case "height":
                        if (anchor.FromMarker != null && anchor.ToMarker != null)
                        {
                            var fromRow = int.TryParse(anchor.FromMarker.RowId?.Text, out var fr) ? fr : 0;
                            anchor.ToMarker.RowId!.Text = (fromRow + int.Parse(value)).ToString();
                        }
                        break;
                    case "alt":
                        var nvProps = anchor.Descendants<XDR.NonVisualDrawingProperties>().FirstOrDefault();
                        if (nvProps != null) nvProps.Description = value;
                        break;
                    default:
                        picUnsupported.Add(key);
                        break;
                }
            }

            drawingsPart.WorksheetDrawing.Save();
            return picUnsupported;
        }

        // Handle /SheetName/table[N]
        var tableSetMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableSetMatch.Success)
        {
            var tableIdx = int.Parse(tableSetMatch.Groups[1].Value);
            var tableParts = worksheet.TableDefinitionParts.ToList();
            if (tableIdx < 1 || tableIdx > tableParts.Count)
                throw new ArgumentException($"Table index {tableIdx} out of range (1..{tableParts.Count})");

            var table = tableParts[tableIdx - 1].Table
                ?? throw new ArgumentException($"Table {tableIdx} has no definition");

            var tblUnsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "name": table.Name = value; break;
                    case "displayname": table.DisplayName = value; break;
                    case "style":
                        var styleInfo = table.GetFirstChild<TableStyleInfo>();
                        if (styleInfo != null) styleInfo.Name = value;
                        else table.AppendChild(new TableStyleInfo
                        {
                            Name = value, ShowFirstColumn = false, ShowLastColumn = false,
                            ShowRowStripes = true, ShowColumnStripes = false
                        });
                        break;
                    case "ref":
                        table.Reference = value.ToUpperInvariant();
                        var af = table.GetFirstChild<AutoFilter>();
                        if (af != null) af.Reference = value.ToUpperInvariant();
                        break;
                    default: tblUnsupported.Add(key); break;
                }
            }

            tableParts[tableIdx - 1].Table!.Save();
            return tblUnsupported;
        }

        // Handle /SheetName/comment[N]
        var commentSetMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentSetMatch.Success)
        {
            var cmtIndex = int.Parse(commentSetMatch.Groups[1].Value);
            var commentsPart = worksheet.WorksheetCommentsPart;
            if (commentsPart?.Comments == null)
                throw new ArgumentException($"No comments found in sheet: {sheetName}");

            var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
            var cmtElement = cmtList?.Elements<Comment>().ElementAtOrDefault(cmtIndex - 1)
                ?? throw new ArgumentException($"Comment [{cmtIndex}] not found");

            var cmtUnsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        cmtElement.CommentText = new CommentText(
                            new Run(
                                new RunProperties(new FontSize { Val = 9 }, new Color { Indexed = 81 },
                                    new RunFont { Val = "Tahoma" }),
                                new Text(value) { Space = SpaceProcessingModeValues.Preserve }
                            )
                        );
                        break;
                    case "author":
                        var authors = commentsPart.Comments.GetFirstChild<Authors>()!;
                        var existingAuthors = authors.Elements<Author>().ToList();
                        var aIdx = existingAuthors.FindIndex(a => a.Text == value);
                        if (aIdx >= 0)
                            cmtElement.AuthorId = (uint)aIdx;
                        else
                        {
                            authors.AppendChild(new Author(value));
                            cmtElement.AuthorId = (uint)existingAuthors.Count;
                        }
                        break;
                    default:
                        cmtUnsupported.Add(key);
                        break;
                }
            }

            commentsPart.Comments.Save();
            return cmtUnsupported;
        }

        // Handle /SheetName/autofilter
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            return SetAutoFilter(worksheet, properties);
        }

        // Handle /SheetName/cf[N]
        var cfSetMatch = Regex.Match(cellRef, @"^cf\[(\d+)\]$");
        if (cfSetMatch.Success)
        {
            var cfIdx = int.Parse(cfSetMatch.Groups[1].Value);
            var ws = GetSheet(worksheet);
            var cfElements = ws.Elements<ConditionalFormatting>().ToList();
            if (cfIdx < 1 || cfIdx > cfElements.Count)
                throw new ArgumentException($"CF {cfIdx} not found (total: {cfElements.Count})");

            var cf = cfElements[cfIdx - 1];
            var unsup = new List<string>();
            var rule = cf.Elements<ConditionalFormattingRule>().FirstOrDefault();

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "sqref":
                        cf.SequenceOfReferences = new ListValue<StringValue>(
                            value.Split(' ').Select(s => new StringValue(s)));
                        break;
                    case "color":
                        var dbColor = rule?.GetFirstChild<DataBar>()?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                        if (dbColor != null) dbColor.Rgb = (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant();
                        else unsup.Add(key);
                        break;
                    case "mincolor":
                        var csColors = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                        if (csColors != null && csColors.Count >= 2)
                            csColors[0].Rgb = (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant();
                        else unsup.Add(key);
                        break;
                    case "maxcolor":
                        var csColors2 = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                        if (csColors2 != null && csColors2.Count >= 2)
                            csColors2[^1].Rgb = (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant();
                        else unsup.Add(key);
                        break;
                    case "iconset":
                        var iconSetEl = rule?.GetFirstChild<IconSet>();
                        if (iconSetEl != null)
                            iconSetEl.IconSetValue = new EnumValue<IconSetValues>(new IconSetValues(value));
                        else unsup.Add(key);
                        break;
                    case "reverse":
                        var isEl = rule?.GetFirstChild<IconSet>();
                        if (isEl != null) isEl.Reverse = bool.Parse(value);
                        else unsup.Add(key);
                        break;
                    case "showvalue":
                        var isEl2 = rule?.GetFirstChild<IconSet>();
                        if (isEl2 != null) isEl2.ShowValue = bool.Parse(value);
                        else unsup.Add(key);
                        break;
                    default:
                        unsup.Add(key);
                        break;
                }
            }
            ReorderWorksheetChildren(ws); ws.Save();
            return unsup;
        }

        // Handle /SheetName/col[X]
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Z]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colName = colMatch.Groups[1].Value.ToUpperInvariant();
            return SetColumn(worksheet, colName, properties);
        }

        // Handle /SheetName/row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            return SetRow(worksheet, rowIdx, properties);
        }

        // Handle /SheetName/chart[N]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartIdx = int.Parse(chartMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart;
            if (drawingsPart == null)
                throw new ArgumentException("No charts in this sheet");
            var chartParts = drawingsPart.ChartParts.ToList();
            if (chartIdx < 1 || chartIdx > chartParts.Count)
                throw new ArgumentException($"Chart {chartIdx} not found");
            var chartPart = chartParts[chartIdx - 1];

            var unsup = new List<string>();
            var chart = chartPart.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            if (chart == null) return unsup;

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "title":
                        var titleEl = chart.Title;
                        if (titleEl != null)
                        {
                            var titleRun = titleEl.Descendants<DocumentFormat.OpenXml.Drawing.Run>().FirstOrDefault();
                            if (titleRun?.Text != null) titleRun.Text.Text = value;
                        }
                        break;
                    default:
                        unsup.Add(key);
                        break;
                }
            }
            chartPart.ChartSpace?.Save();
            return unsup;
        }

        // Handle /SheetName/A1:D1 (range — merge/unmerge)
        if (cellRef.Contains(':'))
        {
            var firstPartRange = cellRef.Split(':')[0];
            bool isRangeRef = Regex.IsMatch(firstPartRange, @"^[A-Z]+\d+$", RegexOptions.IgnoreCase);
            if (isRangeRef)
            {
                return SetRange(worksheet, cellRef.ToUpperInvariant(), properties);
            }
        }

        // Check if path is a cell reference or generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = Regex.IsMatch(firstPart, @"^[A-Z]+\d+", RegexOptions.IgnoreCase);
        if (!isCellRef)
        {
            // Generic XML fallback: navigate to element and set attributes
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                throw new ArgumentException($"Element not found: {cellRef}");
            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            SaveWorksheet(worksheet);
            return unsup;
        }

        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            GetSheet(worksheet).Append(sheetData);
        }

        var cell = FindOrCreateCell(sheetData, cellRef);

        // Separate content props from style props
        var styleProps = new Dictionary<string, string>();
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            if (ExcelStyleManager.IsStyleKey(key))
            {
                styleProps[key] = value;
                continue;
            }

            switch (key.ToLowerInvariant())
            {
                case "value":
                    cell.CellValue = new CellValue(value);
                    // Auto-detect type
                    if (double.TryParse(value, out _))
                        cell.DataType = null; // Number is default
                    else
                    {
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                    break;
                case "formula":
                    cell.CellFormula = new CellFormula(value);
                    cell.CellValue = null;
                    break;
                case "type":
                    cell.DataType = value.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        _ => cell.DataType
                    };
                    break;
                case "clear":
                    cell.CellValue = null;
                    cell.CellFormula = null;
                    break;
                case "link":
                {
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        hyperlinksEl?.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                    }
                    else
                    {
                        var hlRel = worksheet.AddHyperlinkRelationship(new Uri(value), isExternal: true);
                        if (hyperlinksEl == null)
                        {
                            hyperlinksEl = new Hyperlinks();
                            ws.AppendChild(hyperlinksEl);
                        }
                        hyperlinksEl.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                        hyperlinksEl.AppendChild(new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id });
                    }
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                        unsupported.Add(key);
                    break;
            }
        }

        // Apply style properties if any
        if (styleProps.Count > 0)
        {
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var styleManager = new ExcelStyleManager(workbookPart);
            cell.StyleIndex = styleManager.ApplyStyle(cell, styleProps);
        }

        SaveWorksheet(worksheet);
        return unsupported;
    }

    // ==================== Sheet-level Set (freeze panes) ====================

    private List<string> SetSheetLevel(WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "freeze":
                {
                    var sheetViews = ws.GetFirstChild<SheetViews>();
                    if (sheetViews == null)
                    {
                        sheetViews = new SheetViews();
                        ws.InsertAt(sheetViews, 0);
                    }
                    var sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null)
                    {
                        sheetView = new SheetView { WorkbookViewId = 0 };
                        sheetViews.AppendChild(sheetView);
                    }

                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove freeze
                        var existingPane = sheetView.GetFirstChild<Pane>();
                        existingPane?.Remove();
                    }
                    else
                    {
                        // Parse cell reference for freeze position
                        // "A2" = freeze row 1, "B1" = freeze col A, "B2" = freeze row 1 + col A
                        var (col, row) = ParseCellReference(value.ToUpperInvariant());
                        var colSplit = ColumnNameToIndex(col) - 1; // 0-based: B=1 means split at 1
                        var rowSplit = row - 1; // 0-based: 2 means split at 1

                        // Remove existing pane
                        var existingPane = sheetView.GetFirstChild<Pane>();
                        existingPane?.Remove();

                        var activePane = (colSplit > 0 && rowSplit > 0) ? PaneValues.BottomRight
                            : (rowSplit > 0) ? PaneValues.BottomLeft
                            : PaneValues.TopRight;

                        var pane = new Pane
                        {
                            TopLeftCell = value.ToUpperInvariant(),
                            State = PaneStateValues.Frozen,
                            ActivePane = activePane
                        };
                        if (rowSplit > 0) pane.VerticalSplit = rowSplit;
                        if (colSplit > 0) pane.HorizontalSplit = colSplit;

                        sheetView.InsertAt(pane, 0);
                    }
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ReorderWorksheetChildren(ws); ws.Save();
        return unsupported;
    }

    // ==================== Range Set (merge/unmerge) ====================

    private List<string> SetRange(WorksheetPart worksheet, string rangeRef, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "merge":
                {
                    bool doMerge = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);

                    if (doMerge)
                    {
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        if (mergeCells == null)
                        {
                            mergeCells = new MergeCells();
                            ws.AppendChild(mergeCells);
                        }

                        // Avoid duplicate
                        var existing = mergeCells.Elements<MergeCell>()
                            .FirstOrDefault(m => m.Reference?.Value?.Equals(rangeRef, StringComparison.OrdinalIgnoreCase) == true);
                        if (existing == null)
                        {
                            mergeCells.AppendChild(new MergeCell { Reference = rangeRef });
                        }
                    }
                    else
                    {
                        // Unmerge: remove the MergeCell for this range
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        if (mergeCells != null)
                        {
                            var mc = mergeCells.Elements<MergeCell>()
                                .FirstOrDefault(m => m.Reference?.Value?.Equals(rangeRef, StringComparison.OrdinalIgnoreCase) == true);
                            mc?.Remove();

                            // Remove empty MergeCells element
                            if (!mergeCells.HasChildren)
                                mergeCells.Remove();
                        }
                    }
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ReorderWorksheetChildren(ws);
        ReorderWorksheetChildren(ws); ws.Save();
        return unsupported;
    }

    // ==================== Column Set (width, hidden) ====================

    private List<string> SetColumn(WorksheetPart worksheet, string colName, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);
        var colIdx = (uint)ColumnNameToIndex(colName);

        var columns = ws.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData != null)
                ws.InsertBefore(columns, sheetData);
            else
                ws.AppendChild(columns);
        }

        // Find existing column definition or create one
        var col = columns.Elements<Column>()
            .FirstOrDefault(c => c.Min?.Value <= colIdx && c.Max?.Value >= colIdx);
        if (col == null)
        {
            col = new Column { Min = colIdx, Max = colIdx, Width = 8.43, CustomWidth = true };
            columns.AppendChild(col);
        }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "width":
                    col.Width = double.Parse(value);
                    col.CustomWidth = true;
                    break;
                case "hidden":
                    col.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ReorderWorksheetChildren(ws); ws.Save();
        return unsupported;
    }

    // ==================== Row Set (height, hidden) ====================

    private List<string> SetRow(WorksheetPart worksheet, uint rowIdx, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
            throw new ArgumentException("Sheet has no data");

        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null)
        {
            // Create the row
            row = new Row { RowIndex = rowIdx };
            var afterRow = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < rowIdx);
            if (afterRow != null)
                afterRow.InsertAfterSelf(row);
            else
                sheetData.InsertAt(row, 0);
        }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "height":
                    row.Height = double.Parse(value);
                    row.CustomHeight = true;
                    break;
                case "hidden":
                    row.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ReorderWorksheetChildren(ws); ws.Save();
        return unsupported;
    }

    // ==================== AutoFilter Set ====================

    private List<string> SetAutoFilter(WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "range":
                {
                    var autoFilter = ws.GetFirstChild<AutoFilter>();
                    if (autoFilter == null)
                    {
                        autoFilter = new AutoFilter();
                        // AutoFilter goes after SheetData (after MergeCells if present)
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        var sheetData = ws.GetFirstChild<SheetData>();
                        if (mergeCells != null)
                            mergeCells.InsertAfterSelf(autoFilter);
                        else if (sheetData != null)
                            sheetData.InsertAfterSelf(autoFilter);
                        else
                            ws.AppendChild(autoFilter);
                    }
                    autoFilter.Reference = value.ToUpperInvariant();
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ReorderWorksheetChildren(ws); ws.Save();
        return unsupported;
    }
}
