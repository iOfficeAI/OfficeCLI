// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using SpreadsheetDrawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        // Normalize to case-insensitive lookup so camelCase keys (e.g. minColor) match lowercase lookups
        if (properties != null && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        parentPath = NormalizeExcelPath(parentPath);
        switch (type.ToLowerInvariant())
        {
            case "sheet":
                var workbookPart = _doc.WorkbookPart
                    ?? throw new InvalidOperationException("Workbook not found");
                var sheets = GetWorkbook().GetFirstChild<Sheets>()
                    ?? GetWorkbook().AppendChild(new Sheets());

                var name = properties.GetValueOrDefault("name", $"Sheet{sheets.Elements<Sheet>().Count() + 1}");
                var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                newWorksheetPart.Worksheet.Save();

                var sheetId = sheets.Elements<Sheet>().Any()
                    ? sheets.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1
                    : 1;
                var relId = workbookPart.GetIdOfPart(newWorksheetPart);

                sheets.AppendChild(new Sheet { Id = relId, SheetId = (uint)sheetId, Name = name });
                GetWorkbook().Save();
                return $"/{name}";

            case "row":
                var segments = parentPath.TrimStart('/').Split('/', 2);
                var sheetName = segments[0];
                var worksheet = FindWorksheet(sheetName)
                    ?? throw new ArgumentException($"Sheet not found: {sheetName}");
                var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(worksheet).AppendChild(new SheetData());

                var rowIdx = index ?? ((int)(sheetData.Elements<Row>().LastOrDefault()?.RowIndex?.Value ?? 0) + 1);

                // If inserting at an existing position, shift rows down first
                bool needsShift = index.HasValue && sheetData.Elements<Row>().Any(r => r.RowIndex?.Value >= (uint)rowIdx);
                if (needsShift)
                    ShiftRowsDown(worksheet, rowIdx);

                var newRow = new Row { RowIndex = (uint)rowIdx };

                // Create cells if cols specified
                if (properties.TryGetValue("cols", out var colsStr))
                {
                    if (!int.TryParse(colsStr, out var cols) || cols <= 0)
                        throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Expected a positive integer (number of columns to create).");
                    for (int c = 0; c < cols; c++)
                    {
                        var colLetter = IndexToColumnName(c + 1);
                        newRow.AppendChild(new Cell { CellReference = $"{colLetter}{rowIdx}" });
                    }
                }

                // Re-fetch sheetData after potential shift
                sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(worksheet).AppendChild(new SheetData());
                var afterRow = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < (uint)rowIdx);
                if (afterRow != null)
                    afterRow.InsertAfterSelf(newRow);
                else
                    sheetData.InsertAt(newRow, 0);

                SaveWorksheet(worksheet);
                return $"/{sheetName}/row[{rowIdx}]";

            case "cell":
                var cellSegments = parentPath.TrimStart('/').Split('/', 2);
                var cellSheetName = cellSegments[0];
                var cellWorksheet = FindWorksheet(cellSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cellSheetName}");
                var cellSheetData = GetSheet(cellWorksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(cellWorksheet).AppendChild(new SheetData());

                string cellRef;
                if (properties.ContainsKey("ref"))
                {
                    cellRef = properties["ref"];
                }
                else if (properties.ContainsKey("address"))
                {
                    cellRef = properties["address"];
                }
                else
                {
                    // Auto-assign next available cell in row 1
                    var existingRefs = cellSheetData.Descendants<Cell>()
                        .Where(c => c.CellReference?.Value != null)
                        .Select(c => c.CellReference!.Value!)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    int colIdx = 1;
                    while (existingRefs.Contains(IndexToColumnName(colIdx) + "1"))
                        colIdx++;
                    cellRef = IndexToColumnName(colIdx) + "1";
                }
                var cell = FindOrCreateCell(cellSheetData, cellRef);

                if (properties.TryGetValue("value", out var value))
                {
                    cell.CellValue = new CellValue(value);
                    if (!double.TryParse(value, out _))
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
                if (properties.TryGetValue("formula", out var formula))
                {
                    cell.CellFormula = new CellFormula(formula.TrimStart('='));
                    cell.CellValue = null;
                }
                if (properties.TryGetValue("type", out var cellType))
                {
                    if (cellType.Equals("richtext", StringComparison.OrdinalIgnoreCase) ||
                        cellType.Equals("rich", StringComparison.OrdinalIgnoreCase))
                    {
                        // Build a SharedString rich text entry from run1=text:prop=val, run2=text, etc.
                        var wbPart = _doc.WorkbookPart
                            ?? throw new InvalidOperationException("Workbook not found");
                        var sstPart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                            ?? wbPart.AddNewPart<SharedStringTablePart>();
                        SharedStringTable sst;
                        if (sstPart.SharedStringTable != null)
                            sst = sstPart.SharedStringTable;
                        else
                        {
                            sst = new SharedStringTable();
                            sstPart.SharedStringTable = sst;
                        }

                        var ssi = new SharedStringItem();
                        // Collect run1, run2, ... keys in order
                        var runKeys = properties.Keys
                            .Where(k => k.StartsWith("run", StringComparison.OrdinalIgnoreCase) && k.Length > 3 &&
                                        int.TryParse(k.AsSpan(3), out _))
                            .OrderBy(k => int.Parse(k.AsSpan(3).ToString()))
                            .ToList();
                        foreach (var runKey in runKeys)
                        {
                            var runVal = properties[runKey];
                            // Format: "text:prop=val;prop=val" or just "text"
                            var colonIdx = runVal.IndexOf(':');
                            string runText;
                            string[] runProps;
                            if (colonIdx >= 0)
                            {
                                runText = runVal[..colonIdx];
                                runProps = runVal[(colonIdx + 1)..].Split(';');
                            }
                            else
                            {
                                runText = runVal;
                                runProps = [];
                            }
                            var run = new Run();
                            var rp = new RunProperties();
                            foreach (var prop in runProps)
                            {
                                var eqIdx = prop.IndexOf('=');
                                if (eqIdx < 0) continue;
                                var pKey = prop[..eqIdx].Trim().ToLowerInvariant();
                                var pVal = prop[(eqIdx + 1)..].Trim();
                                switch (pKey)
                                {
                                    case "bold" when ParseHelpers.IsTruthy(pVal): rp.AppendChild(new Bold()); break;
                                    case "italic" when ParseHelpers.IsTruthy(pVal): rp.AppendChild(new Italic()); break;
                                    case "strike" when ParseHelpers.IsTruthy(pVal): rp.AppendChild(new Strike()); break;
                                    case "underline":
                                    {
                                        var ul = new Underline();
                                        if (pVal.Equals("double", StringComparison.OrdinalIgnoreCase)) ul.Val = UnderlineValues.Double;
                                        rp.AppendChild(ul);
                                        break;
                                    }
                                    case "size" or "fontsize":
                                        if (double.TryParse(pVal.TrimEnd('p', 't'), out var sz))
                                            rp.AppendChild(new FontSize { Val = sz });
                                        break;
                                    case "color":
                                        rp.AppendChild(new Color { Rgb = new HexBinaryValue(ParseHelpers.NormalizeArgbColor(pVal)) });
                                        break;
                                    case "font" or "fontname":
                                        rp.AppendChild(new RunFont { Val = pVal });
                                        break;
                                }
                            }
                            if (rp.HasChildren)
                            {
                                ReorderRunProperties(rp);
                                run.AppendChild(rp);
                            }
                            run.AppendChild(new Text(runText) { Space = SpaceProcessingModeValues.Preserve });
                            ssi.AppendChild(run);
                        }

                        if (!ssi.HasChildren)
                        {
                            // No runs defined, fall back to plain text
                            var textVal = cell.CellValue?.Text ?? "";
                            ssi.AppendChild(new Text(textVal) { Space = SpaceProcessingModeValues.Preserve });
                        }

                        sst.AppendChild(ssi);
                        sst.Count = (uint)sst.Elements<SharedStringItem>().Count();
                        sst.UniqueCount = sst.Count;

                        var newIdx = sst.Elements<SharedStringItem>().Count() - 1;
                        cell.CellValue = new CellValue(newIdx.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                    else
                    {
                        cell.DataType = cellType.ToLowerInvariant() switch
                        {
                            "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                            "number" or "num" => null,
                            "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                            _ => throw new ArgumentException($"Invalid cell 'type' value '{cellType}'. Valid types: string, number, boolean, richtext.")
                        };
                        // Convert boolean string values to OOXML-compliant 1/0
                        if (cellType.Equals("boolean", StringComparison.OrdinalIgnoreCase) || cellType.Equals("bool", StringComparison.OrdinalIgnoreCase))
                        {
                            var boolText = cell.CellValue?.Text?.Trim().ToLowerInvariant();
                            if (boolText == "true" || boolText == "yes" || boolText == "1")
                                cell.CellValue = new CellValue("1");
                            else if (boolText == "false" || boolText == "no" || boolText == "0")
                                cell.CellValue = new CellValue("0");
                        }
                    }
                }
                if (properties.TryGetValue("clear", out _))
                {
                    cell.CellValue = null;
                    cell.CellFormula = null;
                }

                // Array formula support during Add
                if (properties.TryGetValue("arrayformula", out var arrFormula))
                {
                    var arrRef = properties.GetValueOrDefault("ref", cellRef);
                    cell.CellFormula = new CellFormula(arrFormula.TrimStart('='))
                    {
                        FormulaType = CellFormulaValues.Array,
                        Reference = arrRef
                    };
                    cell.CellValue = null;
                }

                // Hyperlink support during Add
                if (properties.TryGetValue("link", out var linkUrl) && !string.IsNullOrEmpty(linkUrl))
                {
                    var ws = GetSheet(cellWorksheet);
                    var hlUri = new Uri(linkUrl, UriKind.RelativeOrAbsolute);
                    var hlRel = cellWorksheet.AddHyperlinkRelationship(hlUri, isExternal: true);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    if (hyperlinksEl == null)
                    {
                        hyperlinksEl = new Hyperlinks();
                        // Insert in correct OOXML schema position: after conditionalFormatting, before printOptions/pageMargins/pageSetup/drawing etc.
                        OpenXmlElement? insertBefore = ws.GetFirstChild<PrintOptions>();
                        insertBefore ??= ws.GetFirstChild<PageMargins>();
                        insertBefore ??= ws.GetFirstChild<PageSetup>();
                        insertBefore ??= ws.GetFirstChild<SpreadsheetDrawing>();
                        if (insertBefore != null)
                            ws.InsertBefore(hyperlinksEl, insertBefore);
                        else
                            ws.AppendChild(hyperlinksEl);
                    }
                    hyperlinksEl.AppendChild(new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id });
                }

                // Apply style properties if any
                var cellStyleProps = new Dictionary<string, string>();
                foreach (var (key, val) in properties)
                {
                    if (ExcelStyleManager.IsStyleKey(key))
                        cellStyleProps[key] = val;
                }
                if (cellStyleProps.Count > 0)
                {
                    var cellWbPart = _doc.WorkbookPart
                        ?? throw new InvalidOperationException("Workbook not found");
                    var styleManager = new ExcelStyleManager(cellWbPart);
                    cell.StyleIndex = styleManager.ApplyStyle(cell, cellStyleProps);
                }

                DeleteCalcChainIfPresent();
                SaveWorksheet(cellWorksheet);
                return $"/{cellSheetName}/{cellRef}";

            case "namedrange" or "definedname":
            {
                var nrName = properties.GetValueOrDefault("name", "");
                if (string.IsNullOrEmpty(nrName))
                    throw new ArgumentException("'name' property is required for namedrange");
                var refVal = properties.GetValueOrDefault("ref", "");

                var workbook = GetWorkbook();
                var definedNames = workbook.GetFirstChild<DefinedNames>();
                if (definedNames == null)
                {
                    definedNames = new DefinedNames();
                    // OOXML schema order: ...sheets, functionGroups, externalReferences, definedNames, calcPr, oleSize, customWorkbookViews, pivotCaches...
                    // Insert before calcPr, oleSize, customWorkbookViews, pivotCaches, or any later element
                    var insertBefore = (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<CalculationProperties>()
                        ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.OleSize>()
                        ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CustomWorkbookViews>()
                        ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PivotCaches>();
                    if (insertBefore != null)
                        workbook.InsertBefore(definedNames, insertBefore);
                    else
                        workbook.AppendChild(definedNames);
                }

                var dn = new DefinedName(refVal) { Name = nrName };

                if (properties.TryGetValue("scope", out var scope) && !string.IsNullOrEmpty(scope))
                {
                    var nrSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                    var nrSheetIdx = nrSheets?.FindIndex(s =>
                        s.Name?.Value?.Equals(scope, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
                    if (nrSheetIdx >= 0) dn.LocalSheetId = (uint)nrSheetIdx;
                }
                if (properties.TryGetValue("comment", out var nrComment))
                    dn.Comment = nrComment;

                definedNames.AppendChild(dn);
                workbook.Save();

                var nrIdx = definedNames.Elements<DefinedName>().ToList().IndexOf(dn) + 1;
                return $"/namedrange[{nrIdx}]";
            }

            case "comment" or "note":
            {
                var cmtSegments = parentPath.TrimStart('/').Split('/', 2);
                var cmtSheetName = cmtSegments[0];
                // Extract cell reference from path if present (e.g., /Sheet1/A1 -> A1)
                string? cmtRefFromPath = null;
                if (cmtSegments.Length > 1 && Regex.IsMatch(cmtSegments[1], @"^[A-Z]+\d+$", RegexOptions.IgnoreCase))
                    cmtRefFromPath = cmtSegments[1];
                var cmtWorksheet = FindWorksheet(cmtSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cmtSheetName}");

                var cmtRef = properties.GetValueOrDefault("ref") ?? cmtRefFromPath
                    ?? throw new ArgumentException("Property 'ref' is required for comment");
                var cmtText = properties.GetValueOrDefault("text", "");
                var cmtAuthor = properties.GetValueOrDefault("author", "Author");

                var commentsPart = cmtWorksheet.WorksheetCommentsPart
                    ?? cmtWorksheet.AddNewPart<WorksheetCommentsPart>();

                if (commentsPart.Comments == null)
                {
                    commentsPart.Comments = new Comments(
                        new Authors(new Author(cmtAuthor)),
                        new CommentList()
                    );
                }

                var comments = commentsPart.Comments;
                var authors = comments.GetFirstChild<Authors>()!;
                var commentList = comments.GetFirstChild<CommentList>()!;

                uint authorId = 0;
                var existingAuthors = authors.Elements<Author>().ToList();
                var authorIdx = existingAuthors.FindIndex(a => a.Text == cmtAuthor);
                if (authorIdx >= 0)
                    authorId = (uint)authorIdx;
                else
                {
                    authors.AppendChild(new Author(cmtAuthor));
                    authorId = (uint)existingAuthors.Count;
                }

                var comment = new Comment { Reference = cmtRef.ToUpperInvariant(), AuthorId = authorId };
                comment.CommentText = new CommentText(
                    new Run(
                        new RunProperties(new FontSize { Val = 9 }, new Color { Indexed = 81 },
                            new RunFont { Val = "Tahoma" }),
                        new Text(cmtText) { Space = SpaceProcessingModeValues.Preserve }
                    )
                );
                commentList.AppendChild(comment);
                commentsPart.Comments.Save();

                if (!cmtWorksheet.VmlDrawingParts.Any())
                {
                    var vmlPart = cmtWorksheet.AddNewPart<VmlDrawingPart>();
                    using var writer = new System.IO.StreamWriter(vmlPart.GetStream());
                    writer.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype></xml>");
                }

                var cmtIdx = commentList.Elements<Comment>().ToList().IndexOf(comment) + 1;
                return $"/{cmtSheetName}/comment[{cmtIdx}]";
            }

            case "validation":
            case "datavalidation":
            {
                var dvSegments = parentPath.TrimStart('/').Split('/', 2);
                var dvSheetName = dvSegments[0];
                var dvWorksheet = FindWorksheet(dvSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {dvSheetName}");

                var dvSqref = properties.GetValueOrDefault("sqref")
                    ?? properties.GetValueOrDefault("ref")
                    ?? throw new ArgumentException("Property 'sqref' (or 'ref') is required for validation");

                var dv = new DataValidation
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        dvSqref.Split(' ').Select(s => new StringValue(s)))
                };

                if (properties.TryGetValue("type", out var dvType))
                {
                    dv.Type = dvType.ToLowerInvariant() switch
                    {
                        "list" => DataValidationValues.List,
                        "whole" => DataValidationValues.Whole,
                        "decimal" => DataValidationValues.Decimal,
                        "date" => DataValidationValues.Date,
                        "time" => DataValidationValues.Time,
                        "textlength" => DataValidationValues.TextLength,
                        "custom" => DataValidationValues.Custom,
                        _ => throw new ArgumentException($"Unknown validation type: {dvType}. Use: list, whole, decimal, date, time, textLength, custom")
                    };
                }

                if (properties.TryGetValue("operator", out var dvOp))
                {
                    dv.Operator = dvOp.ToLowerInvariant() switch
                    {
                        "between" => DataValidationOperatorValues.Between,
                        "notbetween" => DataValidationOperatorValues.NotBetween,
                        "equal" => DataValidationOperatorValues.Equal,
                        "notequal" => DataValidationOperatorValues.NotEqual,
                        "greaterthan" => DataValidationOperatorValues.GreaterThan,
                        "lessthan" => DataValidationOperatorValues.LessThan,
                        "greaterthanorequal" => DataValidationOperatorValues.GreaterThanOrEqual,
                        "lessthanorequal" => DataValidationOperatorValues.LessThanOrEqual,
                        _ => throw new ArgumentException($"Unknown operator: {dvOp}")
                    };
                }

                if (properties.TryGetValue("formula1", out var dvFormula1))
                {
                    if (dv.Type?.Value == DataValidationValues.List && !dvFormula1.StartsWith("\""))
                        dv.Formula1 = new Formula1($"\"{dvFormula1}\"");
                    else
                        dv.Formula1 = new Formula1(dvFormula1);
                }

                if (properties.TryGetValue("formula2", out var dvFormula2))
                    dv.Formula2 = new Formula2(dvFormula2);

                // Build case-insensitive lookup for validation properties
                var dvProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);

                dv.AllowBlank = !dvProps.TryGetValue("allowBlank", out var dvAllowBlank)
                    || IsTruthy(dvAllowBlank);
                dv.ShowErrorMessage = !dvProps.TryGetValue("showError", out var dvShowError)
                    || IsTruthy(dvShowError);
                dv.ShowInputMessage = !dvProps.TryGetValue("showInput", out var dvShowInput)
                    || IsTruthy(dvShowInput);

                if (dvProps.TryGetValue("errorTitle", out var dvErrorTitle))
                    dv.ErrorTitle = dvErrorTitle;
                if (dvProps.TryGetValue("error", out var dvError))
                    dv.Error = dvError;
                if (dvProps.TryGetValue("promptTitle", out var dvPromptTitle))
                    dv.PromptTitle = dvPromptTitle;
                if (dvProps.TryGetValue("prompt", out var dvPrompt))
                    dv.Prompt = dvPrompt;

                var wsEl = GetSheet(dvWorksheet);
                var dvs = wsEl.GetFirstChild<DataValidations>();
                if (dvs == null)
                {
                    dvs = new DataValidations();
                    var insertAfter = wsEl.GetFirstChild<Hyperlinks>() as OpenXmlElement
                        ?? wsEl.Elements<ConditionalFormatting>().LastOrDefault() as OpenXmlElement
                        ?? wsEl.GetFirstChild<SheetData>() as OpenXmlElement;
                    if (insertAfter is Hyperlinks)
                        insertAfter.InsertBeforeSelf(dvs);
                    else if (insertAfter != null)
                        insertAfter.InsertAfterSelf(dvs);
                    else
                        wsEl.AppendChild(dvs);
                }

                dvs.AppendChild(dv);
                dvs.Count = (uint)dvs.Elements<DataValidation>().Count();

                SaveWorksheet(dvWorksheet);
                var dvIndex = dvs.Elements<DataValidation>().ToList().IndexOf(dv) + 1;
                return $"/{dvSheetName}/validation[{dvIndex}]";
            }

            case "autofilter":
            {
                var afSegments = parentPath.TrimStart('/').Split('/', 2);
                var afSheetName = afSegments[0];
                var afWorksheet = FindWorksheet(afSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {afSheetName}");

                var afRange = properties.GetValueOrDefault("range")
                    ?? throw new ArgumentException("AutoFilter requires 'range' property (e.g. range=A1:F100)");

                var wsElement = GetSheet(afWorksheet);
                var autoFilter = wsElement.GetFirstChild<AutoFilter>();
                if (autoFilter == null)
                {
                    autoFilter = new AutoFilter();
                    // AutoFilter goes after SheetData (after MergeCells if present)
                    var mergeCellsEl = wsElement.GetFirstChild<MergeCells>();
                    var sheetDataEl = wsElement.GetFirstChild<SheetData>();
                    if (mergeCellsEl != null)
                        mergeCellsEl.InsertAfterSelf(autoFilter);
                    else if (sheetDataEl != null)
                        sheetDataEl.InsertAfterSelf(autoFilter);
                    else
                        wsElement.AppendChild(autoFilter);
                }
                autoFilter.Reference = afRange.ToUpperInvariant();

                SaveWorksheet(afWorksheet);
                return $"/{afSheetName}/autofilter";
            }

            case "cf":
            {
                // Dispatch to specific CF type based on "type" property
                var cfType = properties.GetValueOrDefault("type", "databar").ToLowerInvariant();
                return cfType switch
                {
                    "iconset" => Add(parentPath, "iconset", index, properties),
                    "colorscale" => Add(parentPath, "colorscale", index, properties),
                    "formula" => Add(parentPath, "formulacf", index, properties),
                    "topn" or "top10" => Add(parentPath, "topn", index, properties),
                    "aboveaverage" => Add(parentPath, "aboveaverage", index, properties),
                    "uniquevalues" => Add(parentPath, "uniquevalues", index, properties),
                    "duplicatevalues" => Add(parentPath, "duplicatevalues", index, properties),
                    "containstext" => Add(parentPath, "containstext", index, properties),
                    "dateoccurring" or "timeperiod" => Add(parentPath, "dateoccurring", index, properties),
                    _ => Add(parentPath, "conditionalformatting", index, properties)
                };
            }

            case "databar":
            case "conditionalformatting":
            {
                // Dispatch to specific CF type if "type" property is specified
                if (properties.TryGetValue("type", out var cfTypeVal))
                {
                    var cfTypeLower = cfTypeVal.ToLowerInvariant();
                    if (cfTypeLower is "iconset") return Add(parentPath, "iconset", index, properties);
                    if (cfTypeLower is "colorscale") return Add(parentPath, "colorscale", index, properties);
                    if (cfTypeLower is "formula") return Add(parentPath, "formulacf", index, properties);
                    if (cfTypeLower is "topn" or "top10") return Add(parentPath, "topn", index, properties);
                    if (cfTypeLower is "aboveaverage") return Add(parentPath, "aboveaverage", index, properties);
                    if (cfTypeLower is "uniquevalues") return Add(parentPath, "uniquevalues", index, properties);
                    if (cfTypeLower is "duplicatevalues") return Add(parentPath, "duplicatevalues", index, properties);
                    if (cfTypeLower is "containstext") return Add(parentPath, "containstext", index, properties);
                    if (cfTypeLower is "dateoccurring" or "timeperiod") return Add(parentPath, "dateoccurring", index, properties);
                }
                var cfSegments = parentPath.TrimStart('/').Split('/', 2);
                var cfSheetName = cfSegments[0];
                var cfWorksheet = FindWorksheet(cfSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cfSheetName}");

                var sqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("ref", "A1:A10");
                var minVal = properties.ContainsKey("min") ? properties["min"] : (string?)null;
                var maxVal = properties.ContainsKey("max") ? properties["max"] : (string?)null;
                var cfColor = properties.GetValueOrDefault("color", "638EC6");
                var normalizedColor = ParseHelpers.NormalizeArgbColor(cfColor);

                var cfRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = 1
                };
                var dataBar = new DataBar();
                dataBar.Append(new ConditionalFormatValueObject
                {
                    Type = minVal != null ? ConditionalFormatValueObjectValues.Number : ConditionalFormatValueObjectValues.Min,
                    Val = minVal
                });
                dataBar.Append(new ConditionalFormatValueObject
                {
                    Type = maxVal != null ? ConditionalFormatValueObjectValues.Number : ConditionalFormatValueObjectValues.Max,
                    Val = maxVal
                });
                dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedColor });
                cfRule.Append(dataBar);

                var cf = new ConditionalFormatting(cfRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        sqref.Split(' ').Select(s => new StringValue(s)))
                };

                // Insert after sheetData (or after existing elements)
                var wsElement = GetSheet(cfWorksheet);
                var sheetDataEl = wsElement.GetFirstChild<SheetData>();
                if (sheetDataEl != null)
                    sheetDataEl.InsertAfterSelf(cf);
                else
                    wsElement.Append(cf);

                SaveWorksheet(cfWorksheet);
                var dbCfCount = wsElement.Elements<ConditionalFormatting>().Count();
                return $"/{cfSheetName}/cf[{dbCfCount}]";
            }

            case "colorscale":
            {
                var csSegments = parentPath.TrimStart('/').Split('/', 2);
                var csSheetName = csSegments[0];
                var csWorksheet = FindWorksheet(csSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {csSheetName}");

                var csSqref = properties.GetValueOrDefault("sqref", "A1:A10");
                var minColor = properties.GetValueOrDefault("mincolor", "F8696B");
                var maxColor = properties.GetValueOrDefault("maxcolor", "63BE7B");
                var midColor = properties.GetValueOrDefault("midcolor");

                var normalizedMinColor = ParseHelpers.NormalizeArgbColor(minColor);
                var normalizedMaxColor = ParseHelpers.NormalizeArgbColor(maxColor);

                var colorScale = new ColorScale();
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                if (midColor != null)
                    colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percentile, Val = "50" });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedMinColor });
                if (midColor != null)
                {
                    var normalizedMidColor = ParseHelpers.NormalizeArgbColor(midColor);
                    colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedMidColor });
                }
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedMaxColor });

                var csRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = 1
                };
                csRule.Append(colorScale);

                var csCf = new ConditionalFormatting(csRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        csSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var csWsElement = GetSheet(csWorksheet);
                var csSheetDataEl = csWsElement.GetFirstChild<SheetData>();
                if (csSheetDataEl != null)
                    csSheetDataEl.InsertAfterSelf(csCf);
                else
                    csWsElement.Append(csCf);

                SaveWorksheet(csWorksheet);
                var csCfCount = csWsElement.Elements<ConditionalFormatting>().Count();
                return $"/{csSheetName}/cf[{csCfCount}]";
            }

            case "iconset":
            {
                var isSegments = parentPath.TrimStart('/').Split('/', 2);
                var isSheetName = isSegments[0];
                var isWorksheet = FindWorksheet(isSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {isSheetName}");

                var isSqref = properties.GetValueOrDefault("sqref", "A1:A10");
                var iconSetName = properties.GetValueOrDefault("iconset") ?? properties.GetValueOrDefault("icons", "3TrafficLights1");
                var reverse = properties.TryGetValue("reverse", out var revVal) && IsTruthy(revVal);
                var showValue = !properties.TryGetValue("showvalue", out var svVal) || IsTruthy(svVal);

                var iconSetVal = ParseIconSetValues(iconSetName);

                var iconSet = new IconSet { IconSetValue = iconSetVal };
                if (reverse) iconSet.Reverse = true;
                if (!showValue) iconSet.ShowValue = false;

                // Add threshold values based on icon count
                var iconCount = GetIconCount(iconSetName);
                for (int i = 0; i < iconCount; i++)
                {
                    if (i == 0)
                        iconSet.Append(new ConditionalFormatValueObject
                        {
                            Type = ConditionalFormatValueObjectValues.Percent,
                            Val = "0"
                        });
                    else
                        iconSet.Append(new ConditionalFormatValueObject
                        {
                            Type = ConditionalFormatValueObjectValues.Percent,
                            Val = (i * 100 / iconCount).ToString()
                        });
                }

                var isRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.IconSet,
                    Priority = 1
                };
                isRule.Append(iconSet);

                var isCf = new ConditionalFormatting(isRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        isSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var isWsElement = GetSheet(isWorksheet);
                var isSheetDataEl = isWsElement.GetFirstChild<SheetData>();
                if (isSheetDataEl != null)
                    isSheetDataEl.InsertAfterSelf(isCf);
                else
                    isWsElement.Append(isCf);

                SaveWorksheet(isWorksheet);
                var isCfCount = isWsElement.Elements<ConditionalFormatting>().Count();
                return $"/{isSheetName}/cf[{isCfCount}]";
            }

            case "formulacf":
            {
                var fcfSegments = parentPath.TrimStart('/').Split('/', 2);
                var fcfSheetName = fcfSegments[0];
                var fcfWorksheet = FindWorksheet(fcfSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {fcfSheetName}");

                var fcfSqref = properties.GetValueOrDefault("sqref", "A1:A10");
                var fcfFormula = properties.GetValueOrDefault("formula")
                    ?? throw new ArgumentException("Formula-based conditional formatting requires 'formula' property (e.g. formula=$A1>100)");

                // Build DifferentialFormat (dxf) for the formatting
                var dxf = new DifferentialFormat();
                if (properties.TryGetValue("font.color", out var fontColor))
                {
                    var normalizedFontColor = ParseHelpers.NormalizeArgbColor(fontColor);
                    dxf.Append(new Font(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedFontColor }));
                }
                else if (properties.TryGetValue("font.bold", out var fontBold) && IsTruthy(fontBold))
                {
                    dxf.Append(new Font(new Bold()));
                }

                if (properties.TryGetValue("fill", out var fillColor))
                {
                    var normalizedFillColor = ParseHelpers.NormalizeArgbColor(fillColor);
                    dxf.Append(new Fill(new PatternFill(
                        new BackgroundColor { Rgb = normalizedFillColor })
                    { PatternType = PatternValues.Solid }));
                }

                // Handle font.bold when font.color is also set
                if (properties.TryGetValue("font.color", out _) && properties.TryGetValue("font.bold", out var fb2) && IsTruthy(fb2))
                {
                    var existingFont = dxf.GetFirstChild<Font>();
                    existingFont?.Append(new Bold());
                }

                // Add dxf to stylesheet (ensure it exists)
                var fcfWbPart = _doc.WorkbookPart
                    ?? throw new InvalidOperationException("Workbook not found");
                var fcfStyleMgr = new ExcelStyleManager(fcfWbPart);
                fcfStyleMgr.EnsureStylesPart();
                var stylesheet = fcfWbPart.WorkbookStylesPart!.Stylesheet!;

                var dxfs = stylesheet.GetFirstChild<DifferentialFormats>();
                if (dxfs == null)
                {
                    dxfs = new DifferentialFormats { Count = 0 };
                    stylesheet.Append(dxfs);
                }
                dxfs.Append(dxf);
                dxfs.Count = (uint)dxfs.Elements<DifferentialFormat>().Count();
                stylesheet.Save();

                var dxfId = dxfs.Count!.Value - 1;

                var fcfRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1,
                    FormatId = dxfId
                };
                fcfRule.Append(new Formula(fcfFormula));

                var fcfCf = new ConditionalFormatting(fcfRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        fcfSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var fcfWsElement = GetSheet(fcfWorksheet);
                var fcfSheetDataEl = fcfWsElement.GetFirstChild<SheetData>();
                if (fcfSheetDataEl != null)
                    fcfSheetDataEl.InsertAfterSelf(fcfCf);
                else
                    fcfWsElement.Append(fcfCf);

                SaveWorksheet(fcfWorksheet);
                var fcfCfCount = fcfWsElement.Elements<ConditionalFormatting>().Count();
                return $"/{fcfSheetName}/cf[{fcfCfCount}]";
            }

            case "picture":
            case "image":
            {
                var picSegments = parentPath.TrimStart('/').Split('/', 2);
                var picSheetName = picSegments[0];
                var picWorksheet = FindWorksheet(picSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {picSheetName}");

                var imgPath = properties.GetValueOrDefault("path", "") ?? "";
                if (string.IsNullOrEmpty(imgPath))
                    imgPath = properties.GetValueOrDefault("src", "");
                if (string.IsNullOrEmpty(imgPath))
                    throw new ArgumentException("picture requires a 'path' or 'src' property");

                var (px, py, pw, ph) = ParseAnchorBounds(properties, "0", "0", "5", "5");
                var alt = properties.GetValueOrDefault("alt", "");

                var picDrawingsPart = picWorksheet.DrawingsPart
                    ?? picWorksheet.AddNewPart<DrawingsPart>();

                if (picDrawingsPart.WorksheetDrawing == null)
                {
                    picDrawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
                    picDrawingsPart.WorksheetDrawing.Save();

                    if (GetSheet(picWorksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var drawingRelId = picWorksheet.GetIdOfPart(picDrawingsPart);
                        GetSheet(picWorksheet).Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                        SaveWorksheet(picWorksheet);
                    }
                }

                var (xlImgStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
                using var xlImgDispose = xlImgStream;

                var imgPart = picDrawingsPart.AddImagePart(imgPartType);
                imgPart.FeedData(xlImgStream);
                var imgRelId = picDrawingsPart.GetIdOfPart(imgPart);

                var picId = (uint)(picDrawingsPart.WorksheetDrawing.Elements<XDR.TwoCellAnchor>().Count() + 1);
                var anchor = new XDR.TwoCellAnchor(
                    new XDR.FromMarker(
                        new XDR.ColumnId(px.ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId(py.ToString()),
                        new XDR.RowOffset("0")
                    ),
                    new XDR.ToMarker(
                        new XDR.ColumnId((px + pw).ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId((py + ph).ToString()),
                        new XDR.RowOffset("0")
                    ),
                    new XDR.Picture(
                        new XDR.NonVisualPictureProperties(
                            new XDR.NonVisualDrawingProperties { Id = picId, Name = $"Picture {picId}", Description = alt },
                            new XDR.NonVisualPictureDrawingProperties(new Drawing.PictureLocks { NoChangeAspect = true })
                        ),
                        new XDR.BlipFill(
                            new Drawing.Blip { Embed = imgRelId },
                            new Drawing.Stretch(new Drawing.FillRectangle())
                        ),
                        new XDR.ShapeProperties(
                            new Drawing.Transform2D(
                                new Drawing.Offset { X = 0, Y = 0 },
                                new Drawing.Extents { Cx = 0, Cy = 0 }
                            ),
                            new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                        )
                    ),
                    new XDR.ClientData()
                );

                picDrawingsPart.WorksheetDrawing.AppendChild(anchor);
                picDrawingsPart.WorksheetDrawing.Save();

                var picAnchors = picDrawingsPart.WorksheetDrawing.Elements<XDR.TwoCellAnchor>()
                    .Where(a => a.Descendants<XDR.Picture>().Any()).ToList();
                var picIdx = picAnchors.IndexOf(anchor) + 1;

                return $"/{picSheetName}/picture[{picIdx}]";
            }

            case "shape" or "textbox":
            {
                var shpSegments = parentPath.TrimStart('/').Split('/', 2);
                var shpSheetName = shpSegments[0];
                var shpWorksheet = FindWorksheet(shpSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {shpSheetName}");

                var (sx, sy, sw, sh) = ParseAnchorBounds(properties, "1", "1", "5", "3");
                var shpText = properties.GetValueOrDefault("text", "") ?? "";
                var shpName = properties.GetValueOrDefault("name", "");

                var shpDrawingsPart = shpWorksheet.DrawingsPart
                    ?? shpWorksheet.AddNewPart<DrawingsPart>();

                if (shpDrawingsPart.WorksheetDrawing == null)
                {
                    shpDrawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
                    shpDrawingsPart.WorksheetDrawing.Save();

                    if (GetSheet(shpWorksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var drawingRelId = shpWorksheet.GetIdOfPart(shpDrawingsPart);
                        GetSheet(shpWorksheet).Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                        SaveWorksheet(shpWorksheet);
                    }
                }

                var shpId = (uint)(shpDrawingsPart.WorksheetDrawing.Elements<XDR.TwoCellAnchor>().Count() + 1);
                if (string.IsNullOrEmpty(shpName)) shpName = $"Shape {shpId}";

                // Build ShapeProperties
                var spPr = new XDR.ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = 0, Y = 0 },
                        new Drawing.Extents { Cx = 0, Cy = 0 }
                    ),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                );

                // Fill
                if (properties.TryGetValue("fill", out var shpFill))
                {
                    if (shpFill.Equals("none", StringComparison.OrdinalIgnoreCase))
                        spPr.AppendChild(new Drawing.NoFill());
                    else
                    {
                        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(shpFill);
                        var solidFill = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb });
                        spPr.AppendChild(solidFill);
                    }
                }

                // Line/border
                if (properties.TryGetValue("line", out var shpLine))
                {
                    if (shpLine.Equals("none", StringComparison.OrdinalIgnoreCase))
                        spPr.AppendChild(new Drawing.Outline(new Drawing.NoFill()));
                    else
                    {
                        var (lRgb, _) = ParseHelpers.SanitizeColorForOoxml(shpLine);
                        spPr.AppendChild(new Drawing.Outline(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = lRgb })));
                    }
                }

                // Effects (shadow, glow, reflection, softEdge) — shape-level only for shapes with fill
                // For fill=none shapes, shadow/glow go to text-level (rPr) below
                var isNoFillShape = properties.TryGetValue("fill", out var fillCheck) && fillCheck.Equals("none", StringComparison.OrdinalIgnoreCase);
                Drawing.EffectList? shpEffectList = null;
                if (!isNoFillShape)
                {
                    if (properties.TryGetValue("shadow", out var shpShadow) && !shpShadow.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var normalizedShadow = shpShadow.Replace(':', '-');
                        if (IsTruthy(normalizedShadow)) normalizedShadow = "000000";
                        shpEffectList ??= new Drawing.EffectList();
                        shpEffectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedShadow, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                    }
                    if (properties.TryGetValue("glow", out var shpGlow) && !shpGlow.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var normalizedGlow = shpGlow.Replace(':', '-');
                        if (IsTruthy(normalizedGlow)) normalizedGlow = "4472C4";
                        shpEffectList ??= new Drawing.EffectList();
                        shpEffectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedGlow, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                    }
                }
                if (properties.TryGetValue("reflection", out var shpRefl) && !shpRefl.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    shpEffectList ??= new Drawing.EffectList();
                    shpEffectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildReflection(shpRefl));
                }
                if (properties.TryGetValue("softedge", out var shpSoft) && !shpSoft.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    shpEffectList ??= new Drawing.EffectList();
                    shpEffectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildSoftEdge(shpSoft));
                }
                if (shpEffectList != null)
                    spPr.AppendChild(shpEffectList);

                // Build TextBody with runs
                var bodyPr = new Drawing.BodyProperties { Anchor = Drawing.TextAnchoringTypeValues.Center };
                if (properties.TryGetValue("margin", out var shpMargin))
                {
                    var mEmu = (int)(ParseHelpers.SafeParseDouble(shpMargin, "margin") * 12700);
                    bodyPr.LeftInset = mEmu; bodyPr.RightInset = mEmu;
                    bodyPr.TopInset = mEmu; bodyPr.BottomInset = mEmu;
                }
                var txBody = new XDR.TextBody(bodyPr, new Drawing.ListStyle());

                var lines = shpText.Replace("\\n", "\n").Split('\n');
                foreach (var line in lines)
                {
                    var rPr = new Drawing.RunProperties { Language = "en-US" };

                    // Schema order: attributes → solidFill → effectLst → latin/ea
                    if (properties.TryGetValue("size", out var shpSize))
                        rPr.FontSize = (int)Math.Round(ParseHelpers.SafeParseDouble(shpSize, "size") * 100);
                    if (properties.TryGetValue("bold", out var shpBold) && IsTruthy(shpBold))
                        rPr.Bold = true;
                    if (properties.TryGetValue("italic", out var shpItalic) && IsTruthy(shpItalic))
                        rPr.Italic = true;

                    // Fill (color) before fonts
                    if (properties.TryGetValue("color", out var shpColor))
                    {
                        var (cRgb, _) = ParseHelpers.SanitizeColorForOoxml(shpColor);
                        rPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = cRgb }));
                    }

                    // Text-level effects for fill=none shapes
                    var isNoFill = properties.TryGetValue("fill", out var f) && f.Equals("none", StringComparison.OrdinalIgnoreCase);
                    if (isNoFill)
                    {
                        Drawing.EffectList? txtEffects = null;
                        if (properties.TryGetValue("shadow", out var ts) && !ts.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var normalizedTs = ts.Replace(':', '-');
                            if (IsTruthy(normalizedTs)) normalizedTs = "000000";
                            txtEffects ??= new Drawing.EffectList();
                            txtEffects.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedTs, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                        }
                        if (properties.TryGetValue("glow", out var tg) && !tg.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var normalizedTg = tg.Replace(':', '-');
                            if (IsTruthy(normalizedTg)) normalizedTg = "4472C4";
                            txtEffects ??= new Drawing.EffectList();
                            txtEffects.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedTg, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                        }
                        if (txtEffects != null)
                            rPr.AppendChild(txtEffects);
                    }

                    // Fonts last (schema order)
                    if (properties.TryGetValue("font", out var shpFont))
                    {
                        rPr.AppendChild(new Drawing.LatinFont { Typeface = shpFont });
                        rPr.AppendChild(new Drawing.EastAsianFont { Typeface = shpFont });
                    }

                    var pPr = new Drawing.ParagraphProperties();
                    if (properties.TryGetValue("align", out var shpAlign))
                    {
                        pPr.Alignment = shpAlign.ToLowerInvariant() switch
                        {
                            "center" or "c" or "ctr" => Drawing.TextAlignmentTypeValues.Center,
                            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
                            _ => Drawing.TextAlignmentTypeValues.Left
                        };
                    }

                    txBody.AppendChild(new Drawing.Paragraph(
                        pPr,
                        new Drawing.Run(rPr, new Drawing.Text(line))
                    ));
                }

                var shape = new XDR.Shape(
                    new XDR.NonVisualShapeProperties(
                        new XDR.NonVisualDrawingProperties { Id = shpId, Name = shpName },
                        new XDR.NonVisualShapeDrawingProperties()
                    ),
                    spPr,
                    txBody
                );

                var shpAnchor = new XDR.TwoCellAnchor(
                    new XDR.FromMarker(
                        new XDR.ColumnId(sx.ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId(sy.ToString()),
                        new XDR.RowOffset("0")
                    ),
                    new XDR.ToMarker(
                        new XDR.ColumnId((sx + sw).ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId((sy + sh).ToString()),
                        new XDR.RowOffset("0")
                    ),
                    shape,
                    new XDR.ClientData()
                );

                shpDrawingsPart.WorksheetDrawing.AppendChild(shpAnchor);
                shpDrawingsPart.WorksheetDrawing.Save();

                var shpAnchors = shpDrawingsPart.WorksheetDrawing.Elements<XDR.TwoCellAnchor>()
                    .Where(a => a.Descendants<XDR.Shape>().Any()).ToList();
                var shpIdx = shpAnchors.IndexOf(shpAnchor) + 1;

                return $"/{shpSheetName}/shape[{shpIdx}]";
            }

            case "table" or "listobject":
            {
                var tblSegments = parentPath.TrimStart('/').Split('/', 2);
                var tblSheetName = tblSegments[0];
                var tblWorksheet = FindWorksheet(tblSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {tblSheetName}");

                var rangeRef = (properties.GetValueOrDefault("ref") ?? properties.GetValueOrDefault("range")
                    ?? throw new ArgumentException("Property 'ref' or 'range' is required for table")).ToUpperInvariant();

                var existingTableIds = _doc.WorkbookPart!.WorksheetParts
                    .SelectMany(wp => wp.TableDefinitionParts)
                    .Select(tdp => tdp.Table?.Id?.Value ?? 0);
                var tableId = existingTableIds.Any() ? existingTableIds.Max() + 1 : 1;

                var tableName = properties.GetValueOrDefault("name", $"Table{tableId}");
                var displayName = properties.GetValueOrDefault("displayName", tableName);
                var styleName = properties.GetValueOrDefault("style", "TableStyleMedium2");
                var hasHeader = !properties.TryGetValue("headerRow", out var hrVal) || IsTruthy(hrVal);
                var hasTotalRow = properties.TryGetValue("totalRow", out var trVal) && IsTruthy(trVal);

                var rangeParts = rangeRef.Split(':');
                var (startCol, startRow) = ParseCellReference(rangeParts[0]);
                var (endCol, endRow) = ParseCellReference(rangeParts[1]);
                var startColIdx = ColumnNameToIndex(startCol);
                var endColIdx = ColumnNameToIndex(endCol);
                var colCount = endColIdx - startColIdx + 1;

                string[] colNames;
                if (properties.TryGetValue("columns", out var tblColsStr))
                {
                    var userColNames = tblColsStr.Split(',').Select(c => c.Trim()).ToArray();
                    // Pad with default names if fewer columns provided than range requires
                    colNames = new string[colCount];
                    for (int i = 0; i < colCount; i++)
                        colNames[i] = i < userColNames.Length ? userColNames[i] : $"Column{i + 1}";
                }
                else
                {
                    colNames = new string[colCount];
                    if (hasHeader)
                    {
                        var tblSheetData = GetSheet(tblWorksheet).GetFirstChild<SheetData>();
                        var headerRow = tblSheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)startRow);
                        for (int i = 0; i < colCount; i++)
                        {
                            var colLetter = IndexToColumnName(startColIdx + i);
                            var cellRefStr = $"{colLetter}{startRow}";
                            var headerCell = headerRow?.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRefStr);
                            colNames[i] = (headerCell != null ? GetCellDisplayValue(headerCell) : null) ?? $"Column{i + 1}";
                            if (string.IsNullOrEmpty(colNames[i]))
                                colNames[i] = $"Column{i + 1}";
                        }
                    }
                    else
                    {
                        for (int i = 0; i < colCount; i++)
                            colNames[i] = $"Column{i + 1}";
                    }
                }

                var tableDefPart = tblWorksheet.AddNewPart<TableDefinitionPart>();
                var table = new Table
                {
                    Id = (uint)tableId,
                    Name = tableName,
                    DisplayName = displayName,
                    Reference = rangeRef,
                    TotalsRowShown = hasTotalRow
                };
                if (hasTotalRow)
                    table.TotalsRowCount = 1;

                table.AppendChild(new AutoFilter { Reference = rangeRef });

                var tableColumns = new TableColumns { Count = (uint)colCount };
                for (int i = 0; i < colCount; i++)
                    tableColumns.AppendChild(new TableColumn { Id = (uint)(i + 1), Name = colNames[i] });
                table.AppendChild(tableColumns);

                table.AppendChild(new TableStyleInfo
                {
                    Name = styleName,
                    ShowFirstColumn = false,
                    ShowLastColumn = false,
                    ShowRowStripes = true,
                    ShowColumnStripes = false
                });

                // Generate total row content in SheetData when totalRow is enabled
                if (hasTotalRow)
                {
                    var tblSheetData = GetSheet(tblWorksheet).GetFirstChild<SheetData>()
                        ?? GetSheet(tblWorksheet).AppendChild(new SheetData());
                    var totalRowIdx = (uint)endRow;
                    var totalRow = tblSheetData.Elements<Row>()
                        .FirstOrDefault(r => r.RowIndex?.Value == totalRowIdx);
                    if (totalRow == null)
                    {
                        totalRow = new Row { RowIndex = totalRowIdx };
                        // Insert in correct position
                        var lastRow = tblSheetData.Elements<Row>()
                            .Where(r => r.RowIndex?.Value < totalRowIdx)
                            .LastOrDefault();
                        if (lastRow != null)
                            lastRow.InsertAfterSelf(totalRow);
                        else
                            tblSheetData.AppendChild(totalRow);
                    }

                    var tblCols = tableColumns.Elements<TableColumn>().ToList();
                    for (int ci = 0; ci < tblCols.Count; ci++)
                    {
                        var colLetter = IndexToColumnName(startColIdx + ci);
                        var cellRefStr = $"{colLetter}{totalRowIdx}";
                        var existingCell = totalRow.Elements<Cell>()
                            .FirstOrDefault(c => c.CellReference?.Value == cellRefStr);
                        if (existingCell == null)
                        {
                            existingCell = new Cell { CellReference = cellRefStr };
                            totalRow.AppendChild(existingCell);
                        }

                        if (ci == 0)
                        {
                            // First column: label "Total"
                            tblCols[ci].TotalsRowLabel = "Total";
                            existingCell.CellValue = new CellValue("Total");
                            existingCell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        else
                        {
                            // Other columns: SUBTOTAL(109, range) formula for SUM
                            tblCols[ci].TotalsRowFunction = TotalsRowFunctionValues.Sum;
                            var dataStartRow = hasHeader ? startRow + 1 : startRow;
                            var dataEndRow = (int)totalRowIdx - 1;
                            var formulaRange = $"{colLetter}{dataStartRow}:{colLetter}{dataEndRow}";
                            existingCell.CellFormula = new CellFormula($"SUBTOTAL(109,{formulaRange})");
                        }
                    }
                }

                tableDefPart.Table = table;
                tableDefPart.Table.Save();

                var tblWs = GetSheet(tblWorksheet);
                var tableParts = tblWs.GetFirstChild<TableParts>();
                if (tableParts == null)
                {
                    tableParts = new TableParts();
                    tblWs.AppendChild(tableParts);
                }
                tableParts.AppendChild(new TablePart { Id = tblWorksheet.GetIdOfPart(tableDefPart) });
                tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();
                tblWs.Save();

                var tblIdx = tblWorksheet.TableDefinitionParts.ToList().IndexOf(tableDefPart) + 1;
                return $"/{tblSheetName}/table[{tblIdx}]";
            }

            case "chart":
            {
                var chartSegments = parentPath.TrimStart('/').Split('/', 2);
                var chartSheetName = chartSegments[0];
                var chartWorksheet = FindWorksheet(chartSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {chartSheetName}");

                // Parse chart data
                var chartType = properties.FirstOrDefault(kv =>
                    kv.Key.Equals("charttype", StringComparison.OrdinalIgnoreCase)
                    || kv.Key.Equals("type", StringComparison.OrdinalIgnoreCase)).Value
                    ?? "column";
                var chartTitle = properties.GetValueOrDefault("title");
                var categories = ChartHelper.ParseCategories(properties);
                var seriesData = ChartHelper.ParseSeriesData(properties);

                if (seriesData.Count == 0)
                    throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                        "or series1=\"Revenue:100,200,300\"");

                // Create DrawingsPart if needed
                var drawingsPart = chartWorksheet.DrawingsPart
                    ?? chartWorksheet.AddNewPart<DrawingsPart>();

                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
                    drawingsPart.WorksheetDrawing.Save();

                    if (GetSheet(chartWorksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var drawingRelId = chartWorksheet.GetIdOfPart(drawingsPart);
                        GetSheet(chartWorksheet).Append(
                            new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                        SaveWorksheet(chartWorksheet);
                    }
                }

                // Build chart content BEFORE adding part (invalid type throws, must not leave empty part)
                var chartSpace = ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = chartSpace;
                chartPart.ChartSpace.Save();

                // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
                var deferredProps = properties
                    .Where(kv => ChartHelper.DeferredAddKeys.Contains(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (deferredProps.Count > 0)
                    ChartHelper.SetChartProperties(chartPart, deferredProps);

                // Position via TwoCellAnchor
                var fromCol = properties.TryGetValue("x", out var xStr) ? ParseHelpers.SafeParseInt(xStr, "x") : 0;
                var fromRow = properties.TryGetValue("y", out var yStr) ? ParseHelpers.SafeParseInt(yStr, "y") : 0;
                var toCol = properties.TryGetValue("width", out var wStr) ? fromCol + ParseHelpers.SafeParseInt(wStr, "width") : fromCol + 8;
                var toRow = properties.TryGetValue("height", out var hStr) ? fromRow + ParseHelpers.SafeParseInt(hStr, "height") : fromRow + 15;

                var anchor = new XDR.TwoCellAnchor();
                anchor.Append(new XDR.FromMarker(
                    new XDR.ColumnId(fromCol.ToString()),
                    new XDR.ColumnOffset("0"),
                    new XDR.RowId(fromRow.ToString()),
                    new XDR.RowOffset("0")));
                anchor.Append(new XDR.ToMarker(
                    new XDR.ColumnId(toCol.ToString()),
                    new XDR.ColumnOffset("0"),
                    new XDR.RowId(toRow.ToString()),
                    new XDR.RowOffset("0")));

                var chartRelId = drawingsPart.GetIdOfPart(chartPart);
                var graphicFrame = new XDR.GraphicFrame();
                graphicFrame.NonVisualGraphicFrameProperties = new XDR.NonVisualGraphicFrameProperties(
                    new XDR.NonVisualDrawingProperties
                    {
                        Id = (uint)(drawingsPart.WorksheetDrawing.ChildElements.Count + 2),
                        Name = chartTitle ?? "Chart"
                    },
                    new XDR.NonVisualGraphicFrameDrawingProperties()
                );
                graphicFrame.Transform = new XDR.Transform(
                    new Drawing.Offset { X = 0, Y = 0 },
                    new Drawing.Extents { Cx = 0, Cy = 0 }
                );

                var chartRef = new C.ChartReference { Id = chartRelId };
                graphicFrame.Append(new Drawing.Graphic(
                    new Drawing.GraphicData(chartRef)
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                    }
                ));

                anchor.Append(graphicFrame);
                anchor.Append(new XDR.ClientData());
                drawingsPart.WorksheetDrawing.Append(anchor);
                drawingsPart.WorksheetDrawing.Save();

                // Legend
                var legendVal = properties.GetValueOrDefault("legend", "true");
                // Legend is already handled inside ExcelChartBuildChartSpace

                var chartIdx = drawingsPart.ChartParts.ToList().IndexOf(chartPart) + 1;
                return $"/{chartSheetName}/chart[{chartIdx}]";
            }

            case "pivottable" or "pivot":
            {
                var ptSegments = parentPath.TrimStart('/').Split('/', 2);
                var ptSheetName = ptSegments[0];
                var ptWorksheet = FindWorksheet(ptSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {ptSheetName}");

                // Source: "Sheet1!A1:D100" or "A1:D100" (same sheet)
                var sourceSpec = properties.GetValueOrDefault("source", "")
                    ?? properties.GetValueOrDefault("src", "")
                    ?? throw new ArgumentException("pivottable requires 'source' property (e.g. source=Sheet1!A1:D100)");
                if (string.IsNullOrEmpty(sourceSpec))
                    throw new ArgumentException("pivottable requires 'source' property (e.g. source=Sheet1!A1:D100)");

                string sourceSheetName;
                string sourceRef;
                if (sourceSpec.Contains('!'))
                {
                    var srcParts = sourceSpec.Split('!', 2);
                    sourceSheetName = srcParts[0].Trim('\'', '"');
                    sourceRef = srcParts[1];
                }
                else
                {
                    sourceSheetName = ptSheetName;
                    sourceRef = sourceSpec;
                }

                var sourceWorksheet = FindWorksheet(sourceSheetName)
                    ?? throw new ArgumentException($"Source sheet not found: {sourceSheetName}");

                var position = properties.GetValueOrDefault("position", "")
                    ?? properties.GetValueOrDefault("pos", "");
                if (string.IsNullOrEmpty(position))
                {
                    // Auto-position: place after the source data range
                    var rangeEnd = sourceRef.Split(':').Last();
                    var colEndMatch = System.Text.RegularExpressions.Regex.Match(rangeEnd, @"([A-Za-z]+)");
                    var nextCol = colEndMatch.Success ? IndexToColumnName(ColumnNameToIndex(colEndMatch.Value.ToUpperInvariant()) + 2) : "H";
                    position = $"{nextCol}1";
                }

                var ptIdx = PivotTableHelper.CreatePivotTable(
                    _doc.WorkbookPart!, ptWorksheet, sourceWorksheet,
                    sourceSheetName, sourceRef, position, properties);

                return $"/{ptSheetName}/pivottable[{ptIdx}]";
            }

            case "col" or "column":
            {
                var colSegments = parentPath.TrimStart('/').Split('/', 2);
                var colSheetName = colSegments[0];
                var colWorksheet = FindWorksheet(colSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {colSheetName}");

                // Determine insert column: index (1-based) or name from properties
                string insertColName;
                if (properties.TryGetValue("name", out var colNameProp) && !string.IsNullOrEmpty(colNameProp))
                {
                    insertColName = colNameProp.ToUpperInvariant();
                }
                else if (index.HasValue)
                {
                    insertColName = IndexToColumnName(index.Value);
                }
                else
                {
                    // Append after last used column
                    var ws = GetSheet(colWorksheet);
                    var sheetDataForCol = ws.GetFirstChild<SheetData>();
                    int maxColIdx = 0;
                    if (sheetDataForCol != null)
                    {
                        foreach (var r in sheetDataForCol.Elements<Row>())
                            foreach (var cx in r.Elements<Cell>())
                            {
                                if (cx.CellReference?.Value == null) continue;
                                var (c, _) = ParseCellReference(cx.CellReference.Value);
                                var ci = ColumnNameToIndex(c);
                                if (ci > maxColIdx) maxColIdx = ci;
                            }
                    }
                    insertColName = IndexToColumnName(maxColIdx + 1);
                }

                var insertColIdx = ColumnNameToIndex(insertColName);

                // Shift existing columns right if needed
                var colSheetData = GetSheet(colWorksheet).GetFirstChild<SheetData>();
                bool colNeedsShift = colSheetData != null && colSheetData.Elements<Row>()
                    .Any(r => r.Elements<Cell>().Any(c =>
                    {
                        if (c.CellReference?.Value == null) return false;
                        var (col, _) = ParseCellReference(c.CellReference.Value);
                        return ColumnNameToIndex(col) >= insertColIdx;
                    }));
                if (colNeedsShift)
                    ShiftColumnsRight(colWorksheet, insertColIdx);

                // Optionally set column width
                if (properties.TryGetValue("width", out var widthStr) && double.TryParse(widthStr, out var width))
                {
                    var ws = GetSheet(colWorksheet);
                    var columns = ws.GetFirstChild<Columns>() ?? ws.PrependChild(new Columns());
                    columns.AppendChild(new Column
                    {
                        Min = (uint)insertColIdx,
                        Max = (uint)insertColIdx,
                        Width = width,
                        CustomWidth = true
                    });
                }

                SaveWorksheet(colWorksheet);
                return $"/{colSheetName}/col[{insertColName}]";
            }

            case "pagebreak":
            {
                // Route to rowbreak or colbreak based on properties
                if (properties.ContainsKey("col") || properties.ContainsKey("column"))
                    return Add(parentPath, "colbreak", index, properties);
                return Add(parentPath, "rowbreak", index, properties);
            }

            case "rowbreak":
            {
                var rbSegments = parentPath.TrimStart('/').Split('/', 2);
                var rbSheetName = rbSegments[0];
                var rbWorksheet = FindWorksheet(rbSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {rbSheetName}");
                var rbWs = GetSheet(rbWorksheet);

                var rbRowIdx = uint.Parse(properties.GetValueOrDefault("row") ?? properties.GetValueOrDefault("index")
                    ?? throw new ArgumentException("'row' property is required for rowbreak"));

                var rowBreaks = rbWs.GetFirstChild<RowBreaks>();
                if (rowBreaks == null)
                {
                    rowBreaks = new RowBreaks();
                    rbWs.AppendChild(rowBreaks);
                }
                rowBreaks.AppendChild(new Break
                {
                    Id = rbRowIdx,
                    Max = 16383u,
                    ManualPageBreak = true
                });
                rowBreaks.Count = (uint)rowBreaks.Elements<Break>().Count();
                rowBreaks.ManualBreakCount = rowBreaks.Count;
                SaveWorksheet(rbWorksheet);

                var rbIdx = rowBreaks.Elements<Break>().ToList()
                    .FindIndex(b => b.Id?.Value == rbRowIdx) + 1;
                return $"/{rbSheetName}/rowbreak[{rbIdx}]";
            }

            case "colbreak":
            {
                var cbSegments = parentPath.TrimStart('/').Split('/', 2);
                var cbSheetName = cbSegments[0];
                var cbWorksheet = FindWorksheet(cbSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cbSheetName}");
                var cbWs = GetSheet(cbWorksheet);

                var cbColStr = properties.GetValueOrDefault("col") ?? properties.GetValueOrDefault("column")
                    ?? properties.GetValueOrDefault("index")
                    ?? throw new ArgumentException("'col' property is required for colbreak");
                // Accept both numeric index (e.g. "3") and column letter (e.g. "C")
                var cbColIdx = uint.TryParse(cbColStr, out var cbNumVal)
                    ? cbNumVal
                    : (uint)ColumnNameToIndex(cbColStr.ToUpperInvariant());

                var colBreaks = cbWs.GetFirstChild<ColumnBreaks>();
                if (colBreaks == null)
                {
                    colBreaks = new ColumnBreaks();
                    cbWs.AppendChild(colBreaks);
                }
                colBreaks.AppendChild(new Break
                {
                    Id = cbColIdx,
                    Max = 1048575u,
                    ManualPageBreak = true
                });
                colBreaks.Count = (uint)colBreaks.Elements<Break>().Count();
                colBreaks.ManualBreakCount = colBreaks.Count;
                SaveWorksheet(cbWorksheet);

                var cbBrkIdx = colBreaks.Elements<Break>().ToList()
                    .FindIndex(b => b.Id?.Value == cbColIdx) + 1;
                return $"/{cbSheetName}/colbreak[{cbBrkIdx}]";
            }

            case "run":
            {
                // Add a rich text run to a cell: parentPath = /SheetName/CellRef
                var runSegments = parentPath.TrimStart('/').Split('/', 2);
                if (runSegments.Length < 2)
                    throw new ArgumentException("Parent path must be /SheetName/CellRef for adding a run");
                var runSheetName = runSegments[0];
                var runCellRef = runSegments[1].ToUpperInvariant();
                var runWorksheet = FindWorksheet(runSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {runSheetName}");
                var runSheetData = GetSheet(runWorksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(runWorksheet).AppendChild(new SheetData());
                var runCell = FindOrCreateCell(runSheetData, runCellRef);

                var runWbPart = _doc.WorkbookPart
                    ?? throw new InvalidOperationException("Workbook not found");
                var runSstPart = runWbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                    ?? runWbPart.AddNewPart<SharedStringTablePart>();
                SharedStringTable runSst;
                if (runSstPart.SharedStringTable != null)
                    runSst = runSstPart.SharedStringTable;
                else
                {
                    runSst = new SharedStringTable();
                    runSstPart.SharedStringTable = runSst;
                }

                SharedStringItem? runSsi = null;
                if (runCell.DataType?.Value == CellValues.SharedString &&
                    int.TryParse(runCell.CellValue?.Text, out var existingSstIdx))
                {
                    runSsi = runSst.Elements<SharedStringItem>().ElementAtOrDefault(existingSstIdx);
                }
                if (runSsi == null)
                {
                    runSsi = new SharedStringItem();
                    runSst.AppendChild(runSsi);
                    var newSstIdx = runSst.Elements<SharedStringItem>().Count() - 1;
                    runCell.CellValue = new CellValue(newSstIdx.ToString());
                    runCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }

                var newRun = new Run();
                var newRunProps = new RunProperties();
                var runText = properties.GetValueOrDefault("text", "");

                foreach (var (rKey, rVal) in properties)
                {
                    switch (rKey.ToLowerInvariant())
                    {
                        case "bold" when ParseHelpers.IsTruthy(rVal):
                            newRunProps.AppendChild(new Bold()); break;
                        case "italic" when ParseHelpers.IsTruthy(rVal):
                            newRunProps.AppendChild(new Italic()); break;
                        case "strike" when ParseHelpers.IsTruthy(rVal):
                            newRunProps.AppendChild(new Strike()); break;
                        case "underline":
                            if (!string.IsNullOrEmpty(rVal) && rVal != "false" && rVal != "none")
                            {
                                var ul = new Underline();
                                if (rVal.ToLowerInvariant() == "double") ul.Val = UnderlineValues.Double;
                                newRunProps.AppendChild(ul);
                            }
                            break;
                        case "superscript" when ParseHelpers.IsTruthy(rVal):
                            newRunProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript }); break;
                        case "subscript" when ParseHelpers.IsTruthy(rVal):
                            newRunProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript }); break;
                        case "size" or "fontsize":
                            if (double.TryParse(rVal.TrimEnd('p', 't'), out var runSz))
                                newRunProps.AppendChild(new FontSize { Val = runSz });
                            break;
                        case "color":
                            newRunProps.AppendChild(new Color { Rgb = new HexBinaryValue(ParseHelpers.NormalizeArgbColor(rVal)) });
                            break;
                        case "font" or "fontname":
                            newRunProps.AppendChild(new RunFont { Val = rVal }); break;
                    }
                }
                if (newRunProps.HasChildren)
                {
                    ReorderRunProperties(newRunProps);
                    newRun.AppendChild(newRunProps);
                }
                newRun.AppendChild(new Text(runText) { Space = SpaceProcessingModeValues.Preserve });
                runSsi.AppendChild(newRun);

                runSst.Count = (uint)runSst.Elements<SharedStringItem>().Count();
                runSst.UniqueCount = runSst.Count;

                SaveWorksheet(runWorksheet);
                var runIndex = runSsi.Elements<Run>().Count();
                return $"/{runSheetName}/{runCellRef}/run[{runIndex}]";
            }

            case "topn":
            case "aboveaverage":
            case "uniquevalues":
            case "duplicatevalues":
            case "containstext":
            case "dateoccurring":
            {
                var cfNewSegments = parentPath.TrimStart('/').Split('/', 2);
                var cfNewSheetName = cfNewSegments[0];
                var cfNewWorksheet = FindWorksheet(cfNewSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cfNewSheetName}");
                var cfNewSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("ref", "A1:A10");

                ConditionalFormattingRule cfNewRule;
                var typeLower = type.ToLowerInvariant();

                switch (typeLower)
                {
                    case "topn":
                    {
                        var rank = uint.TryParse(properties.GetValueOrDefault("rank", "10"), out var r) ? r : 10u;
                        var percent = ParseHelpers.IsTruthy(properties.GetValueOrDefault("percent", "false"));
                        var bottom = ParseHelpers.IsTruthy(properties.GetValueOrDefault("bottom", "false"));
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.Top10,
                            Priority = 1,
                            Rank = rank,
                            Percent = percent ? true : null,
                            Bottom = bottom ? true : null
                        };
                        break;
                    }
                    case "aboveaverage":
                    {
                        var aboveBelow = properties.GetValueOrDefault("above", "true");
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.AboveAverage,
                            Priority = 1,
                            AboveAverage = ParseHelpers.IsTruthy(aboveBelow) ? null : false
                        };
                        break;
                    }
                    case "uniquevalues":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.UniqueValues,
                            Priority = 1
                        };
                        break;
                    }
                    case "duplicatevalues":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.DuplicateValues,
                            Priority = 1
                        };
                        break;
                    }
                    case "containstext":
                    {
                        var text = properties.GetValueOrDefault("text", "");
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.ContainsText,
                            Priority = 1,
                            Text = text,
                            Operator = ConditionalFormattingOperatorValues.ContainsText
                        };
                        var firstCell = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"NOT(ISERROR(SEARCH(\"{text}\",{firstCell})))"));
                        break;
                    }
                    case "dateoccurring":
                    {
                        var period = properties.GetValueOrDefault("period", "today");
                        var normalizedPeriod = period.ToLowerInvariant() switch
                        {
                            "today" => "today",
                            "yesterday" => "yesterday",
                            "tomorrow" => "tomorrow",
                            "last7days" => "last7Days",
                            "thisweek" => "thisWeek",
                            "lastweek" => "lastWeek",
                            "nextweek" => "nextWeek",
                            "thismonth" => "thisMonth",
                            "lastmonth" => "lastMonth",
                            "nextmonth" => "nextMonth",
                            _ => period
                        };
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.TimePeriod,
                            Priority = 1,
                            TimePeriod = new EnumValue<TimePeriodValues>(normalizedPeriod switch
                            {
                                "today" => TimePeriodValues.Today,
                                "yesterday" => TimePeriodValues.Yesterday,
                                "tomorrow" => TimePeriodValues.Tomorrow,
                                "last7Days" => TimePeriodValues.Last7Days,
                                "thisWeek" => TimePeriodValues.ThisWeek,
                                "lastWeek" => TimePeriodValues.LastWeek,
                                "nextWeek" => TimePeriodValues.NextWeek,
                                "thisMonth" => TimePeriodValues.ThisMonth,
                                "lastMonth" => TimePeriodValues.LastMonth,
                                "nextMonth" => TimePeriodValues.NextMonth,
                                _ => TimePeriodValues.Today
                            })
                        };
                        break;
                    }
                    default:
                        throw new ArgumentException($"Unsupported CF type: {typeLower}");
                }

                // Build DXF formatting if fill/font properties are provided
                var cfNewDxf = new DifferentialFormat();
                bool cfNewHasDxf = false;
                if (properties.TryGetValue("font.color", out var cfNewFontColor))
                {
                    var normalizedFontColor = ParseHelpers.NormalizeArgbColor(cfNewFontColor);
                    cfNewDxf.Append(new Font(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedFontColor }));
                    cfNewHasDxf = true;
                }
                else if (properties.TryGetValue("font.bold", out var cfNewFontBold) && IsTruthy(cfNewFontBold))
                {
                    cfNewDxf.Append(new Font(new Bold()));
                    cfNewHasDxf = true;
                }
                if (properties.TryGetValue("fill", out var cfNewFillColor))
                {
                    var normalizedFillColor = ParseHelpers.NormalizeArgbColor(cfNewFillColor);
                    cfNewDxf.Append(new Fill(new PatternFill(
                        new BackgroundColor { Rgb = normalizedFillColor })
                    { PatternType = PatternValues.Solid }));
                    cfNewHasDxf = true;
                }
                if (properties.TryGetValue("font.color", out _) && properties.TryGetValue("font.bold", out var cfNewFb2) && IsTruthy(cfNewFb2))
                {
                    var existingFont = cfNewDxf.GetFirstChild<Font>();
                    existingFont?.Append(new Bold());
                }

                if (cfNewHasDxf)
                {
                    var cfNewWbPart = _doc.WorkbookPart
                        ?? throw new InvalidOperationException("Workbook not found");
                    var cfNewStyleMgr = new ExcelStyleManager(cfNewWbPart);
                    cfNewStyleMgr.EnsureStylesPart();
                    var cfNewStylesheet = cfNewWbPart.WorkbookStylesPart!.Stylesheet!;
                    var cfNewDxfs = cfNewStylesheet.GetFirstChild<DifferentialFormats>();
                    if (cfNewDxfs == null)
                    {
                        cfNewDxfs = new DifferentialFormats { Count = 0 };
                        cfNewStylesheet.Append(cfNewDxfs);
                    }
                    cfNewDxfs.Append(cfNewDxf);
                    cfNewDxfs.Count = (uint)cfNewDxfs.Elements<DifferentialFormat>().Count();
                    cfNewStylesheet.Save();
                    cfNewRule.FormatId = cfNewDxfs.Count!.Value - 1;
                }

                var cfNewFormatting = new ConditionalFormatting(cfNewRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        cfNewSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var cfNewWs = GetSheet(cfNewWorksheet);
                var cfNewSheetData = cfNewWs.GetFirstChild<SheetData>();
                if (cfNewSheetData != null)
                    cfNewSheetData.InsertAfterSelf(cfNewFormatting);
                else
                    cfNewWs.AppendChild(cfNewFormatting);

                SaveWorksheet(cfNewWorksheet);
                var cfNewCount = cfNewWs.Elements<ConditionalFormatting>().Count();
                return $"/{cfNewSheetName}/cf[{cfNewCount}]";
            }

            case "sparkline":
            {
                var spkSegments = parentPath.TrimStart('/').Split('/', 2);
                var spkSheetName = spkSegments[0];
                var spkWorksheet = FindWorksheet(spkSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {spkSheetName}");

                var spkCell = properties.GetValueOrDefault("cell")
                    ?? throw new ArgumentException("Sparkline requires 'cell' property (e.g. F1)");
                var spkRange = properties.GetValueOrDefault("range")
                    ?? throw new ArgumentException("Sparkline requires 'range' property (e.g. A1:E1)");

                // Determine sparkline type
                var spkTypeStr = properties.GetValueOrDefault("type", "line").ToLowerInvariant();
                var spkType = spkTypeStr switch
                {
                    "column" => X14.SparklineTypeValues.Column,
                    "stacked" => X14.SparklineTypeValues.Stacked,
                    _ => X14.SparklineTypeValues.Line
                };

                // Build the SparklineGroup
                var spkGroup = new X14.SparklineGroup();
                // Only set Type attribute for non-line (line is default in OOXML)
                if (spkType != X14.SparklineTypeValues.Line)
                    spkGroup.Type = spkType;

                // Series color
                var spkColor = properties.GetValueOrDefault("color", "4472C4");
                spkGroup.SeriesColor = new X14.SeriesColor { Rgb = ParseHelpers.NormalizeArgbColor(spkColor) };

                // Negative color
                if (properties.TryGetValue("negativecolor", out var negColor))
                    spkGroup.NegativeColor = new X14.NegativeColor { Rgb = ParseHelpers.NormalizeArgbColor(negColor) };

                // Boolean flags
                if (properties.TryGetValue("markers", out var markersVal) && ParseHelpers.IsTruthy(markersVal))
                    spkGroup.Markers = true;
                if (properties.TryGetValue("highpoint", out var highVal) && ParseHelpers.IsTruthy(highVal))
                    spkGroup.High = true;
                if (properties.TryGetValue("lowpoint", out var lowVal) && ParseHelpers.IsTruthy(lowVal))
                    spkGroup.Low = true;
                if (properties.TryGetValue("firstpoint", out var firstVal) && ParseHelpers.IsTruthy(firstVal))
                    spkGroup.First = true;
                if (properties.TryGetValue("lastpoint", out var lastVal) && ParseHelpers.IsTruthy(lastVal))
                    spkGroup.Last = true;
                if (properties.TryGetValue("negative", out var negVal) && ParseHelpers.IsTruthy(negVal))
                    spkGroup.Negative = true;

                // Marker colors
                if (properties.TryGetValue("highmarkercolor", out var highMC))
                    spkGroup.HighMarkerColor = new X14.HighMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(highMC) };
                if (properties.TryGetValue("lowmarkercolor", out var lowMC))
                    spkGroup.LowMarkerColor = new X14.LowMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(lowMC) };
                if (properties.TryGetValue("firstmarkercolor", out var firstMC))
                    spkGroup.FirstMarkerColor = new X14.FirstMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(firstMC) };
                if (properties.TryGetValue("lastmarkercolor", out var lastMC))
                    spkGroup.LastMarkerColor = new X14.LastMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(lastMC) };
                if (properties.TryGetValue("markerscolor", out var markersMC))
                    spkGroup.MarkersColor = new X14.MarkersColor { Rgb = ParseHelpers.NormalizeArgbColor(markersMC) };

                // Line weight
                if (properties.TryGetValue("lineweight", out var lwVal) && double.TryParse(lwVal, out var lw))
                    spkGroup.LineWeight = lw;

                // Build the Sparkline element
                // Ensure range includes sheet reference
                var spkFormulaRef = spkRange.Contains('!') ? spkRange : $"{spkSheetName}!{spkRange}";
                var sparkline = new X14.Sparkline
                {
                    Formula = new DocumentFormat.OpenXml.Office.Excel.Formula(spkFormulaRef),
                    ReferenceSequence = new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(spkCell)
                };
                var sparklines = new X14.Sparklines();
                sparklines.Append(sparkline);
                spkGroup.Append(sparklines);

                // Add to worksheet extension list
                var spkWs = GetSheet(spkWorksheet);
                var spkExtList = spkWs.GetFirstChild<WorksheetExtensionList>()
                    ?? spkWs.AppendChild(new WorksheetExtensionList());

                // Find existing sparkline extension or create new one
                var spkExt = spkExtList.Elements<WorksheetExtension>()
                    .FirstOrDefault(e => e.Uri == "{05C60535-1F16-4fd2-B633-E4A46CF9E463}");
                X14.SparklineGroups spkGroups;
                if (spkExt != null)
                {
                    spkGroups = spkExt.GetFirstChild<X14.SparklineGroups>()
                        ?? spkExt.AppendChild(new X14.SparklineGroups());
                }
                else
                {
                    spkExt = new WorksheetExtension { Uri = "{05C60535-1F16-4fd2-B633-E4A46CF9E463}" };
                    spkExt.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
                    spkGroups = new X14.SparklineGroups();
                    spkExt.Append(spkGroups);
                    spkExtList.Append(spkExt);
                }

                spkGroups.Append(spkGroup);
                SaveWorksheet(spkWorksheet);

                // Count all sparkline groups to determine index
                var allSpkGroups = spkGroups.Elements<X14.SparklineGroup>().ToList();
                var spkIdx = allSpkGroups.IndexOf(spkGroup) + 1;
                return $"/{spkSheetName}/sparkline[{spkIdx}]";
            }

            default:
            {
                // Generic fallback: create typed element via SDK schema validation
                // Parse parentPath: /<SheetName>/xmlPath...
                var fbSegments = parentPath.TrimStart('/').Split('/', 2);
                var fbSheetName = fbSegments[0];
                var fbWorksheet = FindWorksheet(fbSheetName);
                if (fbWorksheet == null)
                    throw new ArgumentException($"Sheet not found: {fbSheetName}");

                OpenXmlElement fbParent = GetSheet(fbWorksheet);
                if (fbSegments.Length > 1 && !string.IsNullOrEmpty(fbSegments[1]))
                {
                    var xmlSegments = GenericXmlQuery.ParsePathSegments(fbSegments[1]);
                    fbParent = GenericXmlQuery.NavigateByPath(fbParent!, xmlSegments)
                        ?? throw new ArgumentException($"Parent element not found: {parentPath}");
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent!, type, properties, index);
                if (created == null)
                    throw new ArgumentException(
                        $"Unknown element type '{type}' for {parentPath}. " +
                        "Valid types: sheet, row, cell, shape, chart, autofilter, databar, colorscale, iconset, formulacf, comment, namedrange, table, picture, validation, pivottable. " +
                        "Use 'officecli xlsx add' for details.");

                SaveWorksheet(fbWorksheet);

                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }


    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
            throw new ArgumentException("Cannot move an entire sheet. Use move on rows or elements within a sheet.");

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Determine target
        string effectiveParentPath;
        SheetData targetSheetData;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            effectiveParentPath = $"/{sheetName}";
            targetSheetData = sheetData;
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
            var tgtWorksheet = FindWorksheet(tgtSegments[0])
                ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
            targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
                ?? throw new ArgumentException("Target sheet has no data");
        }

        // Find and move the row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = int.Parse(rowMatch.Groups[1].Value);
            // Try ordinal lookup first (Nth row element), then fall back to RowIndex
            var allRows = sheetData.Elements<Row>().ToList();
            var row = (rowIdx >= 1 && rowIdx <= allRows.Count ? allRows[rowIdx - 1] : null)
                ?? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            row.Remove();

            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                    rows[index.Value].InsertBeforeSelf(row);
                else
                    targetSheetData.AppendChild(row);
            }
            else
            {
                targetSheetData.AppendChild(row);
            }

            SaveWorksheet(worksheet);
            var rowIndex = row.RowIndex?.Value ?? (uint)(targetSheetData.Elements<Row>().ToList().IndexOf(row) + 1);
            return $"{effectiveParentPath}/row[{rowIndex}]";
        }

        throw new ArgumentException($"Move not supported for: {elementRef}. Supported: row[N]");
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        // Parse both paths: /SheetName/row[N]
        var seg1 = path1.TrimStart('/').Split('/', 2);
        var seg2 = path2.TrimStart('/').Split('/', 2);
        if (seg1.Length < 2 || seg2.Length < 2)
            throw new ArgumentException("Swap requires element paths (e.g. /Sheet1/row[1])");
        if (seg1[0] != seg2[0])
            throw new ArgumentException("Cannot swap elements across different sheets");

        var sheetName = seg1[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        var rowMatch1 = Regex.Match(seg1[1], @"^row\[(\d+)\]$");
        var rowMatch2 = Regex.Match(seg2[1], @"^row\[(\d+)\]$");
        if (!rowMatch1.Success || !rowMatch2.Success)
            throw new ArgumentException("Swap only supports row[N] elements in Excel");

        var allRows = sheetData.Elements<Row>().ToList();
        var idx1 = int.Parse(rowMatch1.Groups[1].Value);
        var idx2 = int.Parse(rowMatch2.Groups[1].Value);
        var row1 = (idx1 >= 1 && idx1 <= allRows.Count ? allRows[idx1 - 1] : null)
            ?? throw new ArgumentException($"Row {idx1} not found");
        var row2 = (idx2 >= 1 && idx2 <= allRows.Count ? allRows[idx2 - 1] : null)
            ?? throw new ArgumentException($"Row {idx2} not found");

        // Swap RowIndex values and cell references
        var rowIndex1 = row1.RowIndex?.Value ?? (uint)idx1;
        var rowIndex2 = row2.RowIndex?.Value ?? (uint)idx2;
        row1.RowIndex = new DocumentFormat.OpenXml.UInt32Value(rowIndex2);
        row2.RowIndex = new DocumentFormat.OpenXml.UInt32Value(rowIndex1);

        // Update cell references (e.g. A1→A3, B1→B3)
        foreach (var cell in row1.Elements<Cell>())
        {
            if (cell.CellReference?.Value != null)
            {
                var colRef = Regex.Match(cell.CellReference.Value, @"^([A-Z]+)").Groups[1].Value;
                cell.CellReference = $"{colRef}{rowIndex2}";
            }
        }
        foreach (var cell in row2.Elements<Cell>())
        {
            if (cell.CellReference?.Value != null)
            {
                var colRef = Regex.Match(cell.CellReference.Value, @"^([A-Z]+)").Groups[1].Value;
                cell.CellReference = $"{colRef}{rowIndex1}";
            }
        }

        PowerPointHandler.SwapXmlElements(row1, row2);
        SaveWorksheet(worksheet);

        return ($"/{sheetName}/row[{rowIndex2}]", $"/{sheetName}/row[{rowIndex1}]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
            throw new ArgumentException("Cannot copy an entire sheet with --from. Use add --type sheet instead.");

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Find target
        var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
        var tgtWorksheet = FindWorksheet(tgtSegments[0])
            ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
        var targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Target sheet has no data");

        // Copy row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            var clone = (Row)row.CloneNode(true);

            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                    rows[index.Value].InsertBeforeSelf(clone);
                else
                    targetSheetData.AppendChild(clone);
            }
            else
            {
                targetSheetData.AppendChild(clone);
            }

            SaveWorksheet(tgtWorksheet);
            var newRows = targetSheetData.Elements<Row>().ToList();
            var newIdx = newRows.IndexOf(clone) + 1;
            return $"{targetParentPath}/row[{newIdx}]";
        }

        throw new ArgumentException($"Copy not supported for: {elementRef}. Supported: row[N]");
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a worksheet's DrawingsPart
                var sheetName = parentPartPath.TrimStart('/');
                var worksheetPart = FindWorksheet(sheetName)
                    ?? throw new ArgumentException(
                        $"Sheet not found: {sheetName}. Chart must be added under a sheet: add-part <file> /<SheetName> --type chart");

                var drawingsPart = worksheetPart.DrawingsPart
                    ?? worksheetPart.AddNewPart<DrawingsPart>();

                // Initialize DrawingsPart if new
                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing =
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                    drawingsPart.WorksheetDrawing.Save();

                    // Link DrawingsPart to worksheet if not already linked
                    if (GetSheet(worksheetPart).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
                        GetSheet(worksheetPart).Append(
                            new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                        SaveWorksheet(worksheetPart);
                    }
                }

                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                var relId = drawingsPart.GetIdOfPart(chartPart);

                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = drawingsPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/{sheetName}/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }
}
