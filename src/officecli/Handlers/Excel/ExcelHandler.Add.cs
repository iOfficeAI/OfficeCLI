// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
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
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
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
                var newRow = new Row { RowIndex = (uint)rowIdx };

                // Create cells if cols specified
                if (properties.TryGetValue("cols", out var colsStr))
                {
                    if (!int.TryParse(colsStr, out var cols))
                        throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Expected a positive integer (number of columns to create).");
                    for (int c = 0; c < cols; c++)
                    {
                        var colLetter = IndexToColumnName(c + 1);
                        newRow.AppendChild(new Cell { CellReference = $"{colLetter}{rowIdx}" });
                    }
                }

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

                var cellRef = properties.GetValueOrDefault("ref", "A1");
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
                    cell.DataType = cellType.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        _ => throw new ArgumentException($"Invalid cell 'type' value '{cellType}'. Valid types: string, number, boolean.")
                    };
                }
                if (properties.TryGetValue("clear", out _))
                {
                    cell.CellValue = null;
                    cell.CellFormula = null;
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

                SaveWorksheet(cellWorksheet);
                return $"/{cellSheetName}/{cellRef}";

            case "namedrange" or "definedname":
            {
                var nrName = properties.GetValueOrDefault("name", "");
                if (string.IsNullOrEmpty(nrName))
                    throw new ArgumentException("'name' property is required for namedrange");
                var refVal = properties.GetValueOrDefault("ref", "");

                var workbook = GetWorkbook();
                var definedNames = workbook.GetFirstChild<DefinedNames>()
                    ?? workbook.AppendChild(new DefinedNames());

                var dn = new DefinedName(refVal) { Name = nrName };

                if (properties.TryGetValue("scope", out var scope) && !string.IsNullOrEmpty(scope))
                {
                    var nrSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                    var nrSheetIdx = nrSheets?.FindIndex(s =>
                        s.Name?.Value?.Equals(scope, StringComparison.OrdinalIgnoreCase) == true);
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
                }
                var cfSegments = parentPath.TrimStart('/').Split('/', 2);
                var cfSheetName = cfSegments[0];
                var cfWorksheet = FindWorksheet(cfSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cfSheetName}");

                var sqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("ref", "A1:A10");
                var minVal = properties.GetValueOrDefault("min", "0");
                var maxVal = properties.GetValueOrDefault("max", "1");
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
                    Type = ConditionalFormatValueObjectValues.Number,
                    Val = minVal
                });
                dataBar.Append(new ConditionalFormatValueObject
                {
                    Type = ConditionalFormatValueObjectValues.Number,
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
                if (string.IsNullOrEmpty(imgPath) || !File.Exists(imgPath))
                    throw new ArgumentException("picture requires a valid 'path' or 'src' property");

                var px = int.TryParse(properties.GetValueOrDefault("x", "0"), out var xv) ? xv : 0;
                var py = int.TryParse(properties.GetValueOrDefault("y", "0"), out var yv) ? yv : 0;
                var pw = int.TryParse(properties.GetValueOrDefault("width", "5"), out var wv) ? wv : 5;
                var ph = int.TryParse(properties.GetValueOrDefault("height", "5"), out var hv) ? hv : 5;
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

                var ext = Path.GetExtension(imgPath).ToLowerInvariant();
                var imgPartType = ext switch
                {
                    ".png" => ImagePartType.Png,
                    ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                    ".gif" => ImagePartType.Gif,
                    ".bmp" => ImagePartType.Bmp,
                    ".tiff" or ".tif" => ImagePartType.Tiff,
                    ".emf" => ImagePartType.Emf,
                    ".wmf" => ImagePartType.Wmf,
                    _ => throw new ArgumentException($"Unsupported image format: {ext}")
                };

                var imgPart = picDrawingsPart.AddImagePart(imgPartType);
                using (var stream = File.OpenRead(imgPath))
                    imgPart.FeedData(stream);
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
                var (endCol, _) = ParseCellReference(rangeParts[1]);
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
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                SaveWorksheet(fbWorksheet);

                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }

    public void Remove(string path)
    {
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        if (segments.Length == 1)
        {
            // Remove entire sheet
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var sheets = GetWorkbook().GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetName}");

            var sheetCount = sheets!.Elements<Sheet>().Count();
            if (sheetCount <= 1)
                throw new InvalidOperationException($"Cannot remove the last sheet. A workbook must contain at least one sheet.");

            var relId = sheet.Id?.Value;
            sheet.Remove();
            if (relId != null)
                workbookPart.DeletePart(workbookPart.GetPartById(relId));

            // Clean up named ranges referencing the deleted sheet
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                var toRemove = definedNames.Elements<DefinedName>()
                    .Where(dn => dn.Text?.Contains(sheetName + "!", StringComparison.OrdinalIgnoreCase) == true)
                    .ToList();
                foreach (var dn in toRemove) dn.Remove();
                if (!definedNames.HasChildren) definedNames.Remove();
            }

            // Fix ActiveTab to prevent workbook corruption when deleting the last tab
            var remainingCount = sheets!.Elements<Sheet>().Count();
            var bookViews = workbook.GetFirstChild<BookViews>();
            if (bookViews != null)
            {
                foreach (var bv in bookViews.Elements<WorkbookView>())
                {
                    if (bv.ActiveTab?.Value >= (uint)remainingCount)
                        bv.ActiveTab = (uint)Math.Max(0, remainingCount - 1);
                }
            }

            workbook.Save();
            return;
        }

        // Remove cell or row
        var cellRef = segments[1];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Check if it's a row reference like row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
            row?.Remove();

            // Clean up merge cell references that include this row
            var ws = GetSheet(worksheet);
            var mergeCells = ws.GetFirstChild<MergeCells>();
            if (mergeCells != null)
            {
                var toRemove = new List<MergeCell>();
                foreach (var mc in mergeCells.Elements<MergeCell>())
                {
                    var mergeRef = mc.Reference?.Value;
                    if (string.IsNullOrEmpty(mergeRef)) continue;
                    var parts = mergeRef.Split(':');
                    if (parts.Length != 2) continue;
                    var (_, startRow) = ParseCellReference(parts[0]);
                    var (_, endRow) = ParseCellReference(parts[1]);
                    if (rowIdx >= (uint)startRow && rowIdx <= (uint)endRow)
                        toRemove.Add(mc);
                }
                foreach (var mc in toRemove) mc.Remove();
                if (!mergeCells.HasChildren) mergeCells.Remove();
            }

            if (row == null)
            {
                // Row didn't exist as data, but we cleaned up merge references
                SaveWorksheet(worksheet);
                return;
            }
        }
        else
        {
            // Cell reference
            var cell = FindCell(sheetData, cellRef)
                ?? throw new ArgumentException($"Cell {cellRef} not found");
            cell.Remove();
        }

        SaveWorksheet(worksheet);
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
