// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        // Normalize to case-insensitive lookup so camelCase keys match lowercase lookups
        if (properties != null && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        path = NormalizeExcelPath(path);

        // Handle root path "/" — document properties
        if (path == "/")
        {
            var unsupported = new List<string>();
            var pkg = _doc.PackageProperties;
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "title": pkg.Title = value; break;
                    case "author" or "creator": pkg.Creator = value; break;
                    case "subject": pkg.Subject = value; break;
                    case "description": pkg.Description = value; break;
                    case "category": pkg.Category = value; break;
                    case "keywords": pkg.Keywords = value; break;
                    case "lastmodifiedby": pkg.LastModifiedBy = value; break;
                    case "revision": pkg.Revision = value; break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            return unsupported;
        }

        // Handle /SheetName/sparkline[N]
        var sparklineSetMatch = Regex.Match(path.TrimStart('/'), @"^([^/]+)/sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (sparklineSetMatch.Success)
        {
            var spkSheet = sparklineSetMatch.Groups[1].Value;
            var spkIdx = int.Parse(sparklineSetMatch.Groups[2].Value);
            var spkWorksheet = FindWorksheet(spkSheet) ?? throw SheetNotFoundException(spkSheet);
            var spkGroup = GetSparklineGroup(spkWorksheet, spkIdx)
                ?? throw new ArgumentException($"Sparkline[{spkIdx}] not found in sheet '{spkSheet}'");

            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "type":
                        spkGroup.Type = value.ToLowerInvariant() switch
                        {
                            "column" => X14.SparklineTypeValues.Column,
                            "stacked" => X14.SparklineTypeValues.Stacked,
                            _ => null // null = line (default, no attribute)
                        };
                        break;
                    case "color":
                        spkGroup.SeriesColor = new X14.SeriesColor { Rgb = ParseHelpers.NormalizeArgbColor(value) };
                        break;
                    case "negativecolor":
                        spkGroup.NegativeColor = new X14.NegativeColor { Rgb = ParseHelpers.NormalizeArgbColor(value) };
                        break;
                    case "markers":
                        spkGroup.Markers = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                        break;
                    case "highpoint":
                        spkGroup.High = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                        break;
                    case "lowpoint":
                        spkGroup.Low = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                        break;
                    case "firstpoint":
                        spkGroup.First = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                        break;
                    case "lastpoint":
                        spkGroup.Last = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                        break;
                    case "negative":
                        spkGroup.Negative = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                        break;
                    case "lineweight":
                        if (double.TryParse(value, out var lw)) spkGroup.LineWeight = lw;
                        break;
                    default:
                        unsup.Add(key);
                        break;
                }
            }
            SaveWorksheet(spkWorksheet);
            return unsup;
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
                    case "scope":
                        // Set scope like POI's XSSFName.setSheetIndex
                        if (string.IsNullOrEmpty(value) || value.Equals("workbook", StringComparison.OrdinalIgnoreCase))
                        {
                            dn.LocalSheetId = null; // workbook-global scope
                        }
                        else
                        {
                            var nrSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                            var nrSheetIdx = nrSheets?.FindIndex(s =>
                                s.Name?.Value?.Equals(value, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
                            if (nrSheetIdx >= 0)
                                dn.LocalSheetId = (uint)nrSheetIdx;
                            else
                                throw new ArgumentException($"Sheet '{value}' not found for scope");
                        }
                        break;
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
            throw SheetNotFoundException(sheetName);

        // Sheet-level Set (path is just /SheetName)
        if (segments.Length < 2)
        {
            return SetSheetLevel(worksheet, sheetName, properties);
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
                            _ => throw new ArgumentException($"Unknown validation type: '{value}'. Valid types: list, whole, decimal, date, time, textLength, custom.")
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
                    case "operator":
                        dv.Operator = value.ToLowerInvariant() switch
                        {
                            "between" => DataValidationOperatorValues.Between,
                            "notbetween" => DataValidationOperatorValues.NotBetween,
                            "equal" => DataValidationOperatorValues.Equal,
                            "notequal" => DataValidationOperatorValues.NotEqual,
                            "lessthan" => DataValidationOperatorValues.LessThan,
                            "lessthanorequal" => DataValidationOperatorValues.LessThanOrEqual,
                            "greaterthan" => DataValidationOperatorValues.GreaterThan,
                            "greaterthanorequal" => DataValidationOperatorValues.GreaterThanOrEqual,
                            _ => throw new ArgumentException($"Unknown operator: {value}")
                        };
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
                var lk = key.ToLowerInvariant();
                if (TrySetAnchorPosition(anchor, lk, value)) continue;

                var spPr = anchor.Descendants<XDR.ShapeProperties>().FirstOrDefault();
                if (TrySetRotation(spPr, lk, value)) continue;
                if (TrySetShapeEffect(spPr, lk, value)) continue;

                switch (lk)
                {
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

        // Handle /SheetName/shape[N]
        var shapeSetMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (shapeSetMatch.Success)
        {
            var shpIdx = int.Parse(shapeSetMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/shapes");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/shapes");

            var shpAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
                .Where(a => a.Descendants<XDR.Shape>().Any()).ToList();
            if (shpIdx < 1 || shpIdx > shpAnchors.Count)
                throw new ArgumentException($"Shape index {shpIdx} out of range (1..{shpAnchors.Count})");

            var anchor = shpAnchors[shpIdx - 1];
            var shape = anchor.Descendants<XDR.Shape>().First();
            var shpUnsupported = new List<string>();

            foreach (var (key, value) in properties)
            {
                var lk = key.ToLowerInvariant();
                if (TrySetAnchorPosition(anchor, lk, value)) continue;
                if (TrySetRotation(shape.ShapeProperties, lk, value)) continue;

                // For effects on shapes: check if fill=none → text-level, otherwise shape-level
                if (lk is "shadow" or "glow" or "reflection" or "softedge")
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) continue;
                    var isNoFill = spPr.GetFirstChild<Drawing.NoFill>() != null;
                    var normalizedVal = value.Replace(':', '-');

                    if (isNoFill && lk is "shadow" or "glow")
                    {
                        foreach (var run in shape.Descendants<Drawing.Run>())
                        {
                            if (lk == "shadow")
                                OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.OuterShadow>(run, normalizedVal, () =>
                                    OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                            else
                                OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.Glow>(run, normalizedVal, () =>
                                    OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                        }
                    }
                    else
                    {
                        TrySetShapeEffect(spPr, lk, value);
                    }
                    continue;
                }

                switch (lk)
                {
                    case "name":
                    {
                        var nvProps = shape.NonVisualShapeProperties?.GetFirstChild<XDR.NonVisualDrawingProperties>();
                        if (nvProps != null) nvProps.Name = value;
                        break;
                    }
                    case "text":
                    {
                        var txBody = shape.TextBody;
                        if (txBody != null)
                        {
                            var firstPara = txBody.Elements<Drawing.Paragraph>().FirstOrDefault();
                            var pProps = firstPara?.ParagraphProperties?.CloneNode(true);
                            var rProps = firstPara?.Elements<Drawing.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                            txBody.RemoveAllChildren<Drawing.Paragraph>();
                            var lines = value.Replace("\\n", "\n").Split('\n');
                            foreach (var line in lines)
                            {
                                var para = new Drawing.Paragraph();
                                if (pProps != null) para.AppendChild(pProps.CloneNode(true));
                                var run = new Drawing.Run(new Drawing.Text(line));
                                if (rProps != null) run.RunProperties = (Drawing.RunProperties)rProps.CloneNode(true);
                                para.AppendChild(run);
                                txBody.AppendChild(para);
                            }
                        }
                        break;
                    }
                    case "font":
                        foreach (var run in shape.Descendants<Drawing.Run>())
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rPr.RemoveAllChildren<Drawing.LatinFont>();
                            rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                            rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                            rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                        }
                        break;
                    case "size":
                        foreach (var run in shape.Descendants<Drawing.Run>())
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rPr.FontSize = (int)Math.Round(ParseHelpers.ParseFontSize(value) * 100);
                        }
                        break;
                    case "bold":
                        foreach (var run in shape.Descendants<Drawing.Run>())
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rPr.Bold = IsTruthy(value);
                        }
                        break;
                    case "italic":
                        foreach (var run in shape.Descendants<Drawing.Run>())
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rPr.Italic = IsTruthy(value);
                        }
                        break;
                    case "color":
                        foreach (var run in shape.Descendants<Drawing.Run>())
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rPr.RemoveAllChildren<Drawing.SolidFill>();
                            var (cRgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                            OfficeCli.Core.DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                                new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = cRgb }));
                        }
                        break;
                    case "fill":
                    {
                        var spPr = shape.ShapeProperties;
                        if (spPr != null)
                        {
                            spPr.RemoveAllChildren<Drawing.SolidFill>();
                            spPr.RemoveAllChildren<Drawing.NoFill>();
                            if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                                spPr.AppendChild(new Drawing.NoFill());
                            else
                            {
                                var (fRgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                                spPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = fRgb }));
                            }
                        }
                        break;
                    }
                    case "align":
                        foreach (var para in shape.Descendants<Drawing.Paragraph>())
                        {
                            var pPr = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                            pPr.Alignment = value.ToLowerInvariant() switch
                            {
                                "center" or "c" or "ctr" => Drawing.TextAlignmentTypeValues.Center,
                                "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
                                "justify" or "justified" or "j" => Drawing.TextAlignmentTypeValues.Justified,
                                _ => Drawing.TextAlignmentTypeValues.Left
                            };
                        }
                        break;
                    default:
                        shpUnsupported.Add(key);
                        break;
                }
            }

            drawingsPart.WorksheetDrawing.Save();
            return shpUnsupported;
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
                    case "headerrow": table.HeaderRowCount = IsTruthy(value) ? 1u : 0u; break;
                    case "totalrow": table.TotalsRowShown = IsTruthy(value); break;
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
                    case "ref":
                        // Update cell reference (like POI's XSSFComment.setAddress)
                        cmtElement.Reference = value.ToUpperInvariant();
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

        // Handle /SheetName/cf[N] or /SheetName/conditionalformatting[N]
        var cfSetMatch = Regex.Match(cellRef, @"^(?:cf|conditionalformatting)\[(\d+)\]$", RegexOptions.IgnoreCase);
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
                        if (dbColor != null) { dbColor.Rgb = ParseHelpers.NormalizeArgbColor(value); }
                        else unsup.Add(key);
                        break;
                    case "mincolor":
                        var csColors = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                        if (csColors != null && csColors.Count >= 2)
                        { csColors[0].Rgb = ParseHelpers.NormalizeArgbColor(value); }
                        else unsup.Add(key);
                        break;
                    case "maxcolor":
                        var csColors2 = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                        if (csColors2 != null && csColors2.Count >= 2)
                        { csColors2[^1].Rgb = ParseHelpers.NormalizeArgbColor(value); }
                        else unsup.Add(key);
                        break;
                    case "iconset":
                    case "icons":
                        var iconSetEl = rule?.GetFirstChild<IconSet>();
                        if (iconSetEl != null)
                            iconSetEl.IconSetValue = new EnumValue<IconSetValues>(ParseIconSetValues(value));
                        else unsup.Add(key);
                        break;
                    case "reverse":
                        var isEl = rule?.GetFirstChild<IconSet>();
                        if (isEl != null) isEl.Reverse = IsTruthy(value);
                        else unsup.Add(key);
                        break;
                    case "showvalue":
                        var isEl2 = rule?.GetFirstChild<IconSet>();
                        if (isEl2 != null) isEl2.ShowValue = IsTruthy(value);
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

        // Handle /SheetName/col[X] where X is a column letter (A) or numeric index (1)
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z0-9]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colValue = colMatch.Groups[1].Value;
            var colName = int.TryParse(colValue, out var colNumIdx) ? IndexToColumnName(colNumIdx) : colValue.ToUpperInvariant();
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

            var unsup = ChartHelper.SetChartProperties(chartPart, properties);
            chartPart.ChartSpace?.Save();
            return unsup;
        }

        // Handle /SheetName/pivottable[N]
        var pivotSetMatch = Regex.Match(cellRef, @"^pivottable\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (pivotSetMatch.Success)
        {
            var ptIdx = int.Parse(pivotSetMatch.Groups[1].Value);
            var pivotParts = worksheet.PivotTableParts.ToList();
            if (ptIdx < 1 || ptIdx > pivotParts.Count)
                throw new ArgumentException($"PivotTable {ptIdx} not found");
            return PivotTableHelper.SetPivotTableProperties(pivotParts[ptIdx - 1], properties);
        }

        // Handle /SheetName/A1/run[N] (rich text run)
        var runSetMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runSetMatch.Success)
        {
            var runCellRef = runSetMatch.Groups[1].Value.ToUpperInvariant();
            var runIdx = int.Parse(runSetMatch.Groups[2].Value);

            var runSheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
            if (runSheetData == null) throw new ArgumentException("Sheet data not found");
            var runCell = FindOrCreateCell(runSheetData, runCellRef);

            if (runCell.DataType?.Value != CellValues.SharedString ||
                !int.TryParse(runCell.CellValue?.Text, out var sstIdx))
                throw new ArgumentException($"Cell {runCellRef} is not a rich text cell");

            var sstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi = sstPart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx);
            if (ssi == null) throw new ArgumentException($"SharedString entry {sstIdx} not found");

            var runs = ssi.Elements<Run>().ToList();
            if (runIdx < 1 || runIdx > runs.Count)
                throw new ArgumentException($"Run index {runIdx} out of range (1-{runs.Count})");

            var run = runs[runIdx - 1];
            var rProps = run.RunProperties ?? run.PrependChild(new RunProperties());

            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text" or "value":
                        var textEl = run.GetFirstChild<Text>();
                        if (textEl != null) textEl.Text = value;
                        else run.AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                        break;
                    case "bold":
                        rProps.RemoveAllChildren<Bold>();
                        if (ParseHelpers.IsTruthy(value)) rProps.InsertAt(new Bold(), 0);
                        break;
                    case "italic":
                        rProps.RemoveAllChildren<Italic>();
                        if (ParseHelpers.IsTruthy(value)) rProps.AppendChild(new Italic());
                        break;
                    case "strike":
                        rProps.RemoveAllChildren<Strike>();
                        if (ParseHelpers.IsTruthy(value)) rProps.AppendChild(new Strike());
                        break;
                    case "underline":
                        rProps.RemoveAllChildren<Underline>();
                        if (!string.IsNullOrEmpty(value) && value != "false" && value != "none")
                        {
                            var ul = new Underline();
                            if (value.ToLowerInvariant() == "double") ul.Val = UnderlineValues.Double;
                            rProps.AppendChild(ul);
                        }
                        break;
                    case "superscript":
                        rProps.RemoveAllChildren<VerticalTextAlignment>();
                        if (ParseHelpers.IsTruthy(value))
                            rProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript });
                        break;
                    case "subscript":
                        rProps.RemoveAllChildren<VerticalTextAlignment>();
                        if (ParseHelpers.IsTruthy(value))
                            rProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript });
                        break;
                    case "size":
                        rProps.RemoveAllChildren<FontSize>();
                        rProps.AppendChild(new FontSize { Val = ParseHelpers.ParseFontSize(value) });
                        break;
                    case "color":
                        rProps.RemoveAllChildren<Color>();
                        rProps.AppendChild(new Color { Rgb = ParseHelpers.NormalizeArgbColor(value) });
                        break;
                    case "font":
                        rProps.RemoveAllChildren<RunFont>();
                        rProps.AppendChild(new RunFont { Val = value });
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }

            ReorderRunProperties(rProps);
            sstPart!.SharedStringTable!.Save();
            SaveWorksheet(worksheet);
            return unsupported;
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

        // Clone cell for rollback on failure (atomic: no partial modifications)
        var cellBackup = cell.CloneNode(true);

        try
        {
        return SetCellProperties(cell, cellRef, worksheet, properties);
        }
        catch
        {
            // Rollback: restore cell to pre-modification state
            cell.Parent?.ReplaceChild(cellBackup, cell);
            throw;
        }
    }

    private List<string> SetCellProperties(Cell cell, string cellRef, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var unsupported = ApplyCellProperties(cell, cellRef, worksheet, properties);
        // Any mutation to a cell (value, formula, clear) can invalidate the calc chain
        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return unsupported;
    }

    /// <summary>Apply cell properties without saving — caller is responsible for SaveWorksheet.</summary>
    private List<string> ApplyCellProperties(Cell cell, WorksheetPart worksheet, Dictionary<string, string> properties)
        => ApplyCellProperties(cell, cell.CellReference?.Value ?? "", worksheet, properties);

    private List<string> ApplyCellProperties(Cell cell, string cellRef, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
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
                    cell.CellFormula = null; // Clear formula when explicit value is set
                    // If cell is already boolean type, convert true/false to 1/0
                    if (cell.DataType?.Value == CellValues.Boolean)
                    {
                        var bv = value.Trim().ToLowerInvariant();
                        if (bv is "true" or "yes") cell.CellValue = new CellValue("1");
                        else if (bv is "false" or "no") cell.CellValue = new CellValue("0");
                        else cell.CellValue = new CellValue(value);
                    }
                    else
                    {
                        cell.CellValue = new CellValue(value);
                        // Auto-detect type: number or string (boolean only via explicit type=boolean)
                        if (double.TryParse(value, out _))
                            cell.DataType = null; // Number is default
                        else
                        {
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                    }
                    break;
                case "formula":
                    cell.CellFormula = new CellFormula(value.TrimStart('='));
                    cell.CellValue = null;
                    cell.DataType = null; // Formula cells should not retain DataType
                    break;
                case "type":
                    cell.DataType = value.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        "date" => null, // Dates are stored as numbers; format is applied via numberformat below
                        _ => throw new ArgumentException($"Invalid cell 'type' value '{value}'. Valid types: string, number, boolean, date.")
                    };
                    // Convert cell value for boolean type
                    if (value.ToLowerInvariant() is "boolean" or "bool" && cell.CellValue != null)
                    {
                        var cv = cell.CellValue.Text.Trim().ToLowerInvariant();
                        if (cv is "true" or "yes") cell.CellValue = new CellValue("1");
                        else if (cv is "false" or "no") cell.CellValue = new CellValue("0");
                    }
                    // For date type, apply a default date number format unless caller already specifies one
                    if (value.Equals("date", StringComparison.OrdinalIgnoreCase)
                        && !properties.ContainsKey("numberformat") && !properties.ContainsKey("numfmt") && !properties.ContainsKey("format"))
                        styleProps["numberformat"] = "m/d/yy";
                    break;
                case "clear":
                    cell.CellValue = null;
                    cell.CellFormula = null;
                    cell.DataType = null; // Reset type on clear
                    cell.StyleIndex = null; // Also reset style/formatting
                    break;
                case "arrayformula":
                {
                    var arrRef = properties.GetValueOrDefault("ref", cellRef);
                    cell.CellFormula = new CellFormula(value.TrimStart('='))
                    {
                        FormulaType = CellFormulaValues.Array,
                        Reference = arrRef
                    };
                    cell.CellValue = null;
                    break;
                }
                case "link":
                {
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        hyperlinksEl?.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                        if (hyperlinksEl != null && !hyperlinksEl.HasChildren)
                            hyperlinksEl.Remove();
                    }
                    else
                    {
                        var hlUri = new Uri(value, UriKind.RelativeOrAbsolute);
                        var hlRel = worksheet.AddHyperlinkRelationship(hlUri, isExternal: true);
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
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid cell props: value, formula, arrayformula, type, clear, link, bold, italic, strike, underline, superscript, subscript, font.color, font.size, font.name, fill, border.all, alignment.horizontal, numfmt, locked, formulahidden)"
                            : key);
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

        return unsupported;
    }

    // ==================== Sheet-level Set (freeze panes) ====================

    private List<string> SetSheetLevel(WorksheetPart worksheet, string sheetName, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                {
                    // Rename the sheet
                    var workbook = GetWorkbook();
                    var sheets = workbook.Sheets?.Elements<Sheet>().ToList();
                    var sheet = sheets?.FirstOrDefault(s =>
                        s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
                    if (sheet != null)
                    {
                        var oldName = sheet.Name!.Value!;
                        sheet.Name = value;

                        // Excel stores sheet references in formulas as either:
                        //   SimpleSheetName!A1      (no spaces/special chars)
                        //   'Sheet With Spaces'!A1  (name with spaces or special chars)
                        static bool NeedsQuoting(string n) =>
                            n.Any(c => char.IsWhiteSpace(c) || c is '\'' or '[' or ']' or ':' or '*' or '?' or '/' or '\\');
                        static string FormulaRef(string n) => NeedsQuoting(n) ? $"'{n}'" : n;

                        var oldRef = FormulaRef(oldName) + "!";
                        var newRef = FormulaRef(value) + "!";

                        // Update named range references
                        var definedNames = workbook.GetFirstChild<DefinedNames>();
                        if (definedNames != null)
                        {
                            foreach (var dn in definedNames.Elements<DefinedName>())
                            {
                                if (dn.Text != null && dn.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                    dn.Text = dn.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                            }
                        }
                        // Update formula references in all cells across all sheets
                        foreach (var (_, wsPart) in GetWorksheets())
                        {
                            var sd = GetSheet(wsPart).GetFirstChild<SheetData>();
                            if (sd == null) continue;
                            foreach (var cell in sd.Descendants<Cell>())
                            {
                                if (cell.CellFormula?.Text != null &&
                                    cell.CellFormula.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                    cell.CellFormula.Text = cell.CellFormula.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                            }
                            GetSheet(wsPart).Save();
                        }
                        workbook.Save();
                    }
                    break;
                }
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
                case "merge":
                {
                    // Sheet-level merge: value is the range to merge (e.g., "A1:A3")
                    var rangeRef = value.ToUpperInvariant();
                    var mergeCells = ws.GetFirstChild<MergeCells>();
                    if (mergeCells == null)
                    {
                        mergeCells = new MergeCells();
                        ws.AppendChild(mergeCells);
                    }
                    var existing = mergeCells.Elements<MergeCell>()
                        .FirstOrDefault(m => m.Reference?.Value?.Equals(rangeRef, StringComparison.OrdinalIgnoreCase) == true);
                    if (existing == null)
                        mergeCells.AppendChild(new MergeCell { Reference = rangeRef });
                    break;
                }
                case "autofilter":
                {
                    // Set or remove AutoFilter (like POI's XSSFSheet.setAutoFilter)
                    var existingAf = ws.GetFirstChild<AutoFilter>();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        existingAf?.Remove();
                    }
                    else
                    {
                        if (existingAf != null)
                        {
                            existingAf.Reference = value.ToUpperInvariant();
                        }
                        else
                        {
                            var af = new AutoFilter { Reference = value.ToUpperInvariant() };
                            var sheetData = ws.GetFirstChild<SheetData>();
                            if (sheetData != null)
                                sheetData.InsertAfterSelf(af);
                            else
                                ws.AppendChild(af);
                        }
                    }
                    break;
                }
                case "zoom" or "zoomscale":
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
                    sheetView.ZoomScale = ParseHelpers.SafeParseUint(value, "zoom");
                    break;
                }

                case "tabcolor" or "tab_color":
                {
                    var sheetPr = ws.GetFirstChild<SheetProperties>();
                    if (sheetPr == null)
                    {
                        sheetPr = new SheetProperties();
                        ws.InsertAt(sheetPr, 0);
                    }
                    sheetPr.RemoveAllChildren<TabColor>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var colorHex = OfficeCli.Core.ParseHelpers.NormalizeArgbColor(value);
                        sheetPr.AppendChild(new TabColor { Rgb = new HexBinaryValue(colorHex) });
                    }
                    break;
                }

                // ==================== Sheet Protection ====================
                case "protect":
                {
                    var existingSp = ws.GetFirstChild<SheetProtection>();
                    if (ParseHelpers.IsTruthy(value))
                    {
                        if (existingSp == null)
                        {
                            existingSp = new SheetProtection();
                            ws.AppendChild(existingSp);
                        }
                        existingSp.Sheet = true;
                        existingSp.Objects = true;
                        existingSp.Scenarios = true;
                    }
                    else
                    {
                        existingSp?.Remove();
                    }
                    break;
                }
                case "password":
                {
                    var sp = ws.GetFirstChild<SheetProtection>();
                    if (sp == null)
                    {
                        sp = new SheetProtection { Sheet = true, Objects = true, Scenarios = true };
                        ws.AppendChild(sp);
                    }
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        sp.Password = null;
                    else
                    {
                        // Excel legacy password hash (ECMA-376 Part 4, 14.7.1)
                        int hash = 0;
                        for (int ci = value.Length - 1; ci >= 0; ci--)
                        {
                            hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
                            hash ^= value[ci];
                        }
                        hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
                        hash ^= value.Length;
                        hash ^= 0xCE4B;
                        sp.Password = HexBinaryValue.FromString(hash.ToString("X4"));
                    }
                    break;
                }

                // ==================== Print Settings ====================
                case "printarea":
                {
                    var workbook = GetWorkbook();
                    var definedNames = workbook.GetFirstChild<DefinedNames>()
                        ?? workbook.AppendChild(new DefinedNames());
                    // Find sheet index
                    var allSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                    var sheetIdx = allSheets?.FindIndex(s =>
                        s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
                    // Remove existing print area for this sheet
                    var existing = definedNames.Elements<DefinedName>()
                        .Where(d => d.Name == "_xlnm.Print_Area" && d.LocalSheetId?.Value == (uint)sheetIdx)
                        .ToList();
                    foreach (var e in existing) e.Remove();

                    if (!string.IsNullOrEmpty(value) && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var dn = new DefinedName($"{sheetName}!{value}") { Name = "_xlnm.Print_Area" };
                        if (sheetIdx >= 0) dn.LocalSheetId = (uint)sheetIdx;
                        definedNames.AppendChild(dn);
                    }
                    workbook.Save();
                    break;
                }
                case "orientation" or "pageorientation":
                {
                    var pageSetup = ws.GetFirstChild<PageSetup>();
                    if (pageSetup == null)
                    {
                        pageSetup = new PageSetup();
                        ws.AppendChild(pageSetup);
                    }
                    pageSetup.Orientation = value.ToLowerInvariant() == "landscape"
                        ? OrientationValues.Landscape
                        : OrientationValues.Portrait;
                    break;
                }
                case "papersize":
                {
                    var pageSetup = ws.GetFirstChild<PageSetup>();
                    if (pageSetup == null)
                    {
                        pageSetup = new PageSetup();
                        ws.AppendChild(pageSetup);
                    }
                    pageSetup.PaperSize = ParseHelpers.SafeParseUint(value, "paperSize");
                    break;
                }
                case "fittopage":
                {
                    var sheetPr = ws.GetFirstChild<SheetProperties>();
                    if (sheetPr == null)
                    {
                        sheetPr = new SheetProperties();
                        ws.InsertAt(sheetPr, 0);
                    }
                    var psp = sheetPr.GetFirstChild<PageSetupProperties>();
                    if (psp == null)
                    {
                        psp = new PageSetupProperties();
                        sheetPr.AppendChild(psp);
                    }
                    psp.FitToPage = true;

                    var pageSetup = ws.GetFirstChild<PageSetup>();
                    if (pageSetup == null)
                    {
                        pageSetup = new PageSetup();
                        ws.AppendChild(pageSetup);
                    }
                    // Parse "WxH" format (e.g., "1x2" for 1 page wide, 2 pages tall)
                    var fitParts = value.Split('x', 'X');
                    if (fitParts.Length == 2 && uint.TryParse(fitParts[0], out var fw) && uint.TryParse(fitParts[1], out var fh))
                    {
                        pageSetup.FitToWidth = fw;
                        pageSetup.FitToHeight = fh;
                    }
                    else if (ParseHelpers.IsTruthy(value))
                    {
                        pageSetup.FitToWidth = 1;
                        pageSetup.FitToHeight = 1;
                    }
                    break;
                }
                case "header":
                {
                    var hf = ws.GetFirstChild<HeaderFooter>();
                    if (hf == null)
                    {
                        hf = new HeaderFooter();
                        ws.AppendChild(hf);
                    }
                    hf.OddHeader = new OddHeader(value);
                    break;
                }
                case "footer":
                {
                    var hf = ws.GetFirstChild<HeaderFooter>();
                    if (hf == null)
                    {
                        hf = new HeaderFooter();
                        ws.AppendChild(hf);
                    }
                    hf.OddFooter = new OddFooter(value);
                    break;
                }

                // ==================== Sorting ====================
                case "sort":
                {
                    // Remove existing sort state
                    ws.GetFirstChild<SortState>()?.Remove();

                    if (!string.IsNullOrEmpty(value) && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var sd = ws.GetFirstChild<SheetData>();
                        if (sd == null) break;

                        // Value format: "A:asc" or "A:asc,B:desc" with optional range property
                        var sortRange = properties.GetValueOrDefault("range", "");
                        int startRow, endRow;
                        if (string.IsNullOrEmpty(sortRange))
                        {
                            var rows = sd.Elements<Row>().ToList();
                            if (rows.Count == 0) break;
                            var maxCol = rows.SelectMany(r => r.Elements<Cell>())
                                .Select(c => ParseCellReference(c.CellReference?.Value ?? "A1"))
                                .Max(p => ColumnNameToIndex(p.Column));
                            startRow = 1;
                            endRow = rows.Count;
                            sortRange = $"A1:{IndexToColumnName(maxCol)}{rows.Count}";
                        }
                        else
                        {
                            var rangeParts = sortRange.Split(':');
                            startRow = ParseCellReference(rangeParts[0]).Row;
                            endRow = ParseCellReference(rangeParts[1]).Row;
                        }

                        // Parse sort specifications
                        var specs = value.Split(',', StringSplitOptions.RemoveEmptyEntries);
                        var sortKeys = new List<(int ColIndex, bool Descending)>();
                        foreach (var spec in specs)
                        {
                            var specParts = spec.Trim().Split(':');
                            var colName = specParts[0].Trim().ToUpperInvariant();
                            bool descending = specParts.Length > 1 &&
                                specParts[1].Trim().Equals("desc", StringComparison.OrdinalIgnoreCase);
                            sortKeys.Add((ColumnNameToIndex(colName), descending));
                        }

                        // Actually sort the rows in SheetData
                        var rowsInRange = sd.Elements<Row>()
                            .Where(r => r.RowIndex?.Value >= (uint)startRow && r.RowIndex?.Value <= (uint)endRow)
                            .ToList();

                        // Extract sort values for each row
                        string GetCellSortValue(Row row, int colIdx)
                        {
                            var colLetter = IndexToColumnName(colIdx);
                            var cell = row.Elements<Cell>().FirstOrDefault(c =>
                                c.CellReference?.Value?.StartsWith(colLetter, StringComparison.OrdinalIgnoreCase) == true &&
                                ParseCellReference(c.CellReference.Value).Row == (int)(row.RowIndex?.Value ?? 0));
                            return cell != null ? GetCellDisplayValue(cell) : "";
                        }

                        var sorted = rowsInRange.OrderBy(_ => 0); // identity
                        foreach (var (colIdx, desc) in sortKeys)
                        {
                            var col = colIdx;
                            var d = desc;
                            // Always sort by rank ascending (empties last), then by value in requested direction
                            sorted = sorted.ThenBy(r => ParseSortValue(GetCellSortValue(r, col)).Rank);
                            sorted = d
                                ? sorted.ThenByDescending(r => ParseSortValue(GetCellSortValue(r, col)).NumVal)
                                         .ThenByDescending(r => ParseSortValue(GetCellSortValue(r, col)).StrVal)
                                : sorted.ThenBy(r => ParseSortValue(GetCellSortValue(r, col)).NumVal)
                                         .ThenBy(r => ParseSortValue(GetCellSortValue(r, col)).StrVal);
                        }
                        var sortedList = sorted.ToList();

                        // Collect original row indices and reassign
                        var originalIndices = rowsInRange.Select(r => r.RowIndex!.Value).ToList();
                        for (int si = 0; si < sortedList.Count; si++)
                        {
                            var row = sortedList[si];
                            var newRowIdx = originalIndices[si];
                            // Update row index and all cell references
                            row.RowIndex = newRowIdx;
                            foreach (var cell in row.Elements<Cell>())
                            {
                                if (cell.CellReference?.Value != null)
                                {
                                    var (col, _) = ParseCellReference(cell.CellReference.Value);
                                    cell.CellReference = $"{col}{newRowIdx}";
                                }
                            }
                        }

                        // Remove old rows and reinsert in sorted order
                        var beforeRow = sd.Elements<Row>()
                            .LastOrDefault(r => r.RowIndex?.Value < (uint)startRow);
                        foreach (var r in rowsInRange) r.Remove();
                        OpenXmlElement insertAfter = beforeRow ?? (OpenXmlElement)sd;
                        foreach (var row in sortedList)
                        {
                            if (insertAfter == sd)
                            {
                                sd.InsertAt(row, 0);
                                insertAfter = row;
                            }
                            else
                            {
                                insertAfter.InsertAfterSelf(row);
                                insertAfter = row;
                            }
                        }

                        // Write SortState metadata
                        var sortState = new SortState { Reference = sortRange };
                        foreach (var (colIdx, desc) in sortKeys)
                        {
                            var colName = IndexToColumnName(colIdx);
                            var condRef = $"{colName}{startRow}:{colName}{endRow}";
                            var sortCondition = new SortCondition { Reference = condRef };
                            if (desc) sortCondition.Descending = true;
                            sortState.AppendChild(sortCondition);
                        }
                        ws.AppendChild(sortState);
                    }
                    break;
                }

                default:
                    unsupported.Add(unsupported.Count == 0
                        ? $"{key} (valid sheet props: name, freeze, zoom, tabcolor, autofilter, merge, protect, password, printarea, orientation, papersize, fittopage, header, footer, sort)"
                        : key);
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

        // Separate range-level props from cell-level props
        var cellProps = new Dictionary<string, string>();
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
                    // Treat as cell-level property to apply to every cell in the range
                    cellProps[key] = value;
                    break;
            }
        }

        // Apply cell-level properties to every cell in the range (atomic: restore on failure)
        if (cellProps.Count > 0)
        {
            var parts = rangeRef.Split(':');
            var (startCol, startRow) = ParseCellReference(parts[0]);
            var (endCol, endRow) = ParseCellReference(parts[1]);
            var startColIdx = ColumnNameToIndex(startCol);
            var endColIdx = ColumnNameToIndex(endCol);

            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                ws.Append(sheetData);
            }

            // Clone SheetData so we can roll back if any cell fails mid-way
            var sheetDataBackup = (SheetData)sheetData.CloneNode(true);
            try
            {
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++)
                    {
                        var cellRef = $"{IndexToColumnName(colIdx)}{row}";
                        var cell = FindOrCreateCell(sheetData, cellRef);
                        var cellUnsupported = ApplyCellProperties(cell, worksheet, cellProps);
                        // Only add to unsupported once (first cell)
                        if (row == startRow && colIdx == startColIdx)
                            unsupported.AddRange(cellUnsupported);
                    }
                }
            }
            catch
            {
                ws.ReplaceChild(sheetDataBackup, sheetData);
                throw;
            }
        }

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
            var afterCol = columns.Elements<Column>().LastOrDefault(c => (c.Min?.Value ?? 0) < colIdx);
            if (afterCol != null)
                afterCol.InsertAfterSelf(col);
            else
                columns.PrependChild(col);
        }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "width":
                    col.Width = ParseHelpers.SafeParseDouble(value, "width");
                    col.CustomWidth = true;
                    break;
                case "hidden":
                    col.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "outline" or "outlinelevel" or "group":
                    if (!byte.TryParse(value, out var colOutline))
                        throw new ArgumentException($"Invalid 'outline' value: '{value}'. Expected an integer 0-7 (outline/group level).");
                    col.OutlineLevel = colOutline;
                    break;
                case "collapsed":
                    col.Collapsed = value.Equals("true", StringComparison.OrdinalIgnoreCase)
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
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var heightVal) || double.IsNaN(heightVal) || double.IsInfinity(heightVal))
                        throw new ArgumentException($"Invalid 'height' value: '{value}'. Expected a finite number (row height in points, e.g. 15.75).");
                    row.Height = heightVal;
                    row.CustomHeight = true;
                    break;
                case "hidden":
                    row.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "outline" or "outlinelevel" or "group":
                    if (!byte.TryParse(value, out var outlineVal))
                        throw new ArgumentException($"Invalid 'outline' value: '{value}'. Expected an integer 0-7 (outline/group level).");
                    row.OutlineLevel = outlineVal;
                    break;
                case "collapsed":
                    row.Collapsed = value.Equals("true", StringComparison.OrdinalIgnoreCase)
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
