// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for table-like paths (namedrange, validation,
// table column, table, comment, cf, pivot). Mechanically extracted from the
// original god-method Set(); each helper owns one path-pattern's full handling.
public partial class ExcelHandler
{
    private List<string> SetNamedRangeByPath(Match m, Dictionary<string, string> properties)
    {
        var selector = m.Groups[1].Value;
        var workbook = GetWorkbook();
        var definedNames = workbook.GetFirstChild<DefinedNames>()
            ?? throw new ArgumentException("No named ranges found in workbook");

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
                case "volatile":
                    // CONSISTENCY(definedname-volatile): map to the
                    // Function attribute (OOXML's only volatile signal
                    // for defined names) — see ExcelHandler.Add.Tables.cs.
                    if (IsTruthy(value)) dn.Function = true;
                    else dn.Function = null;
                    break;
                case "scope":
                    if (string.IsNullOrEmpty(value) || value.Equals("workbook", StringComparison.OrdinalIgnoreCase))
                    {
                        dn.LocalSheetId = null;
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

    private List<string> SetValidationByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var dvIdx = int.Parse(m.Groups[1].Value);
        var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>()
            ?? throw new ArgumentException("No data validations found in sheet");

        var dvList = dvs.Elements<DataValidation>().ToList();
        if (dvIdx < 1 || dvIdx > dvList.Count)
            throw new ArgumentException($"Validation index {dvIdx} out of range (1-{dvList.Count})");

        var dv = dvList[dvIdx - 1];
        var dvUnsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                // CONSISTENCY(canonical-key): schema canonical key is 'ref';
                // 'sqref' retained as legacy alias.
                case "sqref" or "ref":
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
                    // CONSISTENCY(validation-normalize): use same NormalizeValidationFormula
                    // as Add so range refs (C1:C3, Sheet1!A1:A3) are NOT double-quoted.
                    // Previous code only checked !value.StartsWith("\""), which incorrectly
                    // wrapped range refs that pass through unchanged in Add.
                    dv.Formula1 = new Formula1(NormalizeValidationFormula(value, dv.Type?.Value));
                    break;
                case "formula2":
                    dv.Formula2 = new Formula2(NormalizeValidationFormula(value, dv.Type?.Value));
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
                case "allowblank": dv.AllowBlank = IsTruthy(value); break;
                case "showerror": dv.ShowErrorMessage = IsTruthy(value); break;
                case "errortitle": dv.ErrorTitle = value; break;
                case "error": dv.Error = value; break;
                case "showinput": dv.ShowInputMessage = IsTruthy(value); break;
                case "prompttitle": dv.PromptTitle = value; break;
                case "prompt": dv.Prompt = value; break;
                // CONSISTENCY(validation-errorstyle): errorStyle was supported in Add
                // but missing from Set — silently fell into dvUnsupported.
                case "errorstyle":
                    dv.ErrorStyle = value.ToLowerInvariant() switch
                    {
                        "stop" => DataValidationErrorStyleValues.Stop,
                        "warning" or "warn" => DataValidationErrorStyleValues.Warning,
                        "information" or "info" => DataValidationErrorStyleValues.Information,
                        _ => throw new ArgumentException(
                            $"Unknown errorStyle: '{value}'. Use: stop, warning, information")
                    };
                    break;
                // CONSISTENCY(validation-incelldropdown): inCellDropdown was in Add (inverted
                // OOXML showDropDown semantics) but missing from Set. Also accept raw showDropDown.
                case "incelldropdown":
                    dv.ShowDropDown = !ParseHelpers.IsTruthy(value);
                    break;
                case "showdropdown":
                    dv.ShowDropDown = ParseHelpers.IsTruthy(value);
                    break;
                default: dvUnsupported.Add(key); break;
            }
        }

        SaveWorksheet(worksheet);
        return dvUnsupported;
    }

    // Replace backing embedded part + refresh ProgID. Cleans up the old payload
    // part (CLAUDE.md Known API Quirks rule: always delete the old part on src

    private List<string> SetTableColumnByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var tIdx = int.Parse(m.Groups[1].Value);
        var cIdx = int.Parse(m.Groups[2].Value);
        var tParts = worksheet.TableDefinitionParts.ToList();
        if (tIdx < 1 || tIdx > tParts.Count)
            throw new ArgumentException($"Table index {tIdx} out of range (1..{tParts.Count})");
        var tbl = tParts[tIdx - 1].Table
            ?? throw new ArgumentException($"Table {tIdx} has no definition");
        var tCols = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>().ToList();
        if (tCols == null || cIdx < 1 || cIdx > tCols.Count)
            throw new ArgumentException($"Column index {cIdx} out of range (1..{tCols?.Count ?? 0})");
        var tCol = tCols[cIdx - 1];
        var tcUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                {
                    tCol.Name = value;
                    // Sync the header-row cell so the worksheet matches the
                    // tableColumn @name. Excel rejects mismatch otherwise.
                    var refStr = tbl.Reference?.Value;
                    if (!string.IsNullOrEmpty(refStr) && (tbl.HeaderRowCount?.Value ?? 1) != 0)
                    {
                        var rParts = refStr.Split(':');
                        if (rParts.Length >= 1)
                        {
                            var (startCol, startRow) = ParseCellReference(rParts[0]);
                            var headerColIdx = ColumnNameToIndex(startCol) + (cIdx - 1);
                            var headerColLetter = IndexToColumnName(headerColIdx);
                            var headerCellRef = $"{headerColLetter}{startRow}";
                            var hdrWs = GetSheet(worksheet);
                            var hdrSheetData = hdrWs.GetFirstChild<SheetData>()
                                ?? hdrWs.AppendChild(new SheetData());
                            var hdrCell = FindOrCreateCell(hdrSheetData, headerCellRef);
                            hdrCell.CellValue = new CellValue(value);
                            hdrCell.DataType = CellValues.String;
                        }
                    }
                    break;
                }
                case "totalfunction" or "total":
                    tCol.TotalsRowFunction = value.ToLowerInvariant() switch
                    {
                        "sum" => TotalsRowFunctionValues.Sum,
                        "count" => TotalsRowFunctionValues.Count,
                        "average" or "avg" => TotalsRowFunctionValues.Average,
                        "max" => TotalsRowFunctionValues.Maximum,
                        "min" => TotalsRowFunctionValues.Minimum,
                        "stddev" => TotalsRowFunctionValues.StandardDeviation,
                        "var" => TotalsRowFunctionValues.Variance,
                        "countnums" => TotalsRowFunctionValues.CountNumbers,
                        "none" => TotalsRowFunctionValues.None,
                        "custom" => TotalsRowFunctionValues.Custom,
                        _ => throw new ArgumentException($"Invalid totalFunction: '{value}'.")
                    };
                    break;
                case "totallabel" or "label":
                    tCol.TotalsRowLabel = value;
                    break;
                case "formula":
                    tCol.CalculatedColumnFormula = new CalculatedColumnFormula(value);
                    break;
                default:
                    tcUnsupported.Add(key);
                    break;
            }
        }
        tParts[tIdx - 1].Table!.Save();
        SaveWorksheet(worksheet);
        return tcUnsupported;
    }

    private List<string> SetTableByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var tableIdx = int.Parse(m.Groups[1].Value);
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
                case "totalrow":
                case "showtotals":
                {
                    // CONSISTENCY(table-totalrow): mirror Add — toggling
                    // totalRow on must grow the table ref by one row to host
                    // the totals row OUTSIDE the data area (Excel rejects /
                    // pops a "found a problem" repair otherwise). Toggling
                    // off shrinks the ref symmetrically. AutoFilter ref
                    // tracks the data range only (header..last data row),
                    // so it stays one row shorter than table.Reference when
                    // a totals row is shown.
                    var totalRowEnabled = IsTruthy(value);
                    var prevTotalsCount = table.TotalsRowCount?.Value ?? 0u;
                    var refStr = table.Reference?.Value;
                    if (!string.IsNullOrEmpty(refStr) && refStr.Contains(':'))
                    {
                        var rParts = refStr.Split(':');
                        var (sCol, sRow) = ParseCellReference(rParts[0]);
                        var (eCol, eRow) = ParseCellReference(rParts[1]);
                        if (totalRowEnabled && prevTotalsCount == 0)
                        {
                            eRow += 1;
                        }
                        else if (!totalRowEnabled && prevTotalsCount > 0)
                        {
                            // Shrink only if there is at least one data row left.
                            if (eRow - 1 >= sRow) eRow -= 1;
                        }
                        var newTblRef = $"{sCol}{sRow}:{eCol}{eRow}";
                        table.Reference = newTblRef;
                        var afTbl = table.GetFirstChild<AutoFilter>();
                        if (afTbl != null)
                        {
                            // AutoFilter ref excludes the totals row.
                            var afEndRow = totalRowEnabled ? eRow - 1 : eRow;
                            if (afEndRow < sRow) afEndRow = sRow;
                            afTbl.Reference = $"{sCol}{sRow}:{eCol}{afEndRow}";
                        }
                    }
                    table.TotalsRowShown = totalRowEnabled;
                    table.TotalsRowCount = totalRowEnabled ? 1u : 0u;
                    break;
                }
                case "style":
                {
                    // CONSISTENCY(table-style-validation): mirror Add — short
                    // names like 'medium2' or 'foo' are not valid OOXML
                    // tableStyleInfo @name. Excel silently drops the style
                    // info on open, leaving the user wondering why the
                    // style didn't apply. Reject up-front with a clear
                    // message, same vocabulary as Add (see Helpers.cs
                    // ValidateTableStyleName).
                    // BUG-R9-B2: accept short aliases (medium2, light1, dark1, none).
                    var normalizedStyle = NormalizeTableStyleName(value) ?? value;
                    ValidateTableStyleName(normalizedStyle);
                    var styleInfo = table.GetFirstChild<TableStyleInfo>();
                    if (styleInfo != null) styleInfo.Name = normalizedStyle;
                    else table.AppendChild(new TableStyleInfo
                    {
                        Name = normalizedStyle, ShowFirstColumn = false, ShowLastColumn = false,
                        ShowRowStripes = true, ShowColumnStripes = false
                    });
                    break;
                }
                case "ref":
                {
                    var newRef = value.ToUpperInvariant();
                    // Grow/shrink <x:tableColumns> to match the new column count.
                    // Excel rejects the file when tableColumns.Count mismatches the
                    // ref width. On grow, append default ColumnN entries; on shrink,
                    // trim trailing entries.
                    var newParts = newRef.Split(':');
                    if (newParts.Length == 2)
                    {
                        var (nsc, _) = ParseCellReference(newParts[0]);
                        var (nec, _) = ParseCellReference(newParts[1]);
                        int newColCount = ColumnNameToIndex(nec) - ColumnNameToIndex(nsc) + 1;
                        var tc = table.GetFirstChild<TableColumns>();
                        if (tc != null && newColCount > 0)
                        {
                            var cols = tc.Elements<TableColumn>().ToList();
                            if (newColCount > cols.Count)
                            {
                                var existingIds = cols.Select(c => c.Id?.Value ?? 0u).ToList();
                                var existingNames = new HashSet<string>(
                                    cols.Select(c => c.Name?.Value ?? string.Empty),
                                    StringComparer.OrdinalIgnoreCase);
                                uint nextId = existingIds.Count > 0 ? existingIds.Max() + 1 : 1u;
                                for (int i = cols.Count; i < newColCount; i++)
                                {
                                    var baseName = $"Column{i + 1}";
                                    var name = baseName;
                                    int dedup = 2;
                                    while (!existingNames.Add(name))
                                        name = $"{baseName}{dedup++}";
                                    tc.AppendChild(new TableColumn { Id = nextId++, Name = name });
                                }
                            }
                            else if (newColCount < cols.Count)
                            {
                                for (int i = cols.Count - 1; i >= newColCount; i--)
                                    cols[i].Remove();
                            }
                            tc.Count = (uint)newColCount;
                        }
                    }
                    table.Reference = newRef;
                    var af = table.GetFirstChild<AutoFilter>();
                    if (af != null) af.Reference = newRef;
                    break;
                }
                case "showrowstripes" or "bandedrows" or "bandrows":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowRowStripes = IsTruthy(value);
                    break;
                }
                case "showcolstripes" or "showcolumnstripes" or "bandedcols" or "bandcols":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowColumnStripes = IsTruthy(value);
                    break;
                }
                case "showfirstcolumn" or "firstcol" or "firstcolumn":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowFirstColumn = IsTruthy(value);
                    break;
                }
                case "showlastcolumn" or "lastcol" or "lastcolumn":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowLastColumn = IsTruthy(value);
                    break;
                }
                case var k when k.StartsWith("col[") || k.StartsWith("column["):
                {
                    var tblColMatch = Regex.Match(k, @"^col(?:umn)?\[(\d+)\]\.(.+)$", RegexOptions.IgnoreCase);
                    if (!tblColMatch.Success) { tblUnsupported.Add(key); break; }
                    var colIdx = int.Parse(tblColMatch.Groups[1].Value);
                    var colProp = tblColMatch.Groups[2].Value.ToLowerInvariant();
                    var tableCols = table.GetFirstChild<TableColumns>()?.Elements<TableColumn>().ToList();
                    if (tableCols == null || colIdx < 1 || colIdx > tableCols.Count)
                        throw new ArgumentException($"Column index {colIdx} out of range (1..{tableCols?.Count ?? 0})");
                    var col = tableCols[colIdx - 1];
                    switch (colProp)
                    {
                        case "name": col.Name = value; break;
                        case "totalfunction" or "total":
                            col.TotalsRowFunction = value.ToLowerInvariant() switch
                            {
                                "sum" => TotalsRowFunctionValues.Sum,
                                "count" => TotalsRowFunctionValues.Count,
                                "average" or "avg" => TotalsRowFunctionValues.Average,
                                "max" => TotalsRowFunctionValues.Maximum,
                                "min" => TotalsRowFunctionValues.Minimum,
                                "stddev" => TotalsRowFunctionValues.StandardDeviation,
                                "var" => TotalsRowFunctionValues.Variance,
                                "countnums" => TotalsRowFunctionValues.CountNumbers,
                                "none" => TotalsRowFunctionValues.None,
                                "custom" => TotalsRowFunctionValues.Custom,
                                _ => throw new ArgumentException($"Invalid totalFunction: '{value}'. Valid: sum, count, average, max, min, stddev, var, countNums, none, custom.")
                            };
                            break;
                        case "totallabel" or "label":
                            col.TotalsRowLabel = value;
                            break;
                        case "formula":
                            col.CalculatedColumnFormula = new CalculatedColumnFormula(value);
                            break;
                        default: tblUnsupported.Add(key); break;
                    }
                    break;
                }
                default: tblUnsupported.Add(key); break;
            }
        }

        tableParts[tableIdx - 1].Table!.Save();
        return tblUnsupported;
    }

    private List<string> SetCommentByPath(Match m, WorksheetPart worksheet, string sheetName, Dictionary<string, string> properties)
    {
        var cmtIndex = int.Parse(m.Groups[1].Value);
        var commentsPart = worksheet.WorksheetCommentsPart;
        if (commentsPart?.Comments == null)
            throw new ArgumentException($"No comments found in sheet: {sheetName}");

        var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
        var cmtElement = cmtList?.Elements<Comment>().ElementAtOrDefault(cmtIndex - 1)
            ?? throw new ArgumentException($"Comment [{cmtIndex}] not found");

        var cmtUnsupported = new List<string>();
        // CONSISTENCY(xlsx/comment-font): C8 — font.* props on Set rewrite the
        // single <x:r><x:rPr>, reusing BuildCommentRunProperties. When `text` and
        // `font.*` appear together, text wins the run payload and font.* supplies
        // the rPr. When only font.* appears (no text), preserve the existing run
        // text and just rebuild rPr.
        string? newCmtText = properties.TryGetValue("text", out var tVal)
            ? (tVal ?? string.Empty).Replace("\r\n", "\n")
            : null;
        bool hasFontProp = properties.Keys.Any(k =>
            k.StartsWith("font.", StringComparison.OrdinalIgnoreCase));
        if (newCmtText != null || hasFontProp)
        {
            string runText = newCmtText
                ?? string.Concat(cmtElement.CommentText?.Elements<Run>()
                    .SelectMany(r => r.Elements<Text>()).Select(t => t.Text)
                    ?? Array.Empty<string>());
            cmtElement.CommentText = new CommentText(
                new Run(
                    BuildCommentRunProperties(properties),
                    new Text(runText) { Space = SpaceProcessingModeValues.Preserve }
                )
            );
        }
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                case var k1 when k1.StartsWith("font."):
                    break;
                case "ref":
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

    private List<string> SetCfByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var cfIdx = int.Parse(m.Groups[1].Value);
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
                case "range":
                case "ref":
                    // CONSISTENCY(cf-sqref): accept ref/range/sqref aliases on Set
                    // — same vocabulary as conditionalformatting Add (Add.Cf.cs).
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
                case "midcolor":
                {
                    // 3-stop color scale only — assumes the rule already has min/mid/max.
                    var csColors3 = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                    if (csColors3 != null && csColors3.Count >= 3)
                        csColors3[1].Rgb = ParseHelpers.NormalizeArgbColor(value);
                    else unsup.Add(key);
                    break;
                }
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
                {
                    // showValue applies to both IconSet and DataBar rules.
                    var isEl2 = rule?.GetFirstChild<IconSet>();
                    var dbEl = rule?.GetFirstChild<DataBar>();
                    if (isEl2 != null) isEl2.ShowValue = IsTruthy(value);
                    else if (dbEl != null) dbEl.ShowValue = IsTruthy(value);
                    else unsup.Add(key);
                    break;
                }
                case "minlength":
                {
                    var dbEl = rule?.GetFirstChild<DataBar>();
                    if (dbEl != null && uint.TryParse(value, out var mlen))
                    {
                        dbEl.MinLength = mlen;
                        var x14Db = ResolveX14DataBar(ws, rule!);
                        if (x14Db != null) x14Db.MinLength = mlen;
                    }
                    else unsup.Add(key);
                    break;
                }
                case "maxlength":
                {
                    var dbEl = rule?.GetFirstChild<DataBar>();
                    if (dbEl != null && uint.TryParse(value, out var xlen))
                    {
                        dbEl.MaxLength = xlen;
                        var x14Db = ResolveX14DataBar(ws, rule!);
                        if (x14Db != null) x14Db.MaxLength = xlen;
                    }
                    else unsup.Add(key);
                    break;
                }
                case "negativecolor":
                {
                    var x14Db = rule != null ? ResolveX14DataBar(ws, rule) : null;
                    if (x14Db != null)
                    {
                        x14Db.RemoveAllChildren<X14.NegativeFillColor>();
                        x14Db.Append(new X14.NegativeFillColor { Rgb = ParseHelpers.NormalizeArgbColor(value) });
                    }
                    else unsup.Add(key);
                    break;
                }
                case "axiscolor":
                {
                    var x14Db = rule != null ? ResolveX14DataBar(ws, rule) : null;
                    if (x14Db != null)
                    {
                        x14Db.RemoveAllChildren<X14.BarAxisColor>();
                        x14Db.Append(new X14.BarAxisColor { Rgb = ParseHelpers.NormalizeArgbColor(value) });
                    }
                    else unsup.Add(key);
                    break;
                }
                case "direction":
                {
                    var x14Db = rule != null ? ResolveX14DataBar(ws, rule) : null;
                    if (x14Db != null)
                    {
                        var dirNorm = SchemaKeyNormalizer.Normalize(value);
                        x14Db.Direction = dirNorm switch
                        {
                            "lefttoright" or "ltr" => X14.DataBarDirectionValues.LeftToRight,
                            "righttoleft" or "rtl" => X14.DataBarDirectionValues.RightToLeft,
                            "context" => X14.DataBarDirectionValues.Context,
                            _ => X14.DataBarDirectionValues.Context
                        };
                    }
                    else unsup.Add(key);
                    break;
                }
                case "percent":
                {
                    // top/bottom rules: percent=true treats `rank` as a
                    // percentile (top N%) instead of an absolute count
                    // (top N). Schema declares add/set/get; Add has it
                    // wired but Set was missing.
                    if (rule != null) rule.Percent = IsTruthy(value);
                    else unsup.Add(key);
                    break;
                }
                default:
                    unsup.Add(key);
                    break;
            }
        }
        SaveWorksheet(worksheet);
        return unsup;
    }

    /// <summary>
    /// Resolve the x14:dataBar element paired with a 2007 dataBar rule via x14:id reference.
    /// Returns null if the rule has no x14 extension or the worksheet has no matching x14 cf.
    /// </summary>
    private static X14.DataBar? ResolveX14DataBar(Worksheet ws, ConditionalFormattingRule rule)
    {
        var extList = rule.GetFirstChild<ConditionalFormattingRuleExtensionList>();
        if (extList == null) return null;
        var idExt = extList.Elements<ConditionalFormattingRuleExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}", StringComparison.OrdinalIgnoreCase));
        var refId = idExt?.GetFirstChild<X14.Id>()?.Text;
        if (string.IsNullOrEmpty(refId)) return null;

        const string cfExtUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
        var wsExtList = ws.GetFirstChild<WorksheetExtensionList>();
        if (wsExtList == null) return null;
        foreach (var wsExt in wsExtList.Elements<WorksheetExtension>().Where(e => e.Uri == cfExtUri))
        {
            foreach (var x14Cfs in wsExt.Elements<X14.ConditionalFormattings>())
            foreach (var x14Cf in x14Cfs.Elements<X14.ConditionalFormatting>())
            foreach (var x14Rule in x14Cf.Elements<X14.ConditionalFormattingRule>())
            {
                if (string.Equals(x14Rule.Id?.Value, refId, StringComparison.OrdinalIgnoreCase))
                    return x14Rule.GetFirstChild<X14.DataBar>();
            }
        }
        return null;
    }

    private List<string> SetPivotTableByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var ptIdx = int.Parse(m.Groups[1].Value);
        var pivotParts = worksheet.PivotTableParts.ToList();
        if (ptIdx < 1 || ptIdx > pivotParts.Count)
            throw new ArgumentException($"PivotTable {ptIdx} not found");
        return PivotTableHelper.SetPivotTableProperties(pivotParts[ptIdx - 1], properties);
    }
}
