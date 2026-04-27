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
using ThreadedCmt = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        // Batch Set: if path looks like a selector (not starting with /), Query → Set each
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
        {
            var unsupported = new List<string>();
            var targets = Query(path);
            if (targets.Count == 0)
                throw new ArgumentException($"No elements matched selector: {path}");
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        // Normalize to case-insensitive lookup so camelCase keys match lowercase lookups
        if (properties != null && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        path = NormalizeExcelPath(path);
        path = ResolveSheetIndexInPath(path);

        // Excel only supports find+replace — reject find without replace early (before path dispatch)
        if (properties.ContainsKey("find") && !properties.ContainsKey("replace"))
            throw new ArgumentException("Excel only supports 'find' with 'replace'. Use 'find' + 'replace' for text replacement. find+format (without replace) is not supported in Excel.");
        if (properties.ContainsKey("regex") && properties.ContainsKey("find"))
            throw new ArgumentException("Excel find+replace does not support regex. Remove 'regex' property.");

        // Handle root path "/" — document properties
        if (path == "/")
        {
            // Find & Replace: special handling before document properties
            if (properties.TryGetValue("find", out var findText) && properties.TryGetValue("replace", out var replaceText))
            {
                var count = FindAndReplace(findText, replaceText, null);
                LastFindMatchCount = count;
                var remaining = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
                remaining.Remove("find");
                remaining.Remove("replace");
                if (remaining.Count > 0)
                    return Set(path, remaining);
                return [];
            }

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
                        var lowerKey = key.ToLowerInvariant();
                        if (!TrySetWorkbookSetting(lowerKey, value)
                            && !Core.ThemeHandler.TrySetTheme(_doc.WorkbookPart?.ThemePart, lowerKey, value)
                            && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                                Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                            unsupported.Add(key);
                        break;
                }
            }
            return unsupported;
        }

        // Handle /SheetName/sparkline[N]
        var sparklineSetMatch = Regex.Match(path.TrimStart('/'), @"^([^/]+)/sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (sparklineSetMatch.Success) return SetSparklineByPath(sparklineSetMatch, properties);

        // Handle /namedrange[N] or /namedrange[Name]
        var namedRangeMatch = Regex.Match(path.TrimStart('/'), @"^namedrange\[(.+?)\]$", RegexOptions.IgnoreCase);
        if (namedRangeMatch.Success) return SetNamedRangeByPath(namedRangeMatch, properties);

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
        if (validationSetMatch.Success) return SetValidationByPath(validationSetMatch, worksheet, properties);

        // Handle /SheetName/ole[N]
        var oleSetMatch = Regex.Match(cellRef, @"^(?:ole|object|embed)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (oleSetMatch.Success) return SetOleByPath(oleSetMatch, worksheet, properties);

        // Handle /SheetName/picture[N]
        var picSetMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (picSetMatch.Success) return SetPictureByPath(picSetMatch, worksheet, properties);

        // Handle /SheetName/shape[N]
        var shapeSetMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (shapeSetMatch.Success) return SetShapeByPath(shapeSetMatch, worksheet, properties);

        // Handle /SheetName/slicer[N] — caption/style/columnCount/rowHeight/name
        var slicerSetMatch = Regex.Match(cellRef, @"^slicer\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (slicerSetMatch.Success) return SetSlicerByPath(slicerSetMatch, worksheet, properties);

        // Handle /SheetName/table[N]/columns[M] or /SheetName/table[N]/column[M]
        // CONSISTENCY(table-column-path): mirror the col[M].prop= dotted form already
        // accepted on /Sheet/table[N] by exposing the column as a sub-path so users can
        // address it as a node and call Set with a flat property bag.
        var tableColPathMatch = Regex.Match(cellRef,
            @"^table\[(\d+)\]/(?:columns|column)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableColPathMatch.Success) return SetTableColumnByPath(tableColPathMatch, worksheet, properties);

        // Handle /SheetName/table[N]
        var tableSetMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableSetMatch.Success) return SetTableByPath(tableSetMatch, worksheet, properties);

        // Handle /SheetName/comment[N]
        var commentSetMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentSetMatch.Success) return SetCommentByPath(commentSetMatch, worksheet, sheetName, properties);

        // Handle /SheetName/autofilter
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            return SetAutoFilter(worksheet, properties);
        }

        // Handle /SheetName/cf[N] or /SheetName/conditionalformatting[N]
        var cfSetMatch = Regex.Match(cellRef, @"^(?:cf|conditionalformatting)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cfSetMatch.Success) return SetCfByPath(cfSetMatch, worksheet, properties);

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

        // Handle /SheetName/chart[N]/axis[@role=ROLE]
        var chartAxisSetMatch = Regex.Match(cellRef,
            @"^chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success) return SetChartAxisByPath(chartAxisSetMatch, worksheet, properties);

        // Handle /SheetName/chart[N] or /SheetName/chart[N]/series[K]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success) return SetChartByPath(chartMatch, worksheet, properties);

        // Handle /SheetName/pivottable[N]
        var pivotSetMatch = Regex.Match(cellRef, @"^pivottable\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (pivotSetMatch.Success) return SetPivotTableByPath(pivotSetMatch, worksheet, properties);

        // Handle /SheetName/A1/run[N] (rich text run)
        var runSetMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runSetMatch.Success) return SetCellRunByPath(runSetMatch, worksheet, properties);

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
        // Remove completely empty cells (no value, no formula, no custom style) so that
        // rows with no remaining cells are pruned from XML. This keeps maxRow correct
        // and produces "remove" watch patches instead of "replace" for cleared rows.
        PruneEmptyCell(cell);
        // CONSISTENCY(xlsx/table-autoexpand): eager post-write auto-grow —
        // only fires when the cell still carries a value/formula after prune.
        if (cell.Parent != null && (cell.CellValue != null || cell.CellFormula != null || cell.InlineString != null))
            MaybeExpandTablesForCell(worksheet, cellRef);
        // Any mutation to a cell (value, formula, clear) can invalidate the calc chain
        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return unsupported;
    }

    private void PruneEmptyCell(Cell cell)
    {
        var hasValue = cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.Text);
        var hasFormula = cell.CellFormula != null;
        var hasStyle = cell.StyleIndex != null && cell.StyleIndex.Value != 0;
        if (!hasValue && !hasFormula && !hasStyle)
        {
            var row = cell.Parent as Row;
            cell.Remove();
            if (row != null && !row.Elements<Cell>().Any())
            {
                // Capture sheetData and rowIdx before detaching — row.Parent is null after Remove()
                var sheetData = row.Parent as SheetData;
                var rowIdx = row.RowIndex?.Value;
                row.Remove();
                // Keep row index cache in sync: detached row must not be returned by FindOrCreateRow
                if (sheetData != null && rowIdx.HasValue)
                    _rowIndex?.GetValueOrDefault(sheetData)?.Remove(rowIdx.Value);
            }
        }
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
            if (value is null) continue;
            if (ExcelStyleManager.IsStyleKey(key))
            {
                styleProps[key] = value;
                continue;
            }

            switch (key.ToLowerInvariant())
            {
                case "value" or "text":
                    // bt-3: if the cell already carries a text number format
                    // ("@", numFmtId 49) from a prior `set numberformat=@`,
                    // honor it on subsequent value updates by forcing the cell
                    // to String storage. Skip when the user is overriding the
                    // numberformat in this same call (styleProps captures that
                    // path via IsTextNumberFormat already).
                    bool existingIsTextFmt = false;
                    if (!properties.ContainsKey("numberformat")
                        && !properties.ContainsKey("numfmt")
                        && !properties.ContainsKey("format")
                        && !properties.ContainsKey("type"))
                    {
                        var (existingNumFmtId, existingFmtCode) = ExcelDataFormatter.GetCellFormat(cell, _doc.WorkbookPart);
                        if (existingNumFmtId == 49
                            || (existingFmtCode != null && existingFmtCode.Trim() == "@"))
                            existingIsTextFmt = true;
                    }
                    // R28-B4 — leading apostrophe is Excel's "force text" idiom.
                    // Strip the apostrophe from the stored value and stamp
                    // quotePrefix=1 on the cell xf so Excel renders the value
                    // literally as text without the apostrophe glyph. Cell type
                    // is forced to String below via the local quotePrefixForce flag
                    // (we can't safely add to `properties` mid-foreach).
                    bool quotePrefixForce = false;
                    string effectiveValue = value;
                    if (effectiveValue.StartsWith('\'') && effectiveValue.Length > 1)
                    {
                        effectiveValue = effectiveValue.Substring(1);
                        styleProps["quoteprefix"] = "true";
                        quotePrefixForce = true;
                    }
                    // R13-1: enforce Excel's 32767-char per-cell limit.
                    EnsureCellValueLength(effectiveValue, cell.CellReference?.Value);
                    // R13-3: warn if both value= and formula= supplied — formula
                    // takes precedence below (explicit-formula case runs last and
                    // clears CellValue), so the literal value is silently discarded.
                    if (properties.Any(p => p.Key.Equals("formula", StringComparison.OrdinalIgnoreCase)))
                    {
                        Console.Error.WriteLine(
                            "Warning: Both value= and formula= supplied — using formula, value ignored.");
                    }
                    // Auto-detect formula: value starting with '=' is treated as formula
                    if (effectiveValue.StartsWith('=') && effectiveValue.Length > 1)
                        goto case "formula";
                    var cellValue = effectiveValue.Replace("\\n", "\n"); // Support escaped newlines
                    cell.CellFormula = null; // Clear formula when explicit value is set
                    // If cell is already boolean type, convert true/false to 1/0
                    if (cell.DataType?.Value == CellValues.Boolean)
                    {
                        var bv = cellValue.Trim().ToLowerInvariant();
                        if (bv is "true" or "yes") cell.CellValue = new CellValue("1");
                        else if (bv is "false" or "no") cell.CellValue = new CellValue("0");
                        else cell.CellValue = new CellValue(cellValue);
                    }
                    else
                    {
                        // Check if user explicitly set type
                        var hasExplicitType = properties.Any(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase));
                        var explicitTypeIsString = quotePrefixForce || existingIsTextFmt || (hasExplicitType && properties
                            .Where(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase))
                            .Select(p => p.Value?.ToLowerInvariant())
                            .Any(v => v is "string" or "str"));
                        var explicitTypeIsNumber = hasExplicitType && properties
                            .Where(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase))
                            .Select(p => p.Value?.ToLowerInvariant())
                            .Any(v => v is "number" or "num");

                        // Auto-detect ISO date (only if user did NOT explicitly set type=string)
                        // R13-2: accept date-with-time variants (T and space separators).
                        if (!explicitTypeIsString && TryParseIsoDateFlexible(cellValue, out var dt))
                        {
                            cell.CellValue = new CellValue(dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture));
                            cell.DataType = null;
                            if (!properties.ContainsKey("numberformat") && !properties.ContainsKey("numfmt") && !properties.ContainsKey("format"))
                                styleProps["numberformat"] = "yyyy-mm-dd";
                        }
                        // Auto-detect strings that look like numbers but should be text
                        else if (!explicitTypeIsNumber
                            && ((cellValue.Length > 1 && cellValue.StartsWith('0') && !cellValue.StartsWith("0.") && !cellValue.StartsWith("0,") && cellValue.All(c => char.IsDigit(c)))
                                || (cellValue.All(char.IsDigit) && cellValue.Length > 15)))
                        {
                            cell.CellValue = new CellValue(cellValue);
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        else if (explicitTypeIsString)
                        {
                            // R15-2: honor explicit type=string even for
                            // numeric-looking literals. Without this, Excel
                            // renders 123 as a number despite user intent.
                            cell.CellValue = new CellValue(cellValue);
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        else if (explicitTypeIsNumber)
                        {
                            // R15-2: honor explicit type=number — refuse
                            // non-numeric values rather than silently storing
                            // as string. R32-1: also refuse NaN/Infinity even
                            // though TryParse may accept them — they are not
                            // valid xs:double cell content.
                            if (!double.TryParse(cellValue, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var numDbl)
                                || !double.IsFinite(numDbl))
                                throw new ArgumentException(
                                    $"Cannot store '{cellValue}' as number; use type=string or remove type=");
                            cell.CellValue = new CellValue(cellValue);
                            cell.DataType = null;
                        }
                        else
                        {
                            cell.CellValue = new CellValue(cellValue);
                            // R32-1: double.TryParse("NaN") returns true; without
                            // an IsFinite gate, the cell would be written with
                            // no t= attribute (numeric default) and content
                            // "NaN", which Excel rejects as invalid xs:double.
                            // Force string storage for non-finite doubles,
                            // matching how "Infinity" already behaves.
                            if (double.TryParse(cellValue, out var dbl) && double.IsFinite(dbl))
                                cell.DataType = null;
                            else
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                    }
                    break;
                case "formula":
                    // BUG-R36-03 fix: reject empty/whitespace formula strings.
                    // Storing an empty CellFormula (<x:f/>) is invalid OOXML and causes
                    // Get() to return "=" as the cell text. Treat as clear-formula intent.
                    if (string.IsNullOrWhiteSpace(value))
                        throw new ArgumentException(
                            "Formula cannot be empty or whitespace. " +
                            "To clear a formula use --prop value= (set to a plain value) or --prop clear=true.");
                    RejectCrossWorkbookFormula(value);
                    cell.CellFormula = new CellFormula(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(value.TrimStart('='))));
                    // Try to evaluate and cache the result immediately
                    var evalSheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
                    var evaluator = new Core.FormulaEvaluator(evalSheetData!, _doc.WorkbookPart);
                    var evalResult = evaluator.TryEvaluateFull(value.TrimStart('='));
                    if (evalResult is { IsNumeric: true })
                    {
                        cell.CellValue = new CellValue(evalResult.ToCellValueText());
                        cell.DataType = null;
                    }
                    else if (evalResult is { IsString: true })
                    {
                        cell.CellValue = new CellValue(evalResult.StringValue!);
                        cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                    }
                    else if (evalResult is { IsBool: true })
                    {
                        cell.CellValue = new CellValue(evalResult.ToCellValueText());
                        cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Boolean);
                    }
                    else if (evalResult is { IsError: true })
                    {
                        cell.CellValue = new CellValue(evalResult.ErrorValue!);
                        cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Error);
                    }
                    else
                    {
                        // Formula written but not evaluated — will be calculated when opened in Excel
                        cell.CellValue = null;
                        cell.DataType = null;
                    }
                    // Ensure fullCalcOnLoad so Excel recalculates formulas on open
                    {
                        var workbook = _doc.WorkbookPart!.Workbook!;
                        var calcPr = workbook.GetFirstChild<CalculationProperties>();
                        if (calcPr == null)
                        {
                            calcPr = new CalculationProperties();
                            // OOXML schema order: ...definedNames, calcPr, oleSize, customWorkbookViews, pivotCaches...
                            var insertBefore = (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.OleSize>()
                                ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CustomWorkbookViews>()
                                ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PivotCaches>();
                            if (insertBefore != null)
                                workbook.InsertBefore(calcPr, insertBefore);
                            else
                                workbook.AppendChild(calcPr);
                        }
                        calcPr.FullCalculationOnLoad = true;
                    }
                    break;
                case "type":
                    cell.DataType = value.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        "date" => null, // Dates are stored as numbers; format is applied via numberformat below
                        // CONSISTENCY(cell-type-parity): accept `error`/`err` as in Add.
                        "error" or "err" => new EnumValue<CellValues>(CellValues.Error),
                        _ => throw new ArgumentException($"Invalid cell 'type' value '{value}'. Valid types: string, number, boolean, date, error.")
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
                    RejectCrossWorkbookFormula(value);
                    var arrRef = properties.GetValueOrDefault("ref", cellRef);
                    cell.CellFormula = new CellFormula(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(value.TrimStart('='))))
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
                        // Symmetric to H3 above: when removing a hyperlink,
                        // also drop the implicit Hyperlink cellStyle that
                        // Add/Set installed (blue + underline). User-assigned
                        // explicit styles are preserved — we only revert
                        // StyleIndex values that match the Hyperlink xf.
                        if (cell.StyleIndex != null && cell.StyleIndex.Value != 0)
                        {
                            var wbPart = _doc.WorkbookPart;
                            if (wbPart != null)
                            {
                                var styleManager = new ExcelStyleManager(wbPart);
                                if (styleManager.IsHyperlinkCellStyleXf(cell.StyleIndex.Value))
                                {
                                    cell.StyleIndex = null;
                                    _dirtyStylesheet = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (hyperlinksEl == null)
                        {
                            hyperlinksEl = new Hyperlinks();
                            ws.AppendChild(hyperlinksEl);
                        }
                        hyperlinksEl.Elements<Hyperlink>()
                            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                            .ToList().ForEach(h => h.Remove());
                        // H2: optional tooltip/screenTip from sibling props.
                        var setHlTip = properties.GetValueOrDefault("tooltip")
                            ?? properties.GetValueOrDefault("screenTip")
                            ?? properties.GetValueOrDefault("screentip");
                        // R37-B: also accept bare `SheetName!Cell` (no '#' prefix)
                        // and quoted `'Multi Word'!Cell` as internal targets.
                        // CONSISTENCY(internal-hyperlink): same detection used in Add.Cells.cs.
                        var internalLoc = TryParseInternalHyperlinkLocation(value);
                        if (internalLoc != null)
                        {
                            // Internal target (sheet cell or named range) is
                            // written as an in-document hyperlink via the
                            // `location` attribute, no relationship/target.
                            var hl = new Hyperlink
                            {
                                Reference = cellRef.ToUpperInvariant(),
                                Location = internalLoc
                            };
                            if (!string.IsNullOrEmpty(setHlTip)) hl.Tooltip = setHlTip;
                            hyperlinksEl.AppendChild(hl);
                        }
                        else
                        {
                            var hlUri = new Uri(value, UriKind.RelativeOrAbsolute);
                            var hlRel = worksheet.AddHyperlinkRelationship(hlUri, isExternal: true);
                            var hl = new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id };
                            if (!string.IsNullOrEmpty(setHlTip)) hl.Tooltip = setHlTip;
                            hyperlinksEl.AppendChild(hl);
                        }
                        // H3: apply the built-in "Hyperlink" cellStyle (blue +
                        // underline) if the cell has no user-assigned style.
                        // CONSISTENCY(hyperlink-cellstyle): preserve an
                        // explicit StyleIndex the user already set.
                        if (cell.StyleIndex == null || cell.StyleIndex.Value == 0)
                        {
                            var wbPart = _doc.WorkbookPart
                                ?? throw new InvalidOperationException("Workbook not found");
                            var styleManager = new ExcelStyleManager(wbPart);
                            cell.StyleIndex = styleManager.EnsureHyperlinkCellStyle();
                            _dirtyStylesheet = true;
                        }
                    }
                    break;
                }
                case "merge":
                {
                    // CONSISTENCY(cell-merge): cell Add already accepts
                    // merge=A1:C3 (see ExcelHandler.Add.Cells.cs); cell Set
                    // mirrors it. Empty/false/none/unmerge clears any merge
                    // anchored at this cell.
                    var ws = GetSheet(worksheet);
                    var mergeCellsEl = ws.GetFirstChild<MergeCells>();
                    var clear = string.IsNullOrWhiteSpace(value)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("unmerge", StringComparison.OrdinalIgnoreCase);
                    if (clear)
                    {
                        // Drop any merge whose top-left equals this cell.
                        if (mergeCellsEl != null)
                        {
                            foreach (var mc in mergeCellsEl.Elements<MergeCell>().ToList())
                            {
                                var refStr = mc.Reference?.Value ?? "";
                                var topLeft = refStr.Split(':')[0];
                                if (string.Equals(topLeft, cellRef, StringComparison.OrdinalIgnoreCase))
                                    mc.Remove();
                            }
                            if (!mergeCellsEl.HasChildren) mergeCellsEl.Remove();
                            else mergeCellsEl.Count = (uint)mergeCellsEl.Elements<MergeCell>().Count();
                        }
                    }
                    else
                    {
                        if (mergeCellsEl == null)
                        {
                            mergeCellsEl = new MergeCells();
                            ws.AppendChild(mergeCellsEl);
                        }
                        foreach (var rangeRef in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                            InsertMergeCellChecked(mergeCellsEl, rangeRef);
                        mergeCellsEl.Count = (uint)mergeCellsEl.Elements<MergeCell>().Count();
                    }
                    break;
                }
                case "tooltip":
                case "screentip":
                {
                    // H2: tooltip may also be applied to an EXISTING hyperlink.
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    var existing = hyperlinksEl?.Elements<Hyperlink>()
                        .FirstOrDefault(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
                    if (existing == null)
                    {
                        unsupported.Add($"tooltip (no hyperlink exists on {cellRef}; add a link first)");
                        break;
                    }
                    existing.Tooltip = string.IsNullOrEmpty(value) ? null : value;
                    break;
                }
                default:
                    // Check for known flat-key misuse first, even before generic
                    // attribute fallback — otherwise user typos like `size=14`
                    // would be silently written as unknown XML attributes.
                    var cellHint = CellPropHints.TryGetHint(key);
                    if (cellHint != null)
                    {
                        unsupported.Add(cellHint);
                    }
                    else if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                    {
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid cell props: value, formula, arrayformula, type, clear, link, bold, italic, strike, underline, superscript, subscript, font.color, font.size, font.name, fill, border.all, alignment.horizontal, numfmt, locked, formulahidden)"
                            : key);
                    }
                    break;
            }
        }

        // Apply style properties if any
        if (styleProps.Count > 0)
        {
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var styleManager = new ExcelStyleManager(workbookPart);
            cell.StyleIndex = styleManager.ApplyStyle(cell, styleProps, unsupported);
            _dirtyStylesheet = true;

            // R24-1: numberformat="@" → force text storage. See ExcelHandler.Add.Cells.cs
            // for the matching guard on the Add path.
            if (IsTextNumberFormat(styleProps)
                && cell.DataType?.Value != CellValues.SharedString
                && cell.DataType?.Value != CellValues.InlineString
                && cell.CellFormula == null)
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
        }

        return unsupported;
    }

    // ==================== Sheet-level Set (freeze panes) ====================

    private List<string> SetSheetLevel(WorksheetPart worksheet, string sheetName, Dictionary<string, string> properties)
    {
        // Find & Replace at sheet level
        if (properties.TryGetValue("find", out var findText) && properties.TryGetValue("replace", out var replaceText))
        {
            var count = FindAndReplace(findText, replaceText, worksheet);
            LastFindMatchCount = count;
            var remaining = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            remaining.Remove("find");
            remaining.Remove("replace");
            if (remaining.Count > 0)
                return SetSheetLevel(worksheet, sheetName, remaining);
            return [];
        }

        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                {
                    // Validate sheet name up-front so Excel doesn't reject the file
                    // on open. Rules per Excel:
                    //   - cannot be empty / blank
                    //   - max 31 chars
                    //   - cannot contain  \  /  ?  *  :  [  ]
                    //   - cannot start or end with apostrophe '
                    //   - cannot equal reserved name "History"
                    ValidateSheetName(value);

                    // Rename the sheet
                    var workbook = GetWorkbook();
                    var sheets = workbook.Sheets?.Elements<Sheet>().ToList();
                    var sheet = sheets?.FirstOrDefault(s =>
                        s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
                    if (sheet != null)
                    {
                        var oldName = sheet.Name!.Value!;
                        // R35-1: Excel sheet names are case-insensitive and must be
                        // unique. Match the Add path's duplicate-name check
                        // (ExcelHandler.Add.Cells.cs) so renaming Sheet1→Data when a
                        // "Data" sheet already exists fails up-front rather than
                        // writing two <sheet name="Data"> entries.
                        // CONSISTENCY(sheet-name-unique)
                        if (!oldName.Equals(value, StringComparison.OrdinalIgnoreCase) &&
                            sheets!.Any(s => s != sheet &&
                                s.Name?.Value?.Equals(value, StringComparison.OrdinalIgnoreCase) == true))
                        {
                            throw new ArgumentException(
                                $"A sheet named '{value}' already exists. Sheet names must be unique.");
                        }
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

                        // Update any pivot cache definitions whose WorksheetSource
                        // references the old sheet name. Without this the pivot
                        // cache's stale sheet ref breaks Excel refresh.
                        // CONSISTENCY(sheet-rename-refs)
                        var workbookPart = _doc.WorkbookPart!;
                        foreach (var cacheDefPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
                        {
                            var wsSource = cacheDefPart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                            if (wsSource?.Sheet?.Value != null &&
                                wsSource.Sheet.Value.Equals(oldName, StringComparison.OrdinalIgnoreCase))
                            {
                                wsSource.Sheet = value;
                                cacheDefPart.PivotCacheDefinition!.Save();
                            }
                        }

                        // CONSISTENCY(sheet-rename-refs): chart series formulas
                        // (<c:f>SheetName!$A$1:$B$2</c:f>) must follow the
                        // rename or Excel reopens the file with an "external
                        // links" warning, treating the orphan SheetName!
                        // prefix as a pointer to a separate workbook. Walk
                        // every WorksheetPart's drawing → chart parts and
                        // rewrite the formula text in-place. Both quoted
                        // ('Sheet With Spaces'!) and bare (Sheet1!) forms
                        // are handled because oldRef/newRef already include
                        // the trailing '!' and quoting decision.
                        foreach (var anyWsPart in workbookPart.WorksheetParts)
                        {
                            if (anyWsPart.DrawingsPart == null) continue;
                            foreach (var chartPart in anyWsPart.DrawingsPart.ChartParts)
                            {
                                if (chartPart.ChartSpace == null) continue;
                                bool changed = false;
                                foreach (var f in chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>())
                                {
                                    if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                    {
                                        f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                        changed = true;
                                    }
                                }
                                if (changed) chartPart.ChartSpace.Save();
                            }
                        }

                        // CONSISTENCY(sheet-rename-refs): three more places
                        // carry sheet-qualified formula text in worksheet
                        // XML and need the rename cascaded:
                        //   - sparkline data range  (<xne:f>Sheet1!A1:A4</xne:f>)
                        //   - data validation list  (<x:formula1>Sheet1!A1:A3</x:formula1>)
                        //   - conditional formatting (<x:formula>Sheet1!$A$1</x:formula>)
                        // Walk each worksheet's typed descendants so we
                        // don't accidentally rewrite cell text that happens
                        // to contain the literal substring "Sheet1!".
                        foreach (var anyWsPart2 in workbookPart.WorksheetParts)
                        {
                            var wsRoot = anyWsPart2.Worksheet;
                            if (wsRoot == null) continue;
                            bool wsChanged = false;
                            foreach (var f in wsRoot.Descendants<DocumentFormat.OpenXml.Office.Excel.Formula>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            foreach (var f in wsRoot.Descendants<Formula1>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            foreach (var f in wsRoot.Descendants<Formula2>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            foreach (var f in wsRoot.Descendants<Formula>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            // Internal hyperlinks: <x:hyperlink ref="A1"
                            // location="SheetName!A1"/>. Update the
                            // location attribute when it points at the
                            // renamed sheet.
                            foreach (var hl in wsRoot.Descendants<Hyperlink>())
                            {
                                var loc = hl.Location?.Value;
                                if (loc != null && loc.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    hl.Location = loc.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            if (wsChanged) wsRoot.Save();
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

                        // R18-B3: freeze=A1 means "no freeze". Emitting a <pane> with
                        // no xSplit/ySplit produces invalid OOXML (Excel repairs on
                        // open). Treat A1 as a no-op after clearing the existing pane.
                        if (colSplit <= 0 && rowSplit <= 0)
                            break;

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
                    // Sheet-level merge: value is the range(s) to merge (e.g., "A1:A3" or
                    // "A1:D1,B3:B5" for multiple ranges).
                    // R2-1: Split comma-separated ranges into separate <mergeCell> elements;
                    // Excel rejects a single <mergeCell ref="A1:D1,B3:B5"/>.
                    var mergeCells = ws.GetFirstChild<MergeCells>();
                    if (mergeCells == null)
                    {
                        mergeCells = new MergeCells();
                        ws.AppendChild(mergeCells);
                    }
                    foreach (var part in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                        InsertMergeCellChecked(mergeCells, part.ToUpperInvariant());
                    mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
                    break;
                }
                case "autofilter":
                {
                    // Set or remove AutoFilter (like POI's XSSFSheet.setAutoFilter)
                    var existingAf = ws.GetFirstChild<AutoFilter>();
                    var trimmed = (value ?? "").Trim();
                    var lower = trimmed.ToLowerInvariant();
                    if (string.IsNullOrEmpty(trimmed) || lower is "none" or "false" or "0" or "no" or "off")
                    {
                        existingAf?.Remove();
                    }
                    else if (lower is "true" or "1" or "yes" or "on")
                    {
                        // Reject bare bool — autoFilter requires an explicit range. Otherwise
                        // we'd write Reference="TRUE" as raw text and Get would return "TRUE",
                        // which is invalid OOXML and confuses round-trip. Mirrors Add's
                        // "AutoFilter requires 'range' property" rule.
                        throw new ArgumentException(
                            "autoFilter requires an explicit range (e.g. 'A1:F100'). " +
                            "Use 'false'/'none' to remove an existing autoFilter.");
                    }
                    else
                    {
                        if (existingAf != null)
                        {
                            existingAf.Reference = trimmed.ToUpperInvariant();
                        }
                        else
                        {
                            var af = new AutoFilter { Reference = trimmed.ToUpperInvariant() };
                            var sheetData = ws.GetFirstChild<SheetData>();
                            if (sheetData != null)
                                sheetData.InsertAfterSelf(af);
                            else
                                ws.AppendChild(af);
                        }
                    }
                    break;
                }
                case "autofit":
                {
                    if (ParseHelpers.IsTruthy(value))
                        AutoFitAllColumns(worksheet);
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
                    var zoomVal = ParseHelpers.SafeParseUint(value, "zoom");
                    if (zoomVal < 10 || zoomVal > 400)
                        throw new ArgumentException($"zoom must be between 10 and 400 (got {zoomVal})");
                    sheetView.ZoomScale = zoomVal;
                    sheetView.ZoomScaleNormal = sheetView.ZoomScale;
                    break;
                }
                case "showgridlines" or "gridlines":
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
                    sheetView.ShowGridLines = ParseHelpers.IsTruthy(value);
                    break;
                }
                case "showrowcolheaders" or "showheaders" or "rowcolheaders" or "headings":
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
                    sheetView.ShowRowColHeaders = ParseHelpers.IsTruthy(value);
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
                        // CONSISTENCY(scheme-color): accept scheme-color names
                        // ("accent1"-"accent6", "lt1", "dk1", ...) by mapping
                        // them to TabColor.Theme index. Otherwise fall back to
                        // the numeric color parser for hex/named/rgb() inputs.
                        var themeIndex = ExcelSchemeColorNameToThemeIndex(value);
                        if (themeIndex.HasValue)
                        {
                            sheetPr.AppendChild(new TabColor { Theme = (UInt32Value)themeIndex.Value });
                        }
                        else
                        {
                            var colorHex = OfficeCli.Core.ParseHelpers.NormalizeArgbColor(value);
                            sheetPr.AppendChild(new TabColor { Rgb = new HexBinaryValue(colorHex) });
                        }
                    }
                    break;
                }

                case "hidden":
                case "visibility":
                {
                    // Sheet visibility lives on the workbook-level <sheet> element,
                    // not on the worksheet. Three-state: visible / hidden / veryHidden.
                    var wbSheets = GetWorkbook().GetFirstChild<Sheets>();
                    var wbSheet = wbSheets?.Elements<Sheet>()
                        .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
                    if (wbSheet != null)
                    {
                        var v = (value ?? "").Trim();
                        var keyLower = key.ToLowerInvariant();
                        if (v.Equals("veryHidden", StringComparison.OrdinalIgnoreCase)
                            || v.Equals("very", StringComparison.OrdinalIgnoreCase)
                            || v.Equals("veryhidden", StringComparison.OrdinalIgnoreCase))
                        {
                            wbSheet.State = SheetStateValues.VeryHidden;
                        }
                        else if (v.Equals("hidden", StringComparison.OrdinalIgnoreCase)
                            || (keyLower == "hidden" && ParseHelpers.IsTruthy(v)))
                        {
                            wbSheet.State = SheetStateValues.Hidden;
                        }
                        else if (v.Equals("visible", StringComparison.OrdinalIgnoreCase)
                            || (keyLower == "hidden" && !ParseHelpers.IsTruthy(v))
                            || (keyLower == "visibility" && (string.IsNullOrEmpty(v) || v.Equals("none", StringComparison.OrdinalIgnoreCase))))
                        {
                            wbSheet.State = null;
                        }
                        else
                        {
                            // Unknown value — fall back to truthiness on hidden semantics
                            wbSheet.State = ParseHelpers.IsTruthy(v) ? SheetStateValues.Hidden : null;
                        }
                        GetWorkbook().Save();
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
                // CONSISTENCY(range-action): sort is a region action like merge.
                // Sheet-level path auto-detects the full used range; explicit ranges
                // go through SetRange → SortRangeRows. Keep both entry points in
                // sync. See CLAUDE.md "Consistency > Robustness".
                case "sort":
                {
                    // R7-3: remove ALL sortState children (malformed files may
                    // carry more than one; GetFirstChild leaves stragglers).
                    foreach (var __ss in ws.Descendants<SortState>().ToList()) __ss.Remove();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        break;

                    var sd = ws.GetFirstChild<SheetData>();
                    if (sd == null) sd = ws.AppendChild(new SheetData());
                    var rows = sd.Elements<Row>().ToList();
                    // R12-2: DO NOT early-return on empty sheet here. Empty sheet + invalid
                    // sort spec (e.g. "XFE asc", "AAAA asc", "sort=asc") used to silently
                    // succeed because we bailed before spec validation. Always dispatch into
                    // SortRangeRows so it validates the spec first; if spec is valid and there
                    // is no data, it no-ops cleanly via its existing dataStartRow > row2 guard.
                    int maxCol = 1;
                    foreach (var r in rows)
                        foreach (var c in r.Elements<Cell>())
                        {
                            var cref = c.CellReference?.Value;
                            if (cref == null) continue;
                            maxCol = Math.Max(maxCol, ColumnNameToIndex(ParseCellReference(cref).Column));
                        }
                    int minRowIdx = rows.Count == 0 ? 1 : (int)rows.Min(r => r.RowIndex?.Value ?? 1u);
                    int maxRowIdx = rows.Count == 0 ? 1 : (int)rows.Max(r => r.RowIndex?.Value ?? 1u);

                    // CONSISTENCY(sort-header-default): sortHeader defaults to false
                    // (row 1 participates in the reorder). This matches our general
                    // "caller states intent explicitly" rule and is documented in help.
                    // R4-D1 and R7-4 both proposed auto-detecting headers (type-mismatch
                    // heuristic, first-row-is-string warning). Rejected: heuristic
                    // warnings ship false positives on legitimately-heterogeneous
                    // row-1 data and are spammy in pipelines. Future revisit: make
                    // sortHeader default=true project-wide as a breaking change,
                    // documented in release notes — do NOT add a per-call warning.
                    bool sortHeader = properties.TryGetValue("sortheader", out var shv) && IsTruthy(shv);
                    SortRangeRows(worksheet, 1, minRowIdx, maxCol, maxRowIdx, value, sortHeader);
                    DeleteCalcChainIfPresent();
                    break;
                }
                case "sortheader":
                    // consumed by "sort" case above; ignore silently here so it doesn't show unsupported
                    break;

                default:
                    unsupported.Add(unsupported.Count == 0
                        ? $"{key} (valid sheet props: name, freeze, zoom, showGridLines, showRowColHeaders, tabcolor, autofilter, hidden, merge, protect, password, printarea, orientation, papersize, fittopage, header, footer, sort, sortHeader)"
                        : key);
                    break;
            }
        }

        SaveWorksheet(worksheet);
        return unsupported;
    }

    // ==================== Range Set (merge/unmerge) ====================

    private List<string> SetRange(WorksheetPart worksheet, string rangeRef, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        // Separate range-level props from cell-level props
        var cellProps = new Dictionary<string, string>();
        // CONSISTENCY(range-action): sort/sortHeader are consumed together as a
        // range action (see sheet-level dispatch). If sort is present, apply it
        // after cell-level props are processed.
        string? sortSpec = null;
        bool sortHeader = false;
        // R4-4: reject merge+sort combo up front. SortRangeRows rejects any range
        // containing merged cells, but if merge is applied first in this same call
        // the merge write succeeds, then sort throws, leaving the file in a half-
        // written state. Fail fast before touching the document.
        bool hasMerge = false;
        bool hasSort = false;
        foreach (var (k, _) in properties)
        {
            var kl = k.ToLowerInvariant();
            if (kl == "merge") hasMerge = true;
            else if (kl == "sort") hasSort = true;
        }
        if (hasMerge && hasSort)
            throw new ArgumentException(
                "Cannot apply 'merge' and 'sort' in the same call. Sort rejects merged cells; " +
                "applying both in one call would leave the file half-written. Split into two calls.");
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "sort":
                    sortSpec = value;
                    break;
                case "sortheader":
                    sortHeader = IsTruthy(value);
                    break;
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

                        InsertMergeCellChecked(mergeCells, rangeRef);
                        mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
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
                            else
                                mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
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
                        PruneEmptyCell(cell);
                        // Only add to unsupported once (first cell)
                        if (row == startRow && colIdx == startColIdx)
                            unsupported.AddRange(cellUnsupported);
                    }
                }
            }
            catch
            {
                ws.ReplaceChild(sheetDataBackup, sheetData);
                // sheetData replaced — cached row entries for the old reference are stale
                InvalidateRowIndex();
                throw;
            }
        }

        // Apply sort after cell-level props (range-action handler)
        if (sortSpec != null)
        {
            var parts = rangeRef.Split(':');
            var (sc, sr) = ParseCellReference(parts[0]);
            var (ec, er) = ParseCellReference(parts[1]);
            SortRangeRows(worksheet, ColumnNameToIndex(sc), sr, ColumnNameToIndex(ec), er, sortSpec, sortHeader);
        }

        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return unsupported;
    }

    // ==================== Range Sort (region action) ====================

    /// <summary>
    /// Physically reorder rows in the given range by the given sort keys, then
    /// write sortState metadata. Rejects ranges that intersect merged cells.
    /// sortSpec format: "A asc, B desc" (direction optional, defaults to asc).
    /// Column addressing is column letters only (A, B, AA); column names are not supported.
    /// </summary>
    private void SortRangeRows(WorksheetPart worksheet, int col1, int row1, int col2, int row2,
        string sortSpec, bool sortHeader)
    {
        // Reject empty sort value at the range-level entry. Sheet-level "clear-sort"
        // semantics (sort="" or "none") are handled by the sheet-level dispatcher before
        // reaching here; any empty value that gets here came from a range path and is a
        // user error we should surface loudly.
        if (sortSpec == null || sortSpec.Length == 0 || string.IsNullOrWhiteSpace(sortSpec))
            throw new ArgumentException("sort value cannot be empty");
        if (sortSpec.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            // R7-3: drop every SortState, not just the first.
            var __ws0 = GetSheet(worksheet);
            foreach (var __ss in __ws0.Descendants<SortState>().ToList()) __ss.Remove();
            return;
        }

        // Normalize reversed ranges (e.g. C5:A1 -> A1:C5) so row/column scans cover
        // the intended region and sortState@ref stays well-formed (min:max).
        if (col1 > col2) (col1, col2) = (col2, col1);
        if (row1 > row2) (row1, row2) = (row2, row1);

        var ws = GetSheet(worksheet);
        var sd = ws.GetFirstChild<SheetData>();
        if (sd == null) return;

        // Reject protected sheets unless the protection explicitly allows sort.
        // Per OOXML sheetProtection, @sort defaults to true meaning "sort IS
        // protected" (i.e. blocked). Only @sort="false" exempts sort from the
        // protection and lets it run.
        var protection = ws.GetFirstChild<SheetProtection>();
        if (protection != null && (protection.Sheet?.Value ?? false))
        {
            bool sortBlocked = protection.Sort?.Value ?? true;
            if (sortBlocked)
                throw new InvalidOperationException(
                    "Cannot sort a protected sheet. Unprotect first (or set sheetProtection@sort=\"false\" to allow sorting while protected).");
        }

        // Reject malformed row layout within the sort row range: rows lacking RowIndex,
        // or duplicate RowIndex values. Both cases would cause silent data loss or silent
        // skipped rows in the sort below (RowIndex?.Value >= ... filter drops null;
        // duplicate RowIndex means two rows get mapped to the same target slot).
        // CONSISTENCY(sort-scope): only rows intersecting [row1..row2] are in scope; rows
        // outside the sort range are irrelevant to this action (same scoping rule as the
        // formula rejection below).
        // A row with missing RowIndex is always rejected — it cannot be located in any
        // range, and if it is logically within the sort window the sort filter would drop
        // it silently. That is strictly a data-corruption signal regardless of scope.
        {
            var seen = new HashSet<uint>();
            foreach (var r in sd.Elements<Row>())
            {
                if (r.RowIndex?.Value is not uint ri)
                    throw new InvalidOperationException(
                        "Cannot sort: sheet contains a <row> element without a RowIndex. File is malformed.");
                // Only rows within the sort row range matter for duplicate detection.
                if (ri < (uint)row1 || ri > (uint)row2) continue;
                if (!seen.Add(ri))
                    throw new InvalidOperationException(
                        $"Cannot sort: sheet contains duplicate <row r=\"{ri}\"> entries. File is malformed.");
            }
        }

        // Reject if any merged cell intersects sort range
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>())
            {
                var mref = mc.Reference?.Value;
                if (string.IsNullOrEmpty(mref) || !mref.Contains(':')) continue;
                var mparts = mref.Split(':');
                var (mac, mar) = ParseCellReference(mparts[0]);
                var (mbc, mbr) = ParseCellReference(mparts[1]);
                int maci = ColumnNameToIndex(mac), mbci = ColumnNameToIndex(mbc);
                bool rowsOverlap = !(mbr < row1 || mar > row2);
                bool colsOverlap = !(mbci < col1 || maci > col2);
                if (rowsOverlap && colsOverlap)
                    throw new InvalidOperationException(
                        $"Cannot sort range containing merged cells (found {mref}). Unmerge first or exclude merged cells from the sort range.");
            }
        }

        // Parse sort spec: "A asc, B desc" — default direction is asc
        var sortKeys = new List<(int ColIndex, bool Descending)>();
        foreach (var spec in sortSpec.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            var tokens = spec.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0) continue;
            // Reject trailing junk like "A asc B" instead of silently dropping the tail.
            if (tokens.Length > 2)
                throw new ArgumentException(
                    $"Invalid sort key '{spec.Trim()}': too many tokens. Expected '<col> [asc|desc]'");
            var colName = tokens[0].ToUpperInvariant();
            if (!Regex.IsMatch(colName, @"^[A-Z]+$"))
                throw new ArgumentException(
                    $"Invalid sort column '{tokens[0]}'. Expected column letters (A, B, AA). Column names are not supported; use letters.");
            // R12-3: "asc" and "desc" are direction keywords, not column letters. When a
            // user writes `sort=asc` (forgot the column) the token parses as a column
            // name and produced a misleading "outside the range" error. Reject up-front
            // with a targeted message. Applies regardless of case (Regex above already
            // upper-cased via ToUpperInvariant, so match against "ASC"/"DESC").
            if (colName == "ASC" || colName == "DESC")
                throw new ArgumentException(
                    $"Invalid sort key '{spec.Trim()}': sort key must start with a column letter, not a direction keyword ('{tokens[0]}'). Expected '<col> [asc|desc]'.");
            bool desc = tokens.Length > 1 && tokens[1].Equals("desc", StringComparison.OrdinalIgnoreCase);
            if (tokens.Length > 1 && !desc && !tokens[1].Equals("asc", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException($"Invalid sort direction '{tokens[1]}'. Expected 'asc' or 'desc'.");
            int keyColIdx = ColumnNameToIndex(colName);
            // R11-1 / R12-2: Excel's max column is XFD (16384, 3 letters). Anything
            // that parses past XFD is an invalid column:
            //   - length >= 4 (e.g. "AAAA", "Score"): almost certainly a column name
            //   - length == 3 but > XFD (e.g. "XFE", "ZZZ"): out of Excel's column space
            // Both cases used to fall through to a misleading "outside the range A:B"
            // error (especially pronounced on empty sheets where the range is A:A).
            if (keyColIdx > 16384)
                throw new ArgumentException(
                    $"Invalid sort column '{tokens[0]}'. Column names are not supported; use column letters (A, B, AA, up to XFD).");
            // Key column must lie within the sort range, otherwise the sort is silently
            // a no-op and writes a malformed sortCondition ref.
            if (keyColIdx < col1 || keyColIdx > col2)
                throw new ArgumentException(
                    $"Sort column {colName} is outside the range {IndexToColumnName(col1)}:{IndexToColumnName(col2)}");
            sortKeys.Add((keyColIdx, desc));
        }
        if (sortKeys.Count == 0) return;

        int dataStartRow = sortHeader ? row1 + 1 : row1;
        // R6-2: a sort that can't reorder anything (empty data region, or a
        // single data row) is a no-op. Writing sortState in those cases makes
        // Excel render a bogus sort indicator on a range that was never sorted.
        // Skip the metadata entirely rather than lying about having sorted.
        if (dataStartRow > row2)
        {
            return;
        }

        var rowsInRange = sd.Elements<Row>()
            .Where(r => r.RowIndex?.Value >= (uint)dataStartRow && r.RowIndex?.Value <= (uint)row2)
            .ToList();
        if (rowsInRange.Count <= 1)
        {
            return;
        }

        // CONSISTENCY(sort-scope): formula rejection only applies to cells INSIDE the sort
        // column range. A formula in a cell outside [col1..col2] is untouched by sort
        // (its row may be reordered, but the formula text and its refs stay intact).
        // Helper: test whether a cell's column lies within the sort column range.
        // Name is column-specific: row containment is implied by caller (we iterate
        // only rowsInRange).
        bool CellColumnInSortRange(Cell c)
        {
            var cref = c.CellReference?.Value;
            if (cref == null) return false;
            var (cc, _) = ParseCellReference(cref);
            int ci = ColumnNameToIndex(cc);
            return ci >= col1 && ci <= col2;
        }

        // Reject if any cell in the sort column range carries a shared formula group —
        // sort would corrupt the ref anchor.
        foreach (var r in rowsInRange)
            foreach (var c in r.Elements<Cell>())
                if (CellColumnInSortRange(c) && c.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared)
                    throw new InvalidOperationException(
                        "Cannot sort range containing shared formulas. Rewrite them as per-cell formulas first.");

        // CONSISTENCY(sort-rejects-formulas): same shape as the shared-formula reject above.
        // Sort rewrites each cell's CellReference to the new row index, but the formula text
        // (e.g. "=A2+1000") still encodes the *old* relative addresses. After sort, Excel
        // recalculates against the rewritten ref and silently produces wrong values — a
        // data-corruption bug. A full fix would require parsing every formula and rewriting
        // relative row numbers per the row's new position (handling A1 / $A$1 / A$1 / $A1 /
        // A:B / Sheet!A1 / named ranges), which is high risk for partial-correctness
        // regressions. Until that lands, refuse sort when any data row carries a formula.
        // Known limitation: this does NOT catch formulas *outside* the sort range that
        // reference cells *inside* it; those will also go stale on sort. Same scope as the
        // shared-formula check above (per-row scan only).
        foreach (var r in rowsInRange)
            foreach (var c in r.Elements<Cell>())
                if (CellColumnInSortRange(c) && c.CellFormula != null)
                    throw new InvalidOperationException(
                        $"Cannot sort range containing formulas (cell {c.CellReference?.Value}). " +
                        "Sort would rewrite cell references but leave formula text encoding the old row " +
                        "numbers, silently corrupting results. Rewrite formulas as literal values first " +
                        "(or evaluate and paste-as-values) before sorting.");

        // Materialize sort keys once (O(rows × keys × cells) → O(rows × keys))
        var keyed = rowsInRange.Select(r =>
        {
            var keys = new (int Rank, double NumVal, string StrVal)[sortKeys.Count];
            for (int k = 0; k < sortKeys.Count; k++)
                keys[k] = ParseSortValue(GetCellRawSortValueString(r, sortKeys[k].ColIndex));
            return (Row: r, Keys: keys);
        }).ToList();

        // Stable multi-key sort: first key primary, rest tiebreakers
        IOrderedEnumerable<(Row Row, (int Rank, double NumVal, string StrVal)[] Keys)>? ordered = null;
        for (int i = 0; i < sortKeys.Count; i++)
        {
            int idx = i;
            bool desc = sortKeys[i].Descending;
            if (ordered == null)
            {
                ordered = keyed.OrderBy(x => x.Keys[idx].Rank);
            }
            else
            {
                ordered = ordered.ThenBy(x => x.Keys[idx].Rank);
            }
            // R7-1: use case-insensitive comparer to match Excel's default sort
            // behavior. sortState defaults caseSensitive=false, so the physical
            // order must agree with that metadata declaration. Swapping to
            // OrdinalIgnoreCase also matches Excel's user-visible default.
            ordered = desc
                ? ordered.ThenByDescending(x => x.Keys[idx].NumVal)
                         .ThenByDescending(x => x.Keys[idx].StrVal, StringComparer.OrdinalIgnoreCase)
                : ordered.ThenBy(x => x.Keys[idx].NumVal)
                         .ThenBy(x => x.Keys[idx].StrVal, StringComparer.OrdinalIgnoreCase);
        }
        var sortedRows = ordered!.Select(x => x.Row).ToList();

        // The sorted slots must be assigned by ascending row index; SheetData document
        // order is not guaranteed to be ascending (malformed files, or legitimate writer
        // output), so rely on RowIndex values rather than List position.
        var originalIndices = rowsInRange.Select(r => r.RowIndex!.Value).OrderBy(v => v).ToList();

        // R4-1/2/3: capture old→new row mapping BEFORE mutating row indices so we can
        // rewrite sidecar metadata refs (hyperlinks, comments, dataValidations) that
        // encode absolute cell refs and would otherwise still point at the old rows.
        // Key = old row index (from the row object as it existed pre-sort); Value = new
        // row index it lands on post-sort.
        var oldToNewRow = new Dictionary<uint, uint>(sortedRows.Count);
        for (int i = 0; i < sortedRows.Count; i++)
        {
            var oldIdx = sortedRows[i].RowIndex!.Value;
            var newIdx = originalIndices[i];
            oldToNewRow[oldIdx] = newIdx;
        }

        // Detach from SheetData, invalidate row-index cache
        foreach (var r in rowsInRange) r.Remove();
        InvalidateRowIndex(sd);

        // Rewrite row index + cell refs on sorted rows
        for (int i = 0; i < sortedRows.Count; i++)
        {
            var newIdx = originalIndices[i];
            var r = sortedRows[i];
            r.RowIndex = newIdx;
            foreach (var cell in r.Elements<Cell>())
            {
                var cref = cell.CellReference?.Value;
                if (cref == null) continue;
                var (cc, _) = ParseCellReference(cref);
                cell.CellReference = $"{cc}{newIdx}";
            }
        }

        // R4-1/2/3: rewrite sidecar metadata refs that live outside <sheetData> but
        // encode cell addresses. Only refs pointing into the sort rectangle are
        // rewritten; refs outside are untouched. See CLAUDE.md "Consistency > Robustness"
        // — same philosophy as formula rejection: we do not attempt to rewrite refs
        // that cross the sort boundary (e.g. dataValidation sqref spanning A1:A100 when
        // only A2:A5 sort) because that would require partial-region splitting; instead
        // the cell-anchored model covers the common case and leaves other cases intact.
        RewriteSidecarRefsAfterSort(worksheet, col1, row1, col2, row2, oldToNewRow);

        // Reinsert in sorted order, preserving rows outside the data range
        var beforeRow = sd.Elements<Row>().LastOrDefault(r => r.RowIndex?.Value < (uint)dataStartRow);
        OpenXmlElement insertAfter = beforeRow ?? (OpenXmlElement)sd;
        foreach (var r in sortedRows)
        {
            if (insertAfter == sd) sd.InsertAt(r, 0);
            else insertAfter.InsertAfterSelf(r);
            insertAfter = r;
        }
        InvalidateRowIndex(sd);

        WriteSortState(ws, col1, row1, col2, row2, sortKeys);
    }

    /// <summary>Write sortState metadata. sortState@ref = full range; sortCondition@ref = key column within range.</summary>
    private static void WriteSortState(Worksheet ws, int col1, int row1, int col2, int row2,
        List<(int ColIndex, bool Descending)> sortKeys)
    {
        // R7-3: drop every SortState, not just the first (malformed files may
        // carry duplicates). GetFirstChild would leave the tail behind and the
        // newly-appended state would become the 2nd/3rd, still ambiguous.
        foreach (var __ss in ws.Descendants<SortState>().ToList()) __ss.Remove();
        var fullRef = $"{IndexToColumnName(col1)}{row1}:{IndexToColumnName(col2)}{row2}";
        var ss = new SortState { Reference = fullRef };
        foreach (var (colIdx, desc) in sortKeys)
        {
            var keyRef = $"{IndexToColumnName(colIdx)}{row1}:{IndexToColumnName(colIdx)}{row2}";
            var sc = new SortCondition { Reference = keyRef };
            if (desc) sc.Descending = true;
            ss.AppendChild(sc);
        }
        // Honor OOXML CT_Worksheet schema order. Per ECMA-376 the child sequence that
        // matters here is:
        //   sheetData → sheetCalcPr → sheetProtection → protectedRanges → scenarios
        //     → autoFilter → sortState → dataConsolidate → customSheetViews → mergeCells
        //     → phoneticPr → conditionalFormatting → dataValidations → hyperlinks → ...
        // So sortState must be inserted AFTER the latest present predecessor and BEFORE
        // any later element (mergeCells, hyperlinks, conditionalFormatting, etc.). The
        // previous fallback `sheetData.InsertAfterSelf` placed sortState before mergeCells
        // which violates the schema and is rejected by strict validators.
        var anchor = (OpenXmlElement?)ws.GetFirstChild<AutoFilter>()
            ?? (OpenXmlElement?)ws.GetFirstChild<Scenarios>()
            ?? (OpenXmlElement?)ws.GetFirstChild<ProtectedRanges>()
            ?? (OpenXmlElement?)ws.GetFirstChild<SheetProtection>()
            ?? (OpenXmlElement?)ws.GetFirstChild<SheetCalculationProperties>()
            ?? (OpenXmlElement?)ws.GetFirstChild<SheetData>();
        if (anchor != null)
            anchor.InsertAfterSelf(ss);
        else
            ws.AppendChild(ss);
    }

    /// <summary>
    /// R4-1/2/3: remap sidecar metadata cell refs after a sort. Rewrites any
    /// hyperlink/comment/dataValidation reference that anchors on a single cell
    /// inside the sort rectangle (col1..col2, row1..row2) using the old→new row
    /// mapping. Refs outside the rectangle are left alone; multi-cell refs that
    /// cross the sort boundary are also left alone (same scope-limited philosophy
    /// as the formula-rejection path — see CONSISTENCY(sort-scope)). DataValidation
    /// sqref may contain multiple space-separated tokens; each is processed
    /// independently.
    /// </summary>
    private void RewriteSidecarRefsAfterSort(WorksheetPart worksheet,
        int col1, int row1, int col2, int row2,
        Dictionary<uint, uint> oldToNewRow)
    {
        var ws = GetSheet(worksheet);

        // Helper: is a single cell ref (e.g. "A2") inside the sort rectangle?
        bool CellInRect(string cref, out string col, out uint row)
        {
            col = ""; row = 0;
            if (string.IsNullOrEmpty(cref)) return false;
            if (!System.Text.RegularExpressions.Regex.IsMatch(cref, @"^[A-Za-z]+\d+$")) return false;
            var parsed = ParseCellReference(cref);
            col = parsed.Column;
            row = (uint)parsed.Row;
            int ci = ColumnNameToIndex(col);
            return ci >= col1 && ci <= col2 && row >= (uint)row1 && row <= (uint)row2;
        }

        // ---- Hyperlinks ----
        var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
        if (hyperlinksEl != null)
        {
            foreach (var h in hyperlinksEl.Elements<Hyperlink>())
            {
                var href = h.Reference?.Value;
                if (href == null) continue;
                if (CellInRect(href, out var hc, out var hr) && oldToNewRow.TryGetValue(hr, out var newR))
                {
                    h.Reference = $"{hc.ToUpperInvariant()}{newR}";
                }
            }
        }

        // ---- Comments ----
        var commentsPart = worksheet.WorksheetCommentsPart;
        if (commentsPart?.Comments != null)
        {
            var commentList = commentsPart.Comments.GetFirstChild<CommentList>();
            if (commentList != null)
            {
                bool changed = false;
                foreach (var cmt in commentList.Elements<Comment>())
                {
                    var cref = cmt.Reference?.Value;
                    if (cref == null) continue;
                    if (CellInRect(cref, out var cc, out var cr) && oldToNewRow.TryGetValue(cr, out var newR))
                    {
                        cmt.Reference = $"{cc.ToUpperInvariant()}{newR}";
                        changed = true;
                    }
                }
                if (changed) commentsPart.Comments.Save();
            }
        }

        // ---- Threaded Comments (Excel 365) ----
        // R5-2: threadedComments<N>.xml is a separate part from legacy comments<N>.xml
        // (same storage model: per-cell <threadedComment ref="..."> entries). Rewriting
        // legacy comments but not threaded ones left 365-authored files with threaded
        // bubbles anchored to the wrong rows post-sort. Cell-anchored refs only; any
        // non-single-cell ref is left untouched (same scoping rule as legacy comments).
        foreach (var threadedPart in worksheet.WorksheetThreadedCommentsParts)
        {
            if (threadedPart?.ThreadedComments == null) continue;
            bool tcChanged = false;
            foreach (var tc in threadedPart.ThreadedComments.Elements<ThreadedCmt.ThreadedComment>())
            {
                var tref = tc.Ref?.Value;
                if (tref == null) continue;
                if (CellInRect(tref, out var tcc, out var tcr) && oldToNewRow.TryGetValue(tcr, out var newR))
                {
                    tc.Ref = $"{tcc.ToUpperInvariant()}{newR}";
                    tcChanged = true;
                }
            }
            if (tcChanged) threadedPart.ThreadedComments.Save();
        }

        // ---- DataValidations ----
        var dvs = ws.GetFirstChild<DataValidations>();
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements<DataValidation>())
            {
                var sqref = dv.SequenceOfReferences;
                if (sqref?.InnerText == null) continue;
                // sqref is a space-separated list of ref tokens; each token may be
                // a single cell (A2) or a range (A2:A5). Only single-cell tokens
                // inside the sort rectangle are remapped; multi-cell ranges are
                // left untouched (partial-rect rewrite would require splitting).
                var tokens = sqref.InnerText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                bool changed = false;
                for (int i = 0; i < tokens.Length; i++)
                {
                    var tok = tokens[i];
                    if (tok.Contains(':')) continue; // range token — skip
                    if (CellInRect(tok, out var dc, out var dr) && oldToNewRow.TryGetValue(dr, out var newR))
                    {
                        tokens[i] = $"{dc.ToUpperInvariant()}{newR}";
                        changed = true;
                    }
                }
                if (changed)
                {
                    dv.SequenceOfReferences = new ListValue<StringValue>(
                        tokens.Select(t => new StringValue(t)));
                }
            }
        }

        // ---- ProtectedRanges (R7-2) ----
        // CONSISTENCY(sort-scope): same cell-anchored scoping as dataValidations.
        // Each <protectedRange sqref="..."> carries a space-separated list of
        // ref tokens; only single-cell tokens inside the sort rectangle are
        // remapped. Multi-cell ranges are left intact (partial-rect split would
        // alter which cells are protected, same philosophy as DV/CF).
        var pranges = ws.GetFirstChild<ProtectedRanges>();
        if (pranges != null)
        {
            foreach (var pr in pranges.Elements<ProtectedRange>())
            {
                var sqref = pr.SequenceOfReferences;
                if (sqref?.InnerText == null) continue;
                var tokens = sqref.InnerText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                bool changed = false;
                for (int i = 0; i < tokens.Length; i++)
                {
                    var tok = tokens[i];
                    if (tok.Contains(':')) continue; // range token — skip
                    if (CellInRect(tok, out var pc, out var pRow) && oldToNewRow.TryGetValue(pRow, out var newR))
                    {
                        tokens[i] = $"{pc.ToUpperInvariant()}{newR}";
                        changed = true;
                    }
                }
                if (changed)
                {
                    pr.SequenceOfReferences = new ListValue<StringValue>(
                        tokens.Select(t => new StringValue(t)));
                }
            }
        }

        // ---- ConditionalFormatting (R6-1) ----
        // CONSISTENCY(sort-scope): same cell-anchored scoping as dataValidations.
        // CF sqref is a space-separated list where each token may be a single
        // cell (A2) or a range (A1:A10). Only single-cell tokens inside the sort
        // rectangle are remapped; multi-cell ranges are left untouched — a range
        // that straddles reordered rows cannot be split into the new set of rows
        // without changing which cells the rule covers, so we preserve the
        // authored range verbatim (same partial-rect rule as dataValidations).
        foreach (var cf in ws.Elements<ConditionalFormatting>())
        {
            var sqref = cf.SequenceOfReferences;
            if (sqref?.InnerText == null) continue;
            var tokens = sqref.InnerText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            bool changed = false;
            for (int i = 0; i < tokens.Length; i++)
            {
                var tok = tokens[i];
                if (tok.Contains(':')) continue; // range token — skip
                if (CellInRect(tok, out var cc, out var cr) && oldToNewRow.TryGetValue(cr, out var newR))
                {
                    tokens[i] = $"{cc.ToUpperInvariant()}{newR}";
                    changed = true;
                }
            }
            if (changed)
            {
                cf.SequenceOfReferences = new ListValue<StringValue>(
                    tokens.Select(t => new StringValue(t)));
            }
        }

        // ---- Drawing anchors (R6-4) ----
        // CONSISTENCY(sort-scope): same cell-anchored scoping as dataValidations/CF.
        // Drawing anchors (xdr:twoCellAnchor/xdr:oneCellAnchor) pin shapes, pictures,
        // and charts to a (col,row) pair via xdr:from (and xdr:to for twoCell). RowId
        // is 0-indexed in OOXML, so worksheet row N ↔ RowId = N-1. Before R6-4 the
        // sort path rewrote cell-level sidecars but left drawing RowIds untouched,
        // which dragged pictures off their original anchor row after a reorder.
        //
        // Scoping rule (partial-rect): for TwoCellAnchor both From and To rows must
        // fall inside the sort rectangle for the anchor to move. If only one end is
        // inside, preserve the authored anchor (splitting a rectangle across
        // reordered rows would change which cells the drawing visually covers).
        // OneCellAnchor has only From — remap iff From is inside.
        // Columns aren't affected by row sort, so ColId is never rewritten.
        var drawingsPart = worksheet.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing != null)
        {
            bool drawingChanged = false;
            bool RowInSortRect(uint oneBasedRow) =>
                oneBasedRow >= (uint)row1 && oneBasedRow <= (uint)row2;

            // TwoCellAnchor: remap only if both endpoints' rows are in sort rect.
            foreach (var anchor in drawingsPart.WorksheetDrawing.Elements<XDR.TwoCellAnchor>())
            {
                var from = anchor.FromMarker;
                var to = anchor.ToMarker;
                if (from?.RowId?.Text == null || to?.RowId?.Text == null) continue;
                if (!uint.TryParse(from.RowId.Text, out uint fromRow0)) continue;
                if (!uint.TryParse(to.RowId.Text, out uint toRow0)) continue;
                uint fromRow1 = fromRow0 + 1;
                uint toRow1 = toRow0 + 1;
                if (!RowInSortRect(fromRow1) || !RowInSortRect(toRow1)) continue;
                if (!oldToNewRow.TryGetValue(fromRow1, out uint newFrom1)) continue;
                if (!oldToNewRow.TryGetValue(toRow1, out uint newTo1)) continue;
                from.RowId = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId(
                    (newFrom1 - 1).ToString());
                to.RowId = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId(
                    (newTo1 - 1).ToString());
                drawingChanged = true;
            }

            // OneCellAnchor: remap iff From is in sort rect.
            foreach (var anchor in drawingsPart.WorksheetDrawing.Elements<XDR.OneCellAnchor>())
            {
                var from = anchor.FromMarker;
                if (from?.RowId?.Text == null) continue;
                if (!uint.TryParse(from.RowId.Text, out uint fromRow0)) continue;
                uint fromRow1 = fromRow0 + 1;
                if (!RowInSortRect(fromRow1)) continue;
                if (!oldToNewRow.TryGetValue(fromRow1, out uint newFrom1)) continue;
                from.RowId = new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId(
                    (newFrom1 - 1).ToString());
                drawingChanged = true;
            }

            if (drawingChanged) drawingsPart.WorksheetDrawing.Save();
        }
    }

    /// <summary>Raw cell value for sorting: resolves SharedString/InlineString, skips number formatting. Precise column-letter match (no prefix bug).</summary>
    private string GetCellRawSortValueString(Row row, int colIdx)
    {
        var colLetter = IndexToColumnName(colIdx);
        foreach (var cell in row.Elements<Cell>())
        {
            var cref = cell.CellReference?.Value;
            if (cref == null) continue;
            var (cc, _) = ParseCellReference(cref);
            if (!cc.Equals(colLetter, StringComparison.OrdinalIgnoreCase)) continue;

            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sst?.SharedStringTable != null && int.TryParse(cell.CellValue?.Text, out int idx))
                    return sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx)?.InnerText ?? "";
                return "";
            }
            if (cell.DataType?.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? "";
            return cell.CellValue?.Text ?? "";
        }
        return "";
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
                    col.Width = ParseColWidthChars(value);
                    col.CustomWidth = true;
                    break;
                case "hidden":
                    col.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "outline" or "outlinelevel" or "group":
                    // DEFERRED(xlsx/row-height-validation) RC2: Excel outline level max is 7.
                    if (!byte.TryParse(value, out var colOutline) || colOutline > 7)
                        throw new ArgumentException($"Invalid 'outline' value: '{value}'. Expected an integer 0-7 (outline/group level).");
                    col.OutlineLevel = colOutline;
                    break;
                case "collapsed":
                    col.Collapsed = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "autofit":
                    if (ParseHelpers.IsTruthy(value))
                    {
                        var autoFitWidth = CalculateAutoFitWidth(worksheet, colName);
                        col.Width = autoFitWidth;
                        col.CustomWidth = true;
                    }
                    break;
                default:
                    // Long-tail Column attribute (CT_Col attrs beyond width/
                    // hidden/outlineLevel/collapsed/customWidth — e.g. style,
                    // bestFit, phonetic). Set as raw OOXML attribute. Symmetric
                    // with the column Get reader which now uses
                    // FillUnknownAttrProps for unrecognized attrs. Preserve
                    // original case (OOXML attribute names are case-sensitive).
                    col.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", key, "", value));
                    break;
            }
        }

        SaveWorksheet(worksheet);
        return unsupported;
    }

    // ==================== Column Auto-Fit ====================

    private double CalculateAutoFitWidth(WorksheetPart worksheet, string colName)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        var colIdx = ColumnNameToIndex(colName);
        double maxLen = 0;

        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value;
                    if (cellRef == null) continue;
                    var (cellCol, _) = ParseCellReference(cellRef);
                    if (ColumnNameToIndex(cellCol) != colIdx) continue;

                    var text = GetCellDisplayValue(cell);
                    var textWidth = ParseHelpers.EstimateTextWidthInChars(text);
                    if (textWidth > maxLen)
                        maxLen = textWidth;
                }
            }
        }

        // Approximate width: characters * 1.1 + 2 for padding, minimum 8
        return Math.Max(maxLen * 1.1 + 2, 8);
    }

    private void AutoFitAllColumns(WorksheetPart worksheet)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        // Collect all used column indices
        var usedColumns = new HashSet<int>();
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;
                var (cellCol, _) = ParseCellReference(cellRef);
                usedColumns.Add(ColumnNameToIndex(cellCol));
            }
        }

        if (usedColumns.Count == 0) return;

        var columns = ws.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            ws.InsertBefore(columns, sheetData);
        }

        foreach (var colIdx in usedColumns.OrderBy(c => c))
        {
            var colName = IndexToColumnName(colIdx);
            var width = CalculateAutoFitWidth(worksheet, colName);
            var uColIdx = (uint)colIdx;

            var col = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min?.Value <= uColIdx && c.Max?.Value >= uColIdx);
            if (col == null)
            {
                col = new Column { Min = uColIdx, Max = uColIdx, Width = width, CustomWidth = true };
                var afterCol = columns.Elements<Column>().LastOrDefault(c => (c.Min?.Value ?? 0) < uColIdx);
                if (afterCol != null)
                    afterCol.InsertAfterSelf(col);
                else
                    columns.PrependChild(col);
            }
            else
            {
                col.Width = width;
                col.CustomWidth = true;
            }
        }

        SaveWorksheet(worksheet);
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
                    row.Height = ParseRowHeightPoints(value);
                    row.CustomHeight = true;
                    break;
                case "hidden":
                    row.Hidden = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                case "outline" or "outlinelevel" or "group":
                    // DEFERRED(xlsx/row-height-validation) RC2: Excel outline level max is 7.
                    if (!byte.TryParse(value, out var outlineVal) || outlineVal > 7)
                        throw new ArgumentException($"Invalid 'outline' value: '{value}'. Expected an integer 0-7 (outline/group level).");
                    row.OutlineLevel = outlineVal;
                    break;
                case "collapsed":
                    row.Collapsed = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value == "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    // Long-tail Row attribute (CT_Row attrs beyond height/
                    // hidden/outlineLevel/collapsed — e.g. spans, style, ph,
                    // thickTop, thickBot, customFormat). Symmetric with the
                    // row Get reader. Preserve original case.
                    row.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", key, "", value));
                    break;
            }
        }

        SaveWorksheet(worksheet);
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

        SaveWorksheet(worksheet);
        return unsupported;
    }
}
