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
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // Normalize to case-insensitive lookup so camelCase keys (e.g. minColor) match lowercase lookups
        if (properties != null && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        parentPath = NormalizeExcelPath(parentPath);
        parentPath = ResolveSheetIndexInPath(parentPath);
        switch (type.ToLowerInvariant())
        {
            case "sheet":
                var workbookPart = _doc.WorkbookPart
                    ?? throw new InvalidOperationException("Workbook not found");
                var sheets = GetWorkbook().GetFirstChild<Sheets>()
                    ?? GetWorkbook().AppendChild(new Sheets());

                var name = properties.GetValueOrDefault("name", $"Sheet{sheets.Elements<Sheet>().Count() + 1}");
                if (sheets.Elements<Sheet>().Any(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)))
                {
                    if (_initialSheetNames.Contains(name))
                    {
                        // Sheet existed when the file was opened — treat as idempotent no-op
                        return $"/{name}";
                    }
                    throw new ArgumentException($"A sheet named '{name}' already exists. Sheet names must be unique.");
                }
                var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                newWorksheetPart.Worksheet.Save();

                var sheetId = sheets.Elements<Sheet>().Any()
                    ? sheets.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1
                    : 1;
                var relId = workbookPart.GetIdOfPart(newWorksheetPart);

                var newSheet = new Sheet { Id = relId, SheetId = (uint)sheetId, Name = name };
                if (properties.TryGetValue("position", out var posStr)
                    && int.TryParse(posStr, out var pos)
                    && pos >= 0
                    && pos < sheets.Elements<Sheet>().Count())
                {
                    var refSheet = sheets.Elements<Sheet>().ElementAt(pos);
                    sheets.InsertBefore(newSheet, refSheet);
                }
                else
                {
                    sheets.AppendChild(newSheet);
                }
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

                if (needsShift)
                    DeleteCalcChainIfPresent();
                SaveWorksheet(worksheet);
                return $"/{sheetName}/row[{rowIdx}]";

            case "cell":
                var cellSegments = parentPath.TrimStart('/').Split('/', 2);
                var cellSheetName = cellSegments[0];
                var cellWorksheet = FindWorksheet(cellSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cellSheetName}");
                var cellSheetData = GetSheet(cellWorksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(cellWorksheet).AppendChild(new SheetData());

                // R7-1: if path tail is a cell-ref (e.g. /Sheet1/Z99), treat it
                // as the target address — equivalent to --prop ref=Z99. Parity
                // with the `comment` case below which already does this.
                string? cellRefFromPath = null;
                if (cellSegments.Length > 1 && Regex.IsMatch(cellSegments[1], @"^[A-Z]+\d+$", RegexOptions.IgnoreCase))
                    cellRefFromPath = cellSegments[1].ToUpperInvariant();

                string cellRef;
                if (properties.ContainsKey("ref"))
                {
                    cellRef = properties["ref"];
                    if (cellRefFromPath != null && !cellRefFromPath.Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                        Console.Error.WriteLine($"warning: path tail '{cellRefFromPath}' does not match --prop ref='{cellRef}'; using ref='{cellRef}'.");
                }
                else if (properties.ContainsKey("address"))
                {
                    cellRef = properties["address"];
                    if (cellRefFromPath != null && !cellRefFromPath.Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                        Console.Error.WriteLine($"warning: path tail '{cellRefFromPath}' does not match --prop address='{cellRef}'; using address='{cellRef}'.");
                }
                else if (cellRefFromPath != null)
                {
                    cellRef = cellRefFromPath;
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
                    // Auto-detect formula: value starting with '=' is treated as formula
                    if (value.StartsWith('=') && value.Length > 1)
                    {
                        cell.CellFormula = new CellFormula(value.TrimStart('='));
                        cell.CellValue = null;
                    }
                    else
                    {
                        // CONSISTENCY(formula-stale): writing a literal value must
                        // clear any prior CellFormula on the same cell. Otherwise
                        // the old formula re-evaluates on open / in html preview
                        // and overrides the literal the caller just set.
                        cell.CellFormula = null;
                        // R2-2: strip XML-illegal chars (e.g. U+0000) from the cell
                        // value before it gets serialized to sheet1.xml. Without
                        // this, a NUL byte from upstream data would crash every
                        // downstream save (including the pivot cache write).
                        var safeValue = OfficeCli.Core.PivotTableHelper.SanitizeXmlText(value);
                        cell.CellValue = new CellValue(safeValue);
                        if (!double.TryParse(safeValue, out _))
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                }
                if (properties.TryGetValue("formula", out var formula))
                {
                    // Strip a leading '=' (formula-bar copy) and reject
                    // literal `{...}` array-formula wrapping — users must use
                    // the dedicated `arrayformula=` prop for that, since
                    // `<x:f>{=...}</x:f>` causes Excel to reject the file.
                    var fTrim = formula.TrimStart('=').Trim();
                    if (fTrim.StartsWith("{") && fTrim.EndsWith("}"))
                        throw new ArgumentException("Literal braces '{...}' around a formula create an Excel-rejected file. Use --prop arrayformula=... (without braces) to declare a CSE array formula.");
                    cell.CellFormula = new CellFormula(fTrim);
                    cell.CellValue = null;
                }
                // CE1: allow `runs=<json>` without an explicit `type=richtext`.
                if (!properties.ContainsKey("type") && properties.ContainsKey("runs"))
                    properties["type"] = "richtext";
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

                        // Gather runs from either: (a) runs=<JSON array> or
                        // (b) legacy run1=, run2=, ... mini-spec syntax.
                        // CE1 fix: `runs=[{"text":"Hello","bold":true,...},...]`
                        // is now the preferred, documented form.
                        var gatheredRuns = new List<(string text, Dictionary<string, string> props)>();
                        if (properties.TryGetValue("runs", out var runsJson) && !string.IsNullOrWhiteSpace(runsJson))
                        {
                            try
                            {
                                using var jdoc = System.Text.Json.JsonDocument.Parse(runsJson);
                                if (jdoc.RootElement.ValueKind != System.Text.Json.JsonValueKind.Array)
                                    throw new ArgumentException("'runs' must be a JSON array of run objects.");
                                foreach (var el in jdoc.RootElement.EnumerateArray())
                                {
                                    if (el.ValueKind != System.Text.Json.JsonValueKind.Object)
                                        throw new ArgumentException("Each run in 'runs' must be a JSON object.");
                                    string text = "";
                                    var pd = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                    foreach (var p in el.EnumerateObject())
                                    {
                                        var sv = p.Value.ValueKind switch
                                        {
                                            System.Text.Json.JsonValueKind.True => "true",
                                            System.Text.Json.JsonValueKind.False => "false",
                                            System.Text.Json.JsonValueKind.Null => "",
                                            System.Text.Json.JsonValueKind.Number => p.Value.GetRawText(),
                                            _ => p.Value.GetString() ?? ""
                                        };
                                        if (p.NameEquals("text")) text = sv;
                                        else pd[p.Name] = sv;
                                    }
                                    gatheredRuns.Add((text, pd));
                                }
                            }
                            catch (System.Text.Json.JsonException jex)
                            {
                                throw new ArgumentException($"Invalid JSON for 'runs': {jex.Message}");
                            }
                        }
                        else
                        {
                            // Legacy path: run1=text:prop=val;prop=val, run2=...
                            var runKeys = properties.Keys
                                .Where(k => k.StartsWith("run", StringComparison.OrdinalIgnoreCase) && k.Length > 3 &&
                                            int.TryParse(k.AsSpan(3), out _))
                                .OrderBy(k => int.Parse(k.AsSpan(3).ToString()))
                                .ToList();
                            foreach (var runKey in runKeys)
                            {
                                var runVal = properties[runKey];
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
                                var pd = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                foreach (var prop in runProps)
                                {
                                    var eqIdx = prop.IndexOf('=');
                                    if (eqIdx < 0) continue;
                                    pd[prop[..eqIdx].Trim()] = prop[(eqIdx + 1)..].Trim();
                                }
                                gatheredRuns.Add((runText, pd));
                            }
                        }

                        foreach (var (runText, pd) in gatheredRuns)
                        {
                            var run = new Run();
                            var rp = new RunProperties();
                            foreach (var kv in pd)
                            {
                                var pKey = kv.Key.ToLowerInvariant();
                                var pVal = kv.Value;
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
                                    case "superscript" when ParseHelpers.IsTruthy(pVal):
                                        rp.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript });
                                        break;
                                    case "subscript" when ParseHelpers.IsTruthy(pVal):
                                        rp.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript });
                                        break;
                                    case "size" or "fontsize":
                                        if (double.TryParse(pVal.TrimEnd('p', 't'), out var sz))
                                            rp.AppendChild(new FontSize { Val = sz });
                                        break;
                                    case "color":
                                        rp.AppendChild(new Color { Rgb = new HexBinaryValue(ParseHelpers.NormalizeArgbColor(pVal)) });
                                        break;
                                    case "font" or "fontname" or "name":
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
                            // CONSISTENCY(cell-type-parity): Bug #4 — Add must accept
                            // the same type tokens as Set (ExcelHandler.Set.cs line 1105).
                            // Dates are stored as numeric OADate, so DataType stays null;
                            // the date-shaped cell value serialization and default
                            // numberformat are applied right after this switch.
                            "date" => null,
                            // CE16 — accept `type=error value="#N/A"|"#DIV/0!"|...` →
                            // emits <x:c t="e"><x:v>#N/A</x:v></x:c>. Standard
                            // Excel error tokens: #N/A, #DIV/0!, #REF!, #NAME?,
                            // #NULL!, #NUM!, #VALUE!, #GETTING_DATA.
                            "error" or "err" => new EnumValue<CellValues>(CellValues.Error),
                            _ => throw new ArgumentException($"Invalid cell 'type' value '{cellType}'. Valid types: string, number, boolean, date, error, richtext.")
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
                        // CONSISTENCY(cell-type-parity): mirror Set's value auto-detect
                        // path (ExcelHandler.Set.cs lines 1025-1033) — parse the cell
                        // value as an ISO date and write it back as an OADate double so
                        // Excel renders it as a real date instead of a literal string.
                        if (cellType.Equals("date", StringComparison.OrdinalIgnoreCase))
                        {
                            var dateText = cell.CellValue?.Text?.Trim();
                            if (!string.IsNullOrEmpty(dateText)
                                && DateTime.TryParseExact(dateText,
                                    new[] { "yyyy-MM-dd", "yyyy/MM/dd", "yyyy-MM-dd HH:mm:ss" },
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    System.Globalization.DateTimeStyles.None, out var dt))
                            {
                                cell.CellValue = new CellValue(
                                    dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture));
                            }
                            // Apply a default date number format unless the caller
                            // already supplied one — matches Set's type=date guard.
                            if (!properties.ContainsKey("numberformat")
                                && !properties.ContainsKey("numfmt")
                                && !properties.ContainsKey("format"))
                            {
                                properties["numberformat"] = "yyyy-mm-dd";
                            }
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
                    var hl = new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id };
                    // H2: tooltip (OOXML @tooltip) — Excel surfaces it as a
                    // ScreenTip when the cell is hovered in read mode.
                    var hlTip = properties.GetValueOrDefault("tooltip")
                        ?? properties.GetValueOrDefault("screenTip")
                        ?? properties.GetValueOrDefault("screentip");
                    if (!string.IsNullOrEmpty(hlTip))
                        hl.Tooltip = hlTip;
                    hyperlinksEl.AppendChild(hl);
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
                    _dirtyStylesheet = true;
                }
                else if (properties.ContainsKey("link") && !string.IsNullOrEmpty(properties["link"]))
                {
                    // H3: give hyperlink cells the built-in "Hyperlink" cellStyle
                    // (blue + underline) when the user did not supply explicit
                    // styling — so they render as proper links in real Excel.
                    // CONSISTENCY(hyperlink-cellstyle): explicit font=/color= wins.
                    var cellWbPart = _doc.WorkbookPart
                        ?? throw new InvalidOperationException("Workbook not found");
                    var styleManager = new ExcelStyleManager(cellWbPart);
                    cell.StyleIndex = styleManager.EnsureHyperlinkCellStyle();
                    _dirtyStylesheet = true;
                }

                // CONSISTENCY(xlsx/table-autoexpand): eager post-write auto-grow
                // for tables flagged with autoExpand=true. Matches Excel's
                // "type below a table → table grows" UX.
                MaybeExpandTablesForCell(cellWorksheet, cellRef);

                DeleteCalcChainIfPresent();
                SaveWorksheet(cellWorksheet);
                return $"/{cellSheetName}/{cellRef}";

            case "namedrange" or "definedname":
            {
                // R4-4: accept `/namedrange[NAME]` path form so users don't
                // have to repeat the name in --prop name=. Path brackets take
                // precedence only when --prop name= is absent (explicit prop
                // still wins on mismatch, to keep other `/namedrange[N]` int
                // indexing semantics elsewhere in the handler usable as-is).
                var pathNrName = "";
                {
                    var mNr = System.Text.RegularExpressions.Regex.Match(
                        parentPath, @"^/namedrange\[([^\]]+)\]/?$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (mNr.Success)
                    {
                        var captured = mNr.Groups[1].Value;
                        // Only treat as a name if it is not a pure integer
                        // (preserves existing `/namedrange[1]` semantics).
                        if (!int.TryParse(captured, out _))
                            pathNrName = captured;
                    }
                }
                var nrName = properties.GetValueOrDefault("name", pathNrName);
                if (string.IsNullOrEmpty(nrName))
                    throw new ArgumentException("'name' property is required for namedrange");
                // Per OOXML §18.2.5: defined-name identifiers must start with
                // letter/underscore/backslash, contain only letter/digit/
                // underscore/period/backslash, and must not parse as a cell
                // reference. Otherwise Excel rejects the file with 0x800A03EC.
                if (!System.Text.RegularExpressions.Regex.IsMatch(nrName, @"^[A-Za-z_\\][A-Za-z0-9_\\.]*$"))
                    throw new ArgumentException($"Invalid defined-name '{nrName}': must start with a letter/underscore and contain only letters, digits, underscores, or periods (no spaces).");
                if (LooksLikeCellReference(nrName))
                    throw new ArgumentException($"Invalid defined-name '{nrName}': name parses as a cell reference; choose a different name.");
                // `refersTo` is the common Excel-documented alias for `ref`;
                // silently map it so users don't end up with an empty
                // <x:definedName/> that corrupts the file.
                var refVal = properties.GetValueOrDefault("ref",
                    properties.GetValueOrDefault("refersTo",
                        properties.GetValueOrDefault("formula", "")));
                // R7-2: per ECMA-376 §18.2.5, <x:definedName> content must NOT
                // have a leading '=' (unlike the formula-bar form in Excel UI).
                // Excel rejects the file with 0x800A03EC if '=' is present.
                if (refVal.StartsWith('='))
                    refVal = refVal.TrimStart('=');

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

                // R7-3: if the defined-name body is a formula (not just a pure
                // range reference), set fullCalcOnLoad so Excel recomputes on
                // first open — otherwise the name evaluates to 0 until the
                // user triggers a recalc.
                if (LooksLikeFormulaBody(refVal))
                {
                    var calcPr = workbook.GetFirstChild<CalculationProperties>();
                    if (calcPr == null)
                    {
                        calcPr = new CalculationProperties();
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

                // CONSISTENCY(overlap-reject): duplicate comment on the same
                // cell is ambiguous — mirror the table T4 overlap-reject
                // pattern. User must `remove comment` first to replace it.
                var cmtRefUpper = cmtRef.ToUpperInvariant();
                if (commentList.Elements<Comment>().Any(c =>
                        string.Equals(c.Reference?.Value, cmtRefUpper, StringComparison.OrdinalIgnoreCase)))
                    throw new ArgumentException(
                        $"comment already exists on {cmtRefUpper}. Remove it first before adding a new comment.");

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
                // Support user-supplied `\n` (literal two-char sequence from
                // CLI) and real LF as line breaks — Excel renders the
                // preserved newline in the comment body. Matches the shape
                // `text` behavior documented in add-shape help.
                var cmtNormalized = (cmtText ?? "").Replace("\r\n", "\n").Replace("\\n", "\n");
                comment.CommentText = new CommentText(
                    new Run(
                        BuildCommentRunProperties(properties),
                        new Text(cmtNormalized) { Space = SpaceProcessingModeValues.Preserve }
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
                    dv.Formula1 = new Formula1(NormalizeValidationFormula(dvFormula1, dv.Type?.Value));
                }

                if (properties.TryGetValue("formula2", out var dvFormula2))
                    dv.Formula2 = new Formula2(NormalizeValidationFormula(dvFormula2, dv.Type?.Value));

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

                // V6 — errorStyle: stop (default), warning, information.
                if (dvProps.TryGetValue("errorStyle", out var dvErrStyle))
                {
                    dv.ErrorStyle = dvErrStyle.ToLowerInvariant() switch
                    {
                        "stop" => DataValidationErrorStyleValues.Stop,
                        "warning" or "warn" => DataValidationErrorStyleValues.Warning,
                        "information" or "info" => DataValidationErrorStyleValues.Information,
                        _ => throw new ArgumentException(
                            $"Unknown errorStyle: {dvErrStyle}. Use: stop, warning, information")
                    };
                }

                // V7 — showDropDown / inCellDropdown. OOXML `showDropDown`
                // has INVERTED semantics: true = HIDE the in-cell arrow.
                // Expose it as `inCellDropdown` (user-friendly sense) and
                // the raw `showDropDown` (OOXML sense).
                if (dvProps.TryGetValue("inCellDropdown", out var dvInCell))
                    dv.ShowDropDown = !ParseHelpers.IsTruthy(dvInCell);
                else if (dvProps.TryGetValue("showDropDown", out var dvShowDd))
                    dv.ShowDropDown = ParseHelpers.IsTruthy(dvShowDd);

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

                // CONSISTENCY(cellref-validate): reject garbage refs (e.g. "BADREF")
                // so Excel doesn't silently open with an invalid <x:autoFilter ref="...">.
                if (!Regex.IsMatch(afRange.Trim(),
                        @"^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$",
                        RegexOptions.IgnoreCase))
                    throw new ArgumentException(
                        $"Invalid 'range' value: '{afRange}'. Expected a cell range like 'A1:F100' or 'A1'.");

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

                // AF1: per-column criteria. Syntax: criteriaN.OP=VAL where
                // N is 0-based column offset from the filter range's
                // leftmost column and OP is one of:
                //   equals, contains, gt, lt, top, blanks, nonBlanks
                // Each distinct N builds one <x:filterColumn colId="N">.
                // Previous criteria for the same N are replaced.
                var criteriaGroups = new Dictionary<uint, List<(string op, string val)>>();
                foreach (var (k, v) in properties)
                {
                    var cm = Regex.Match(k, @"^criteria(\d+)\.([A-Za-z]+)$");
                    if (!cm.Success) continue;
                    var colId = uint.Parse(cm.Groups[1].Value);
                    var op = cm.Groups[2].Value.ToLowerInvariant();
                    if (!criteriaGroups.TryGetValue(colId, out var list))
                        criteriaGroups[colId] = list = new List<(string, string)>();
                    list.Add((op, v));
                }
                // Strip any prior filterColumn entries so a re-Add is idempotent
                foreach (var fc in autoFilter.Elements<FilterColumn>().ToList())
                    fc.Remove();
                foreach (var (colId, entries) in criteriaGroups.OrderBy(kv => kv.Key))
                {
                    var filterColumn = new FilterColumn { ColumnId = colId };
                    // Dispatch by operator family. Top-N, Blanks, value-list,
                    // and dynamicFilter build dedicated child elements;
                    // text/number ops feed into <customFilters>.
                    var customEntries = new List<(FilterOperatorValues fop, string val)>();
                    bool customFilterAnd = false;
                    bool handledDedicated = false;
                    foreach (var (op, rawVal) in entries)
                    {
                        switch (op)
                        {
                            case "equals":
                                customEntries.Add((FilterOperatorValues.Equal, rawVal));
                                break;
                            case "notequals":
                                customEntries.Add((FilterOperatorValues.NotEqual, rawVal));
                                break;
                            case "contains":
                            {
                                var wild = rawVal.Contains('*') ? rawVal : $"*{rawVal}*";
                                customEntries.Add((FilterOperatorValues.Equal, wild));
                                break;
                            }
                            case "doesnotcontain":
                            {
                                var wild = rawVal.Contains('*') ? rawVal : $"*{rawVal}*";
                                customEntries.Add((FilterOperatorValues.NotEqual, wild));
                                break;
                            }
                            case "beginswith":
                            {
                                var wild = rawVal.EndsWith("*") ? rawVal : $"{rawVal}*";
                                customEntries.Add((FilterOperatorValues.Equal, wild));
                                break;
                            }
                            case "endswith":
                            {
                                var wild = rawVal.StartsWith("*") ? rawVal : $"*{rawVal}";
                                customEntries.Add((FilterOperatorValues.Equal, wild));
                                break;
                            }
                            case "gt":
                                customEntries.Add((FilterOperatorValues.GreaterThan, rawVal));
                                break;
                            case "gte":
                                customEntries.Add((FilterOperatorValues.GreaterThanOrEqual, rawVal));
                                break;
                            case "lt":
                                customEntries.Add((FilterOperatorValues.LessThan, rawVal));
                                break;
                            case "lte":
                                customEntries.Add((FilterOperatorValues.LessThanOrEqual, rawVal));
                                break;
                            case "between":
                            case "notbetween":
                            {
                                var parts = rawVal.Split(',');
                                if (parts.Length != 2)
                                    throw new ArgumentException(
                                        $"criteria{colId}.{op} requires 'lo,hi', got: '{rawVal}'");
                                var lo = parts[0].Trim();
                                var hi = parts[1].Trim();
                                if (op == "between")
                                {
                                    customEntries.Add((FilterOperatorValues.GreaterThanOrEqual, lo));
                                    customEntries.Add((FilterOperatorValues.LessThanOrEqual, hi));
                                    customFilterAnd = true;
                                }
                                else
                                {
                                    // notBetween = lt lo OR gt hi (Excel default OR)
                                    customEntries.Add((FilterOperatorValues.LessThan, lo));
                                    customEntries.Add((FilterOperatorValues.GreaterThan, hi));
                                }
                                break;
                            }
                            case "top":
                            case "toppercent":
                            case "bottom":
                            case "bottompercent":
                            {
                                if (!double.TryParse(rawVal, System.Globalization.NumberStyles.Any,
                                        System.Globalization.CultureInfo.InvariantCulture, out var topN))
                                    throw new ArgumentException(
                                        $"criteria{colId}.{op} requires a numeric value, got: '{rawVal}'");
                                filterColumn.Top10 = new Top10
                                {
                                    Top = op == "top" || op == "toppercent",
                                    Percent = op == "toppercent" || op == "bottompercent",
                                    Val = topN
                                };
                                handledDedicated = true;
                                break;
                            }
                            case "blanks":
                                if (IsTruthy(rawVal))
                                {
                                    filterColumn.Filters = new Filters { Blank = true };
                                    handledDedicated = true;
                                }
                                break;
                            case "nonblanks":
                                if (IsTruthy(rawVal))
                                {
                                    customEntries.Add((FilterOperatorValues.NotEqual, ""));
                                }
                                break;
                            case "values":
                            {
                                // Discrete value-list filter: comma-separated
                                // (split+trim empty; escape \, not supported).
                                var vals = rawVal.Split(',')
                                    .Select(s => s.Trim())
                                    .Where(s => s.Length > 0)
                                    .ToList();
                                var filters = filterColumn.Filters ?? (filterColumn.Filters = new Filters());
                                foreach (var v in vals)
                                    filters.AppendChild(new Filter { Val = v });
                                handledDedicated = true;
                                break;
                            }
                            case "dynamic":
                            {
                                var dyn = new DynamicFilter
                                {
                                    Type = new EnumValue<DynamicFilterValues>(new DynamicFilterValues(rawVal))
                                };
                                filterColumn.DynamicFilter = dyn;
                                handledDedicated = true;
                                break;
                            }
                            default:
                                throw new ArgumentException(
                                    $"Unsupported criteria operator: '{op}'. Valid: equals, notEquals, contains, doesNotContain, beginsWith, endsWith, gt, gte, lt, lte, between, notBetween, top, topPercent, bottom, bottomPercent, blanks, nonBlanks, values, dynamic.");
                        }
                    }
                    if (customEntries.Count > 0 && !handledDedicated)
                    {
                        var cf = new CustomFilters();
                        if (customFilterAnd)
                            cf.And = true;
                        foreach (var (fop, val) in customEntries)
                            cf.AppendChild(new CustomFilter
                            {
                                Operator = fop,
                                Val = val
                            });
                        filterColumn.CustomFilters = cf;
                    }
                    autoFilter.AppendChild(filterColumn);
                }

                SaveWorksheet(afWorksheet);
                return $"/{afSheetName}/autofilter";
            }

            case "cf":
            {
                // Dispatch to specific CF type based on "type" (primary) or "rule" (alias) property.
                // R2-2: `rule=cellIs` is also accepted — user expectation from real Excel vocabulary
                // (Excel calls these "rules", OOXML calls them cfRule "type").
                var cfType = (properties.GetValueOrDefault("type")
                    ?? properties.GetValueOrDefault("rule")
                    ?? "databar").ToLowerInvariant();
                return cfType switch
                {
                    "iconset" => Add(parentPath, "iconset", position, properties),
                    "colorscale" => Add(parentPath, "colorscale", position, properties),
                    "formula" or "expression" => Add(parentPath, "formulacf", position, properties),
                    "cellis" => Add(parentPath, "cellis", position, properties),
                    "topn" or "top10" => Add(parentPath, "topn", position, properties),
                    "aboveaverage" => Add(parentPath, "aboveaverage", position, properties),
                    "uniquevalues" => Add(parentPath, "uniquevalues", position, properties),
                    "duplicatevalues" => Add(parentPath, "duplicatevalues", position, properties),
                    "containstext" => Add(parentPath, "containstext", position, properties),
                    "dateoccurring" or "timeperiod" => Add(parentPath, "dateoccurring", position, properties),
                    "belowaverage" or "containsblanks" or "notcontainsblanks" or "containserrors" or "notcontainserrors" or "contains" or "notcontains" or "beginswith" or "endswith"
                        => Add(parentPath, "cfextended", position, properties),
                    _ => Add(parentPath, "conditionalformatting", position, properties)
                };
            }

            case "databar":
            case "conditionalformatting":
            {
                // Dispatch to specific CF type if "type" or "rule" property is specified.
                // R2-2: `rule=` is an accepted alias for `type=` (matches Excel UI vocabulary).
                var cfTypeProp = properties.GetValueOrDefault("type") ?? properties.GetValueOrDefault("rule");
                if (cfTypeProp != null)
                {
                    var cfTypeLower = cfTypeProp.ToLowerInvariant();
                    if (cfTypeLower is "iconset") return Add(parentPath, "iconset", position, properties);
                    if (cfTypeLower is "colorscale") return Add(parentPath, "colorscale", position, properties);
                    if (cfTypeLower is "formula" or "expression") return Add(parentPath, "formulacf", position, properties);
                    if (cfTypeLower is "cellis") return Add(parentPath, "cellis", position, properties);
                    if (cfTypeLower is "topn" or "top10") return Add(parentPath, "topn", position, properties);
                    if (cfTypeLower is "aboveaverage") return Add(parentPath, "aboveaverage", position, properties);
                    if (cfTypeLower is "uniquevalues") return Add(parentPath, "uniquevalues", position, properties);
                    if (cfTypeLower is "duplicatevalues") return Add(parentPath, "duplicatevalues", position, properties);
                    if (cfTypeLower is "containstext") return Add(parentPath, "containstext", position, properties);
                    if (cfTypeLower is "dateoccurring" or "timeperiod") return Add(parentPath, "dateoccurring", position, properties);
                    if (cfTypeLower is "belowaverage" or "containsblanks" or "notcontainsblanks" or "containserrors" or "notcontainserrors" or "contains" or "notcontains" or "beginswith" or "endswith")
                        return Add(parentPath, "cfextended", position, properties);
                }
                var cfSegments = parentPath.TrimStart('/').Split('/', 2);
                var cfSheetName = cfSegments[0];
                var cfWorksheet = FindWorksheet(cfSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cfSheetName}");

                var sqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range") ?? properties.GetValueOrDefault("ref", "A1:A10");
                var minVal = properties.ContainsKey("min") ? properties["min"] : (string?)null;
                var maxVal = properties.ContainsKey("max") ? properties["max"] : (string?)null;
                var cfColor = properties.GetValueOrDefault("color", "638EC6");
                var normalizedColor = ParseHelpers.NormalizeArgbColor(cfColor);

                var cfRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = NextCfPriority(GetSheet(cfWorksheet))
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
                // CF6 — dataBar `showValue=false` hides the cell's numeric
                // value under the bar. Defaults to true in OOXML; only emit
                // the attribute when the user opted out.
                if (properties.TryGetValue("showValue", out var dbShowVal) && !ParseHelpers.IsTruthy(dbShowVal))
                    dataBar.ShowValue = false;
                ApplyStopIfTrue(cfRule, properties);

                var cf = new ConditionalFormatting(cfRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        sqref.Split(' ').Select(s => new StringValue(s)))
                };

                var wsElement = GetSheet(cfWorksheet);
                InsertConditionalFormatting(wsElement, cf);

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

                var csSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range", "A1:A10");
                var minColor = properties.GetValueOrDefault("mincolor", "F8696B");
                var maxColor = properties.GetValueOrDefault("maxcolor", "63BE7B");
                var midColor = properties.GetValueOrDefault("midcolor");

                var normalizedMinColor = ParseHelpers.NormalizeArgbColor(minColor);
                var normalizedMaxColor = ParseHelpers.NormalizeArgbColor(maxColor);

                // CF5 — accept user-supplied midpoint percentile (`midpoint=50`, default 50).
                var midPointStr = properties.GetValueOrDefault("midpoint")
                    ?? properties.GetValueOrDefault("midPoint")
                    ?? "50";
                var colorScale = new ColorScale();
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                if (midColor != null)
                    colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percentile, Val = midPointStr });
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
                    Priority = NextCfPriority(GetSheet(csWorksheet))
                };
                csRule.Append(colorScale);
                ApplyStopIfTrue(csRule, properties);

                var csCf = new ConditionalFormatting(csRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        csSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var csWsElement = GetSheet(csWorksheet);
                InsertConditionalFormatting(csWsElement, csCf);

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

                var isSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range", "A1:A10");
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
                    Priority = NextCfPriority(GetSheet(isWorksheet))
                };
                isRule.Append(iconSet);
                ApplyStopIfTrue(isRule, properties);

                var isCf = new ConditionalFormatting(isRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        isSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var isWsElement = GetSheet(isWorksheet);
                InsertConditionalFormatting(isWsElement, isCf);

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

                var fcfSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range", "A1:A10");
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
                _dirtyStylesheet = true;

                var dxfId = dxfs.Count!.Value - 1;

                var fcfRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.Expression,
                    Priority = NextCfPriority(GetSheet(fcfWorksheet)),
                    FormatId = dxfId
                };
                fcfRule.Append(new Formula(fcfFormula));
                ApplyStopIfTrue(fcfRule, properties);

                var fcfCf = new ConditionalFormatting(fcfRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        fcfSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var fcfWsElement = GetSheet(fcfWorksheet);
                InsertConditionalFormatting(fcfWsElement, fcfCf);

                SaveWorksheet(fcfWorksheet);
                var fcfCfCount = fcfWsElement.Elements<ConditionalFormatting>().Count();
                return $"/{fcfSheetName}/cf[{fcfCfCount}]";
            }

            case "cellis":
            {
                // R2-2: cellIs conditional formatting — compare each cell value against
                // a literal (or formula) using one of greaterThan/lessThan/... operators.
                var cisSegments = parentPath.TrimStart('/').Split('/', 2);
                var cisSheetName = cisSegments[0];
                var cisWorksheet = FindWorksheet(cisSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cisSheetName}");

                var cisSqref = properties.GetValueOrDefault("sqref")
                    ?? properties.GetValueOrDefault("range", "A1:A10");
                var opStr = (properties.GetValueOrDefault("operator") ?? "greaterThan").Trim();
                var opVal = opStr.ToLowerInvariant() switch
                {
                    "greaterthan" or "gt" or ">" => ConditionalFormattingOperatorValues.GreaterThan,
                    "lessthan" or "lt" or "<" => ConditionalFormattingOperatorValues.LessThan,
                    "greaterthanorequal" or "gte" or ">=" => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
                    "lessthanorequal" or "lte" or "<=" => ConditionalFormattingOperatorValues.LessThanOrEqual,
                    "equal" or "eq" or "=" or "==" => ConditionalFormattingOperatorValues.Equal,
                    "notequal" or "ne" or "!=" or "<>" => ConditionalFormattingOperatorValues.NotEqual,
                    "between" => ConditionalFormattingOperatorValues.Between,
                    "notbetween" => ConditionalFormattingOperatorValues.NotBetween,
                    _ => throw new ArgumentException(
                        $"Unsupported cellIs operator '{opStr}'. Valid: greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, equal, notEqual, between, notBetween.")
                };

                var primary = properties.GetValueOrDefault("value")
                    ?? properties.GetValueOrDefault("formula")
                    ?? properties.GetValueOrDefault("value1")
                    ?? throw new ArgumentException("cellIs conditional formatting requires 'value' property (e.g. value=50).");
                var secondary = properties.GetValueOrDefault("value2")
                    ?? properties.GetValueOrDefault("maxvalue");

                // Build DifferentialFormat (dxf)
                var cisDxf = new DifferentialFormat();
                if (properties.TryGetValue("font.color", out var cisFontColor))
                {
                    var normalizedFontColor = ParseHelpers.NormalizeArgbColor(cisFontColor);
                    cisDxf.Append(new Font(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedFontColor }));
                }
                if (properties.TryGetValue("font.bold", out var cisFontBold) && IsTruthy(cisFontBold))
                {
                    var existingFont = cisDxf.GetFirstChild<Font>();
                    if (existingFont != null) existingFont.Append(new Bold());
                    else cisDxf.Append(new Font(new Bold()));
                }
                if (properties.TryGetValue("fill", out var cisFill))
                {
                    var normalizedFill = ParseHelpers.NormalizeArgbColor(cisFill);
                    cisDxf.Append(new Fill(new PatternFill(
                        new BackgroundColor { Rgb = normalizedFill })
                    { PatternType = PatternValues.Solid }));
                }

                var cisWbPart = _doc.WorkbookPart
                    ?? throw new InvalidOperationException("Workbook not found");
                var cisStyleMgr = new ExcelStyleManager(cisWbPart);
                cisStyleMgr.EnsureStylesPart();
                var cisStylesheet = cisWbPart.WorkbookStylesPart!.Stylesheet!;
                var cisDxfs = cisStylesheet.GetFirstChild<DifferentialFormats>();
                if (cisDxfs == null)
                {
                    cisDxfs = new DifferentialFormats { Count = 0 };
                    cisStylesheet.Append(cisDxfs);
                }
                cisDxfs.Append(cisDxf);
                cisDxfs.Count = (uint)cisDxfs.Elements<DifferentialFormat>().Count();
                _dirtyStylesheet = true;
                var cisDxfId = cisDxfs.Count!.Value - 1;

                var cisRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.CellIs,
                    Priority = NextCfPriority(GetSheet(cisWorksheet)),
                    FormatId = cisDxfId,
                    Operator = opVal
                };
                cisRule.Append(new Formula(primary));
                if ((opVal == ConditionalFormattingOperatorValues.Between
                     || opVal == ConditionalFormattingOperatorValues.NotBetween)
                    && secondary != null)
                {
                    cisRule.Append(new Formula(secondary));
                }
                ApplyStopIfTrue(cisRule, properties);

                var cisCf = new ConditionalFormatting(cisRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        cisSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var cisWsElement = GetSheet(cisWorksheet);
                InsertConditionalFormatting(cisWsElement, cisCf);

                SaveWorksheet(cisWorksheet);
                var cisCfCount = cisWsElement.Elements<ConditionalFormatting>().Count();
                return $"/{cisSheetName}/cf[{cisCfCount}]";
            }

            case "ole":
            case "oleobject":
            case "object":
            case "embed":
            {
                // ---- Excel OLE insertion (modern form, Office 2010+) ----
                //
                // Structure produced:
                //   Worksheet > oleObjects > oleObject(progId, shapeId, r:id=embedRel)
                //     > objectPr(defaultSize=0, r:id=iconRel)
                //       > anchor(moveWithCells=1)
                //         > from(col, colOff, row, rowOff)
                //         > to  (col, colOff, row, rowOff)
                //
                // We skip the legacy VML shape that Excel historically
                // generates as a fallback — when the modern objectPr/anchor
                // is present, Office 2010+ renders from it directly. The
                // constraint-required shapeId still needs a value, so we
                // allocate one in the legal range (1-67098623) unique per
                // worksheet. For round-trip fidelity, we also create an
                // empty legacy VmlDrawingPart and register the shapeId
                // there so the relationship target exists.
                var oleSheetSegs = parentPath.TrimStart('/').Split('/', 2);
                var oleSheetName = oleSheetSegs[0];
                var oleWorksheet = FindWorksheet(oleSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {oleSheetName}");

                var oleSrc = OfficeCli.Core.OleHelper.RequireSource(properties);
                OfficeCli.Core.OleHelper.WarnOnUnknownOleProps(properties);

                // CONSISTENCY(excel-ole-display): Excel OLE does not have a
                // DrawAspect concept — worksheet objects are always shown as
                // icons via objectPr/anchor, so 'display' would be a no-op.
                // Set already rejects it; Add must too, for symmetry.
                if (properties.ContainsKey("display"))
                    throw new ArgumentException(
                        "'display' property is not supported for Excel OLE "
                        + "(Excel always shows objects as icon). Remove --prop display.");

                // CONSISTENCY(ole-name): Word/PPT OLE accept --prop name=... and
                // round-trip it via Get. SpreadsheetML x:oleObject has no Name
                // attribute in the schema, so there is nowhere to persist it.
                // Throw explicitly rather than silently dropping the value —
                // keep 'name' in KnownOleProps so Word/PPT still accept it.
                if (properties.ContainsKey("name"))
                    throw new ArgumentException(
                        "'name' property is not supported for Excel OLE "
                        + "(Spreadsheet OleObject schema has no Name attribute). Remove --prop name.");

                // 1. Embedded payload.
                var (oleEmbedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(oleWorksheet, oleSrc, _filePath);

                // 2. Icon preview image part.
                var (_, oleIconRelId) = OfficeCli.Core.OleHelper.CreateIconPart(oleWorksheet, properties);

                // 3. Resolve ProgID.
                var oleProgId = OfficeCli.Core.OleHelper.ResolveProgId(properties, oleSrc);

                // 4. Anchor: accept either cell range "B2:E6" or x/y/width/height (column units).
                // CONSISTENCY(ole-width-units): sub-cell precision is carried in
                // ColumnOffset/RowOffset (EMU) so unit-qualified widths like
                // "6cm" survive a round-trip. When the user passes a cell range
                // or a bare integer cell count, the remainder offsets are 0 and
                // behavior matches the legacy whole-cell path.
                int oleFromCol, oleFromRow, oleToCol, oleToRow;
                // FromMarker offsets are always zero (anchor starts at cell boundary);
                // ToMarker offsets carry the sub-cell EMU remainder for unit-qualified
                // width/height inputs, preserving round-trip precision.
                const long oleFromColOff = 0, oleFromRowOff = 0;
                long oleToColOff = 0, oleToRowOff = 0;
                if (properties.TryGetValue("anchor", out var oleAnchorStr) && !string.IsNullOrWhiteSpace(oleAnchorStr))
                {
                    // CONSISTENCY(ole-width-units): anchor= defines the full
                    // rectangle (start+end cells), so width/height on the same
                    // Add call would be ambiguous and are silently dropped.
                    // Warn loudly rather than fail, so existing scripts keep
                    // working but users notice the dropped value.
                    if (properties.ContainsKey("width") || properties.ContainsKey("height"))
                        Console.Error.WriteLine(
                            "Warning: 'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
                    var m = Regex.Match(oleAnchorStr, @"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", RegexOptions.IgnoreCase);
                    if (!m.Success)
                        throw new ArgumentException($"Invalid anchor: '{oleAnchorStr}'. Expected e.g. 'B2' or 'B2:E6'.");
                    // CONSISTENCY(xdr-coords): XDR ColumnId/RowId are 0-based;
                    // ColumnNameToIndex returns 1-based, so subtract 1 here.
                    oleFromCol = ColumnNameToIndex(m.Groups[1].Value) - 1;
                    oleFromRow = int.Parse(m.Groups[2].Value) - 1;
                    if (m.Groups[3].Success)
                    {
                        oleToCol = ColumnNameToIndex(m.Groups[3].Value) - 1;
                        oleToRow = int.Parse(m.Groups[4].Value) - 1;
                    }
                    else
                    {
                        oleToCol = oleFromCol + 2;
                        oleToRow = oleFromRow + 3;
                    }
                }
                else
                {
                    var (ax, ay, awEmu, ahEmu) = ParseAnchorBoundsEmu(properties, "1", "1", "3", "4");
                    oleFromCol = ax;
                    oleFromRow = ay;
                    // Split the EMU extent into (whole cells, sub-cell offset).
                    // EmuPerCol/Row constants live in ExcelHandler.Helpers.cs.
                    long wholeCols = awEmu / EmuPerColApprox;
                    long remCols = awEmu % EmuPerColApprox;
                    long wholeRows = ahEmu / EmuPerRowApprox;
                    long remRows = ahEmu % EmuPerRowApprox;
                    oleToCol = ax + (int)wholeCols;
                    oleToRow = ay + (int)wholeRows;
                    oleToColOff = remCols;
                    oleToRowOff = remRows;
                }

                // 5. Ensure the legacy VmlDrawingPart exists and carry an
                //    empty shape placeholder referencing our shapeId. This
                //    keeps the schema happy without writing VML rendering
                //    logic — Excel 2010+ renders from objectPr/anchor anyway.
                var oleVmlPart = oleWorksheet.VmlDrawingParts.FirstOrDefault()
                    ?? oleWorksheet.AddNewPart<VmlDrawingPart>();
                // Allocate a unique shapeId per worksheet (1025+N is the
                // conventional Excel starting point for legacy VML shapes).
                var existingOleCount = GetSheet(oleWorksheet).Descendants<OleObject>().Count();
                uint oleShapeId = (uint)(1025 + existingOleCount);
                EnsureExcelVmlShapeForOle(oleVmlPart, oleShapeId, oleFromCol, oleFromRow, oleToCol, oleToRow);

                // Ensure worksheet references the VML drawing part.
                var oleWsElement = GetSheet(oleWorksheet);
                if (oleWsElement.GetFirstChild<LegacyDrawing>() == null)
                {
                    var vmlRelId = oleWorksheet.GetIdOfPart(oleVmlPart);
                    // LegacyDrawing must sit after the AutoFilter/Phonetic
                    // region per schema order — safe to insert before the
                    // last known printing-related elements. Use InsertAfter
                    // relative to AutoFilter when present, else append.
                    var lgd = new LegacyDrawing { Id = vmlRelId };
                    var pageSetup = oleWsElement.GetFirstChild<PageSetup>();
                    if (pageSetup != null)
                        oleWsElement.InsertAfter(lgd, pageSetup);
                    else
                        oleWsElement.AppendChild(lgd);
                }

                // 6. Build the oleObject element + objectPr/anchor.
                var oleObj = new OleObject
                {
                    ProgId = oleProgId,
                    ShapeId = oleShapeId,
                    Id = oleEmbedRelId,
                };
                var objectPr = new EmbeddedObjectProperties
                {
                    DefaultSize = false,
                    Id = oleIconRelId,
                };
                var anchor = new ObjectAnchor { MoveWithCells = true };
                anchor.AppendChild(new FromMarker(
                    new XDR.ColumnId(oleFromCol.ToString()),
                    new XDR.ColumnOffset(oleFromColOff.ToString()),
                    new XDR.RowId(oleFromRow.ToString()),
                    new XDR.RowOffset(oleFromRowOff.ToString())));
                anchor.AppendChild(new ToMarker(
                    new XDR.ColumnId(oleToCol.ToString()),
                    new XDR.ColumnOffset(oleToColOff.ToString()),
                    new XDR.RowId(oleToRow.ToString()),
                    new XDR.RowOffset(oleToRowOff.ToString())));
                objectPr.AppendChild(anchor);
                oleObj.AppendChild(objectPr);

                // 7. Find/create oleObjects collection and append.
                var oleObjects = oleWsElement.GetFirstChild<OleObjects>();
                if (oleObjects == null)
                {
                    oleObjects = new OleObjects();
                    // Schema: oleObjects sits between picture and controls;
                    // safest is after tableParts if present, else before
                    // pageSetup, else append.
                    var insertBefore = oleWsElement.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.ExtensionList>()
                        ?? (OpenXmlElement?)null;
                    if (insertBefore != null)
                        oleWsElement.InsertBefore(oleObjects, insertBefore);
                    else
                        oleWsElement.AppendChild(oleObjects);
                }
                oleObjects.AppendChild(oleObj);

                SaveWorksheet(oleWorksheet);

                var oleCount = oleWsElement.Descendants<OleObject>().Count();
                return $"/{oleSheetName}/ole[{oleCount}]";
            }

            case "picture":
            case "image":
            case "img":
            {
                var picSegments = parentPath.TrimStart('/').Split('/', 2);
                var picSheetName = picSegments[0];
                var picWorksheet = FindWorksheet(picSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {picSheetName}");

                if (!properties.TryGetValue("path", out var imgPath)
                    && !properties.TryGetValue("src", out imgPath))
                    throw new ArgumentException("'src' property is required for picture type");

                // CONSISTENCY(picture-emu): use ParseAnchorBoundsEmu like OLE,
                // so width/height accept unit-qualified strings ("6cm", "2in")
                // in addition to bare integer cell counts.
                var (px, py, pwEmu, phEmu) = ParseAnchorBoundsEmu(properties, "0", "0", "5", "5");
                // P9: accept `altText=` as alias for `alt=`.
                var alt = properties.GetValueOrDefault("alt")
                    ?? properties.GetValueOrDefault("altText")
                    ?? properties.GetValueOrDefault("alttext", "");

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

                // CONSISTENCY(svg-dual-rep): same dual-representation as Word
                // and PPT — main r:embed points to a PNG fallback, SVG is
                // referenced via a:blip/a:extLst asvg:svgBlip.
                string imgRelId;
                string? xlSvgRelId = null;
                if (imgPartType == ImagePartType.Svg)
                {
                    var svgPart = picDrawingsPart.AddImagePart(ImagePartType.Svg);
                    svgPart.FeedData(xlImgStream);
                    xlSvgRelId = picDrawingsPart.GetIdOfPart(svgPart);

                    if (properties.TryGetValue("fallback", out var xlFallback) && !string.IsNullOrWhiteSpace(xlFallback))
                    {
                        var (fbRaw, fbType) = OfficeCli.Core.ImageSource.Resolve(xlFallback);
                        using var fbDispose = fbRaw;
                        var fbPart = picDrawingsPart.AddImagePart(fbType);
                        fbPart.FeedData(fbRaw);
                        imgRelId = picDrawingsPart.GetIdOfPart(fbPart);
                    }
                    else
                    {
                        var pngPart = picDrawingsPart.AddImagePart(ImagePartType.Png);
                        pngPart.FeedData(new MemoryStream(
                            OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                        imgRelId = picDrawingsPart.GetIdOfPart(pngPart);
                    }
                }
                else
                {
                    var imgPart = picDrawingsPart.AddImagePart(imgPartType);
                    imgPart.FeedData(xlImgStream);
                    imgRelId = picDrawingsPart.GetIdOfPart(imgPart);
                }

                var picId = picDrawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
                    .Select(p => (uint?)p.Id?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;
                // CONSISTENCY(picture-emu): split EMU extent into whole-cell
                // count + sub-cell offset, matching the OLE anchor path.
                long picWholeCols = pwEmu / EmuPerColApprox;
                long picRemCols = pwEmu % EmuPerColApprox;
                long picWholeRows = phEmu / EmuPerRowApprox;
                long picRemRows = phEmu % EmuPerRowApprox;

                // DEFERRED(xlsx/picture-anchor-mode) P12: honor `anchorMode=`
                // oneCell|absolute|twoCell. Default remains twoCell for back-compat.
                // oneCell → <xdr:oneCellAnchor> with from + ext; picture auto-scales
                //           if the column/row containing "from" is resized.
                // absolute → <xdr:absoluteAnchor> with pos (x/y EMU) + ext; picture
                //            does not move or resize with cells.
                // twoCell  → <xdr:twoCellAnchor> with from + to markers (default).
                //
                // CONSISTENCY(ole-width-units): `anchor=B2:E6` (cell-range) is
                // parsed here the same way as the OLE and shape branches; it
                // implies anchorMode=twoCell. `anchor=oneCell|twoCell|absolute`
                // is still honored as the mode for back-compat. Explicit
                // `anchorMode=` always wins. When both `anchor=<range>` and
                // `x/y/width/height` are supplied, anchor wins with a warning
                // (same convention as the shape/OLE branches).
                var picAnchorRaw = properties.GetValueOrDefault("anchor");
                var picAnchorModeExplicit = properties.GetValueOrDefault("anchorMode");
                bool picHasRange = false;
                int picRangeFromCol = 0, picRangeFromRow = 0, picRangeToCol = -1, picRangeToRow = -1;
                // `anchor=` is either a cell-range ("B2" / "B2:E6") or an
                // anchorMode token ("oneCell"/"twoCell"/"absolute"). Prefer the
                // cell-range interpretation; fall back to mode-token only when
                // the value is a recognized token. Explicit `anchorMode=` wins
                // the mode selection regardless.
                if (!string.IsNullOrWhiteSpace(picAnchorRaw) && !IsAnchorModeToken(picAnchorRaw))
                {
                    if (!TryParseCellRangeAnchor(picAnchorRaw, out picRangeFromCol, out picRangeFromRow, out picRangeToCol, out picRangeToRow))
                        throw new ArgumentException($"Invalid anchor: '{picAnchorRaw}'. Expected e.g. 'B2', 'B2:E6', or one of 'oneCell'/'twoCell'/'absolute'.");
                    picHasRange = true;
                    if (properties.ContainsKey("width") || properties.ContainsKey("height")
                        || properties.ContainsKey("x") || properties.ContainsKey("y"))
                        Console.Error.WriteLine(
                            "Warning: 'x'/'y'/'width'/'height' are ignored when 'anchor' is a cell range (anchor defines the full rectangle).");
                }
                var picAnchorMode = (picAnchorModeExplicit
                    ?? (picHasRange ? "twoCell" : picAnchorRaw)
                    ?? "twoCell").Trim().ToLowerInvariant();

                var picShape = BuildPictureElementWithTransform(picId, alt ?? "", imgRelId, xlSvgRelId, properties);

                // For oneCell / absolute anchors the size is carried by an <xdr:ext>
                // element instead of a To marker, so we must also stamp the extent
                // onto the picture's Transform2D so rotation / flip metadata plus
                // the rendered size stay in sync.
                if (picAnchorMode is "onecell" or "absolute")
                {
                    var picXfrm = picShape.Descendants<Drawing.Transform2D>().FirstOrDefault();
                    if (picXfrm != null)
                    {
                        var ext2d = picXfrm.Extents ?? new Drawing.Extents();
                        ext2d.Cx = pwEmu;
                        ext2d.Cy = phEmu;
                        picXfrm.Extents = ext2d;
                    }
                }

                OpenXmlElement anchor;
                switch (picAnchorMode)
                {
                    case "onecell":
                    {
                        int oneFromCol = picHasRange ? picRangeFromCol : px;
                        int oneFromRow = picHasRange ? picRangeFromRow : py;
                        var oneAnchor = new XDR.OneCellAnchor(
                            new XDR.FromMarker(
                                new XDR.ColumnId(oneFromCol.ToString()),
                                new XDR.ColumnOffset("0"),
                                new XDR.RowId(oneFromRow.ToString()),
                                new XDR.RowOffset("0")
                            ),
                            new XDR.Extent { Cx = pwEmu, Cy = phEmu },
                            picShape,
                            new XDR.ClientData()
                        );
                        anchor = oneAnchor;
                        break;
                    }
                    case "absolute":
                    {
                        // Absolute anchor pos: accept `x=`/`y=` in the same unit
                        // syntax as width/height (bare EMU, or "1in", "2cm").
                        long absX = 0, absY = 0;
                        if (properties.TryGetValue("x", out var absXs))
                            absX = OfficeCli.Core.EmuConverter.ParseEmu(absXs);
                        if (properties.TryGetValue("y", out var absYs))
                            absY = OfficeCli.Core.EmuConverter.ParseEmu(absYs);
                        var absAnchor = new XDR.AbsoluteAnchor(
                            new XDR.Position { X = absX, Y = absY },
                            new XDR.Extent { Cx = pwEmu, Cy = phEmu },
                            picShape,
                            new XDR.ClientData()
                        );
                        anchor = absAnchor;
                        break;
                    }
                    default:
                    {
                        int twoFromCol, twoFromRow, twoToCol, twoToRow;
                        long twoToColOff, twoToRowOff;
                        if (picHasRange)
                        {
                            twoFromCol = picRangeFromCol;
                            twoFromRow = picRangeFromRow;
                            if (picRangeToCol >= 0)
                            {
                                twoToCol = picRangeToCol;
                                twoToRow = picRangeToRow;
                                twoToColOff = 0;
                                twoToRowOff = 0;
                            }
                            else
                            {
                                // Single-cell range in twoCell mode: fall back to width/height extent.
                                twoToCol = twoFromCol + (int)picWholeCols;
                                twoToRow = twoFromRow + (int)picWholeRows;
                                twoToColOff = picRemCols;
                                twoToRowOff = picRemRows;
                            }
                        }
                        else
                        {
                            twoFromCol = px;
                            twoFromRow = py;
                            twoToCol = px + (int)picWholeCols;
                            twoToRow = py + (int)picWholeRows;
                            twoToColOff = picRemCols;
                            twoToRowOff = picRemRows;
                        }
                        anchor = new XDR.TwoCellAnchor(
                            new XDR.FromMarker(
                                new XDR.ColumnId(twoFromCol.ToString()),
                                new XDR.ColumnOffset("0"),
                                new XDR.RowId(twoFromRow.ToString()),
                                new XDR.RowOffset("0")
                            ),
                            new XDR.ToMarker(
                                new XDR.ColumnId(twoToCol.ToString()),
                                new XDR.ColumnOffset(twoToColOff.ToString()),
                                new XDR.RowId(twoToRow.ToString()),
                                new XDR.RowOffset(twoToRowOff.ToString())
                            ),
                            picShape,
                            new XDR.ClientData()
                        );
                        break;
                    }
                }

                picDrawingsPart.WorksheetDrawing.AppendChild(anchor);

                // P10: picture decorative=true — emit <a:extLst><a:ext uri="...">
                // <a16:decorative val="1"/></a:ext></a:extLst> under <xdr:cNvPr>.
                // Requires declaring xmlns:a16 on the drawing root; mirrors the
                // sparkline pattern of adding namespaces idempotently.
                if (properties.TryGetValue("decorative", out var picDec) && IsTruthy(picDec))
                {
                    var picCNvPrDec = anchor.Descendants<XDR.NonVisualDrawingProperties>().FirstOrDefault();
                    if (picCNvPrDec != null)
                    {
                        const string a16Ns = "http://schemas.microsoft.com/office/drawing/2014/main";
                        var wsDrawingRoot = picDrawingsPart.WorksheetDrawing;
                        if (wsDrawingRoot.LookupNamespace("a16") == null)
                            wsDrawingRoot.AddNamespaceDeclaration("a16", a16Ns);
                        var decInner = new OpenXmlUnknownElement("a16", "decorative", a16Ns);
                        decInner.SetAttribute(new OpenXmlAttribute("", "val", "", "1"));
                        var ext = new Drawing.Extension { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };
                        ext.Append(decInner);
                        var extLst = picCNvPrDec.GetFirstChild<Drawing.ExtensionList>()
                            ?? picCNvPrDec.AppendChild(new Drawing.ExtensionList());
                        extLst.Append(ext);
                    }
                }

                // P8: picture-level hyperlink — <a:hlinkClick> under <xdr:cNvPr>.
                // External URL → add rel on DrawingsPart, reference its rId.
                // Internal (starts with '#') → no rel, use Location attribute.
                // CONSISTENCY(xlsx-hyperlink): mirrors cell link handling in
                // commit 60e1455.
                var picHlink = properties.GetValueOrDefault("hyperlink")
                    ?? properties.GetValueOrDefault("link");
                if (!string.IsNullOrWhiteSpace(picHlink))
                {
                    var picCNvPr = anchor.Descendants<XDR.NonVisualDrawingProperties>().FirstOrDefault();
                    if (picCNvPr != null)
                    {
                        Drawing.HyperlinkOnClick hlClick;
                        if (picHlink.StartsWith("#"))
                        {
                            // No rel, no @r:id — pure in-document jump via @location.
                            hlClick = new Drawing.HyperlinkOnClick { Id = "" };
                            hlClick.SetAttribute(new OpenXmlAttribute(
                                "", "location", "", picHlink.Substring(1)));
                        }
                        else
                        {
                            var hlUri = new Uri(picHlink, UriKind.RelativeOrAbsolute);
                            var hlRel = picDrawingsPart.AddHyperlinkRelationship(hlUri, isExternal: true);
                            hlClick = new Drawing.HyperlinkOnClick { Id = hlRel.Id };
                        }
                        picCNvPr.AppendChild(hlClick);
                    }
                }

                picDrawingsPart.WorksheetDrawing.Save();

                // DEFERRED(xlsx/picture-anchor-mode) P12: enumerate all anchor
                // kinds (twoCell / oneCell / absolute) when counting picture slots.
                var picAnchors = picDrawingsPart.WorksheetDrawing
                    .Elements<OpenXmlElement>()
                    .Where(a => (a is XDR.TwoCellAnchor || a is XDR.OneCellAnchor || a is XDR.AbsoluteAnchor)
                        && a.Descendants<XDR.Picture>().Any())
                    .ToList();
                var picIdx = picAnchors.IndexOf(anchor) + 1;

                return $"/{picSheetName}/picture[{picIdx}]";
            }

            case "shape" or "textbox":
            {
                var shpSegments = parentPath.TrimStart('/').Split('/', 2);
                var shpSheetName = shpSegments[0];
                var shpWorksheet = FindWorksheet(shpSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {shpSheetName}");

                // CONSISTENCY(ole-width-units): accept `anchor=B2:F7` as a cell
                // range (same grammar as OLE's anchor=), alongside the legacy
                // x/y/width/height (column/row units) form. When both are
                // supplied, warn and let anchor= win — it defines the full
                // rectangle, so width/height are ambiguous.
                int sx, sy, sw, sh;
                if (properties.TryGetValue("anchor", out var shpAnchorStr) && !string.IsNullOrWhiteSpace(shpAnchorStr))
                {
                    if (properties.ContainsKey("width") || properties.ContainsKey("height")
                        || properties.ContainsKey("x") || properties.ContainsKey("y"))
                        Console.Error.WriteLine(
                            "Warning: 'x'/'y'/'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
                    if (!TryParseCellRangeAnchor(shpAnchorStr, out var sxFrom, out var syFrom, out var sxTo, out var syTo))
                        throw new ArgumentException($"Invalid anchor: '{shpAnchorStr}'. Expected e.g. 'B2' or 'B2:F7'.");
                    sx = sxFrom;
                    sy = syFrom;
                    if (sxTo < 0) { sxTo = sx + 4; syTo = sy + 2; }
                    sw = sxTo - sx;
                    sh = syTo - sy;
                }
                else
                {
                    (sx, sy, sw, sh) = ParseAnchorBounds(properties, "1", "1", "5", "3");
                }
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

                var shpId = shpDrawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
                    .Select(p => (uint?)p.Id?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;
                if (string.IsNullOrEmpty(shpName)) shpName = $"Shape {shpId}";

                // CONSISTENCY(shape-preset): map `preset=` to a:prstGeom prst value
                // using the same token set PowerPointHandler.ParsePresetShape accepts.
                // textbox ignores preset (always "rect"). Default for shape: "rect".
                var shpPreset = Drawing.ShapeTypeValues.Rectangle;
                if (string.Equals(type, "shape", StringComparison.OrdinalIgnoreCase)
                    && properties.TryGetValue("preset", out var shpPresetRaw)
                    && !string.IsNullOrWhiteSpace(shpPresetRaw))
                    shpPreset = ParseExcelShapePreset(shpPresetRaw);

                // Build ShapeProperties
                var shpXfrm = new Drawing.Transform2D(
                    new Drawing.Offset { X = 0, Y = 0 },
                    new Drawing.Extents { Cx = 0, Cy = 0 }
                );
                ApplyTransform2DRotationFlip(shpXfrm, properties);
                var spPr = new XDR.ShapeProperties(
                    shpXfrm,
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = shpPreset }
                );

                // Fill — single-color `fill=` OR gradient `gradientFill=C1-C2[-C3][:angle]`.
                // SH6/shape-gradient-fill: keep `fill=` strictly single-color; gradient has its own prop
                // to avoid ambiguity (FF0000-0000FF could otherwise collide with single ARGB literals).
                if (properties.TryGetValue("gradientFill", out var shpGradFill)
                    && !string.IsNullOrWhiteSpace(shpGradFill))
                {
                    spPr.AppendChild(BuildShapeGradientFill(shpGradFill));
                }
                else if (properties.TryGetValue("fill", out var shpFill))
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
                        if (IsValidBooleanString(normalizedShadow) && IsTruthy(normalizedShadow)) normalizedShadow = "000000";
                        shpEffectList ??= new Drawing.EffectList();
                        shpEffectList.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedShadow, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                    }
                    if (properties.TryGetValue("glow", out var shpGlow) && !shpGlow.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var normalizedGlow = shpGlow.Replace(':', '-');
                        if (IsValidBooleanString(normalizedGlow) && IsTruthy(normalizedGlow)) normalizedGlow = "4472C4";
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

                    // R2-3: accept both bare (`size`, `bold`, `color`, `font`) and `font.*`
                    // sub-prop forms (`font.size`, `font.bold`, `font.color`, `font.name`,
                    // `font.italic`, `font.underline`) for consistency with cell/comment.
                    // Schema order: attributes → solidFill → effectLst → latin/ea
                    string? rawSize = properties.GetValueOrDefault("size")
                        ?? properties.GetValueOrDefault("font.size");
                    if (rawSize != null)
                        rPr.FontSize = (int)Math.Round(ParseHelpers.ParseFontSize(rawSize) * 100);

                    string? rawBold = properties.GetValueOrDefault("bold")
                        ?? properties.GetValueOrDefault("font.bold");
                    if (rawBold != null && IsTruthy(rawBold))
                        rPr.Bold = true;

                    string? rawItalic = properties.GetValueOrDefault("italic")
                        ?? properties.GetValueOrDefault("font.italic");
                    if (rawItalic != null && IsTruthy(rawItalic))
                        rPr.Italic = true;

                    if (properties.TryGetValue("font.underline", out var shpUnder)
                        || properties.TryGetValue("underline", out shpUnder))
                    {
                        var uv = shpUnder.ToLowerInvariant();
                        rPr.Underline = uv switch
                        {
                            "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                            "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                            "none" or "false" => Drawing.TextUnderlineValues.None,
                            _ => Drawing.TextUnderlineValues.Single
                        };
                    }

                    // Fill (color) before fonts
                    string? rawColor = properties.GetValueOrDefault("color")
                        ?? properties.GetValueOrDefault("font.color");
                    if (rawColor != null)
                    {
                        var (cRgb, _) = ParseHelpers.SanitizeColorForOoxml(rawColor);
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
                            if (IsValidBooleanString(normalizedTs) && IsTruthy(normalizedTs)) normalizedTs = "000000";
                            txtEffects ??= new Drawing.EffectList();
                            txtEffects.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedTs, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                        }
                        if (properties.TryGetValue("glow", out var tg) && !tg.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var normalizedTg = tg.Replace(':', '-');
                            if (IsValidBooleanString(normalizedTg) && IsTruthy(normalizedTg)) normalizedTg = "4472C4";
                            txtEffects ??= new Drawing.EffectList();
                            txtEffects.AppendChild(OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedTg, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                        }
                        if (txtEffects != null)
                            rPr.AppendChild(txtEffects);
                    }

                    // Fonts last (schema order). Accept `font=Arial` or `font.name=Arial`.
                    string? rawFontName = properties.GetValueOrDefault("font.name")
                        ?? properties.GetValueOrDefault("font");
                    if (rawFontName != null)
                    {
                        rPr.AppendChild(new Drawing.LatinFont { Typeface = rawFontName });
                        rPr.AppendChild(new Drawing.EastAsianFont { Typeface = rawFontName });
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

                // T4 — reject a new table whose ref overlaps any existing table on
                // the same sheet. Excel silently corrupts the file otherwise.
                foreach (var existingTdp in tblWorksheet.TableDefinitionParts)
                {
                    var existing = existingTdp.Table;
                    if (existing?.Reference?.Value is not string existingRef) continue;
                    if (RangesOverlap(rangeRef, existingRef))
                        throw new ArgumentException(
                            $"Table ref overlaps existing table '{existing.Name?.Value ?? existing.DisplayName?.Value}' ({existingRef})");
                }

                var existingTableIds = _doc.WorkbookPart!.WorksheetParts
                    .SelectMany(wp => wp.TableDefinitionParts)
                    .Select(tdp => tdp.Table?.Id?.Value ?? 0);
                var tableId = existingTableIds.Any() ? existingTableIds.Max() + 1 : 1;

                var userProvidedName = properties.ContainsKey("name");
                var tableName = SanitizeTableIdentifier(
                    properties.GetValueOrDefault("name", $"Table{tableId}"),
                    userProvided: userProvidedName);
                // displayName defaults to the (already-sanitized) tableName; if
                // name was user-provided it flows through verbatim so Excel
                // shows the same identifier the user asked for.
                var userProvidedDisplay = properties.ContainsKey("displayName");
                var displayName = SanitizeTableIdentifier(
                    properties.GetValueOrDefault("displayName", tableName),
                    userProvided: userProvidedDisplay || userProvidedName);
                var styleName = properties.GetValueOrDefault("style", "TableStyleMedium2");
                // T6 — validate style name against the built-in whitelist +
                // any workbook-level customStyles. Unknown names silently
                // fell through to Excel which would either ignore or
                // reject the file; prefer an explicit ArgumentException.
                ValidateTableStyleName(styleName);
                // T1 — accept `showHeader=false` alias alongside `headerRow=false`.
                var hasHeader = !(properties.TryGetValue("headerRow", out var hrVal) && !IsTruthy(hrVal))
                             && !(properties.TryGetValue("showHeader", out var shVal) && !IsTruthy(shVal));
                // CONSISTENCY(table-totalrow): accept `showTotals=true` alias
                // alongside `totalRow=true` (mirrors the `showHeader` alias
                // pattern above for users coming from Office API vocabulary).
                var hasTotalRow = (properties.TryGetValue("totalRow", out var trVal) && IsTruthy(trVal))
                               || (properties.TryGetValue("showTotals", out var stVal) && IsTruthy(stVal));

                var rangeParts = rangeRef.Split(':');
                var (startCol, startRow) = ParseCellReference(rangeParts[0]);
                var (endCol, endRow) = ParseCellReference(rangeParts[1]);
                var startColIdx = ColumnNameToIndex(startCol);
                var endColIdx = ColumnNameToIndex(endCol);
                var colCount = endColIdx - startColIdx + 1;

                // T5-ext: autoExpand=true probes the sheet for contiguous
                // non-empty rows immediately below the declared ref and grows
                // endRow to include them. Mirrors Excel's "Table expand when
                // you type below" behavior at Add time.
                if (properties.TryGetValue("autoExpand", out var autoExpandRaw) && IsTruthy(autoExpandRaw))
                {
                    var sheetDataForProbe = GetSheet(tblWorksheet).GetFirstChild<SheetData>();
                    if (sheetDataForProbe != null)
                    {
                        int probeRow = endRow + 1;
                        while (true)
                        {
                            var probe = sheetDataForProbe.Elements<Row>()
                                .FirstOrDefault(r => r.RowIndex?.Value == (uint)probeRow);
                            if (probe == null) break;
                            // non-empty = at least one cell in the column
                            // span carries a CellValue or InlineString.
                            bool anyNonEmpty = false;
                            for (int ci = 0; ci < colCount; ci++)
                            {
                                var cLetter = IndexToColumnName(startColIdx + ci);
                                var cRef = $"{cLetter}{probeRow}";
                                var probeCell = probe.Elements<Cell>()
                                    .FirstOrDefault(c => c.CellReference?.Value == cRef);
                                if (probeCell == null) continue;
                                if (probeCell.CellValue != null || probeCell.InlineString != null)
                                {
                                    anyNonEmpty = true;
                                    break;
                                }
                            }
                            if (!anyNonEmpty) break;
                            endRow = probeRow;
                            probeRow++;
                        }
                        rangeRef = $"{startCol}{startRow}:{endCol}{endRow}";
                    }
                }

                // CONSISTENCY(table-totalrow): a:totalsRowShown MUST point at a row
                // OUTSIDE the data area. Previously we reused endRow as the totals
                // row, which overwrote whatever data lived on that last row. Expand
                // the ref by one row so the totals row is appended below the data
                // instead of stamping over it.
                if (hasTotalRow)
                {
                    endRow += 1;
                    rangeRef = $"{startCol}{startRow}:{endCol}{endRow}";
                }

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
                            // Excel rejects a table whose header cell is typed
                            // as a number. Convert the cell to an inline string
                            // so the header reads as text, and tableColumn name
                            // (read above) still matches the cell's visible
                            // value exactly — Excel also requires that match.
                            if (headerCell != null && (headerCell.DataType == null || headerCell.DataType.Value == CellValues.Number))
                            {
                                var text = colNames[i];
                                headerCell.DataType = CellValues.InlineString;
                                headerCell.CellValue = null;
                                headerCell.InlineString = new InlineString(new Text(text));
                            }
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
                if (!hasHeader)
                    table.HeaderRowCount = 0;

                table.AppendChild(new AutoFilter { Reference = rangeRef });

                // Dedupe duplicate column names (Excel also trips on those).
                var usedColNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < colCount; i++)
                {
                    var baseName = colNames[i];
                    var cn = baseName;
                    var dedupIdx = 2;
                    while (!usedColNames.Add(cn))
                        cn = $"{baseName}{dedupIdx++}";
                    colNames[i] = cn;
                }

                var tableColumns = new TableColumns { Count = (uint)colCount };
                for (int i = 0; i < colCount; i++)
                    tableColumns.AppendChild(new TableColumn { Id = (uint)(i + 1), Name = colNames[i] });
                table.AppendChild(tableColumns);

                // T7-ext: `columns.N.dxfId=<id>` stamps dataDxfId on the
                // target tableColumn (N is 1-based). The id must reference
                // an existing workbook differentialFormats entry; we do not
                // synthesize new dxfs here — users who want inline style
                // values should register a dxf first via `add dxf` (or the
                // underlying APIs) and then reference it.
                var tblColList = tableColumns.Elements<TableColumn>().ToList();
                foreach (var (rawKey, rawVal) in properties)
                {
                    var m = Regex.Match(rawKey, @"^columns?\.(\d+)\.dxfId$",
                        RegexOptions.IgnoreCase);
                    if (!m.Success) continue;
                    var n = int.Parse(m.Groups[1].Value);
                    if (n < 1 || n > tblColList.Count) continue;
                    if (!uint.TryParse(rawVal, out var dxfId))
                        throw new ArgumentException(
                            $"columns.{n}.dxfId requires a numeric dxf id, got: '{rawVal}'");
                    tblColList[n - 1].DataFormatId = dxfId;
                }

                // T2 — wire the banded rows/columns + first/last column
                // flags onto the TableStyleInfo. Each accepts `showX` or
                // its alias; default matches the old hard-coded values so
                // omitting them is identical to previous behavior.
                table.AppendChild(new TableStyleInfo
                {
                    Name = styleName,
                    ShowFirstColumn = properties.TryGetValue("showFirstColumn", out var sfc)
                        ? IsTruthy(sfc) : false,
                    ShowLastColumn = properties.TryGetValue("showLastColumn", out var slc)
                        ? IsTruthy(slc) : false,
                    ShowRowStripes = properties.TryGetValue("showBandedRows", out var sbr)
                        ? IsTruthy(sbr) : true,
                    ShowColumnStripes = properties.TryGetValue("showBandedColumns", out var sbc)
                        ? IsTruthy(sbc) : false
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
                    // Per-column totalsRowFunction tokens: "none,sum,average"
                    // → first col = label/none, rest = sum, average. If the
                    // user didn't pass it, default to "none" on col0 + "sum"
                    // on the rest (legacy behavior).
                    string[] trfTokens = properties.TryGetValue("totalsRowFunction", out var trfRaw)
                        ? trfRaw.Split(',').Select(s => s.Trim()).ToArray()
                        : Array.Empty<string>();
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

                        var tokRaw = ci < trfTokens.Length ? trfTokens[ci].ToLowerInvariant() : "";
                        var (trfEnum, subtotalCode) = MapTotalsRowFunction(tokRaw);

                        if (ci == 0 && (tokRaw == "" || tokRaw == "none" || tokRaw == "label"))
                        {
                            // First column: label "Total"
                            tblCols[ci].TotalsRowLabel = "Total";
                            existingCell.CellValue = new CellValue("Total");
                            existingCell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        else if (trfEnum == TotalsRowFunctionValues.None)
                        {
                            // Skip — leave cell empty, no function set.
                        }
                        else
                        {
                            // Default non-first column (no explicit token) = SUM
                            if (ci > 0 && tokRaw == "")
                            {
                                trfEnum = TotalsRowFunctionValues.Sum;
                                subtotalCode = 109;
                            }
                            tblCols[ci].TotalsRowFunction = trfEnum;
                            var dataStartRow = hasHeader ? startRow + 1 : startRow;
                            var dataEndRow = (int)totalRowIdx - 1;
                            var formulaRange = $"{colLetter}{dataStartRow}:{colLetter}{dataEndRow}";
                            existingCell.CellFormula = new CellFormula($"SUBTOTAL({subtotalCode},{formulaRange})");
                        }
                    }

                    // T10: per-column custom totalsFormula override. Syntax:
                    //   columns.N.totalsFormula="=SUM(Table1[Sales])/2"
                    // where N is 1-based. This sets the column's
                    // totalsRowFunction to "custom" + writes <calculatedColumnFormula>,
                    // and replaces the SUBTOTAL cell formula with the user's.
                    foreach (var (rawKey, rawVal) in properties)
                    {
                        var m = Regex.Match(rawKey, @"^columns?\.(\d+)\.totalsFormula$",
                            RegexOptions.IgnoreCase);
                        if (!m.Success) continue;
                        var n = int.Parse(m.Groups[1].Value);
                        if (n < 1 || n > tblCols.Count) continue;
                        var ci = n - 1;
                        var colLetter = IndexToColumnName(startColIdx + ci);
                        var cellRefStr = $"{colLetter}{totalRowIdx}";
                        var existingCell = totalRow.Elements<Cell>()
                            .FirstOrDefault(c => c.CellReference?.Value == cellRefStr)
                            ?? totalRow.AppendChild(new Cell { CellReference = cellRefStr });

                        var customFormula = rawVal.TrimStart('=');
                        tblCols[ci].TotalsRowFunction = TotalsRowFunctionValues.Custom;
                        tblCols[ci].TotalsRowLabel = null;
                        tblCols[ci].TotalsRowFormula = new TotalsRowFormula(customFormula);
                        existingCell.CellFormula = new CellFormula(customFormula);
                        existingCell.CellValue = null;
                        existingCell.DataType = null;
                    }
                }

                // CONSISTENCY(xlsx/table-autoexpand): persist the opt-in flag as
                // a custom-namespace attribute on <x:table> so eager auto-grow
                // survives reopen. Real Excel ignores unknown-namespace attrs.
                if (properties.TryGetValue("autoExpand", out var aeRaw) && IsTruthy(aeRaw))
                    SetTableAutoExpandMarker(table, true);

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
                SaveWorksheet(tblWorksheet);

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

                // Support dataRange: read cell data from worksheet and build series with cell references
                string[]? categories;
                List<(string name, double[] values)> seriesData;
                var dataRangeStr = properties.FirstOrDefault(kv =>
                    kv.Key.Equals("datarange", StringComparison.OrdinalIgnoreCase)
                    || kv.Key.Equals("dataRange", StringComparison.OrdinalIgnoreCase)
                    || kv.Key.Equals("range", StringComparison.OrdinalIgnoreCase)).Value;
                if (!string.IsNullOrEmpty(dataRangeStr))
                {
                    (seriesData, categories) = ParseDataRangeForChart(dataRangeStr, chartSheetName, properties);
                }
                else
                {
                    categories = ChartHelper.ParseCategories(properties);
                    seriesData = ChartHelper.ParseSeriesData(properties);
                }

                if (seriesData.Count == 0)
                    throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                        "or dataRange=\"Sheet1!A1:D5\" or series1=\"Revenue:100,200,300\"");

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

                // Position via TwoCellAnchor (shared by both standard and extended charts)
                // CONSISTENCY(ole-width-units): accept `anchor=D2:J18` as a cell
                // range (same grammar as OLE, shape, picture). When both
                // `anchor=<range>` and `x/y/width/height` are supplied, anchor
                // wins with a warning — matches shape/picture/OLE convention.
                int fromCol, fromRow, toCol, toRow;
                if (properties.TryGetValue("anchor", out var chartAnchorStr) && !string.IsNullOrWhiteSpace(chartAnchorStr))
                {
                    if (properties.ContainsKey("width") || properties.ContainsKey("height")
                        || properties.ContainsKey("x") || properties.ContainsKey("y"))
                        Console.Error.WriteLine(
                            "Warning: 'x'/'y'/'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
                    if (!TryParseCellRangeAnchor(chartAnchorStr, out var cxFrom, out var cyFrom, out var cxTo, out var cyTo))
                        throw new ArgumentException($"Invalid anchor: '{chartAnchorStr}'. Expected e.g. 'D2' or 'D2:J18'.");
                    fromCol = cxFrom;
                    fromRow = cyFrom;
                    if (cxTo < 0) { cxTo = fromCol + 8; cyTo = fromRow + 15; }
                    toCol = cxTo;
                    toRow = cyTo;
                }
                else
                {
                    fromCol = properties.TryGetValue("x", out var xStr) ? ParseHelpers.SafeParseInt(xStr, "x") : 0;
                    fromRow = properties.TryGetValue("y", out var yStr) ? ParseHelpers.SafeParseInt(yStr, "y") : 0;
                    toCol = properties.TryGetValue("width", out var wStr) ? fromCol + ParseHelpers.SafeParseInt(wStr, "width") : fromCol + 8;
                    toRow = properties.TryGetValue("height", out var hStr) ? fromRow + ParseHelpers.SafeParseInt(hStr, "height") : fromRow + 15;
                }

                // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
                if (ChartExBuilder.IsExtendedChartType(chartType))
                {
                    var cxChartSpace = ChartExBuilder.BuildExtendedChartSpace(
                        chartType, chartTitle, categories, seriesData, properties);
                    var extChartPart = drawingsPart.AddNewPart<ExtendedChartPart>();
                    extChartPart.ChartSpace = cxChartSpace;
                    extChartPart.ChartSpace.Save();

                    // CONSISTENCY(chartex-sidecars): every Office-canonical
                    // chartEx part requires two sidecar parts linked via
                    // relationships: a ChartStylePart (chs:chartStyle) and a
                    // ChartColorStylePart (chs:colorStyle). Excel rejects
                    // files that have the chartEx body but lack these
                    // sidecars (silent "We found a problem" repair that
                    // DELETES the entire drawing containing the chart —
                    // slicers and all other anchors get collateral-damaged).
                    // The SDK validator doesn't flag this because each part
                    // is independently schema-valid; it's only the absence
                    // of the sidecar relationship that Excel trips on.
                    //
                    // chartStyle is built by ChartExStyleBuilder; an
                    // optional chartStyle=N prop on the caller picks a
                    // numbered style variant, default = 0.
                    var styleVariant = properties.GetValueOrDefault("chartStyle")
                                    ?? properties.GetValueOrDefault("chartstyle")
                                    ?? "default";
                    var stylePart = extChartPart.AddNewPart<ChartStylePart>();
                    using (var styleStream = ChartExStyleBuilder.BuildChartStyleXml(chartType, styleVariant))
                        stylePart.FeedData(styleStream);
                    var colorStylePart = extChartPart.AddNewPart<ChartColorStylePart>();
                    using (var colorStream = LoadChartExResource("chartex-colors.xml"))
                        colorStylePart.FeedData(colorStream);

                    var cxRelId = drawingsPart.GetIdOfPart(extChartPart);
                    var cxAnchor = new XDR.TwoCellAnchor();
                    cxAnchor.Append(new XDR.FromMarker(
                        new XDR.ColumnId(fromCol.ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId(fromRow.ToString()),
                        new XDR.RowOffset("0")));
                    cxAnchor.Append(new XDR.ToMarker(
                        new XDR.ColumnId(toCol.ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId(toRow.ToString()),
                        new XDR.RowOffset("0")));

                    var cxGraphicFrame = new XDR.GraphicFrame();
                    var cxExistingIds = drawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
                        .Select(p => (uint?)p.Id?.Value ?? 0u)
                        .DefaultIfEmpty(1u)
                        .Max();
                    var cxFrameId = cxExistingIds + 1;
                    cxGraphicFrame.NonVisualGraphicFrameProperties = new XDR.NonVisualGraphicFrameProperties(
                        new XDR.NonVisualDrawingProperties
                        {
                            Id = cxFrameId,
                            Name = chartTitle ?? "Chart"
                        },
                        new XDR.NonVisualGraphicFrameDrawingProperties()
                    );
                    cxGraphicFrame.Transform = new XDR.Transform(
                        new Drawing.Offset { X = 0, Y = 0 },
                        new Drawing.Extents { Cx = 0, Cy = 0 }
                    );

                    var cxChartRef = new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId { Id = cxRelId };
                    cxGraphicFrame.Append(new Drawing.Graphic(
                        new Drawing.GraphicData(cxChartRef)
                        {
                            Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex"
                        }
                    ));

                    cxAnchor.Append(cxGraphicFrame);
                    cxAnchor.Append(new XDR.ClientData());
                    drawingsPart.WorksheetDrawing.Append(cxAnchor);
                    drawingsPart.WorksheetDrawing.Save();

                    // Count all charts (both regular and extended)
                    var totalCharts = CountExcelCharts(drawingsPart);
                    return $"/{chartSheetName}/chart[{totalCharts}]";
                }

                // Build chart content BEFORE adding part (invalid type throws, must not leave empty part)
                var chartSpace = ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = chartSpace;
                chartPart.ChartSpace.Save();

                // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
                var deferredProps = properties
                    .Where(kv => ChartHelper.IsDeferredKey(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (deferredProps.Count > 0)
                    ChartHelper.SetChartProperties(chartPart, deferredProps);

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
                // Compute a unique cNvPr ID: use max existing ID + 1 to avoid duplicates after deletion
                var existingIds = drawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
                    .Select(p => (uint?)p.Id?.Value ?? 0u)
                    .DefaultIfEmpty(1u)
                    .Max();
                var chartFrameId = existingIds + 1;
                graphicFrame.NonVisualGraphicFrameProperties = new XDR.NonVisualGraphicFrameProperties(
                    new XDR.NonVisualDrawingProperties
                    {
                        Id = chartFrameId,
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

                // Legend is already handled inside BuildChartSpace

                var chartIdx = CountExcelCharts(drawingsPart);
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

                // R8-7: incidental whitespace around the source spec or its
                // components (" Sheet1 ! A1:D10 ") is a common paste-from-docs
                // artefact. Trim the whole string and both sides of the '!'
                // split so the downstream sheet/range lookup sees clean values.
                sourceSpec = sourceSpec.Trim();

                // R8-3: external workbook refs such as [other.xlsx]Sheet1!A1:D10
                // used to fall through to FindWorksheet and surface as the
                // misleading "Source sheet not found: [other.xlsx]Sheet1".
                // Detect the '[' prefix up front and throw a clear error so
                // users know the feature is not supported rather than blaming
                // a missing sheet.
                if (sourceSpec.StartsWith("["))
                    throw new ArgumentException(
                        "External workbook references are not supported in pivot source. "
                        + "Use a local sheet name (e.g. Sheet1!A1:D10)");

                string sourceSheetName;
                string sourceRef;
                if (sourceSpec.Contains('!'))
                {
                    var srcParts = sourceSpec.Split('!', 2);
                    sourceSheetName = srcParts[0].Trim().Trim('\'', '"').Trim();
                    sourceRef = srcParts[1].Trim();
                }
                else
                {
                    sourceSheetName = ptSheetName;
                    sourceRef = sourceSpec;
                }

                var sourceWorksheet = FindWorksheet(sourceSheetName)
                    ?? throw new ArgumentException($"Source sheet not found: {sourceSheetName}");

                var ptPosition = (properties.GetValueOrDefault("position", "")
                    ?? properties.GetValueOrDefault("pos", ""))
                    ?.Replace("$", ""); // CONSISTENCY(dollar-strip): parity with source ref handling
                if (string.IsNullOrEmpty(ptPosition))
                {
                    // Auto-position: place after the source data range
                    var rangeEnd = sourceRef.Split(':').Last();
                    var colEndMatch = System.Text.RegularExpressions.Regex.Match(rangeEnd, @"([A-Za-z]+)");
                    var nextCol = colEndMatch.Success ? IndexToColumnName(ColumnNameToIndex(colEndMatch.Value.ToUpperInvariant()) + 2) : "H";
                    ptPosition = $"{nextCol}1";
                }

                // R26-1: validate that the pivot output fits within sheet dimensions
                // before writing any cache/pivot parts. A position near the sheet edge
                // can produce an end-location beyond XFD1048576, which causes a
                // partial-write: cache parts are already saved when the render stage
                // discovers the overflow and throws, leaving a corrupt zip.
                {
                    const int ExcelMaxCol = 16384; // XFD
                    const int ExcelMaxRow = 1048576;
                    var srcRefParts = sourceRef.Replace("$", "").Split(':');
                    if (srcRefParts.Length == 2)
                    {
                        var (srcStartCol, srcStartRow) = ParseCellReference(srcRefParts[0].Trim().ToUpperInvariant());
                        var (srcEndCol, srcEndRow)     = ParseCellReference(srcRefParts[1].Trim().ToUpperInvariant());
                        int nSourceCols = ColumnNameToIndex(srcEndCol) - ColumnNameToIndex(srcStartCol) + 1;
                        int nDataRows   = srcEndRow - srcStartRow; // header excluded
                        var (anchorColStr, anchorRow) = ParseCellReference(ptPosition.ToUpperInvariant());
                        int anchorColIdx = ColumnNameToIndex(anchorColStr);
                        // Conservative lower-bound: pivot needs at least nSourceCols columns
                        // (row-label cols + value cols + grand-total col) and at least
                        // nDataRows + 2 rows (header + data rows + grand-total row).
                        int minEndColIdx = anchorColIdx + nSourceCols - 1;
                        int minEndRow    = anchorRow + nDataRows + 1;
                        if (minEndColIdx > ExcelMaxCol || minEndRow > ExcelMaxRow)
                        {
                            throw new ArgumentException(
                                $"pivot at {ptPosition} does not fit: computed end col={minEndColIdx} row={minEndRow} exceeds sheet dimensions (max XFD1048576)");
                        }
                    }
                }

                var ptIdx = PivotTableHelper.CreatePivotTable(
                    _doc.WorkbookPart!, ptWorksheet, sourceWorksheet,
                    sourceSheetName, sourceRef, ptPosition, properties);

                return $"/{ptSheetName}/pivottable[{ptIdx}]";
            }

            case "slicer":
            {
                return AddSlicer(parentPath, properties);
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
                {
                    ShiftColumnsRight(colWorksheet, insertColIdx);
                    DeleteCalcChainIfPresent();
                }

                // Optionally set column width (accepts bare char units or unit-qualified)
                if (properties.TryGetValue("width", out var widthStr) && !string.IsNullOrWhiteSpace(widthStr))
                {
                    var width = ParseColWidthChars(widthStr);
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
                    return Add(parentPath, "colbreak", position, properties);
                return Add(parentPath, "rowbreak", position, properties);
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
            case "cfextended":
            {
                var cfNewSegments = parentPath.TrimStart('/').Split('/', 2);
                var cfNewSheetName = cfNewSegments[0];
                var cfNewWorksheet = FindWorksheet(cfNewSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cfNewSheetName}");
                var cfNewSqref = properties.GetValueOrDefault("sqref") ?? properties.GetValueOrDefault("range") ?? properties.GetValueOrDefault("ref", "A1:A10");
                var cfNewPriority = NextCfPriority(GetSheet(cfNewWorksheet));

                ConditionalFormattingRule cfNewRule;
                var typeLower = type.ToLowerInvariant();
                // For cfextended dispatch, the actual requested sub-type is in
                // properties["type"] (the user-facing switch; the outer `type`
                // variable is literal "cfextended" here).
                if (typeLower == "cfextended")
                    typeLower = (properties.GetValueOrDefault("type", "") ?? "").ToLowerInvariant();

                switch (typeLower)
                {
                    case "topn":
                    {
                        // Accept both `rank=` (OOXML attribute name) and `top=`
                        // (user-facing alias documented in the topn help).
                        var rankStr = properties.GetValueOrDefault("rank")
                            ?? properties.GetValueOrDefault("top")
                            ?? properties.GetValueOrDefault("bottomN")
                            ?? "10";
                        var rank = uint.TryParse(rankStr, out var r) ? r : 10u;
                        var percent = ParseHelpers.IsTruthy(properties.GetValueOrDefault("percent", "false"));
                        var bottom = ParseHelpers.IsTruthy(properties.GetValueOrDefault("bottom", "false"));
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.Top10,
                            Priority = cfNewPriority,
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
                            Priority = cfNewPriority,
                            AboveAverage = ParseHelpers.IsTruthy(aboveBelow) ? null : false
                        };
                        break;
                    }
                    case "uniquevalues":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.UniqueValues,
                            Priority = cfNewPriority
                        };
                        break;
                    }
                    case "duplicatevalues":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.DuplicateValues,
                            Priority = cfNewPriority
                        };
                        break;
                    }
                    case "containstext":
                    {
                        var text = properties.GetValueOrDefault("text", "");
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.ContainsText,
                            Priority = cfNewPriority,
                            Text = text,
                            Operator = ConditionalFormattingOperatorValues.ContainsText
                        };
                        var firstCell = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"NOT(ISERROR(SEARCH(\"{text}\",{firstCell})))"));
                        break;
                    }
                    case "dateoccurring":
                    {
                        // Accept both `period=` (docs/canonical) and `timePeriod=`
                        // (OOXML attribute spelling) as input aliases.
                        var period = properties.GetValueOrDefault("period")
                            ?? properties.GetValueOrDefault("timePeriod")
                            ?? properties.GetValueOrDefault("timeperiod")
                            ?? "today";
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
                            Priority = cfNewPriority,
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
                    case "belowaverage":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.AboveAverage,
                            Priority = cfNewPriority,
                            AboveAverage = false
                        };
                        break;
                    }
                    case "containsblanks":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.ContainsBlanks,
                            Priority = cfNewPriority
                        };
                        var fc0 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"LEN(TRIM({fc0}))=0"));
                        break;
                    }
                    case "notcontainsblanks":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.NotContainsBlanks,
                            Priority = cfNewPriority
                        };
                        var fc1 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"LEN(TRIM({fc1}))>0"));
                        break;
                    }
                    case "containserrors":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.ContainsErrors,
                            Priority = cfNewPriority
                        };
                        var fc2 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"ISERROR({fc2})"));
                        break;
                    }
                    case "notcontainserrors":
                    {
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.NotContainsErrors,
                            Priority = cfNewPriority
                        };
                        var fc3 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"NOT(ISERROR({fc3}))"));
                        break;
                    }
                    case "contains":
                    {
                        var ctext = properties.GetValueOrDefault("text", "");
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.ContainsText,
                            Priority = cfNewPriority,
                            Text = ctext,
                            Operator = ConditionalFormattingOperatorValues.ContainsText
                        };
                        var fc4 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"NOT(ISERROR(SEARCH(\"{ctext}\",{fc4})))"));
                        break;
                    }
                    case "notcontains":
                    {
                        var nctext = properties.GetValueOrDefault("text", "");
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.NotContainsText,
                            Priority = cfNewPriority,
                            Text = nctext,
                            Operator = ConditionalFormattingOperatorValues.NotContains
                        };
                        var fc5 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"ISERROR(SEARCH(\"{nctext}\",{fc5}))"));
                        break;
                    }
                    case "beginswith":
                    {
                        var btext = properties.GetValueOrDefault("text", "");
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.BeginsWith,
                            Priority = cfNewPriority,
                            Text = btext,
                            Operator = ConditionalFormattingOperatorValues.BeginsWith
                        };
                        var fc6 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"LEFT({fc6},{btext.Length})=\"{btext}\""));
                        break;
                    }
                    case "endswith":
                    {
                        var etext = properties.GetValueOrDefault("text", "");
                        cfNewRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.EndsWith,
                            Priority = cfNewPriority,
                            Text = etext,
                            Operator = ConditionalFormattingOperatorValues.EndsWith
                        };
                        var fc7 = cfNewSqref.Split(':')[0].TrimStart('$');
                        cfNewRule.AppendChild(new Formula($"RIGHT({fc7},{etext.Length})=\"{etext}\""));
                        break;
                    }
                    default:
                        throw new ArgumentException($"Unsupported CF type: {typeLower}");
                }

                ApplyStopIfTrue(cfNewRule, properties);

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
                    _dirtyStylesheet = true;
                    cfNewRule.FormatId = cfNewDxfs.Count!.Value - 1;
                }

                var cfNewFormatting = new ConditionalFormatting(cfNewRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        cfNewSqref.Split(' ').Select(s => new StringValue(s)))
                };

                var cfNewWs = GetSheet(cfNewWorksheet);
                InsertConditionalFormatting(cfNewWs, cfNewFormatting);

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
                    ?? properties.GetValueOrDefault("data")
                    ?? throw new ArgumentException("Sparkline requires 'range' (or 'data') property (e.g. A1:E1)");

                // Determine sparkline type
                var spkTypeStr = properties.GetValueOrDefault("type", "line").ToLowerInvariant();
                var spkType = spkTypeStr switch
                {
                    "column" => X14.SparklineTypeValues.Column,
                    "stacked" or "winloss" or "win-loss" => X14.SparklineTypeValues.Stacked,
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

                // Ensure worksheet root declares mc:Ignorable="x14" so Excel opts-in
                // to the x14 extension namespace where sparklines live. Without this,
                // Excel silently drops the entire extLst block and no sparklines render.
                var spkWsRoot = spkWs;
                const string spkMcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                const string spkX14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
                if (spkWsRoot.LookupNamespace("mc") == null)
                    spkWsRoot.AddNamespaceDeclaration("mc", spkMcNs);
                if (spkWsRoot.LookupNamespace("x14") == null)
                    spkWsRoot.AddNamespaceDeclaration("x14", spkX14Ns);
                var spkIgnorable = spkWsRoot.MCAttributes?.Ignorable?.Value ?? "";
                if (!spkIgnorable.Split(' ').Contains("x14"))
                {
                    spkWsRoot.MCAttributes ??= new MarkupCompatibilityAttributes();
                    spkWsRoot.MCAttributes.Ignorable = string.IsNullOrEmpty(spkIgnorable) ? "x14" : $"{spkIgnorable} x14";
                }

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
                        "Valid types: sheet, row, cell, shape, chart, ole (object, embed), autofilter, databar, colorscale, iconset, formulacf, comment, namedrange, table, picture, validation, pivottable. " +
                        "Use 'officecli xlsx add' for details.");

                SaveWorksheet(fbWorksheet);

                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }


    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
        {
            // Move (reorder) the sheet within the workbook.
            // CONSISTENCY(move-anchor): mirrors PowerPointHandler.Move slide reorder —
            // supports --index / --after /Sheet2 / --before /Sheet3.
            var workbook = GetWorkbook();
            var sheets = workbook.GetFirstChild<Sheets>()
                ?? throw new InvalidOperationException("Workbook has no sheets element");
            var sheetEl = sheets.Elements<Sheet>().FirstOrDefault(s =>
                string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException($"Sheet not found: {sheetName}");

            // Resolve after/before anchor BEFORE removing sheetEl.
            static string ExtractAnchorSheetName(string raw) =>
                (raw.StartsWith("/") ? raw[1..] : raw).Split('/', 2)[0];

            Sheet? afterAnchor = null, beforeAnchor = null;
            if (position?.After != null)
            {
                var anchorName = ExtractAnchorSheetName(position.After);
                afterAnchor = sheets.Elements<Sheet>().FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, anchorName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"After anchor not found: {position.After}");
            }
            else if (position?.Before != null)
            {
                var anchorName = ExtractAnchorSheetName(position.Before);
                beforeAnchor = sheets.Elements<Sheet>().FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, anchorName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"Before anchor not found: {position.Before}");
            }
            else if (index == null)
            {
                throw new ArgumentException("One of --index, --after, or --before is required when moving a sheet");
            }

            sheetEl.Remove();

            if (afterAnchor != null)
            {
                afterAnchor.InsertAfterSelf(sheetEl);
            }
            else if (beforeAnchor != null)
            {
                beforeAnchor.InsertBeforeSelf(sheetEl);
            }
            else
            {
                var targetIndex = index!.Value;
                var sheetList = sheets.Elements<Sheet>().ToList();
                if (targetIndex >= 0 && targetIndex < sheetList.Count)
                    sheetList[targetIndex].InsertBeforeSelf(sheetEl);
                else
                    sheets.AppendChild(sheetEl);
            }
            workbook.Save();
            return $"/{sheetName}";
        }

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

    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
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
