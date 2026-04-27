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

// Per-element-type Add helpers for cell-grid paths (sheet, row, cell, col, run, page/row/col-breaks). Mechanically extracted from the Add() god-method.
public partial class ExcelHandler
{
    private string AddSheet(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("Workbook not found");
        var sheets = GetWorkbook().GetFirstChild<Sheets>()
            ?? GetWorkbook().AppendChild(new Sheets());

        var name = properties.GetValueOrDefault("name", $"Sheet{sheets.Elements<Sheet>().Count() + 1}");
        // CONSISTENCY(sheet-name-validation): mirror Set's name validation
        // (ExcelHandler.Set.cs L1777) so Add and Set both reject names Excel
        // would refuse to open. Only validate when explicitly user-supplied —
        // the auto-generated SheetN default is always safe.
        if (properties.ContainsKey("name"))
            ValidateSheetName(name);
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

        // Add/Set symmetry (CLAUDE.md): apply autoFilter / tabColor / hidden
        // at creation time by funneling into the same code paths Set uses,
        // so property bags accepted by Set are also accepted by Add.
        var sheetLevelForwarded = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (properties.TryGetValue("autoFilter", out var addAf)) sheetLevelForwarded["autofilter"] = addAf;
        if (properties.TryGetValue("tabColor", out var addTc)) sheetLevelForwarded["tabcolor"] = addTc;
        if (sheetLevelForwarded.Count > 0)
            SetSheetLevel(newWorksheetPart, name, sheetLevelForwarded);

        // Sheet-state (hidden) lives on the workbook-level Sheet element,
        // not on the Worksheet, so it can't route through SetSheetLevel.
        if (properties.TryGetValue("hidden", out var addHidden) && ParseHelpers.IsTruthy(addHidden))
            newSheet.State = SheetStateValues.Hidden;

        GetWorkbook().Save();
        return $"/{name}";
    }

    private string AddRow(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
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

        // CONSISTENCY(add-set-symmetry): accept height/hidden at creation
        // time, mirroring SetRow semantics (ExcelHandler.Set.cs L3157-3164).
        if (properties.TryGetValue("height", out var addRowHeight) && !string.IsNullOrWhiteSpace(addRowHeight))
        {
            newRow.Height = ParseRowHeightPoints(addRowHeight);
            newRow.CustomHeight = true;
        }
        if (properties.TryGetValue("hidden", out var addRowHidden))
        {
            newRow.Hidden = addRowHidden.Equals("true", StringComparison.OrdinalIgnoreCase)
                || addRowHidden == "1" || addRowHidden.Equals("yes", StringComparison.OrdinalIgnoreCase);
        }

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

        // R33-2: this AddRow mutated sheetData directly (bypassing
        // FindOrCreateRow). If the row-index cache was already populated
        // by a prior cell op on the same sheet, it now lacks the new row
        // — a subsequent AddCell at the same row index would cache-miss
        // and create a duplicate <x:row r="N">, producing an
        // Excel-rejected file. Invalidate the cache to force a rescan.
        InvalidateRowIndex(sheetData);

        if (needsShift)
            DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return $"/{sheetName}/row[{rowIdx}]";
    }

    private string AddCell(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
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

        // CONSISTENCY(cell-value-alias): Set accepts "text" as alias for
        // "value" (see WordHandler.Set cell text handling); mirror that here.
        if (!properties.ContainsKey("value") && properties.TryGetValue("text", out var textAlias))
            properties["value"] = textAlias;

        if (properties.TryGetValue("value", out var value))
        {
            // R28-B4 — leading apostrophe is Excel's "force text" idiom.
            // Strip the apostrophe and stamp quotePrefix=1 on the cell xf.
            // Mirrors the Set path; see ExcelHandler.Set.cs case "value".
            if (value.StartsWith('\'') && value.Length > 1)
            {
                value = value.Substring(1);
                properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase)
                {
                    ["value"] = value,
                    ["quoteprefix"] = "true",
                };
                if (!properties.ContainsKey("type"))
                    properties["type"] = "string";
            }
            // R13-1: reject values longer than Excel's 32767-char limit
            // before doing any conversion/serialization.
            EnsureCellValueLength(value, cellRef);
            // R13-3: if both value= and formula= are supplied, formula wins
            // (established precedence — formula is written after value) but
            // the discarded value is easy to miss. Warn on stderr.
            if (properties.ContainsKey("formula"))
            {
                Console.Error.WriteLine(
                    "Warning: Both value= and formula= supplied — using formula, value ignored.");
            }
            // Auto-detect formula: value starting with '=' is treated as formula
            if (value.StartsWith('=') && value.Length > 1)
            {
                RejectCrossWorkbookFormula(value);
                cell.CellFormula = new CellFormula(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(value.TrimStart('='))));
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
                // R32-1: double.TryParse("NaN") returns true with double.NaN,
                // which would write <c><v>NaN</v></c> with no t= — invalid
                // xs:double content that crashes Excel. Force string type for
                // any non-finite double (NaN/Infinity), matching the
                // already-string behavior of "Infinity"/"-Infinity" (which
                // TryParse rejects under default culture).
                if (!double.TryParse(safeValue, out var dbl) || !double.IsFinite(dbl))
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
            RejectCrossWorkbookFormula(fTrim);
            cell.CellFormula = new CellFormula(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(fTrim)));
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
                    // R13-2: accept ISO date-with-time (T separator) as well.
                    if (!string.IsNullOrEmpty(dateText)
                        && TryParseIsoDateFlexible(dateText, out var dt))
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
            RejectCrossWorkbookFormula(arrFormula);
            var arrRef = properties.GetValueOrDefault("ref", cellRef);
            cell.CellFormula = new CellFormula(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(arrFormula.TrimStart('='))))
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
            // H2: tooltip (OOXML @tooltip) — Excel surfaces it as a
            // ScreenTip when the cell is hovered in read mode.
            var hlTip = properties.GetValueOrDefault("tooltip")
                ?? properties.GetValueOrDefault("screenTip")
                ?? properties.GetValueOrDefault("screentip");
            // R37-B: detect internal `[#]Sheet!Cell` (and quoted variants);
            // emit as @location with no relationship.
            // CONSISTENCY(internal-hyperlink): same detection used in Set.cs.
            var addInternalLoc = TryParseInternalHyperlinkLocation(linkUrl);
            if (addInternalLoc != null)
            {
                var hl = new Hyperlink
                {
                    Reference = cellRef.ToUpperInvariant(),
                    Location = addInternalLoc
                };
                if (!string.IsNullOrEmpty(hlTip)) hl.Tooltip = hlTip;
                hyperlinksEl.AppendChild(hl);
            }
            else
            {
                var hlUri = new Uri(linkUrl, UriKind.RelativeOrAbsolute);
                var hlRel = cellWorksheet.AddHyperlinkRelationship(hlUri, isExternal: true);
                var hl = new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id };
                if (!string.IsNullOrEmpty(hlTip))
                    hl.Tooltip = hlTip;
                hyperlinksEl.AppendChild(hl);
            }
        }

        // CONSISTENCY(cell-prop-hints): mirror Set's CellPropHints check
        // here. Before the style filter runs, flag any ambiguous flat
        // keys (e.g. `color` — is it font.color or fill?) as unsupported.
        // Without this, Add silently drops the key while Set loudly
        // rejects it — inconsistent, and the caller's intent is lost.
        var cellHintMessages = new List<string>();
        foreach (var (key, _) in properties)
        {
            var hint = CellPropHints.TryGetHint(key);
            if (hint != null)
                cellHintMessages.Add(hint);
        }
        if (cellHintMessages.Count > 0)
            throw new ArgumentException(
                "Unsupported cell property: " + string.Join("; ", cellHintMessages));

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

            // R24-1: when caller explicitly chose the text number format ("@"),
            // force the cell into String storage so leading zeros and any
            // non-numeric content survive the round-trip. Without this, a
            // value like "00456" gets written as <x:v>00456</x:v> with no
            // t="str" and Excel reparses it as 456 on open.
            if (IsTextNumberFormat(cellStyleProps)
                && cell.DataType?.Value != CellValues.SharedString
                && cell.DataType?.Value != CellValues.InlineString
                && cell.CellFormula == null)
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
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

        // R20-02: accept `merge=A1:C3` on cell Add (parity with `set`).
        // This is the same merge logic used by Set range action; we
        // apply it post-creation so users can merge in a single Add
        // call instead of needing a follow-up set.
        if (properties.TryGetValue("merge", out var mergeRange) && !string.IsNullOrWhiteSpace(mergeRange))
        {
            var sheetEl = GetSheet(cellWorksheet);
            var mergeCellsEl = sheetEl.GetFirstChild<MergeCells>();
            if (mergeCellsEl == null)
            {
                mergeCellsEl = new MergeCells();
                sheetEl.AppendChild(mergeCellsEl);
            }
            foreach (var rangeRef in mergeRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var existing = mergeCellsEl.Elements<MergeCell>()
                    .Any(mc => string.Equals(mc.Reference?.Value, rangeRef, StringComparison.OrdinalIgnoreCase));
                if (!existing)
                    mergeCellsEl.AppendChild(new MergeCell { Reference = rangeRef });
            }
            mergeCellsEl.Count = (uint)mergeCellsEl.Elements<MergeCell>().Count();
        }

        DeleteCalcChainIfPresent();
        SaveWorksheet(cellWorksheet);
        return $"/{cellSheetName}/{cellRef}";
    }

    private string AddCol(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var colSegments = parentPath.TrimStart('/').Split('/', 2);
        var colSheetName = colSegments[0];
        var colWorksheet = FindWorksheet(colSheetName)
            ?? throw new ArgumentException($"Sheet not found: {colSheetName}");

        // Determine insert column: index (1-based) or name/letter from properties
        // CONSISTENCY(col-letter-prop): accept col=, letter=, column= as aliases of name=
        // matching how `colbreak` (case "colbreak" above) accepts col/column/index.
        string insertColName;
        string? colLetterProp = null;
        if (properties.TryGetValue("name", out var colNameProp) && !string.IsNullOrEmpty(colNameProp))
            colLetterProp = colNameProp;
        else if (properties.TryGetValue("col", out var colProp) && !string.IsNullOrEmpty(colProp))
            colLetterProp = colProp;
        else if (properties.TryGetValue("letter", out var letterProp) && !string.IsNullOrEmpty(letterProp))
            colLetterProp = letterProp;
        else if (properties.TryGetValue("column", out var columnProp) && !string.IsNullOrEmpty(columnProp))
            colLetterProp = columnProp;

        if (!string.IsNullOrEmpty(colLetterProp))
        {
            // Accept either column letter (e.g. "B") or numeric index (e.g. "2")
            insertColName = uint.TryParse(colLetterProp, out var colNumIdx)
                ? IndexToColumnName((int)colNumIdx)
                : colLetterProp.ToUpperInvariant();
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

        // CONSISTENCY(add-set-symmetry): always materialize a <col> element so
        // Get/Query can find the column even when no width/hidden was supplied.
        // Width/Hidden are attached only when the caller provides them.
        bool hasColWidth = properties.TryGetValue("width", out var widthStr) && !string.IsNullOrWhiteSpace(widthStr);
        bool hasColHidden = properties.TryGetValue("hidden", out var addColHidden);
        {
            var ws = GetSheet(colWorksheet);
            var columns = ws.GetFirstChild<Columns>() ?? ws.PrependChild(new Columns());
            // Idempotent: if a Column with exact Min==Max==insertColIdx already exists,
            // update it rather than appending a duplicate.
            var existingCol = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min?.Value == (uint)insertColIdx && c.Max?.Value == (uint)insertColIdx);
            var newCol = existingCol ?? new Column
            {
                Min = (uint)insertColIdx,
                Max = (uint)insertColIdx
            };
            if (hasColWidth)
            {
                newCol.Width = ParseColWidthChars(widthStr!);
                newCol.CustomWidth = true;
            }
            if (hasColHidden)
            {
                newCol.Hidden = addColHidden!.Equals("true", StringComparison.OrdinalIgnoreCase)
                    || addColHidden == "1" || addColHidden.Equals("yes", StringComparison.OrdinalIgnoreCase);
            }
            if (existingCol == null)
                columns.AppendChild(newCol);
        }

        SaveWorksheet(colWorksheet);
        return $"/{colSheetName}/col[{insertColName}]";
    }

    private string AddRun(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
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

    private string AddPageBreak(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // Route to rowbreak or colbreak based on properties
        if (properties.ContainsKey("col") || properties.ContainsKey("column"))
            return Add(parentPath, "colbreak", position, properties);
        return Add(parentPath, "rowbreak", position, properties);
    }

    private string AddRowBreak(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
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

    private string AddColBreak(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
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

}
