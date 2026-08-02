// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        // Probe ifExists UNCONDITIONALLY: it is only acted on when the name
        // collides, but reading it inside that branch meant the tracking
        // dictionary reported it as unsupported_property on every clean
        // subtree-dump replay (the emitter always emits it).
        var claimIfExists = properties.TryGetValue("ifExists", out var ifExistsVal)
            && ifExistsVal.Equals("use", StringComparison.OrdinalIgnoreCase);
        var caseMatch = sheets.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        if (caseMatch != null)
        {
            // ifExists=use: claim the existing sheet as a success no-op.
            // Emitted by single-sheet subtree dumps, whose batch must replay
            // both onto a workbook lacking the sheet (create it) and back
            // onto one that already has it (merge into it) — the hard
            // duplicate-name error below broke the second case.
            if (claimIfExists)
                return $"/{caseMatch.Name}";
            // Distinguish the BlankDocCreator-shipped placeholder sheet
            // (untouched, claimable by the first Add) from a real
            // user-created sheet (collision is a genuine error). The
            // placeholder is identifiable as: workbook holds exactly one
            // sheet, that sheet's worksheet has empty SheetData, no
            // sheetView properties beyond defaults, no tabColor — i.e.
            // a fresh `Create blank → first Add` flow.
            var caseExact = string.Equals(caseMatch.Name, name, StringComparison.Ordinal);
            var isPlaceholder = sheets.Elements<Sheet>().Count() == 1
                && IsPristineWorksheet(workbookPart, caseMatch);
            // Placeholder claim is only meaningful when the caller actually
            // supplies a sheet-level prop that would mutate the placeholder
            // (autoFilter / tabColor / hidden). Without any such prop the
            // "claim" is a true no-op and indistinguishable from a duplicate-
            // name collision — reject so callers don't see fake success.
            var hasClaimableProp = properties.ContainsKey("autoFilter")
                || properties.ContainsKey("tabColor")
                || properties.ContainsKey("hidden");
            if (!caseExact || !isPlaceholder || !hasClaimableProp)
            {
                throw new ArgumentException(
                    $"A sheet named '{caseMatch.Name}' already exists. Sheet names must be unique.");
            }

            // Placeholder claim: route any supplied autoFilter / tabColor /
            // hidden through Set so the user's intent applies — the previous
            // silent no-op branch dropped them, which is what motivated
            // rejecting duplicates outright.
            var existingPart = (WorksheetPart)workbookPart.GetPartById(caseMatch.Id!);
            var sheetMerged = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (properties.TryGetValue("autoFilter", out var dupAf)) sheetMerged["autofilter"] = dupAf;
            if (properties.TryGetValue("tabColor", out var dupTc)) sheetMerged["tabcolor"] = dupTc;
            if (sheetMerged.Count > 0)
                SetSheetLevel(existingPart, name, sheetMerged);
            if (properties.TryGetValue("hidden", out var dupHidden) && ParseHelpers.IsTruthy(dupHidden))
                caseMatch.State = SheetStateValues.Hidden;
            return $"/sheet[@name='{name}']";
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

            // localSheetId on <definedName> is a 0-based position into
            // <sheets>; inserting mid-list shifts every sheet at/after the
            // insert point up by one, so scoped names (printArea, print
            // titles, scoped named ranges) must shift with them or they
            // silently rebind to the sheet now occupying the old position.
            var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                foreach (var dn in definedNames.Elements<DefinedName>())
                {
                    var lid = dn.LocalSheetId?.Value;
                    if (lid.HasValue && lid.Value >= (uint)pos)
                        dn.LocalSheetId = lid.Value + 1;
                }
            }
        }
        else
        {
            sheets.AppendChild(newSheet);
        }

        // Add/Set symmetry (the project conventions): apply autoFilter / tabColor / hidden
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

    /// <summary>
    /// Returns true when the worksheet behind <paramref name="sheet"/> looks
    /// like the BlankDocCreator placeholder: empty SheetData, no tabColor,
    /// no autoFilter, default visibility. Used by AddSheet to decide whether
    /// a duplicate-name Add is the legacy "claim the blank's auto-Sheet1"
    /// pattern (idempotent) or a genuine user collision (throw).
    /// </summary>
    private static bool IsPristineWorksheet(WorkbookPart workbookPart, Sheet sheet)
    {
        if (sheet.State != null && sheet.State.Value != SheetStateValues.Visible) return false;
        if (sheet.Id?.Value == null) return false;
        if (workbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart wsp) return false;
        var ws = wsp.Worksheet;
        if (ws == null) return false;
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData != null && sheetData.Elements<Row>().Any()) return false;
        var props = ws.GetFirstChild<SheetProperties>();
        if (props?.GetFirstChild<TabColor>() != null) return false;
        if (ws.Descendants<AutoFilter>().Any()) return false;
        return true;
    }

    private string AddRow(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var segments = parentPath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? GetSheet(worksheet).AppendChild(new SheetData());

        // Resolve --before / --after anchors (same shape as Excel CopyFrom):
        // anchor must be /<sheetName>/row[K] in the same sheet.
        // CONSISTENCY(zero-based-index): per project convention, position.Index
        // is 0-based across all formats (--index 0 = head, --index 1 = before
        // 2nd slot). xlsx Row uses a 1-based RowIndex internally, so +1 here
        // and let the existing branch keep treating `index` as a 1-based row
        // number (which is also what the anchor branch below produces).
        int? index = position?.Index.HasValue == true ? position!.Index!.Value + 1 : (int?)null;
        if (index == null && position != null && (position.After != null || position.Before != null))
        {
            int FindAnchorRow(string anchorPath)
            {
                var aSegs = anchorPath.TrimStart('/').Split('/', 2);
                if (aSegs.Length < 2)
                    throw new ArgumentException(
                        $"Anchor must be a row path like /{sheetName}/row[K], got: {anchorPath}");
                if (!aSegs[0].Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException(
                        $"Anchor sheet '{aSegs[0]}' must match target sheet '{sheetName}'");
                var am = Regex.Match(aSegs[1], @"^row\[(\d+)\]$");
                if (!am.Success)
                    throw new ArgumentException(
                        $"Anchor must be a row path like /{sheetName}/row[K], got: {anchorPath}");
                return (int)uint.Parse(am.Groups[1].Value);
            }
            // For row insertion, --before /Sheet1/row[5] means "the new row
            // takes the row[5] slot, original row[5] shifts to row[6]". So
            // resolved index == anchor row number. --after /Sheet1/row[5]
            // means index == anchor + 1.
            if (position.Before != null) index = FindAnchorRow(position.Before);
            else index = FindAnchorRow(position.After!) + 1;
        }

        var rowIdx = index ?? ((int)(sheetData.Elements<Row>().LastOrDefault()?.RowIndex?.Value ?? 0) + 1);

        // Excel's row space tops out at 1048576 (2^20). The append branch
        // above silently produced row[1048577+] when row[1048576] already
        // existed, writing a file Excel rejects on open. Mirror the Set
        // path's bound check (ExcelHandler.Set.cs row index guard) so the
        // overflow surfaces as a clean invalid_value at Add time.
        if (rowIdx < 1 || rowIdx > 1048576)
            throw new ArgumentException(
                $"Invalid row index {rowIdx}. Valid row range is 1-1048576.");

        // If inserting at an existing position, shift everything at/below it
        // down. Gate only on "inserting at a position" (index set), NOT on the
        // presence of cell data at/below — sheet-level structures (CF / merge /
        // dataValidation) anchored on still-empty cells must shift too. Mirrors
        // AddCol, which calls ShiftColumnsRight on every positional insert
        // (CONSISTENCY(add-row-col-shift)). When nothing sits at/below rowIdx this
        // is a harmless no-op.
        // Validate all props BEFORE the structural shift (same atomicity rule
        // as AddCol): a height/outline parse failure after ShiftRowsDown left
        // the shift applied even though the add reported an error.
        double? parsedRowHeight = null;
        if (properties.TryGetValue("height", out var addRowHeight) && !string.IsNullOrWhiteSpace(addRowHeight))
            parsedRowHeight = ParseRowHeightPoints(addRowHeight);
        byte? parsedRowOutline = null;
        if (properties.TryGetValue("outline", out var addRowOutline)
            || properties.TryGetValue("outlinelevel", out addRowOutline)
            || properties.TryGetValue("group", out addRowOutline))
        {
            if (!byte.TryParse(addRowOutline, out var addRowOutlineVal) || addRowOutlineVal > 7)
                throw new ArgumentException($"Invalid 'outline' value: '{addRowOutline}'. Expected an integer 0-7 (outline/group level).");
            parsedRowOutline = addRowOutlineVal;
        }

        bool needsShift = index.HasValue;
        if (needsShift)
            ShiftRowsDown(worksheet, rowIdx);

        var newRow = new Row { RowIndex = (uint)rowIdx };

        // CONSISTENCY(add-set-symmetry): accept height/hidden at creation
        // time, mirroring SetRow semantics (ExcelHandler.Set.cs L3157-3164).
        if (parsedRowHeight is { } rh)
        {
            newRow.Height = rh;
            newRow.CustomHeight = true;
        }
        if (properties.TryGetValue("hidden", out var addRowHidden))
        {
            newRow.Hidden = addRowHidden.Equals("true", StringComparison.OrdinalIgnoreCase)
                || addRowHidden == "1" || addRowHidden.Equals("yes", StringComparison.OrdinalIgnoreCase);
        }
        // CONSISTENCY(add-set-symmetry): accept outline/group + collapsed at
        // creation, mirroring SetRow (ExcelHandler.Set.cs L2823-2832).
        if (parsedRowOutline is { } rowOutlineVal)
            newRow.OutlineLevel = rowOutlineVal;
        if (properties.TryGetValue("collapsed", out var addRowCollapsed))
        {
            newRow.Collapsed = addRowCollapsed.Equals("true", StringComparison.OrdinalIgnoreCase)
                || addRowCollapsed == "1" || addRowCollapsed.Equals("yes", StringComparison.OrdinalIgnoreCase);
        }

        // Create cells if cols specified
        if (properties.TryGetValue("cols", out var colsStr))
        {
            if (!int.TryParse(colsStr, out var cols) || cols <= 0)
                throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Expected a positive integer (number of columns to create).");
            // CONSISTENCY(table-row-cN): pptx AddRow accepts c1=/c2=/... to
            // populate the new row's cells (PowerPointHandler.Add.Table.cs
            // L332). Mirror it here so xlsx `add row --prop cols=N c1=...`
            // is a one-shot row create + fill instead of needing N follow-up
            // cell Sets. Only materialize a <c> when the caller actually
            // supplied content for that column — pre-emitting empty <c r=...>
            // shells would diverge from Excel's stored form (empty cells are
            // simply absent) and make Get("/Sheet/An") report "" instead of
            // "(empty)".
            for (int c = 0; c < cols; c++)
            {
                if (!properties.TryGetValue($"c{c + 1}", out var cellText) || cellText == null)
                    continue;
                var colLetter = IndexToColumnName(c + 1);
                EnsureCellValueLength(cellText, $"{colLetter}{rowIdx}");
                var safe = OfficeCli.Core.PivotTableHelper.SanitizeXmlText(cellText);
                var newCell = new Cell
                {
                    CellReference = $"{colLetter}{rowIdx}",
                    CellValue = new CellValue(safe),
                    DataType = new EnumValue<CellValues>(CellValues.String),
                };
                newRow.AppendChild(newCell);
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
        // R10-2: also honor a cell[<ref>] path tail (e.g. /Sheet1/cell[C5]) so
        // `add /Sheet1/cell[C5] cell` lands at C5 instead of silently snapping
        // to A1. Mirrors the bare-cellref tail above and the row[N] tail below;
        // without it, "cell[C5]" matched neither regex and auto-assign chose A1.
        else if (cellSegments.Length > 1)
        {
            var cellPathMatch = Regex.Match(cellSegments[1], @"^cell\[([A-Z]+\d+)\]$", RegexOptions.IgnoreCase);
            if (cellPathMatch.Success)
                cellRefFromPath = cellPathMatch.Groups[1].Value.ToUpperInvariant();
        }

        // BUG-R41-B6: also honor a row[N] path tail (e.g. /Sheet1/row[5]) so
        // `add /Sheet1/row[5] cell` lands on row 5 instead of silently snapping
        // to row 1. Without this, the row[N] segment was completely ignored:
        // the auto-assign branch below always picked row 1, and `--prop ref=A1`
        // overrode the row index too. Encode the row-from-path as a 1-based
        // row index and apply it later wherever a row choice is made.
        uint? rowIndexFromPath = null;
        if (cellSegments.Length > 1)
        {
            var rowPathMatch = Regex.Match(cellSegments[1], @"^row\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (rowPathMatch.Success)
                rowIndexFromPath = uint.Parse(rowPathMatch.Groups[1].Value);
        }

        // BUG-R2: an unrecognized parent path tail (e.g. r[2], foo[2], xyz[5])
        // matched none of the three accepted forms above (bare cell-ref,
        // cell[<ref>], row[N]), leaving both cellRefFromPath and rowIndexFromPath
        // null. The auto-assign branch below then silently snapped to row 1,
        // misplacing data without warning. Reject it instead. (Bare /Sheet1 with
        // cellSegments.Length == 1 is unaffected — that's a legitimate
        // auto-append target.) No r[N] alias is added: Get/Query don't support
        // r[N], so accepting it on Add would break Add/Get symmetry.
        if (cellSegments.Length > 1
            && cellRefFromPath == null
            && rowIndexFromPath == null)
        {
            throw new ArgumentException(
                $"Unrecognized cell parent path segment '{cellSegments[1]}'. " +
                $"Expected a cell reference (e.g. A2), cell[A2], or row[N].");
        }

        string cellRef;
        // BUG-R36-B1: when --prop arrayformula= is supplied with --prop ref=A1:C3,
        // the range is the spill region, not a single cell address. Detect it and
        // resolve cellRef to the top-left so FindOrCreateCell doesn't reject the
        // colon. The full range is still passed through to arrayformula below via
        // properties["ref"].
        string? arrayFormulaRefRange = null;
        if (properties.ContainsKey("ref"))
        {
            cellRef = properties["ref"];
            if (cellRef.Contains(':') && properties.ContainsKey("arrayformula"))
            {
                arrayFormulaRefRange = cellRef;
                var topLeft = cellRef.Split(':', 2)[0];
                if (!Regex.IsMatch(topLeft, @"^[A-Z]+\d+$", RegexOptions.IgnoreCase))
                    throw new ArgumentException($"Invalid cell reference: '{cellRef}'");
                cellRef = topLeft.ToUpperInvariant();
            }
            if (cellRefFromPath != null && !cellRefFromPath.Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                Console.Error.WriteLine($"warning: path tail '{cellRefFromPath}' does not match --prop ref='{properties["ref"]}'; using ref='{properties["ref"]}'.");
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
            // BUG-R41-B6: if the parent path supplies a row index (/Sheet1/row[5]),
            // auto-assign within that row instead of always defaulting to row 1.
            var targetRow = rowIndexFromPath ?? 1;
            var existingRefs = cellSheetData.Descendants<Cell>()
                .Where(c => c.CellReference?.Value != null)
                .Select(c => c.CellReference!.Value!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            int colIdx = 1;
            while (existingRefs.Contains(IndexToColumnName(colIdx) + targetRow))
                colIdx++;
            cellRef = IndexToColumnName(colIdx) + targetRow;
        }

        // BUG-R41-B6: if both /Sheet1/row[N] and an explicit ref/address (or
        // path-tail cell-ref) were supplied, the row index in the address
        // wins, but warn when they disagree so the operator notices.
        if (rowIndexFromPath.HasValue)
        {
            var refRowMatch = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
            if (refRowMatch.Success && uint.Parse(refRowMatch.Groups[2].Value) != rowIndexFromPath.Value)
                Console.Error.WriteLine(
                    $"warning: path row[{rowIndexFromPath.Value}] does not match cell ref '{cellRef}' row; using ref's row.");
        }
        // --prop shift=right|down: before materializing the new cell, push
        // existing cells in the same row (right) or column (down) by 1.
        // Mirrors Excel UI's "Insert Cells > Shift cells right / down".
        // Same scope cap as RemoveCellWithShift: only intra-row/col cellRefs
        // are rewritten — formulas, mergeCells, CF/DV/hyperlinks/tables that
        // span the affected row/col are NOT adjusted. For full row/col insert
        // with all relations, use add --type row / --type col.
        if (properties.TryGetValue("shift", out var shiftVal) && !string.IsNullOrEmpty(shiftVal))
        {
            var shiftDir = shiftVal.ToLowerInvariant();
            if (shiftDir is not ("right" or "down"))
                throw new ArgumentException(
                    $"--prop shift={shiftVal} not valid for add cell. Use 'right' or 'down'.");
            var (shiftCol, shiftRow) = ParseCellReference(cellRef);
            var shiftColIdx = ColumnNameToIndex(shiftCol);
            if (shiftDir == "right")
                ShiftCellsRightInRow(cellSheetData, (uint)shiftRow, shiftColIdx);
            else
                ShiftCellsDownInColumn(cellSheetData, shiftCol, shiftRow);
        }

        // Atomicity: validate a type=boolean value BEFORE FindOrCreateCell
        // appends the cell to the sheet. A throw AFTER the cell is created
        // used to leave a corrupt <c t="b"><v>garbage</v></c> persisted on
        // disk (real Excel then refuses the file, 0x800A03EC) even though the
        // Add reported an error. The later in-switch check stays as a
        // defense-in-depth guard.
        {
            var upfrontType = properties.GetValueOrDefault("type")?.ToLowerInvariant();
            var upfrontValue = (properties.GetValueOrDefault("value")
                ?? properties.GetValueOrDefault("text"))?.Trim().ToLowerInvariant();
            if ((upfrontType is "boolean" or "bool") && !string.IsNullOrEmpty(upfrontValue)
                && upfrontValue is not ("true" or "false" or "yes" or "no" or "1" or "0"))
                throw new ArgumentException(
                    $"Cannot store '{properties.GetValueOrDefault("value") ?? properties.GetValueOrDefault("text")}' as boolean; " +
                    "value must be true/false, yes/no, or 1/0. Use type=string to keep the literal text.");
        }

        // Atomicity: FindOrCreateCell materializes a <c> stub if the cell did
        // not exist. A validation throw further down (bad textRotation, bad
        // color, bad merge ref, ...) must not leave that stub — or the value
        // already written into it — persisted while the command reports
        // Error/exit 1. Capture pre-existence, then roll the new cell back on
        // any throw. Mirrors the Set-side rollback (ExcelHandler.Set.cs).
        var cellPreExisted = cellSheetData.Elements<Row>()
            .SelectMany(r => r.Elements<Cell>())
            .Any(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));

        var cell = FindOrCreateCell(cellSheetData, cellRef);
        // Clone for rollback of a pre-existing cell (restore original state);
        // a newly created cell is removed instead (see catch below).
        var cellBackup = cell.CloneNode(true);

        try
        {

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
            // Auto-detect formula: value starting with '=' is treated as
            // formula — UNLESS type=string was supplied (explicitly, or forced
            // by the apostrophe branch above). Mirrors the Set-path gate; see
            // ExcelHandler.Set.cs case "value".
            var addForcedString = properties.TryGetValue("type", out var addTypeVal)
                && addTypeVal.Equals("string", StringComparison.OrdinalIgnoreCase);
            if (!addForcedString && value.StartsWith('=') && value.Length > 1)
            {
                RejectCrossWorkbookFormula(value);
                ValidateFormulaCellRefs(value);
                // CONSISTENCY(value-child-uniqueness): a <c> may hold at most
                // one value child. Drop any stale <is> placeholder (table
                // header cells emit <is><t>ColumnN</t></is>) before writing
                // the formula; a cell with both <f>/<v> and <is> is invalid
                // OOXML that real Excel rejects with 0x800A03EC.
                cell.RemoveAllChildren<InlineString>();
                cell.CellFormula = new CellFormula(Core.PivotTableHelper.SanitizeXmlText(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(value.TrimStart('=')))));
                cell.CellValue = null;
                // CONSISTENCY(cell-formula-calc): Set and import both stamp
                // fullCalcOnLoad AND evaluate/cache the result when writing a
                // formula; Add doing neither left a bare <f> whose Get type
                // diverged from the replayed file (String vs Error).
                EnsureFullCalcOnLoad();
                {
                    var addEvalSd = GetSheet(cellWorksheet).GetFirstChild<SheetData>();
                    if (addEvalSd != null)
                        WriteFormulaResultToCell(cell,
                            new Core.FormulaEvaluator(addEvalSd, _doc.WorkbookPart).TryEvaluateFull(value.TrimStart('=')));
                }
            }
            else
            {
                // CONSISTENCY(formula-stale): writing a literal value must
                // clear any prior CellFormula on the same cell. Otherwise
                // the old formula re-evaluates on open / in html preview
                // and overrides the literal the caller just set.
                cell.CellFormula = null;
                // CONSISTENCY(value-child-uniqueness): also drop any stale
                // <is> inline-string child (table header placeholders write
                // <is><t>ColumnN</t></is>). A cell carrying both <v> and <is>
                // is invalid OOXML — real Excel refuses to open it (0x800A03EC).
                cell.RemoveAllChildren<InlineString>();
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
                // R-bt-1: an ISO-date value combined with an explicit date-like
                // numberformat and no type= means the user unambiguously wants
                // a date cell. Without coercion the numberformat is inert on a
                // t="str" cell — displays/sorts/filters as text in real Excel.
                // Import's type detection already coerces ISO dates; mirror it
                // here for this intent-clear case only (bare ISO strings
                // without a format keep their existing string semantics).
                if (!properties.ContainsKey("type")
                    && (properties.GetValueOrDefault("numberformat")
                        ?? properties.GetValueOrDefault("numfmt")
                        ?? properties.GetValueOrDefault("format")) is { } nfRaw
                    && System.Text.RegularExpressions.Regex.IsMatch(nfRaw, "[ymdhs]", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    && TryParseIsoDateFlexible(safeValue.Trim(), out var inferredDate)
                    && inferredDate >= new System.DateTime(1900, 1, 1))
                {
                    cell.CellValue = new CellValue(
                        ExcelDataFormatter.ToExcelSerial(inferredDate, IsWorkbookDate1904()).ToString(System.Globalization.CultureInfo.InvariantCulture));
                    cell.DataType = null;
                }
                else if (!double.TryParse(safeValue, out var dbl) || !double.IsFinite(dbl))
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                else
                    // R-fuzz2-1: TryParse accepts spellings Excel's <v> parser
                    // does not ("+5", "1,234", padded). Store the canonical
                    // form; literal digits are preserved when already canonical.
                    cell.CellValue = new CellValue(NormalizeNumericCellText(safeValue, dbl));
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
            ValidateFormulaCellRefs(fTrim);
            var addCellFormula = new CellFormula(Core.PivotTableHelper.SanitizeXmlText(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(fTrim))));
            // Dynamic-array functions (SORT/FILTER/UNIQUE/SEQUENCE/XLOOKUP/LET/etc.)
            // carry t="array" ref="<cellRef>" on the cell-level CellFormula PLUS a
            // cm cell-metadata index into an XLDAPR record (EnsureDynamicArrayMetadata)
            // — t="array" alone is a legacy CSE array locked to the anchor; the
            // XLDAPR metadata is what makes Excel 365 spill. The anchor reference is
            // the single cell being written; Excel recomputes the spill extent and
            // fills adjacent cells at runtime.
            if (Core.ModernFunctionQualifier.IsDynamicArrayFormula(fTrim) && cell.CellReference?.Value != null)
            {
                addCellFormula.FormulaType = CellFormulaValues.Array;
                addCellFormula.Reference = cell.CellReference.Value;
                EnsureDynamicArrayMetadata(cell);
            }
            // CONSISTENCY(value-child-uniqueness): clear any stale <is> so the
            // cell never carries both a formula and an inline string (invalid
            // OOXML, 0x800A03EC in real Excel).
            cell.RemoveAllChildren<InlineString>();
            cell.CellFormula = addCellFormula;
            cell.CellValue = null;
            EnsureFullCalcOnLoad(); // CONSISTENCY(cell-formula-calc): see value-branch note.
            {
                var addEvalSd2 = GetSheet(cellWorksheet).GetFirstChild<SheetData>();
                if (addEvalSd2 != null)
                    WriteFormulaResultToCell(cell,
                        new Core.FormulaEvaluator(addEvalSd2, _doc.WorkbookPart).TryEvaluateFull(fTrim));
            }
        }
        // CE1: allow `runs=<json>` without an explicit `type=richtext`.
        if (!properties.ContainsKey("type") && properties.ContainsKey("runs"))
            properties["type"] = "richtext";
        if (properties.TryGetValue("type", out var cellType))
        {
            if (cellType.Equals("richtext", StringComparison.OrdinalIgnoreCase) ||
                cellType.Equals("rich", StringComparison.OrdinalIgnoreCase))
            {
                ApplyRichTextToCell(cell, properties);
            }
            else
            {
                // Validate a boolean retype BEFORE mutating DataType. When the
                // cell already holds text and type=boolean arrives with no new
                // value, the switch below would stamp t="b" onto that text and
                // only the later check would throw — leaving a corrupt
                // <c t="b"><v>hello</v></c> Excel refuses (0x800A03EC). The
                // R114 upfront guard only sees the incoming value=, not the
                // existing cell text, so guard that here too.
                if ((cellType.Equals("boolean", StringComparison.OrdinalIgnoreCase)
                        || cellType.Equals("bool", StringComparison.OrdinalIgnoreCase))
                    && cell.CellValue?.Text?.Trim().ToLowerInvariant() is { Length: > 0 } existingBool
                    && existingBool is not ("true" or "false" or "yes" or "no" or "1" or "0"))
                    throw new ArgumentException(
                        $"Cannot store '{cell.CellValue?.Text}' as boolean; value must be true/false, yes/no, or 1/0. " +
                        "Use type=string to keep the literal text.");
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
                    else if (!string.IsNullOrEmpty(boolText))
                        // A t="b" cell whose value isn't 0/1 makes real Excel
                        // refuse the whole file (0x800A03EC). Reject up front,
                        // mirroring the type=date guard.
                        throw new ArgumentException(
                            $"Cannot store '{cell.CellValue?.Text}' as boolean; value must be true/false, yes/no, or 1/0. " +
                            "Use type=string to keep the literal text.");
                }
                // A type=number cell stores its value in <v> with no t=
                // attribute, so a non-numeric value produces spec-invalid
                // numeric content (<v>notanumber</v>) that makes real Excel
                // refuse the whole file (0x800A03EC) while schema validation
                // stays green. Reject up front, mirroring the boolean/date
                // guards above.
                if (cellType.ToLowerInvariant() is "number" or "num")
                {
                    var numText = cell.CellValue?.Text?.Trim();
                    if (!string.IsNullOrEmpty(numText)
                        && (!double.TryParse(numText, System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var numDbl)
                            || !double.IsFinite(numDbl)))
                        throw new ArgumentException(
                            $"Cannot store '{cell.CellValue?.Text}' as number; value must be a finite numeric literal. " +
                            "Use type=string to keep the literal text.");
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
                        // Mirrors Set's pre-1900 guard: Excel's serial epoch is
                        // 1899-12-30; earlier dates round-trip as the epoch and
                        // mislead the user. Reject instead of silently clamping.
                        if (dt < new System.DateTime(1900, 1, 1))
                            throw new ArgumentException(
                                $"Cannot store '{dateText}' as date; Excel does not support dates before 1900-01-01 " +
                                $"(serial epoch is 1899-12-30). Use type=string to keep the literal text.");
                        cell.CellValue = new CellValue(
                            ExcelDataFormatter.ToExcelSerial(dt, IsWorkbookDate1904()).ToString(System.Globalization.CultureInfo.InvariantCulture));
                    }
                    else if (!string.IsNullOrEmpty(dateText))
                    {
                        // BUG-FIX(B10): if user said type=date but the value isn't
                        // parseable, refuse to leave a date-shaped string in a
                        // numeric-styled cell — that produces invalid OOXML.
                        throw new ArgumentException(
                            $"Cannot store '{dateText}' as date; value must be ISO 8601 (yyyy-MM-dd) " +
                            $"and represent a real calendar day. Use type=string to keep the literal text.");
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

        // R8-3: phonetic guides (Japanese furigana, CJK ruby). The cell's
        // base text is promoted into the shared-string table with an <rPh>
        // child carrying the phonetic reading; the worksheet's default
        // <phoneticPr> is created if absent. Stamps a single phonetic run
        // spanning the entire base text (sb=0 / eb=len) — sufficient for
        // the canonical use case (one reading per cell). Multi-segment
        // phonetic runs are out of scope for the minimum viable surface;
        // callers that need them can submit raw OOXML through extension
        // attrs in a follow-up.
        if (properties.TryGetValue("phonetic", out var phoneticText)
            && !string.IsNullOrEmpty(phoneticText))
        {
            ApplyPhoneticToCell(cell, cellWorksheet, phoneticText, properties);
        }

        // Array formula support during Add
        if (properties.TryGetValue("arrayformula", out var arrFormula))
        {
            // arrayformula=true|1|yes is flag intent ("make my formula= an
            // array formula"), not formula text. Writing it verbatim replaced
            // the real formula with the literal string "true" — silent
            // corruption. Substitute the companion formula= text; without one
            // there is nothing to convert, so reject clearly.
            if (arrFormula.Equals("true", StringComparison.OrdinalIgnoreCase)
                || arrFormula == "1"
                || arrFormula.Equals("yes", StringComparison.OrdinalIgnoreCase))
            {
                arrFormula = properties.GetValueOrDefault("formula")
                    ?? cell.CellFormula?.Text
                    ?? throw new ArgumentException(
                        "arrayformula=true requires a formula: pass the text directly (arrayformula=\"B1:B3*C1:C3\") or combine with formula=.");
            }
            RejectCrossWorkbookFormula(arrFormula);
            ValidateFormulaCellRefs(arrFormula);
            // BUG-R36-B1: if ref was a range (A1:C3), use the full range as
            // arrRef so the array formula spills correctly; otherwise default
            // to the single cellRef.
            var arrRef = arrayFormulaRefRange ?? properties.GetValueOrDefault("ref", cellRef);
            // CONSISTENCY(value-child-uniqueness): drop any stale <is> placeholder
            // so the cell holds a single value child (invalid otherwise).
            cell.RemoveAllChildren<InlineString>();
            cell.CellFormula = new CellFormula(Core.PivotTableHelper.SanitizeXmlText(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(arrFormula.TrimStart('=')))))
            {
                FormulaType = CellFormulaValues.Array,
                Reference = arrRef
            };
            EnsureFullCalcOnLoad(); // CONSISTENCY(cell-formula-calc)
            cell.CellValue = null;
        }

        // Hyperlink support during Add
        if ((properties.TryGetValue("link", out var linkUrl) || properties.TryGetValue("url", out linkUrl)) && !string.IsNullOrEmpty(linkUrl))
        {
            // Validate the scheme BEFORE creating the <hyperlinks> container
            // (same fix as the Set path): a rejected scheme used to leave an
            // empty schema-invalid <x:hyperlinks/> behind — Excel 0x800A03EC.
            var addLinkIsInternal = ResolveInternalHyperlinkLocation(linkUrl) != null;
            if (!addLinkIsInternal)
                Core.HyperlinkUriValidator.RequireSafeScheme(linkUrl, "link");
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
            // H2b: display (OOXML @display) — friendly text Excel shows for
            // the link. Handler-as-truth: consumed here so it is not reported
            // unsupported (schema hyperlink.json documents cell display=).
            var hlDisplay = properties.GetValueOrDefault("display");
            var addInternalLoc = ResolveInternalHyperlinkLocation(linkUrl);
            if (addInternalLoc != null)
            {
                var hl = new Hyperlink
                {
                    Reference = cellRef.ToUpperInvariant(),
                    Location = addInternalLoc
                };
                if (!string.IsNullOrEmpty(hlTip)) hl.Tooltip = hlTip;
                if (!string.IsNullOrEmpty(hlDisplay)) hl.Display = hlDisplay;
                hyperlinksEl.AppendChild(hl);
            }
            else
            {
                // CONSISTENCY(hyperlink-scheme-allowlist): see Set.cs cell link.
                Core.HyperlinkUriValidator.RequireSafeScheme(linkUrl, "link");
                var hlUri = new Uri(linkUrl, UriKind.RelativeOrAbsolute);
                var hlRel = cellWorksheet.AddHyperlinkRelationship(hlUri, isExternal: true);
                var hl = new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id };
                if (!string.IsNullOrEmpty(hlTip))
                    hl.Tooltip = hlTip;
                if (!string.IsNullOrEmpty(hlDisplay)) hl.Display = hlDisplay;
                hyperlinksEl.AppendChild(hl);
            }
        }

        // In-cell image ("Place in Cell" richValue) during Add — parity with
        // the Set cell `image=` case (ExcelHandler.Set.Cells.cs).
        if (properties.TryGetValue("image", out var inCellImg) && !string.IsNullOrEmpty(inCellImg)
            && !inCellImg.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            // CONSISTENCY(picture-alt): same alt aliases as the picture element.
            var inCellAlt = properties.GetValueOrDefault("alt")
                ?? properties.GetValueOrDefault("altText")
                ?? properties.GetValueOrDefault("alttext")
                ?? properties.GetValueOrDefault("description")
                ?? properties.GetValueOrDefault("image.alt");
            if (inCellAlt != null) Core.ParseHelpers.ValidateXmlText(inCellAlt, "alt");
            SetInCellImage(cell, inCellImg, inCellAlt);
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

        // Apply style properties if any. Use TryGetValue per key so the
        // TrackingPropertyDictionary comparer marks each style key as
        // accessed — bare foreach over the upcast Dictionary<,> base type
        // bypasses the recording GetEnumerator override and leaves
        // legitimately-consumed keys (bold, align, color, ...) reported
        // as UNSUPPORTED while their values silently take effect.
        var cellStyleProps = new Dictionary<string, string>();
        foreach (var key in properties.Keys.ToList())
        {
            if (ExcelStyleManager.IsStyleKey(key) && properties.TryGetValue(key, out var val))
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

        // DATA-CORRUPTION(xlsx/table-header-name): if this write landed on a
        // table header cell, keep <tableColumn name> in sync with the header
        // text — Excel rejects the file otherwise.
        MaybeSyncTableHeaderName(cellWorksheet, cellRef);

        // R20-02: accept `merge=A1:C3` on cell Add (parity with `set`).
        // This is the same merge logic used by Set range action; we
        // apply it post-creation so users can merge in a single Add
        // call instead of needing a follow-up set.
        if (properties.TryGetValue("merge", out var mergeRange) && !string.IsNullOrWhiteSpace(mergeRange))
        {
            var refList = mergeRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            // CONSISTENCY(merge-empty-container): pre-validate before container
            // creation — see ExcelHandler.Helpers.ValidateMergeRefLiteral.
            foreach (var r in refList) ValidateMergeRefLiteral(r);
            var sheetEl = GetSheet(cellWorksheet);
            var mergeCellsEl = sheetEl.GetFirstChild<MergeCells>();
            if (mergeCellsEl == null)
            {
                mergeCellsEl = new MergeCells();
                sheetEl.AppendChild(mergeCellsEl);
            }
            // CONSISTENCY(merge-comma): comma in *prop value* is the supported
            // batch form (here, in cell Set, and in sheet Set) — split into
            // separate <mergeCell> elements. Comma in *path* is rejected by
            // InsertMergeCellChecked since path is a single-target locator.
            foreach (var rangeRef in refList)
                InsertMergeCellChecked(mergeCellsEl, rangeRef, cellWorksheet);
            mergeCellsEl.Count = (uint)mergeCellsEl.Elements<MergeCell>().Count();
        }

        DeleteCalcChainIfPresent();
        SaveWorksheet(cellWorksheet);
        return $"/{cellSheetName}/{cellRef}";
        }
        catch
        {
            if (cellPreExisted)
            {
                // Restore the pre-existing cell to its original state so a
                // failed Add makes no partial change (mirrors Set-side rollback).
                cell.Parent?.ReplaceChild(cellBackup, cell);
            }
            else
            {
                // Newly created by this Add — remove the stub (and its now-empty
                // row) so a failed create leaves no ghost cell/value behind.
                var newRow = cell.Parent as Row;
                cell.Remove();
                if (newRow != null && !newRow.Elements<Cell>().Any())
                {
                    var sd = newRow.Parent as SheetData;
                    var rIdx = newRow.RowIndex?.Value;
                    newRow.Remove();
                    if (sd != null && rIdx.HasValue)
                        _rowIndex?.GetValueOrDefault(sd)?.Remove(rIdx.Value);
                }
            }
            throw;
        }
    }

    private string AddCol(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var colSegments = parentPath.TrimStart('/').Split('/', 2);
        var colSheetName = colSegments[0];
        var colWorksheet = FindWorksheet(colSheetName)
            ?? throw new ArgumentException($"Sheet not found: {colSheetName}");

        // Resolve --before / --after anchors, mirroring AddRow. Anchor must
        // be /<sheetName>/col[L] in the same sheet; --before takes the
        // anchor's slot, --after lands one column to the right.
        int? index = position?.Index;
        if (index == null && position != null && (position.After != null || position.Before != null))
        {
            int FindAnchorColIndex(string anchorPath)
            {
                var aSegs = anchorPath.TrimStart('/').Split('/', 2);
                if (aSegs.Length < 2)
                    throw new ArgumentException(
                        $"Anchor must be a column path like /{colSheetName}/col[L], got: {anchorPath}");
                if (!aSegs[0].Equals(colSheetName, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException(
                        $"Anchor sheet '{aSegs[0]}' must match target sheet '{colSheetName}'");
                var am = Regex.Match(aSegs[1], @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
                if (!am.Success)
                    throw new ArgumentException(
                        $"Anchor must be a column path like /{colSheetName}/col[L], got: {anchorPath}");
                return ColumnNameToIndex(am.Groups[1].Value.ToUpperInvariant());
            }
            if (position.Before != null) index = FindAnchorColIndex(position.Before);
            else index = FindAnchorColIndex(position.After!) + 1;
        }

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

        // A positional insert ALWAYS shifts existing data + metadata right,
        // exactly like AddRow's ShiftRowsDown gate (needsShift = index.HasValue).
        // The earlier guard skipped the shift whenever a single-column <col>
        // already sat at insertColIdx — but that is true for ANY column merely
        // carrying a stored width, so a formatted worksheet's `add col` silently
        // dropped the shift and corrupted data. Insert is a purely structural
        // shift, never gated on stored column width (standard spreadsheet
        // semantics); mutating width/hidden in place is the job of
        // `set /Sheet/col[X]`.
        // CONSISTENCY(add-row-col-shift): mirror AddRow's positional-insert gate.
        // Validate EVERY property BEFORE the structural shift: a parse failure
        // after ShiftColumnsRight left the shift applied ("Error" + data moved
        // one column right anyway) and an empty <cols/> shell on disk —
        // schema-invalid (cols requires >= 1 col child), so a REJECTED add
        // corrupted a previously-fine file (0x800A03EC in real Excel).
        bool hasColWidth = properties.TryGetValue("width", out var widthStr) && !string.IsNullOrWhiteSpace(widthStr);
        double parsedColWidth = hasColWidth ? ParseColWidthChars(widthStr!) : 0;
        bool hasColHidden = properties.TryGetValue("hidden", out var addColHidden);
        byte? parsedColOutline = null;
        if (properties.TryGetValue("outline", out var addColOutline)
            || properties.TryGetValue("outlinelevel", out addColOutline)
            || properties.TryGetValue("group", out addColOutline))
        {
            if (!byte.TryParse(addColOutline, out var addColOutlineVal) || addColOutlineVal > 7)
                throw new ArgumentException($"Invalid 'outline' value: '{addColOutline}'. Expected an integer 0-7 (outline/group level).");
            parsedColOutline = addColOutlineVal;
        }

        bool colNeedsShift = index.HasValue || !string.IsNullOrEmpty(colLetterProp);
        if (colNeedsShift)
        {
            ShiftColumnsRight(colWorksheet, insertColIdx);
            DeleteCalcChainIfPresent();
        }

        // CONSISTENCY(add-set-symmetry): always materialize a <col> element so
        // Get/Query can find the column even when no width/hidden was supplied.
        // Width/Hidden are attached only when the caller provides them.
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
                newCol.Width = parsedColWidth;
                newCol.CustomWidth = true;
            }
            else if (newCol.Width == null)
            {
                // A <col> with no width attribute renders ZERO-width in real
                // Excel (verified via remote render: the inserted column and
                // its cells visually vanish while Get/dump still see them).
                // Materializing for Get/Query symmetry must therefore carry
                // the sheet's default width explicitly.
                newCol.Width = ws.SheetFormatProperties?.DefaultColumnWidth?.Value
                    ?? ws.SheetFormatProperties?.BaseColumnWidth?.Value + 0.43
                    ?? 8.43;
            }
            if (hasColHidden)
            {
                newCol.Hidden = addColHidden!.Equals("true", StringComparison.OrdinalIgnoreCase)
                    || addColHidden == "1" || addColHidden.Equals("yes", StringComparison.OrdinalIgnoreCase);
            }
            // CONSISTENCY(add-set-symmetry): accept outline/group + collapsed at
            // creation, mirroring SetColumn (ExcelHandler.Set.cs L2624-2632).
            if (parsedColOutline is { } outlineVal)
                newCol.OutlineLevel = outlineVal;
            if (properties.TryGetValue("collapsed", out var addColCollapsed))
            {
                newCol.Collapsed = addColCollapsed.Equals("true", StringComparison.OrdinalIgnoreCase)
                    || addColCollapsed == "1" || addColCollapsed.Equals("yes", StringComparison.OrdinalIgnoreCase);
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
            // Converting a plain-string / inline-string / numeric cell to rich
            // text: preserve the cell's existing content as the first unstyled
            // run instead of silently discarding it. Without this, adding the
            // first run to a cell that already had a value threw the value away
            // (e.g. value="Hello World" + add run "Hi" → cell became just "Hi").
            string? runExistingText = runCell.DataType?.Value == CellValues.InlineString
                ? runCell.InlineString?.InnerText
                : runCell.CellValue?.Text;

            runSsi = new SharedStringItem();
            if (!string.IsNullOrEmpty(runExistingText))
                runSsi.AppendChild(new Run(
                    new Text(runExistingText) { Space = SpaceProcessingModeValues.Preserve }));
            runCell.RemoveAllChildren<InlineString>();
            runSst.AppendChild(runSsi);
            var newSstIdx = runSst.Elements<SharedStringItem>().Count() - 1;
            runCell.CellValue = new CellValue(newSstIdx.ToString());
            runCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        var newRun = new Run();
        var newRunProps = new RunProperties();
        var runText = properties.GetValueOrDefault("text", "");
        OfficeCli.Core.ParseHelpers.ValidateXmlText(runText, "run text");

        // CONSISTENCY(tracking-dict): read each prop via TryGetValue (not a
        // foreach over the dictionary). The foreach went through
        // IEnumerable.GetEnumerator on the Dictionary<> static type, which does
        // NOT fire TrackingPropertyDictionary's shadow GetEnumerator — so applied
        // keys like bold/italic were never marked accessed and surfaced as a
        // false unsupported_property (exit 2). See the project conventions tracking pitfalls.
        // Each helper accepts the short key plus its font.* alias.
        if ((properties.TryGetValue("bold", out var rBold) && ParseHelpers.IsTruthy(rBold)) ||
            (properties.TryGetValue("font.bold", out var rFBold) && ParseHelpers.IsTruthy(rFBold)))
            newRunProps.AppendChild(new Bold());
        if ((properties.TryGetValue("italic", out var rItalic) && ParseHelpers.IsTruthy(rItalic)) ||
            (properties.TryGetValue("font.italic", out var rFItalic) && ParseHelpers.IsTruthy(rFItalic)))
            newRunProps.AppendChild(new Italic());
        if (properties.TryGetValue("strike", out var rStrike) && ParseHelpers.IsTruthy(rStrike))
            newRunProps.AppendChild(new Strike());
        if (properties.TryGetValue("underline", out var rUl)
            && !string.IsNullOrEmpty(rUl) && rUl != "false" && rUl != "none")
        {
            var ul = new Underline();
            if (rUl.ToLowerInvariant() == "double") ul.Val = UnderlineValues.Double;
            newRunProps.AppendChild(ul);
        }
        if (properties.TryGetValue("superscript", out var rSup) && ParseHelpers.IsTruthy(rSup))
            newRunProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript });
        if (properties.TryGetValue("subscript", out var rSub) && ParseHelpers.IsTruthy(rSub))
            newRunProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript });
        if ((properties.TryGetValue("size", out var rSize) ||
             properties.TryGetValue("fontsize", out rSize) ||
             properties.TryGetValue("font.size", out rSize))
            && double.TryParse(rSize.TrimEnd('p', 't'), out var runSz))
            newRunProps.AppendChild(new FontSize { Val = runSz });
        if (properties.TryGetValue("color", out var rColor) ||
            properties.TryGetValue("font.color", out rColor))
            newRunProps.AppendChild(new Color { Rgb = new HexBinaryValue(ParseHelpers.NormalizeArgbColor(rColor)) });
        if (properties.TryGetValue("font", out var rFont) ||
            properties.TryGetValue("fontname", out rFont) ||
            properties.TryGetValue("font.name", out rFont))
            newRunProps.AppendChild(new RunFont { Val = rFont });
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
        // A break id of 0 or beyond the grid fails the schema's Min/Max
        // constraints — reject up front instead of writing invalid OOXML.
        if (rbRowIdx < 1 || rbRowIdx > 1048576)
            throw new ArgumentException(
                $"Invalid 'row' value: '{rbRowIdx}'. Row breaks must be between 1 and 1048576.");

        var rowBreaks = rbWs.GetFirstChild<RowBreaks>();
        if (rowBreaks == null)
        {
            rowBreaks = new RowBreaks();
            rbWs.AppendChild(rowBreaks);
        }
        // Optional restricted column span (min/max) — mirrors the Set path so a
        // dump-emitted `add rowbreak row=N min=.. max=..` reproduces a
        // non-full-width break. Defaults to full width (max 16383) when absent.
        var rbBreak = new Break { Id = rbRowIdx, Max = 16383u, ManualPageBreak = true };
        if (properties.TryGetValue("min", out var rbMinS) && uint.TryParse(rbMinS, out var rbMin))
            rbBreak.Min = rbMin;
        if (properties.TryGetValue("max", out var rbMaxS) && uint.TryParse(rbMaxS, out var rbMax))
            rbBreak.Max = rbMax;
        if (properties.TryGetValue("manual", out var rbMan))
            rbBreak.ManualPageBreak = IsTruthy(rbMan);
        rowBreaks.AppendChild(rbBreak);
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
        // Same schema Min/Max guard as rowbreak: 0 / beyond-XFD ids write
        // invalid OOXML that only surfaces at validate/open time.
        if (cbColIdx < 1 || cbColIdx > 16384)
            throw new ArgumentException(
                $"Invalid 'col' value: '{cbColStr}'. Column breaks must be between 1 and 16384 (A-XFD).");

        var colBreaks = cbWs.GetFirstChild<ColumnBreaks>();
        if (colBreaks == null)
        {
            colBreaks = new ColumnBreaks();
            cbWs.AppendChild(colBreaks);
        }
        // Optional restricted row span (min/max) — mirrors the Set path.
        var cbBreak = new Break { Id = cbColIdx, Max = 1048575u, ManualPageBreak = true };
        if (properties.TryGetValue("min", out var cbMinS) && uint.TryParse(cbMinS, out var cbMin))
            cbBreak.Min = cbMin;
        if (properties.TryGetValue("max", out var cbMaxS) && uint.TryParse(cbMaxS, out var cbMax))
            cbBreak.Max = cbMax;
        if (properties.TryGetValue("manual", out var cbMan))
            cbBreak.ManualPageBreak = IsTruthy(cbMan);
        colBreaks.AppendChild(cbBreak);
        colBreaks.Count = (uint)colBreaks.Elements<Break>().Count();
        colBreaks.ManualBreakCount = colBreaks.Count;
        SaveWorksheet(cbWorksheet);

        var cbBrkIdx = colBreaks.Elements<Break>().ToList()
            .FindIndex(b => b.Id?.Value == cbColIdx) + 1;
        return $"/{cbSheetName}/colbreak[{cbBrkIdx}]";
    }

    /// <summary>
    /// Build a SharedString rich-text entry for <paramref name="cell"/> from
    /// `runs=<JSON array>` or legacy `run1=text:prop=val;…` syntax. Reused by
    /// Add (when the user passes type=richtext) and by Set (so type=richtext
    /// is symmetric — see CONSISTENCY(cell-type-parity)).
    /// </summary>
    private void ApplyRichTextToCell(Cell cell, Dictionary<string, string> properties)
    {
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
                    OfficeCli.Core.ParseHelpers.ValidateXmlText(text, "richtext run text");
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

        // Drop empty-text runs that carry formatting props: real Excel refuses
        // the whole workbook (0x800A03EC) when a formatting-only empty <r> is
        // not the last run in the <si>. They render nothing, so dropping is
        // lossless; keep one if removing all would leave an empty <si>.
        var effectiveRuns = gatheredRuns
            .Where(r => !(string.IsNullOrEmpty(r.text) && r.props.Count > 0)).ToList();
        if (effectiveRuns.Count == 0 && gatheredRuns.Count > 0)
            effectiveRuns.Add(gatheredRuns[^1]);

        foreach (var (runText, pd) in effectiveRuns)
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

    /// <summary>
    /// Stamp a phonetic guide (furigana / CJK ruby) on a cell. Promotes the
    /// cell's base text into the shared-string table (existing SST entry
    /// is reused when one with the same base text is found) and appends an
    /// <c>&lt;rPh&gt;</c> run carrying the phonetic reading. Also seeds
    /// the worksheet's <c>&lt;phoneticPr&gt;</c> default block — without
    /// it, Excel suppresses the rendered guide regardless of what the
    /// SSI contains. R8-3.
    /// </summary>
    private void ApplyPhoneticToCell(Cell cell, WorksheetPart wsPart,
        string phoneticText, Dictionary<string, string> properties)
    {
        // 1) Resolve the cell's base text.
        string baseText;
        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(cell.CellValue?.Text, out var existingIdx))
        {
            var existingSstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var existingSsi = existingSstPart?.SharedStringTable?
                .Elements<SharedStringItem>().ElementAtOrDefault(existingIdx);
            baseText = existingSsi?.Text?.Text
                ?? string.Concat(existingSsi?.Elements<Run>().Select(r => r.Text?.Text ?? "")
                    ?? Enumerable.Empty<string>());
        }
        else
        {
            baseText = cell.CellValue?.Text ?? "";
        }
        if (string.IsNullOrEmpty(baseText))
            throw new ArgumentException(
                "phonetic requires a non-empty cell value (the base text the phonetic guide annotates).");

        // 2) Build a fresh SSI: <si><t>baseText</t><rPh sb=0 eb=len><t>phonetic</t></rPh></si>
        var wbPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("Workbook not found");
        var sstPart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
            ?? wbPart.AddNewPart<SharedStringTablePart>();
        var sst = sstPart.SharedStringTable ??= new SharedStringTable();

        var ssi = new SharedStringItem(
            new Text(baseText) { Space = SpaceProcessingModeValues.Preserve });
        var rPh = new PhoneticRun(
                new Text(phoneticText) { Space = SpaceProcessingModeValues.Preserve })
        {
            BaseTextStartIndex = 0u,
            EndingBaseIndex = (uint)baseText.Length,
        };
        ssi.AppendChild(rPh);

        sst.AppendChild(ssi);
        sst.Count = (uint)sst.Elements<SharedStringItem>().Count();
        sst.UniqueCount = sst.Count;

        var newIdx = sst.Elements<SharedStringItem>().Count() - 1;
        cell.CellValue = new CellValue(newIdx.ToString());
        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        // 3) Ensure the worksheet has a <phoneticPr> block — Excel only
        // renders <rPh> when the worksheet supplies a default font / type.
        var ws = GetSheet(wsPart);
        if (ws.GetFirstChild<PhoneticProperties>() == null)
        {
            var phoneticPr = new PhoneticProperties
            {
                FontId = 1u,
                Type = PhoneticValues.FullWidthKatakana,
                Alignment = PhoneticAlignmentValues.Distributed,
            };
            // Schema position: phoneticPr lives between mergeCells and
            // conditionalFormatting (CT_Worksheet — see ordering comment in
            // ExcelHandler.Set.cs:2004). Use the schema-aware sheet child
            // inserter rather than a plain AppendChild.
            InsertPhoneticPropertiesInOrder(ws, phoneticPr);
        }
    }

    /// <summary>
    /// Insert a <c>&lt;phoneticPr&gt;</c> at its CT_Worksheet schema slot.
    /// Predecessors (mergeCells / customSheetViews / dataConsolidate /
    /// sortState / autoFilter / scenarios / protectedRanges /
    /// sheetProtection / sheetCalcPr / sheetData) come before; successors
    /// (conditionalFormatting / dataValidations / hyperlinks / printOptions /
    /// pageMargins / pageSetup / drawing / etc.) come after.
    /// </summary>
    private static void InsertPhoneticPropertiesInOrder(Worksheet ws, PhoneticProperties pr)
    {
        OpenXmlElement? after = null;
        Type[] preds =
        [
            typeof(SheetData),
            typeof(SheetCalculationProperties),
            typeof(SheetProtection),
            typeof(ProtectedRanges),
            typeof(Scenarios),
            typeof(AutoFilter),
            typeof(SortState),
            typeof(DataConsolidate),
            typeof(CustomSheetViews),
            typeof(MergeCells),
        ];
        foreach (var t in preds)
        {
            var hit = ws.ChildElements.FirstOrDefault(c => c.GetType() == t);
            if (hit != null) after = hit; // last match wins — schema-latest predecessor
        }
        if (after != null) ws.InsertAfter(pr, after);
        else ws.PrependChild(pr);
    }
}
