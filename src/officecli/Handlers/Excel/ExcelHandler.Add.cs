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

public partial class ExcelHandler
{
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        Modified = true;
        var index = position?.Index;
        // Normalize to case-insensitive lookup so camelCase keys (e.g. minColor) match lowercase lookups.
        // Preserve TrackingPropertyDictionary so handler-as-truth read
        // tracking survives — its comparer wraps OrdinalIgnoreCase already.
        if (properties != null
            && properties is not OfficeCli.Core.TrackingPropertyDictionary
            && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        parentPath = NormalizeExcelPath(parentPath);
        parentPath = ResolveSheetIndexInPath(parentPath);

        // Reject element types that belong to a different document format up
        // front. Without this, "add /xlsx --type slide" fell into AddDefault,
        // tried to resolve the parent sheet from an empty path segment, and
        // produced a misleading "Sheet not found: " error with an empty name.
        // Naming the wrong-format type explicitly tells the caller where to
        // look (e.g. use the .pptx variant) instead of sending them on a
        // sheet-permission hunt.
        var typeLower = type.ToLowerInvariant();
        if (typeLower is "slide" or "slidemaster" or "slidelayout" or "notes"
            or "paragraph" or "p" or "field"
            or "section" or "header" or "footer")
        {
            var sourceFormat = typeLower switch
            {
                "slide" or "slidemaster" or "slidelayout" or "notes" => "pptx",
                "paragraph" or "p" or "field" => "docx/pptx",
                "section" or "header" or "footer" => "docx",
                _ => "another format"
            };
            throw new ArgumentException(
                $"Invalid element type '{type}' for xlsx files (belongs to {sourceFormat}). " +
                "Valid values: sheet, row, cell, col, namedrange, comment, validation, autofilter, " +
                "cf, databar, colorscale, iconset, formulacf, cellis, ole, picture, shape, slicer, " +
                "sparkline, table, chart, pivottable.");
        }

        switch (type.ToLowerInvariant())
        {
            case "sheet":
                return AddSheet(parentPath, type, position, properties);

            case "row":
                return AddRow(parentPath, type, position, properties);

            case "cell":
                return AddCell(parentPath, type, position, properties);

            case "namedrange" or "definedname" or "name":
                return AddNamedRange(parentPath, type, position, properties);

            case "comment" or "note":
                return AddComment(parentPath, type, position, properties);

            case "validation":
            case "datavalidation":
                return AddValidation(parentPath, type, position, properties);

            case "autofilter":
                return AddAutoFilter(parentPath, type, position, properties);

            case "cf":
                return AddCf(parentPath, type, position, properties);

            case "databar":
            case "conditionalformatting":
                return AddDataBar(parentPath, type, position, properties);

            case "colorscale":
                return AddColorScale(parentPath, type, position, properties);

            case "iconset":
                return AddIconSet(parentPath, type, position, properties);

            case "formulacf":
                return AddFormulaCf(parentPath, type, position, properties);

            case "cellis":
                return AddCellIs(parentPath, type, position, properties);

            case "ole":
            case "oleobject":
            case "object":
            case "embed":
                return AddOle(parentPath, type, position, properties);

            case "picture":
            case "image":
            case "img":
                return AddPicture(parentPath, type, position, properties);

            case "shape" or "textbox":
                return AddShape(parentPath, type, position, properties);

            case "table" or "listobject":
                return AddTable(parentPath, type, position, properties);

            case "chart":
                return AddChart(parentPath, type, position, properties);

            // BUG-002: schema chart-series.json declares operations.add on a
            // chart parent, but only pptx had a dispatch entry ("series").
            // Accept the schema element name and the pptx alias alike.
            case "series" or "chart-series" or "chartseries":
                return AddChartSeries(parentPath, properties);

            case "pivottable" or "pivot":
                return AddPivotTable(parentPath, type, position, properties);

            case "slicer":
                return AddSlicer(parentPath, type, position, properties);

            case "col" or "column":
                return AddCol(parentPath, type, position, properties);

            case "pagebreak":
                return AddPageBreak(parentPath, type, position, properties);

            case "rowbreak":
                return AddRowBreak(parentPath, type, position, properties);

            case "colbreak":
                return AddColBreak(parentPath, type, position, properties);

            case "run":
                return AddRun(parentPath, type, position, properties);

            case "topn":
            case "aboveaverage":
            case "uniquevalues":
            case "duplicatevalues":
            case "containstext":
            case "dateoccurring":
            case "cfextended":
                return AddCfExtended(parentPath, type, position, properties);

            case "sparkline":
                return AddSparkline(parentPath, type, position, properties);

            default:
                return AddDefault(parentPath, type, position, properties);
        }
    }

    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position, Dictionary<string, string>? properties = null)
    {
        // xlsx has no track-change concept; `properties` is accepted for IDocumentHandler parity but ignored.
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

            // Self-move guard: moving a sheet after/before itself is a no-op.
            // Removing first detaches sheetEl, then InsertAfterSelf/InsertBeforeSelf
            // throws "parent is null" and the sheet is lost (data loss).
            if (ReferenceEquals(afterAnchor, sheetEl) || ReferenceEquals(beforeAnchor, sheetEl))
                return $"/{sheetName}";

            // localSheetId on <definedName> is a 0-based position into
            // <sheets>; capture the pre-move order so scoped names can be
            // remapped to the sheets' new positions after the reorder.
            var preMoveOrder = sheets.Elements<Sheet>().ToList();

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

            // Remap sheet-scoped defined names (printArea, print titles,
            // scoped named ranges) from old positions to new ones — an
            // unremapped localSheetId silently rebinds the name to whatever
            // sheet now occupies the old position.
            var postMoveOrder = sheets.Elements<Sheet>().ToList();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                foreach (var dn in definedNames.Elements<DefinedName>())
                {
                    var lid = dn.LocalSheetId?.Value;
                    if (lid == null || lid >= preMoveOrder.Count) continue;
                    var newIdx = postMoveOrder.IndexOf(preMoveOrder[(int)lid.Value]);
                    if (newIdx >= 0 && newIdx != lid.Value) dn.LocalSheetId = (uint)newIdx;
                }
            }

            // Mark the document modified so Dispose flushes it. Without this,
            // a `using (h) h.Move(...)` (no explicit Save) is silently dropped
            // by the !Modified byte-preserving discard path in Dispose.
            Modified = true;
            workbook.Save();
            return $"/{sheetName}";
        }

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Determine the target sheet's SheetData. The result path is built from
        // the resolved target SHEET (below), NOT the raw --to: a row/col/cell
        // lives directly under a sheet, so a non-sheet --to like /Sheet1/row[2]
        // must not leak into the result path (it used to produce a doubled
        // /Sheet1/row[2]/row[3]). Only the sheet segment of --to is meaningful.
        SheetData targetSheetData;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            targetSheetData = sheetData;
        }
        else
        {
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
            var row = (rowIdx >= 1 && rowIdx <= allRows.Count ? allRows[PathIndex.ToArrayIndex(rowIdx)] : null)
                ?? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");

            // Resolve --before / --after anchors to a 0-based document-order
            // position in the target sheet. Anchor must be /<TargetSheet>/row[K].
            // Resolved BEFORE removing the moved row so the anchor is found by
            // its current position.
            int? targetIndex = index;
            string targetSheetName = string.IsNullOrEmpty(targetParentPath)
                ? sheetName
                : targetParentPath.TrimStart('/').Split('/', 2)[0];
            if (targetIndex == null && position != null && (position.After != null || position.Before != null))
            {
                int FindAnchorRowPos(string anchorPath)
                {
                    var aSegs = anchorPath.TrimStart('/').Split('/', 2);
                    if (aSegs.Length < 2)
                        throw new ArgumentException(
                            $"Anchor must be a row path like /{targetSheetName}/row[K], got: {anchorPath}");
                    if (!aSegs[0].Equals(targetSheetName, StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException(
                            $"Anchor sheet '{aSegs[0]}' must match target sheet '{targetSheetName}'");
                    var am = Regex.Match(aSegs[1], @"^row\[(\d+)\]$");
                    if (!am.Success)
                        throw new ArgumentException(
                            $"Anchor must be a row path like /{targetSheetName}/row[K], got: {anchorPath}");
                    var anchorRowIdx = uint.Parse(am.Groups[1].Value);
                    var pos = targetSheetData.Elements<Row>().ToList()
                        .FindIndex(r => r.RowIndex?.Value == anchorRowIdx);
                    if (pos < 0)
                        throw new ArgumentException($"Anchor row {anchorRowIdx} not found in {targetSheetName}");
                    return pos;
                }
                if (position.Before != null) targetIndex = FindAnchorRowPos(position.Before);
                else targetIndex = FindAnchorRowPos(position.After!) + 1;
            }

            // If the moved row sits before the anchor in the same sheet,
            // removing it shifts everything (including the anchor) up by one.
            // Adjust the resolved target index so it still points at the
            // intended slot in post-remove document order.
            if (targetIndex.HasValue && targetSheetData == sheetData)
            {
                var srcPos = sheetData.Elements<Row>().ToList().IndexOf(row);
                if (srcPos >= 0 && srcPos < targetIndex.Value)
                    targetIndex = targetIndex.Value - 1;
            }

            // Snapshot every row's old RowIndex (per sheet) so we can build
            // an oldToNew renumber map after the reposition + renumber. The
            // map drives formula and range-ref rewriting so cross-row
            // references follow the moved content.
            var srcOldIdx = sheetData.Elements<Row>().ToDictionary(r => r, r => (int)(r.RowIndex?.Value ?? 0));
            Dictionary<Row, int>? tgtOldIdx = null;
            if (targetSheetData != sheetData)
                tgtOldIdx = targetSheetData.Elements<Row>().ToDictionary(r => r, r => (int)(r.RowIndex?.Value ?? 0));

            // Mark modified before the first irreversible mutation so Dispose
            // flushes (the !Modified path would otherwise discard the move).
            Modified = true;
            row.Remove();

            if (targetIndex.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (targetIndex.Value >= 0 && targetIndex.Value < rows.Count)
                    rows[targetIndex.Value].InsertBeforeSelf(row);
                else
                    targetSheetData.AppendChild(row);
            }
            else
            {
                targetSheetData.AppendChild(row);
            }

            // Renumber every row in document order so Excel reads them in the
            // intended sequence — Excel ignores XML document order and uses
            // <row r='N'> as the source of truth. Without renumbering, a move
            // operation appears to do nothing on reopen.
            //
            // Limitation: this collapses any gaps the original sheet may have
            // had (e.g. rows 1, 3, 5 → rows 1, 2, 3). Sheets with intentional
            // RowIndex gaps are unusual; if the user needs gap preservation,
            // they should perform the move via direct cell-level set ops.
            RenumberRowsAndCellRefs(targetSheetData);
            if (targetSheetData != sheetData)
                RenumberRowsAndCellRefs(sheetData);

            // Build oldToNew row-index maps and apply to formula text +
            // range-bearing structures (mergeCells, CF/DV sqref, autoFilter,
            // hyperlinks, table refs). Without this, formulas like A1==A3
            // would still read literal A3 after the move, defeating the
            // 'follow content' contract.
            var srcMap = BuildRowRenumberMap(srcOldIdx);
            var srcSheetWs = worksheet;
            ApplyRowRenumberToSheet(srcSheetWs, sheetName, srcMap);
            if (targetSheetData != sheetData && tgtOldIdx != null)
            {
                var tgtMap = BuildRowRenumberMap(tgtOldIdx);
                var tgtWsPart = GetWorksheets().FirstOrDefault(w => GetSheet(w.Part).GetFirstChild<SheetData>() == targetSheetData).Part;
                if (tgtWsPart != null)
                    ApplyRowRenumberToSheet(tgtWsPart, GetWorksheets().First(w => w.Part == tgtWsPart).Name, tgtMap);
            }

            SaveWorksheet(worksheet);
            if (targetSheetData != sheetData)
            {
                var tgtWs = GetWorksheets().FirstOrDefault(w => GetSheet(w.Part).GetFirstChild<SheetData>() == targetSheetData).Part;
                if (tgtWs != null) SaveWorksheet(tgtWs);
            }
            var newRowIndex = row.RowIndex?.Value ?? 0u;
            return $"/{targetSheetName}/row[{newRowIndex}]";
        }

        // Move col[L]: shuffle cells across the affected column band, renumber
        // <col> metadata, and remap formulas + range refs via FormulaRefShifter
        // ApplyColRenumberMap. Same scope rules as row move (single sheet,
        // anchor must be col[L] in same sheet).
        var colMatch = Regex.Match(elementRef, @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var srcColLetter = colMatch.Groups[1].Value.ToUpperInvariant();
            var srcColIdx = ColumnNameToIndex(srcColLetter);

            // Resolve target. Default behavior (no position): append after the
            // last used column.
            int? targetColIdx = null;
            if (position?.Index.HasValue == true)
                targetColIdx = position.Index!.Value;
            else if (position?.Before != null || position?.After != null)
            {
                int FindAnchorColIdx(string anchorPath)
                {
                    var aSegs = anchorPath.TrimStart('/').Split('/', 2);
                    if (aSegs.Length < 2)
                        throw new ArgumentException(
                            $"Anchor must be a col path like /{sheetName}/col[L], got: {anchorPath}");
                    if (!aSegs[0].Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException(
                            $"Anchor sheet '{aSegs[0]}' must match source sheet '{sheetName}'");
                    var am = Regex.Match(aSegs[1], @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
                    if (!am.Success)
                        throw new ArgumentException(
                            $"Anchor must be a col path like /{sheetName}/col[L], got: {anchorPath}");
                    return ColumnNameToIndex(am.Groups[1].Value.ToUpperInvariant());
                }
                if (position.Before != null) targetColIdx = FindAnchorColIdx(position.Before);
                else targetColIdx = FindAnchorColIdx(position.After!) + 1;
            }
            else
            {
                // Append after last used column.
                int maxCol = 1;
                foreach (var r in sheetData.Elements<Row>())
                    foreach (var c in r.Elements<Cell>())
                        if (c.CellReference?.Value != null)
                            maxCol = Math.Max(maxCol, ColumnNameToIndex(ParseCellReference(c.CellReference.Value).Column));
                targetColIdx = maxCol + 1;
            }

            int target = targetColIdx!.Value;
            if (target == srcColIdx || target == srcColIdx + 1)
            {
                // No-op: moving a col to its own slot or right after itself.
                return $"/{sheetName}/col[{srcColLetter}]";
            }

            // Mark modified before mutating so Dispose flushes (past the no-op
            // early return above; the !Modified path would else discard the move).
            Modified = true;

            // Build the col renumber map. Two cases:
            //   src < target: cols (src+1)..(target-1) shift left by 1; src moves to (target-1).
            //   src > target: cols target..(src-1) shift right by 1; src moves to target.
            var colMap = new Dictionary<int, int>();
            if (srcColIdx < target)
            {
                for (int i = srcColIdx + 1; i < target; i++) colMap[i] = i - 1;
                colMap[srcColIdx] = target - 1;
            }
            else
            {
                for (int i = target; i < srcColIdx; i++) colMap[i] = i + 1;
                colMap[srcColIdx] = target;
            }

            // Apply map to cell references in sheetData.
            foreach (var r in sheetData.Elements<Row>())
            {
                foreach (var c in r.Elements<Cell>())
                {
                    if (c.CellReference?.Value == null) continue;
                    var (col, row) = ParseCellReference(c.CellReference.Value);
                    var oldIdx = ColumnNameToIndex(col);
                    if (colMap.TryGetValue(oldIdx, out var newIdx))
                        c.CellReference = $"{IndexToColumnName(newIdx)}{row}";
                }
                // After remap, cells in a row may be out of left-to-right order;
                // OOXML expects ascending CellReference within a row.
                var sortedCells = r.Elements<Cell>()
                    .OrderBy(c => c.CellReference?.Value == null ? 0 : ColumnNameToIndex(ParseCellReference(c.CellReference.Value).Column))
                    .ToList();
                r.RemoveAllChildren<Cell>();
                foreach (var c in sortedCells) r.AppendChild(c);
            }

            // Apply map to <col> metadata (width/style entries).
            var ws = GetSheet(worksheet);
            var columns = ws.GetFirstChild<Columns>();
            if (columns != null)
            {
                foreach (var colEl in columns.Elements<Column>().ToList())
                {
                    var minOld = (int)(colEl.Min?.Value ?? 0);
                    var maxOld = (int)(colEl.Max?.Value ?? 0);
                    // Only handle the simple case of single-column entries
                    // (min == max). Multi-col runs spanning the moved band are
                    // left as-is — user-meaningful collisions are rare and
                    // post-renumber a multi-col run can't always be expressed
                    // as a single Column element either.
                    if (minOld == maxOld && colMap.TryGetValue(minOld, out var newIdx))
                    {
                        colEl.Min = (uint)newIdx;
                        colEl.Max = (uint)newIdx;
                    }
                }
                // Sort col entries ascending for OOXML schema validity.
                var sortedCols = columns.Elements<Column>()
                    .OrderBy(c => c.Min?.Value ?? 0).ToList();
                columns.RemoveAllChildren<Column>();
                foreach (var c in sortedCols) columns.AppendChild(c);
            }

            // Remap formulas + range-bearing structures via the col shifter.
            ApplyColRenumberToSheet(worksheet, sheetName, colMap);

            DeleteCalcChainIfPresent();
            SaveWorksheet(worksheet);
            int newSrcIdx = colMap[srcColIdx];
            return $"/{sheetName}/col[{IndexToColumnName(newSrcIdx)}]";
        }

        throw new ArgumentException($"Move not supported for: {elementRef}. Supported: row[N], col[L]");
    }

    /// <summary>
    /// Build {old → new} row-index map from a snapshot taken before the
    /// move + renumber. Rows whose old and new index match are omitted (the
    /// shifter treats absent keys as no-op).
    /// </summary>
    private static Dictionary<int, int> BuildRowRenumberMap(Dictionary<Row, int> oldIdxByRow)
    {
        var map = new Dictionary<int, int>(oldIdxByRow.Count);
        foreach (var (row, oldIdx) in oldIdxByRow)
        {
            int newIdx = (int)(row.RowIndex?.Value ?? 0u);
            if (newIdx != 0 && newIdx != oldIdx)
                map[oldIdx] = newIdx;
        }
        return map;
    }

    /// <summary>
    /// Apply an oldToNew row-index map to every formula and range-bearing
    /// structure on the sheet (mergeCells, CF/DV sqref, autoFilter,
    /// hyperlinks, table refs). Range refs whose endpoints invert after
    /// renumber are left unchanged (best-effort: post-renumber they no
    /// longer express a contiguous A1 region).
    /// </summary>
    private void ApplyRowRenumberToSheet(WorksheetPart worksheet, string sheetName, IReadOnlyDictionary<int, int> map)
    {
        if (map.Count == 0) return;
        // Drawing markers use 0-based row indices; the renumber map is keyed by
        // 1-based row indices. Translate so anchored pictures/charts follow the
        // rows they sit on (parity with insert/delete's rowMarkerShift).
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => RemapRowsInRangeRef(r, map),
            formulaTextMapper: f => Core.FormulaRefShifter.ApplyRowRenumberMap(f, sheetName, sheetName, map),
            rowMarkerShift: m => map.TryGetValue(m + 1, out var n) ? n - 1 : m,
            crossSheetFormulaMapper: (other, f) => Core.FormulaRefShifter.ApplyRowRenumberMap(f, other, sheetName, map));
    }

    private void ApplyColRenumberToSheet(WorksheetPart worksheet, string sheetName, IReadOnlyDictionary<int, int> map)
    {
        if (map.Count == 0) return;
        // Drawing markers use 0-based column indices; the renumber map is keyed
        // by 1-based column indices. Translate so anchored objects follow the
        // columns they sit on (parity with insert/delete's colMarkerShift).
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => RemapColsInRangeRef(r, map),
            formulaTextMapper: f => Core.FormulaRefShifter.ApplyColRenumberMap(f, sheetName, sheetName, map),
            colMarkerShift: m => map.TryGetValue(m + 1, out var n) ? n - 1 : m,
            crossSheetFormulaMapper: (other, f) => Core.FormulaRefShifter.ApplyColRenumberMap(f, other, sheetName, map));
    }

    private static string? RemapColsInRangeRef(string? refStr, IReadOnlyDictionary<int, int> map)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        var colVals = new List<int>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var match = System.Text.RegularExpressions.Regex.Match(part, @"^([A-Z]+)(\d+)$");
                if (!match.Success) { shifted.Add(part); colVals.Add(-1); continue; }
                var col = match.Groups[1].Value;
                var oldColIdx = ColumnNameToIndex(col);
                var row = match.Groups[2].Value;
                var newCol = map.TryGetValue(oldColIdx, out var n) ? IndexToColumnName(n) : col;
                shifted.Add($"{newCol}{row}");
                colVals.Add(map.TryGetValue(oldColIdx, out var ni) ? ni : oldColIdx);
            }
            catch { shifted.Add(part); colVals.Add(-1); }
        }
        if (colVals.Count == 2 && colVals[0] > 0 && colVals[1] > 0 && colVals[0] > colVals[1])
            return null;
        return string.Join(":", shifted);
    }

    // ApplyRowRenumberToWorkbookDefinedNames / ApplyColRenumberToWorkbookDefinedNames
    // removed — defined-names are now rewritten by section 8 of
    // ApplySheetRangeMutations (the formulaTextMapper passed in).

    /// <summary>
    /// Apply the row-renumber map to a range-style ref like 'B2:D5' or 'A1'.
    /// Returns null if any endpoint's row is absent from the map AND the
    /// other endpoint is in the map (would produce a malformed range), or
    /// if the resulting endpoints invert.
    /// </summary>
    private static string? RemapRowsInRangeRef(string? refStr, IReadOnlyDictionary<int, int> map)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        var rowVals = new List<int>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var match = System.Text.RegularExpressions.Regex.Match(part, @"^([A-Z]+)(\d+)$");
                if (!match.Success) { shifted.Add(part); rowVals.Add(-1); continue; }
                var col = match.Groups[1].Value;
                var oldRow = int.Parse(match.Groups[2].Value);
                var newRow = map.TryGetValue(oldRow, out var n) ? n : oldRow;
                shifted.Add($"{col}{newRow}");
                rowVals.Add(newRow);
            }
            catch { shifted.Add(part); rowVals.Add(-1); }
        }
        // Range endpoint sanity: if both rows are valid and start > end, abort.
        if (rowVals.Count == 2 && rowVals[0] > 0 && rowVals[1] > 0 && rowVals[0] > rowVals[1])
            return null;
        return string.Join(":", shifted);
    }

    /// <summary>
    /// Walk every Row in document order and reassign RowIndex to its 1-based
    /// position, then rewrite every cell's CellReference to match the new
    /// row number. Used after Move to make Excel honor the document-order
    /// rearrangement.
    /// </summary>
    private void RenumberRowsAndCellRefs(SheetData sheetData)
    {
        InvalidateRowIndex(sheetData);
        uint newIdx = 1;
        foreach (var row in sheetData.Elements<Row>())
        {
            row.RowIndex = newIdx;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value == null) continue;
                var (col, _) = ParseCellReference(cell.CellReference.Value);
                cell.CellReference = $"{col}{newIdx}";
            }
            newIdx++;
        }
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

        var rowIndex1 = row1.RowIndex?.Value ?? (uint)idx1;
        var rowIndex2 = row2.RowIndex?.Value ?? (uint)idx2;

        // Snapshot every row's old RowIndex before the swap so we can build the
        // oldToNew renumber map afterwards. This routes the swap through the
        // same machinery as Move (ApplyRowRenumberToSheet): external formulas,
        // CF/DV sqref, mergeCells, chart series refs and drawing anchors all
        // follow the swapped content, and formula caches are refreshed at flush
        // so cachedValue stays consistent with the formula.
        var oldIdx = sheetData.Elements<Row>().ToDictionary(r => r, r => (int)(r.RowIndex?.Value ?? 0));

        // Mark modified before the swap so Dispose flushes it (the !Modified
        // byte-preserving discard path would otherwise drop the whole swap).
        Modified = true;

        // Physically exchange the two rows in document order, then renumber by
        // document order — mirrors Move's reposition + RenumberRowsAndCellRefs.
        PowerPointHandler.SwapXmlElements(row1, row2);
        RenumberRowsAndCellRefs(sheetData);

        var map = BuildRowRenumberMap(oldIdx);
        ApplyRowRenumberToSheet(worksheet, sheetName, map);

        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);

        return ($"/{sheetName}/row[{rowIndex2}]", $"/{sheetName}/row[{rowIndex1}]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
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

            // Resolve --after/--before anchors to a 0-based row position in
            // the target sheet. Anchor format must be `/SheetName/row[K]`.
            // Mismatch (different sheet, non-row anchor, missing row) → throw.
            int? index = null;
            if (position != null)
            {
                var rowsList = targetSheetData.Elements<Row>().ToList();
                int FindAnchorRowIndex(string anchorPath)
                {
                    var aSegs = anchorPath.TrimStart('/').Split('/', 2);
                    if (aSegs.Length < 2)
                        throw new ArgumentException(
                            $"Anchor must be a row path like /{tgtSegments[0]}/row[K], got: {anchorPath}");
                    if (!aSegs[0].Equals(tgtSegments[0], StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException(
                            $"Anchor sheet '{aSegs[0]}' must match target sheet '{tgtSegments[0]}'");
                    var am = Regex.Match(aSegs[1], @"^row\[(\d+)\]$");
                    if (!am.Success)
                        throw new ArgumentException(
                            $"Anchor must be a row path like /{tgtSegments[0]}/row[K], got: {anchorPath}");
                    var anchorRowIdx = uint.Parse(am.Groups[1].Value);
                    var pos = rowsList.FindIndex(r => r.RowIndex?.Value == anchorRowIdx);
                    if (pos < 0)
                        throw new ArgumentException($"Anchor row {anchorRowIdx} not found in {tgtSegments[0]}");
                    return pos;
                }
                index = position.Resolve(FindAnchorRowIndex, rowsList.Count);
            }

            // R8-1: CloneNode preserves the source row's RowIndex and every
            // cell's CellReference (e.g. "A1","B1"). Without rewriting these,
            // the new row collides with the source (Excel shows one row at
            // rowIdx, A2 appears empty) or is silently ignored. Compute the
            // new rowIndex from the target sheet and rewrite all cell refs.
            uint newRowIndex;
            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                {
                    newRowIndex = rows[index.Value].RowIndex?.Value ?? (uint)(index.Value + 1);
                    // Shift existing rows at/after this position down by 1
                    ShiftRowsDown(tgtWorksheet, (int)newRowIndex);
                    // Re-fetch sheetData (ShiftRowsDown may reorder)
                    targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()!;
                    var afterRow = targetSheetData.Elements<Row>()
                        .LastOrDefault(r => (r.RowIndex?.Value ?? 0) < newRowIndex);
                    if (afterRow != null) afterRow.InsertAfterSelf(clone);
                    else targetSheetData.InsertAt(clone, 0);
                }
                else
                {
                    newRowIndex = (targetSheetData.Elements<Row>()
                        .LastOrDefault()?.RowIndex?.Value ?? 0u) + 1;
                    targetSheetData.AppendChild(clone);
                }
            }
            else
            {
                newRowIndex = (targetSheetData.Elements<Row>()
                    .LastOrDefault()?.RowIndex?.Value ?? 0u) + 1;
                targetSheetData.AppendChild(clone);
            }

            clone.RowIndex = newRowIndex;
            int copyDeltaRow = (int)newRowIndex - (int)rowIdx;
            string targetSheetName = tgtSegments[0];
            foreach (var c in clone.Elements<Cell>())
            {
                var oldRef = c.CellReference?.Value;
                if (string.IsNullOrEmpty(oldRef)) continue;
                var m = Regex.Match(oldRef, @"^([A-Z]+)\d+$", RegexOptions.IgnoreCase);
                if (m.Success)
                    c.CellReference = $"{m.Groups[1].Value.ToUpperInvariant()}{newRowIndex}";

                // Apply copy-delta to formulas inside cloned cells so that
                // relative refs follow the new anchor row. Excel UI does this
                // automatically for "Insert Copied Cells" / paste. Refs to
                // other sheets are left untouched (sheet-scope guard).
                if (c.CellFormula != null && !string.IsNullOrEmpty(c.CellFormula.Text) && copyDeltaRow != 0)
                {
                    c.CellFormula.Text = Core.FormulaRefShifter.ApplyCopyDelta(
                        c.CellFormula.Text, targetSheetName, sheetName,
                        deltaCol: 0, deltaRow: copyDeltaRow);
                }
            }

            // mergeCells live in the sheet-level <mergeCells> container, not
            // inside the row's subtree, so CloneNode misses them. Walk the
            // SOURCE sheet's mergeCells for entries whose start AND end rows
            // both equal the source row index (single-row merges within the
            // copied row), and add a corresponding mergeCell at the new row
            // index. Multi-row merges that include the source row are out of
            // scope for row-copy semantics — they belong to a region, not a
            // single row.
            var srcMergeCells = GetSheet(worksheet).GetFirstChild<MergeCells>();
            if (srcMergeCells != null)
            {
                var newMergesToAdd = new List<string>();
                foreach (var mc in srcMergeCells.Elements<MergeCell>())
                {
                    var refStr = mc.Reference?.Value;
                    if (string.IsNullOrEmpty(refStr)) continue;
                    var parts = refStr.Split(':');
                    if (parts.Length != 2) continue;
                    var ms = Regex.Match(parts[0], @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
                    var me = Regex.Match(parts[1], @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
                    if (!ms.Success || !me.Success) continue;
                    if (uint.Parse(ms.Groups[2].Value) == rowIdx
                        && uint.Parse(me.Groups[2].Value) == rowIdx)
                    {
                        newMergesToAdd.Add(
                            $"{ms.Groups[1].Value.ToUpperInvariant()}{newRowIndex}:" +
                            $"{me.Groups[1].Value.ToUpperInvariant()}{newRowIndex}");
                    }
                }
                if (newMergesToAdd.Count > 0)
                {
                    var tgtSheetEl = GetSheet(tgtWorksheet);
                    var tgtMergeCells = tgtSheetEl.GetFirstChild<MergeCells>()
                        ?? tgtSheetEl.AppendChild(new MergeCells());
                    foreach (var newRef in newMergesToAdd)
                        tgtMergeCells.AppendChild(new MergeCell { Reference = newRef });
                    tgtMergeCells.Count = (uint)tgtMergeCells.Elements<MergeCell>().Count();
                }
            }

            // Mark modified so Dispose flushes the copy (the !Modified
            // byte-preserving discard path would otherwise drop it).
            Modified = true;
            SaveWorksheet(tgtWorksheet);
            return $"{targetParentPath}/row[{newRowIndex}]";
        }

        // Copy col[L] — mirror of the row case. Snapshot cells from the
        // source column before any shift; resolve target col from anchor or
        // index; ShiftColumnsRight at the target col (handles all displacement
        // for cellRef + col metadata + mergeCells + CF/DV/autoFilter +
        // hyperlinks + tables + namedRanges + cross-sheet formula refs); then
        // insert the snapshotted cells at the target col with delta-shifted
        // formulas. Single-col merges fully contained in the source column
        // are replicated at the target column.
        var colMatch = Regex.Match(elementRef, @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var srcColLetter = colMatch.Groups[1].Value.ToUpperInvariant();
            var srcColIdx = ColumnNameToIndex(srcColLetter);

            // Resolve target col index. With no position → append after
            // the last used column.
            int targetColIdx;
            if (position?.Index.HasValue == true)
            {
                targetColIdx = position.Index.Value > 0 ? position.Index.Value : 1;
            }
            else if (position?.Before != null || position?.After != null)
            {
                int FindAnchorColIdx(string anchorPath)
                {
                    var aSegs = anchorPath.TrimStart('/').Split('/', 2);
                    if (aSegs.Length < 2)
                        throw new ArgumentException(
                            $"Anchor must be a col path like /{tgtSegments[0]}/col[L], got: {anchorPath}");
                    if (!aSegs[0].Equals(tgtSegments[0], StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException(
                            $"Anchor sheet '{aSegs[0]}' must match target sheet '{tgtSegments[0]}'");
                    var am = Regex.Match(aSegs[1], @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
                    if (!am.Success)
                        throw new ArgumentException(
                            $"Anchor must be a col path like /{tgtSegments[0]}/col[L], got: {anchorPath}");
                    return ColumnNameToIndex(am.Groups[1].Value.ToUpperInvariant());
                }
                if (position.Before != null) targetColIdx = FindAnchorColIdx(position.Before);
                else targetColIdx = FindAnchorColIdx(position.After!) + 1;
            }
            else
            {
                int maxCol = 0;
                foreach (var r in sheetData.Elements<Row>())
                    foreach (var c in r.Elements<Cell>())
                        if (c.CellReference?.Value != null)
                            maxCol = Math.Max(maxCol, ColumnNameToIndex(ParseCellReference(c.CellReference.Value).Column));
                targetColIdx = maxCol + 1;
            }

            // Snapshot source col cells (clones) BEFORE any shift, keyed by
            // row number so we can recreate them at the target col.
            var srcCellClones = new List<(uint Row, Cell Clone)>();
            foreach (var r in sheetData.Elements<Row>())
            {
                var cell = r.Elements<Cell>().FirstOrDefault(c =>
                {
                    if (c.CellReference?.Value == null) return false;
                    return ParseCellReference(c.CellReference.Value).Column
                        .Equals(srcColLetter, StringComparison.OrdinalIgnoreCase);
                });
                if (cell != null && r.RowIndex?.Value != null)
                    srcCellClones.Add((r.RowIndex.Value, (Cell)cell.CloneNode(true)));
            }

            // Snapshot single-col merges fully contained in the source col.
            var srcSingleColMerges = new List<(uint StartRow, uint EndRow)>();
            var srcMergeCells = GetSheet(worksheet).GetFirstChild<MergeCells>();
            if (srcMergeCells != null)
            {
                foreach (var mc in srcMergeCells.Elements<MergeCell>())
                {
                    var refStr = mc.Reference?.Value;
                    if (string.IsNullOrEmpty(refStr)) continue;
                    var parts = refStr.Split(':');
                    if (parts.Length != 2) continue;
                    var (sCol, sRow) = ParseCellReference(parts[0]);
                    var (eCol, eRow) = ParseCellReference(parts[1]);
                    if (sCol.Equals(srcColLetter, StringComparison.OrdinalIgnoreCase)
                        && eCol.Equals(srcColLetter, StringComparison.OrdinalIgnoreCase))
                    {
                        srcSingleColMerges.Add(((uint)sRow, (uint)eRow));
                    }
                }
            }

            // Make room at target col. ShiftColumnsRight handles all
            // sheet-wide displacement (cellRef, col meta, mergeCells, CF/DV,
            // autoFilter, hyperlinks, tables, namedRanges, formulas).
            ShiftColumnsRight(tgtWorksheet, targetColIdx);

            // Account for the source col having been shifted right by 1 if
            // it was at or after the target.
            int effectiveSrcColIdx = srcColIdx >= targetColIdx ? srcColIdx + 1 : srcColIdx;
            int copyDeltaCol = targetColIdx - effectiveSrcColIdx;

            // Insert snapshotted cell clones into the target col.
            var tgtSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()!;
            string targetColLetter = IndexToColumnName(targetColIdx);
            foreach (var (srcRowNum, clone) in srcCellClones)
            {
                clone.CellReference = $"{targetColLetter}{srcRowNum}";

                // Delta-shift formulas inside the clone: relative refs follow
                // the new anchor column.
                if (clone.CellFormula != null && !string.IsNullOrEmpty(clone.CellFormula.Text) && copyDeltaCol != 0)
                {
                    clone.CellFormula.Text = Core.FormulaRefShifter.ApplyCopyDelta(
                        clone.CellFormula.Text, tgtSegments[0], sheetName,
                        deltaCol: copyDeltaCol, deltaRow: 0);
                }

                var targetRow = tgtSheetData.Elements<Row>()
                    .FirstOrDefault(r => r.RowIndex?.Value == srcRowNum);
                if (targetRow == null)
                {
                    // Materialize the row in correct ascending order.
                    targetRow = new Row { RowIndex = srcRowNum };
                    var afterRow = tgtSheetData.Elements<Row>()
                        .LastOrDefault(r => (r.RowIndex?.Value ?? 0) < srcRowNum);
                    if (afterRow != null) afterRow.InsertAfterSelf(targetRow);
                    else tgtSheetData.InsertAt(targetRow, 0);
                }
                // Insert clone at the correct in-row position (ascending col).
                var afterCell = targetRow.Elements<Cell>()
                    .LastOrDefault(c => c.CellReference?.Value != null
                        && ColumnNameToIndex(ParseCellReference(c.CellReference.Value).Column) < targetColIdx);
                if (afterCell != null) afterCell.InsertAfterSelf(clone);
                else targetRow.InsertAt(clone, 0);
            }

            // Replicate single-col merges at the target col.
            if (srcSingleColMerges.Count > 0)
            {
                var tgtSheetEl = GetSheet(tgtWorksheet);
                var tgtMergeCells = tgtSheetEl.GetFirstChild<MergeCells>()
                    ?? tgtSheetEl.AppendChild(new MergeCells());
                foreach (var (sRow, eRow) in srcSingleColMerges)
                    tgtMergeCells.AppendChild(new MergeCell {
                        Reference = $"{targetColLetter}{sRow}:{targetColLetter}{eRow}"
                    });
                tgtMergeCells.Count = (uint)tgtMergeCells.Elements<MergeCell>().Count();
            }

            // Mark modified so Dispose flushes the copy (the !Modified
            // byte-preserving discard path would otherwise drop it).
            Modified = true;
            DeleteCalcChainIfPresent();
            SaveWorksheet(tgtWorksheet);
            return $"{targetParentPath}/col[{targetColLetter}]";
        }

        throw new ArgumentException($"Copy not supported for: {elementRef}. Supported: row[N], col[L]");
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

            case "drawing-group":
            {
                // Verbatim DrawingML group carrier for xlsx dump→batch.
                // The full hosting anchor is preserved because flattening a
                // <xdr:grpSp> loses the child coordinate system, z-order,
                // styles and the fact that the objects are grouped. Only
                // hyperlink relationships are carried here; dump falls back
                // to semantic leaf shapes when a group references package
                // parts such as images/charts.
                var groupSheetName = parentPartPath.TrimStart('/');
                var groupWorksheet = FindWorksheet(groupSheetName)
                    ?? throw new ArgumentException(
                        $"Sheet not found: {groupSheetName}. drawing-group must be added under a sheet.");
                properties ??= new Dictionary<string, string>();
                var anchorXml = properties.GetValueOrDefault("anchor-xml")
                    ?? throw new ArgumentException(
                        "'anchor-xml' property is required for drawing-group (verbatim xdr anchor XML)");

                XDR.TwoCellAnchor groupAnchor;
                try
                {
                    groupAnchor = new XDR.TwoCellAnchor(anchorXml);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException(
                        $"drawing-group anchor XML is not a valid xdr:twoCellAnchor: {ex.Message}", ex);
                }
                if (groupAnchor.GetFirstChild<XDR.GroupShape>() == null)
                    throw new ArgumentException(
                        "drawing-group anchor XML must contain a top-level xdr:grpSp.");

                List<DumpDrawingHyperlinkSpec> groupHyperlinks;
                try
                {
                    groupHyperlinks = DecodeDumpDrawingHyperlinks(
                        properties.GetValueOrDefault("hyperlinks") ?? "");
                }
                catch (FormatException ex)
                {
                    throw new ArgumentException(
                        $"drawing-group 'hyperlinks' carrier is invalid: {ex.Message}", ex);
                }

                Modified = true;
                var groupDrawingsPart = groupWorksheet.DrawingsPart
                    ?? groupWorksheet.AddNewPart<DrawingsPart>();
                if (groupDrawingsPart.WorksheetDrawing == null)
                {
                    groupDrawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
                    groupDrawingsPart.WorksheetDrawing.Save();
                }
                var groupSheet = GetSheet(groupWorksheet);
                if (groupSheet.GetFirstChild<SpreadsheetDrawing>() == null)
                {
                    var drawingRelId = groupWorksheet.GetIdOfPart(groupDrawingsPart);
                    groupSheet.Append(new SpreadsheetDrawing { Id = drawingRelId });
                    SaveWorksheet(groupWorksheet);
                }

                // Relationship IDs are scoped to the destination drawing part.
                // Create fresh IDs (avoids collisions with pictures/charts
                // emitted earlier), then rewrite every r:id/r:embed/r:link in
                // the verbatim group anchor that referenced the source ID.
                foreach (var hyperlink in groupHyperlinks)
                {
                    if (string.IsNullOrEmpty(hyperlink.Id) || string.IsNullOrEmpty(hyperlink.Target))
                        throw new ArgumentException(
                            "drawing-group hyperlink entries require non-empty Id and Target.");
                    var uri = new Uri(hyperlink.Target, UriKind.RelativeOrAbsolute);
                    var replayRel = groupDrawingsPart.AddHyperlinkRelationship(
                        uri, hyperlink.IsExternal);
                    RemapDrawingRelationshipId(groupAnchor, hyperlink.Id, replayRel.Id);
                }

                groupDrawingsPart.WorksheetDrawing.AppendChild(groupAnchor);
                groupDrawingsPart.WorksheetDrawing.Save();
                var groupIndex = groupDrawingsPart.WorksheetDrawing
                    .Elements<XDR.TwoCellAnchor>()
                    .Count(a => a.GetFirstChild<XDR.GroupShape>() != null);
                return ("group", $"/{groupSheetName}/group[{groupIndex}]");
            }

            case "chartex":
            {
                // Extended (cx:) chart carrier for dump→batch round-trip.
                // chartEx has no semantic add vocabulary — waterfall/funnel/
                // sunburst charts are carried VERBATIM: the caller pins the
                // source rIds so the graphicFrame slice raw-set into the
                // drawing resolves without rewriting. Mirrors the pptx
                // SmartArt add-part pattern (pinned rIds + raw payload).
                // Props: rid (required), xml (base64 cx:chartSpace),
                // colors-rid/colors-xml, style-rid/style-xml (optional
                // sub-parts — Excel-authored chartEx always carries both;
                // dropping them dangles the main part's rels).
                var cxSheetName = parentPartPath.TrimStart('/');
                var cxWorksheet = FindWorksheet(cxSheetName)
                    ?? throw new ArgumentException(
                        $"Sheet not found: {cxSheetName}. chartex must be added under a sheet: add-part <file> /<SheetName> --type chartex");
                properties ??= new Dictionary<string, string>();
                var cxRid = properties.GetValueOrDefault("rid")
                    ?? throw new ArgumentException("'rid' property is required for chartex (pinned relationship id)");
                var cxXmlB64 = properties.GetValueOrDefault("xml")
                    ?? throw new ArgumentException("'xml' property is required for chartex (base64 cx:chartSpace XML)");

                var cxDrawingsPart = cxWorksheet.DrawingsPart
                    ?? cxWorksheet.AddNewPart<DrawingsPart>();
                if (cxDrawingsPart.WorksheetDrawing == null)
                {
                    cxDrawingsPart.WorksheetDrawing =
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                    cxDrawingsPart.WorksheetDrawing.Save();
                    if (GetSheet(cxWorksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var cxDrawRelId = cxWorksheet.GetIdOfPart(cxDrawingsPart);
                        GetSheet(cxWorksheet).Append(
                            new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = cxDrawRelId });
                        SaveWorksheet(cxWorksheet);
                    }
                }

                var extChartPart = cxDrawingsPart.AddNewPart<ExtendedChartPart>(cxRid);
                using (var s = extChartPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write))
                {
                    var bytes = Convert.FromBase64String(cxXmlB64);
                    s.Write(bytes, 0, bytes.Length);
                }
                foreach (var (ridKey, xmlKey, subType) in new[]
                {
                    ("colors-rid", "colors-xml", "colors"),
                    ("style-rid", "style-xml", "style"),
                })
                {
                    var subRid = properties.GetValueOrDefault(ridKey);
                    var subXml = properties.GetValueOrDefault(xmlKey);
                    if (string.IsNullOrEmpty(subRid) || string.IsNullOrEmpty(subXml)) continue;
                    OpenXmlPart subPart = subType == "colors"
                        ? extChartPart.AddNewPart<ChartColorStylePart>(subRid)
                        : extChartPart.AddNewPart<ChartStylePart>(subRid);
                    using var ss = subPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write);
                    var subBytes = Convert.FromBase64String(subXml!);
                    ss.Write(subBytes, 0, subBytes.Length);
                }
                return (cxRid, $"/{cxSheetName}/chartex");
            }

            case "ole":
            {
                // Verbatim OLE carrier for dump→batch round-trip. Mirrors the
                // pptx add-part ole contract (pinned rIds + base64 payloads)
                // but is all-in-one: Excel's OLE anatomy spans the worksheet
                // (<oleObjects> child + embed/icon rels), the VML drawing
                // (anchor shape) and <legacyDrawing>, all of which must stay
                // consistent — so the handler wires everything here instead
                // of leaving XML splicing to a companion raw-set.
                // Props: rid + data (+content-type/extension) = payload part;
                // icon-rid + icon-data (+icon-content-type) = objectPr image;
                // vml-shape = the <v:shape> anchor XML verbatim;
                // object-xml = the <oleObjects> CHILD element verbatim
                // (mc:AlternateContent or bare oleObject, pinned rIds inside).
                var oleSheetName = parentPartPath.TrimStart('/');
                var oleWs = FindWorksheet(oleSheetName)
                    ?? throw new ArgumentException(
                        $"Sheet not found: {oleSheetName}. ole must be added under a sheet: add-part <file> /<SheetName> --type ole");
                properties ??= new Dictionary<string, string>();
                var oleRid = properties.GetValueOrDefault("rid")
                    ?? throw new ArgumentException("'rid' property is required for ole (pinned payload relationship id)");
                var oleDataB64 = properties.GetValueOrDefault("data")
                    ?? throw new ArgumentException("'data' property is required for ole (base64 payload bytes)");
                var oleObjectXml = properties.GetValueOrDefault("object-xml")
                    ?? throw new ArgumentException("'object-xml' property is required for ole (verbatim oleObjects child element)");
                byte[] oleBytes;
                try { oleBytes = Convert.FromBase64String(oleDataB64); }
                catch (FormatException) { throw new ArgumentException("add-part ole: 'data' is not valid base64"); }

                var oleCt = properties.GetValueOrDefault("content-type")
                    ?? "application/vnd.openxmlformats-officedocument.oleObject";
                var oleExt = properties.GetValueOrDefault("extension") ?? ".bin";
                if (!oleExt.StartsWith('.')) oleExt = "." + oleExt;

                // Kind comes from the dump (source part type), because content
                // type alone cannot classify legacy package formats (.xls
                // carries application/vnd.ms-excel, not an OOXML CT). Fallback
                // for hand-written batches that omit ole-kind: package iff the
                // CT is a non-oleObject openxmlformats CT.
                var oleKind = properties.GetValueOrDefault("ole-kind")
                    ?? (oleCt.StartsWith(
                            "application/vnd.openxmlformats-officedocument.", StringComparison.OrdinalIgnoreCase)
                        && !oleCt.Equals(
                            "application/vnd.openxmlformats-officedocument.oleObject", StringComparison.OrdinalIgnoreCase)
                        ? "package" : "object");
                // PartTypeInfo's target extension is dot-prefixed (".docx");
                // a bare "docx" silently falls back to ".bin" part names.
                OfficeCli.Core.OleHelper.AddEmbeddedPartFromBytes(
                    oleWs, oleBytes, oleKind, oleCt, oleExt, oleRid);

                // Icon image (objectPr r:id target).
                var iconRid = properties.GetValueOrDefault("icon-rid");
                var iconB64 = properties.GetValueOrDefault("icon-data");
                if (!string.IsNullOrEmpty(iconRid) && !string.IsNullOrEmpty(iconB64))
                {
                    var iconCt = properties.GetValueOrDefault("icon-content-type") ?? "image/x-emf";
                    var iconType = iconCt switch
                    {
                        "image/png" => ImagePartType.Png,
                        "image/jpeg" => ImagePartType.Jpeg,
                        "image/gif" => ImagePartType.Gif,
                        "image/bmp" => ImagePartType.Bmp,
                        "image/x-wmf" => ImagePartType.Wmf,
                        _ => ImagePartType.Emf,
                    };
                    var iconPart = oleWs.AddImagePart(iconType, iconRid!);
                    var iconBytes = Convert.FromBase64String(iconB64!);
                    using var s = new MemoryStream(iconBytes);
                    iconPart.FeedData(s);
                }

                // VML anchor shape + <legacyDrawing> reference.
                var vmlShapeXml = properties.GetValueOrDefault("vml-shape");
                if (!string.IsNullOrEmpty(vmlShapeXml))
                {
                    if (!oleWs.VmlDrawingParts.Any())
                    {
                        var vmlPart = oleWs.AddNewPart<VmlDrawingPart>();
                        using var writer = new System.IO.StreamWriter(vmlPart.GetStream());
                        writer.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"></xml>");
                    }
                    InsertVmlShapeXml(oleWs.VmlDrawingParts.First(), vmlShapeXml!);
                    var wsEl = GetSheet(oleWs);
                    if (wsEl.GetFirstChild<LegacyDrawing>() == null)
                    {
                        wsEl.AppendChild(new LegacyDrawing { Id = oleWs.GetIdOfPart(oleWs.VmlDrawingParts.First()) });
                    }
                }

                // <oleObjects> child, verbatim, inserted in schema order.
                var oleWsElement = GetSheet(oleWs);
                var oleObjects = oleWsElement.GetFirstChild<OleObjects>();
                if (oleObjects == null)
                {
                    oleObjects = new OleObjects();
                    oleWsElement.AppendChild(oleObjects);
                }
                OpenXmlElement oleChild = oleObjectXml.Contains("AlternateContent", StringComparison.Ordinal)
                    ? new DocumentFormat.OpenXml.AlternateContent(oleObjectXml)
                    : new OleObject(oleObjectXml);
                oleObjects.AppendChild(oleChild);
                ReorderWorksheetChildren(oleWsElement);
                SaveWorksheet(oleWs);

                return (oleRid, $"/{oleSheetName}/ole");
            }

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, chartex, drawing-group, ole");
        }
    }

    private static void RemapDrawingRelationshipId(
        OpenXmlElement root, string sourceId, string destinationId)
    {
        const string relNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        foreach (var element in root.Descendants().Prepend(root))
        {
            foreach (var attr in element.GetAttributes()
                .Where(a => a.NamespaceUri == relNs && a.Value == sourceId)
                .ToList())
            {
                element.SetAttribute(new OpenXmlAttribute(
                    attr.Prefix, attr.LocalName, attr.NamespaceUri, destinationId));
            }
        }
    }

    /// <summary>Append a verbatim &lt;v:shape&gt; slice before the VML part's
    /// closing &lt;/xml&gt;. Shared by the OLE carrier (comment shapes go
    /// through AppendCommentVmlShape which builds the shape itself).</summary>
    private static void InsertVmlShapeXml(VmlDrawingPart vmlPart, string shapeXml)
    {
        string xml;
        using (var reader = new System.IO.StreamReader(vmlPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read)))
            xml = reader.ReadToEnd();
        var closeIdx = xml.LastIndexOf("</xml>", StringComparison.OrdinalIgnoreCase);
        if (closeIdx < 0) return;
        xml = xml[..closeIdx] + shapeXml + xml[closeIdx..];
        using var writer = new System.IO.StreamWriter(vmlPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write));
        writer.Write(xml);
    }
}
