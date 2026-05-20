// Copyright 2025 OfficeCLI (officecli.ai)
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

            case "namedrange" or "definedname":
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
            return $"{effectiveParentPath}/row[{newRowIndex}]";
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
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => RemapRowsInRangeRef(r, map),
            formulaTextMapper: f => Core.FormulaRefShifter.ApplyRowRenumberMap(f, sheetName, sheetName, map));
    }

    private void ApplyColRenumberToSheet(WorksheetPart worksheet, string sheetName, IReadOnlyDictionary<int, int> map)
    {
        if (map.Count == 0) return;
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => RemapColsInRangeRef(r, map),
            formulaTextMapper: f => Core.FormulaRefShifter.ApplyColRenumberMap(f, sheetName, sheetName, map));
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

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }
}
