// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    // ==================== Range Set (merge/unmerge) ====================

    private List<string> SetRange(WorksheetPart worksheet, string rangeRef, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        // Separate range-level props from cell-level props
        var cellProps = new Dictionary<string, string>();
        string? stylePreset = null;
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
                    bool doSweep = value.Equals("sweep", StringComparison.OrdinalIgnoreCase);

                    // CONSISTENCY(cell-merge): cell-anchor Set accepts merge=<range>
                    // as the merge target; range-path Set must mirror that. If the
                    // value is a range-shaped literal that matches the path's rangeRef,
                    // treat it as merge=true (the range is already encoded in the path).
                    // A mismatch is a path-vs-value disagreement and must not be silent
                    // — historically it fell through to the unmerge branch and on a
                    // blank sheet became a silent no-op (issue #108).
                    if (!doMerge && !doSweep
                        && SingleMergeRefPattern.IsMatch(value.ToUpperInvariant()))
                    {
                        if (string.Equals(value, rangeRef, StringComparison.OrdinalIgnoreCase))
                        {
                            doMerge = true;
                        }
                        else
                        {
                            throw new ArgumentException(
                                $"merge value '{value}' does not match the range path '{rangeRef}'. " +
                                $"The range is already encoded in the path — pass merge=true, " +
                                $"or fix the path to match the intended merge range.");
                        }
                    }

                    if (doMerge)
                    {
                        // CONSISTENCY(merge-empty-container): pre-validate before container
                        // creation — see ExcelHandler.Helpers.ValidateMergeRefLiteral.
                        ValidateMergeRefLiteral(rangeRef);
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        if (mergeCells == null)
                        {
                            mergeCells = new MergeCells();
                            ws.AppendChild(mergeCells);
                        }

                        // CONSISTENCY(merge-comma): path is a single-target locator, not
                        // a list. Disjoint multi-range merges go through prop value form
                        // (`--prop merge=A1:B1,A2:B2`), at sheet- or cell-anchored set.
                        // A comma in the path itself is rejected by the guard inside
                        // InsertMergeCellChecked with an actionable message.
                        InsertMergeCellChecked(mergeCells, rangeRef, worksheet);
                        mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
                    }
                    else if (doSweep)
                    {
                        // Explicit "I know this is destructive": clear every merge whose ref
                        // lies entirely inside this range. Idempotent no-op when none.
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        if (mergeCells != null)
                        {
                            var contained = FindMergesContainedIn(mergeCells, rangeRef);
                            foreach (var refStr in contained)
                            {
                                var mc = mergeCells.Elements<MergeCell>()
                                    .FirstOrDefault(m => m.Reference?.Value == refStr);
                                mc?.Remove();
                            }
                            if (!mergeCells.HasChildren) mergeCells.Remove();
                            else mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
                        }
                    }
                    else
                    {
                        // Unmerge: remove the MergeCell whose ref exactly matches this range.
                        // CONSISTENCY(merge-precision): exact-match only. If the range covers
                        // sub-merges but does not equal one, fail with the precise refs the
                        // caller should use, rather than silently sweeping or no-op'ing.
                        // Pass merge=sweep to clear all sub-merges at once.
                        var mergeCells = ws.GetFirstChild<MergeCells>();
                        if (mergeCells != null)
                        {
                            var mc = mergeCells.Elements<MergeCell>()
                                .FirstOrDefault(m => m.Reference?.Value?.Equals(rangeRef, StringComparison.OrdinalIgnoreCase) == true);
                            if (mc != null)
                            {
                                mc.Remove();
                            }
                            else
                            {
                                var contained = FindMergesContainedIn(mergeCells, rangeRef);
                                if (contained.Count > 0)
                                {
                                    throw new CliException(
                                        $"Range {rangeRef} does not match an existing merge but contains {contained.Count} merge(s): " +
                                        string.Join(", ", contained) + ".")
                                    {
                                        Code = "merge_not_exact",
                                        Suggestion = $"Call merge=false on each precise range (e.g. /SheetName/{contained[0]} --prop merge=false), " +
                                                     $"or pass merge=sweep to clear all sub-merges in {rangeRef} at once.",
                                        ValidValues = contained.ToArray(),
                                    };
                                }
                                // else: nothing to unmerge anywhere in the range — idempotent no-op.
                            }

                            // Remove empty MergeCells element
                            if (!mergeCells.HasChildren)
                                mergeCells.Remove();
                            else
                                mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
                        }
                    }
                    break;
                }
                case "stylepreset":
                    // Range style preset (03 §6) — expanded after cellProps so
                    // the preset bundle wins conflicts deterministically.
                    stylePreset = value;
                    break;
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

        // Apply the style preset after explicit cellProps (preset wins conflicts;
        // its own per-row expansion is atomic via its internal sheetData backup).
        if (stylePreset != null)
            unsupported.AddRange(ApplyStylePresetRange(worksheet, rangeRef, stylePreset, cellProps));

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
            // Accept both the space form ("A desc") and the colon form
            // ("A:desc") — the latter is exactly what Get surfaces for a sort
            // state, so dump→replay of the canonical readback works. ':' is
            // never otherwise valid in a sort key, so treating it as a
            // separator is unambiguous.
            var tokens = spec.Trim().Split(new[] { ' ', '\t', ':' }, StringSplitOptions.RemoveEmptyEntries);
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
        // rewritten; refs outside are untouched. See the project conventions "Consistency > Robustness"
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

}
