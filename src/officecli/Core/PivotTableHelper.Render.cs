// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

internal static partial class PivotTableHelper
{
    // ==================== Pivot Output Renderer ====================

    /// <summary>
    /// Compute the pivot's aggregation matrix from columnData and write the
    /// rendered cells into targetSheet's SheetData. Mirrors what real Excel writes
    /// on save: literal cells with computed values, NOT a definition that Excel
    /// recomputes on open.
    ///
    /// Supported (v1): exactly 1 row field × 1 col field × 1 data field, with
    /// aggregator in {sum, count, average, min, max}, plus row/column/grand totals.
    /// Other configurations leave sheetData empty and emit a stderr warning so
    /// the file still validates and opens, just without rendered data.
    ///
    /// Layout (verified against Excel-authored sample):
    ///     Row 0:  [data caption] [col field caption]
    ///     Row 1:  [row field caption] [col label 1] [col label 2] ... [总计]
    ///     Row 2:  [row label 1]       [v]            [v]              [row total 1]
    ///     ...
    ///     Row N:  [总计]              [col total 1] [col total 2] ... [grand total]
    /// </summary>
    private static void RenderPivotIntoSheet(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields,
        List<int>? filterFieldIndices = null,
        uint?[]? columnStyleIds = null)
    {
        RenderPivotIntoSheetCore(targetSheet, position, headers, columnData,
            rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, columnStyleIds);
        // Every specialized renderer appends fresh <row>/<c> elements without
        // probing for pre-existing ones. When the pivot's footprint lands on
        // rows/cells that already exist (e.g. the anchor overlaps the source
        // range on the same sheet), that leaves duplicate row indices and —
        // fatally — duplicate cell references, which real Excel refuses
        // (0x800A03EC) while the CLI reported success. Merge duplicates
        // (pivot output wins, it was appended last) and restore row order.
        NormalizeRenderedSheetData(targetSheet);
    }

    /// <summary>Merge duplicate-RowIndex rows (later cells win per reference)
    /// and re-sort rows/cells, so a pivot rendered over existing content
    /// yields valid OOXML instead of duplicate r= entries.</summary>
    private static void NormalizeRenderedSheetData(WorksheetPart targetSheet)
    {
        var sheetData = targetSheet.Worksheet?.GetFirstChild<SheetData>();
        if (sheetData == null) return;
        var byIdx = new Dictionary<uint, Row>();
        foreach (var row in sheetData.Elements<Row>().ToList())
        {
            var idx = row.RowIndex?.Value ?? 0;
            if (!byIdx.TryGetValue(idx, out var first))
            {
                byIdx[idx] = row;
                continue;
            }
            foreach (var cell in row.Elements<Cell>().ToList())
            {
                cell.Remove();
                var cr = cell.CellReference?.Value;
                if (cr != null)
                    first.Elements<Cell>()
                        .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cr, StringComparison.OrdinalIgnoreCase))
                        ?.Remove();
                first.AppendChild(cell);
            }
            row.Remove();
        }
        // Restore row order and per-row cell column order.
        var orderedRows = sheetData.Elements<Row>()
            .OrderBy(r => r.RowIndex?.Value ?? 0).ToList();
        foreach (var r in orderedRows) r.Remove();
        foreach (var r in orderedRows)
        {
            var orderedCells = r.Elements<Cell>()
                .OrderBy(c =>
                {
                    var v = c.CellReference?.Value;
                    if (string.IsNullOrEmpty(v)) return int.MaxValue;
                    try { return ColToIndex(ParseCellRef(v!).col); }
                    catch { return int.MaxValue; }
                })
                .ToList();
            foreach (var c in orderedCells) c.Remove();
            foreach (var c in orderedCells) r.AppendChild(c);
            sheetData.AppendChild(r);
        }
    }

    private static void RenderPivotIntoSheetCore(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields,
        List<int>? filterFieldIndices = null,
        uint?[]? columnStyleIds = null)
    {
        // Per-data-field style index: pivot value cells for data field d inherit
        // the source column's StyleIndex (number format). A null entry means the
        // source cell had no explicit style → pivot cell stays General.
        int dataFieldCount = Math.Max(1, valueFields.Count);
        var valueStyleIds = new uint?[dataFieldCount];
        if (columnStyleIds != null)
        {
            for (int d = 0; d < valueFields.Count; d++)
            {
                var srcIdx = valueFields[d].idx;
                if (srcIdx >= 0 && srcIdx < columnStyleIds.Length)
                    valueStyleIds[d] = columnStyleIds[srcIdx];
            }
        }

        // v3 limits: dispatch based on field-count combinations.
        //   1 row × 1 col × K data → single-row K-data renderer below
        //   2 row × 1 col × 1 data → multi-row renderer (RenderMultiRowPivot)
        //   1 row × 2 col × 1 data → multi-col renderer (RenderMultiColPivot)
        // Other combinations fall back to empty skeleton with a warning.
        // N≥3 row or col fields → general tree-based renderer (handles arbitrary depth).
        // N≤2 cases continue to use the specialized renderers below for byte-level
        // backward compatibility (regression-tested via test-samples/pivot_baselines).
        //
        // Non-compact layouts (outline/tabular) always route through the general
        // renderer because specialized renderers hardcode compact-mode column
        // placement (all row labels in one column). The general renderer handles
        // multi-column row labels for outline/tabular.
        if (ActiveLayoutMode != "compact" && valueFields.Count >= 1)
        {
            RenderGeneralPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, valueStyleIds);
            return;
        }

        // Compact + multi-row + subtotals OFF also routes through the general
        // renderer. The N=2 specialized RenderMultiRowPivot lacks the
        // compactLabelRows path (label-only parent rows + indented children) and
        // falls back to an "outer / inner" string-concat hack on the first
        // leaf, which doesn't match Excel. The general renderer treats N≥2
        // compact+nosubtotals uniformly via its compactLabelRows branch.
        if (ActiveLayoutMode == "compact" && !ActiveDefaultSubtotal
            && rowFieldIndices.Count >= 2 && valueFields.Count >= 1)
        {
            RenderGeneralPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, valueStyleIds);
            return;
        }

        // Catch-all for field combinations not handled by the specialized N≤2
        // renderers below: 0×0, 0×1, 0×2, 2×0. RenderGeneralPivot handles
        // empty row/col axes naturally via empty AxisTrees.
        if (valueFields.Count >= 1
            && (rowFieldIndices.Count == 0 || (rowFieldIndices.Count == 2 && colFieldIndices.Count == 0)))
        {
            RenderGeneralPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, valueStyleIds);
            return;
        }

        if (rowFieldIndices.Count >= 3 || colFieldIndices.Count >= 3)
        {
            // CONSISTENCY(no-values-noop): RenderGeneralPivot dereferences
            // valueFields[0] for the data column anchor and crashes when the
            // user has moved every field to an axis (no values left). Skip
            // rendering — the pivotDef + cache survive so a subsequent Set
            // re-adds values cleanly.
            if (valueFields.Count == 0)
            {
                WarnRenderAdvisory(
                    "pivot has no value fields; skipping cell render. " +
                    "Add a value field to materialize the table.",
                    "pivot_render_skipped");
                return;
            }
            RenderGeneralPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, valueStyleIds);
            return;
        }

        if (rowFieldIndices.Count == 2 && colFieldIndices.Count == 2 && valueFields.Count >= 1)
        {
            RenderMatrixPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, valueStyleIds);
            return;
        }
        if (rowFieldIndices.Count == 2 && colFieldIndices.Count == 1 && valueFields.Count >= 1)
        {
            RenderMultiRowPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, valueStyleIds);
            return;
        }
        if (rowFieldIndices.Count == 1 && colFieldIndices.Count == 2 && valueFields.Count >= 1)
        {
            RenderMultiColPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices, valueStyleIds);
            return;
        }

        // Accept 1×1×K AND 1×0×K (rows-only). The 1×0 layout collapses the
        // column axis to a single synthetic bucket so the same matrix code
        // below produces one data column ("Total <name>" / value name) plus
        // the rightmost grand-total column.
        bool rowsOnly = rowFieldIndices.Count == 1 && colFieldIndices.Count == 0 && valueFields.Count >= 1;
        if (!rowsOnly && (rowFieldIndices.Count != 1 || colFieldIndices.Count != 1 || valueFields.Count < 1))
        {
            WarnRenderAdvisory(
                "pivot rendering currently supports 1×0×K, 1×1×K, 2×1×1, or 1×2×1 field combinations. " +
                "The file will open but the pivot will appear empty. " +
                "Use Excel's Refresh button to populate it manually.",
                "pivot_render_skipped");
            return;
        }

        var rowFieldIdx = rowFieldIndices[0];
        var colFieldIdx = rowsOnly ? -1 : colFieldIndices[0];
        var rowFieldName = headers[rowFieldIdx];
        // CONSISTENCY(rows-only-pivot): no col field → use empty caption so
        // the layout collapses cleanly. The K-column header path uses the
        // value field name as the only visible column label.
        var colFieldName = rowsOnly ? "" : headers[colFieldIdx];
        int K = valueFields.Count;

        var rowValues = columnData[rowFieldIdx];
        // Synthetic single-bucket col axis for rows-only: every source row
        // collapses into one column so Reduce/Aggregate machinery below stays
        // structurally identical to the 1×1×K path.
        var colValues = rowsOnly ? new string[rowValues.Length] : columnData[colFieldIdx];
        if (rowsOnly)
        {
            for (int i = 0; i < colValues.Length; i++) colValues[i] = "__total__";
        }

        // Unique row/col labels in cache order (alphabetical ordinal).
        var uniqueRows = rowValues.Where(v => !string.IsNullOrEmpty(v)).Distinct()
            .OrderByAxis(v => v).ToList();
        var uniqueCols = colValues.Where(v => !string.IsNullOrEmpty(v)).Distinct()
            .OrderByAxis(v => v).ToList();

        // Bucket source values per (rowLabel, colLabel, dataFieldIdx) so each data
        // field is aggregated independently. The aggregator function differs per
        // data field (sum/count/avg/...) so each bucket carries its own reducer.
        // Two data fields on the same source column are common (e.g. sum + count
        // of 金额) and produce two independent buckets keyed by their dataFieldIdx
        // in valueFields.
        var perBucket = new Dictionary<(string r, string c, int d), List<string>>();
        var perDataField = new List<List<string>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<string>());

        for (int i = 0; i < rowValues.Length; i++)
        {
            var rv = rowValues.Length > i ? rowValues[i] : null;
            var cv = colValues.Length > i ? colValues[i] : null;
            if (string.IsNullOrEmpty(rv) || string.IsNullOrEmpty(cv)) continue;

            for (int d = 0; d < K; d++)
            {
                var dataIdx = valueFields[d].idx;
                var dataValues = columnData[dataIdx];
                if (i >= dataValues.Length) continue;
                var raw = dataValues[i];

                var key = (rv, cv, d);
                if (!perBucket.TryGetValue(key, out var list))
                {
                    list = new List<string>();
                    perBucket[key] = list;
                }
                list.Add(raw);
                perDataField[d].Add(raw);
            }
        }

        double Reduce(IEnumerable<string> values, string func) => ReducePivotValues(values, func);

        // Compute the K-deep cell matrix + row/col/grand totals per data field.
        // matrix[r, c, d] = reduce(values for row r, col c, data field d)
        // rowTotals[r, d], colTotals[c, d], grandTotals[d] follow the same shape.
        var matrix = new double?[uniqueRows.Count, uniqueCols.Count, K];
        var rowTotals = new double[uniqueRows.Count, K];
        var colTotals = new double[uniqueCols.Count, K];
        var grandTotals = new double[K];
        for (int d = 0; d < K; d++)
        {
            var func = valueFields[d].func;
            for (int r = 0; r < uniqueRows.Count; r++)
            {
                var rowAll = new List<string>();
                for (int c = 0; c < uniqueCols.Count; c++)
                {
                    if (perBucket.TryGetValue((uniqueRows[r], uniqueCols[c], d), out var bucket) && bucket.Count > 0)
                    {
                        matrix[r, c, d] = Reduce(bucket, func);
                        rowAll.AddRange(bucket);
                    }
                }
                rowTotals[r, d] = Reduce(rowAll, func);
            }
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                var colAll = new List<string>();
                for (int r = 0; r < uniqueRows.Count; r++)
                {
                    if (perBucket.TryGetValue((uniqueRows[r], uniqueCols[c], d), out var bucket))
                        colAll.AddRange(bucket);
                }
                colTotals[c, d] = Reduce(colAll, func);
            }
            grandTotals[d] = Reduce(perDataField[d], func);
        }

        // showDataAs post-processing: transform raw aggregates into ratio /
        // running-total forms before they hit sheetData. Done per data field
        // so sum + percent_of_total can coexist in the same pivot. Cell values
        // for a data field are normalized against the corresponding total,
        // matching Excel's Show Values As semantics. See ParseShowDataAs for
        // the supported mode strings.
        //
        // Row/col/grand totals are transformed alongside the matrix so the
        // rendered totals stay consistent with the transformed data cells
        // (e.g. under percent_of_total, the grand total becomes 1.0).
        for (int d = 0; d < K; d++)
        {
            var mode = valueFields[d].showAs;
            ApplyShowDataAs1x1(mode, matrix, rowTotals, colTotals, grandTotals, uniqueRows.Count, uniqueCols.Count, d);
        }

        // ===== Write cells =====
        // For K=1, layout is 2 header rows: caption + col labels.
        // For K>1, layout is 3 header rows: caption + col labels + per-data-field
        // names repeated under each col label group. This matches the Excel sample
        // multi_data_authored.xlsx exactly.
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalColLabel = ActiveGrandTotalCaption;

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // ----- Row 0 (caption row) -----
        // Single data field: data field name in row-label col, col field name in first data col.
        // Multi data field: empty in row-label col, col field name (or "Values" placeholder) in first data col.
        var captionRow = new Row { RowIndex = (uint)anchorRow };
        if (K == 1)
            captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, valueFields[0].name));
        captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, colFieldName));
        sheetData.AppendChild(captionRow);

        // ----- Row 1 (col label row) -----
        // K=1: row field caption + col labels + grand total label
        // K>1: empty row-label cell + col labels at first col of each K-group + grand total labels
        var colLabelRowIdx = anchorRow + 1;
        var colLabelRow = new Row { RowIndex = (uint)colLabelRowIdx };
        if (K == 1)
        {
            colLabelRow.AppendChild(MakeStringCell(anchorColIdx, colLabelRowIdx, rowFieldName));
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                // Rows-only: the synthetic "__total__" bucket is invisible; show
                // the value field name as the single data column header.
                var label = rowsOnly ? valueFields[0].name : uniqueCols[c];
                colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + c, colLabelRowIdx, label));
            }
            // CONSISTENCY(grand-totals): rowGrandTotals=false drops the rightmost
            // 总计 column entirely — header label, per-row totals, and the grand
            // total row's rightmost cells all gated on ActiveRowGrandTotals.
            // For rows-only the only data column already IS the value's grand
            // total, so we suppress the duplicate trailing 总计 column.
            if (ActiveRowGrandTotals && !rowsOnly)
                colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + uniqueCols.Count, colLabelRowIdx, totalColLabel));
        }
        else if (rowsOnly)
        {
            // R4-2: rows-only multi-data pivot has a synthetic "__total__"
            // col bucket and its K data cells ARE the grand totals, so we
            // skip the col-label row entirely (no sentinel, no "Total Sum").
            // Data field names are emitted on a dedicated row below.
        }
        else
        {
            // First col of each K-group gets the col label; the K-1 cells after are
            // visually spanned in Excel's renderer but we leave them empty in
            // sheetData (Excel handles the visual span via colItems metadata).
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                int colStart = anchorColIdx + 1 + c * K;
                colLabelRow.AppendChild(MakeStringCell(colStart, colLabelRowIdx, uniqueCols[c]));
            }
            // Grand total area: K cells, one per data field, labeled "Total <name>"
            if (ActiveRowGrandTotals)
            {
                int totalStart = anchorColIdx + 1 + uniqueCols.Count * K;
                for (int d = 0; d < K; d++)
                    colLabelRow.AppendChild(MakeStringCell(totalStart + d, colLabelRowIdx, "Total " + valueFields[d].name));
            }
        }
        sheetData.AppendChild(colLabelRow);

        // ----- Row 2 (data field name row, only when K>1) -----
        int firstDataRow;
        if (K > 1)
        {
            var dfNameRowIdx = anchorRow + 2;
            var dfNameRow = new Row { RowIndex = (uint)dfNameRowIdx };
            // row label column gets the row field name
            dfNameRow.AppendChild(MakeStringCell(anchorColIdx, dfNameRowIdx, rowFieldName));
            // Repeat data field names under each col label group
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                for (int d = 0; d < K; d++)
                {
                    int colIdx = anchorColIdx + 1 + c * K + d;
                    dfNameRow.AppendChild(MakeStringCell(colIdx, dfNameRowIdx, valueFields[d].name));
                }
            }
            // No data field names under the grand total cols — row 1 already
            // labeled them with "Total <name>" so they are self-describing.
            sheetData.AppendChild(dfNameRow);
            firstDataRow = anchorRow + 3;
        }
        else
        {
            firstDataRow = anchorRow + 2;
        }

        // ----- Data rows -----
        for (int r = 0; r < uniqueRows.Count; r++)
        {
            var rowIdx = firstDataRow + r;
            var dataRow = new Row { RowIndex = (uint)rowIdx };
            dataRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, uniqueRows[r]));
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                for (int d = 0; d < K; d++)
                {
                    int colIdx = anchorColIdx + 1 + c * K + d;
                    var v = matrix[r, c, d];
                    if (v.HasValue)
                        dataRow.AppendChild(MakeNumericCell(colIdx, rowIdx, v.Value, valueStyleIds[d]));
                }
            }
            // Row totals — K cells (one per data field).
            // CONSISTENCY(grand-totals): gated on ActiveRowGrandTotals so the
            // rightmost 总计 column disappears entirely when grandTotals=none|cols.
            // Rows-only: the K data cells already ARE the row totals (single
            // synthetic col bucket), so the trailing duplicate is omitted.
            if (ActiveRowGrandTotals && !rowsOnly)
            {
                int rowTotalStart = anchorColIdx + 1 + uniqueCols.Count * K;
                for (int d = 0; d < K; d++)
                    dataRow.AppendChild(MakeNumericCell(rowTotalStart + d, rowIdx, rowTotals[r, d], valueStyleIds[d]));
            }
            sheetData.AppendChild(dataRow);
        }

        // ----- Grand total row -----
        // CONSISTENCY(grand-totals): the entire bottom 总计 row is omitted
        // when ActiveColGrandTotals is false (grandTotals=none|rows). The
        // rightmost cells inside the row are independently gated on
        // ActiveRowGrandTotals so grandTotals=cols still renders the bottom
        // row but without the trailing K row-grand cells.
        if (ActiveColGrandTotals)
        {
            var grandRowIdx = firstDataRow + uniqueRows.Count;
            var grandRow = new Row { RowIndex = (uint)grandRowIdx };
            grandRow.AppendChild(MakeStringCell(anchorColIdx, grandRowIdx, totalColLabel));
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                for (int d = 0; d < K; d++)
                {
                    int colIdx = anchorColIdx + 1 + c * K + d;
                    grandRow.AppendChild(MakeNumericCell(colIdx, grandRowIdx, colTotals[c, d], valueStyleIds[d]));
                }
            }
            if (ActiveRowGrandTotals && !rowsOnly)
            {
                int grandTotalStart = anchorColIdx + 1 + uniqueCols.Count * K;
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(grandTotalStart + d, grandRowIdx, grandTotals[d], valueStyleIds[d]));
            }
            sheetData.AppendChild(grandRow);
        }

        // Page filter cells: rendered ABOVE the table at rows
        // (anchorRow - filterCount - 1) ... (anchorRow - 2). One row per filter
        // field, with field name in the row-label column and "(All)" in the
        // adjacent data column. Row (anchorRow - 1) is left empty as a visual gap.
        //
        // Page filters are NOT inside <location ref/> per ECMA-376; they are
        // separate visual cells whose presence is signalled by the rowPageCount /
        // colPageCount attributes on pivotTableDefinition (already set in
        // BuildPivotTableDefinition). Excel pairs the filter cells with the pivot
        // by their position above the location range.
        //
        // If there isn't enough room above (e.g. user anchored at F1), we skip the
        // visible cells but the pivot definition still tags them as page fields,
        // so the dropdowns appear in Excel's pivot UI even without the cell labels.
        if (filterFieldIndices != null && filterFieldIndices.Count > 0)
        {
            var requiredHeadroom = filterFieldIndices.Count + 1; // filter rows + 1 gap
            if (anchorRow > requiredHeadroom)
            {
                var firstFilterRow = anchorRow - requiredHeadroom;
                for (int fi = 0; fi < filterFieldIndices.Count; fi++)
                {
                    var fIdx = filterFieldIndices[fi];
                    if (fIdx < 0 || fIdx >= headers.Length) continue;
                    var rowIdx = firstFilterRow + fi;
                    var filterRow = new Row { RowIndex = (uint)rowIdx };
                    filterRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, headers[fIdx]));
                    // Round-trip preservation: if the user has manually set a
                    // locale-specific label (e.g. "(全部)" / "(Tous)") on this
                    // filter cell in a previous edit, keep it. Fall back to the
                    // English default only when the cell is missing or empty.
                    var filterAllLabel = ReadExistingStringAtOrDefault(
                        targetSheet, sheetData, anchorColIdx + 1, rowIdx, "(All)");
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, filterAllLabel));
                    // Insert in row order: existing rows in sheetData start at
                    // anchorRow, so prepend the filter rows to the front.
                    sheetData.InsertAt(filterRow, fi);
                }
            }
            else
            {
                WarnRenderAdvisory(
                    $"pivot at {position} has {filterFieldIndices.Count} page filter(s) " +
                    $"but only {anchorRow - 1} row(s) of headroom above. " +
                    "Filter cells will not be visible in the host sheet, but the filter dropdowns " +
                    "will still appear in Excel's pivot UI. Move the pivot to a lower anchor row " +
                    $"(at least row {requiredHeadroom + 1}) to render the filter cells.",
                    "pivot_filter_cells_skipped");
            }
        }

        ws.Save();
    }

    /// <summary>
    /// Render a 2-row-field pivot. Compact-mode layout (verified against
    /// multi_row_authored.xlsx with rows=地区,城市):
    ///
    ///     A                  B           C           D
    ///   3 [data caption]     [col field caption]
    ///   4 Row Labels         咖啡        奶茶        Grand Total
    ///   5 华东                200        260         460          <- outer subtotal
    ///   6   上海              200        150         350
    ///   7   杭州                         110         110
    ///   8 华北                215        85          300          <- outer subtotal
    ///   ...
    ///   N Grand Total        595        345         940
    ///
    /// Both outer and inner labels live in column A (compact mode collapses the
    /// row-label area into a single column, with Excel auto-indenting inners
    /// visually). Each outer value gets its own subtotal row showing the
    /// aggregate across all its existing inners; only (outer, inner) pairs that
    /// actually appear in the source data are rendered (Excel does not enumerate
    /// empty cartesian cells).
    ///
    /// Multi data fields (K>1) are not yet supported in this code path — would
    /// need to extend col multiplication and add the third "data field name"
    /// header row. v4 expansion. Tracked.
    /// </summary>
    private static void RenderMultiRowPivot(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields,
        List<int>? filterFieldIndices,
        uint?[] valueStyleIds)
    {
        var outerFieldIdx = rowFieldIndices[0];
        var innerFieldIdx = rowFieldIndices[1];
        var colFieldIdx = colFieldIndices[0];
        int K = valueFields.Count;

        var outerVals = columnData[outerFieldIdx];
        var innerVals = columnData[innerFieldIdx];
        var colVals = columnData[colFieldIdx];
        var colFieldName = headers[colFieldIdx];

        // Build the same (outer → [inners]) groups used by BuildMultiRowItems so
        // the rendered cells match the rowItems indices position-for-position.
        var groups = BuildOuterInnerGroups(outerFieldIdx, innerFieldIdx, columnData);
        var uniqueCols = colVals.Where(v => !string.IsNullOrEmpty(v)).Distinct()
            .OrderByAxis(v => v).ToList();

        // Aggregate per (outer, inner, col, dataFieldIdx). For K=1 the d
        // dimension is degenerate but the same data structure works uniformly.
        var leafBucket = new Dictionary<(string o, string i, string c, int d), List<string>>();
        var perDataField = new List<List<string>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<string>());

        for (int i = 0; i < outerVals.Length; i++)
        {
            var ov = outerVals.Length > i ? outerVals[i] : null;
            var iv = innerVals.Length > i ? innerVals[i] : null;
            var cv = colVals.Length > i ? colVals[i] : null;
            if (string.IsNullOrEmpty(ov) || string.IsNullOrEmpty(iv) || string.IsNullOrEmpty(cv)) continue;

            for (int d = 0; d < K; d++)
            {
                var dataIdx = valueFields[d].idx;
                var dataValues = columnData[dataIdx];
                if (i >= dataValues.Length) continue;
                var raw = dataValues[i];

                var key = (ov, iv, cv, d);
                if (!leafBucket.TryGetValue(key, out var list))
                {
                    list = new List<string>();
                    leafBucket[key] = list;
                }
                list.Add(raw);
                perDataField[d].Add(raw);
            }
        }

        double Reduce(IEnumerable<string> values, string func) => ReducePivotValues(values, func);

        // The closures below compute the cell values per (row pos, col pos, d)
        // by reducing raw value lists. Each closure takes a data field index d
        // so each data field aggregates with its own function (sum/count/avg/...).
        double LeafCell(string outer, string inner, string col, int d)
            => leafBucket.TryGetValue((outer, inner, col, d), out var b) && b.Count > 0
                ? Reduce(b, valueFields[d].func) : double.NaN;

        double OuterSubtotalForCol(string outer, string col, int d)
        {
            var all = new List<string>();
            foreach (var (o, inners) in groups)
                if (o == outer)
                    foreach (var inner in inners)
                        if (leafBucket.TryGetValue((outer, inner, col, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double LeafRowTotal(string outer, string inner, int d)
        {
            var all = new List<string>();
            foreach (var col in uniqueCols)
                if (leafBucket.TryGetValue((outer, inner, col, d), out var b))
                    all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterRowTotal(string outer, int d)
        {
            var all = new List<string>();
            foreach (var (o, inners) in groups)
                if (o == outer)
                    foreach (var inner in inners)
                        foreach (var col in uniqueCols)
                            if (leafBucket.TryGetValue((outer, inner, col, d), out var b))
                                all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double ColTotal(string col, int d)
        {
            var all = new List<string>();
            foreach (var (outer, inners) in groups)
                foreach (var inner in inners)
                    if (leafBucket.TryGetValue((outer, inner, col, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        // ===== Write cells =====
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalLabel = ActiveGrandTotalCaption;

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // Helper: column index of leaf cell for col label c, data field d.
        int LeafColIdx(int c, int d) => anchorColIdx + 1 + c * K + d;
        // Helper: column index of grand-total cell for data field d.
        int GrandTotalColIdx(int d) => anchorColIdx + 1 + uniqueCols.Count * K + d;

        // CONSISTENCY(grand-totals): mirror the 1×1×K renderer's gating. Right
        // grand-total column = ActiveRowGrandTotals; bottom grand-total row =
        // ActiveColGrandTotals. Cached once per render call.
        bool emitRowGrand = ActiveRowGrandTotals;
        bool emitColGrand = ActiveColGrandTotals;

        // ----- Row 0 (caption row) -----
        // K=1: data field name + col field name
        // K>1: empty + col field name (data caption is implicit per col group)
        var captionRow = new Row { RowIndex = (uint)anchorRow };
        if (K == 1)
            captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, valueFields[0].name));
        captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, colFieldName));
        sheetData.AppendChild(captionRow);

        // ----- Row 1 (col label row) -----
        // K=1: row field name + col labels + 总计
        // K>1: empty + col labels at first col of each K-group + "Total <name>" cells
        var colLabelRowIdx = anchorRow + 1;
        var colLabelRow = new Row { RowIndex = (uint)colLabelRowIdx };
        if (K == 1)
        {
            colLabelRow.AppendChild(MakeStringCell(anchorColIdx, colLabelRowIdx, headers[outerFieldIdx]));
            for (int c = 0; c < uniqueCols.Count; c++)
                colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + c, colLabelRowIdx, uniqueCols[c]));
            if (emitRowGrand)
                colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + uniqueCols.Count, colLabelRowIdx, totalLabel));
        }
        else
        {
            for (int c = 0; c < uniqueCols.Count; c++)
                colLabelRow.AppendChild(MakeStringCell(LeafColIdx(c, 0), colLabelRowIdx, uniqueCols[c]));
            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    colLabelRow.AppendChild(MakeStringCell(GrandTotalColIdx(d), colLabelRowIdx, "Total " + valueFields[d].name));
            }
        }
        sheetData.AppendChild(colLabelRow);

        // ----- Row 2 (data field name row, only when K>1) -----
        int firstDataRow;
        if (K > 1)
        {
            var dfNameRowIdx = anchorRow + 2;
            var dfNameRow = new Row { RowIndex = (uint)dfNameRowIdx };
            dfNameRow.AppendChild(MakeStringCell(anchorColIdx, dfNameRowIdx, headers[outerFieldIdx]));
            for (int c = 0; c < uniqueCols.Count; c++)
                for (int d = 0; d < K; d++)
                    dfNameRow.AppendChild(MakeStringCell(LeafColIdx(c, d), dfNameRowIdx, valueFields[d].name));
            sheetData.AppendChild(dfNameRow);
            firstDataRow = anchorRow + 3;
        }
        else
        {
            firstDataRow = anchorRow + 2;
        }

        // CONSISTENCY(subtotals-opts): cache the subtotals toggle once per
        // render call. When off, skip the outer subtotal row emit AND change
        // the leaf row label from "inner only" to "outer > inner" so each
        // group is still visually identifiable in compact mode.
        bool emitSubtotals = ActiveDefaultSubtotal;

        // ----- Data rows -----
        int currentRow = firstDataRow;
        foreach (var (outer, inners) in groups)
        {
            if (emitSubtotals)
            {
                // Outer subtotal row: K cells per col + K cells in grand total area.
                var subRow = new Row { RowIndex = (uint)currentRow };
                subRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, outer));
                for (int c = 0; c < uniqueCols.Count; c++)
                {
                    bool any = HasAnyValueInOuterCol(outer, uniqueCols[c], groups, leafBucket, K);
                    for (int d = 0; d < K; d++)
                    {
                        var v = OuterSubtotalForCol(outer, uniqueCols[c], d);
                        if (any || v != 0)
                            subRow.AppendChild(MakeNumericCell(LeafColIdx(c, d), currentRow, v, valueStyleIds[d]));
                    }
                }
                if (emitRowGrand)
                {
                    for (int d = 0; d < K; d++)
                        subRow.AppendChild(MakeNumericCell(GrandTotalColIdx(d), currentRow, OuterRowTotal(outer, d), valueStyleIds[d]));
                }
                sheetData.AppendChild(subRow);
                currentRow++;
            }

            // Leaf rows for each existing (outer, inner) combo.
            bool firstLeafOfGroup = true;
            foreach (var inner in inners)
            {
                var leafRow = new Row { RowIndex = (uint)currentRow };
                // When subtotals are off, prefix the FIRST leaf of each group
                // with the outer label so users can still tell which group
                // they're in. Subsequent leaves just carry the inner label
                // (Excel's compact mode already indents them under the outer).
                var label = (!emitSubtotals && firstLeafOfGroup)
                    ? $"{outer} / {inner}"
                    : inner;
                leafRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, label));
                firstLeafOfGroup = false;
                for (int c = 0; c < uniqueCols.Count; c++)
                {
                    for (int d = 0; d < K; d++)
                    {
                        var v = LeafCell(outer, inner, uniqueCols[c], d);
                        if (!double.IsNaN(v))
                            leafRow.AppendChild(MakeNumericCell(LeafColIdx(c, d), currentRow, v, valueStyleIds[d]));
                    }
                }
                if (emitRowGrand)
                {
                    for (int d = 0; d < K; d++)
                        leafRow.AppendChild(MakeNumericCell(GrandTotalColIdx(d), currentRow, LeafRowTotal(outer, inner, d), valueStyleIds[d]));
                }
                sheetData.AppendChild(leafRow);
                currentRow++;
            }
        }

        // Grand total row.
        if (emitColGrand)
        {
            var grandRow = new Row { RowIndex = (uint)currentRow };
            grandRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, totalLabel));
            for (int c = 0; c < uniqueCols.Count; c++)
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(LeafColIdx(c, d), currentRow, ColTotal(uniqueCols[c], d), valueStyleIds[d]));
            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(GrandTotalColIdx(d), currentRow,
                        Reduce(perDataField[d], valueFields[d].func), valueStyleIds[d]));
            }
            sheetData.AppendChild(grandRow);
        }

        // Page filter cells reuse the single-row path's logic — same shape, same
        // layout above the table. RenderPivotIntoSheet handles them; we don't
        // duplicate the code, but if the user really needs filters with 2 row
        // fields, they should still get rendered. v4 candidate to factor out.
        // (Currently filters on multi-row pivots will write the page filter
        // markers in the pivot definition but no visible filter cells above
        // the table. Same warning is emitted.)
        if (filterFieldIndices != null && filterFieldIndices.Count > 0)
        {
            var requiredHeadroom = filterFieldIndices.Count + 1;
            if (anchorRow > requiredHeadroom)
            {
                var firstFilterRow = anchorRow - requiredHeadroom;
                for (int fi = 0; fi < filterFieldIndices.Count; fi++)
                {
                    var fIdx = filterFieldIndices[fi];
                    if (fIdx < 0 || fIdx >= headers.Length) continue;
                    var rowIdx = firstFilterRow + fi;
                    var filterRow = new Row { RowIndex = (uint)rowIdx };
                    filterRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, headers[fIdx]));
                    // Round-trip preservation: if the user has manually set a
                    // locale-specific label (e.g. "(全部)" / "(Tous)") on this
                    // filter cell in a previous edit, keep it. Fall back to the
                    // English default only when the cell is missing or empty.
                    var filterAllLabel = ReadExistingStringAtOrDefault(
                        targetSheet, sheetData, anchorColIdx + 1, rowIdx, "(All)");
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, filterAllLabel));
                    sheetData.InsertAt(filterRow, fi);
                }
            }
        }

        ws.Save();
    }

    /// <summary>
    /// Render a 1-row × 2-col pivot with hierarchical column subtotals. Compact
    /// mode layout (verified against multi_col_authored.xlsx, cols=产品,包装):
    ///
    ///     A          B        C        D            E         F        G          H
    ///   3 [data cap] [col field caption]
    ///   4            咖啡                            奶茶
    ///   5 Row Labels 罐装     袋装     咖啡 Total    罐装      袋装     奶茶 Tot.  Grand Total
    ///   6 华东       200               200           150                150        350
    ///   7 华北       120      80       200           85                 85         285
    ///   ...
    ///   N Grand Tot. 320      80       400           195       150      345        745
    ///
    /// Each outer col value gets its own subtotal column, then a final grand
    /// total column. Only (outer, inner) col combinations that exist in the
    /// data are rendered (matching Excel's behavior). Three header rows total
    /// (caption, outer col labels, inner col labels) — same as the multi-data
    /// case, so firstDataRow=3.
    ///
    /// Limitation: K=1 data field only. Multi-col + multi-data is a v4
    /// expansion; the col layout would multiply by K just like the single-col
    /// multi-data path does.
    /// </summary>
    private static void RenderMultiColPivot(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields,
        List<int>? filterFieldIndices,
        uint?[] valueStyleIds)
    {
        var rowFieldIdx = rowFieldIndices[0];
        var outerColIdx = colFieldIndices[0];
        var innerColIdx = colFieldIndices[1];
        int K = valueFields.Count;

        var rowVals = columnData[rowFieldIdx];
        var outerColVals = columnData[outerColIdx];
        var innerColVals = columnData[innerColIdx];

        var colGroups = BuildOuterInnerGroups(outerColIdx, innerColIdx, columnData);
        var uniqueRows = rowVals.Where(v => !string.IsNullOrEmpty(v)).Distinct()
            .OrderByAxis(v => v).ToList();

        // Aggregate per (row, outerCol, innerCol, dataFieldIdx). For K=1 the d
        // dimension is degenerate but the same data structure works uniformly.
        var leafBucket = new Dictionary<(string r, string oc, string ic, int d), List<string>>();
        var perDataField = new List<List<string>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<string>());

        for (int i = 0; i < rowVals.Length; i++)
        {
            var rv = rowVals.Length > i ? rowVals[i] : null;
            var ocv = outerColVals.Length > i ? outerColVals[i] : null;
            var icv = innerColVals.Length > i ? innerColVals[i] : null;
            if (string.IsNullOrEmpty(rv) || string.IsNullOrEmpty(ocv) || string.IsNullOrEmpty(icv)) continue;

            for (int d = 0; d < K; d++)
            {
                var dataIdx = valueFields[d].idx;
                var dataValues = columnData[dataIdx];
                if (i >= dataValues.Length) continue;
                var raw = dataValues[i];

                var key = (rv, ocv, icv, d);
                if (!leafBucket.TryGetValue(key, out var list))
                {
                    list = new List<string>();
                    leafBucket[key] = list;
                }
                list.Add(raw);
                perDataField[d].Add(raw);
            }
        }

        double Reduce(IEnumerable<string> values, string func) => ReducePivotValues(values, func);

        // Per-(row, outerCol, innerCol, d) reductions over raw values.
        double LeafCell(string row, string outerCol, string innerCol, int d)
            => leafBucket.TryGetValue((row, outerCol, innerCol, d), out var b) && b.Count > 0
                ? Reduce(b, valueFields[d].func) : double.NaN;

        double OuterColSubtotalForRow(string row, string outerCol, int d)
        {
            var all = new List<string>();
            foreach (var (oc, inners) in colGroups)
                if (oc == outerCol)
                    foreach (var inner in inners)
                        if (leafBucket.TryGetValue((row, outerCol, inner, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double RowGrandTotal(string row, int d)
        {
            var all = new List<string>();
            foreach (var (oc, inners) in colGroups)
                foreach (var inner in inners)
                    if (leafBucket.TryGetValue((row, oc, inner, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double LeafColTotal(string outerCol, string innerCol, int d)
        {
            var all = new List<string>();
            foreach (var row in uniqueRows)
                if (leafBucket.TryGetValue((row, outerCol, innerCol, d), out var b))
                    all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterColTotal(string outerCol, int d)
        {
            var all = new List<string>();
            foreach (var (oc, inners) in colGroups)
                if (oc == outerCol)
                    foreach (var inner in inners)
                        foreach (var row in uniqueRows)
                            if (leafBucket.TryGetValue((row, outerCol, inner, d), out var b))
                                all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        // ===== Write cells =====
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalLabel = ActiveGrandTotalCaption;

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // CONSISTENCY(grand-totals): cache the grand totals toggles once per
        // render call. emitRowGrand controls the right grand-total column
        // block; emitColGrand controls the bottom grand-total row.
        bool emitRowGrand = ActiveRowGrandTotals;
        bool emitColGrand = ActiveColGrandTotals;

        // Pre-compute absolute column indices. K data fields multiply the leaf
        // and subtotal positions by K. Layout (left to right):
        //   row label
        //   For each outer:
        //     For each inner:                            K cells (data fields)
        //     subtotal:                                  K cells (per-data subtotal)
        //   grand total:                                 K cells (per-data grand)
        // The grand total column block is skipped entirely when emitRowGrand=false.
        // CONSISTENCY(subtotals-opts): cached once per render call.
        bool emitSubtotals = ActiveDefaultSubtotal;

        var leafColPositions = new Dictionary<(string outer, string inner, int d), int>();
        var subtotalColPositions = new Dictionary<(string outer, int d), int>();
        var grandTotalColPositions = new int[K];
        int currentCol = anchorColIdx + 1;
        foreach (var (outer, inners) in colGroups)
        {
            foreach (var inner in inners)
            {
                for (int d = 0; d < K; d++)
                {
                    leafColPositions[(outer, inner, d)] = currentCol;
                    currentCol++;
                }
            }
            if (emitSubtotals)
            {
                for (int d = 0; d < K; d++)
                {
                    subtotalColPositions[(outer, d)] = currentCol;
                    currentCol++;
                }
            }
        }
        if (emitRowGrand)
        {
            for (int d = 0; d < K; d++)
            {
                grandTotalColPositions[d] = currentCol;
                currentCol++;
            }
        }

        // ----- Header rows -----
        // K=1 → 3 header rows (caption, outer col labels, inner col labels)
        // K>1 → 4 header rows (caption, outer col labels + subtotal/grand-total
        //                      labels in same row, inner col labels, data field names)
        if (K == 1)
        {
            // Row 0 (caption): data field name + col field name.
            var captionRow = new Row { RowIndex = (uint)anchorRow };
            captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, valueFields[0].name));
            captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, headers[outerColIdx]));
            sheetData.AppendChild(captionRow);

            // Row 1 (outer col header): outer col label at first leaf col of each group.
            var outerHeaderRowIdx = anchorRow + 1;
            var outerHeaderRow = new Row { RowIndex = (uint)outerHeaderRowIdx };
            foreach (var (outer, inners) in colGroups)
            {
                int firstLeafCol = leafColPositions[(outer, inners[0], 0)];
                outerHeaderRow.AppendChild(MakeStringCell(firstLeafCol, outerHeaderRowIdx, outer));
            }
            sheetData.AppendChild(outerHeaderRow);

            // Row 2 (inner col header): row field caption + inner col labels +
            //                            "<outer> Total" at subtotal cols + "总计" at grand.
            var innerHeaderRowIdx = anchorRow + 2;
            var innerHeaderRow = new Row { RowIndex = (uint)innerHeaderRowIdx };
            innerHeaderRow.AppendChild(MakeStringCell(anchorColIdx, innerHeaderRowIdx, headers[rowFieldIdx]));
            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                    innerHeaderRow.AppendChild(MakeStringCell(leafColPositions[(outer, inner, 0)], innerHeaderRowIdx, inner));
                if (emitSubtotals)
                    innerHeaderRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, 0)], innerHeaderRowIdx, outer + " Total"));
            }
            if (emitRowGrand)
                innerHeaderRow.AppendChild(MakeStringCell(grandTotalColPositions[0], innerHeaderRowIdx, totalLabel));
            sheetData.AppendChild(innerHeaderRow);
        }
        else
        {
            // Row 0 (caption): only the col field caption (no data caption when K>1).
            var captionRow = new Row { RowIndex = (uint)anchorRow };
            captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, headers[outerColIdx]));
            sheetData.AppendChild(captionRow);

            // Row 1 (outer col header): outer label at first leaf col of group +
            // per-subtotal labels "<outer> <data field>" + grand total labels
            // "Total <data field>". This is verified against multi_col_K_authored.xlsx
            // where the subtotal labels live in row 4 (the outer header row) NOT
            // in the inner-label or data-field rows below.
            var outerHeaderRowIdx = anchorRow + 1;
            var outerHeaderRow = new Row { RowIndex = (uint)outerHeaderRowIdx };
            foreach (var (outer, inners) in colGroups)
            {
                int firstLeafCol = leafColPositions[(outer, inners[0], 0)];
                outerHeaderRow.AppendChild(MakeStringCell(firstLeafCol, outerHeaderRowIdx, outer));
                if (emitSubtotals)
                {
                    for (int d = 0; d < K; d++)
                        outerHeaderRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, d)],
                            outerHeaderRowIdx, $"{outer} {valueFields[d].name}"));
                }
            }
            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    outerHeaderRow.AppendChild(MakeStringCell(grandTotalColPositions[d],
                        outerHeaderRowIdx, $"Total {valueFields[d].name}"));
            }
            sheetData.AppendChild(outerHeaderRow);

            // Row 2 (inner col header): inner label at the first data col of each
            // (outer, inner) sub-group. Subtotal/grand-total cols are EMPTY in this
            // row (their labels live one row above).
            var innerHeaderRowIdx = anchorRow + 2;
            var innerHeaderRow = new Row { RowIndex = (uint)innerHeaderRowIdx };
            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                    innerHeaderRow.AppendChild(MakeStringCell(leafColPositions[(outer, inner, 0)],
                        innerHeaderRowIdx, inner));
            }
            sheetData.AppendChild(innerHeaderRow);

            // Row 3 (data field name row): row field caption + data field name at
            // every leaf col. Subtotal/grand-total cols stay empty (already labeled
            // in the outer header row above).
            var dfNameRowIdx = anchorRow + 3;
            var dfNameRow = new Row { RowIndex = (uint)dfNameRowIdx };
            dfNameRow.AppendChild(MakeStringCell(anchorColIdx, dfNameRowIdx, headers[rowFieldIdx]));
            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                    for (int d = 0; d < K; d++)
                        dfNameRow.AppendChild(MakeStringCell(leafColPositions[(outer, inner, d)],
                            dfNameRowIdx, valueFields[d].name));
            }
            sheetData.AppendChild(dfNameRow);
        }

        // ----- Data rows -----
        int firstDataRow = anchorRow + (K == 1 ? 3 : 4);
        for (int r = 0; r < uniqueRows.Count; r++)
        {
            var rowIdx = firstDataRow + r;
            var dataRow = new Row { RowIndex = (uint)rowIdx };
            dataRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, uniqueRows[r]));

            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                {
                    for (int d = 0; d < K; d++)
                    {
                        var v = LeafCell(uniqueRows[r], outer, inner, d);
                        if (!double.IsNaN(v))
                            dataRow.AppendChild(MakeNumericCell(leafColPositions[(outer, inner, d)], rowIdx, v, valueStyleIds[d]));
                    }
                }
                if (emitSubtotals)
                {
                    // Outer col subtotal cells (K per outer).
                    bool any = HasAnyValueInRowOuter(uniqueRows[r], outer, colGroups, leafBucket, K);
                    for (int d = 0; d < K; d++)
                    {
                        var sub = OuterColSubtotalForRow(uniqueRows[r], outer, d);
                        if (sub != 0 || any)
                            dataRow.AppendChild(MakeNumericCell(subtotalColPositions[(outer, d)], rowIdx, sub, valueStyleIds[d]));
                    }
                }
            }

            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    dataRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], rowIdx, RowGrandTotal(uniqueRows[r], d), valueStyleIds[d]));
            }
            sheetData.AppendChild(dataRow);
        }

        // Grand total row.
        if (emitColGrand)
        {
            int grandRowIdx = firstDataRow + uniqueRows.Count;
            var grandRow = new Row { RowIndex = (uint)grandRowIdx };
            grandRow.AppendChild(MakeStringCell(anchorColIdx, grandRowIdx, totalLabel));
            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                    for (int d = 0; d < K; d++)
                        grandRow.AppendChild(MakeNumericCell(leafColPositions[(outer, inner, d)], grandRowIdx,
                            LeafColTotal(outer, inner, d), valueStyleIds[d]));
                if (emitSubtotals)
                {
                    for (int d = 0; d < K; d++)
                        grandRow.AppendChild(MakeNumericCell(subtotalColPositions[(outer, d)], grandRowIdx, OuterColTotal(outer, d), valueStyleIds[d]));
                }
            }
            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], grandRowIdx,
                        Reduce(perDataField[d], valueFields[d].func), valueStyleIds[d]));
            }
            sheetData.AppendChild(grandRow);
        }

        // Page filter cells (same logic as the single-row renderer).
        if (filterFieldIndices != null && filterFieldIndices.Count > 0)
        {
            var requiredHeadroom = filterFieldIndices.Count + 1;
            if (anchorRow > requiredHeadroom)
            {
                var firstFilterRow = anchorRow - requiredHeadroom;
                for (int fi = 0; fi < filterFieldIndices.Count; fi++)
                {
                    var fIdx = filterFieldIndices[fi];
                    if (fIdx < 0 || fIdx >= headers.Length) continue;
                    var rowIdx = firstFilterRow + fi;
                    var filterRow = new Row { RowIndex = (uint)rowIdx };
                    filterRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, headers[fIdx]));
                    // Round-trip preservation: if the user has manually set a
                    // locale-specific label (e.g. "(全部)" / "(Tous)") on this
                    // filter cell in a previous edit, keep it. Fall back to the
                    // English default only when the cell is missing or empty.
                    var filterAllLabel = ReadExistingStringAtOrDefault(
                        targetSheet, sheetData, anchorColIdx + 1, rowIdx, "(All)");
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, filterAllLabel));
                    sheetData.InsertAt(filterRow, fi);
                }
            }
        }

        ws.Save();
    }

    /// <summary>
    /// Render a 2-row × 2-col × 1-data matrix pivot. The cross product of
    /// hierarchical rows (multi-row layout) with hierarchical columns
    /// (multi-col layout). Verified against matrix_authored.xlsx.
    ///
    /// Layout (rows=地区,城市 cols=产品,包装 values=金额:sum):
    ///   Row 0 (caption):       [data caption] [col field caption]
    ///   Row 1 (outer col hdr):                  咖啡            奶茶
    ///   Row 2 (inner col hdr): [row field nm]   罐装  袋装  咖啡 Total  罐装  袋装  奶茶 Total  Grand Total
    ///   Row 3 onwards:
    ///     For each row outer in display order:
    ///       Outer subtotal row: [outer]   <values across all cols>
    ///       For each (existing) inner:
    ///         Leaf row:         [inner]   <values for this leaf>
    ///   Last row: [总计] <col grand totals>
    ///
    /// Cell value semantics (all reduce raw value lists, never pre-aggregated):
    ///   - (outer row sub, leaf col):    sum over (rOuter, *, cOuter, cInner)
    ///   - (outer row sub, col sub):     sum over (rOuter, *, cOuter, *)
    ///   - (outer row sub, grand col):   sum over (rOuter, *, *, *)
    ///   - (leaf row, leaf col):         sum over (rOuter, rInner, cOuter, cInner)
    ///   - (leaf row, col sub):          sum over (rOuter, rInner, cOuter, *)
    ///   - (leaf row, grand col):        sum over (rOuter, rInner, *, *)
    ///   - (grand row, leaf col):        sum over (*, *, cOuter, cInner)
    ///   - (grand row, col sub):         sum over (*, *, cOuter, *)
    ///   - (grand row, grand col):       sum over (*, *, *, *)
    ///
    /// K=1 only. 2×2×K (matrix + multi-data) is rare and tracked as v5.
    /// </summary>
    private static void RenderMatrixPivot(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields,
        List<int>? filterFieldIndices,
        uint?[] valueStyleIds)
    {
        var rowOuterIdx = rowFieldIndices[0];
        var rowInnerIdx = rowFieldIndices[1];
        var colOuterIdx = colFieldIndices[0];
        var colInnerIdx = colFieldIndices[1];
        int K = valueFields.Count;

        var rowOuterVals = columnData[rowOuterIdx];
        var rowInnerVals = columnData[rowInnerIdx];
        var colOuterVals = columnData[colOuterIdx];
        var colInnerVals = columnData[colInnerIdx];

        var rowGroups = BuildOuterInnerGroups(rowOuterIdx, rowInnerIdx, columnData);
        var colGroups = BuildOuterInnerGroups(colOuterIdx, colInnerIdx, columnData);

        // Aggregate per (rowOuter, rowInner, colOuter, colInner, dataFieldIdx).
        // 5-tuple bucket — combines the 4-tuple matrix bucket with K data fields.
        var bucket = new Dictionary<(string ro, string ri, string co, string ci, int d), List<string>>();
        var perDataField = new List<List<string>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<string>());

        for (int i = 0; i < rowOuterVals.Length; i++)
        {
            var ro = rowOuterVals.Length > i ? rowOuterVals[i] : null;
            var ri = rowInnerVals.Length > i ? rowInnerVals[i] : null;
            var co = colOuterVals.Length > i ? colOuterVals[i] : null;
            var ci = colInnerVals.Length > i ? colInnerVals[i] : null;
            if (string.IsNullOrEmpty(ro) || string.IsNullOrEmpty(ri)
                || string.IsNullOrEmpty(co) || string.IsNullOrEmpty(ci)) continue;

            for (int d = 0; d < K; d++)
            {
                var dataIdx = valueFields[d].idx;
                var dataValues = columnData[dataIdx];
                if (i >= dataValues.Length) continue;
                var raw = dataValues[i];

                var key = (ro, ri, co, ci, d);
                if (!bucket.TryGetValue(key, out var list))
                {
                    list = new List<string>();
                    bucket[key] = list;
                }
                list.Add(raw);
                perDataField[d].Add(raw);
            }
        }

        double Reduce(IEnumerable<string> values, string func) => ReducePivotValues(values, func);

        // The 9 cell-value closures from the K=1 path now each take a data
        // field index d so the right aggregator is applied per cell.
        double LeafCell(string ro, string ri, string co, string ci, int d)
            => bucket.TryGetValue((ro, ri, co, ci, d), out var b) && b.Count > 0
                ? Reduce(b, valueFields[d].func) : double.NaN;

        double LeafRowColSub(string ro, string ri, string co, int d)
        {
            var all = new List<string>();
            foreach (var (oc, inners) in colGroups)
                if (oc == co)
                    foreach (var inner in inners)
                        if (bucket.TryGetValue((ro, ri, co, inner, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double LeafRowGrandTotal(string ro, string ri, int d)
        {
            var all = new List<string>();
            foreach (var (oc, inners) in colGroups)
                foreach (var inner in inners)
                    if (bucket.TryGetValue((ro, ri, oc, inner, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterRowLeafCell(string ro, string co, string ci, int d)
        {
            var all = new List<string>();
            foreach (var (g, inners) in rowGroups)
                if (g == ro)
                    foreach (var inner in inners)
                        if (bucket.TryGetValue((ro, inner, co, ci, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterRowColSub(string ro, string co, int d)
        {
            var all = new List<string>();
            foreach (var (g, rinners) in rowGroups)
                if (g == ro)
                    foreach (var rinner in rinners)
                        foreach (var (oc, cinners) in colGroups)
                            if (oc == co)
                                foreach (var cinner in cinners)
                                    if (bucket.TryGetValue((ro, rinner, co, cinner, d), out var b))
                                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterRowGrandTotal(string ro, int d)
        {
            var all = new List<string>();
            foreach (var (g, rinners) in rowGroups)
                if (g == ro)
                    foreach (var rinner in rinners)
                        foreach (var (oc, cinners) in colGroups)
                            foreach (var cinner in cinners)
                                if (bucket.TryGetValue((ro, rinner, oc, cinner, d), out var b))
                                    all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double GrandRowLeafCol(string co, string ci, int d)
        {
            var all = new List<string>();
            foreach (var (g, rinners) in rowGroups)
                foreach (var rinner in rinners)
                    if (bucket.TryGetValue((g, rinner, co, ci, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double GrandRowColSub(string co, int d)
        {
            var all = new List<string>();
            foreach (var (g, rinners) in rowGroups)
                foreach (var rinner in rinners)
                    foreach (var (oc, cinners) in colGroups)
                        if (oc == co)
                            foreach (var cinner in cinners)
                                if (bucket.TryGetValue((g, rinner, co, cinner, d), out var b))
                                    all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        // ===== Write cells =====
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalLabel = ActiveGrandTotalCaption;

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // CONSISTENCY(grand-totals): cache the grand totals toggles once per
        // render call. emitRowGrand = right column block; emitColGrand = bottom row.
        bool emitRowGrand = ActiveRowGrandTotals;
        bool emitColGrand = ActiveColGrandTotals;

        // CONSISTENCY(subtotals-opts): cached once per render call. When off,
        // skip per-group outer subtotal row and column position allocation,
        // header labels, and cell writes in all 9 intersections below.
        bool emitSubtotals = ActiveDefaultSubtotal;

        // Pre-compute K-aware col positions: each (outer, inner) leaf gets K
        // cells, each outer subtotal gets K cells, K final grand total cells.
        // Grand total column block is skipped entirely when emitRowGrand=false.
        var leafColPositions = new Dictionary<(string outer, string inner, int d), int>();
        var subtotalColPositions = new Dictionary<(string outer, int d), int>();
        var grandTotalColPositions = new int[K];
        int currentCol = anchorColIdx + 1;
        foreach (var (outer, inners) in colGroups)
        {
            foreach (var inner in inners)
            {
                for (int d = 0; d < K; d++)
                {
                    leafColPositions[(outer, inner, d)] = currentCol;
                    currentCol++;
                }
            }
            if (emitSubtotals)
            {
                for (int d = 0; d < K; d++)
                {
                    subtotalColPositions[(outer, d)] = currentCol;
                    currentCol++;
                }
            }
        }
        if (emitRowGrand)
        {
            for (int d = 0; d < K; d++)
            {
                grandTotalColPositions[d] = currentCol;
                currentCol++;
            }
        }

        // ----- Header rows -----
        // K=1 → 3 header rows (caption + outer col + inner col)
        // K>1 → 4 header rows (caption + outer col + inner col + data field name)
        if (K == 1)
        {
            // Row 0: data caption + col field caption.
            var captionRow = new Row { RowIndex = (uint)anchorRow };
            captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, valueFields[0].name));
            captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, headers[colOuterIdx]));
            sheetData.AppendChild(captionRow);

            // Row 1: outer col labels at first leaf col of each group.
            var outerHdrRowIdx = anchorRow + 1;
            var outerHdrRow = new Row { RowIndex = (uint)outerHdrRowIdx };
            foreach (var (outer, inners) in colGroups)
            {
                int firstLeafCol = leafColPositions[(outer, inners[0], 0)];
                outerHdrRow.AppendChild(MakeStringCell(firstLeafCol, outerHdrRowIdx, outer));
            }
            sheetData.AppendChild(outerHdrRow);

            // Row 2: row outer field name + inner col labels + "<outer> Total" + 总计.
            var innerHdrRowIdx = anchorRow + 2;
            var innerHdrRow = new Row { RowIndex = (uint)innerHdrRowIdx };
            innerHdrRow.AppendChild(MakeStringCell(anchorColIdx, innerHdrRowIdx, headers[rowOuterIdx]));
            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                    innerHdrRow.AppendChild(MakeStringCell(leafColPositions[(outer, inner, 0)],
                        innerHdrRowIdx, inner));
                if (emitSubtotals)
                    innerHdrRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, 0)], innerHdrRowIdx, outer + " Total"));
            }
            if (emitRowGrand)
                innerHdrRow.AppendChild(MakeStringCell(grandTotalColPositions[0], innerHdrRowIdx, totalLabel));
            sheetData.AppendChild(innerHdrRow);
        }
        else
        {
            // Row 0 (caption): only the col field caption (no data caption when K>1).
            var captionRow = new Row { RowIndex = (uint)anchorRow };
            captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, headers[colOuterIdx]));
            sheetData.AppendChild(captionRow);

            // Row 1 (outer col): outer label at first leaf col + per-subtotal labels
            // "<outer> <data field>" + "Total <data field>" at grand total cols.
            var outerHdrRowIdx = anchorRow + 1;
            var outerHdrRow = new Row { RowIndex = (uint)outerHdrRowIdx };
            foreach (var (outer, inners) in colGroups)
            {
                int firstLeafCol = leafColPositions[(outer, inners[0], 0)];
                outerHdrRow.AppendChild(MakeStringCell(firstLeafCol, outerHdrRowIdx, outer));
                if (emitSubtotals)
                {
                    for (int d = 0; d < K; d++)
                        outerHdrRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, d)],
                            outerHdrRowIdx, $"{outer} {valueFields[d].name}"));
                }
            }
            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    outerHdrRow.AppendChild(MakeStringCell(grandTotalColPositions[d],
                        outerHdrRowIdx, $"Total {valueFields[d].name}"));
            }
            sheetData.AppendChild(outerHdrRow);

            // Row 2 (inner col): inner label at the first data col of each (outer, inner) sub-group.
            var innerHdrRowIdx = anchorRow + 2;
            var innerHdrRow = new Row { RowIndex = (uint)innerHdrRowIdx };
            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                    innerHdrRow.AppendChild(MakeStringCell(leafColPositions[(outer, inner, 0)],
                        innerHdrRowIdx, inner));
            }
            sheetData.AppendChild(innerHdrRow);

            // Row 3 (data field name): row outer field name + data field name at every leaf col.
            var dfNameRowIdx = anchorRow + 3;
            var dfNameRow = new Row { RowIndex = (uint)dfNameRowIdx };
            dfNameRow.AppendChild(MakeStringCell(anchorColIdx, dfNameRowIdx, headers[rowOuterIdx]));
            foreach (var (outer, inners) in colGroups)
            {
                foreach (var inner in inners)
                    for (int d = 0; d < K; d++)
                        dfNameRow.AppendChild(MakeStringCell(leafColPositions[(outer, inner, d)],
                            dfNameRowIdx, valueFields[d].name));
            }
            sheetData.AppendChild(dfNameRow);
        }

        // ----- Data rows: alternate (outer subtotal row + leaf rows) per row group -----
        int firstDataRow = anchorRow + (K == 1 ? 3 : 4);
        int currentRowIdx = firstDataRow;
        foreach (var (rowOuter, rowInners) in rowGroups)
        {
            if (emitSubtotals)
            {
                // Outer subtotal row.
                var outerSubRow = new Row { RowIndex = (uint)currentRowIdx };
                outerSubRow.AppendChild(MakeStringCell(anchorColIdx, currentRowIdx, rowOuter));
                foreach (var (colOuter, colInners) in colGroups)
                {
                    foreach (var colInner in colInners)
                    {
                        bool any = HasAnyValueInOuterRowCol(rowOuter, colOuter, colInner, rowGroups, bucket, K);
                        for (int d = 0; d < K; d++)
                        {
                            var v = OuterRowLeafCell(rowOuter, colOuter, colInner, d);
                            if (v != 0 || any)
                                outerSubRow.AppendChild(MakeNumericCell(leafColPositions[(colOuter, colInner, d)], currentRowIdx, v, valueStyleIds[d]));
                        }
                    }
                    bool anyOuter = HasAnyValueInOuterRowOuterCol(rowOuter, colOuter, rowGroups, colGroups, bucket, K);
                    for (int d = 0; d < K; d++)
                    {
                        var sub = OuterRowColSub(rowOuter, colOuter, d);
                        if (sub != 0 || anyOuter)
                            outerSubRow.AppendChild(MakeNumericCell(subtotalColPositions[(colOuter, d)], currentRowIdx, sub, valueStyleIds[d]));
                    }
                }
                if (emitRowGrand)
                {
                    for (int d = 0; d < K; d++)
                        outerSubRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], currentRowIdx, OuterRowGrandTotal(rowOuter, d), valueStyleIds[d]));
                }
                sheetData.AppendChild(outerSubRow);
                currentRowIdx++;
            }

            // Leaf rows for each existing inner of this row outer.
            // When subtotals are off, prefix the first leaf with the outer label
            // so users can still identify which group the row belongs to.
            bool firstLeafOfGroup = true;
            foreach (var rowInner in rowInners)
            {
                var leafRow = new Row { RowIndex = (uint)currentRowIdx };
                var label = (!emitSubtotals && firstLeafOfGroup)
                    ? $"{rowOuter} / {rowInner}"
                    : rowInner;
                leafRow.AppendChild(MakeStringCell(anchorColIdx, currentRowIdx, label));
                firstLeafOfGroup = false;
                foreach (var (colOuter, colInners) in colGroups)
                {
                    foreach (var colInner in colInners)
                    {
                        for (int d = 0; d < K; d++)
                        {
                            var v = LeafCell(rowOuter, rowInner, colOuter, colInner, d);
                            if (!double.IsNaN(v))
                                leafRow.AppendChild(MakeNumericCell(leafColPositions[(colOuter, colInner, d)], currentRowIdx, v, valueStyleIds[d]));
                        }
                    }
                    if (emitSubtotals)
                    {
                        bool any = HasAnyValueInLeafRowCol(rowOuter, rowInner, colOuter, colGroups, bucket, K);
                        for (int d = 0; d < K; d++)
                        {
                            var sub = LeafRowColSub(rowOuter, rowInner, colOuter, d);
                            if (sub != 0 || any)
                                leafRow.AppendChild(MakeNumericCell(subtotalColPositions[(colOuter, d)], currentRowIdx, sub, valueStyleIds[d]));
                        }
                    }
                }
                if (emitRowGrand)
                {
                    for (int d = 0; d < K; d++)
                        leafRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], currentRowIdx, LeafRowGrandTotal(rowOuter, rowInner, d), valueStyleIds[d]));
                }
                sheetData.AppendChild(leafRow);
                currentRowIdx++;
            }
        }

        // Grand total row.
        if (emitColGrand)
        {
            var grandRow = new Row { RowIndex = (uint)currentRowIdx };
            grandRow.AppendChild(MakeStringCell(anchorColIdx, currentRowIdx, totalLabel));
            foreach (var (colOuter, colInners) in colGroups)
            {
                foreach (var colInner in colInners)
                    for (int d = 0; d < K; d++)
                        grandRow.AppendChild(MakeNumericCell(leafColPositions[(colOuter, colInner, d)], currentRowIdx,
                            GrandRowLeafCol(colOuter, colInner, d), valueStyleIds[d]));
                if (emitSubtotals)
                {
                    for (int d = 0; d < K; d++)
                        grandRow.AppendChild(MakeNumericCell(subtotalColPositions[(colOuter, d)], currentRowIdx, GrandRowColSub(colOuter, d), valueStyleIds[d]));
                }
            }
            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], currentRowIdx,
                        Reduce(perDataField[d], valueFields[d].func), valueStyleIds[d]));
            }
            sheetData.AppendChild(grandRow);
        }

        // Page filter cells (same logic as the other renderers).
        if (filterFieldIndices != null && filterFieldIndices.Count > 0)
        {
            var requiredHeadroom = filterFieldIndices.Count + 1;
            if (anchorRow > requiredHeadroom)
            {
                var firstFilterRow = anchorRow - requiredHeadroom;
                for (int fi = 0; fi < filterFieldIndices.Count; fi++)
                {
                    var fIdx = filterFieldIndices[fi];
                    if (fIdx < 0 || fIdx >= headers.Length) continue;
                    var rowIdx = firstFilterRow + fi;
                    var filterRow = new Row { RowIndex = (uint)rowIdx };
                    filterRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, headers[fIdx]));
                    // Round-trip preservation: if the user has manually set a
                    // locale-specific label (e.g. "(全部)" / "(Tous)") on this
                    // filter cell in a previous edit, keep it. Fall back to the
                    // English default only when the cell is missing or empty.
                    var filterAllLabel = ReadExistingStringAtOrDefault(
                        targetSheet, sheetData, anchorColIdx + 1, rowIdx, "(All)");
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, filterAllLabel));
                    sheetData.InsertAt(filterRow, fi);
                }
            }
        }

        ws.Save();
    }

    // ==================== General Tree-Based Renderer (N≥3 axis fields) ====================

    /// <summary>
    /// Render a pivot with arbitrary depth on either axis using AxisTree
    /// abstraction. Currently engaged for N_row≥3 OR N_col≥3 (the cases that
    /// the specialized RenderMultiRow/Col/Matrix renderers do not handle).
    ///
    /// Layout strategy:
    ///   - Compact mode: row labels collapse into a single column (col A)
    ///                   regardless of N_row. firstDataCol = 1.
    ///   - Each internal row tree node emits an outer-subtotal row before its
    ///     children. Each leaf tree node emits a leaf row.
    ///   - Each internal col tree node emits an outer-subtotal col AFTER its
    ///     children (matching multi-col convention). Each leaf node emits a
    ///     leaf data col.
    ///   - K data fields multiply the col area by K (K cells per leaf, K cells
    ///     per col subtotal, K final grand totals).
    ///   - Header rows: 1 caption + N_col rows (one per col field level) +
    ///                  optional 1 data field name row (when K>1) = 1 + N_col + (K>1?1:0)
    ///
    /// Cell value semantics: for each (row pos, col pos, dataField d), reduce
    /// raw values from rows whose row-field tuple matches BOTH the row path
    /// prefix AND the col path prefix. Subtotal positions widen the prefix
    /// match (e.g. an outer-row subtotal at depth 1 in a depth-3 row tree
    /// matches all source rows whose first-field value equals the path[0]).
    /// </summary>
    private static void RenderGeneralPivot(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields,
        List<int>? filterFieldIndices,
        uint?[] valueStyleIds)
    {
        int K = Math.Max(1, valueFields.Count);
        var rowTree = BuildAxisTree(rowFieldIndices, columnData);
        var colTree = BuildAxisTree(colFieldIndices, columnData);

        // Walk both trees in display order. Each entry is the absolute display
        // position relative to the start of the data area.
        // CONSISTENCY(subtotals-opts): when off, drop all subtotal positions
        // (internal tree nodes) from both axes. Leaf positions keep their
        // relative ordering, and the grand total column block is still
        // controlled separately by ActiveRow/ColGrandTotals below.
        //
        // Exception: compact mode keeps row-axis internal nodes as label-only
        // rows even when subtotals are off. Excel's compact layout displays
        // parent group headers (e.g. product name) as separate indented rows
        // without aggregated values, so users can see the hierarchy.
        bool emitSubtotals = ActiveDefaultSubtotal;
        bool compactLabelRows = !emitSubtotals && ActiveLayoutMode == "compact"
            && rowFieldIndices.Count >= 2;
        var rowPositions = WalkAxisTree(rowTree, isCol: false)
            .Where(p => emitSubtotals || !p.isSubtotal || compactLabelRows).ToList();
        var colPositions = WalkAxisTree(colTree, isCol: true)
            .Where(p => emitSubtotals || !p.isSubtotal).ToList();

        // Build per-source-row tuples once so cell value lookups are O(rows × K)
        // instead of O(rows × cells × N).
        int srcRowCount = columnData.Count > 0 ? columnData[0].Length : 0;
        var rowFieldVals = new string[srcRowCount][];
        var colFieldVals = new string[srcRowCount][];
        for (int r = 0; r < srcRowCount; r++)
        {
            rowFieldVals[r] = new string[rowFieldIndices.Count];
            colFieldVals[r] = new string[colFieldIndices.Count];
            for (int l = 0; l < rowFieldIndices.Count; l++)
            {
                var fi = rowFieldIndices[l];
                rowFieldVals[r][l] = (fi >= 0 && fi < columnData.Count && r < columnData[fi].Length)
                    ? columnData[fi][r] : null!;
            }
            for (int l = 0; l < colFieldIndices.Count; l++)
            {
                var fi = colFieldIndices[l];
                colFieldVals[r][l] = (fi >= 0 && fi < columnData.Count && r < columnData[fi].Length)
                    ? columnData[fi][r] : null!;
            }
        }

        var dataValues = valueFields.Select(field => columnData[field.idx]).ToArray();

        double Reduce(IEnumerable<string> values, string func) => ReducePivotValues(values, func);

        // Compute the value at (rowNode, colNode, dataFieldIdx).
        // Subtotal nodes have shorter Path arrays than leaves; the prefix match
        // automatically widens the set of source rows that contribute.
        double ComputeCell(AxisNode rowNode, AxisNode colNode, int d)
        {
            var rPath = rowNode.Path;
            var cPath = colNode.Path;
            var collected = new List<string>();
            for (int r = 0; r < srcRowCount; r++)
            {
                bool match = true;
                for (int l = 0; l < rPath.Length && match; l++)
                    if (rowFieldVals[r][l] != rPath[l]) match = false;
                for (int l = 0; l < cPath.Length && match; l++)
                    if (colFieldVals[r][l] != cPath[l]) match = false;
                if (!match) continue;

                // Skip rows where ANY row-axis or col-axis field is empty (mirrors
                // the specialized renderers' validity gate).
                for (int l = 0; l < rowFieldIndices.Count && match; l++)
                    if (string.IsNullOrEmpty(rowFieldVals[r][l])) match = false;
                for (int l = 0; l < colFieldIndices.Count && match; l++)
                    if (string.IsNullOrEmpty(colFieldVals[r][l])) match = false;
                if (!match) continue;

                if (r < dataValues[d].Length) collected.Add(dataValues[d][r]);
            }
            return Reduce(collected, valueFields[d].func);
        }

        // showDataAs post-processing for the tree-based renderer. The inline
        // 1×1×K path uses ApplyShowDataAs1x1 on a pre-computed matrix;
        // RenderGeneralPivot computes cells on demand via ComputeCell, so the
        // transformation has to wrap each per-cell value with division by the
        // relevant total. Mirrors the same set of modes (Excel's "Show Values
        // As" semantics):
        //
        //   normal             — passthrough
        //   percent_of_total   — v / grandTotal[d]
        //   percent_of_row     — v / rowTotal(rowNode, d)
        //   percent_of_col     — v / colTotal(colNode, d)
        //   running_total      — NOT YET (needs col ordering context; raw value
        //                        is returned, callers add a TODO marker)
        //
        // Recomputes totals per call for simplicity. With small pivots the
        // O(rowsXcols) per-row scan is fine; bigger pivots could memoize but
        // we keep the wrapper simple to match Excel's correctness first.
        var grandAxisNode = new AxisNode(string.Empty, 0, Array.Empty<string>());
        double TransformShowAs(double rawValue, AxisNode rowNode, AxisNode colNode, int d)
        {
            var mode = (valueFields[d].showAs ?? "").ToLowerInvariant();
            switch (mode)
            {
                case "":
                case "normal":
                    return rawValue;
                case "percent_of_total" or "percentoftotal" or "percent":
                {
                    var gt = ComputeCell(grandAxisNode, grandAxisNode, d);
                    return gt == 0 ? rawValue : rawValue / gt;
                }
                case "percent_of_row" or "percentofrow":
                {
                    var rt = ComputeCell(rowNode, grandAxisNode, d);
                    return rt == 0 ? rawValue : rawValue / rt;
                }
                case "percent_of_col" or "percent_of_column" or "percentofcol" or "percentofcolumn":
                {
                    var ct = ComputeCell(grandAxisNode, colNode, d);
                    return ct == 0 ? rawValue : rawValue / ct;
                }
                case "running_total" or "runningtotal":
                {
                    // Cumulative sum across col positions. ONLY applied when
                    // the aggregate is "sum" (the only case where running of
                    // per-cell values equals Excel's "aggregate over the
                    // cumulative set"). For var/varP/avg/count/etc., Excel
                    // re-aggregates against the cumulative leaf set, which
                    // we don't have machinery for; raw value is returned and
                    // Refresh in Excel computes the correct cumulative.
                    var func = (valueFields[d].func ?? "").ToLowerInvariant();
                    if (func != "sum" && func != "") return rawValue;
                    if (colPositions.Count == 0) return rawValue;
                    double running = 0;
                    bool found = false;
                    foreach (var (pNode, pIsLeaf, pIsSubtotal) in colPositions)
                    {
                        if (pIsSubtotal) continue;
                        running += ComputeCell(rowNode, pNode, d);
                        if (object.ReferenceEquals(pNode, colNode))
                        {
                            found = true;
                            break;
                        }
                    }
                    return found ? running : rawValue;
                }
                default:
                    return rawValue;
            }
        }

        double ComputeCellShowAs(AxisNode rowNode, AxisNode colNode, int d)
            => TransformShowAs(ComputeCell(rowNode, colNode, d), rowNode, colNode, d);

        bool HasAnyValue(AxisNode rowNode, AxisNode colNode)
        {
            var rPath = rowNode.Path;
            var cPath = colNode.Path;
            for (int r = 0; r < srcRowCount; r++)
            {
                bool match = true;
                for (int l = 0; l < rPath.Length && match; l++)
                    if (rowFieldVals[r][l] != rPath[l]) match = false;
                for (int l = 0; l < cPath.Length && match; l++)
                    if (colFieldVals[r][l] != cPath[l]) match = false;
                if (!match) continue;
                for (int d = 0; d < K; d++)
                    if (r < dataValues[d].Length && !string.IsNullOrEmpty(dataValues[d][r])) return true;
            }
            return false;
        }

        // ===== Write cells =====
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalLabel = ActiveGrandTotalCaption;

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // CONSISTENCY(grand-totals): cache the grand totals toggles once per
        // render call. emitRowGrand → right grand total column block;
        // emitColGrand → bottom grand total row.
        bool emitRowGrand = ActiveRowGrandTotals;
        bool emitColGrand = ActiveColGrandTotals;

        // Compact-form row-label indentation: for pivots with 2+ row fields,
        // Excel's canonical compact layout puts every row field into col A with
        // progressively deeper cell alignment indents (level 1 = indent 0,
        // level 2 = indent 1, ...). The indent is a cell style, not a rowItem
        // attribute — verified against Excel-authored test_encrypted.xlsx.
        // Build a cached indent→styleIndex map so the renderer resolves each
        // distinct depth to a single cellXfs entry. Lazy: only initialized
        // when rowFieldIndices.Count >= 2.
        var workbookPart = targetSheet.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
        var indentStyleByLevel = new Dictionary<int, uint>();
        ExcelStyleManager? styleManager = null;
        if (rowFieldIndices.Count >= 2 && workbookPart != null)
            styleManager = new ExcelStyleManager(workbookPart);

        uint GetIndentStyleIndex(int indentLevel)
        {
            if (indentLevel <= 0 || styleManager == null) return 0u;
            if (indentStyleByLevel.TryGetValue(indentLevel, out var cached)) return cached;
            // ApplyStyle mutates a temp cell but returns the xfIndex we need.
            var probe = new Cell();
            var styleIdx = styleManager.ApplyStyle(probe, new Dictionary<string, string>
            {
                ["alignment.horizontal"] = "left",
                ["alignment.indent"] = indentLevel.ToString(System.Globalization.CultureInfo.InvariantCulture)
            });
            indentStyleByLevel[indentLevel] = styleIdx;
            return styleIdx;
        }

        // Pre-compute absolute col indices for every col position × data field.
        // colPositions does not include the grand total column — that's tracked
        // separately so the writer doesn't accidentally include it inside the
        // per-outer subtotal block.
        int colCells = colPositions.Count * K;
        // Compact: all row fields share one column → firstDataCol = anchor + 1
        // Outline/Tabular: one column per row field → firstDataCol = anchor + N
        int rowLabelCols = ActiveLayoutMode == "compact"
            ? 1
            : Math.Max(1, rowFieldIndices.Count);
        int firstDataCol = anchorColIdx + rowLabelCols;
        var colIdxByPosition = new int[colPositions.Count, K];
        for (int p = 0; p < colPositions.Count; p++)
            for (int d = 0; d < K; d++)
                colIdxByPosition[p, d] = firstDataCol + p * K + d;
        int grandTotalColStart = firstDataCol + colCells;  // unused when !emitRowGrand

        // Header rows. Layout depends on (N_col, K):
        //   - colN == 0 && K == 1: single header row with row-label caption
        //                          + data field name.
        //   - colN == 0 && K >  1: two header rows — R0 carries the "Values"
        //                          axis caption at col B, R1 carries the
        //                          row-label caption at col A plus K data
        //                          field names across cols B..B+K-1. Excel
        //                          injects a synthetic col field (x=-2) for
        //                          multi-data no-col pivots; the rendered
        //                          sheetData must match that axis shape.
        //   - colN >= 1: 1 caption row + N_col field-label rows + optional
        //                dfRow when K>1.
        //   Must stay in sync with ComputePivotGeometry and BuildLocation.
        int headerRows;
        if (colFieldIndices.Count == 0)
            headerRows = K > 1 ? 2 : 1;
        else
            headerRows = 1 + colFieldIndices.Count + (K > 1 ? 1 : 0);

        // Helper: write row field header labels into the label columns.
        // Compact: single caption at anchorColIdx (first row field name).
        // Outline/Tabular: one header per row field, each in its own column.
        void WriteRowFieldHeaders(Row row, int rowIndex)
        {
            if (ActiveLayoutMode == "compact")
            {
                var caption = rowFieldIndices.Count > 0
                    ? headers[rowFieldIndices[0]]
                    : "Row Labels";
                row.AppendChild(MakeStringCell(anchorColIdx, rowIndex, caption));
            }
            else
            {
                for (int f = 0; f < rowFieldIndices.Count; f++)
                    row.AppendChild(MakeStringCell(anchorColIdx + f, rowIndex, headers[rowFieldIndices[f]]));
            }
        }

        if (colFieldIndices.Count == 0)
        {
            if (K > 1)
            {
                // R0: "Values" axis caption at first data col.
                var valuesCaptionRow = new Row { RowIndex = (uint)anchorRow };
                valuesCaptionRow.AppendChild(MakeStringCell(firstDataCol, anchorRow, "Values"));
                sheetData.AppendChild(valuesCaptionRow);

                // R1: row-label caption(s), K data field names.
                int dfHeaderRowIdx = anchorRow + 1;
                var dfHeaderRow = new Row { RowIndex = (uint)dfHeaderRowIdx };
                WriteRowFieldHeaders(dfHeaderRow, dfHeaderRowIdx);
                for (int d = 0; d < K; d++)
                    dfHeaderRow.AppendChild(MakeStringCell(firstDataCol + d, dfHeaderRowIdx,
                        valueFields[d].name));
                sheetData.AppendChild(dfHeaderRow);
            }
            else
            {
                // Single header row: row-label caption(s), single data field name.
                var headerRow = new Row { RowIndex = (uint)anchorRow };
                WriteRowFieldHeaders(headerRow, anchorRow);
                headerRow.AppendChild(MakeStringCell(firstDataCol, anchorRow, valueFields[0].name));
                sheetData.AppendChild(headerRow);
            }
        }
        else
        {
            // Row 0 (caption): col field caption (the outermost col field name) at
            // first data col position. For K=1 the row-label col also gets the
            // single data field name.
            var captionRow = new Row { RowIndex = (uint)anchorRow };
            if (K == 1)
                captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, valueFields[0].name));
            captionRow.AppendChild(MakeStringCell(firstDataCol, anchorRow,
                headers[colFieldIndices[0]]));
            sheetData.AppendChild(captionRow);
        }

        // Rows 1..N_col (col field header rows). For each level L (1..N_col), the
        // L-th col field's labels are written at the first leaf col of every node
        // at depth L in the col tree. Subtotal cols at level L get their label
        // here too (for the outermost level when K>1, we put the subtotal labels
        // in the outermost header row, matching the multi-col K>1 ground truth).
        for (int level = 1; level <= colFieldIndices.Count; level++)
        {
            int headerRowIdx = anchorRow + level;
            var headerRow = new Row { RowIndex = (uint)headerRowIdx };
            // Row label column header on the LAST col-field row carries the
            // row field name(s) (when K=1) or stays empty (when K>1
            // because the data-field-name row below carries it).
            if (level == colFieldIndices.Count && K == 1 && rowFieldIndices.Count > 0)
                WriteRowFieldHeaders(headerRow, headerRowIdx);

            for (int p = 0; p < colPositions.Count; p++)
            {
                var (node, isLeaf, isSubtotal) = colPositions[p];
                // Internal-node label appears at THIS row only when level matches
                // the node's depth, AND it appears at the FIRST data col of its
                // descendants (i.e. the position of the first leaf in its subtree).
                if (isSubtotal)
                {
                    // For each internal node N at depth L, the subtotal label
                    // pattern depends on which row we're on:
                    //   - At header row L (matching the node's depth): emit the
                    //     parent-style label "<parent path tail>" at the first
                    //     leaf col of N's subtree.
                    //   - At the LAST col-field header row (level == N_col): emit
                    //     the "<node label> Total" at THIS subtotal col position.
                    if (level == node.Depth)
                    {
                        // Subtotal cols don't carry inner labels; the label here
                        // is the node's own label, written at THIS subtotal col.
                        // Match the multi-col single-data convention: "<outer> Total".
                        if (K == 1)
                            headerRow.AppendChild(MakeStringCell(colIdxByPosition[p, 0], headerRowIdx,
                                node.Label + " Total"));
                        else
                        {
                            // Multi-data: emit per-data-field labels.
                            for (int d = 0; d < K; d++)
                                headerRow.AppendChild(MakeStringCell(colIdxByPosition[p, d], headerRowIdx,
                                    $"{node.Label} {valueFields[d].name}"));
                        }
                    }
                    continue;
                }

                // Leaf node: emit the label corresponding to THIS header level.
                // Only at the level where the node's path-element matches (depth).
                if (level <= node.Path.Length)
                {
                    // Write at the FIRST leaf of any contiguous group sharing the
                    // same prefix at this level. Approximation: write at every
                    // leaf, but Excel deduplicates visually via colItems metadata.
                    // Simpler implementation: just write the label at this leaf
                    // for the level matching its current depth in the tree.
                    if (level == node.Path.Length)
                    {
                        // Innermost level for this leaf: emit at first data col.
                        headerRow.AppendChild(MakeStringCell(colIdxByPosition[p, 0], headerRowIdx, node.Label));
                    }
                    else
                    {
                        // Outer ancestor levels: emit the ancestor label only at
                        // the first leaf of the ancestor's subtree (positions
                        // sharing path[level-1] = ancestor's label, AND this is
                        // the first such position).
                        // Find the previous position; if its path[level-1] differs
                        // OR there is no previous, this is the start of a new group.
                        bool isFirst = (p == 0);
                        if (!isFirst)
                        {
                            var (prevNode, _, prevIsSub) = colPositions[p - 1];
                            // Skip subtotal cols when checking "previous leaf in group"
                            // — subtotals belong to a different ancestor than their
                            // following leaves.
                            if (prevIsSub) isFirst = true;
                            else
                            {
                                var prev = prevNode;
                                if (level - 1 >= prev.Path.Length || level - 1 >= node.Path.Length
                                    || prev.Path[level - 1] != node.Path[level - 1])
                                    isFirst = true;
                            }
                        }
                        if (isFirst && level - 1 < node.Path.Length)
                            headerRow.AppendChild(MakeStringCell(colIdxByPosition[p, 0], headerRowIdx,
                                node.Path[level - 1]));
                    }
                }
            }

            // Grand total column header label appears at the LAST col header row
            // for K=1. For K>1 the label belongs on the data-field-name row
            // below (alongside "Sum of Sales"/"Sum of Qty"), not on the col
            // header row — see the K>1 block right after this loop.
            if (level == colFieldIndices.Count && emitRowGrand && K == 1)
            {
                headerRow.AppendChild(MakeStringCell(grandTotalColStart, headerRowIdx, totalLabel));
            }
            sheetData.AppendChild(headerRow);
        }

        // Optional data field name row (K>1). Only emitted when colN >= 1;
        // the colN == 0 path above already wrote a single combined header row
        // carrying the row-label caption + data field names, so running this
        // block would write duplicate cells at anchorRow.
        if (K > 1 && colFieldIndices.Count > 0)
        {
            int dfRowIdx = anchorRow + headerRows - 1;
            var dfRow = new Row { RowIndex = (uint)dfRowIdx };
            if (rowFieldIndices.Count > 0)
                WriteRowFieldHeaders(dfRow, dfRowIdx);
            for (int p = 0; p < colPositions.Count; p++)
            {
                var (_, isLeaf, isSubtotal) = colPositions[p];
                if (isSubtotal) continue; // Subtotal cols already labelled in their header row above.
                for (int d = 0; d < K; d++)
                    dfRow.AppendChild(MakeStringCell(colIdxByPosition[p, d], dfRowIdx, valueFields[d].name));
            }
            // K>1 grand total column captions ("Total Sum of Sales" /
            // "Total Sum of Qty") sit on the data-field-name row, NOT on the
            // col-header row above — that row carries the col-axis labels
            // (Q1/Q2/...) and would visually misalign the grand total caption
            // with its values otherwise.
            if (emitRowGrand)
            {
                for (int d = 0; d < K; d++)
                    dfRow.AppendChild(MakeStringCell(grandTotalColStart + d, dfRowIdx,
                        $"Total {valueFields[d].name}"));
            }
            sheetData.AppendChild(dfRow);
        }

        // Data + grand total rows.
        int firstDataRowIdx = anchorRow + headerRows;
        int blankRowOffset = 0; // extra rows inserted for insertBlankRow
        for (int rp = 0; rp < rowPositions.Count; rp++)
        {
            var (rowNode, rIsLeaf, rIsSubtotal) = rowPositions[rp];
            int rowIdx = firstDataRowIdx + rp + blankRowOffset;
            var row = new Row { RowIndex = (uint)rowIdx };
            if (ActiveLayoutMode == "compact")
            {
                // Compact-mode: all labels in one column with indentation.
                // level 1 (outermost row field) gets no indent (style 0),
                // level 2 gets indent 1, level 3 gets indent 2, etc.
                var rowLabelCell = MakeStringCell(anchorColIdx, rowIdx, rowNode.Label);
                var indentStyle = GetIndentStyleIndex(rowNode.Depth - 1);
                if (indentStyle != 0) rowLabelCell.StyleIndex = indentStyle;
                row.AppendChild(rowLabelCell);
            }
            else
            {
                // Outline/Tabular: each row field level writes to its own column.
                // rowNode.Depth is 1-based; the label goes at column (anchor + depth - 1).
                // Tabular subtotal rows append " Total" to match Excel — the
                // subtotal row sits AFTER its leaves so the suffix disambiguates
                // it from a leaf row of the same name. Outline subtotals sit
                // BEFORE leaves and act as group headers, so they keep the
                // bare label (matches Excel's outline mode).
                // CONSISTENCY(subtotal-total-suffix): mirrors col-axis subtotal
                // labels at PivotTableHelper.Render.cs:1981.
                int labelCol = anchorColIdx + rowNode.Depth - 1;
                string labelText = rowNode.Label;
                if (rIsSubtotal && ActiveLayoutMode == "tabular")
                    labelText = rowNode.Label + " Total";
                row.AppendChild(MakeStringCell(labelCol, rowIdx, labelText));
                // Ancestor labels for non-compact leaf rows. Two modes:
                //   - repeatLabels=true: write every ancestor on every leaf,
                //     unconditionally (Excel's "Repeat All Item Labels" toggle).
                //   - default: per-level diff against the previous row's path.
                //     A given ancestor level is written only if its value
                //     changed from the previous row. The previous row may be
                //     a subtotal (path shorter than leaf) or another leaf —
                //     either way the diff gives the correct answer:
                //       * outline+subtotals=on: prev subtotal already carries
                //         the outer label, so its path matches → diff skips it
                //       * outline+subtotals=off: parent labels appear on first
                //         leaf of each group; intermediate transitions stay
                //         visible
                //       * tabular+subtotals=on: after an inner subtotal at
                //         depth L, the next leaf only re-writes ancestors that
                //         actually changed (NOT the still-same outer ones)
                //       * tabular+subtotals=off: same as outline+subtotals=off
                // CONSISTENCY(first-of-group-ancestors): one rule for every
                // non-compact leaf — per-level diff is what Excel does.
                if (rowNode.Depth >= 2)
                {
                    bool repeatAll = ActiveRepeatItemLabels;
                    if (repeatAll)
                    {
                        for (int anc = 0; anc < rowNode.Depth - 1; anc++)
                            row.InsertBefore(
                                MakeStringCell(anchorColIdx + anc, rowIdx, rowNode.Path[anc]),
                                row.FirstChild);
                    }
                    else if (rIsLeaf)
                    {
                        string[]? prevPath = null;
                        if (rp > 0)
                        {
                            var (prevNode, _, _) = rowPositions[rp - 1];
                            prevPath = prevNode.Path;
                        }
                        for (int anc = 0; anc < rowNode.Depth - 1; anc++)
                        {
                            bool changed = prevPath == null
                                || anc >= prevPath.Length
                                || prevPath[anc] != rowNode.Path[anc];
                            if (changed)
                                row.InsertBefore(
                                    MakeStringCell(anchorColIdx + anc, rowIdx, rowNode.Path[anc]),
                                    row.FirstChild);
                        }
                    }
                }
            }

            // Label-only rows: compact internal nodes with subtotals off
            // get the label but no aggregated values (mirrors Excel's compact
            // layout where parent group headers have no data).
            bool isLabelOnly = compactLabelRows && rIsSubtotal && !emitSubtotals;

            if (!isLabelOnly)
            {
                if (colPositions.Count > 0)
                {
                    for (int cp = 0; cp < colPositions.Count; cp++)
                    {
                        var (colNode, cIsLeaf, cIsSubtotal) = colPositions[cp];
                        bool any = HasAnyValue(rowNode, colNode);
                        for (int d = 0; d < K; d++)
                        {
                            var v = ComputeCellShowAs(rowNode, colNode, d);
                            // Skip 0-value cells when there are no underlying values to
                            // mirror Excel's behavior of leaving sparse intersections blank.
                            if (any || v != 0)
                                row.AppendChild(MakeNumericCell(colIdxByPosition[cp, d], rowIdx, v, valueStyleIds[d]));
                        }
                    }
                }
                else
                {
                    // No col fields: K value cells written directly. The empty
                    // colNode matches all source rows so ComputeCell aggregates
                    // across the entire dataset for the given row path.
                    var emptyColNode = new AxisNode(string.Empty, 0, Array.Empty<string>());
                    for (int d = 0; d < K; d++)
                    {
                        var v = ComputeCellShowAs(rowNode, emptyColNode, d);
                        row.AppendChild(MakeNumericCell(firstDataCol + d, rowIdx, v, valueStyleIds[d]));
                    }
                }
            }

            // Grand total cells (per data field) — the row's value across all cols.
            // Only applies when there ARE col fields; without col fields the value
            // cells already aggregate across all rows (no per-row grand total needed).
            if (emitRowGrand && !isLabelOnly && colPositions.Count > 0)
            {
                var grandRowNode = new AxisNode(string.Empty, 0, Array.Empty<string>());
                for (int d = 0; d < K; d++)
                    row.AppendChild(MakeNumericCell(grandTotalColStart + d, rowIdx,
                        ComputeCellShowAs(rowNode, grandRowNode, d), valueStyleIds[d]));
            }
            sheetData.AppendChild(row);

            // insertBlankRow: insert an empty row after each outer group's
            // last entry. With subtotals ON, that's the depth-1 subtotal row;
            // with subtotals OFF those positions are filtered out, so we
            // detect end-of-group as "the next row's outermost path element
            // differs from this row's, OR there is no next row before the
            // grand total". This works for tabular/outline/compact alike.
            bool insertBlank = false;
            if (ActiveInsertBlankRow)
            {
                if (rIsSubtotal && rowNode.Depth == 1)
                    insertBlank = true;
                else if (!emitSubtotals && rIsLeaf && rowNode.Path.Length >= 1)
                {
                    bool isLastInGroup;
                    if (rp == rowPositions.Count - 1)
                        isLastInGroup = true;
                    else
                    {
                        var (nextNode, _, _) = rowPositions[rp + 1];
                        isLastInGroup = nextNode.Path.Length == 0
                            || nextNode.Path[0] != rowNode.Path[0];
                    }
                    insertBlank = isLastInGroup;
                }
            }
            if (insertBlank)
            {
                blankRowOffset++;
                var blankRow = new Row { RowIndex = (uint)(rowIdx + 1) };
                sheetData.AppendChild(blankRow);
            }
        }

        // Final grand total row.
        if (emitColGrand)
        {
            int grandRowIdx = firstDataRowIdx + rowPositions.Count + blankRowOffset;
            var grandRow = new Row { RowIndex = (uint)grandRowIdx };
            grandRow.AppendChild(MakeStringCell(anchorColIdx, grandRowIdx, totalLabel));
            var grandRowNodeFinal = new AxisNode(string.Empty, 0, Array.Empty<string>());
            if (colPositions.Count > 0)
            {
                for (int cp = 0; cp < colPositions.Count; cp++)
                {
                    var (colNode, _, _) = colPositions[cp];
                    for (int d = 0; d < K; d++)
                    {
                        var v = ComputeCellShowAs(grandRowNodeFinal, colNode, d);
                        grandRow.AppendChild(MakeNumericCell(colIdxByPosition[cp, d], grandRowIdx, v, valueStyleIds[d]));
                    }
                }
            }
            else
            {
                // No col fields: write K value cells directly at firstDataCol.
                var emptyColNode = new AxisNode(string.Empty, 0, Array.Empty<string>());
                for (int d = 0; d < K; d++)
                {
                    var v = ComputeCellShowAs(grandRowNodeFinal, emptyColNode, d);
                    grandRow.AppendChild(MakeNumericCell(firstDataCol + d, grandRowIdx, v, valueStyleIds[d]));
                }
            }
            if (emitRowGrand && colPositions.Count > 0)
            {
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(grandTotalColStart + d, grandRowIdx,
                        ComputeCellShowAs(grandRowNodeFinal, grandRowNodeFinal, d), valueStyleIds[d]));
            }
            sheetData.AppendChild(grandRow);
        }

        // Page filter cells (same logic as the other renderers).
        if (filterFieldIndices != null && filterFieldIndices.Count > 0)
        {
            var requiredHeadroom = filterFieldIndices.Count + 1;
            if (anchorRow > requiredHeadroom)
            {
                var firstFilterRow = anchorRow - requiredHeadroom;
                for (int fi = 0; fi < filterFieldIndices.Count; fi++)
                {
                    var fIdx = filterFieldIndices[fi];
                    if (fIdx < 0 || fIdx >= headers.Length) continue;
                    var rowIdx = firstFilterRow + fi;
                    var filterRow = new Row { RowIndex = (uint)rowIdx };
                    filterRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, headers[fIdx]));
                    // Round-trip preservation: if the user has manually set a
                    // locale-specific label (e.g. "(全部)" / "(Tous)") on this
                    // filter cell in a previous edit, keep it. Fall back to the
                    // English default only when the cell is missing or empty.
                    var filterAllLabel = ReadExistingStringAtOrDefault(
                        targetSheet, sheetData, anchorColIdx + 1, rowIdx, "(All)");
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, filterAllLabel));
                    sheetData.InsertAt(filterRow, fi);
                }
            }
        }

        ws.Save();
    }

    /// <summary>
    /// Helper for RenderMatrixPivot: true if (rowOuter, *, colOuter, colInner)
    /// has any non-empty leaf bucket across any data field.
    /// </summary>
    private static bool HasAnyValueInOuterRowCol(string rowOuter, string colOuter, string colInner,
        List<(string outer, List<string> inners)> rowGroups,
        Dictionary<(string ro, string ri, string co, string ci, int d), List<string>> bucket,
        int dataFieldCount)
    {
        foreach (var (g, inners) in rowGroups)
        {
            if (g != rowOuter) continue;
            foreach (var inner in inners)
                for (int d = 0; d < dataFieldCount; d++)
                    if (bucket.TryGetValue((rowOuter, inner, colOuter, colInner, d), out var b) && b.Count > 0)
                        return true;
        }
        return false;
    }

    /// <summary>
    /// Helper for RenderMatrixPivot: true if (rowOuter, *, colOuter, *) has any
    /// non-empty bucket across any data field.
    /// </summary>
    private static bool HasAnyValueInOuterRowOuterCol(string rowOuter, string colOuter,
        List<(string outer, List<string> inners)> rowGroups,
        List<(string outer, List<string> inners)> colGroups,
        Dictionary<(string ro, string ri, string co, string ci, int d), List<string>> bucket,
        int dataFieldCount)
    {
        foreach (var (g, rinners) in rowGroups)
        {
            if (g != rowOuter) continue;
            foreach (var rinner in rinners)
                foreach (var (oc, cinners) in colGroups)
                    if (oc == colOuter)
                        foreach (var cinner in cinners)
                            for (int d = 0; d < dataFieldCount; d++)
                                if (bucket.TryGetValue((rowOuter, rinner, colOuter, cinner, d), out var b) && b.Count > 0)
                                    return true;
        }
        return false;
    }

    /// <summary>
    /// Helper for RenderMatrixPivot: true if (rowOuter, rowInner, colOuter, *)
    /// has any non-empty bucket across any data field.
    /// </summary>
    private static bool HasAnyValueInLeafRowCol(string rowOuter, string rowInner, string colOuter,
        List<(string outer, List<string> inners)> colGroups,
        Dictionary<(string ro, string ri, string co, string ci, int d), List<string>> bucket,
        int dataFieldCount)
    {
        foreach (var (oc, cinners) in colGroups)
        {
            if (oc != colOuter) continue;
            foreach (var cinner in cinners)
                for (int d = 0; d < dataFieldCount; d++)
                    if (bucket.TryGetValue((rowOuter, rowInner, colOuter, cinner, d), out var b) && b.Count > 0)
                        return true;
        }
        return false;
    }

    /// <summary>
    /// Helper for RenderMultiColPivot: like HasAnyValueInOuterCol but flipped
    /// (checks if a (row, outerCol) pair has any non-empty leaf bucket across
    /// the outer's inners and any data field). Used to decide whether to
    /// write a 0-valued subtotal cell or skip it entirely on a sparse row.
    /// </summary>
    private static bool HasAnyValueInRowOuter(string row, string outerCol,
        List<(string outer, List<string> inners)> colGroups,
        Dictionary<(string r, string oc, string ic, int d), List<string>> leafBucket,
        int dataFieldCount)
    {
        foreach (var (oc, inners) in colGroups)
        {
            if (oc != outerCol) continue;
            foreach (var inner in inners)
                for (int d = 0; d < dataFieldCount; d++)
                    if (leafBucket.TryGetValue((row, outerCol, inner, d), out var b) && b.Count > 0)
                        return true;
        }
        return false;
    }

    /// <summary>
    /// Helper for the multi-row renderer: returns true if the (outer, col)
    /// pair has at least one non-empty leaf bucket across any of the K data
    /// fields. Used to decide whether to write a 0-valued subtotal cell or
    /// skip it entirely (Excel writes nothing rather than a literal 0 for
    /// genuinely empty (outer, col) intersections).
    /// </summary>
    private static bool HasAnyValueInOuterCol(string outer, string col,
        List<(string outer, List<string> inners)> groups,
        Dictionary<(string o, string i, string c, int d), List<string>> leafBucket,
        int dataFieldCount)
    {
        foreach (var (o, inners) in groups)
        {
            if (o != outer) continue;
            foreach (var inner in inners)
                for (int d = 0; d < dataFieldCount; d++)
                    if (leafBucket.TryGetValue((outer, inner, col, d), out var b) && b.Count > 0)
                        return true;
        }
        return false;
    }

    /// <summary>
    /// Build an inline-string cell. We use inline strings (t="inlineStr" + &lt;is&gt;)
    /// rather than the SharedStringTable because the renderer is self-contained
    /// and adding entries to the SST would require coordinating with whatever
    /// other handler code touches the workbook's strings — out of scope for v1.
    /// </summary>
    private static Cell MakeStringCell(int colIdx, int rowIdx, string text)
    {
        return new Cell
        {
            CellReference = $"{IndexToCol(colIdx)}{rowIdx}",
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(text ?? string.Empty))
        };
    }

    /// <summary>
    /// Read the string value of an existing cell at (colIdx, rowIdx) and
    /// return it if non-empty, otherwise return <paramref name="defaultValue"/>.
    /// Used by the page filter renderers to preserve a user-localized filter
    /// label (e.g. "(全部)") on round-trip through <c>RebuildFieldAreas</c>,
    /// instead of overwriting it with our English default "(All)".
    ///
    /// Resolves both InlineString cells and SharedString cells; falls back to
    /// the raw CellValue text if neither matches. Missing row / missing cell /
    /// empty text all return the default.
    /// </summary>
    private static string ReadExistingStringAtOrDefault(
        WorksheetPart targetSheet, SheetData sheetData,
        int colIdx, int rowIdx, string defaultValue)
    {
        var cellRef = $"{IndexToCol(colIdx)}{rowIdx}";
        var row = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx);
        if (row == null) return defaultValue;
        var cell = row.Elements<Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == cellRef);
        if (cell == null) return defaultValue;

        // InlineString: text is embedded in the cell.
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            var inline = cell.InlineString?.Text?.Text ?? cell.InlineString?.InnerText;
            if (!string.IsNullOrEmpty(inline)) return inline;
            return defaultValue;
        }

        // SharedString: CellValue holds the SST index; resolve via workbook.
        if (cell.DataType?.Value == CellValues.SharedString
            && cell.CellValue?.Text is { } sstIdxStr
            && int.TryParse(sstIdxStr, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var sstIdx))
        {
            var wbPart = targetSheet.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
            var sst = wbPart?.SharedStringTablePart?.SharedStringTable;
            if (sst != null)
            {
                var items = sst.Elements<SharedStringItem>().ToList();
                if (sstIdx >= 0 && sstIdx < items.Count)
                {
                    var txt = items[sstIdx].Text?.Text ?? items[sstIdx].InnerText;
                    if (!string.IsNullOrEmpty(txt)) return txt;
                }
            }
            return defaultValue;
        }

        // String-typed (legacy) or untyped: fall back to raw CellValue.
        if (cell.CellValue?.Text is { Length: > 0 } cv) return cv;

        return defaultValue;
    }

    /// <summary>
    /// Numeric cell with the value serialized using invariant culture.
    /// When <paramref name="styleIndex"/> is provided, the cell carries that
    /// styles.xml cellXfs index — used to inherit the source column's number
    /// format (currency, percentage, custom format) onto pivot value cells so
    /// the pivot displays "¥1,234.50" rather than the raw "1234.5".
    /// </summary>
    private static Cell MakeNumericCell(int colIdx, int rowIdx, double value, uint? styleIndex = null)
    {
        var cell = new Cell
        {
            CellReference = $"{IndexToCol(colIdx)}{rowIdx}",
            CellValue = new CellValue(value.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
        };
        if (styleIndex.HasValue)
            cell.StyleIndex = styleIndex.Value;
        return cell;
    }

    // CONSISTENCY(numfmt-warning): JSON mode queues the advisory for the
    // envelope's warnings[]; plain mode keeps the stderr line (which the
    // resident server lifts via BuildWarnings).
    private static void WarnRenderAdvisory(string message, string code)
    {
        if (WarningContext.IsActive)
            WarningContext.Add(message, code);
        else
            Console.Error.WriteLine($"WARNING: {message}");
    }
}
