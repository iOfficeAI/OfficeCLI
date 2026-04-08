// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Helper for building and reading pivot tables.
/// Manages PivotTableCacheDefinitionPart (workbook-level) and PivotTablePart (worksheet-level).
/// </summary>
internal static class PivotTableHelper
{
    // ==================== Axis sort options ====================
    //
    // Axis labels on every level are sorted through a single comparer that
    // CreatePivotTable / SetPivotTableProperties publishes into _axisSortMode
    // for the duration of the operation. Every sort site below reads
    // ActiveAxisComparer / ActiveAxisDescending rather than hard-coding
    // StringComparer.Ordinal.
    //
    // Why ThreadStatic instead of a parameter: the sort opts have to reach
    // ~15 deeply-nested call sites (cache builders, pivotField items writers,
    // per-level index maps, 5 specialized renderers). Threading a parameter
    // through all of them would balloon 15+ signatures with pass-through
    // boilerplate. The CLI is single-threaded per pivot operation, so
    // ThreadStatic is safe and dramatically less invasive.
    //
    // Supported modes:
    //   "asc"         — StringComparer.Ordinal ascending (DEFAULT, preserves
    //                   byte-level regression baselines)
    //   "desc"        — StringComparer.Ordinal descending
    //   "locale"      — CurrentCulture ascending (pinyin for zh-CN, etc.)
    //   "locale-desc" — CurrentCulture descending
    [ThreadStatic] private static string? _axisSortMode;

    private static IComparer<string> ActiveAxisComparer => _axisSortMode switch
    {
        "locale" or "locale-desc" => StringComparer.CurrentCulture,
        _ => StringComparer.Ordinal
    };

    private static bool ActiveAxisDescending => _axisSortMode switch
    {
        "desc" or "locale-desc" => true,
        _ => false
    };

    /// <summary>
    /// Set axis sort mode from the pivot properties and return a token that
    /// restores the previous value on Dispose. Usage:
    ///   using (PushAxisSortMode(properties)) { ... build pivot ... }
    /// </summary>
    private static IDisposable PushAxisSortMode(Dictionary<string, string> properties)
    {
        var prev = _axisSortMode;
        if (properties.TryGetValue("sort", out var mode) && !string.IsNullOrWhiteSpace(mode))
            _axisSortMode = mode.Trim().ToLowerInvariant();
        return new SortModeScope(prev);
    }

    private sealed class SortModeScope : IDisposable
    {
        private readonly string? _prev;
        public SortModeScope(string? prev) { _prev = prev; }
        public void Dispose() { _axisSortMode = _prev; }
    }

    /// <summary>
    /// Apply axis ordering (ascending/descending) to an OrderBy clause using
    /// the currently-active sort mode. All axis sort sites use this helper.
    /// </summary>
    private static IOrderedEnumerable<T> OrderByAxis<T>(this IEnumerable<T> source, Func<T, string> keySelector)
    {
        return ActiveAxisDescending
            ? source.OrderByDescending(keySelector, ActiveAxisComparer)
            : source.OrderBy(keySelector, ActiveAxisComparer);
    }

    /// <summary>
    /// Create a pivot table on the target worksheet.
    /// </summary>
    /// <param name="workbookPart">The workbook part</param>
    /// <param name="targetSheet">Worksheet where the pivot table will be placed</param>
    /// <param name="sourceSheet">Worksheet containing the source data</param>
    /// <param name="sourceSheetName">Name of the source worksheet</param>
    /// <param name="sourceRef">Source data range (e.g. "A1:D100")</param>
    /// <param name="position">Top-left cell for the pivot table (e.g. "F1")</param>
    /// <param name="properties">Configuration: rows, cols, values, filters, style, name</param>
    /// <returns>The 1-based index of the created pivot table</returns>
    internal static int CreatePivotTable(
        WorkbookPart workbookPart,
        WorksheetPart targetSheet,
        WorksheetPart sourceSheet,
        string sourceSheetName,
        string sourceRef,
        string position,
        Dictionary<string, string> properties)
    {
        // Publish the axis sort mode (asc/desc/locale/locale-desc) so every
        // sort site below — cache builder, pivotField items writer, per-level
        // index maps, specialized renderers — reads the same comparer.
        using var _sortScope = PushAxisSortMode(properties);

        // 1. Read source data to build cache
        var (headers, columnData, columnStyleIds) = ReadSourceData(sourceSheet, sourceRef);
        if (headers.Length == 0)
            throw new ArgumentException("Source range has no data");

        // 2. Parse field assignments from properties
        var rowFields = ParseFieldList(properties, "rows", headers);
        var colFields = ParseFieldList(properties, "cols", headers);
        var filterFields = ParseFieldList(properties, "filters", headers);
        var valueFields = ParseValueFields(properties, "values", headers);

        // Auto-assign: if no values specified, use the first numeric column
        if (valueFields.Count == 0)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                if (!rowFields.Contains(i) && !colFields.Contains(i) && !filterFields.Contains(i)
                    && columnData[i].All(v => double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _)))
                {
                    valueFields.Add((i, "sum", $"Sum of {headers[i]}"));
                    break;
                }
            }
        }

        // 3. Generate unique cache ID
        uint cacheId = 0;
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("Workbook is missing");
        var pivotCaches = workbook.GetFirstChild<PivotCaches>();
        if (pivotCaches != null)
            cacheId = pivotCaches.Elements<PivotCache>().Select(pc => pc.CacheId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;

        // 4. Create PivotTableCacheDefinitionPart at workbook level
        var cachePart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        var cacheRelId = workbookPart.GetIdOfPart(cachePart);

        // Build cache definition + per-field shared-item index maps. The maps are
        // needed to write pivotCacheRecords below: each non-numeric field value is
        // referenced as <x v="N"/> where N is the value's position in sharedItems.
        var (cacheDef, fieldNumeric, fieldValueIndex) =
            BuildCacheDefinition(sourceSheetName, sourceRef, headers, columnData);
        cachePart.PivotCacheDefinition = cacheDef;
        cachePart.PivotCacheDefinition.Save();

        // 4b. Create PivotTableCacheRecordsPart and write one record per source row.
        // Without records, Excel rejects the file with "PivotTable report is invalid"
        // because saveData defaults to true. Writing real records also makes the file
        // self-contained for non-refreshing consumers (POI, third-party parsers).
        var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        recordsPart.PivotCacheRecords = BuildCacheRecords(columnData, fieldNumeric, fieldValueIndex);
        recordsPart.PivotCacheRecords.Save();

        // The pivotCacheDefinition element MUST carry an r:id attribute pointing to the
        // records part — Excel uses it to find records, not the package _rels alone.
        // LibreOffice writes this in xepivotxml.cxx:280 (FSNS(XML_r, XML_id)). Without
        // this attribute the file looks structurally complete but Excel rejects it.
        cacheDef.Id = cachePart.GetIdOfPart(recordsPart);
        cachePart.PivotCacheDefinition.Save();

        // Register in workbook's PivotCaches
        if (pivotCaches == null)
        {
            pivotCaches = new PivotCaches();
            workbook.AppendChild(pivotCaches);
        }
        pivotCaches.AppendChild(new PivotCache { CacheId = cacheId, Id = cacheRelId });
        workbook.Save();

        // 5. Create PivotTablePart at worksheet level
        var pivotPart = targetSheet.AddNewPart<PivotTablePart>();
        // Link pivot table to cache definition
        pivotPart.AddPart(cachePart);

        var pivotName = properties.GetValueOrDefault("name", $"PivotTable{cacheId + 1}");
        var style = properties.GetValueOrDefault("style", "PivotStyleLight16");

        // Resolve per-column numFmtId from the source StyleIndex so we can stamp
        // it onto DataField elements below. Excel uses DataField.NumberFormatId
        // as the PRIMARY display driver for pivot values — the cell-level
        // StyleIndex alone is not enough; without this, Excel renders pivot
        // values as plain General-format numbers even though the rendered cells
        // carry the correct style.
        var columnNumFmtIds = ResolveColumnNumFmtIds(workbookPart, columnStyleIds);

        var pivotDef = BuildPivotTableDefinition(
            pivotName, cacheId, position, headers, columnData,
            rowFields, colFields, filterFields, valueFields, style, columnNumFmtIds);
        pivotPart.PivotTableDefinition = pivotDef;
        pivotPart.PivotTableDefinition.Save();

        // 6. RENDER the pivot output into the target sheet's <sheetData>.
        //
        // This is the critical step that distinguishes a "valid pivot file Excel
        // accepts" from a "pivot file Excel actually displays". Excel does NOT
        // recompute pivots from cache on open — it reads the rendered cells
        // directly from sheetData, exactly like any other range. We verified this
        // by inspecting an Excel-authored sample (excel_authored.xlsx → sheet2.xml):
        // every aggregated cell is a literal <c><v>200</v></c> element.
        //
        // Without this step the pivot opens as an empty drop-down skeleton — the
        // structure is valid but there is nothing to display. POI / Open XML SDK
        // suffer from exactly the same limitation; this is the lift that turns
        // officecli into a real pivot writer rather than a definition-only one.
        //
        // For unsupported configurations (multiple row/col fields, multiple data
        // fields, page filters), the renderer falls back to writing nothing, which
        // gives Excel an empty sheetData and the same skeleton-only behavior.
        // Those configs are tracked as a v2 expansion.
        RenderPivotIntoSheet(
            targetSheet, position, headers, columnData,
            rowFields, colFields, valueFields, filterFields, columnStyleIds);

        // Return 1-based index
        return targetSheet.PivotTableParts.ToList().IndexOf(pivotPart) + 1;
    }

    // ==================== Axis Tree (general N-level row/col abstraction) ====================
    //
    // For N≥3 row or col fields the existing specialized renderers (1×1, 2×1,
    // 1×2, 2×2 with K data variants) cannot be extended without an N² explosion
    // in case count. The AxisTree abstraction below replaces them with a single
    // recursive tree representation:
    //
    //   - The root has one child per unique value of the FIRST (outermost) field
    //   - Each level-L node has one child per unique value of the (L+1)-th field
    //     that appears in the source data PAIRED WITH the parent's path
    //   - Leaves are at depth N (i.e. path length = N field values)
    //
    // Example for rows=[地区, 城市, 区]:
    //   root
    //   ├── 华东
    //   │   ├── 上海
    //   │   │   ├── 浦东
    //   │   │   └── 徐汇
    //   │   └── 杭州
    //   │       └── 西湖
    //   └── 华北
    //       └── 北京
    //           ├── 朝阳
    //           └── 海淀
    //
    // Walk order produces (in display sequence): outer subtotals at internal
    // nodes + leaf rows at leaves + grand total at the very end. For 2D pivots
    // both row and col axes use independent AxisTrees and the renderer walks
    // them in lockstep.
    //
    // This abstraction is currently used ONLY for N≥3 cases via the dispatch in
    // RenderPivotIntoSheet. The 8 existing N≤2 cases continue to use their
    // specialized renderers (regression-tested via test-samples/pivot_baselines).

    /// <summary>
    /// One node in the axis tree. Represents either an internal node (subtotal
    /// row/col) or a leaf node (specific data row/col). Children are sorted in
    /// ordinal display order to keep rowItems/colItems indices consistent with
    /// the corresponding pivotField items list.
    /// </summary>
    private sealed class AxisNode
    {
        /// <summary>The label for this node (e.g. "华东"). Empty string for the root.</summary>
        public string Label { get; }
        /// <summary>0 = root, 1 = outermost field, 2 = next inner, ..., N = leaf level.</summary>
        public int Depth { get; }
        /// <summary>Path from root: [outerVal, ..., this.Label]. Length == Depth.</summary>
        public string[] Path { get; }
        /// <summary>Child nodes in ordinal display order. Empty for leaves.</summary>
        public List<AxisNode> Children { get; } = new();

        public AxisNode(string label, int depth, string[] path)
        {
            Label = label;
            Depth = depth;
            Path = path;
        }

        public bool IsLeaf => Children.Count == 0;
    }

    /// <summary>
    /// Build an AxisTree from columnData given the field indices for an axis.
    /// Only paths that actually appear in the source data are included — Excel
    /// does not enumerate empty cartesian intersections at any level.
    /// </summary>
    private static AxisNode BuildAxisTree(List<int> fieldIndices, List<string[]> columnData)
    {
        var root = new AxisNode(string.Empty, 0, Array.Empty<string>());
        if (fieldIndices.Count == 0 || columnData.Count == 0)
            return root;

        var rowCount = columnData[fieldIndices[0]].Length;
        // For each source row, walk down the tree, creating child nodes as needed.
        for (int r = 0; r < rowCount; r++)
        {
            var current = root;
            var validPath = true;
            var path = new string[fieldIndices.Count];

            for (int level = 0; level < fieldIndices.Count; level++)
            {
                var fieldIdx = fieldIndices[level];
                if (fieldIdx < 0 || fieldIdx >= columnData.Count) { validPath = false; break; }
                var values = columnData[fieldIdx];
                if (r >= values.Length) { validPath = false; break; }
                var v = values[r];
                if (string.IsNullOrEmpty(v)) { validPath = false; break; }
                path[level] = v;

                // Find or create child for this value at this level.
                var child = current.Children.FirstOrDefault(c => c.Label == v);
                if (child == null)
                {
                    var childPath = new string[level + 1];
                    Array.Copy(path, childPath, level + 1);
                    child = new AxisNode(v, level + 1, childPath);
                    current.Children.Add(child);
                }
                current = child;
            }

            // Drop the row entirely if any field had an empty value — matches the
            // "skip rows with missing values" semantics of the specialized renderers.
            _ = validPath;
        }

        // Sort children at every level using the same StringComparer.Ordinal that
        // BuildOuterInnerGroups and AppendFieldItems use, so the rowItems indices
        // line up with the pivotField items list.
        SortAxisTreeRecursive(root);
        return root;
    }

    private static void SortAxisTreeRecursive(AxisNode node)
    {
        var cmp = ActiveAxisComparer;
        var sign = ActiveAxisDescending ? -1 : 1;
        node.Children.Sort((a, b) => sign * cmp.Compare(a.Label, b.Label));
        foreach (var c in node.Children) SortAxisTreeRecursive(c);
    }

    /// <summary>
    /// Walk the tree in display order, yielding each node alongside whether it's
    /// a subtotal (internal) or a leaf, plus its absolute display row/col index
    /// (relative to the start of the data area).
    ///
    /// Display order for row axis is "pre-order": for each internal node, emit
    /// the subtotal row first, then recurse into children. The order matches
    /// what BuildMultiRowItems already produces for N=2 and what Excel writes
    /// for N≥3 in compact mode.
    ///
    /// For col axis it's the same plus an additional subtotal column AFTER the
    /// children of each internal node — Excel writes the col subtotal column
    /// to the right of the inner cols, not to the left like the row subtotal.
    /// </summary>
    private static IEnumerable<(AxisNode node, bool isLeaf, bool isSubtotal)> WalkAxisTree(
        AxisNode root, bool isCol)
    {
        // Skip the synthetic root, walk its children in order.
        foreach (var child in root.Children)
            foreach (var entry in WalkAxisTreeRecursive(child, isCol))
                yield return entry;
    }

    private static IEnumerable<(AxisNode node, bool isLeaf, bool isSubtotal)> WalkAxisTreeRecursive(
        AxisNode node, bool isCol)
    {
        if (node.IsLeaf)
        {
            yield return (node, true, false);
            yield break;
        }

        // Row axis convention: outer subtotal row appears BEFORE the children.
        // Col axis convention: outer subtotal col appears AFTER the children
        //                     (matches multi_col_authored.xlsx ground truth).
        if (!isCol)
            yield return (node, false, true);

        foreach (var child in node.Children)
            foreach (var entry in WalkAxisTreeRecursive(child, isCol))
                yield return entry;

        if (isCol)
            yield return (node, false, true);
    }

    /// <summary>Count all internal nodes (subtotal positions) in a tree.</summary>
    private static int CountSubtotalNodes(AxisNode root)
    {
        int count = 0;
        void Recurse(AxisNode n)
        {
            if (!n.IsLeaf && n.Depth > 0) count++;
            foreach (var c in n.Children) Recurse(c);
        }
        Recurse(root);
        return count;
    }

    /// <summary>Count all leaf nodes in a tree.</summary>
    private static int CountLeafNodes(AxisNode root)
    {
        int count = 0;
        void Recurse(AxisNode n)
        {
            if (n.IsLeaf && n.Depth > 0) count++;
            else foreach (var c in n.Children) Recurse(c);
        }
        Recurse(root);
        return count;
    }

    // ==================== Geometry & Cache Readback Helpers ====================

    /// <summary>Computed pivot table extent — anchor + bounding range + key offsets.</summary>
    private readonly struct PivotGeometry
    {
        public PivotGeometry(int anchorCol, int anchorRow, int width, int height, int rowLabelCols, string rangeRef)
        {
            AnchorCol = anchorCol;
            AnchorRow = anchorRow;
            Width = width;
            Height = height;
            RowLabelCols = rowLabelCols;
            RangeRef = rangeRef;
        }
        public int AnchorCol { get; }
        public int AnchorRow { get; }
        public int Width { get; }
        public int Height { get; }
        public int RowLabelCols { get; }
        public string RangeRef { get; }
    }

    /// <summary>
    /// Compute the bounding range and row-label column count for a pivot at the
    /// given anchor with the given field assignments. Used by both initial creation
    /// (BuildPivotTableDefinition) and post-Set rebuild (RebuildFieldAreas) so the
    /// two paths agree on layout.
    ///
    /// Layout assumes the standard compact/outline mode with:
    ///   width  = max(1, rowFieldCount)                    // row labels
    ///          + max(1, colUnique) * max(1, valueCount)    // data cells
    ///          + (colFieldCount > 0 ? 1 : 0)               // grand total column
    ///   height = (colFieldCount > 0 ? 2 : 1)               // header rows
    ///          + max(1, rowUnique)                          // data rows
    ///          + 1                                          // grand total row
    /// Page filter rows are excluded from the range per ECMA-376.
    /// </summary>
    private static PivotGeometry ComputePivotGeometry(
        string position, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string name)> valueFields)
    {
        int dataFieldCount = Math.Max(1, valueFields.Count);
        int rowLabelCols = 1; // Compact mode

        int valueCols, totalCols, dataRowCount, headerRows;

        // N≥3 on either axis: use AxisTree for both width and height counts.
        // N≤2: keep the existing specialized formulas (regression-tested).
        if (rowFieldIndices.Count >= 3 || colFieldIndices.Count >= 3)
        {
            var rowTree = BuildAxisTree(rowFieldIndices, columnData);
            var colTree = BuildAxisTree(colFieldIndices, columnData);

            // Display row count = subtotal positions + leaf positions
            // (the grand total row is added separately below).
            int rowSubtotals = CountSubtotalNodes(rowTree);
            int rowLeaves = CountLeafNodes(rowTree);
            dataRowCount = rowSubtotals + rowLeaves;

            int colSubtotals = CountSubtotalNodes(colTree);
            int colLeaves = CountLeafNodes(colTree);
            // Per col position: K cells. Plus K grand totals.
            valueCols = (colSubtotals + colLeaves) * dataFieldCount;
            totalCols = dataFieldCount;

            // Header rows: 1 caption + N_col field-label rows + (K>1 ? 1 : 0).
            headerRows = 1 + Math.Max(1, colFieldIndices.Count) + (dataFieldCount > 1 ? 1 : 0);
        }
        else if (colFieldIndices.Count >= 2)
        {
            var groups = BuildOuterInnerGroups(
                colFieldIndices[0], colFieldIndices[1], columnData);
            valueCols = groups.Sum(g => (g.inners.Count + 1) * dataFieldCount);
            totalCols = dataFieldCount;

            if (rowFieldIndices.Count >= 2)
            {
                var rowGroups = BuildOuterInnerGroups(
                    rowFieldIndices[0], rowFieldIndices[1], columnData);
                dataRowCount = rowGroups.Sum(g => 1 + g.inners.Count);
            }
            else
            {
                dataRowCount = Math.Max(1, ProductOfUniqueValues(rowFieldIndices, columnData));
            }
            headerRows = dataFieldCount > 1 ? 4 : 3;
        }
        else
        {
            int colUnique = ProductOfUniqueValues(colFieldIndices, columnData);
            valueCols = Math.Max(1, colUnique) * dataFieldCount;
            totalCols = colFieldIndices.Count > 0 ? dataFieldCount : 0;

            if (rowFieldIndices.Count >= 2)
            {
                var rowGroups = BuildOuterInnerGroups(
                    rowFieldIndices[0], rowFieldIndices[1], columnData);
                dataRowCount = rowGroups.Sum(g => 1 + g.inners.Count);
            }
            else
            {
                dataRowCount = Math.Max(1, ProductOfUniqueValues(rowFieldIndices, columnData));
            }

            if (colFieldIndices.Count > 0)
                headerRows = dataFieldCount > 1 ? 3 : 2;
            else
                headerRows = dataFieldCount > 1 ? 2 : 1;
        }

        int width = rowLabelCols + valueCols + totalCols;
        int height = headerRows + dataRowCount + 1;

        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var endColIdx = anchorColIdx + width - 1;
        var endRow = anchorRow + height - 1;
        var rangeRef = $"{position}:{IndexToCol(endColIdx)}{endRow}";

        return new PivotGeometry(anchorColIdx, anchorRow, width, height, rowLabelCols, rangeRef);
    }

    /// <summary>
    /// Reconstruct the per-field columnData from the cache definition + records.
    /// Used by RebuildFieldAreas after Set: the source sheet may not be readily
    /// reachable, but the cache holds the original values (string fields via
    /// sharedItems index, numeric fields directly in &lt;n v=...&gt;). This makes
    /// the rebuild self-contained on the cache part alone.
    /// </summary>
    private static (string[] headers, List<string[]> columnData) ReadColumnDataFromCache(
        PivotCacheDefinition cacheDef, PivotCacheRecords? records)
    {
        var cacheFields = cacheDef.GetFirstChild<CacheFields>();
        if (cacheFields == null) return (Array.Empty<string>(), new List<string[]>());

        var fieldList = cacheFields.Elements<CacheField>().ToList();
        var headers = fieldList.Select(cf => cf.Name?.Value ?? "").ToArray();
        var fieldCount = fieldList.Count;

        // Pre-resolve each field's sharedItems string lookup table (index → text).
        // Numeric fields without enumerated items leave the table empty; their
        // values come straight from <n v=...> in the records below.
        var perFieldStrings = new List<List<string>>(fieldCount);
        for (int f = 0; f < fieldCount; f++)
        {
            var items = fieldList[f].GetFirstChild<SharedItems>();
            var list = new List<string>();
            if (items != null)
            {
                foreach (var child in items.ChildElements)
                {
                    list.Add(child switch
                    {
                        StringItem s => s.Val?.Value ?? string.Empty,
                        NumberItem n => n.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                        DateTimeItem d => d.Val?.Value.ToString("yyyy-MM-dd") ?? string.Empty,
                        BooleanItem b => b.Val?.Value == true ? "true" : "false",
                        _ => string.Empty
                    });
                }
            }
            perFieldStrings.Add(list);
        }

        var recordList = records?.Elements<PivotCacheRecord>().ToList() ?? new List<PivotCacheRecord>();
        var columnData = new List<string[]>(fieldCount);
        for (int f = 0; f < fieldCount; f++)
            columnData.Add(new string[recordList.Count]);

        for (int r = 0; r < recordList.Count; r++)
        {
            var record = recordList[r];
            var children = record.ChildElements.ToList();
            for (int f = 0; f < fieldCount && f < children.Count; f++)
            {
                columnData[f][r] = children[f] switch
                {
                    FieldItem fi when fi.Val?.Value is uint idx
                        && idx < perFieldStrings[f].Count
                        => perFieldStrings[f][(int)idx],
                    NumberItem n => n.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                    StringItem s => s.Val?.Value ?? string.Empty,
                    DateTimeItem d => d.Val?.Value.ToString("yyyy-MM-dd") ?? string.Empty,
                    BooleanItem b => b.Val?.Value == true ? "true" : "false",
                    _ => string.Empty
                };
            }
        }

        return (headers, columnData);
    }

    /// <summary>
    /// Remove every cell in sheetData that falls inside the given pivot range.
    /// Called before re-rendering so stale cells from the previous pivot layout
    /// (e.g. row totals from a wider configuration) do not leak through.
    /// </summary>
    private static void ClearPivotRangeCells(SheetData sheetData, string rangeRef)
    {
        var parts = rangeRef.Split(':');
        if (parts.Length != 2) return;
        var (startCol, startRow) = ParseCellRef(parts[0]);
        var (endCol, endRow) = ParseCellRef(parts[1]);
        var startColIdx = ColToIndex(startCol);
        var endColIdx = ColToIndex(endCol);

        var rowsToRemove = new List<Row>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var rIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rIdx < startRow || rIdx > endRow) continue;

            var cellsToRemove = row.Elements<Cell>()
                .Where(c =>
                {
                    var cref = c.CellReference?.Value ?? "";
                    var (cc, _) = ParseCellRef(cref);
                    var ci = ColToIndex(cc);
                    return ci >= startColIdx && ci <= endColIdx;
                })
                .ToList();
            foreach (var c in cellsToRemove) c.Remove();

            // If the row is now empty AND was entirely inside the pivot, drop it
            // entirely so we don't leave stray <row r="N"/> elements behind.
            if (!row.Elements<Cell>().Any())
                rowsToRemove.Add(row);
        }
        foreach (var r in rowsToRemove) r.Remove();
    }

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
        List<(int idx, string func, string name)> valueFields,
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
        if (rowFieldIndices.Count >= 3 || colFieldIndices.Count >= 3)
        {
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

        if (rowFieldIndices.Count != 1 || colFieldIndices.Count != 1 || valueFields.Count < 1)
        {
            Console.Error.WriteLine(
                "WARNING: pivot rendering currently supports 1×1×K, 2×1×1, or 1×2×1 field combinations. " +
                "The file will open but the pivot will appear empty. " +
                "Use Excel's Refresh button to populate it manually.");
            return;
        }

        var rowFieldIdx = rowFieldIndices[0];
        var colFieldIdx = colFieldIndices[0];
        var rowFieldName = headers[rowFieldIdx];
        var colFieldName = headers[colFieldIdx];
        int K = valueFields.Count;

        var rowValues = columnData[rowFieldIdx];
        var colValues = columnData[colFieldIdx];

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
        var perBucket = new Dictionary<(string r, string c, int d), List<double>>();
        var perDataField = new List<List<double>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<double>());

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
                if (!double.TryParse(dataValues[i], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var num)) continue;

                var key = (rv, cv, d);
                if (!perBucket.TryGetValue(key, out var list))
                {
                    list = new List<double>();
                    perBucket[key] = list;
                }
                list.Add(num);
                perDataField[d].Add(num);
            }
        }

        double Reduce(IEnumerable<double> values, string func) => ReducePivotValues(values, func);

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
                var rowAll = new List<double>();
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
                var colAll = new List<double>();
                for (int r = 0; r < uniqueRows.Count; r++)
                {
                    if (perBucket.TryGetValue((uniqueRows[r], uniqueCols[c], d), out var bucket))
                        colAll.AddRange(bucket);
                }
                colTotals[c, d] = Reduce(colAll, func);
            }
            grandTotals[d] = Reduce(perDataField[d], func);
        }

        // ===== Write cells =====
        // For K=1, layout is 2 header rows: caption + col labels.
        // For K>1, layout is 3 header rows: caption + col labels + per-data-field
        // names repeated under each col label group. This matches the Excel sample
        // multi_data_authored.xlsx exactly.
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalColLabel = "总计";

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
                colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + c, colLabelRowIdx, uniqueCols[c]));
            colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + uniqueCols.Count, colLabelRowIdx, totalColLabel));
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
            int totalStart = anchorColIdx + 1 + uniqueCols.Count * K;
            for (int d = 0; d < K; d++)
                colLabelRow.AppendChild(MakeStringCell(totalStart + d, colLabelRowIdx, "Total " + valueFields[d].name));
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
            int rowTotalStart = anchorColIdx + 1 + uniqueCols.Count * K;
            for (int d = 0; d < K; d++)
                dataRow.AppendChild(MakeNumericCell(rowTotalStart + d, rowIdx, rowTotals[r, d], valueStyleIds[d]));
            sheetData.AppendChild(dataRow);
        }

        // ----- Grand total row -----
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
        int grandTotalStart = anchorColIdx + 1 + uniqueCols.Count * K;
        for (int d = 0; d < K; d++)
            grandRow.AppendChild(MakeNumericCell(grandTotalStart + d, grandRowIdx, grandTotals[d], valueStyleIds[d]));
        sheetData.AppendChild(grandRow);

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
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, "(All)"));
                    // Insert in row order: existing rows in sheetData start at
                    // anchorRow, so prepend the filter rows to the front.
                    sheetData.InsertAt(filterRow, fi);
                }
            }
            else
            {
                Console.Error.WriteLine(
                    $"WARNING: pivot at {position} has {filterFieldIndices.Count} page filter(s) " +
                    $"but only {anchorRow - 1} row(s) of headroom above. " +
                    "Filter cells will not be visible in the host sheet, but the filter dropdowns " +
                    "will still appear in Excel's pivot UI. Move the pivot to a lower anchor row " +
                    $"(at least row {requiredHeadroom + 1}) to render the filter cells.");
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
        List<(int idx, string func, string name)> valueFields,
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
        var leafBucket = new Dictionary<(string o, string i, string c, int d), List<double>>();
        var perDataField = new List<List<double>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<double>());

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
                if (!double.TryParse(dataValues[i], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var num)) continue;

                var key = (ov, iv, cv, d);
                if (!leafBucket.TryGetValue(key, out var list))
                {
                    list = new List<double>();
                    leafBucket[key] = list;
                }
                list.Add(num);
                perDataField[d].Add(num);
            }
        }

        double Reduce(IEnumerable<double> values, string func) => ReducePivotValues(values, func);

        // The closures below compute the cell values per (row pos, col pos, d)
        // by reducing raw value lists. Each closure takes a data field index d
        // so each data field aggregates with its own function (sum/count/avg/...).
        double LeafCell(string outer, string inner, string col, int d)
            => leafBucket.TryGetValue((outer, inner, col, d), out var b) && b.Count > 0
                ? Reduce(b, valueFields[d].func) : double.NaN;

        double OuterSubtotalForCol(string outer, string col, int d)
        {
            var all = new List<double>();
            foreach (var (o, inners) in groups)
                if (o == outer)
                    foreach (var inner in inners)
                        if (leafBucket.TryGetValue((outer, inner, col, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double LeafRowTotal(string outer, string inner, int d)
        {
            var all = new List<double>();
            foreach (var col in uniqueCols)
                if (leafBucket.TryGetValue((outer, inner, col, d), out var b))
                    all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterRowTotal(string outer, int d)
        {
            var all = new List<double>();
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
            var all = new List<double>();
            foreach (var (outer, inners) in groups)
                foreach (var inner in inners)
                    if (leafBucket.TryGetValue((outer, inner, col, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        // ===== Write cells =====
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalLabel = "总计";

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
            colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + uniqueCols.Count, colLabelRowIdx, totalLabel));
        }
        else
        {
            for (int c = 0; c < uniqueCols.Count; c++)
                colLabelRow.AppendChild(MakeStringCell(LeafColIdx(c, 0), colLabelRowIdx, uniqueCols[c]));
            for (int d = 0; d < K; d++)
                colLabelRow.AppendChild(MakeStringCell(GrandTotalColIdx(d), colLabelRowIdx, "Total " + valueFields[d].name));
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

        // ----- Data rows -----
        int currentRow = firstDataRow;
        foreach (var (outer, inners) in groups)
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
            for (int d = 0; d < K; d++)
                subRow.AppendChild(MakeNumericCell(GrandTotalColIdx(d), currentRow, OuterRowTotal(outer, d), valueStyleIds[d]));
            sheetData.AppendChild(subRow);
            currentRow++;

            // Leaf rows for each existing (outer, inner) combo.
            foreach (var inner in inners)
            {
                var leafRow = new Row { RowIndex = (uint)currentRow };
                leafRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, inner));
                for (int c = 0; c < uniqueCols.Count; c++)
                {
                    for (int d = 0; d < K; d++)
                    {
                        var v = LeafCell(outer, inner, uniqueCols[c], d);
                        if (!double.IsNaN(v))
                            leafRow.AppendChild(MakeNumericCell(LeafColIdx(c, d), currentRow, v, valueStyleIds[d]));
                    }
                }
                for (int d = 0; d < K; d++)
                    leafRow.AppendChild(MakeNumericCell(GrandTotalColIdx(d), currentRow, LeafRowTotal(outer, inner, d), valueStyleIds[d]));
                sheetData.AppendChild(leafRow);
                currentRow++;
            }
        }

        // Grand total row.
        var grandRow = new Row { RowIndex = (uint)currentRow };
        grandRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, totalLabel));
        for (int c = 0; c < uniqueCols.Count; c++)
            for (int d = 0; d < K; d++)
                grandRow.AppendChild(MakeNumericCell(LeafColIdx(c, d), currentRow, ColTotal(uniqueCols[c], d), valueStyleIds[d]));
        for (int d = 0; d < K; d++)
            grandRow.AppendChild(MakeNumericCell(GrandTotalColIdx(d), currentRow,
                Reduce(perDataField[d], valueFields[d].func), valueStyleIds[d]));
        sheetData.AppendChild(grandRow);

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
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, "(All)"));
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
        List<(int idx, string func, string name)> valueFields,
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
        var leafBucket = new Dictionary<(string r, string oc, string ic, int d), List<double>>();
        var perDataField = new List<List<double>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<double>());

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
                if (!double.TryParse(dataValues[i], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var num)) continue;

                var key = (rv, ocv, icv, d);
                if (!leafBucket.TryGetValue(key, out var list))
                {
                    list = new List<double>();
                    leafBucket[key] = list;
                }
                list.Add(num);
                perDataField[d].Add(num);
            }
        }

        double Reduce(IEnumerable<double> values, string func) => ReducePivotValues(values, func);

        // Per-(row, outerCol, innerCol, d) reductions over raw values.
        double LeafCell(string row, string outerCol, string innerCol, int d)
            => leafBucket.TryGetValue((row, outerCol, innerCol, d), out var b) && b.Count > 0
                ? Reduce(b, valueFields[d].func) : double.NaN;

        double OuterColSubtotalForRow(string row, string outerCol, int d)
        {
            var all = new List<double>();
            foreach (var (oc, inners) in colGroups)
                if (oc == outerCol)
                    foreach (var inner in inners)
                        if (leafBucket.TryGetValue((row, outerCol, inner, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double RowGrandTotal(string row, int d)
        {
            var all = new List<double>();
            foreach (var (oc, inners) in colGroups)
                foreach (var inner in inners)
                    if (leafBucket.TryGetValue((row, oc, inner, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double LeafColTotal(string outerCol, string innerCol, int d)
        {
            var all = new List<double>();
            foreach (var row in uniqueRows)
                if (leafBucket.TryGetValue((row, outerCol, innerCol, d), out var b))
                    all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterColTotal(string outerCol, int d)
        {
            var all = new List<double>();
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
        var totalLabel = "总计";

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // Pre-compute absolute column indices. K data fields multiply the leaf
        // and subtotal positions by K. Layout (left to right):
        //   row label
        //   For each outer:
        //     For each inner:                            K cells (data fields)
        //     subtotal:                                  K cells (per-data subtotal)
        //   grand total:                                 K cells (per-data grand)
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
            for (int d = 0; d < K; d++)
            {
                subtotalColPositions[(outer, d)] = currentCol;
                currentCol++;
            }
        }
        for (int d = 0; d < K; d++)
        {
            grandTotalColPositions[d] = currentCol;
            currentCol++;
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
                innerHeaderRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, 0)], innerHeaderRowIdx, outer + " Total"));
            }
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
                for (int d = 0; d < K; d++)
                    outerHeaderRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, d)],
                        outerHeaderRowIdx, $"{outer} {valueFields[d].name}"));
            }
            for (int d = 0; d < K; d++)
                outerHeaderRow.AppendChild(MakeStringCell(grandTotalColPositions[d],
                    outerHeaderRowIdx, $"Total {valueFields[d].name}"));
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
                // Outer col subtotal cells (K per outer).
                bool any = HasAnyValueInRowOuter(uniqueRows[r], outer, colGroups, leafBucket, K);
                for (int d = 0; d < K; d++)
                {
                    var sub = OuterColSubtotalForRow(uniqueRows[r], outer, d);
                    if (sub != 0 || any)
                        dataRow.AppendChild(MakeNumericCell(subtotalColPositions[(outer, d)], rowIdx, sub, valueStyleIds[d]));
                }
            }

            for (int d = 0; d < K; d++)
                dataRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], rowIdx, RowGrandTotal(uniqueRows[r], d), valueStyleIds[d]));
            sheetData.AppendChild(dataRow);
        }

        // Grand total row.
        int grandRowIdx = firstDataRow + uniqueRows.Count;
        var grandRow = new Row { RowIndex = (uint)grandRowIdx };
        grandRow.AppendChild(MakeStringCell(anchorColIdx, grandRowIdx, totalLabel));
        foreach (var (outer, inners) in colGroups)
        {
            foreach (var inner in inners)
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(leafColPositions[(outer, inner, d)], grandRowIdx,
                        LeafColTotal(outer, inner, d), valueStyleIds[d]));
            for (int d = 0; d < K; d++)
                grandRow.AppendChild(MakeNumericCell(subtotalColPositions[(outer, d)], grandRowIdx, OuterColTotal(outer, d), valueStyleIds[d]));
        }
        for (int d = 0; d < K; d++)
            grandRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], grandRowIdx,
                Reduce(perDataField[d], valueFields[d].func), valueStyleIds[d]));
        sheetData.AppendChild(grandRow);

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
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, "(All)"));
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
        List<(int idx, string func, string name)> valueFields,
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
        var bucket = new Dictionary<(string ro, string ri, string co, string ci, int d), List<double>>();
        var perDataField = new List<List<double>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<double>());

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
                if (!double.TryParse(dataValues[i], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var num)) continue;

                var key = (ro, ri, co, ci, d);
                if (!bucket.TryGetValue(key, out var list))
                {
                    list = new List<double>();
                    bucket[key] = list;
                }
                list.Add(num);
                perDataField[d].Add(num);
            }
        }

        double Reduce(IEnumerable<double> values, string func) => ReducePivotValues(values, func);

        // The 9 cell-value closures from the K=1 path now each take a data
        // field index d so the right aggregator is applied per cell.
        double LeafCell(string ro, string ri, string co, string ci, int d)
            => bucket.TryGetValue((ro, ri, co, ci, d), out var b) && b.Count > 0
                ? Reduce(b, valueFields[d].func) : double.NaN;

        double LeafRowColSub(string ro, string ri, string co, int d)
        {
            var all = new List<double>();
            foreach (var (oc, inners) in colGroups)
                if (oc == co)
                    foreach (var inner in inners)
                        if (bucket.TryGetValue((ro, ri, co, inner, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double LeafRowGrandTotal(string ro, string ri, int d)
        {
            var all = new List<double>();
            foreach (var (oc, inners) in colGroups)
                foreach (var inner in inners)
                    if (bucket.TryGetValue((ro, ri, oc, inner, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterRowLeafCell(string ro, string co, string ci, int d)
        {
            var all = new List<double>();
            foreach (var (g, inners) in rowGroups)
                if (g == ro)
                    foreach (var inner in inners)
                        if (bucket.TryGetValue((ro, inner, co, ci, d), out var b))
                            all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double OuterRowColSub(string ro, string co, int d)
        {
            var all = new List<double>();
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
            var all = new List<double>();
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
            var all = new List<double>();
            foreach (var (g, rinners) in rowGroups)
                foreach (var rinner in rinners)
                    if (bucket.TryGetValue((g, rinner, co, ci, d), out var b))
                        all.AddRange(b);
            return Reduce(all, valueFields[d].func);
        }

        double GrandRowColSub(string co, int d)
        {
            var all = new List<double>();
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
        var totalLabel = "总计";

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // Pre-compute K-aware col positions: each (outer, inner) leaf gets K
        // cells, each outer subtotal gets K cells, K final grand total cells.
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
            for (int d = 0; d < K; d++)
            {
                subtotalColPositions[(outer, d)] = currentCol;
                currentCol++;
            }
        }
        for (int d = 0; d < K; d++)
        {
            grandTotalColPositions[d] = currentCol;
            currentCol++;
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
                innerHdrRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, 0)], innerHdrRowIdx, outer + " Total"));
            }
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
                for (int d = 0; d < K; d++)
                    outerHdrRow.AppendChild(MakeStringCell(subtotalColPositions[(outer, d)],
                        outerHdrRowIdx, $"{outer} {valueFields[d].name}"));
            }
            for (int d = 0; d < K; d++)
                outerHdrRow.AppendChild(MakeStringCell(grandTotalColPositions[d],
                    outerHdrRowIdx, $"Total {valueFields[d].name}"));
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
            for (int d = 0; d < K; d++)
                outerSubRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], currentRowIdx, OuterRowGrandTotal(rowOuter, d), valueStyleIds[d]));
            sheetData.AppendChild(outerSubRow);
            currentRowIdx++;

            // Leaf rows for each existing inner of this row outer.
            foreach (var rowInner in rowInners)
            {
                var leafRow = new Row { RowIndex = (uint)currentRowIdx };
                leafRow.AppendChild(MakeStringCell(anchorColIdx, currentRowIdx, rowInner));
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
                    bool any = HasAnyValueInLeafRowCol(rowOuter, rowInner, colOuter, colGroups, bucket, K);
                    for (int d = 0; d < K; d++)
                    {
                        var sub = LeafRowColSub(rowOuter, rowInner, colOuter, d);
                        if (sub != 0 || any)
                            leafRow.AppendChild(MakeNumericCell(subtotalColPositions[(colOuter, d)], currentRowIdx, sub, valueStyleIds[d]));
                    }
                }
                for (int d = 0; d < K; d++)
                    leafRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], currentRowIdx, LeafRowGrandTotal(rowOuter, rowInner, d), valueStyleIds[d]));
                sheetData.AppendChild(leafRow);
                currentRowIdx++;
            }
        }

        // Grand total row.
        var grandRow = new Row { RowIndex = (uint)currentRowIdx };
        grandRow.AppendChild(MakeStringCell(anchorColIdx, currentRowIdx, totalLabel));
        foreach (var (colOuter, colInners) in colGroups)
        {
            foreach (var colInner in colInners)
                for (int d = 0; d < K; d++)
                    grandRow.AppendChild(MakeNumericCell(leafColPositions[(colOuter, colInner, d)], currentRowIdx,
                        GrandRowLeafCol(colOuter, colInner, d), valueStyleIds[d]));
            for (int d = 0; d < K; d++)
                grandRow.AppendChild(MakeNumericCell(subtotalColPositions[(colOuter, d)], currentRowIdx, GrandRowColSub(colOuter, d), valueStyleIds[d]));
        }
        for (int d = 0; d < K; d++)
            grandRow.AppendChild(MakeNumericCell(grandTotalColPositions[d], currentRowIdx,
                Reduce(perDataField[d], valueFields[d].func), valueStyleIds[d]));
        sheetData.AppendChild(grandRow);

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
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, "(All)"));
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
        List<(int idx, string func, string name)> valueFields,
        List<int>? filterFieldIndices,
        uint?[] valueStyleIds)
    {
        int K = Math.Max(1, valueFields.Count);
        var rowTree = BuildAxisTree(rowFieldIndices, columnData);
        var colTree = BuildAxisTree(colFieldIndices, columnData);

        // Walk both trees in display order. Each entry is the absolute display
        // position relative to the start of the data area.
        var rowPositions = WalkAxisTree(rowTree, isCol: false).ToList();
        var colPositions = WalkAxisTree(colTree, isCol: true).ToList();

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

        // Numeric value cache per data field. Pre-parse so we don't double_parse
        // every cell access. NaN encodes "not a number / skip".
        var dataNums = new double[K][];
        for (int d = 0; d < K; d++)
        {
            var dataIdx = valueFields[d].idx;
            var values = (dataIdx >= 0 && dataIdx < columnData.Count) ? columnData[dataIdx] : Array.Empty<string>();
            dataNums[d] = new double[srcRowCount];
            for (int r = 0; r < srcRowCount; r++)
            {
                if (r >= values.Length || string.IsNullOrEmpty(values[r])
                    || !double.TryParse(values[r], System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var n))
                    dataNums[d][r] = double.NaN;
                else
                    dataNums[d][r] = n;
            }
        }

        double Reduce(IEnumerable<double> values, string func) => ReducePivotValues(values, func);

        // Compute the value at (rowNode, colNode, dataFieldIdx).
        // Subtotal nodes have shorter Path arrays than leaves; the prefix match
        // automatically widens the set of source rows that contribute.
        double ComputeCell(AxisNode rowNode, AxisNode colNode, int d)
        {
            var rPath = rowNode.Path;
            var cPath = colNode.Path;
            var collected = new List<double>();
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

                var v = dataNums[d][r];
                if (!double.IsNaN(v)) collected.Add(v);
            }
            return Reduce(collected, valueFields[d].func);
        }

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
                    if (!double.IsNaN(dataNums[d][r])) return true;
            }
            return false;
        }

        // ===== Write cells =====
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalLabel = "总计";

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // Pre-compute absolute col indices for every col position × data field.
        // colPositions does not include the grand total column — that's tracked
        // separately so the writer doesn't accidentally include it inside the
        // per-outer subtotal block.
        int colCells = colPositions.Count * K;
        int firstDataCol = anchorColIdx + 1;
        var colIdxByPosition = new int[colPositions.Count, K];
        for (int p = 0; p < colPositions.Count; p++)
            for (int d = 0; d < K; d++)
                colIdxByPosition[p, d] = firstDataCol + p * K + d;
        int grandTotalColStart = firstDataCol + colCells;

        // Header rows. Layout depends on (N_col, K):
        //   - 1 caption row (row 0)
        //   - N_col header rows (one per col field level, top→bottom = outer→inner)
        //   - Optionally 1 data-field-name row when K>1
        int headerRows = 1 + Math.Max(1, colFieldIndices.Count) + (K > 1 ? 1 : 0);

        // Row 0 (caption): col field caption (the outermost col field name) at
        // first data col position. For K=1 the row-label col also gets the
        // single data field name.
        var captionRow = new Row { RowIndex = (uint)anchorRow };
        if (K == 1)
            captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, valueFields[0].name));
        if (colFieldIndices.Count > 0)
            captionRow.AppendChild(MakeStringCell(firstDataCol, anchorRow,
                headers[colFieldIndices[0]]));
        sheetData.AppendChild(captionRow);

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
            // outermost row field name (when K=1) or stays empty (when K>1
            // because the data-field-name row below carries it).
            if (level == colFieldIndices.Count && K == 1 && rowFieldIndices.Count > 0)
                headerRow.AppendChild(MakeStringCell(anchorColIdx, headerRowIdx, headers[rowFieldIndices[0]]));

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
            // (or in the K>1 case it's spread across all data field columns).
            if (level == colFieldIndices.Count)
            {
                if (K == 1)
                    headerRow.AppendChild(MakeStringCell(grandTotalColStart, headerRowIdx, totalLabel));
                else
                    for (int d = 0; d < K; d++)
                        headerRow.AppendChild(MakeStringCell(grandTotalColStart + d, headerRowIdx,
                            $"Total {valueFields[d].name}"));
            }
            sheetData.AppendChild(headerRow);
        }

        // Optional data field name row (K>1).
        if (K > 1)
        {
            int dfRowIdx = anchorRow + headerRows - 1;
            var dfRow = new Row { RowIndex = (uint)dfRowIdx };
            if (rowFieldIndices.Count > 0)
                dfRow.AppendChild(MakeStringCell(anchorColIdx, dfRowIdx, headers[rowFieldIndices[0]]));
            for (int p = 0; p < colPositions.Count; p++)
            {
                var (_, isLeaf, isSubtotal) = colPositions[p];
                if (isSubtotal) continue; // Subtotal cols already labelled in their header row above.
                for (int d = 0; d < K; d++)
                    dfRow.AppendChild(MakeStringCell(colIdxByPosition[p, d], dfRowIdx, valueFields[d].name));
            }
            sheetData.AppendChild(dfRow);
        }

        // Data + grand total rows.
        int firstDataRowIdx = anchorRow + headerRows;
        for (int rp = 0; rp < rowPositions.Count; rp++)
        {
            var (rowNode, rIsLeaf, rIsSubtotal) = rowPositions[rp];
            int rowIdx = firstDataRowIdx + rp;
            var row = new Row { RowIndex = (uint)rowIdx };
            row.AppendChild(MakeStringCell(anchorColIdx, rowIdx, rowNode.Label));

            for (int cp = 0; cp < colPositions.Count; cp++)
            {
                var (colNode, cIsLeaf, cIsSubtotal) = colPositions[cp];
                bool any = HasAnyValue(rowNode, colNode);
                for (int d = 0; d < K; d++)
                {
                    var v = ComputeCell(rowNode, colNode, d);
                    // Skip 0-value cells when there are no underlying values to
                    // mirror Excel's behavior of leaving sparse intersections blank.
                    if (any || v != 0)
                        row.AppendChild(MakeNumericCell(colIdxByPosition[cp, d], rowIdx, v, valueStyleIds[d]));
                }
            }

            // Grand total cells (per data field) — the row's value across all cols.
            var grandRowNode = new AxisNode(string.Empty, 0, Array.Empty<string>());
            for (int d = 0; d < K; d++)
                row.AppendChild(MakeNumericCell(grandTotalColStart + d, rowIdx,
                    ComputeCell(rowNode, grandRowNode, d), valueStyleIds[d]));
            sheetData.AppendChild(row);
        }

        // Final grand total row.
        int grandRowIdx = firstDataRowIdx + rowPositions.Count;
        var grandRow = new Row { RowIndex = (uint)grandRowIdx };
        grandRow.AppendChild(MakeStringCell(anchorColIdx, grandRowIdx, totalLabel));
        var grandRowNodeFinal = new AxisNode(string.Empty, 0, Array.Empty<string>());
        for (int cp = 0; cp < colPositions.Count; cp++)
        {
            var (colNode, _, _) = colPositions[cp];
            for (int d = 0; d < K; d++)
            {
                var v = ComputeCell(grandRowNodeFinal, colNode, d);
                grandRow.AppendChild(MakeNumericCell(colIdxByPosition[cp, d], grandRowIdx, v, valueStyleIds[d]));
            }
        }
        for (int d = 0; d < K; d++)
            grandRow.AppendChild(MakeNumericCell(grandTotalColStart + d, grandRowIdx,
                ComputeCell(grandRowNodeFinal, grandRowNodeFinal, d), valueStyleIds[d]));
        sheetData.AppendChild(grandRow);

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
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, "(All)"));
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
        Dictionary<(string ro, string ri, string co, string ci, int d), List<double>> bucket,
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
        Dictionary<(string ro, string ri, string co, string ci, int d), List<double>> bucket,
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
        Dictionary<(string ro, string ri, string co, string ci, int d), List<double>> bucket,
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
        Dictionary<(string r, string oc, string ic, int d), List<double>> leafBucket,
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
        Dictionary<(string o, string i, string c, int d), List<double>> leafBucket,
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

    // ==================== Source Data Reader ====================

    private static (string[] headers, List<string[]> columnData, uint?[] columnStyleIds) ReadSourceData(
        WorksheetPart sourceSheet, string sourceRef)
    {
        var ws = sourceSheet.Worksheet ?? throw new InvalidOperationException("Worksheet missing");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null) return (Array.Empty<string>(), new List<string[]>(), Array.Empty<uint?>());

        // Parse range "A1:D100"
        var parts = sourceRef.Replace("$", "").Split(':');
        if (parts.Length != 2) throw new ArgumentException($"Invalid source range: {sourceRef}");

        var (startCol, startRow) = ParseCellRef(parts[0]);
        var (endCol, endRow) = ParseCellRef(parts[1]);

        var startColIdx = ColToIndex(startCol);
        var endColIdx = ColToIndex(endCol);
        var colCount = endColIdx - startColIdx + 1;

        // Read all rows in range. We also capture the StyleIndex of the first
        // non-empty data cell per column (skipping the header row) so pivot
        // value cells can inherit the source column's number format. This
        // mirrors how Excel's pivot engine picks the column format: it looks
        // at the data-area formatting, not the header.
        var rows = new List<string[]>();
        var columnStyleIds = new uint?[colCount];
        var sst = sourceSheet.OpenXmlPackage is SpreadsheetDocument doc
            ? doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
            : null;

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;

            var values = new string[colCount];
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value ?? "";
                var (cn, _) = ParseCellRef(cellRef);
                var ci = ColToIndex(cn) - startColIdx;
                if (ci < 0 || ci >= colCount) continue;

                values[ci] = GetCellText(cell, sst);

                // Capture style from first non-header data cell per column.
                // rowIdx > startRow skips the header row; we keep the first
                // one we encounter and ignore subsequent rows.
                if (rowIdx > startRow && columnStyleIds[ci] == null && cell.StyleIndex?.Value is uint sIdx && sIdx != 0)
                    columnStyleIds[ci] = sIdx;
            }
            rows.Add(values);
        }

        if (rows.Count == 0) return (Array.Empty<string>(), new List<string[]>(), Array.Empty<uint?>());

        // First row = headers (ensure no nulls)
        var headers = rows[0].Select(h => h ?? "").ToArray();
        // Remaining rows = data, transposed to column-major for cache
        var columnDataList = new List<string[]>();
        for (int c = 0; c < colCount; c++)
        {
            var colVals = new string[rows.Count - 1];
            for (int r = 1; r < rows.Count; r++)
                colVals[r - 1] = rows[r][c] ?? "";
            columnDataList.Add(colVals);
        }

        return (headers, columnDataList, columnStyleIds);
    }

    private static string GetCellText(Cell cell, SharedStringTablePart? sst)
    {
        // Handle InlineString cells (t="inlineStr") — used by openpyxl and some other tools
        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? "";

        var value = cell.CellValue?.Text ?? "";
        if (cell.DataType?.Value == CellValues.SharedString && sst?.SharedStringTable != null)
        {
            if (int.TryParse(value, out int idx))
            {
                var item = sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? value;
            }
        }
        return value;
    }

    // ==================== Cache Definition Builder ====================

    private static (PivotCacheDefinition def, bool[] fieldNumeric, Dictionary<string, int>[] fieldValueIndex)
        BuildCacheDefinition(
            string sourceSheetName, string sourceRef,
            string[] headers, List<string[]> columnData)
    {
        var recordCount = columnData.Count > 0 ? columnData[0].Length : 0;

        // refreshOnLoad=1 tells Excel to re-render the pivot from the cache when the
        // file is opened. We need this because officecli (a pure DOM library) does NOT
        // have a pivot computation engine — we cannot materialize the rendered cells
        // into sheetData ourselves. Real Excel/LibreOffice DO write rendered cells on
        // save (verified against pivot5.xlsx and pivot_dark1.xlsx fixtures), so opening
        // their files shows data immediately. Without refreshOnLoad, our pivot-only
        // sheet would render empty even though the cache and definition are valid.
        //
        // Trade-off: Excel may prompt for trust before refreshing, and consumers that
        // do not implement refresh (POI, third-party parsers) will still see an empty
        // sheet. The proper long-term fix is a built-in render engine; this flag is
        // the lowest-cost workaround until that lands.
        var cacheDef = new PivotCacheDefinition
        {
            CreatedVersion = 3,
            MinRefreshableVersion = 3,
            RefreshedVersion = 3,
            RecordCount = (uint)recordCount,
            RefreshOnLoad = true
        };

        // CacheSource -> WorksheetSource
        var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
        cacheSource.AppendChild(new WorksheetSource
        {
            Reference = sourceRef,
            Sheet = sourceSheetName
        });
        cacheDef.AppendChild(cacheSource);

        // CacheFields — also build per-field metadata used to write records:
        //   - fieldNumeric[i]: true if field i is numeric (records emit <n v=".."/>)
        //   - fieldValueIndex[i]: value→sharedItems index map for non-numeric fields
        //     (records emit <x v="N"/> referencing this index)
        var fieldNumeric = new bool[headers.Length];
        var fieldValueIndex = new Dictionary<string, int>[headers.Length];

        var cacheFields = new CacheFields { Count = (uint)headers.Length };
        for (int i = 0; i < headers.Length; i++)
        {
            var fieldName = string.IsNullOrEmpty(headers[i]) ? $"Column{i + 1}" : headers[i];
            var values = i < columnData.Count ? columnData[i] : Array.Empty<string>();
            cacheFields.AppendChild(BuildCacheField(fieldName, values, out fieldNumeric[i], out fieldValueIndex[i]));
        }
        cacheDef.AppendChild(cacheFields);

        return (cacheDef, fieldNumeric, fieldValueIndex);
    }

    private static CacheField BuildCacheField(
        string name, string[] values, out bool isNumeric, out Dictionary<string, int> valueIndex)
    {
        var field = new CacheField { Name = name, NumberFormatId = 0u };
        isNumeric = values.Length > 0 && values.All(v =>
            string.IsNullOrEmpty(v) || double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _));
        valueIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        var sharedItems = new SharedItems();

        // MIXED strategy — verified against Microsoft's own pivot5.xlsx (in
        // OPEN-XML-SDK test fixtures, authored by real Excel):
        //
        //   • Numeric fields: emit ONLY containsNumber/minValue/maxValue metadata,
        //     no enumerated items, no count attribute. Records reference values
        //     directly via <n v="..."/>.
        //   • String fields: enumerate every unique value as <s v="..."/> with
        //     count attribute. Records reference them by index via <x v="N"/>.
        //
        // I previously experimented with LibreOffice's uniform strategy (always
        // enumerate, always index-reference), but Microsoft's actual format is
        // the mixed one — and matching the real Excel format is the safest bet
        // for round-trip compatibility. The uniform strategy is technically valid
        // OOXML but introduces an asymmetry that Excel handles less reliably
        // (numeric data fields with item enumeration have failed to render in
        // testing, even though the file passes schema validation).
        if (isNumeric && values.Any(v => !string.IsNullOrEmpty(v)))
        {
            var nums = values.Where(v => !string.IsNullOrEmpty(v))
                .Select(v => double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            sharedItems.ContainsSemiMixedTypes = false;
            sharedItems.ContainsString = false;
            sharedItems.ContainsNumber = true;
            sharedItems.MinValue = nums.Min();
            sharedItems.MaxValue = nums.Max();
            // No items enumerated, no count — records emit <n v="..."/> directly.
        }
        else
        {
            var uniqueValues = values
                .Where(v => !string.IsNullOrEmpty(v))
                .Distinct()
                .OrderByAxis(v => v)
                .ToList();
            sharedItems.Count = (uint)uniqueValues.Count;
            for (int i = 0; i < uniqueValues.Count; i++)
            {
                var v = uniqueValues[i];
                sharedItems.AppendChild(new StringItem { Val = v });
                if (!valueIndex.ContainsKey(v))
                    valueIndex[v] = i;
            }
        }

        field.AppendChild(sharedItems);
        return field;
    }

    // ==================== Cache Records Builder ====================

    /// <summary>
    /// Build pivotCacheRecords using the MIXED strategy verified against Microsoft's
    /// own pivot5.xlsx test fixture:
    ///
    ///   <r>
    ///     <x v="0"/>     <!-- string field, references sharedItems[0] -->
    ///     <x v="2"/>     <!-- string field, references sharedItems[2] -->
    ///     <n v="702"/>   <!-- numeric field, value written directly -->
    ///     <m/>           <!-- empty/missing value -->
    ///   </r>
    ///
    /// String fields use indexed references (<x v="N"/>) into the per-field
    /// sharedItems list; numeric fields use NumberItem (<n v="V"/>) directly,
    /// because their cacheField only carries min/max metadata, not enumerated items.
    /// </summary>
    private static PivotCacheRecords BuildCacheRecords(
        List<string[]> columnData, bool[] fieldNumeric, Dictionary<string, int>[] fieldValueIndex)
    {
        var recordCount = columnData.Count > 0 ? columnData[0].Length : 0;
        var fieldCount = columnData.Count;
        var records = new PivotCacheRecords { Count = (uint)recordCount };

        for (int r = 0; r < recordCount; r++)
        {
            var record = new PivotCacheRecord();
            for (int f = 0; f < fieldCount; f++)
            {
                var v = columnData[f][r];
                if (string.IsNullOrEmpty(v))
                {
                    record.AppendChild(new MissingItem());
                }
                else if (fieldNumeric[f])
                {
                    record.AppendChild(new NumberItem
                    {
                        Val = double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)
                    });
                }
                else if (fieldValueIndex[f].TryGetValue(v, out var idx))
                {
                    // FieldItem = <x v="N"/> in OpenXml SDK, references sharedItems[N].
                    record.AppendChild(new FieldItem { Val = (uint)idx });
                }
                else
                {
                    // Defensive: value missing from the per-field index map. Should
                    // not occur since the map is built from the same columnData;
                    // emit <m/> rather than a dangling reference.
                    record.AppendChild(new MissingItem());
                }
            }
            records.AppendChild(record);
        }

        return records;
    }

    // ==================== Pivot Table Definition Builder ====================

    /// <summary>
    /// Resolve each source column's StyleIndex into the numFmtId that Excel
    /// actually needs on DataField. Returns null entries for columns whose
    /// source cell had no explicit style (→ General) so the caller can leave
    /// DataField.NumberFormatId unset.
    /// </summary>
    private static uint?[] ResolveColumnNumFmtIds(WorkbookPart workbookPart, uint?[] columnStyleIds)
    {
        var result = new uint?[columnStyleIds.Length];
        var stylesPart = workbookPart.WorkbookStylesPart;
        var cellXfs = stylesPart?.Stylesheet?.CellFormats?.Elements<CellFormat>().ToList();
        if (cellXfs == null) return result;
        for (int i = 0; i < columnStyleIds.Length; i++)
        {
            var sIdx = columnStyleIds[i];
            if (!sIdx.HasValue) continue;
            if (sIdx.Value >= cellXfs.Count) continue;
            var xf = cellXfs[(int)sIdx.Value];
            var numFmtId = xf.NumberFormatId?.Value;
            // numFmtId == 0 is General → no-op, skip so DataField stays plain
            if (numFmtId.HasValue && numFmtId.Value != 0)
                result[i] = numFmtId.Value;
        }
        return result;
    }

    private static PivotTableDefinition BuildPivotTableDefinition(
        string name, uint cacheId, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<int> filterFieldIndices, List<(int idx, string func, string name)> valueFields,
        string styleName,
        uint?[]? columnNumFmtIds = null)
    {
        var pivotDef = new PivotTableDefinition
        {
            Name = name,
            CacheId = cacheId,
            DataCaption = "Values",
            CreatedVersion = 3,
            MinRefreshableVersion = 3,
            UpdatedVersion = 3,
            ApplyNumberFormats = false,
            ApplyBorderFormats = false,
            ApplyFontFormats = false,
            ApplyPatternFormats = false,
            ApplyAlignmentFormats = false,
            ApplyWidthHeightFormats = true,
            UseAutoFormatting = true,
            ItemPrintTitles = true,
            MultipleFieldFilters = false,
            Indent = 0u,
            // outline + outlineData are emitted by both Microsoft Excel (pivot5.xlsx)
            // and LibreOffice (pivot_dark1.xlsx). They select the "outline" layout —
            // the default presentation where row labels stack into one column. Without
            // these, Excel falls back to a layout that's not fully wired through and
            // refuses to render the data area.
            Outline = true,
            OutlineData = true,
            // Caption attributes — when present, Excel uses these strings instead
            // of its locale-default "Row Labels" / "Column Labels" / "Grand Total".
            // Without these the rendered cells we wrote into sheetData ("地区",
            // "产品", "总计") get visually overlaid by Excel's English defaults
            // because the pivot's caption layer takes precedence over cell content
            // when the corresponding caption attribute is empty/missing.
            RowHeaderCaption = rowFieldIndices.Count > 0 ? headers[rowFieldIndices[0]] : "Rows",
            ColumnHeaderCaption = colFieldIndices.Count > 0 ? headers[colFieldIndices[0]] : "Columns",
            GrandTotalCaption = "总计"
        };

        // Use typed property setters to ensure correct schema order

        // Compute the pivot's geometry (range + offsets) via shared helper, so the
        // initial CreatePivotTable path and the post-Set RebuildFieldAreas path
        // produce identical results.
        var geom = ComputePivotGeometry(
            position, columnData, rowFieldIndices, colFieldIndices, valueFields);
        pivotDef.Location = new Location
        {
            Reference = geom.RangeRef,
            FirstHeaderRow = 1u,
            FirstDataRow = (colFieldIndices.Count >= 2 && valueFields.Count > 1) ? 4u
                         : ((valueFields.Count > 1 || colFieldIndices.Count >= 2) ? 3u : 2u),
            FirstDataColumn = (uint)geom.RowLabelCols
        };

        // Page filters: presence is signalled by the <pageFields> element + the
        // pivotField axis="axisPage" marker, both written further down. ECMA-376
        // also defines optional rowPageCount / colPageCount attributes here, but
        // OpenXml SDK 3.3.0 does not model them and rejects them as unknown
        // during schema validation. Excel recognizes the filter without them
        // (verified empirically and in pivot_dark1.xlsx, which has filters but
        // no page count attributes). Tracked as a v2 polish item if any consumer
        // turns out to require them.

        // PivotFields — one per source column
        var pivotFields = new PivotFields { Count = (uint)headers.Length };
        for (int i = 0; i < headers.Length; i++)
        {
            var pf = new PivotField { ShowAll = false };
            var values = i < columnData.Count ? columnData[i] : Array.Empty<string>();
            var isNumeric = values.Length > 0 && values.All(v =>
                string.IsNullOrEmpty(v) || double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _));

            if (rowFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisRow;
                if (!isNumeric) AppendFieldItems(pf, values);
            }
            else if (colFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisColumn;
                if (!isNumeric) AppendFieldItems(pf, values);
            }
            else if (filterFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisPage;
                if (!isNumeric) AppendFieldItems(pf, values);
            }
            else if (valueFields.Any(vf => vf.idx == i))
            {
                pf.DataField = true;
            }

            pivotFields.AppendChild(pf);
        }
        pivotDef.PivotFields = pivotFields;

        // RowFields — the synthetic <field x="-2"/> sentinel for multiple data
        // fields belongs to whichever axis (rows or columns) actually displays
        // the data field labels. The default is dataOnRows=false, so multi-data
        // labels go in COLUMNS — meaning the sentinel appears in colFields, NOT
        // rowFields. Only add the sentinel here when there are no col fields and
        // therefore data must flow in the row dimension.
        if (rowFieldIndices.Count > 0)
        {
            var rf = new RowFields();
            foreach (var idx in rowFieldIndices)
                rf.AppendChild(new Field { Index = idx });
            if (valueFields.Count > 1 && colFieldIndices.Count == 0)
                rf.AppendChild(new Field { Index = -2 });
            rf.Count = (uint)rf.Elements<Field>().Count();
            pivotDef.RowFields = rf;
        }

        // RowItems — describes the row-label layout. Without this, Excel renders only the
        // pivot's drop-down chrome but no actual data cells (the layout we observed earlier).
        // Pattern verified against LibreOffice's pivot_dark1.xlsx test fixture:
        //   <rowItems count="K+1">
        //     <i><x/></i>            <-- index 0 (shorthand: omit v attribute)
        //     <i><x v="1"/></i>      <-- index 1
        //     ...
        //     <i t="grand"><x/></i>  <-- grand total row
        //   </rowItems>
        // The <x v="N"/> values index into the corresponding pivotField's <items> list,
        // which we already populate via AppendFieldItems in BuildPivotTableDefinition above.
        // Single row field only: multi-row-field cartesian-product layout is a v2 concern.
        if (rowFieldIndices.Count > 0)
            pivotDef.RowItems = (RowItems)BuildAxisItems(rowFieldIndices, columnData, isRow: true, dataFieldCount: 1);

        // ColumnFields — when there are 2+ data fields, append the synthetic
        // <field x="-2"/> sentinel that tells Excel "data field labels go in
        // the column dimension here". Verified against multi_data_authored.xlsx:
        // a 1-row × 1-col × 2-data pivot writes <colFields count="2">
        // <field x="1"/><field x="-2"/></colFields>. Without this sentinel
        // Excel still opens the file but renders the K data fields stacked
        // incorrectly. RebuildFieldAreas already handles this; the initial
        // build path was missing the sentinel.
        if (colFieldIndices.Count > 0 || valueFields.Count > 1)
        {
            var cf = new ColumnFields();
            foreach (var idx in colFieldIndices)
                cf.AppendChild(new Field { Index = idx });
            if (valueFields.Count > 1)
                cf.AppendChild(new Field { Index = -2 });
            cf.Count = (uint)cf.Elements<Field>().Count();
            pivotDef.ColumnFields = cf;
        }

        // ColumnItems — same shape as RowItems but for the column-label layout.
        // Even when there are NO column fields, ECMA-376 requires a <colItems> with one
        // empty <i/> placeholder; LibreOffice's writeRowColumnItems empty-case branch
        // (xepivotxml.cxx:1008-1014) writes exactly that.
        pivotDef.ColumnItems = (ColumnItems)BuildAxisItems(
            colFieldIndices, columnData, isRow: false, dataFieldCount: valueFields.Count);

        // PageFields (filters)
        if (filterFieldIndices.Count > 0)
        {
            var pf = new PageFields { Count = (uint)filterFieldIndices.Count };
            foreach (var idx in filterFieldIndices)
                pf.AppendChild(new PageField { Field = idx, Hierarchy = -1 });
            pivotDef.PageFields = pf;
        }

        // DataFields
        if (valueFields.Count > 0)
        {
            var df = new DataFields { Count = (uint)valueFields.Count };
            foreach (var (idx, func, displayName) in valueFields)
            {
                // BaseField/BaseItem: Excel ignores these when ShowDataAs is normal,
                // but LibreOffice and Excel both emit them unconditionally on every
                // dataField (verified against pivot_dark1.xlsx and other LO fixtures).
                // Following the verified pattern rather than my earlier "omit them"
                // theory — being closer to what real producers write reduces the risk
                // of triggering picky consumers.
                var dataField = new DataField
                {
                    Name = displayName,
                    Field = (uint)idx,
                    Subtotal = ParseSubtotal(func),
                    BaseField = 0,
                    BaseItem = 0u
                };
                // Inherit the source column's numFmtId so Excel displays
                // pivot values using the same format as the source (currency,
                // percent, etc.). DataField.NumberFormatId is the primary
                // display driver — cell-level StyleIndex alone is ignored by
                // Excel for pivot values.
                if (columnNumFmtIds != null && idx >= 0 && idx < columnNumFmtIds.Length
                    && columnNumFmtIds[idx] is uint nfid)
                {
                    dataField.NumberFormatId = nfid;
                }
                df.AppendChild(dataField);
            }
            pivotDef.DataFields = df;
        }

        // Style
        pivotDef.PivotTableStyle = new PivotTableStyle
        {
            Name = styleName,
            ShowRowHeaders = true,
            ShowColumnHeaders = true,
            ShowRowStripes = false,
            ShowColumnStripes = false,
            ShowLastColumn = true
        };

        return pivotDef;
    }

    /// <summary>
    /// Build the &lt;rowItems&gt; or &lt;colItems&gt; layout block. Excel uses this to
    /// know how to expand row/column labels in the rendered pivot.
    ///
    /// Single data field (K=1):
    ///   <rowItems count="K+1">
    ///     <i><x/></i>            <-- index 0 (shorthand: omit v)
    ///     <i><x v="1"/></i>
    ///     ...
    ///     <i t="grand"><x/></i>
    ///   </rowItems>
    ///
    /// Multi-data field on the column axis (K>1, only used for ColumnItems):
    ///   <colItems count="(L+1)*K">
    ///     <i><x/><x/></i>                     <-- col label 0, data field 0
    ///     <i r="1" i="1"><x v="1"/></i>       <-- col label 0, data field 1 (r=1 = repeat prev x)
    ///     <i><x v="1"/><x/></i>               <-- col label 1, data field 0
    ///     <i r="1" i="1"><x v="1"/></i>       <-- col label 1, data field 1
    ///     ...
    ///     <i t="grand"><x/></i>               <-- grand total, data field 0
    ///     <i t="grand" i="1"><x/></i>         <-- grand total, data field 1
    ///   </colItems>
    /// Verified against multi_data_authored.xlsx (a 1×1×2 pivot from real Excel).
    ///
    /// Empty axis: single &lt;i/&gt; placeholder (LibreOffice writeRowColumnItems
    /// empty-case branch in xepivotxml.cxx:1008-1014).
    ///
    /// Limitation: still only single-axis-field cases are correct. Multi-row-field
    /// cartesian-product layouts need a deeper expansion tracked as v2.
    /// </summary>
    private static OpenXmlElement BuildAxisItems(
        List<int> fieldIndices, List<string[]> columnData, bool isRow, int dataFieldCount = 1)
    {
        OpenXmlCompositeElement container = isRow
            ? new RowItems()
            : new ColumnItems();

        // Empty axis: write a single empty <i/>. LibreOffice does this unconditionally
        // when there's nothing to render — Excel needs the placeholder. When there are
        // multiple data fields on the column axis but no col field, we still need
        // K entries (one per data field) instead of just one — handled below.
        if (fieldIndices.Count == 0)
        {
            if (!isRow && dataFieldCount > 1)
            {
                // Data-only column axis: K entries, each marked with i="d".
                for (int d = 0; d < dataFieldCount; d++)
                {
                    var item = new RowItem();
                    if (d > 0) item.Index = (uint)d;
                    item.AppendChild(new MemberPropertyIndex());
                    container.AppendChild(item);
                }
                SetAxisCount(container, dataFieldCount);
            }
            else
            {
                container.AppendChild(new RowItem());
                SetAxisCount(container, 1);
            }
            return container;
        }

        // N≥3 axis: route to tree-based items writer that uses LCP encoding
        // (longest common prefix) to compress arbitrary-depth path encoding.
        // Falls back to specialized N=2 path below for byte-level backward
        // compat with the regression baseline.
        if (fieldIndices.Count >= 3)
        {
            return BuildTreeAxisItems(fieldIndices, columnData, isRow, dataFieldCount);
        }

        // Multi-col case (N>=2 col fields, only used for ColumnItems).
        //
        // Pattern (verified against multi_col_authored.xlsx with cols=产品,包装):
        //   For each outer col value O:
        //     <i><x v="O"/><x v="0"/></i>           <- O + first inner (2 x children)
        //     For each subsequent inner I (sorted):
        //       <i r="1"><x v="I"/></i>             <- repeat outer, just give inner
        //     <i t="default"><x v="O"/></i>          <- O subtotal column
        //   <i t="grand"><x/></i>                   <- final grand total column
        //
        // Compared to BuildMultiRowItems: col subtotals use t="default" (not the
        // bare-<i> form rows use), and the leaf entries have 2 x children for
        // the first inner of each group instead of just 1.
        if (!isRow && fieldIndices.Count >= 2)
        {
            return BuildMultiColItems(fieldIndices, columnData, dataFieldCount);
        }

        // Multi-row case (N>=2 row fields, only used for RowItems).
        //
        // Pattern (verified against multi_row_authored.xlsx with 2 row fields,
        // where the user manually built a pivot with rows=地区,城市):
        //   For each outer value O in display order:
        //     <i><x v="O"/></i>                     <- outer subtotal row (1 x child)
        //     For each inner value I that exists in (O, *):
        //       <i r="1"><x v="I"/></i>             <- leaf row (r=1 = repeat outer)
        //   <i t="grand"><x/></i>                   <- final grand total
        //
        // The "1 x child only" form is treated by Excel as the outer-level
        // subtotal row (it shows aggregate across all this outer's inners). Leaf
        // rows use r='1' to mean "the first 1 member is inherited from the
        // previous row" (the outer index), so the leaf only needs its own inner
        // index as a single x child.
        //
        // This implementation supports exactly N=2 row fields. N>=3 would need a
        // recursive expansion at every non-leaf level — tracked as v4.
        if (isRow && fieldIndices.Count >= 2)
        {
            return BuildMultiRowItems(fieldIndices, columnData);
        }

        // Single field: one <i> per unique value, then a grand-total entry.
        // Multi-field is not yet supported — fall back to the first field's values
        // so the file is at least openable; rendering will be incomplete.
        var fieldIdx = fieldIndices[0];
        if (fieldIdx < 0 || fieldIdx >= columnData.Count)
        {
            container.AppendChild(new RowItem());
            SetAxisCount(container, 1);
            return container;
        }

        var uniqueCount = columnData[fieldIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .Count();

        // Multi-data on column axis: each col label gets K entries, then K grand totals.
        // The first entry per col label has TWO <x> children (col index + data field 0);
        // subsequent entries use r="1" to repeat the col index and bump i to the data
        // field number.
        if (!isRow && dataFieldCount > 1)
        {
            for (int i = 0; i < uniqueCount; i++)
            {
                // Entry for data field 0: <i><x v="i"/><x v="0"/></i>
                var first = new RowItem();
                if (i == 0)
                    first.AppendChild(new MemberPropertyIndex());
                else
                    first.AppendChild(new MemberPropertyIndex { Val = i });
                first.AppendChild(new MemberPropertyIndex());
                container.AppendChild(first);

                // Entries for data fields 1..K-1: <i r="1" i="d"><x v="d"/></i>
                for (int d = 1; d < dataFieldCount; d++)
                {
                    var rep = new RowItem
                    {
                        RepeatedItemCount = 1u,
                        Index = (uint)d
                    };
                    if (d == 0)
                        rep.AppendChild(new MemberPropertyIndex());
                    else
                        rep.AppendChild(new MemberPropertyIndex { Val = d });
                    container.AppendChild(rep);
                }
            }

            // Grand totals: K entries marked t="grand", with i=d for d>0.
            for (int d = 0; d < dataFieldCount; d++)
            {
                var gt = new RowItem { ItemType = ItemValues.Grand };
                if (d > 0) gt.Index = (uint)d;
                gt.AppendChild(new MemberPropertyIndex());
                container.AppendChild(gt);
            }

            SetAxisCount(container, uniqueCount * dataFieldCount + dataFieldCount);
            return container;
        }

        // Single-data layout (original path): K data rows + 1 grand total.
        for (int i = 0; i < uniqueCount; i++)
        {
            var item = new RowItem();
            if (i == 0)
                item.AppendChild(new MemberPropertyIndex());
            else
                item.AppendChild(new MemberPropertyIndex { Val = i });
            container.AppendChild(item);
        }

        // Grand total entry — always present in the default layout.
        var grandTotal = new RowItem { ItemType = ItemValues.Grand };
        grandTotal.AppendChild(new MemberPropertyIndex());
        container.AppendChild(grandTotal);

        SetAxisCount(container, uniqueCount + 1);
        return container;
    }

    /// <summary>
    /// Compute the (outer → ordered list of inners) groupings for a 2-row-field
    /// pivot. Only (outer, inner) combinations that actually appear in the
    /// source data are included — Excel does not enumerate empty cartesian
    /// cells in compact mode. Output is sorted by ordinal: outer keys first,
    /// then each outer's inner list. Used by both BuildMultiRowItems (XML
    /// rowItems generation) and the renderer (cell layout).
    /// </summary>
    private static List<(string outer, List<string> inners)> BuildOuterInnerGroups(
        int outerFieldIdx, int innerFieldIdx, List<string[]> columnData)
    {
        var outerVals = columnData[outerFieldIdx];
        var innerVals = columnData[innerFieldIdx];
        var n = outerVals.Length;

        var seen = new HashSet<(string, string)>();
        var combos = new List<(string outer, string inner)>();
        for (int i = 0; i < n; i++)
        {
            var ov = outerVals[i];
            var iv = innerVals[i];
            if (string.IsNullOrEmpty(ov) || string.IsNullOrEmpty(iv)) continue;
            if (seen.Add((ov, iv)))
                combos.Add((ov, iv));
        }

        // Sort using the active axis comparer so display order matches the
        // pivotField items list (which sorts via the same comparer). This
        // keeps rowItems indices in sync with rendered cell labels.
        return combos
            .GroupBy(c => c.outer, StringComparer.Ordinal)  // equality, not ordering
            .OrderByAxis(g => g.Key)
            .Select(g => (g.Key, g.Select(c => c.inner)
                .OrderByAxis(v => v).ToList()))
            .ToList();
    }

    /// <summary>
    /// Build the &lt;rowItems&gt; element for a 2-row-field pivot. Emits one
    /// outer-subtotal row per unique outer value plus one leaf row per
    /// (outer, inner) combination that exists in the data, then the grand
    /// total. See BuildOuterInnerGroups for the grouping logic.
    /// </summary>
    private static OpenXmlElement BuildMultiRowItems(
        List<int> fieldIndices, List<string[]> columnData)
    {
        var container = new RowItems();
        if (fieldIndices.Count < 2 || fieldIndices[0] >= columnData.Count || fieldIndices[1] >= columnData.Count)
        {
            container.AppendChild(new RowItem());
            container.Count = 1u;
            return container;
        }

        var outerIdx = fieldIndices[0];
        var innerIdx = fieldIndices[1];
        var groups = BuildOuterInnerGroups(outerIdx, innerIdx, columnData);

        // Pre-compute the value→pivotField-items-index map for both row fields.
        // The pivotField items list is built with StringComparer.Ordinal in
        // AppendFieldItems below, so we mirror the same ordering here to keep
        // the indices consistent.
        var outerOrder = columnData[outerIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .OrderByAxis(v => v)
            .Select((v, i) => (v, i))
            .ToDictionary(t => t.v, t => t.i, StringComparer.Ordinal);
        var innerOrder = columnData[innerIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .OrderByAxis(v => v)
            .Select((v, i) => (v, i))
            .ToDictionary(t => t.v, t => t.i, StringComparer.Ordinal);

        int count = 0;
        foreach (var (outer, inners) in groups)
        {
            // Outer subtotal row: <i><x v="outerIdx"/></i>
            var outerEntry = new RowItem();
            var outerPivIdx = outerOrder[outer];
            if (outerPivIdx == 0)
                outerEntry.AppendChild(new MemberPropertyIndex());
            else
                outerEntry.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
            container.AppendChild(outerEntry);
            count++;

            // Leaf rows for each inner of this outer: <i r="1"><x v="innerIdx"/></i>
            foreach (var inner in inners)
            {
                var leafEntry = new RowItem { RepeatedItemCount = 1u };
                var innerPivIdx = innerOrder[inner];
                if (innerPivIdx == 0)
                    leafEntry.AppendChild(new MemberPropertyIndex());
                else
                    leafEntry.AppendChild(new MemberPropertyIndex { Val = innerPivIdx });
                container.AppendChild(leafEntry);
                count++;
            }
        }

        // Grand total row.
        var grand = new RowItem { ItemType = ItemValues.Grand };
        grand.AppendChild(new MemberPropertyIndex());
        container.AppendChild(grand);
        count++;

        container.Count = (uint)count;
        return container;
    }

    /// <summary>
    /// Build the &lt;colItems&gt; element for a 2-col-field pivot, supporting K
    /// data fields. Mirrors BuildMultiRowItems but uses the col-subtotal
    /// pattern (t="default") instead of the bare-i form rows use, and the
    /// first leaf of each outer group emits 2 x children (outer + inner).
    ///
    /// For K&gt;1 (multi-col + multi-data, e.g. 1×2×2), each leaf and each
    /// subtotal/grand-total entry is multiplied by K, with the additional
    /// data field entries using r='2' (repeat outer + inner) and i='d' to
    /// flag the data field index. Verified against multi_col_K_authored.xlsx.
    /// </summary>
    private static OpenXmlElement BuildMultiColItems(
        List<int> fieldIndices, List<string[]> columnData, int dataFieldCount)
    {
        var container = new ColumnItems();
        if (fieldIndices.Count < 2 || fieldIndices[0] >= columnData.Count || fieldIndices[1] >= columnData.Count)
        {
            container.AppendChild(new RowItem());
            container.Count = 1u;
            return container;
        }

        var outerIdx = fieldIndices[0];
        var innerIdx = fieldIndices[1];
        var groups = BuildOuterInnerGroups(outerIdx, innerIdx, columnData);

        // Value → pivotField-items-index map (alphabetical ordinal sort).
        var outerOrder = columnData[outerIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .OrderByAxis(v => v)
            .Select((v, i) => (v, i))
            .ToDictionary(t => t.v, t => t.i, StringComparer.Ordinal);
        var innerOrder = columnData[innerIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .OrderByAxis(v => v)
            .Select((v, i) => (v, i))
            .ToDictionary(t => t.v, t => t.i, StringComparer.Ordinal);

        int K = Math.Max(1, dataFieldCount);
        int count = 0;
        foreach (var (outer, inners) in groups)
        {
            var outerPivIdx = outerOrder[outer];

            for (int idx = 0; idx < inners.Count; idx++)
            {
                var inner = inners[idx];
                var innerPivIdx = innerOrder[inner];

                // First leaf of (this outer, this inner): K entries (one per data field).
                // The very first entry has the full path; subsequent K-1 use r=2 (repeat
                // outer + inner) to compress the encoding.
                for (int d = 0; d < K; d++)
                {
                    if (d == 0)
                    {
                        // First data field: full path.
                        // For new outer (idx==0): 2 or 3 x children (outer + inner + maybe d).
                        //   With K==1: just outer + inner = 2 x children.
                        //   With K>1: outer + inner + first data = 3 x children.
                        // For new inner (idx>0) with new outer leaf area: r=1 (repeat outer)
                        //   With K==1: r=1, then inner = 1 x child total.
                        //   With K>1: r=1, then inner + first data = 2 x children.
                        if (idx == 0)
                        {
                            // First leaf of new outer: write everything fresh.
                            var first = new RowItem();
                            if (outerPivIdx == 0) first.AppendChild(new MemberPropertyIndex());
                            else first.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
                            if (innerPivIdx == 0) first.AppendChild(new MemberPropertyIndex());
                            else first.AppendChild(new MemberPropertyIndex { Val = innerPivIdx });
                            if (K > 1)
                            {
                                // First data field index = 0 → bare <x/>
                                first.AppendChild(new MemberPropertyIndex());
                            }
                            container.AppendChild(first);
                        }
                        else
                        {
                            // Inner shift within same outer: r=1 keeps outer.
                            var rep = new RowItem { RepeatedItemCount = 1u };
                            if (innerPivIdx == 0) rep.AppendChild(new MemberPropertyIndex());
                            else rep.AppendChild(new MemberPropertyIndex { Val = innerPivIdx });
                            if (K > 1) rep.AppendChild(new MemberPropertyIndex());
                            container.AppendChild(rep);
                        }
                    }
                    else
                    {
                        // Additional data field for the same (outer, inner): r=2 keeps
                        // outer + inner, i=d marks the data field, x v=d gives the index.
                        var rep = new RowItem { RepeatedItemCount = 2u, Index = (uint)d };
                        if (d == 0) rep.AppendChild(new MemberPropertyIndex());
                        else rep.AppendChild(new MemberPropertyIndex { Val = d });
                        container.AppendChild(rep);
                    }
                    count++;
                }
            }

            // Outer subtotal columns: K entries with t="default", x v=outer, i=d for d>0.
            for (int d = 0; d < K; d++)
            {
                var sub = new RowItem { ItemType = ItemValues.Default };
                if (d > 0) sub.Index = (uint)d;
                if (outerPivIdx == 0) sub.AppendChild(new MemberPropertyIndex());
                else sub.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
                container.AppendChild(sub);
                count++;
            }
        }

        // Grand total columns: K entries with t="grand", x=0, i=d for d>0.
        for (int d = 0; d < K; d++)
        {
            var grand = new RowItem { ItemType = ItemValues.Grand };
            if (d > 0) grand.Index = (uint)d;
            grand.AppendChild(new MemberPropertyIndex());
            container.AppendChild(grand);
            count++;
        }

        container.Count = (uint)count;
        return container;
    }

    /// <summary>
    /// Generic axis-items writer for N≥3 row or col fields. Walks the AxisTree
    /// in display order and emits RowItem entries with longest-common-prefix
    /// (LCP) compression for the &lt;i r="K"&gt; repeat attribute.
    ///
    /// Pattern (verified by extending the N=2 patterns recursively):
    ///   - Each entry has 1 logical "path" of length = entry depth (subtotals
    ///     have shorter paths than leaves).
    ///   - r = LCP(this.path, prev.path). x children = path elements after the LCP.
    ///   - For N=2 cases this naturally collapses to the existing
    ///     BuildMultiRowItems / BuildMultiColItems output (verified by hand).
    ///   - Row axis: subtotals are bare &lt;i&gt; entries. They sit BEFORE their
    ///     children in walk order.
    ///   - Col axis: subtotals are &lt;i t="default"&gt; entries that always emit
    ///     r=0 + 1 x child for the path's last (and only) element. They sit
    ///     AFTER their children in walk order. This matches the empirical
    ///     observation that Excel "resets" the inheritance chain at every
    ///     col-axis subtotal.
    ///   - Grand total: &lt;i t="grand"&gt; with bare &lt;x/&gt;, always r=0.
    ///
    /// For K>1 on the column axis, each logical entry (leaf, subtotal, grand)
    /// is multiplied by K, mirroring the BuildMultiColItems pattern:
    ///   - Leaf d=0: LCP-compressed path + 1 extra &lt;x/&gt; for data field 0.
    ///   - Leaf d∈[1,K): r=path.Length, i=d, 1 &lt;x v=d/&gt;. (The whole
    ///     non-data path is inherited from d=0; i=d flags this as "same
    ///     cell position, different data field".)
    ///   - Subtotal d=0: as in K=1 (r=0 + 1 x child for path[last]).
    ///   - Subtotal d∈[1,K): same x child, add i=d attribute.
    ///   - Grand d=0: bare &lt;x/&gt;. Grand d∈[1,K): bare &lt;x/&gt; + i=d.
    /// Row axis is never K-multiplied regardless of K — verified against
    /// 2x1x1 vs 2x1xK baselines where rowItems.count is identical.
    /// </summary>
    private static OpenXmlElement BuildTreeAxisItems(
        List<int> fieldIndices, List<string[]> columnData, bool isRow, int dataFieldCount)
    {
        var container = isRow
            ? (OpenXmlCompositeElement)new RowItems()
            : new ColumnItems();

        var tree = BuildAxisTree(fieldIndices, columnData);

        // Pre-compute per-level value→index maps so the emitted <x v="N"/>
        // references match the corresponding pivotField items list (which
        // we sort with StringComparer.Ordinal in AppendFieldItems).
        var perLevelOrder = new Dictionary<string, int>[fieldIndices.Count];
        for (int level = 0; level < fieldIndices.Count; level++)
        {
            var fi = fieldIndices[level];
            if (fi < 0 || fi >= columnData.Count) { perLevelOrder[level] = new Dictionary<string, int>(); continue; }
            perLevelOrder[level] = columnData[fi]
                .Where(v => !string.IsNullOrEmpty(v))
                .Distinct()
                .OrderByAxis(v => v)
                .Select((v, i) => (v, i))
                .ToDictionary(t => t.v, t => t.i, StringComparer.Ordinal);
        }

        // Collect entries by walking the tree in display order. Each entry is a
        // (path, type) pair where type ∈ {leaf, subtotal, grand}.
        var entries = new List<(string[] path, string kind)>(); // kind: "leaf" | "subtotal" | "grand"
        void Walk(AxisNode node)
        {
            if (node.IsLeaf)
            {
                entries.Add((node.Path, "leaf"));
                return;
            }
            // Skip the synthetic root (Depth=0).
            if (!isRow && node.Depth > 0)
            {
                // Col axis: children before subtotal.
                foreach (var c in node.Children) Walk(c);
                entries.Add((node.Path, "subtotal"));
            }
            else if (isRow && node.Depth > 0)
            {
                // Row axis: subtotal before children.
                entries.Add((node.Path, "subtotal"));
                foreach (var c in node.Children) Walk(c);
            }
            else
            {
                // Synthetic root, just recurse.
                foreach (var c in node.Children) Walk(c);
            }
        }
        Walk(tree);
        entries.Add((Array.Empty<string>(), "grand"));

        // K>1 multiplies col-axis entries by K (one per data field). Row axis
        // stays 1 entry per logical row regardless of K.
        int K = Math.Max(1, dataFieldCount);
        bool kMultiply = !isRow && K > 1;

        // Emit entries with LCP compression. Col-axis subtotals are special-cased
        // to always emit r=0 + 1 x child for the outer index (Excel's empirical
        // convention — col subtotals "reset" the inheritance chain).
        string[] prevPath = Array.Empty<string>();
        int emittedCount = 0;
        foreach (var (path, kind) in entries)
        {
            if (kind == "grand")
            {
                // K entries on col axis, 1 entry on row axis. Each is a bare
                // <x/> (v=0), with i=d on d∈[1,K) for col axis.
                int grandCount = kMultiply ? K : 1;
                for (int d = 0; d < grandCount; d++)
                {
                    var gt = new RowItem { ItemType = ItemValues.Grand };
                    if (d > 0) gt.Index = (uint)d;
                    gt.AppendChild(new MemberPropertyIndex());
                    container.AppendChild(gt);
                    emittedCount++;
                }
                prevPath = path;
                continue;
            }

            if (kind == "subtotal" && !isRow)
            {
                // Col-axis subtotal: always r=0 + 1 x child for the deepest
                // index in the path (the immediate-parent value). Verified
                // against multi_col_authored.xlsx. For K>1, emit K of these
                // with i=d attribute on d∈[1,K).
                int lastLevel = path.Length - 1;
                int lastIdx = perLevelOrder[lastLevel].TryGetValue(path[lastLevel], out var li) ? li : 0;
                for (int d = 0; d < K; d++)
                {
                    var sub = new RowItem { ItemType = ItemValues.Default };
                    if (d > 0) sub.Index = (uint)d;
                    if (lastIdx == 0) sub.AppendChild(new MemberPropertyIndex());
                    else sub.AppendChild(new MemberPropertyIndex { Val = lastIdx });
                    container.AppendChild(sub);
                    emittedCount++;
                }
                // Reset prev so the next entry doesn't try to inherit through
                // the subtotal's truncated path. The next leaf in a new outer
                // group will write a fresh path from r=0.
                prevPath = path;
                continue;
            }

            // Leaf entries (both row and col) and row subtotals use LCP encoding.
            var item = new RowItem();
            int lcp = 0;
            while (lcp < path.Length && lcp < prevPath.Length && path[lcp] == prevPath[lcp]) lcp++;
            if (lcp > 0) item.RepeatedItemCount = (uint)lcp;
            for (int i = lcp; i < path.Length; i++)
            {
                int idx = perLevelOrder[i].TryGetValue(path[i], out var pi) ? pi : 0;
                if (idx == 0) item.AppendChild(new MemberPropertyIndex());
                else item.AppendChild(new MemberPropertyIndex { Val = idx });
            }
            // For col-axis leaves with K>1, append one extra <x/> for the
            // first data field (index 0 = bare <x/>). The K-1 subsequent
            // entries below handle the remaining data fields.
            if (kMultiply && kind == "leaf")
            {
                item.AppendChild(new MemberPropertyIndex());
            }
            // Defensive: an entry with no x children (e.g. an empty path with
            // no LCP slack) would be malformed. Always ensure at least one.
            if (!item.Elements<MemberPropertyIndex>().Any())
                item.AppendChild(new MemberPropertyIndex());

            container.AppendChild(item);
            emittedCount++;

            // K>1 col-axis leaf: emit K-1 more entries that inherit the full
            // path (r=path.Length) and carry i=d to mark the data field.
            if (kMultiply && kind == "leaf")
            {
                for (int d = 1; d < K; d++)
                {
                    var rep = new RowItem
                    {
                        RepeatedItemCount = (uint)path.Length,
                        Index = (uint)d
                    };
                    rep.AppendChild(new MemberPropertyIndex { Val = d });
                    container.AppendChild(rep);
                    emittedCount++;
                }
            }

            prevPath = path;
        }

        SetAxisCount(container, emittedCount);
        return container;
    }

    /// <summary>Set the count attribute on RowItems / ColumnItems uniformly.</summary>
    private static void SetAxisCount(OpenXmlCompositeElement container, int count)
    {
        if (container is RowItems ri) ri.Count = (uint)count;
        else if (container is ColumnItems ci) ci.Count = (uint)count;
    }

    private static void AppendFieldItems(PivotField pf, string[] values)
    {
        var unique = values.Where(v => !string.IsNullOrEmpty(v)).Distinct().OrderByAxis(v => v).ToList();
        var items = new Items { Count = (uint)(unique.Count + 1) };
        for (int i = 0; i < unique.Count; i++)
            items.AppendChild(new Item { Index = (uint)i });
        items.AppendChild(new Item { ItemType = ItemValues.Default }); // grand total
        pf.AppendChild(items);
    }

    // ==================== Readback ====================

    internal static void ReadPivotTableProperties(PivotTableDefinition pivotDef, DocumentNode node)
    {
        if (pivotDef.Name?.HasValue == true) node.Format["name"] = pivotDef.Name.Value;
        if (pivotDef.CacheId?.HasValue == true) node.Format["cacheId"] = pivotDef.CacheId.Value;

        var location = pivotDef.GetFirstChild<Location>();
        if (location?.Reference?.HasValue == true) node.Format["location"] = location.Reference.Value;

        // Count fields
        var pivotFields = pivotDef.GetFirstChild<PivotFields>();
        if (pivotFields != null)
            node.Format["fieldCount"] = pivotFields.Elements<PivotField>().Count();

        // Row fields
        var rowFields = pivotDef.RowFields;
        if (rowFields != null)
        {
            var indices = rowFields.Elements<Field>().Where(f => f.Index?.Value >= 0).Select(f => f.Index!.Value).ToList();
            if (indices.Count > 0)
                node.Format["rowFields"] = string.Join(",", indices);
        }

        // Column fields
        var colFields = pivotDef.ColumnFields;
        if (colFields != null)
        {
            var indices = colFields.Elements<Field>().Where(f => f.Index?.Value >= 0).Select(f => f.Index!.Value).ToList();
            if (indices.Count > 0)
                node.Format["colFields"] = string.Join(",", indices);
        }

        // Page/filter fields
        var pageFields = pivotDef.PageFields;
        if (pageFields != null)
        {
            var indices = pageFields.Elements<PageField>().Select(f => f.Field?.Value ?? -1).Where(v => v >= 0).ToList();
            if (indices.Count > 0)
                node.Format["filterFields"] = string.Join(",", indices);
        }

        // Data fields (use typed property for reliable access)
        var dataFields = pivotDef.DataFields;
        if (dataFields != null)
        {
            var dfList = dataFields.Elements<DataField>().ToList();
            node.Format["dataFieldCount"] = dfList.Count;
            for (int i = 0; i < dfList.Count; i++)
            {
                var df = dfList[i];
                var dfName = df.Name?.Value ?? "";
                var dfFunc = df.Subtotal?.InnerText ?? "sum";
                var dfField = df.Field?.Value ?? 0;
                node.Format[$"dataField{i + 1}"] = $"{dfName}:{dfFunc}:{dfField}";
            }
        }

        // Style
        var styleInfo = pivotDef.PivotTableStyle;
        if (styleInfo?.Name?.HasValue == true)
            node.Format["style"] = styleInfo.Name.Value;
    }

    internal static List<string> SetPivotTableProperties(PivotTablePart pivotPart, Dictionary<string, string> properties)
    {
        // Publish sort mode for this Set operation so the re-rendered items /
        // renderers use the requested order. Sort only affects the rendered
        // layout — sharedItems order in the cache is fixed at Create time.
        using var _sortScope = PushAxisSortMode(properties);

        var unsupported = new List<string>();
        var pivotDef = pivotPart.PivotTableDefinition;
        if (pivotDef == null) { unsupported.AddRange(properties.Keys); return unsupported; }

        // Collect field-area properties separately — they require a coordinated rebuild
        var fieldAreaProps = new Dictionary<string, string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    pivotDef.Name = value;
                    break;
                case "style":
                {
                    pivotDef.PivotTableStyle = new PivotTableStyle
                    {
                        Name = value,
                        ShowRowHeaders = true,
                        ShowColumnHeaders = true,
                        ShowRowStripes = false,
                        ShowColumnStripes = false,
                        ShowLastColumn = true
                    };
                    break;
                }
                case "rows":
                case "cols" or "columns":
                case "values":
                case "filters":
                    fieldAreaProps[key.ToLowerInvariant() == "columns" ? "cols" : key.ToLowerInvariant()] = value;
                    break;
                case "sort":
                    // Already consumed by PushAxisSortMode at the top of this
                    // method; re-rendering below reads _axisSortMode directly.
                    // Trigger a re-render even if no field areas changed so
                    // the layout reflects the new sort.
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters"))
                    {
                        // Seed an empty entry so RebuildFieldAreas runs with
                        // current field assignments and re-renders with the
                        // new sort.
                        fieldAreaProps["__sort_only__"] = value;
                    }
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        // If any field areas were specified, rebuild them
        if (fieldAreaProps.Count > 0)
            RebuildFieldAreas(pivotPart, pivotDef, fieldAreaProps);

        pivotDef.Save();
        return unsupported;
    }

    /// <summary>
    /// Rebuild pivot table field areas (rows, cols, values, filters).
    /// For areas not specified in changes, preserves the current assignment.
    /// Two-layer update: (1) PivotField.Axis/DataField, (2) RowFields/ColumnFields/PageFields/DataFields.
    /// </summary>
    private static void RebuildFieldAreas(PivotTablePart pivotPart, PivotTableDefinition pivotDef,
        Dictionary<string, string> changes)
    {
        // Get headers from cache definition
        var cachePart = pivotPart.GetPartsOfType<PivotTableCacheDefinitionPart>().FirstOrDefault();
        if (cachePart?.PivotCacheDefinition == null) return;

        var cacheFields = cachePart.PivotCacheDefinition.GetFirstChild<CacheFields>();
        if (cacheFields == null) return;

        var headers = cacheFields.Elements<CacheField>().Select(cf => cf.Name?.Value ?? "").ToArray();
        if (headers.Length == 0) return;

        // Read current assignments for areas NOT being changed
        var currentRows = ReadCurrentFieldIndices(pivotDef.RowFields?.Elements<Field>(), f => f.Index?.Value ?? -1);
        var currentCols = ReadCurrentFieldIndices(pivotDef.ColumnFields?.Elements<Field>(), f => f.Index?.Value ?? -1);
        var currentFilters = ReadCurrentFieldIndices(pivotDef.PageFields?.Elements<PageField>(), f => f.Field?.Value ?? -1);
        var currentValues = ReadCurrentDataFields(pivotDef.DataFields);

        // Parse new assignments (or keep current)
        // If user specified a non-empty value but nothing resolved, warn via stderr
        var rowFieldIndices = changes.ContainsKey("rows")
            ? ParseFieldListWithWarning(changes, "rows", headers)
            : currentRows;
        var colFieldIndices = changes.ContainsKey("cols")
            ? ParseFieldListWithWarning(changes, "cols", headers)
            : currentCols;
        var filterFieldIndices = changes.ContainsKey("filters")
            ? ParseFieldListWithWarning(changes, "filters", headers)
            : currentFilters;
        var valueFields = changes.ContainsKey("values")
            ? ParseValueFieldsWithWarning(changes, "values", headers)
            : currentValues;

        // Layer 1: Reset all PivotField axis/dataField, then re-assign
        var pivotFields = pivotDef.PivotFields;
        if (pivotFields == null) return;

        var pfList = pivotFields.Elements<PivotField>().ToList();
        for (int i = 0; i < pfList.Count; i++)
        {
            var pf = pfList[i];
            // Clear axis and dataField
            pf.Axis = null;
            pf.DataField = null;
            pf.RemoveAllChildren<Items>();

            // Determine if this field's cache data is numeric (for Items generation)
            var isNumeric = IsFieldNumeric(cacheFields, i);

            if (rowFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisRow;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
            }
            else if (colFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisColumn;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
            }
            else if (filterFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisPage;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
            }
            else if (valueFields.Any(vf => vf.idx == i))
            {
                pf.DataField = true;
            }
        }

        // Layer 2: Rebuild area reference lists
        // RowFields
        if (rowFieldIndices.Count > 0)
        {
            var rf = new RowFields { Count = (uint)rowFieldIndices.Count };
            foreach (var idx in rowFieldIndices)
                rf.AppendChild(new Field { Index = idx });
            // -2 sentinel for multiple value fields displayed in rows
            if (valueFields.Count > 1 && colFieldIndices.Count == 0)
            {
                rf.AppendChild(new Field { Index = -2 });
                rf.Count = (uint)rf.Elements<Field>().Count();
            }
            pivotDef.RowFields = rf;
        }
        else
        {
            pivotDef.RowFields = null;
        }

        // ColumnFields
        if (colFieldIndices.Count > 0 || valueFields.Count > 1)
        {
            var cf = new ColumnFields();
            foreach (var idx in colFieldIndices)
                cf.AppendChild(new Field { Index = idx });
            // -2 sentinel for multiple value fields in columns
            if (valueFields.Count > 1)
                cf.AppendChild(new Field { Index = -2 });
            cf.Count = (uint)cf.Elements<Field>().Count();
            pivotDef.ColumnFields = cf;
        }
        else
        {
            pivotDef.ColumnFields = null;
        }

        // PageFields (filters)
        if (filterFieldIndices.Count > 0)
        {
            var pf = new PageFields { Count = (uint)filterFieldIndices.Count };
            foreach (var idx in filterFieldIndices)
                pf.AppendChild(new PageField { Field = idx, Hierarchy = -1 });
            pivotDef.PageFields = pf;
        }
        else
        {
            pivotDef.PageFields = null;
        }

        // Re-read the source sheet's column styles so both (a) the DataField's
        // NumberFormatId (Excel's primary pivot-value display driver) and
        // (b) the value-cell StyleIndex stay in sync with the source column's
        // currency/percent/custom format across Set operations.
        uint?[]? sourceColumnStyleIds = null;
        uint?[]? sourceColumnNumFmtIds = null;
        var wbPart = pivotPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault()
            ?.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
        var wsSource = cachePart.PivotCacheDefinition.CacheSource?.WorksheetSource;
        if (wbPart != null && wsSource?.Sheet?.Value is string srcSheetName
            && wsSource.Reference?.Value is string srcRef)
        {
            var sheetRef = wbPart.Workbook?.Sheets?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value == srcSheetName);
            if (sheetRef?.Id?.Value is string relId
                && wbPart.GetPartById(relId) is WorksheetPart srcWsPart)
            {
                try
                {
                    var (_, _, ids) = ReadSourceData(srcWsPart, srcRef);
                    sourceColumnStyleIds = ids;
                    sourceColumnNumFmtIds = ResolveColumnNumFmtIds(wbPart, ids);
                }
                catch { /* best-effort: Set still succeeds with General format */ }
            }
        }

        // DataFields
        if (valueFields.Count > 0)
        {
            var df = new DataFields { Count = (uint)valueFields.Count };
            foreach (var (idx, func, displayName) in valueFields)
            {
                // BaseField/BaseItem: Excel ignores these when ShowDataAs is normal,
                // but LibreOffice and Excel both emit them unconditionally on every
                // dataField (verified against pivot_dark1.xlsx and other LO fixtures).
                // Following the verified pattern rather than my earlier "omit them"
                // theory — being closer to what real producers write reduces the risk
                // of triggering picky consumers.
                var dataField = new DataField
                {
                    Name = displayName,
                    Field = (uint)idx,
                    Subtotal = ParseSubtotal(func),
                    BaseField = 0,
                    BaseItem = 0u
                };
                if (sourceColumnNumFmtIds != null && idx >= 0 && idx < sourceColumnNumFmtIds.Length
                    && sourceColumnNumFmtIds[idx] is uint nfid)
                {
                    dataField.NumberFormatId = nfid;
                }
                df.AppendChild(dataField);
            }
            pivotDef.DataFields = df;
        }
        else
        {
            pivotDef.DataFields = null;
        }

        // Update Location with the full new geometry — range, offsets, FirstDataCol —
        // not just FirstDataColumn. The previous incremental approach left a stale
        // range covering the old layout, which made Excel render only the original
        // bounds even when fields were added or removed.
        var oldLocation = pivotDef.Location;
        var oldRangeRef = oldLocation?.Reference?.Value;
        var anchorRefForGeometry = oldRangeRef?.Split(':')[0]
            ?? oldLocation?.Reference?.Value
            ?? "A1";

        // Reconstruct columnData from the cache so the geometry helper and the
        // renderer below can compute new extents without re-reading the source sheet.
        var (cacheHeaders, cacheColumnData) = ReadColumnDataFromCache(
            cachePart.PivotCacheDefinition,
            cachePart.GetPartsOfType<PivotTableCacheRecordsPart>().FirstOrDefault()?.PivotCacheRecords);

        var newGeom = ComputePivotGeometry(
            anchorRefForGeometry, cacheColumnData, rowFieldIndices, colFieldIndices, valueFields);

        pivotDef.Location = new Location
        {
            Reference = newGeom.RangeRef,
            FirstHeaderRow = 1u,
            FirstDataRow = 2u,
            FirstDataColumn = (uint)newGeom.RowLabelCols
        };

        // Rebuild RowItems / ColumnItems for the new field assignments. The previous
        // configuration's row/col layout no longer matches; without these the rendered
        // skeleton would still describe the old shape.
        if (rowFieldIndices.Count > 0)
            pivotDef.RowItems = (RowItems)BuildAxisItems(rowFieldIndices, cacheColumnData, isRow: true, dataFieldCount: 1);
        else
            pivotDef.RowItems = null;
        pivotDef.ColumnItems = (ColumnItems)BuildAxisItems(
            colFieldIndices, cacheColumnData, isRow: false, dataFieldCount: valueFields.Count);

        // Refresh caption attributes — they pin to the row/col field's header name,
        // so reassigning fields means the visible caption changes too.
        pivotDef.RowHeaderCaption = rowFieldIndices.Count > 0 ? cacheHeaders[rowFieldIndices[0]] : "Rows";
        pivotDef.ColumnHeaderCaption = colFieldIndices.Count > 0 ? cacheHeaders[colFieldIndices[0]] : "Columns";

        // Re-render the materialized cells. Find the host worksheet via the pivot
        // part's parent — pivotPart is owned by exactly one WorksheetPart so this
        // is unambiguous in v1 (no shared pivot tables).
        var hostSheet = pivotPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault();
        if (hostSheet != null)
        {
            var ws = hostSheet.Worksheet;
            var sheetData = ws?.GetFirstChild<SheetData>();
            if (ws != null && sheetData != null)
            {
                // Clear the OLD rendered cells before drawing the new layout. The
                // new geometry might be smaller (fewer cols → stale right-hand cells)
                // OR larger (more rows → safe overwrite), so we always wipe the union
                // of old and new bounds. Old range first, then new range — the new
                // render writes into the cleared area immediately after.
                if (!string.IsNullOrEmpty(oldRangeRef))
                    ClearPivotRangeCells(sheetData, oldRangeRef);
                ClearPivotRangeCells(sheetData, newGeom.RangeRef);

                RenderPivotIntoSheet(
                    hostSheet, anchorRefForGeometry, cacheHeaders, cacheColumnData,
                    rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices,
                    sourceColumnStyleIds);
            }
        }
    }

    private static List<int> ReadCurrentFieldIndices<T>(IEnumerable<T>? elements, Func<T, int> getIndex)
    {
        if (elements == null) return new List<int>();
        return elements.Select(getIndex).Where(i => i >= 0).ToList();
    }

    private static List<(int idx, string func, string name)> ReadCurrentDataFields(DataFields? dataFields)
    {
        if (dataFields == null) return new List<(int, string, string)>();
        return dataFields.Elements<DataField>().Select(df => (
            idx: (int)(df.Field?.Value ?? 0),
            func: df.Subtotal?.InnerText ?? "sum",
            name: df.Name?.Value ?? ""
        )).ToList();
    }

    private static bool IsFieldNumeric(CacheFields cacheFields, int index)
    {
        var cf = cacheFields.Elements<CacheField>().ElementAtOrDefault(index);
        var sharedItems = cf?.GetFirstChild<SharedItems>();
        if (sharedItems == null) return false;
        return sharedItems.ContainsNumber?.Value == true && sharedItems.ContainsString?.Value != true;
    }

    private static void AppendFieldItemsFromCache(PivotField pf, CacheFields cacheFields, int index)
    {
        var cf = cacheFields.Elements<CacheField>().ElementAtOrDefault(index);
        var sharedItems = cf?.GetFirstChild<SharedItems>();
        var count = sharedItems?.Elements<StringItem>().Count() ?? 0;
        if (count == 0) return;

        var items = new Items { Count = (uint)(count + 1) };
        for (int i = 0; i < count; i++)
            items.AppendChild(new Item { Index = (uint)i });
        items.AppendChild(new Item { ItemType = ItemValues.Default }); // grand total
        pf.AppendChild(items);
    }

    // ==================== Parse Helpers ====================

    private static List<int> ParseFieldListWithWarning(Dictionary<string, string> props, string key, string[] headers)
    {
        var result = ParseFieldList(props, key, headers);
        if (result.Count == 0 && props.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value))
        {
            var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
            Console.Error.WriteLine($"WARNING: No matching fields for {key}={value}. Available: {available}");
        }
        return result;
    }

    private static List<(int idx, string func, string name)> ParseValueFieldsWithWarning(
        Dictionary<string, string> props, string key, string[] headers)
    {
        var result = ParseValueFields(props, key, headers);
        if (result.Count == 0 && props.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value))
        {
            var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
            Console.Error.WriteLine($"WARNING: No matching fields for {key}={value}. Available: {available}");
        }
        return result;
    }

    private static List<int> ParseFieldList(Dictionary<string, string> props, string key, string[] headers)
    {
        if (!props.TryGetValue(key, out var value) || string.IsNullOrEmpty(value))
            return new List<int>();

        return value.Split(',').Select(f =>
        {
            var name = f.Trim();
            // Try as column index first
            if (int.TryParse(name, out var idx)) return idx;
            // Try as header name
            for (int i = 0; i < headers.Length; i++)
                if (headers[i] != null && headers[i].Equals(name, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }).Where(i => i >= 0 && i < headers.Length).ToList();
    }

    private static List<(int idx, string func, string name)> ParseValueFields(
        Dictionary<string, string> props, string key, string[] headers)
    {
        if (!props.TryGetValue(key, out var value) || string.IsNullOrEmpty(value))
            return new List<(int, string, string)>();

        var result = new List<(int idx, string func, string name)>();
        foreach (var spec in value.Split(','))
        {
            // Format: "FieldName:func" or "FieldName" (default sum)
            var parts = spec.Trim().Split(':');
            var fieldName = parts[0].Trim();
            var func = parts.Length > 1 ? parts[1].Trim().ToLowerInvariant() : "sum";

            int fieldIdx = -1;
            if (int.TryParse(fieldName, out var idx)) fieldIdx = idx;
            else
            {
                for (int i = 0; i < headers.Length; i++)
                    if (headers[i] != null && headers[i].Equals(fieldName, StringComparison.OrdinalIgnoreCase)) { fieldIdx = i; break; }
            }

            if (fieldIdx >= 0 && fieldIdx < headers.Length)
            {
                var displayName = $"{char.ToUpper(func[0])}{func[1..]} of {headers[fieldIdx]}";
                result.Add((fieldIdx, func, displayName));
            }
        }
        return result;
    }

    private static DataConsolidateFunctionValues ParseSubtotal(string func)
    {
        return func.ToLowerInvariant() switch
        {
            "sum" => DataConsolidateFunctionValues.Sum,
            "count" => DataConsolidateFunctionValues.Count,
            "countnums" or "countnum" => DataConsolidateFunctionValues.CountNumbers,
            "average" or "avg" => DataConsolidateFunctionValues.Average,
            "max" => DataConsolidateFunctionValues.Maximum,
            "min" => DataConsolidateFunctionValues.Minimum,
            "product" => DataConsolidateFunctionValues.Product,
            "stddev" or "std" => DataConsolidateFunctionValues.StandardDeviation,
            "stddevp" or "stdp" => DataConsolidateFunctionValues.StandardDeviationP,
            "var" or "variance" => DataConsolidateFunctionValues.Variance,
            "varp" => DataConsolidateFunctionValues.VarianceP,
            _ => DataConsolidateFunctionValues.Sum
        };
    }

    /// <summary>
    /// Aggregate a bag of numeric values using the given subtotal function.
    /// Matches LibreOffice's ScDPAggData semantics (sc/source/core/data/dptabres.cxx):
    ///   sum / product / min / max / count : trivial
    ///   countNums : count of numeric entries (identical to count here because
    ///     the caller only places parsed numerics into the bag)
    ///   average : arithmetic mean
    ///   stdDev  : sample std-dev  (sqrt(Σ(x-μ)²/(n-1))), requires n≥2
    ///   stdDevp : population std-dev (sqrt(Σ(x-μ)²/n)), requires n≥1
    ///   var     : sample variance (Σ(x-μ)²/(n-1)), requires n≥2
    ///   varp    : population variance (Σ(x-μ)²/n), requires n≥1
    /// Returns 0 for empty input and for stdDev/var when n&lt;2, matching the
    /// existing 0-on-empty convention that the rest of the renderer assumes.
    /// </summary>
    private static double ReducePivotValues(IEnumerable<double> values, string func)
    {
        var arr = values as double[] ?? values.ToArray();
        if (arr.Length == 0) return 0;
        switch (func.ToLowerInvariant())
        {
            case "sum": return arr.Sum();
            case "count": return arr.Length;
            case "countnums":
            case "countnum": return arr.Length;
            case "average":
            case "avg": return arr.Average();
            case "min": return arr.Min();
            case "max": return arr.Max();
            case "product":
                double p = 1;
                foreach (var v in arr) p *= v;
                return p;
            case "stddev":
            case "std":
            {
                if (arr.Length < 2) return 0;
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return Math.Sqrt(sq / (arr.Length - 1));
            }
            case "stddevp":
            case "stdp":
            {
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return Math.Sqrt(sq / arr.Length);
            }
            case "var":
            case "variance":
            {
                if (arr.Length < 2) return 0;
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return sq / (arr.Length - 1);
            }
            case "varp":
            {
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return sq / arr.Length;
            }
            default: return arr.Sum();
        }
    }

    private static (string col, int row) ParseCellRef(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        var col = cellRef[..i].ToUpperInvariant();
        var row = int.TryParse(cellRef[i..], out var r) ? r : 1;
        return (col, row);
    }

    private static int ColToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    private static string IndexToCol(int index)
    {
        // Inverse of ColToIndex (1-based: A=1, Z=26, AA=27, ...)
        var sb = new System.Text.StringBuilder();
        while (index > 0)
        {
            int rem = (index - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            index = (index - 1) / 26;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Multiply the cardinality (distinct non-empty values) of each field in the
    /// given index list. Used to size the pivot table's rendered area for the
    /// Location.ref range. Returns 1 when the list is empty (so layout math stays
    /// safe in pivots that have only column fields, only row fields, etc.).
    /// </summary>
    private static int ProductOfUniqueValues(List<int> fieldIndices, List<string[]> columnData)
    {
        if (fieldIndices.Count == 0) return 1;
        int product = 1;
        foreach (var idx in fieldIndices)
        {
            if (idx < 0 || idx >= columnData.Count) continue;
            var unique = columnData[idx].Where(v => !string.IsNullOrEmpty(v)).Distinct().Count();
            product *= Math.Max(1, unique);
        }
        return product;
    }
}
