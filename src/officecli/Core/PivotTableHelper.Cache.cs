// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

internal static partial class PivotTableHelper
{
    // ==================== Date Grouping Preprocessing ====================

    /// <summary>
    /// Metadata describing one date-grouped derived field. Used by the cache
    /// builder to emit native Excel <c>&lt;fieldGroup&gt;</c> XML that makes
    /// Excel recognize the derived field as a proper date bucket (required
    /// for the rendered layout to appear — without this, Excel detects a
    /// "fieldGroup shape mismatch" and falls back to grand-total only).
    /// </summary>
    private sealed class DateGroupSpec
    {
        /// <summary>Index of the original date field in the final columnData list.</summary>
        public int BaseFieldIdx { get; set; }
        /// <summary>Index of this derived field in the final columnData list.</summary>
        public int DerivedFieldIdx { get; set; }
        /// <summary>Grouping kind: "year" / "quarter" / "month" / "day".</summary>
        public string Grouping { get; set; } = "";
        /// <summary>Minimum date observed across the source column.</summary>
        public DateTime? MinDate { get; set; }
        /// <summary>Maximum date observed across the source column.</summary>
        public DateTime? MaxDate { get; set; }
    }

    /// <summary>
    /// Scans rows/cols/filters properties for <c>fieldName:grouping</c> syntax
    /// and creates a new virtual column per unique (field, grouping) pair. The
    /// original property strings are rewritten in-place so downstream
    /// ParseFieldList sees clean names.
    ///
    /// Example: input properties
    ///     rows = "日期:year,日期:quarter"
    ///     cols = "产品"
    /// With source columns [日期, 产品, 金额], returns:
    ///     headers    = [日期, 产品, 金额, 日期 (Year), 日期 (Quarter)]
    ///     columnData = [orig days, products, amounts, year labels, quarter labels]
    ///     dateGroups = [ {Base=0, Derived=3, Grouping=year}, {Base=0, Derived=4, Grouping=quarter} ]
    /// And mutates properties to:
    ///     rows = "日期 (Year),日期 (Quarter)"
    ///
    /// Multiple field specs referencing the same (field, grouping) pair share
    /// the single virtual column. Rows that don't parse as dates pass through
    /// unchanged so columns with a few stray non-date rows don't break.
    /// </summary>
    private static (string[] headers, List<string[]> columnData, List<DateGroupSpec> dateGroups) ApplyDateGrouping(
        string[] headers, List<string[]> columnData, Dictionary<string, string> properties)
    {
        // Track virtual columns keyed by (srcIdx, grouping). Value = new
        // column's header name, used to rewrite property references.
        var virtualColumns = new Dictionary<(int srcIdx, string grouping), string>();

        bool RewriteFieldListProp(string propKey)
        {
            if (!properties.TryGetValue(propKey, out var raw) || string.IsNullOrEmpty(raw))
                return false;

            var parts = raw.Split(',');
            var outParts = new List<string>(parts.Length);
            bool changed = false;

            foreach (var p in parts)
            {
                var spec = p.Trim();
                if (spec.Length == 0) continue;

                // Grouping suffix is allowed only if the prefix matches an
                // existing header. Otherwise the ':' might be part of the
                // field name (unlikely in practice but allowed by the parser)
                // and we must not mangle it.
                var colonIdx = spec.LastIndexOf(':');
                if (colonIdx <= 0 || colonIdx == spec.Length - 1)
                {
                    outParts.Add(spec);
                    continue;
                }

                var fieldName = spec.Substring(0, colonIdx).Trim();
                var grouping = spec.Substring(colonIdx + 1).Trim().ToLowerInvariant();
                if (grouping != "year" && grouping != "quarter"
                    && grouping != "month" && grouping != "day")
                {
                    outParts.Add(spec);
                    continue;
                }

                // Locate the source field.
                int srcIdx = -1;
                for (int i = 0; i < headers.Length; i++)
                {
                    if (headers[i] != null && headers[i].Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        srcIdx = i;
                        break;
                    }
                }
                if (srcIdx < 0)
                {
                    outParts.Add(spec);
                    continue;
                }

                if (!virtualColumns.TryGetValue((srcIdx, grouping), out var virtName))
                {
                    virtName = $"{fieldName} ({CapitalizeFirst(grouping)})";
                    virtualColumns[(srcIdx, grouping)] = virtName;
                }
                outParts.Add(virtName);
                changed = true;
            }

            if (changed)
                properties[propKey] = string.Join(",", outParts);
            return changed;
        }

        bool any = false;
        any |= RewriteFieldListProp("rows");
        any |= RewriteFieldListProp("cols");
        any |= RewriteFieldListProp("columns");
        any |= RewriteFieldListProp("filters");

        var dateGroups = new List<DateGroupSpec>();

        if (!any || virtualColumns.Count == 0)
            return (headers, columnData, dateGroups);

        // Materialize each virtual column AND record a DateGroupSpec so the
        // cache builder can emit <fieldGroup> XML. Output ordering follows
        // the insertion order of virtualColumns (first reference in props).
        // Also walk the source date column once to find min/max for the
        // rangePr startDate/endDate attributes Excel requires.
        var newHeaders = new List<string>(headers);
        foreach (var ((srcIdx, grouping), virtName) in virtualColumns)
        {
            var src = columnData[srcIdx];
            var derived = new string[src.Length];
            DateTime? min = null, max = null;
            for (int r = 0; r < src.Length; r++)
            {
                derived[r] = BucketDateValue(src[r], grouping);
                if (TryParseSourceDate(src[r], out var dt))
                {
                    if (!min.HasValue || dt < min.Value) min = dt;
                    if (!max.HasValue || dt > max.Value) max = dt;
                }
            }
            newHeaders.Add(virtName);
            columnData.Add(derived);
            dateGroups.Add(new DateGroupSpec
            {
                BaseFieldIdx = srcIdx,
                DerivedFieldIdx = columnData.Count - 1,
                Grouping = grouping,
                MinDate = min,
                MaxDate = max,
            });
        }

        return (newHeaders.ToArray(), columnData, dateGroups);
    }

    /// <summary>
    /// Parse a cell value as a DateTime, handling both string form
    /// ("2024-01-05") and Excel's OLE serial number form ("45296"). Used by
    /// ApplyDateGrouping to find the min/max needed for fieldGroup rangePr.
    /// </summary>
    private static bool TryParseSourceDate(string raw, out DateTime dt)
    {
        dt = default;
        if (string.IsNullOrEmpty(raw)) return false;
        // CONSISTENCY(timezone): Use AssumeUniversal+AdjustToUniversal so the parsed
        // DateTime has Kind=Utc and no timezone shift occurs when OpenXML SDK serializes
        // it. AssumeLocal would produce Kind=Local which the SDK converts to UTC on
        // write, shifting dates by the local UTC offset (e.g. UTC+8 shifts Jan 15 → Jan 14).
        if (DateTime.TryParse(raw, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out dt))
            return true;
        if (double.TryParse(raw, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var serial))
        {
            try { dt = DateTime.FromOADate(serial); return true; }
            catch { return false; }
        }
        return false;
    }

    /// <summary>
    /// Transform a raw cell value into a date bucket label for the given
    /// grouping. Accepts either a formatted date string ("2024-01-05") or
    /// Excel's serial number form ("45296"). Unparseable values pass through
    /// unchanged.
    /// </summary>
    private static string BucketDateValue(string raw, string grouping)
    {
        if (string.IsNullOrEmpty(raw)) return raw ?? string.Empty;

        DateTime dt;
        // CONSISTENCY(timezone): match TryParseSourceDate — use AssumeUniversal to
        // avoid Kind=Local which shifts dates by local UTC offset during serialization.
        if (!DateTime.TryParse(raw, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out dt))
        {
            if (double.TryParse(raw, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var serial))
            {
                try { dt = DateTime.FromOADate(serial); }
                catch { return raw; }
            }
            else
            {
                return raw;
            }
        }

        // Bucket labels must match the canonical names emitted by
        // ComputeDateGroupBuckets (Qtr1..Qtr4 / Jan..Dec / 1..31) so the
        // cache's groupItems and the renderer's columnData agree on bucket
        // identity. Cross-year disambiguation for quarter/month/day is
        // handled by the year field (if present as a sibling row/col).
        return grouping switch
        {
            "year"    => dt.Year.ToString("D4", System.Globalization.CultureInfo.InvariantCulture),
            "quarter" => $"Qtr{(dt.Month - 1) / 3 + 1}",
            "month"   => MonthShortName(dt.Month),
            "day"     => dt.Day.ToString(System.Globalization.CultureInfo.InvariantCulture),
            _         => raw,
        };
    }

    private static string MonthShortName(int month)
        => month switch
        {
            1  => "Jan", 2  => "Feb", 3  => "Mar", 4  => "Apr",
            5  => "May", 6  => "Jun", 7  => "Jul", 8  => "Aug",
            9  => "Sep", 10 => "Oct", 11 => "Nov", 12 => "Dec",
            _  => month.ToString(System.Globalization.CultureInfo.InvariantCulture),
        };

    private static string CapitalizeFirst(string s)
        => string.IsNullOrEmpty(s) ? s : char.ToUpperInvariant(s[0]) + s.Substring(1);

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
        // R6-3: reject columns beyond Excel's hard max (XFD = 16384). Previously
        // XFE / XFZ / ZZZZ silently parsed into oversized indices, produced a
        // giant colCount, and either crashed deep in the renderer or wrote an
        // invalid source range into the cache.
        const int ExcelMaxColumn = 16384; // XFD
        if (startColIdx > ExcelMaxColumn)
            throw new ArgumentException($"Column {startCol} out of range (max: XFD)");
        if (endColIdx > ExcelMaxColumn)
            throw new ArgumentException($"Column {endCol} out of range (max: XFD)");
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
        // Error cells (DataType=Error, e.g. #DIV/0!) must not be treated as string values.
        // Return the sentinel so BuildCacheField can emit ErrorItem instead of StringItem.
        if (cell.DataType?.Value == CellValues.Error)
            return ErrorCellSentinel;

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
            string[] headers, List<string[]> columnData,
            HashSet<int>? axisFieldIndices = null,
            List<DateGroupSpec>? dateGroups = null,
            uint?[]? columnNumFmtIds = null)
    {
        var recordCount = columnData.Count > 0 ? columnData[0].Length : 0;

        // RenderPivotIntoSheet now materializes all pivot cells into sheetData
        // (including the N≥3 general renderer), so Excel can display the pre-
        // rendered values directly without a cache refresh. Do NOT set
        // RefreshOnLoad — it causes Excel to clear the pre-rendered cells and
        // attempt a live rebuild from the cache definition. If the rebuild
        // fails (e.g. complex N≥3 rowItems structure, security policy blocking
        // refresh, or WPS Office's limited pivot support), the user sees an
        // empty pivot skeleton instead of the correct data. Real Excel/
        // files likewise ship rendered cells without refreshOnLoad.
        var cacheDef = new PivotCacheDefinition
        {
            CreatedVersion = 3,
            MinRefreshableVersion = 3,
            RefreshedVersion = 3,
            RecordCount = (uint)recordCount
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
        //
        // Date group handling:
        //   - Base date field gets standard enumerated items PLUS a <fieldGroup
        //     par="N"/> pointer to the FIRST derived field (Excel's convention).
        //   - Each derived field writes a synthetic cacheField with
        //     databaseField="0", a <fieldGroup base="baseIdx"> containing
        //     <rangePr groupBy="..." startDate=".." endDate=".." /> and a
        //     <groupItems> list of string labels — including LEADING/TRAILING
        //     sentinels ("<startDate" / ">endDate") that Excel requires.
        //   - Derived fields emit NO entries in pivotCacheRecords (databaseField=0).
        //     BuildCacheRecords in the caller must skip them, which we signal by
        //     setting fieldNumeric[derivedIdx] = false AND leaving fieldValueIndex
        //     entries pointing into the enumerated shared items of the synthetic
        //     field. See BuildCacheRecords for the skip logic.
        var fieldNumeric = new bool[headers.Length];
        var fieldValueIndex = new Dictionary<string, int>[headers.Length];

        // Build quick lookups from the date group specs.
        var derivedByIdx = new Dictionary<int, DateGroupSpec>();
        var baseFields = new HashSet<int>();
        if (dateGroups != null)
        {
            foreach (var g in dateGroups)
            {
                derivedByIdx[g.DerivedFieldIdx] = g;
                baseFields.Add(g.BaseFieldIdx);
            }
        }

        var cacheFields = new CacheFields { Count = (uint)headers.Length };
        for (int i = 0; i < headers.Length; i++)
        {
            var fieldName = string.IsNullOrEmpty(headers[i]) ? $"Column{i + 1}" : headers[i];
            var values = i < columnData.Count ? columnData[i] : Array.Empty<string>();

            // R19-1: per-column source numFmtId (date/currency/etc.) to stamp
            // on the cacheField so the pivot renders values with the same
            // formatting as the source column. Null means "General" and we
            // leave the default in place.
            uint? srcNumFmtId = (columnNumFmtIds != null && i < columnNumFmtIds.Length)
                ? columnNumFmtIds[i] : null;

            if (derivedByIdx.TryGetValue(i, out var spec))
            {
                // Derived date group field — synthesized, no records entries.
                var derived = BuildDateGroupDerivedCacheField(fieldName, spec,
                    out fieldValueIndex[i]);
                if (srcNumFmtId.HasValue) derived.NumberFormatId = srcNumFmtId.Value;
                cacheFields.AppendChild(derived);
                fieldNumeric[i] = false; // records should skip this field
                continue;
            }

            if (baseFields.Contains(i))
            {
                // Base date field — enumerate date items (not a plain numeric
                // column) and add a <fieldGroup par="N"/> pointing at the first
                // derived field for this base. Records for this field emit
                // <x v="N"/> referencing the enumerated date items.
                int parIdx = derivedByIdx
                    .Where(kv => kv.Value.BaseFieldIdx == i)
                    .Min(kv => kv.Key);
                var baseField = BuildDateGroupBaseCacheField(fieldName, values, parIdx,
                    out fieldValueIndex[i]);
                // Prefer the source column's numFmtId when present; else keep
                // the builder's 164u default (yyyy-mm-dd).
                if (srcNumFmtId.HasValue) baseField.NumberFormatId = srcNumFmtId.Value;
                cacheFields.AppendChild(baseField);
                fieldNumeric[i] = false;
                continue;
            }

            // Axis fields (row/col/filter) go through the string/indexed path
            // even when their values parse as numeric, so pivotField items
            // indices and cache record references stay in sync.
            bool forceStringIndexed = axisFieldIndices?.Contains(i) == true;
            var plainField = BuildCacheField(
                fieldName, values, out fieldNumeric[i], out fieldValueIndex[i], forceStringIndexed);
            if (srcNumFmtId.HasValue) plainField.NumberFormatId = srcNumFmtId.Value;
            cacheFields.AppendChild(plainField);
        }
        cacheDef.AppendChild(cacheFields);

        return (cacheDef, fieldNumeric, fieldValueIndex);
    }

    private static CacheField BuildCacheField(
        string name, string[] values, out bool isNumeric, out Dictionary<string, int> valueIndex,
        bool forceStringIndexed = false)
    {
        var field = new CacheField { Name = name, NumberFormatId = 0u };
        // Exclude error-cell sentinels from the numeric check — they are neither
        // numeric nor regular strings; they will be emitted as ErrorItem elements.
        bool valuesAreNumeric = values.Length > 0 && values.All(v =>
            string.IsNullOrEmpty(v) || v == ErrorCellSentinel
            || double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _));
        // When forceStringIndexed is true (axis fields), report isNumeric=false
        // so downstream record-writing code uses the valueIndex map to emit
        // <x v="N"/> references instead of <n v="..."/> direct values. The
        // local 'valuesAreNumeric' still determines which sharedItems branch
        // we take below.
        isNumeric = valuesAreNumeric && !forceStringIndexed;
        valueIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        var sharedItems = new SharedItems();

        // MIXED strategy — verified against canonical Excel-authored pivots:
        //
        //   • Numeric fields: emit ONLY containsNumber/minValue/maxValue metadata,
        //     no enumerated items, no count attribute. Records reference values
        //     directly via <n v="..."/>.
        //   • String fields: enumerate every unique value as <s v="..."/> with
        //     count attribute. Records reference them by index via <x v="N"/>.
        //
        // A uniform strategy (always enumerate, always index-reference) is
        // technically valid OOXML but introduces an asymmetry Excel handles
        // less reliably (numeric data fields with item enumeration have failed
        // to render in testing, even though the file passes schema validation).
        bool hasErrorCells = values.Any(v => v == ErrorCellSentinel);
        if (isNumeric && values.Any(v => !string.IsNullOrEmpty(v) && v != ErrorCellSentinel))
        {
            var nums = values.Where(v => !string.IsNullOrEmpty(v) && v != ErrorCellSentinel)
                .Select(v => double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            sharedItems.ContainsSemiMixedTypes = false;
            sharedItems.ContainsString = false;
            sharedItems.ContainsNumber = true;
            sharedItems.MinValue = nums.Min();
            sharedItems.MaxValue = nums.Max();
            // No string items enumerated — records emit <n v="..."/> or index ref for errors.
        }
        else
        {
            var uniqueValues = values
                .Where(v => !string.IsNullOrEmpty(v) && v != ErrorCellSentinel)
                .Distinct()
                .OrderByAxis(v => v)
                .ToList();
            // Error cells occupy their own ErrorItem slots after the string items.
            var uniqueErrors = values
                .Where(v => v == ErrorCellSentinel)
                .Distinct()
                .ToList();
            int totalCount = uniqueValues.Count + uniqueErrors.Count;
            sharedItems.Count = (uint)totalCount;
            if (hasErrorCells)
            {
                sharedItems.ContainsSemiMixedTypes = false;
            }
            for (int i = 0; i < uniqueValues.Count; i++)
            {
                var v = uniqueValues[i];
                // R2-2: strip XML-illegal chars (e.g. U+0000) before writing.
                sharedItems.AppendChild(new StringItem { Val = SanitizeXmlText(v) });
                if (!valueIndex.ContainsKey(v))
                    valueIndex[v] = i;
            }
            // Emit ErrorItem elements for error-cell sentinels.
            for (int i = 0; i < uniqueErrors.Count; i++)
            {
                sharedItems.AppendChild(new ErrorItem { Val = "#VALUE!" });
                valueIndex[ErrorCellSentinel] = uniqueValues.Count + i;
            }
            // OOXML requires longText="1" when any string exceeds 255 chars.
            // Without it, Excel reports "problem with some content" and repairs.
            if (uniqueValues.Any(v => v.Length > 255))
                sharedItems.LongText = true;
        }

        field.AppendChild(sharedItems);
        return field;
    }

    // ==================== Date Group Cache Field Builders ====================

    /// <summary>
    /// Build the base date cacheField for a date-grouped column. Enumerates
    /// every parsed source date as a <c>&lt;d v="..."/&gt;</c> shared item and
    /// appends a <c>&lt;fieldGroup par="N"/&gt;</c> pointing at the first
    /// derived field for this base (Excel convention: even when there are
    /// multiple derived fields — year + quarter + month — only the lowest
    /// par index is written on the base).
    ///
    /// Verified against Excel-authored /tmp/date_authored.xlsx: the base
    /// field has <c>containsDate="1"</c>, enumerated ISO-format dates, no
    /// <c>containsString</c>/<c>containsNumber</c> attributes.
    /// </summary>
    private static CacheField BuildDateGroupBaseCacheField(
        string name, string[] values, int parDerivedIdx,
        out Dictionary<string, int> valueIndex)
    {
        var field = new CacheField { Name = name, NumberFormatId = 164u };
        valueIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        // Collect unique parsed dates in source order. Excel enumerates them
        // in the order they first appear in the data, which keeps the cache
        // record indices stable and human-readable.
        var uniqueDates = new List<DateTime>();
        var dateToIdx = new Dictionary<DateTime, int>();
        DateTime? min = null, max = null;
        for (int r = 0; r < values.Length; r++)
        {
            if (!TryParseSourceDate(values[r], out var dt)) continue;
            if (!dateToIdx.ContainsKey(dt))
            {
                dateToIdx[dt] = uniqueDates.Count;
                uniqueDates.Add(dt);
            }
            if (!min.HasValue || dt < min.Value) min = dt;
            if (!max.HasValue || dt > max.Value) max = dt;
        }

        var sharedItems = new SharedItems
        {
            ContainsSemiMixedTypes = false,
            ContainsNonDate = false,
            ContainsDate = true,
            ContainsString = false,
            Count = (uint)uniqueDates.Count
        };
        if (min.HasValue) sharedItems.MinDate = min.Value;
        if (max.HasValue) sharedItems.MaxDate = max.Value;

        foreach (var dt in uniqueDates)
        {
            sharedItems.AppendChild(new DateTimeItem { Val = dt });
        }

        // Populate the value→index map so BuildCacheRecords can resolve each
        // source row's date value to the correct sharedItems index. The map
        // keys are the ORIGINAL raw cell values (not the normalized dates),
        // since that's what the record writer will look up.
        for (int r = 0; r < values.Length; r++)
        {
            var raw = values[r];
            if (string.IsNullOrEmpty(raw)) continue;
            if (valueIndex.ContainsKey(raw)) continue;
            if (TryParseSourceDate(raw, out var dt) && dateToIdx.TryGetValue(dt, out var idx))
                valueIndex[raw] = idx;
        }

        field.AppendChild(sharedItems);

        // <fieldGroup par="N"/> — the "par" attribute points at the FIRST
        // derived field for this base. Verified against /tmp/date_authored.xlsx
        // where the base had par=3 pointing at the Quarters field at idx 3.
        field.AppendChild(new FieldGroup { ParentId = (uint)parDerivedIdx });
        return field;
    }

    /// <summary>
    /// Build a derived date-group cacheField (Year / Quarter / Month / Day)
    /// with <c>databaseField="0"</c> and a synthetic <c>&lt;fieldGroup base=&gt;
    /// &lt;rangePr groupBy="..."/&gt; &lt;groupItems&gt;...&lt;/groupItems&gt;
    /// &lt;/fieldGroup&gt;</c> structure.
    ///
    /// The groupItems list follows Excel's sentinel convention: a leading
    /// <c>&lt;startDate</c> and trailing <c>&gt;endDate</c> sentinel bracket
    /// the real buckets. Excel uses sentinel indices (0 and last) internally
    /// to mark "out of range" values, but for our purposes only the middle
    /// real buckets matter. The renderer writes bucket labels directly into
    /// sheetData so the sentinel placeholder semantics are moot.
    ///
    /// The valueIndex map lets BuildCacheRecords resolve each source row's
    /// bucketed LABEL value back into a groupItems index ≥ 1 (skipping the
    /// leading sentinel). Derived fields do NOT emit records entries because
    /// databaseField="0", but we still populate the map defensively.
    /// </summary>
    private static CacheField BuildDateGroupDerivedCacheField(
        string name, DateGroupSpec spec, out Dictionary<string, int> valueIndex)
    {
        valueIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        var field = new CacheField
        {
            Name = name,
            NumberFormatId = 0u,
            DatabaseField = false  // Derived — not backed by a record column
        };

        // Compute bucket labels for the grouping. The order and count must
        // match Excel's convention because rowItems/colItems reference these
        // indices. Year buckets are per-year observed in the data; quarter
        // labels use the Qtr1..Qtr4 short form Excel writes natively.
        List<string> buckets = ComputeDateGroupBuckets(spec);

        // Wrap the buckets with Excel's sentinel items:
        //   idx 0:        "<startDate"
        //   idx 1..N:     real buckets (Qtr1, Qtr2, ...; 2024, 2025, ...)
        //   idx N+1:      ">endDate"
        var startSentinel = spec.MinDate.HasValue
            ? "<" + spec.MinDate.Value.ToString("yyyy.MM.dd", System.Globalization.CultureInfo.InvariantCulture)
            : "<start";
        // Guard against DateTime.MaxValue overflow: if MaxDate is already the
        // last representable day, clamp AddDays(1) to DateTime.MaxValue itself
        // so the sentinel label and OOXML EndDate remain well-formed.
        var endSentinel = spec.MaxDate.HasValue
            ? ">" + (spec.MaxDate.Value < DateTime.MaxValue.Date
                ? spec.MaxDate.Value.AddDays(1)
                : spec.MaxDate.Value)
                .ToString("yyyy.MM.dd", System.Globalization.CultureInfo.InvariantCulture)
            : ">end";

        var allItems = new List<string>(buckets.Count + 2);
        allItems.Add(startSentinel);
        allItems.AddRange(buckets);
        allItems.Add(endSentinel);

        // Populate valueIndex so raw bucket labels (the ones our renderer
        // wrote into columnData) resolve to the correct groupItems index.
        for (int i = 0; i < buckets.Count; i++)
        {
            valueIndex[buckets[i]] = i + 1; // +1 for leading sentinel
        }

        var fieldGroup = new FieldGroup { Base = (uint)spec.BaseFieldIdx };

        var rangePr = new RangeProperties
        {
            GroupBy = spec.Grouping switch
            {
                "year"    => GroupByValues.Years,
                "quarter" => GroupByValues.Quarters,
                "month"   => GroupByValues.Months,
                "day"     => GroupByValues.Days,
                _         => GroupByValues.Days,
            },
        };
        if (spec.MinDate.HasValue) rangePr.StartDate = spec.MinDate.Value;
        // CONSISTENCY(date-boundary-clamp): same AddDays(1) guard as endSentinel above.
        if (spec.MaxDate.HasValue) rangePr.EndDate = spec.MaxDate.Value < DateTime.MaxValue.Date
            ? spec.MaxDate.Value.AddDays(1)
            : spec.MaxDate.Value;
        fieldGroup.AppendChild(rangePr);

        var groupItems = new GroupItems { Count = (uint)allItems.Count };
        foreach (var label in allItems)
            // R2-2: defensive sanitize — date labels are code-generated so
            // they shouldn't contain control chars, but keep parity with the
            // sharedItems writer in case a format spec ever changes.
            groupItems.AppendChild(new StringItem { Val = SanitizeXmlText(label) });
        fieldGroup.AppendChild(groupItems);

        field.AppendChild(fieldGroup);
        return field;
    }

    /// <summary>
    /// Compute the ordered list of bucket labels for a given date group spec.
    /// These labels are FIXED across years (matching Excel's native
    /// behavior): quarter → Qtr1..Qtr4, month → Jan..Dec, day → 1..31.
    /// Year is the exception: it returns the actual observed years.
    ///
    /// Excel treats quarter/month/day as CATEGORICAL fields — the same
    /// "Qtr1" bucket applies to all years in the data. Different years of
    /// the same quarter disambiguate in the rendered pivot via the
    /// rowItems/colItems (year_idx, quarter_idx) tuple, not via label
    /// text. Verified against /tmp/date_authored.xlsx where quarters
    /// enumerated exactly 4 buckets regardless of year range.
    ///
    /// This is critical: if we emit non-standard labels like "2024-Q1"
    /// (which we initially did), Excel's pivot engine crashes when
    /// parsing month grouping because it expects Jan..Dec format. The
    /// buckets below are the canonical names Excel writes natively.
    /// </summary>
    private static List<string> ComputeDateGroupBuckets(DateGroupSpec spec)
    {
        var result = new List<string>();
        switch (spec.Grouping)
        {
            case "year":
                // Years ARE actual — observed years in the data.
                if (!spec.MinDate.HasValue || !spec.MaxDate.HasValue) return result;
                for (int y = spec.MinDate.Value.Year; y <= spec.MaxDate.Value.Year; y++)
                    result.Add(y.ToString("D4", System.Globalization.CultureInfo.InvariantCulture));
                break;

            case "quarter":
                // Fixed set regardless of year range.
                result.AddRange(new[] { "Qtr1", "Qtr2", "Qtr3", "Qtr4" });
                break;

            case "month":
                // Fixed set. Excel uses 3-letter English month abbreviations
                // (Jan..Dec) in its native format — verified against Excel's
                // quarter-grouping output which emits "Qtr1..Qtr4". We follow
                // the same short-form convention for months.
                result.AddRange(new[]
                {
                    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
                });
                break;

            case "day":
                // Fixed set — day-of-month 1..31.
                for (int d = 1; d <= 31; d++)
                    result.Add(d.ToString(System.Globalization.CultureInfo.InvariantCulture));
                break;
        }
        return result;
    }

    // ==================== Cache Records Builder ====================

    /// <summary>
    /// Build pivotCacheRecords using the MIXED strategy:
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
        List<string[]> columnData, bool[] fieldNumeric, Dictionary<string, int>[] fieldValueIndex,
        HashSet<int>? skipFieldIndices = null)
    {
        var recordCount = columnData.Count > 0 ? columnData[0].Length : 0;
        var fieldCount = columnData.Count;
        var records = new PivotCacheRecords { Count = (uint)recordCount };

        for (int r = 0; r < recordCount; r++)
        {
            var record = new PivotCacheRecord();
            for (int f = 0; f < fieldCount; f++)
            {
                // Derived date-group fields carry databaseField="0" and therefore
                // don't contribute entries to pivotCacheRecords — they're computed
                // on-the-fly by Excel from the base date field's <fieldGroup>
                // <rangePr>/<groupItems> definition. Skip them here so the record
                // column count matches the non-derived fields.
                if (skipFieldIndices?.Contains(f) == true) continue;

                var v = columnData[f][r];
                if (string.IsNullOrEmpty(v))
                {
                    record.AppendChild(new MissingItem());
                }
                else if (v == ErrorCellSentinel)
                {
                    // Error cell — reference the ErrorItem in sharedItems if indexed, or
                    // emit MissingItem for numeric fields that have no sharedItems index.
                    if (fieldValueIndex[f].TryGetValue(v, out var errIdx))
                        record.AppendChild(new FieldItem { Val = (uint)errIdx });
                    else
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

    // ==================== Pivot cache sharing (design change) ====================
    //
    // Excel's contract is "one pivotCache per source range, shared by all
    // pivots that reference that source". OfficeCLI originally created a new
    // cache per pivot (one cacheDefinition + cacheRecords part each), which
    // bloated files and diverged from Excel behavior on refresh (refreshing
    // the source under one pivot did not propagate to its sibling pivot
    // because they each owned a private cache snapshot).
    //
    // The three sites that need to honor sharing are:
    //   - Add: reuse an existing cache if any pivot already binds to a
    //     source-equivalent cacheSource (NormalizePivotSource).
    //   - Remove: derived ref-counting — only delete the cache part when
    //     the LAST pivot referring to it is removed. This is what
    //     PrunePivotCacheIfOrphan already does (it scans every remaining
    //     pivottable part); the design rule just makes that load-bearing.
    //   - Set source: copy-on-write. If the cache is currently shared, do
    //     not mutate it in place — clone cacheDefinition + cacheRecords +
    //     workbook.PivotCaches entry so this pivot points to a fresh,
    //     private cache; siblings continue to see their original cache.

    /// <summary>
    /// Normalize a pivot source spec ("<sheet>!<range>" or just "<range>") into
    /// a canonical key for equality comparison across two pivots' cacheSource.
    /// First-version coverage: in-workbook explicit sheet+range references.
    ///
    /// Normalization rules:
    ///   - sheet name: outer single/double quotes stripped ('Sheet 1' == Sheet 1),
    ///     comparison case-insensitive (matches Excel sheet-name semantics).
    ///   - range: '$' absolute markers stripped ($A$1:$B$3 == A1:B3), column
    ///     letters uppercased.
    ///   - Whitespace around '!' / commas trimmed.
    ///
    /// Deferred (returns the input untouched, so they fall through to
    /// "different source" → no sharing — safe default):
    ///   - named ranges (e.g. <c>SalesData</c>) — would need wb-level resolve.
    ///   - external workbook references (<c>[Other.xlsx]Sheet1!A1:C5</c>).
    ///   - structured-ref table references (<c>Table1[#All]</c>).
    /// </summary>
    internal static string NormalizePivotSource(string? sheetName, string? rangeRef)
    {
        if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeRef))
            return $"{sheetName}|{rangeRef}";
        var s = sheetName.Trim();
        // Strip a single layer of matching quotes.
        if (s.Length >= 2 && ((s[0] == '\'' && s[^1] == '\'') || (s[0] == '"' && s[^1] == '"')))
            s = s.Substring(1, s.Length - 2);
        var r = rangeRef.Trim().Replace("$", "").ToUpperInvariant();
        return s.ToUpperInvariant() + "!" + r;
    }

    /// <summary>
    /// Find an existing PivotTableCacheDefinitionPart whose cacheSource is
    /// equivalent (after normalization) to the given (sheet, ref) target.
    /// Walks every pivot table part in the workbook and dedupes by part.
    /// Returns null if no match.
    /// </summary>
    internal static PivotTableCacheDefinitionPart? FindMatchingCachePart(
        WorkbookPart workbookPart, string sheetName, string rangeRef)
    {
        var target = NormalizePivotSource(sheetName, rangeRef);
        var seen = new HashSet<PivotTableCacheDefinitionPart>();
        foreach (var ws in workbookPart.WorksheetParts)
        {
            foreach (var pp in ws.PivotTableParts)
            {
                var cp = pp.PivotTableCacheDefinitionPart;
                if (cp == null || !seen.Add(cp)) continue;
                var ws2 = cp.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                if (ws2 == null) continue;
                var key = NormalizePivotSource(ws2.Sheet?.Value, ws2.Reference?.Value);
                if (key == target) return cp;
            }
        }
        return null;
    }

    /// <summary>
    /// Count how many distinct PivotTablePart instances (across all worksheets)
    /// reference the given cacheDefinitionPart. Used to decide whether Set
    /// source must clone (CoW) or may mutate the cache in place.
    /// </summary>
    internal static int CountCacheReferrers(WorkbookPart workbookPart,
        PivotTableCacheDefinitionPart cachePart)
    {
        int n = 0;
        foreach (var ws in workbookPart.WorksheetParts)
            foreach (var pp in ws.PivotTableParts)
                if (pp.PivotTableCacheDefinitionPart == cachePart) n++;
        return n;
    }

    /// <summary>
    /// Clone a PivotTableCacheDefinitionPart (and its child
    /// PivotTableCacheRecordsPart, if any) into a brand-new workbook-level
    /// part, then register a new <pivotCache> entry with a fresh cacheId.
    /// Used by Set source's copy-on-write path when the original cache is
    /// shared with other pivots.
    ///
    /// The clone copies XML content via streaming, so any subsequent mutation
    /// to the new cache cannot leak back to the original.
    /// </summary>
    internal static PivotTableCacheDefinitionPart CloneCachePartForCoW(
        WorkbookPart workbookPart, PivotTableCacheDefinitionPart original)
    {
        var clone = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        // Copy the cache definition XML stream wholesale.
        using (var src = original.GetStream(FileMode.Open, FileAccess.Read))
        using (var dst = clone.GetStream(FileMode.Create, FileAccess.Write))
        {
            src.CopyTo(dst);
        }
        // Force the SDK to re-read the part so subsequent edits go through
        // its strongly-typed PivotCacheDefinition view.
        _ = clone.PivotCacheDefinition;

        // Clone the records part (if present) under the new cache part.
        var origRecords = original.GetPartsOfType<PivotTableCacheRecordsPart>().FirstOrDefault();
        if (origRecords != null)
        {
            var cloneRecords = clone.AddNewPart<PivotTableCacheRecordsPart>();
            using (var src = origRecords.GetStream(FileMode.Open, FileAccess.Read))
            using (var dst = cloneRecords.GetStream(FileMode.Create, FileAccess.Write))
            {
                src.CopyTo(dst);
            }
            // Re-point the cacheDef's r:id to the new records part.
            if (clone.PivotCacheDefinition != null)
                clone.PivotCacheDefinition.Id = clone.GetIdOfPart(cloneRecords);
        }

        // Register a new <pivotCache> entry in workbook.xml with a fresh cacheId.
        var wb = workbookPart.Workbook
            ?? throw new InvalidOperationException("Workbook is missing");
        var pivotCaches = wb.GetFirstChild<PivotCaches>();
        if (pivotCaches == null)
        {
            pivotCaches = new PivotCaches();
            var insertBefore = wb.GetFirstChild<WebPublishing>()
                ?? wb.GetFirstChild<FileRecoveryProperties>()
                ?? (OpenXmlElement?)wb.GetFirstChild<WebPublishObjects>();
            if (insertBefore != null)
                wb.InsertBefore(pivotCaches, insertBefore);
            else
                wb.AppendChild(pivotCaches);
        }
        uint newCacheId = pivotCaches.Elements<PivotCache>()
            .Select(pc => pc.CacheId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;
        pivotCaches.AppendChild(new PivotCache
        {
            CacheId = newCacheId,
            Id = workbookPart.GetIdOfPart(clone)
        });
        wb.Save();
        return clone;
    }

    /// <summary>
    /// Look up the cacheId (in workbook.xml's pivotCaches) for a given
    /// cacheDefinitionPart by matching its r:id.
    /// </summary>
    internal static uint? GetCacheIdForPart(WorkbookPart workbookPart,
        PivotTableCacheDefinitionPart cachePart)
    {
        string? rid = null;
        try { rid = workbookPart.GetIdOfPart(cachePart); } catch { return null; }
        var caches = workbookPart.Workbook?.GetFirstChild<PivotCaches>();
        if (caches == null) return null;
        foreach (var pc in caches.Elements<PivotCache>())
            if (pc.Id?.Value == rid) return pc.CacheId?.Value;
        return null;
    }

    // ==================== Pivot source resolution v2 (B6 v2) ====================
    //
    // v1 only matched explicit "<sheet>!<range>" literally. v2 adds two
    // common forms so cache sharing also works when the user authors with:
    //   • Structured table refs:     Table1[#All] / Table1 / Table1[#Data]
    //   • Workbook-/sheet-scoped name: SalesData  /  Sheet1!SalesData
    //
    // Anything we can't or shouldn't resolve (external workbook ref,
    // dynamic OFFSET-based names, [ColumnName] single-column refs, multi-
    // range names) falls through to the v1 string-key path: safe — the
    // caller treats it as "different source" and creates an independent
    // cache rather than risking an incorrect share.

    /// <summary>
    /// Try to resolve a pivot source spec into an explicit (sheet, range)
    /// tuple. Returns null on fall-through (caller should keep using the
    /// original spec). The returned range is always an "A1:C100"-style
    /// rectangular reference suitable for cacheSource WorksheetSource.
    ///
    /// Resolution priority:
    ///   1. Explicit "Sheet!Range" → returned as-is (after quote/$ trim).
    ///   2. Bare "A1:C100" with defaultSheet supplied → (defaultSheet, range).
    ///   3. Token containing '[' → structured table ref.
    ///      Supported: Table1, Table1[#All], Table1[#Data]
    ///      Unsupported (returns null): [#Headers], [#Totals], [ColumnName],
    ///      compound like Table1[[#Data],[Col]].
    ///   4. Single token (no '!' / no '[') → defined-name lookup.
    ///      Supported: workbook-scoped or sheet-scoped name whose body is
    ///      a single "Sheet!Range" reference.
    ///      Unsupported (returns null): dynamic body (OFFSET / INDEX),
    ///      multi-range body ("Sheet1!A1:C5,Sheet1!E1:G5").
    /// </summary>
    internal static (string sheet, string rangeRef)? ResolvePivotSourceSpec(
        WorkbookPart workbookPart, string sourceSpec, string? defaultSheet = null)
    {
        if (string.IsNullOrWhiteSpace(sourceSpec)) return null;
        var spec = sourceSpec.Trim();

        // External workbook ref — explicitly unsupported, fall through.
        if (spec.StartsWith("[")) return null;

        // 3. Structured table ref — must be checked BEFORE the explicit
        // "!" path because Excel allows e.g. Table1 in a context that
        // happens not to contain '!'. We also catch Table1[#All] before
        // ambiguous '[' parsing in defined names.
        if (spec.Contains('['))
        {
            return ResolveTableRef(workbookPart, spec);
        }

        // 1. Explicit Sheet!Range
        if (spec.Contains('!'))
        {
            var parts = spec.Split('!', 2);
            var sheet = parts[0].Trim();
            if (sheet.Length >= 2 && ((sheet[0] == '\'' && sheet[^1] == '\'')
                                    || (sheet[0] == '"' && sheet[^1] == '"')))
                sheet = sheet.Substring(1, sheet.Length - 2);
            var rangePart = parts[1].Trim();

            // A "Sheet1!SalesData"-style sheet-qualified defined name —
            // rangePart is a single identifier, not an A1 reference.
            if (LooksLikeRangeRef(rangePart))
                return (sheet, rangePart);

            // Sheet-qualified defined name: try sheet-scoped first.
            var resolvedName = ResolveDefinedName(workbookPart, rangePart, sheetScopeName: sheet);
            return resolvedName;
        }

        // 2/4. Bare token — could be range, table name, or defined name.
        // Try table first because a bare "Tbl1" matches LooksLikeRangeRef
        // (letters+digits) but isn't a real cell reference; only fall back
        // to range when no matching table exists.
        var asTable = ResolveTableRef(workbookPart, spec);
        if (asTable != null) return asTable;
        if (LooksLikeRangeRef(spec))
        {
            return defaultSheet != null ? (defaultSheet, spec) : null;
        }
        return ResolveDefinedName(workbookPart, spec, sheetScopeName: null);
    }

    private static bool LooksLikeRangeRef(string s)
    {
        // "A1" or "A1:C100" — letters then digits, optionally colon-range.
        // Tolerant of surrounding $ markers.
        var t = s.Replace("$", "").Trim();
        return System.Text.RegularExpressions.Regex.IsMatch(t,
            @"^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$");
    }

    private static (string sheet, string rangeRef)? ResolveTableRef(
        WorkbookPart workbookPart, string spec)
    {
        // Parse: TableName  |  TableName[#All]  |  TableName[#Data]
        // Reject anything more complex (column refs, compound).
        var bracketIdx = spec.IndexOf('[');
        string tableName;
        string modifier; // "" (no brackets), "#All", "#Data", or other
        if (bracketIdx < 0)
        {
            tableName = spec.Trim();
            modifier = "";
        }
        else
        {
            tableName = spec.Substring(0, bracketIdx).Trim();
            var inside = spec.Substring(bracketIdx); // "[#All]" etc
            // Must be exactly "[#X]" with no comma / nested brackets
            var m = System.Text.RegularExpressions.Regex.Match(inside,
                @"^\[(#[A-Za-z]+)\]$");
            if (!m.Success) return null; // [ColumnName] / compound — fall through
            modifier = m.Groups[1].Value;
        }

        if (string.IsNullOrEmpty(tableName)) return null;

        // Walk every TableDefinitionPart to find the table by name.
        foreach (var ws in workbookPart.WorksheetParts)
        {
            foreach (var tdp in ws.TableDefinitionParts)
            {
                var t = tdp.Table;
                if (t == null) continue;
                var name = t.Name?.Value ?? t.DisplayName?.Value;
                if (name == null) continue;
                if (!string.Equals(name, tableName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var refStr = t.Reference?.Value;
                if (string.IsNullOrEmpty(refStr)) return null;
                var sheetName = LookupSheetNameForPart(workbookPart, ws);
                if (sheetName == null) return null;

                bool tableHasHeaders = (t.HeaderRowCount?.Value ?? 1u) >= 1u;

                switch (modifier)
                {
                    case "":
                    case "#All":
                        // Whole table including header (if any).
                        return (sheetName, refStr.Replace("$", "").ToUpperInvariant());

                    case "#Data":
                    {
                        // Strip header row. If table has no headers, same as #All.
                        if (!tableHasHeaders)
                            return (sheetName, refStr.Replace("$", "").ToUpperInvariant());
                        var stripped = StripHeaderRow(refStr);
                        if (stripped == null) return null;
                        return (sheetName, stripped);
                    }

                    default:
                        // #Headers / #Totals / unknown — fall through; let
                        // string key handle (no share, but no crash).
                        return null;
                }
            }
        }
        return null; // Table name not found — fall through.
    }

    private static (string sheet, string rangeRef)? ResolveDefinedName(
        WorkbookPart workbookPart, string nameToken, string? sheetScopeName)
    {
        var definedNames = workbookPart.Workbook?.GetFirstChild<DefinedNames>();
        if (definedNames == null) return null;

        // Build sheet index map for LocalSheetId resolution.
        // sheets are 0-indexed in the order they appear under <sheets>.
        var sheetByLocalId = new Dictionary<uint, string>();
        var sheetsElem = workbookPart.Workbook?.GetFirstChild<Sheets>();
        if (sheetsElem != null)
        {
            uint i = 0;
            foreach (var s in sheetsElem.Elements<Sheet>())
            {
                if (s.Name?.Value != null) sheetByLocalId[i] = s.Name.Value;
                i++;
            }
        }

        // First pass: matching name + matching scope (sheet-scoped if requested,
        // else workbook-scoped). Second pass: relax sheet scope.
        DefinedName? bestMatch = null;
        uint? scopeSheetIdHint = null;
        if (sheetScopeName != null)
        {
            foreach (var kv in sheetByLocalId)
                if (string.Equals(kv.Value, sheetScopeName, StringComparison.OrdinalIgnoreCase))
                    scopeSheetIdHint = kv.Key;
        }

        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            var dnName = dn.Name?.Value;
            if (dnName == null) continue;
            if (!string.Equals(dnName, nameToken, StringComparison.OrdinalIgnoreCase)) continue;

            // Scope priority:
            //   - If caller passed a sheetScopeName: prefer sheet-scoped to
            //     that sheet; fall back to workbook-scoped.
            //   - Else: prefer workbook-scoped (LocalSheetId == null).
            if (scopeSheetIdHint.HasValue && dn.LocalSheetId?.Value == scopeSheetIdHint.Value)
            { bestMatch = dn; break; }
            if (sheetScopeName == null && dn.LocalSheetId == null)
            { bestMatch = dn; break; }
            if (bestMatch == null) bestMatch = dn; // fallback any
        }
        if (bestMatch == null) return null;

        var body = bestMatch.Text?.Trim();
        if (string.IsNullOrEmpty(body)) return null;
        if (body.StartsWith("=")) body = body.Substring(1).Trim();

        // Multi-range bodies (commas) — fall through.
        if (body.Contains(',')) return null;
        // Dynamic bodies (OFFSET/INDEX/INDIRECT) — fall through.
        if (System.Text.RegularExpressions.Regex.IsMatch(body,
                @"\b(OFFSET|INDEX|INDIRECT|CHOOSE)\s*\(", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return null;

        // Body should look like "Sheet1!$A$1:$C$5" or "'Sheet 1'!A1:C5".
        if (!body.Contains('!')) return null;
        var parts = body.Split('!', 2);
        var sheet = parts[0].Trim();
        if (sheet.Length >= 2 && ((sheet[0] == '\'' && sheet[^1] == '\'')
                                || (sheet[0] == '"' && sheet[^1] == '"')))
            sheet = sheet.Substring(1, sheet.Length - 2);
        var rangePart = parts[1].Trim();
        if (!LooksLikeRangeRef(rangePart)) return null;
        return (sheet, rangePart.Replace("$", "").ToUpperInvariant());
    }

    private static string? LookupSheetNameForPart(
        WorkbookPart workbookPart, WorksheetPart targetWs)
    {
        var sheets = workbookPart.Workbook?.GetFirstChild<Sheets>();
        if (sheets == null) return null;
        foreach (var s in sheets.Elements<Sheet>())
        {
            var rid = s.Id?.Value;
            if (rid == null) continue;
            try
            {
                if (workbookPart.GetPartById(rid) == targetWs)
                    return s.Name?.Value;
            }
            catch { /* missing rel — ignore */ }
        }
        return null;
    }

    private static string? StripHeaderRow(string reference)
    {
        // "$A$1:$C$5" → "A2:C5" (strip $, increment start row).
        var clean = reference.Replace("$", "").ToUpperInvariant();
        var parts = clean.Split(':');
        if (parts.Length != 2) return null;
        var m1 = System.Text.RegularExpressions.Regex.Match(parts[0], @"^([A-Z]+)(\d+)$");
        var m2 = System.Text.RegularExpressions.Regex.Match(parts[1], @"^([A-Z]+)(\d+)$");
        if (!m1.Success || !m2.Success) return null;
        if (!int.TryParse(m1.Groups[2].Value, out var startRow)) return null;
        var newStart = $"{m1.Groups[1].Value}{startRow + 1}";
        return $"{newStart}:{parts[1]}";
    }
}
