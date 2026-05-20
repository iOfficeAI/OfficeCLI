// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

internal static partial class PivotTableHelper
{
    internal static List<string> SetPivotTableProperties(PivotTablePart pivotPart, Dictionary<string, string> properties)
    {
        // R12-2 / R12-3: normalize alias keys (row→rows, rowFields→rows,
        // columngrandtotals→colgrandtotals) so Set accepts the same aliases
        // as Add and the switch below binds to canonical keys.
        properties = NormalizePivotProperties(properties);

        // Publish sort mode for this Set operation so the re-rendered items /
        // renderers use the requested order. Sort only affects the rendered
        // layout — sharedItems order in the cache is fixed at Create time.
        using var _sortScope = PushAxisSortMode(properties);
        // CONSISTENCY(thread-static-pivot-opts): grand totals options ride
        // through the same ambient scope as sort.
        using var _gtScope = PushGrandTotalsOptions(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for subtotals.
        using var _subScope = PushSubtotalsOptions(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for layout mode.
        using var _layoutScope = PushLayoutMode(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for repeatItemLabels.
        using var _repeatScope = PushRepeatItemLabels(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for insertBlankRow.
        using var _blankRowScope = PushInsertBlankRow(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for grandTotalCaption.
        using var _captionScope = PushGrandTotalCaption(properties);

        var unsupported = new List<string>();
        var pivotDef = pivotPart.PivotTableDefinition;
        if (pivotDef == null) { unsupported.AddRange(properties.Keys); return unsupported; }

        // Seed the thread-static grand-totals scope from the CURRENT definition
        // when the caller did not explicitly pass the keys. This keeps prior
        // toggles sticky across unrelated Set operations (e.g. `set rows=...`
        // must not silently re-enable grand totals that were turned off earlier).
        // OOXML attribute → internal flag mapping:
        //   RowGrandTotals (bottom row)    → _colGrandTotals
        //   ColumnGrandTotals (right col)  → _rowGrandTotals
        if (!_rowGrandTotals.HasValue && pivotDef.ColumnGrandTotals?.Value == false)
            _rowGrandTotals = false;
        if (!_colGrandTotals.HasValue && pivotDef.RowGrandTotals?.Value == false)
            _colGrandTotals = false;

        // Seed layout sticky state: detect current layout from definition
        // attributes when the caller did not explicitly pass layout=. This keeps
        // the layout stable across unrelated Set operations (e.g. `set rows=...`
        // must not silently revert an outline pivot to compact).
        if (_layoutMode == null)
        {
            if (pivotDef.Compact?.Value == false)
            {
                var firstAxisField = pivotDef.PivotFields?.Elements<PivotField>()
                    .FirstOrDefault(pf => pf.Axis != null);
                if (firstAxisField?.Outline?.Value == false)
                    _layoutMode = "tabular";
                else
                    _layoutMode = "outline";
            }
            // else: compact (default) — _layoutMode stays null → ActiveLayoutMode returns "compact"
        }

        // Seed subtotals sticky state: if any existing row/col pivotField has
        // DefaultSubtotal=false, assume the user previously turned subtotals off
        // and the current Set (which didn't re-specify it) should preserve that.
        if (!_defaultSubtotal.HasValue && pivotDef.PivotFields != null)
        {
            foreach (var pf in pivotDef.PivotFields.Elements<PivotField>())
            {
                if (pf.DefaultSubtotal?.Value == false)
                {
                    _defaultSubtotal = false;
                    break;
                }
            }
        }

        // Collect field-area properties separately — they require a coordinated rebuild
        var fieldAreaProps = new Dictionary<string, string>();

        // R15-2: Pre-scan for field-area keys so RefreshPivotCacheFromSource
        // can skip validation of axes the same Set call is about to overwrite.
        var pendingAreaKeys = new Dictionary<string, string>();
        foreach (var (k, v) in properties)
        {
            var lk = k.ToLowerInvariant();
            if (lk == "rows" || lk == "cols" || lk == "columns" || lk == "values" || lk == "filters")
                pendingAreaKeys[lk == "columns" ? "cols" : lk] = v;
        }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    // R16-2: validate via shared helper so Set rejects
                    // empty / whitespace / control-char names just like Add.
                    // CONSISTENCY(pivot-name-validation): same rules, same
                    // error messages for both Add and Set paths.
                    pivotDef.Name = ValidatePivotName(value);
                    break;
                case "source":
                case "src":
                    // R10-1: refreshing the pivot's source range MUST also
                    // refresh the cache definition's CacheFields and the
                    // CacheRecords part. Otherwise RebuildFieldAreas reads
                    // headers from the stale cache and rejects fields that
                    // exist in the new range. Run the refresh BEFORE the
                    // field-area rebuild so any newly-added columns from the
                    // new range are visible to header validation.
                    RefreshPivotCacheFromSource(pivotPart, value, pendingAreaKeys);
                    // Force RebuildFieldAreas to run even if the caller did
                    // not pass any rows/cols/values keys, so the existing
                    // PivotField axis assignments get re-rendered against
                    // the new (possibly resized) header list.
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                        && !fieldAreaProps.ContainsKey("__sort_only__"))
                    {
                        fieldAreaProps["__sort_only__"] = "";
                    }
                    break;
                case "style":
                {
                    // Preserve existing style-info bool toggles so a bare
                    // `style=PivotStyleMedium9` does not clobber a previously-
                    // set showRowStripes=true. EnsurePivotTableStyle creates
                    // the element with defaults if absent; only the Name is
                    // overwritten here.
                    var styleInfo = EnsurePivotTableStyle(pivotDef);
                    styleInfo.Name = value;
                    break;
                }
                case "showrowstripes":
                case "showcolstripes":
                case "showcolumnstripes":
                case "showrowheaders":
                case "showcolheaders":
                case "showcolumnheaders":
                case "showlastcolumn":
                {
                    // Individual <pivotTableStyleInfo> bool toggles. Route
                    // through the shared ApplyPivotStyleInfoProps helper so
                    // Add and Set share the exact same validation + alias
                    // rules (col/column siblings) and neither path can
                    // diverge on which OOXML attribute a key maps to.
                    ApplyPivotStyleInfoProps(
                        EnsurePivotTableStyle(pivotDef),
                        new Dictionary<string, string> { [key] = value });
                    break;
                }
                case "rows":
                case "cols" or "columns":
                case "values":
                case "filters":
                    fieldAreaProps[key.ToLowerInvariant() == "columns" ? "cols" : key.ToLowerInvariant()] = value;
                    break;
                case "aggregate":
                case "showdataas":
                    // CONSISTENCY(aggregate-override / showdataas): these two
                    // sibling keys mutate per-value-field semantics. They piggy-
                    // back on the same RebuildFieldAreas pass that 'values' uses,
                    // so we hand them through verbatim and let the rebuild path
                    // (which always re-parses the value field list, even when
                    // 'values' was not in this Set call) pick them up.
                    fieldAreaProps[key.ToLowerInvariant()] = value;
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
                case "grandtotals":
                case "rowgrandtotals":
                case "colgrandtotals":
                case "columngrandtotals":
                    // Already consumed by PushGrandTotalsOptions at the top of
                    // this method. Trigger a re-render so geometry / items /
                    // cells all reflect the new toggle. Mirrors "sort".
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                        && !fieldAreaProps.ContainsKey("__sort_only__"))
                    {
                        fieldAreaProps["__sort_only__"] = value;
                    }
                    break;
                case "subtotals":
                case "defaultsubtotal":
                    // Already consumed by PushSubtotalsOptions at the top of
                    // this method. Trigger a re-render (mirrors grandtotals).
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                        && !fieldAreaProps.ContainsKey("__sort_only__"))
                    {
                        fieldAreaProps["__sort_only__"] = value;
                    }
                    break;
                case "layout":
                {
                    // Already consumed by PushLayoutMode at the top of this
                    // method. Apply definition-level + per-field attributes
                    // immediately, then trigger a re-render for geometry change
                    // (rowLabelCols depends on layout mode).
                    var lower = (value ?? "").Trim().ToLowerInvariant();
                    // Definition-level attributes
                    if (lower == "compact")
                    {
                        pivotDef.Compact = null; // revert to default true
                        pivotDef.CompactData = null;
                        pivotDef.Outline = true;
                        pivotDef.OutlineData = true;
                    }
                    else if (lower == "outline")
                    {
                        pivotDef.Compact = false;
                        pivotDef.CompactData = false;
                        pivotDef.Outline = true;
                        pivotDef.OutlineData = true;
                    }
                    else // tabular
                    {
                        pivotDef.Compact = false;
                        pivotDef.CompactData = false;
                        pivotDef.Outline = null;
                        pivotDef.OutlineData = null;
                    }
                    // Per-field attributes
                    if (pivotDef.PivotFields != null)
                    {
                        foreach (var pf in pivotDef.PivotFields.Elements<PivotField>())
                        {
                            pf.Compact = (lower == "compact") ? null : (BooleanValue)false;
                            pf.Outline = (lower == "tabular") ? (BooleanValue)false : null;
                        }
                    }
                    // Trigger re-render for geometry change
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                        && !fieldAreaProps.ContainsKey("__sort_only__"))
                    {
                        fieldAreaProps["__sort_only__"] = "";
                    }
                    break;
                }
                case "repeatlabels":
                {
                    // Write or remove the x14:pivotTableDefinition fillDownLabelsDefault
                    // extension element. Also trigger re-render so materialized cells
                    // reflect the label repetition.
                    bool enable = ParseHelpers.IsTruthy(value);
                    const string x14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
                    var extLst = pivotDef.GetFirstChild<PivotTableDefinitionExtensionList>();
                    // Remove any existing fillDownLabels extension
                    if (extLst != null)
                    {
                        var toRemove = extLst.Elements<PivotTableDefinitionExtension>()
                            .Where(e => e.Uri == "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}")
                            .ToList();
                        foreach (var e in toRemove) e.Remove();
                        if (!extLst.HasChildren) extLst.Remove();
                    }
                    if (enable)
                    {
                        var ext = new PivotTableDefinitionExtension
                        {
                            Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"
                        };
                        var x14PivotDef = new OpenXmlUnknownElement("x14", "pivotTableDefinition", x14Ns);
                        x14PivotDef.SetAttribute(new OpenXmlAttribute("fillDownLabelsDefault", "", "1"));
                        x14PivotDef.AddNamespaceDeclaration("x14", x14Ns);
                        ext.AppendChild(x14PivotDef);
                        extLst = pivotDef.GetFirstChild<PivotTableDefinitionExtensionList>()
                            ?? pivotDef.AppendChild(new PivotTableDefinitionExtensionList());
                        extLst.AppendChild(ext);
                    }
                    // Trigger re-render
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                        && !fieldAreaProps.ContainsKey("__sort_only__"))
                    {
                        fieldAreaProps["__sort_only__"] = "";
                    }
                    break;
                }
                case "grandtotalcaption":
                {
                    pivotDef.GrandTotalCaption = value?.Trim() ?? "Grand Total";
                    // Trigger re-render so materialized cells reflect the new caption
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                        && !fieldAreaProps.ContainsKey("__sort_only__"))
                    {
                        fieldAreaProps["__sort_only__"] = "";
                    }
                    break;
                }
                case "blankrows":
                {
                    bool enable = ParseHelpers.IsTruthy(value);
                    // Set insertBlankRow on the outermost row field
                    if (pivotDef.PivotFields != null && pivotDef.RowFields != null)
                    {
                        var rowFields = pivotDef.RowFields.Elements<Field>().ToList();
                        if (rowFields.Count >= 2)
                        {
                            var firstIdx = (int)(rowFields[0].Index?.Value ?? 0);
                            var pf = pivotDef.PivotFields.Elements<PivotField>()
                                .ElementAtOrDefault(firstIdx);
                            if (pf != null)
                                pf.InsertBlankRow = enable ? true : null;
                        }
                    }
                    // Trigger re-render
                    if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                        && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                        && !fieldAreaProps.ContainsKey("__sort_only__"))
                    {
                        fieldAreaProps["__sort_only__"] = "";
                    }
                    break;
                }
                default:
                {
                    // R15-4: accept `dataField{N}.showAs=<token>` as the
                    // write-side counterpart of the Get readback key. N is
                    // 1-indexed over the current DataFields list; map to
                    // the positional `showdataas` list so RebuildFieldAreas
                    // can apply the transform through its existing showAs
                    // override path. Consistency with the Get readback
                    // symmetry rule: users copy a key from Get and Set it
                    // back without learning a second vocabulary.
                    var lkDf = key.ToLowerInvariant();
                    if (lkDf.StartsWith("datafield") && lkDf.EndsWith(".showas"))
                    {
                        var idxStr = lkDf.Substring("datafield".Length,
                            lkDf.Length - "datafield".Length - ".showas".Length);
                        if (int.TryParse(idxStr, out var oneBasedIdx) && oneBasedIdx >= 1)
                        {
                            var existingDf = pivotDef.DataFields?.Elements<DataField>().ToList();
                            var dfCount = existingDf?.Count ?? 0;
                            if (oneBasedIdx > dfCount)
                                throw new ArgumentException(
                                    $"dataField{oneBasedIdx}.showAs: index out of range " +
                                    $"(1..{dfCount} data field(s) defined)");

                            // Build / extend the positional showdataas list
                            // so slot oneBasedIdx-1 carries the new token,
                            // leaving earlier slots empty (RebuildFieldAreas
                            // treats empty slot as "keep current").
                            fieldAreaProps.TryGetValue("showdataas", out var existingShow);
                            var slots = existingShow?.Split(',').Select(s => s.Trim()).ToList()
                                        ?? new List<string>();
                            while (slots.Count < oneBasedIdx) slots.Add("");
                            slots[oneBasedIdx - 1] = value;
                            fieldAreaProps["showdataas"] = string.Join(",", slots);

                            // Force RebuildFieldAreas to run even without
                            // any rows/cols/values/filters in this call.
                            if (!fieldAreaProps.ContainsKey("rows") && !fieldAreaProps.ContainsKey("cols")
                                && !fieldAreaProps.ContainsKey("values") && !fieldAreaProps.ContainsKey("filters")
                                && !fieldAreaProps.ContainsKey("__sort_only__"))
                            {
                                fieldAreaProps["__sort_only__"] = "";
                            }
                            break;
                        }
                    }
                    unsupported.Add(key);
                    break;
                }
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

        // CONSISTENCY(field-area-dedup): a field cannot be in two axes at
        // once. When a Set call moves a field into one axis, it must drop
        // out of any other axis it currently sits on. Without this dedup,
        // `set rows=X` can leave X in both currentCols and the new rows
        // list, which Excel renders as a corrupt pivotTableDefinition.
        // Precedence: the most-recently-set axis wins; areas not touched
        // in this Set call shed any field that was just claimed elsewhere.
        var valueFields = changes.ContainsKey("values")
            ? ParseValueFieldsWithWarning(changes, "values", headers)
            : currentValues;

        if (changes.ContainsKey("rows"))
        {
            colFieldIndices = colFieldIndices.Where(i => !rowFieldIndices.Contains(i)).ToList();
            filterFieldIndices = filterFieldIndices.Where(i => !rowFieldIndices.Contains(i)).ToList();
            // R15-1 parity: claimed row field also drops from values axis.
            valueFields = valueFields.Where(vf => !rowFieldIndices.Contains(vf.idx)).ToList();
        }
        if (changes.ContainsKey("cols"))
        {
            rowFieldIndices = rowFieldIndices.Where(i => !colFieldIndices.Contains(i)).ToList();
            filterFieldIndices = filterFieldIndices.Where(i => !colFieldIndices.Contains(i)).ToList();
            valueFields = valueFields.Where(vf => !colFieldIndices.Contains(vf.idx)).ToList();
        }
        if (changes.ContainsKey("filters"))
        {
            rowFieldIndices = rowFieldIndices.Where(i => !filterFieldIndices.Contains(i)).ToList();
            colFieldIndices = colFieldIndices.Where(i => !filterFieldIndices.Contains(i)).ToList();
            // R15-1: without this, `set filters=Sales` leaves Sales in both
            // DataFields and PageFields, producing a corrupt pivot with
            // duplicate assignment on the same cacheField.
            valueFields = valueFields.Where(vf => !filterFieldIndices.Contains(vf.idx)).ToList();
        }
        if (changes.ContainsKey("values"))
        {
            var valueIdxSet = valueFields.Select(vf => vf.idx).ToHashSet();
            rowFieldIndices = rowFieldIndices.Where(i => !valueIdxSet.Contains(i)).ToList();
            colFieldIndices = colFieldIndices.Where(i => !valueIdxSet.Contains(i)).ToList();
            filterFieldIndices = filterFieldIndices.Where(i => !valueIdxSet.Contains(i)).ToList();
        }

        // CONSISTENCY(aggregate-override / showdataas in Set): when only the
        // sibling keys were passed (values list unchanged), apply them to
        // the existing value-field list positionally so users can mutate
        // func / showAs without restating the whole values spec.
        if (!changes.ContainsKey("values"))
        {
            string[]? aggOverride = null;
            string[]? showOverride = null;
            if (changes.TryGetValue("aggregate", out var aggSpec) && !string.IsNullOrEmpty(aggSpec))
                aggOverride = aggSpec.Split(',').Select(s => s.Trim().ToLowerInvariant()).ToArray();
            if (changes.TryGetValue("showdataas", out var showSpec) && !string.IsNullOrEmpty(showSpec))
                showOverride = showSpec.Split(',').Select(s => s.Trim().ToLowerInvariant()).ToArray();
            if (aggOverride != null || showOverride != null)
            {
                for (int i = 0; i < valueFields.Count; i++)
                {
                    var (idx, func, showAs, name) = valueFields[i];
                    var funcChanged = false;
                    if (aggOverride != null && i < aggOverride.Length && !string.IsNullOrEmpty(aggOverride[i]))
                    {
                        if (!string.Equals(func, aggOverride[i], StringComparison.OrdinalIgnoreCase))
                            funcChanged = true;
                        func = aggOverride[i];
                    }
                    if (showOverride != null && i < showOverride.Length && !string.IsNullOrEmpty(showOverride[i]))
                        showAs = showOverride[i];
                    // R15-5: when aggregate changes, regenerate the display
                    // name so the DataField header shows "Count of Sales"
                    // instead of the stale "Sum of Sales". Only rewrite when
                    // the current name still matches the canonical
                    // "<AggDisplay> of <sourceHeader>" shape — future explicit
                    // user-provided names would then survive untouched.
                    if (funcChanged && idx >= 0 && idx < headers.Length)
                    {
                        var sourceHeader = headers[idx];
                        if (LooksLikeAutoDataFieldName(name, sourceHeader))
                            name = $"{AggregateDisplayName(func)} of {sourceHeader}";
                    }
                    valueFields[i] = (idx, func, showAs, name);
                }
            }
        }

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
            pf.DefaultSubtotal = null;
            pf.RemoveAllChildren<Items>();
            // CONSISTENCY(thread-static-pivot-opts): layout-dependent per-field
            // attributes. Mirrors BuildPivotTableDefinition per-field logic.
            var layoutMode = ActiveLayoutMode;
            pf.Compact = (layoutMode == "compact") ? null : (BooleanValue)false;
            pf.Outline = (layoutMode == "tabular") ? (BooleanValue)false : null;

            // Determine if this field's cache data is numeric (for Items generation)
            var isNumeric = IsFieldNumeric(cacheFields, i);

            bool onAxis = false;
            if (rowFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisRow;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
                onAxis = true;
            }
            else if (colFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisColumn;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
                onAxis = true;
            }
            else if (filterFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisPage;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
                onAxis = true;
            }
            else if (valueFields.Any(vf => vf.idx == i))
            {
                pf.DataField = true;
            }

            // CONSISTENCY(subtotals-opts): mirror BuildPivotTableDefinition — the
            // defaultSubtotal attribute lives on every axis field, gated on the
            // Set-time scope (seeded from existing state earlier if not passed).
            if (onAxis && !ActiveDefaultSubtotal)
                pf.DefaultSubtotal = false;
        }

        // Layer 2: Rebuild area reference lists
        // RowFields
        if (rowFieldIndices.Count > 0)
        {
            // The -2 sentinel belongs to the column axis only (dataOnRows=false
            // is the default and we never flip it). ColumnFields below adds it
            // unconditionally for valueFields.Count > 1, so do not duplicate
            // it on the row axis.
            var rf = new RowFields { Count = (uint)rowFieldIndices.Count };
            foreach (var idx in rowFieldIndices)
                rf.AppendChild(new Field { Index = idx });
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
            foreach (var (idx, func, showAs, displayName) in valueFields)
            {
                // BaseField/BaseItem: Excel ignores these when ShowDataAs is normal,
                // but Excel both emit them unconditionally on every
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
                var sda = ParseShowDataAs(showAs);
                if (sda.HasValue) dataField.ShowDataAs = sda.Value;
                if (sourceColumnNumFmtIds != null && idx >= 0 && idx < sourceColumnNumFmtIds.Length
                    && sourceColumnNumFmtIds[idx] is uint nfid)
                {
                    dataField.NumberFormatId = nfid;
                }
                // CONSISTENCY(percent-numfmt): mirror Add path — percent_* showAs
                // overrides any inherited numFmtId so values render as percentages.
                if (IsPercentShowAs(showAs))
                {
                    dataField.NumberFormatId = 10u;
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

        pivotDef.Location = BuildLocation(newGeom, rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices.Count);

        // Sync grand-totals attributes. Only touch when the caller explicitly
        // set them in this Set call (_*.HasValue); otherwise leave whatever
        // the definition already carried so repeated Sets don't clobber an
        // earlier toggle. OOXML mapping: internal _rowGrandTotals controls
        // the right column → OOXML ColumnGrandTotals; _colGrandTotals controls
        // the bottom row → OOXML RowGrandTotals.
        if (_rowGrandTotals.HasValue)
            pivotDef.ColumnGrandTotals = _rowGrandTotals.Value ? null : (BooleanValue)false;
        if (_colGrandTotals.HasValue)
            pivotDef.RowGrandTotals = _colGrandTotals.Value ? null : (BooleanValue)false;

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

                // Collapse any duplicate <row r="N"> elements produced by the
                // re-render interacting with other pivots in the same sheet.
                // See DedupeSheetDataRows docstring.
                DedupeSheetDataRows(sheetData);
            }
        }
    }

    private static List<int> ReadCurrentFieldIndices<T>(IEnumerable<T>? elements, Func<T, int> getIndex)
    {
        if (elements == null) return new List<int>();
        return elements.Select(getIndex).Where(i => i >= 0).ToList();
    }

    private static List<(int idx, string func, string showAs, string name)> ReadCurrentDataFields(DataFields? dataFields)
    {
        if (dataFields == null) return new List<(int, string, string, string)>();
        return dataFields.Elements<DataField>().Select(df => (
            idx: (int)(df.Field?.Value ?? 0),
            func: df.Subtotal?.InnerText ?? "sum",
            showAs: df.ShowDataAs?.InnerText ?? "normal",
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

        // CONSISTENCY(subtotals-opts): mirror AppendFieldItems — the trailing
        // <item t="default"/> is the field-level subtotal sentinel, gated on
        // ActiveDefaultSubtotal.
        bool emitSub = ActiveDefaultSubtotal;
        var items = new Items { Count = (uint)(count + (emitSub ? 1 : 0)) };
        for (int i = 0; i < count; i++)
            items.AppendChild(new Item { Index = (uint)i });
        if (emitSub)
            items.AppendChild(new Item { ItemType = ItemValues.Default });
        pf.AppendChild(items);
    }
}
