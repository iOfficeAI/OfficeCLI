// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

internal static partial class PivotTableHelper
{
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

    // ==================== Pivot style info helpers ====================
    //
    // PivotTableStyle carries both the style NAME and five bool layout
    // toggles (showRowStripes, showColStripes, showRowHeaders,
    // showColHeaders, showLastColumn). CONSISTENCY(canonical-format-key):
    // every toggle is a first-class Set key with a canonical lowercase
    // form matching ReadPivotTableProperties output. The helper below is
    // the single ensure-or-create site so Add and Set never diverge on
    // defaults, and style-name changes preserve existing toggles.

    /// <summary>
    /// Return the pivot's existing &lt;pivotTableStyleInfo&gt; element, creating
    /// one with the project-standard defaults if absent. Callers then
    /// mutate individual attributes in place. Defaults match the hard-
    /// coded values previously duplicated in CreatePivotTable and the
    /// Set 'style' case (row/col headers on, stripes off, last column on).
    /// </summary>
    private static PivotTableStyle EnsurePivotTableStyle(PivotTableDefinition pivotDef)
    {
        if (pivotDef.PivotTableStyle == null)
        {
            pivotDef.PivotTableStyle = new PivotTableStyle
            {
                ShowRowHeaders = true,
                ShowColumnHeaders = true,
                ShowRowStripes = false,
                ShowColumnStripes = false,
                ShowLastColumn = true
            };
        }
        return pivotDef.PivotTableStyle;
    }

    /// <summary>
    /// Strict bool parser for pivot style toggles. Accepts true/false/1/0/
    /// yes/no/on/off (case-insensitive) and throws ArgumentException on
    /// anything else. CONSISTENCY(strict-enums): matches the sort-mode and
    /// showdataas reject-unknown behavior introduced in the recent pivot
    /// validation sweep — silent fallbacks mask typos.
    /// </summary>
    private static bool ParsePivotStyleBool(string key, string value)
    {
        switch ((value ?? "").Trim().ToLowerInvariant())
        {
            case "true": case "1": case "yes": case "on": return true;
            case "false": case "0": case "no": case "off": return false;
            default:
                throw new ArgumentException(
                    $"invalid {key}: '{value}'. Valid: true, false");
        }
    }

    /// <summary>
    /// Apply the five &lt;pivotTableStyleInfo&gt; bool attributes from the
    /// caller's properties dict onto an existing PivotTableStyle element.
    /// Only keys actually present in the dict are applied, so Set
    /// operations can change one toggle without clobbering the others.
    /// Accepts both canonical (showColStripes) and OOXML-verbatim
    /// (showColumnStripes) spellings for the "col/column" siblings,
    /// matching the existing alias policy.
    /// </summary>
    private static void ApplyPivotStyleInfoProps(
        PivotTableStyle styleInfo,
        Dictionary<string, string> properties)
    {
        foreach (var (rawKey, value) in properties)
        {
            switch (rawKey.ToLowerInvariant())
            {
                case "showrowstripes":
                    styleInfo.ShowRowStripes = ParsePivotStyleBool(rawKey, value);
                    break;
                case "showcolstripes":
                case "showcolumnstripes":
                    styleInfo.ShowColumnStripes = ParsePivotStyleBool(rawKey, value);
                    break;
                case "showrowheaders":
                    styleInfo.ShowRowHeaders = ParsePivotStyleBool(rawKey, value);
                    break;
                case "showcolheaders":
                case "showcolumnheaders":
                    styleInfo.ShowColumnHeaders = ParsePivotStyleBool(rawKey, value);
                    break;
                case "showlastcolumn":
                    styleInfo.ShowLastColumn = ParsePivotStyleBool(rawKey, value);
                    break;
            }
        }
    }

    private static PivotTableDefinition BuildPivotTableDefinition(
        string name, uint cacheId, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<int> filterFieldIndices, List<(int idx, string func, string showAs, string name)> valueFields,
        string styleName,
        uint?[]? columnNumFmtIds = null,
        List<DateGroupSpec>? dateGroups = null)
    {
        var pivotDef = new PivotTableDefinition
        {
            Name = name,
            CacheId = cacheId,
            DataCaption = "Values",
            CreatedVersion = 3,
            MinRefreshableVersion = 3,
            // UpdatedVersion=4 marks this pivot as "last saved by Excel 2010"
            // — the minimum required for Excel to attach slicers. With =3
            // (Excel 2007), Excel silently refuses to bind slicers to the
            // pivot table and the slicer drawing renders blank. See
            // slicer repro: only the <pivotTableDefinition updatedVersion>
            // needed to change for the slicer to appear.
            UpdatedVersion = 4,
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
            // Caption attributes — when present, Excel uses these strings instead
            // of its locale-default "Row Labels" / "Column Labels" / "Grand Total".
            // Without these the rendered cells we wrote into sheetData ("地区",
            // "产品", "总计") get visually overlaid by Excel's English defaults
            // because the pivot's caption layer takes precedence over cell content
            // when the corresponding caption attribute is empty/missing.
            RowHeaderCaption = rowFieldIndices.Count > 0 ? headers[rowFieldIndices[0]] : "Rows",
            ColumnHeaderCaption = colFieldIndices.Count > 0 ? headers[colFieldIndices[0]] : "Columns",
            GrandTotalCaption = ActiveGrandTotalCaption
        };

        // Layout-dependent attributes on PivotTableDefinition.
        // Compact: compact=default(true), outline=true, outlineData=true
        // Outline: compact=false, compactData=false, outline=true, outlineData=true
        // Tabular: compact=false, compactData=false, outline=default, outlineData=default
        var layoutMode = ActiveLayoutMode;
        if (layoutMode == "outline" || layoutMode == "tabular")
        {
            pivotDef.Compact = false;
            pivotDef.CompactData = false;
        }
        if (layoutMode != "tabular")
        {
            pivotDef.Outline = true;
            pivotDef.OutlineData = true;
        }

        // Grand totals toggles. Both attributes default to true in ECMA-376 —
        // only emit when the user opted out, matching real Excel + LibreOffice
        // serialization behavior.
        // OOXML attribute mapping (ECMA-376, empirically verified):
        //   RowGrandTotals    = BOTTOM grand total ROW  (→ internal _colGrandTotals)
        //   ColumnGrandTotals = RIGHT grand total COLUMN (→ internal _rowGrandTotals)
        if (!ActiveRowGrandTotals) pivotDef.ColumnGrandTotals = false;
        if (!ActiveColGrandTotals) pivotDef.RowGrandTotals = false;

        // Use typed property setters to ensure correct schema order

        // Compute the pivot's geometry (range + offsets) via shared helper, so the
        // initial CreatePivotTable path and the post-Set RebuildFieldAreas path
        // produce identical results.
        var geom = ComputePivotGeometry(
            position, columnData, rowFieldIndices, colFieldIndices, valueFields);
        pivotDef.Location = BuildLocation(geom, rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices.Count);

        // Page filters: presence is signalled by the <pageFields> element + the
        // pivotField axis="axisPage" marker, both written further down. ECMA-376
        // also defines optional rowPageCount / colPageCount attributes here, but
        // OpenXml SDK 3.3.0 does not model them and rejects them as unknown
        // during schema validation. Excel recognizes the filter without them
        // (verified empirically and in pivot_dark1.xlsx, which has filters but
        // no page count attributes). Tracked as a v2 polish item if any consumer
        // turns out to require them.

        // Derived date-group fields need their pivotField items count to
        // match the FIXED bucket count (month=12, quarter=4, day=31, year=
        // observed years), not just the values present in the source data.
        // Excel validates the cache groupItems count against the pivotField
        // items count and crashes if they mismatch (verified with 'months'
        // grouping — Excel for Mac hit a hard crash during parser on
        // item-count mismatch).
        var derivedFieldByIdx = new Dictionary<int, DateGroupSpec>();
        if (dateGroups != null)
            foreach (var g in dateGroups) derivedFieldByIdx[g.DerivedFieldIdx] = g;

        // PivotFields — one per source column
        var pivotFields = new PivotFields { Count = (uint)headers.Length };
        for (int i = 0; i < headers.Length; i++)
        {
            var pf = new PivotField { ShowAll = false };
            // Layout-dependent per-field attributes.
            // Compact: compact=default(true), outline=default(true)
            // Outline: compact=false, outline=default(true)
            // Tabular: compact=false, outline=false
            if (layoutMode == "outline" || layoutMode == "tabular")
                pf.Compact = false;
            if (layoutMode == "tabular")
                pf.Outline = false;
            var values = i < columnData.Count ? columnData[i] : Array.Empty<string>();
            var isNumeric = values.Length > 0 && values.All(v =>
                string.IsNullOrEmpty(v) || double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _));

            // Axis fields (row/col/filter) MUST enumerate <items> regardless of
            // whether the values look numeric. The "skip items for numeric
            // fields" optimization is only valid for data/value fields, whose
            // values are referenced directly via <n v="..."/> in cache records.
            // Row/col/filter fields are referenced by INDEX through the
            // pivotField items list, so omitting the list leaves rowItems /
            // colItems entries dangling. Failure mode verified against a
            // date-grouped pivot where year bucket values "2024"/"2025" parse
            // as numeric but render as labels — Excel showed only the grand
            // total row instead of the year hierarchy.
            // R6-2: a field can be on an axis AND a data field at the same
            // time (e.g. rows=Region values=Region:count). The axis flag and
            // the DataField flag are independent, so check each of them
            // separately instead of if/else-if which silently dropped the
            // DataField marker.
            bool isDerivedDateGroup = derivedFieldByIdx.ContainsKey(i);
            bool onAxis = false;
            if (rowFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisRow;
                onAxis = true;
                // PV4: persist axis sort as OOXML sortType="ascending|descending"
                // on each row pivotField. Previously only affected rendering
                // order at write-time; Excel reopens reset to source order.
                if (_axisSortMode is string pvSort)
                {
                    if (pvSort.Equals("desc", StringComparison.OrdinalIgnoreCase)
                        || pvSort.Equals("locale-desc", StringComparison.OrdinalIgnoreCase))
                        pf.SortType = FieldSortValues.Descending;
                    else if (pvSort.Equals("asc", StringComparison.OrdinalIgnoreCase)
                        || pvSort.Equals("locale", StringComparison.OrdinalIgnoreCase))
                        pf.SortType = FieldSortValues.Ascending;
                }
                // PV5: repeatItemLabels ("Repeat All Item Labels") lands on
                // every outer row pivotField (all row fields except the
                // innermost — repeating the leaf would be redundant). This
                // is the per-field knob; the prior workbook-wide
                // fillDownLabelsDefault ext was a default-for-future-pivots,
                // not a knob affecting the current pivot.
                if (ActiveRepeatItemLabels)
                {
                    int rowFieldPos = rowFieldIndices.IndexOf(i);
                    bool isInnermost = rowFieldPos == rowFieldIndices.Count - 1;
                    if (!isInnermost)
                    {
                        // x14 extension on pivotField: <x14:pivotField ... /> with
                        // repeatItemLabels="1" wrapped in <x:extLst><x:ext uri=...>.
                        // The attribute is a 2009 extension, not part of the
                        // base schema (Open XML SDK 3.4 PivotField has no
                        // property for it), so we synthesize the ext element.
                        const string x14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
                        var pfExt = new PivotFieldExtension
                        {
                            Uri = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}"
                        };
                        var x14Pf = new OpenXmlUnknownElement("x14", "pivotField", x14Ns);
                        x14Pf.SetAttribute(new OpenXmlAttribute("repeatItemLabels", "", "1"));
                        x14Pf.AddNamespaceDeclaration("x14", x14Ns);
                        pfExt.AppendChild(x14Pf);
                        var pfExtLst = pf.GetFirstChild<PivotFieldExtensionList>()
                            ?? pf.AppendChild(new PivotFieldExtensionList());
                        pfExtLst.AppendChild(pfExt);
                    }
                }
            }
            else if (colFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisColumn;
                onAxis = true;
            }
            else if (filterFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisPage;
                onAxis = true;
            }
            if (onAxis)
            {
                if (isDerivedDateGroup)
                    AppendFixedBucketItems(pf, derivedFieldByIdx[i]);
                else
                    AppendFieldItems(pf, values);
                // CONSISTENCY(subtotals-opts): defaultSubtotal=false on the
                // pivotField tells Excel this axis field does not contribute
                // an outer-level subtotal. Only emit the attribute when the
                // user opted out (default true matches ECMA-376).
                if (!ActiveDefaultSubtotal)
                    pf.DefaultSubtotal = false;
            }
            if (valueFields.Any(vf => vf.idx == i))
            {
                pf.DataField = true;
            }
            // insertBlankRow: Excel sets this on ALL pivotFields (not just
            // axis fields) when "Insert Blank Line After Each Item" is enabled.
            if (ActiveInsertBlankRow)
                pf.InsertBlankRow = true;

            _ = isNumeric; // kept for readability; consumed only by data fields above

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
            // Note: the synthetic <field x="-2"/> sentinel for multi-data labels
            // belongs only on the column axis (default dataOnRows=false). The
            // ColumnFields branch below unconditionally adds it when there are
            // 2+ data fields, so we must NOT also add it here.
            var rf = new RowFields();
            foreach (var idx in rowFieldIndices)
                rf.AppendChild(new Field { Index = idx });
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
            foreach (var (idx, func, showAs, displayName) in valueFields)
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
                var sda = ParseShowDataAs(showAs);
                if (sda.HasValue) dataField.ShowDataAs = sda.Value;
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
                // showDataAs=percent_* always renders as a fraction in [0,1],
                // regardless of source column format. Override to built-in
                // numFmtId 10 ("0.00%") so Excel displays "43.08%" instead of
                // the bare "0.43" the source format would produce.
                if (IsPercentShowAs(showAs))
                {
                    dataField.NumberFormatId = 10u;
                }
                df.AppendChild(dataField);
            }
            pivotDef.DataFields = df;
        }

        // Style: create with project-standard defaults via the shared
        // EnsurePivotTableStyle helper so Set and Add never diverge on
        // defaults. The caller (CreatePivotTable) overlays any user-
        // supplied style-info toggles via ApplyPivotStyleInfoProps before
        // the definition is saved.
        var styleInfo = EnsurePivotTableStyle(pivotDef);
        styleInfo.Name = styleName;

        // PV5: "Repeat All Item Labels" is set per-pivotField in the loop
        // above (pf.RepeatItemLabels = true on outer row fields), replacing
        // the previous workbook-wide x14 fillDownLabelsDefault ext which was
        // a default-for-future-pivots, not a knob for the current pivot.

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

        // CONSISTENCY(grand-totals): emit the t="grand" sentinel entries only
        // when the corresponding axis toggle is on. rowItems' grand = bottom row
        // = _colGrandTotals; colItems' grand = right column = _rowGrandTotals.
        bool emitGrand = isRow ? ActiveColGrandTotals : ActiveRowGrandTotals;

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

            int extra = 0;
            if (emitGrand)
            {
                // Grand totals: K entries marked t="grand", with i=d for d>0.
                for (int d = 0; d < dataFieldCount; d++)
                {
                    var gt = new RowItem { ItemType = ItemValues.Grand };
                    if (d > 0) gt.Index = (uint)d;
                    gt.AppendChild(new MemberPropertyIndex());
                    container.AppendChild(gt);
                }
                extra = dataFieldCount;
            }

            SetAxisCount(container, uniqueCount * dataFieldCount + extra);
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

        if (emitGrand)
        {
            // Grand total entry — omitted when the corresponding axis toggle is off.
            var grandTotal = new RowItem { ItemType = ItemValues.Grand };
            grandTotal.AppendChild(new MemberPropertyIndex());
            container.AppendChild(grandTotal);
            SetAxisCount(container, uniqueCount + 1);
        }
        else
        {
            SetAxisCount(container, uniqueCount);
        }
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

        // CONSISTENCY(subtotals-opts): subtotal position depends on layout:
        //   compact/outline: subtotal BEFORE leaves (subtotalTop)
        //   tabular: subtotal AFTER leaves (matches Excel-authored tabular pivots)
        //
        // When subtotals are on:
        //   compact/outline: outer subtotal row first, then leaves with r=1
        //   tabular: first leaf has full (outer,inner) path, rest r=1,
        //            then subtotal with t="default" after all leaves
        // When subtotals are off: first leaf has full path, rest r=1
        bool emitSubtotals = ActiveDefaultSubtotal;
        bool tabularMode = ActiveLayoutMode == "tabular";
        int count = 0;
        foreach (var (outer, inners) in groups)
        {
            var outerPivIdx = outerOrder[outer];

            if (emitSubtotals && !tabularMode)
            {
                // Compact/outline: outer subtotal row BEFORE leaves
                var outerEntry = new RowItem();
                if (outerPivIdx == 0)
                    outerEntry.AppendChild(new MemberPropertyIndex());
                else
                    outerEntry.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
                container.AppendChild(outerEntry);
                count++;
            }

            // Leaf rows for each inner of this outer.
            // In tabular mode (or when subtotals are off), the FIRST leaf of
            // each outer group spells the full (outer, inner) path; subsequent
            // leaves use r=1. In compact/outline with subtotals, every leaf
            // uses r=1 to inherit from the subtotal row above.
            for (int li = 0; li < inners.Count; li++)
            {
                var inner = inners[li];
                var innerPivIdx = innerOrder[inner];
                bool needsFullPath = (tabularMode || !emitSubtotals) && li == 0;
                var leafEntry = needsFullPath
                    ? new RowItem()
                    : new RowItem { RepeatedItemCount = 1u };
                if (needsFullPath)
                {
                    // Full (outer, inner) path.
                    if (outerPivIdx == 0)
                        leafEntry.AppendChild(new MemberPropertyIndex());
                    else
                        leafEntry.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
                }
                if (innerPivIdx == 0)
                    leafEntry.AppendChild(new MemberPropertyIndex());
                else
                    leafEntry.AppendChild(new MemberPropertyIndex { Val = innerPivIdx });
                container.AppendChild(leafEntry);
                count++;
            }

            if (emitSubtotals && tabularMode)
            {
                // Tabular: outer subtotal row AFTER leaves, with t="default"
                var subtotalEntry = new RowItem { ItemType = ItemValues.Default };
                if (outerPivIdx == 0)
                    subtotalEntry.AppendChild(new MemberPropertyIndex());
                else
                    subtotalEntry.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
                container.AppendChild(subtotalEntry);
                count++;
            }

            // insertBlankRow: emit <i t="blank"> after each group
            if (ActiveInsertBlankRow)
            {
                var blankEntry = new RowItem { ItemType = ItemValues.Blank };
                if (outerPivIdx == 0)
                    blankEntry.AppendChild(new MemberPropertyIndex());
                else
                    blankEntry.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
                container.AppendChild(blankEntry);
                count++;
            }
        }

        // CONSISTENCY(grand-totals): rowItems' grand entry = bottom grand total
        // row, gated on _colGrandTotals. Omit entirely when the user opted out.
        if (ActiveColGrandTotals)
        {
            var grand = new RowItem { ItemType = ItemValues.Grand };
            grand.AppendChild(new MemberPropertyIndex());
            container.AppendChild(grand);
            count++;
        }

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

            // CONSISTENCY(subtotals-opts): skip the per-outer subtotal column
            // block entirely when subtotals are off. Col-axis subtotals use
            // t="default" (not the bare <i> row pattern).
            if (ActiveDefaultSubtotal)
            {
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
        }

        // CONSISTENCY(grand-totals): colItems' grand entries = right grand total
        // column(s), gated on _rowGrandTotals. Omit entirely when the user opted out.
        if (ActiveRowGrandTotals)
        {
            // Grand total columns: K entries with t="grand", x=0, i=d for d>0.
            for (int d = 0; d < K; d++)
            {
                var grand = new RowItem { ItemType = ItemValues.Grand };
                if (d > 0) grand.Index = (uint)d;
                grand.AppendChild(new MemberPropertyIndex());
                container.AppendChild(grand);
                count++;
            }
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
        // CONSISTENCY(subtotals-opts): when subtotals are off, skip emitting
        // the "subtotal" entries for every internal node. Leaf entries still
        // go in as normal, and the grand sentinel is handled below based on
        // ActiveRow/ColGrandTotals.
        bool emitSubtotals = ActiveDefaultSubtotal;
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
                if (emitSubtotals)
                    entries.Add((node.Path, "subtotal"));
            }
            else if (isRow && node.Depth > 0)
            {
                // Row axis: subtotal before children.
                if (emitSubtotals)
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
        // CONSISTENCY(grand-totals): row-axis tree grand = bottom row (→ _colGrandTotals);
        // col-axis tree grand = right column (→ _rowGrandTotals). Skip the grand
        // sentinel entirely when the corresponding toggle is off.
        bool emitGrand = isRow ? ActiveColGrandTotals : ActiveRowGrandTotals;
        if (emitGrand)
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
        // CONSISTENCY(subtotals-opts): trailing <item t="default"/> is the
        // field-level subtotal sentinel. Must be omitted when defaultSubtotal=0
        // or Excel rejects with "problem with some content" validation error.
        bool emitSub = ActiveDefaultSubtotal;
        var items = new Items { Count = (uint)(unique.Count + (emitSub ? 1 : 0)) };
        for (int i = 0; i < unique.Count; i++)
            items.AppendChild(new Item { Index = (uint)i });
        if (emitSub)
            items.AppendChild(new Item { ItemType = ItemValues.Default });
        pf.AppendChild(items);
    }

    /// <summary>
    /// Append pivot field <items> for a derived date-group field. The item
    /// count MUST match the cache's groupItems count — Excel validates the
    /// two and crashes (hard parser abort on macOS) when they mismatch.
    ///
    /// cache groupItems = N buckets + 2 sentinels
    /// pivotField items = N + 2 sentinels + 1 grand-total (default)
    ///
    /// Item indices run 0..N+1 referencing groupItems directly (including
    /// the sentinels), then the final <item t="default"/> entry is the
    /// grand total row/col. Verified against /tmp/date_authored.xlsx.
    /// </summary>
    private static void AppendFixedBucketItems(PivotField pf, DateGroupSpec spec)
    {
        var buckets = ComputeDateGroupBuckets(spec);
        int totalGroupItems = buckets.Count + 2; // + leading/trailing sentinels
        var items = new Items { Count = (uint)(totalGroupItems + 1) };
        for (int i = 0; i < totalGroupItems; i++)
            items.AppendChild(new Item { Index = (uint)i });
        items.AppendChild(new Item { ItemType = ItemValues.Default });
        pf.AppendChild(items);
    }

}
