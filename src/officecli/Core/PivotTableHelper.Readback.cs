// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

internal static partial class PivotTableHelper
{
    // ==================== Readback ====================

    internal static void ReadPivotTableProperties(PivotTableDefinition pivotDef, DocumentNode node, PivotTablePart? pivotPart = null)
    {
        if (pivotDef.Name?.HasValue == true) node.Format["name"] = pivotDef.Name.Value;
        if (pivotDef.CacheId?.HasValue == true) node.Format["cacheId"] = pivotDef.CacheId.Value;

        var location = pivotDef.GetFirstChild<Location>();
        if (location?.Reference?.HasValue == true) node.Format["location"] = location.Reference.Value;

        // R15-3: Round-trip the source range so `Get`'s output is symmetric
        // with the `source=Sheet1!A1:C3` input form accepted by Add/Set.
        // Pull from the cache definition's WorksheetSource (Sheet + Reference);
        // emit the "Sheet!Ref" form, or just "Ref" when the sheet attribute
        // is absent (same-sheet fallback used by BuildCacheDefinition).
        if (pivotPart != null)
        {
            var cachePartForSrc = pivotPart.GetPartsOfType<PivotTableCacheDefinitionPart>().FirstOrDefault();
            var wsSrc = cachePartForSrc?.PivotCacheDefinition?.CacheSource?.WorksheetSource;
            if (wsSrc?.Reference?.HasValue == true)
            {
                var refVal = wsSrc.Reference.Value;
                var sheetVal = wsSrc.Sheet?.Value;
                node.Format["source"] = string.IsNullOrEmpty(sheetVal)
                    ? refVal!
                    : $"{sheetVal}!{refVal}";
            }
        }

        // Count fields
        var pivotFields = pivotDef.GetFirstChild<PivotFields>();
        if (pivotFields != null)
            node.Format["fieldCount"] = pivotFields.Elements<PivotField>().Count();

        // R3-1: resolve field indices to cacheField names for rowFields /
        // colFields / filters readback. dataField{N} already emits names, so
        // consistency requires the same here. Fall back to numeric index only
        // when the cache can't be loaded (defensive, should not happen for
        // well-formed files).
        string[]? fieldNames = null;
        if (pivotPart != null)
        {
            var cachePart = pivotPart.GetPartsOfType<PivotTableCacheDefinitionPart>().FirstOrDefault();
            var cacheFields = cachePart?.PivotCacheDefinition?.GetFirstChild<CacheFields>();
            if (cacheFields != null)
                fieldNames = cacheFields.Elements<CacheField>().Select(cf => cf.Name?.Value ?? "").ToArray();
        }
        string ResolveFieldName(uint idx)
        {
            if (fieldNames != null && idx < fieldNames.Length && !string.IsNullOrEmpty(fieldNames[idx]))
                return fieldNames[idx];
            return idx.ToString();
        }

        // Row fields
        var rowFields = pivotDef.RowFields;
        if (rowFields != null)
        {
            var names = rowFields.Elements<Field>().Where(f => f.Index?.Value >= 0).Select(f => ResolveFieldName((uint)f.Index!.Value)).ToList();
            if (names.Count > 0)
                // R4-1: canonical key matches input ('rows=' on Add/Set).
                // Legacy 'rowFields' output key removed in favor of single
                // canonical key per CLAUDE.md "Canonical DocumentNode.Format Rules".
                node.Format["rows"] = string.Join(",", names);
        }

        // Column fields
        var colFields = pivotDef.ColumnFields;
        if (colFields != null)
        {
            var names = colFields.Elements<Field>().Where(f => f.Index?.Value >= 0).Select(f => ResolveFieldName((uint)f.Index!.Value)).ToList();
            if (names.Count > 0)
                // R4-1: canonical key matches input ('cols=' on Add/Set).
                node.Format["cols"] = string.Join(",", names);
        }

        // Page/filter fields
        var pageFields = pivotDef.PageFields;
        if (pageFields != null)
        {
            var names = pageFields.Elements<PageField>().Select(f => f.Field?.Value ?? -1).Where(v => v >= 0).Select(v => ResolveFieldName((uint)v)).ToList();
            if (names.Count > 0)
                // R2-3: canonical key matches input ('filters=' on Add/Set).
                // Legacy 'filterFields' output key removed in favor of single
                // canonical key per CLAUDE.md "Canonical DocumentNode.Format Rules".
                node.Format["filters"] = string.Join(",", names);
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
                // CONSISTENCY(canonical-format-key): showDataAs round-trips
                // through its own structured Format key rather than being
                // packed into the dataField{N} colon string. Existing
                // dataField{N} schema (name:func:fieldIdx) stays untouched.
                // 'normal' is the absent/default value, omitted from output.
                if (df.ShowDataAs != null && df.ShowDataAs.InnerText != "normal" && !string.IsNullOrEmpty(df.ShowDataAs.InnerText))
                {
                    node.Format[$"dataField{i + 1}.showAs"] = ShowDataAsToCanonicalToken(df.ShowDataAs);
                }
            }
        }
        // CONSISTENCY(pivot-sort-readonly): the 'sortByField' Format key
        // (emitted below after the subtotals block) surfaces per-pivotField
        // SortType from real-world files (e.g. Excel-authored pivots). The
        // writer still applies 'sort=' globally and does not persist per-field
        // AutoSort — so Set can't round-trip 'sortByField'. See
        // CONSISTENCY(pivot-sort-store) v2 candidate for full AutoSort support.

        // Layout form readback. Detect from definition-level compact attribute
        // and per-pivotField outline attribute.
        // Compact = compact=true or absent (default), outline fields = default
        // Outline = compact=false, pivotField outline = default (true)
        // Tabular = compact=false, pivotField outline = false
        {
            bool defCompact = pivotDef.Compact?.Value ?? true;
            string layout = "compact";
            if (!defCompact)
            {
                var firstAxisPf = pivotFields?.Elements<PivotField>()
                    .FirstOrDefault(pf => pf.Axis != null);
                bool fieldOutline = firstAxisPf?.Outline?.Value ?? true;
                layout = fieldOutline ? "outline" : "tabular";
            }
            node.Format["layout"] = layout;
        }

        // grandTotalCaption readback
        {
            var caption = pivotDef.GrandTotalCaption?.Value;
            if (!string.IsNullOrEmpty(caption) && caption != "Grand Total")
                node.Format["grandTotalCaption"] = caption;
        }

        // insertBlankRow readback — check outermost row axis field
        if (pivotFields != null)
        {
            var rowAxisFields = pivotFields.Elements<PivotField>()
                .Where(pf => pf.Axis?.Value == PivotTableAxisValues.AxisRow)
                .ToList();
            if (rowAxisFields.Count > 0 && rowAxisFields[0].InsertBlankRow?.Value == true)
                node.Format["blankRows"] = "true";
        }

        // repeatItemLabels (fillDownLabelsDefault in x14:pivotTableDefinition)
        {
            bool repeatLabels = false;
            var extLst = pivotDef.GetFirstChild<PivotTableDefinitionExtensionList>();
            if (extLst != null)
            {
                foreach (var ext in extLst.Elements<PivotTableDefinitionExtension>())
                {
                    foreach (var child in ext.ChildElements)
                    {
                        if (child.LocalName != "pivotTableDefinition") continue;
                        // Open XML SDK v3's GetAttribute(local, ns) throws
                        // KeyNotFoundException when the attribute is absent —
                        // which is the common case here since Excel only
                        // emits fillDownLabelsDefault when the user enables
                        // "Repeat Item Labels". Enumerate attributes and
                        // tolerate absence instead.
                        var attr = child.GetAttributes()
                            .FirstOrDefault(a => a.LocalName == "fillDownLabelsDefault");
                        if (attr.Value == "1")
                        {
                            repeatLabels = true;
                            break;
                        }
                    }
                    if (repeatLabels) break;
                }
            }
            if (repeatLabels)
                node.Format["repeatLabels"] = "true";
        }

        // Style
        var styleInfo = pivotDef.PivotTableStyle;
        if (styleInfo?.Name?.HasValue == true)
            node.Format["style"] = styleInfo.Name.Value;
        // <pivotTableStyleInfo> bool toggles. Emit as "true"/"false" strings
        // for symmetry with the Set input form (accepts true/false/1/0/on/off
        // via ParsePivotStyleBool; Get emits the canonical true/false pair
        // so a round-trip Get → Set is a no-op). Defaults (row/col headers
        // on, stripes off, last column on) are surfaced explicitly rather
        // than being elided, so consumers reading the dict never have to
        // know which value is the OOXML default.
        if (styleInfo != null)
        {
            node.Format["showRowHeaders"] = (styleInfo.ShowRowHeaders?.Value ?? true) ? "true" : "false";
            node.Format["showColHeaders"] = (styleInfo.ShowColumnHeaders?.Value ?? true) ? "true" : "false";
            node.Format["showRowStripes"] = (styleInfo.ShowRowStripes?.Value ?? false) ? "true" : "false";
            node.Format["showColStripes"] = (styleInfo.ShowColumnStripes?.Value ?? false) ? "true" : "false";
            node.Format["showLastColumn"] = (styleInfo.ShowLastColumn?.Value ?? true) ? "true" : "false";
        }

        // R11-3: Grand totals readback. Both attributes default to true in
        // OOXML, so emit "true" when absent (default) and reflect explicit
        // false. Canonical key matches Add/Set input ('rowGrandTotals' /
        // 'colGrandTotals') per CLAUDE.md canonical Format rules.
        node.Format["rowGrandTotals"] = (pivotDef.RowGrandTotals?.Value ?? true) ? "true" : "false";
        node.Format["colGrandTotals"] = (pivotDef.ColumnGrandTotals?.Value ?? true) ? "true" : "false";

        // R20-1: subtotals readback. Inspect axis pivotFields (those with
        // Axis != null) and aggregate their DefaultSubtotal flags.
        // - All false  → "off"  (user set subtotals=off)
        // - All true / missing → "on"  (default OOXML behaviour)
        // - Mixed       → omit key  (per-field subtotals is a v2 feature)
        // Canonical key "subtotals" matches Add/Set input form.
        if (pivotFields != null)
        {
            var axisFields = pivotFields.Elements<PivotField>()
                .Where(pf => pf.Axis != null)
                .ToList();
            if (axisFields.Count > 0)
            {
                // DefaultSubtotal attribute defaults to true when absent (ECMA-376 § 18.10.1.69).
                var defaultSubtotalValues = axisFields
                    .Select(pf => pf.DefaultSubtotal?.Value ?? true)
                    .ToList();
                bool allOff = defaultSubtotalValues.All(v => !v);
                bool allOn  = defaultSubtotalValues.All(v => v);
                if (allOff)
                    node.Format["subtotals"] = "off";
                else if (allOn)
                    node.Format["subtotals"] = "on";
                // mixed: omit key (v2 per-field subtotals feature)
            }

            // R27-1: three per-pivotField readback surfaces, each emitted as
            // a csv of field-name or field-name:value pairs. All three keys
            // are read-only — officecli's writer doesn't yet round-trip any
            // of them, and Add/Set inputs remain untouched (see
            // CONSISTENCY(pivot-sort-readonly), CONSISTENCY(collapsed-items-readonly),
            // CONSISTENCY(axis-datafield-readonly) below). The purpose is to
            // surface real-world OOXML pivot features during query/get so
            // users inspecting files authored in Excel (or ClosedXML) don't
            // see silent information loss.
            //
            // Key names intentionally distinct from the Add/Set input form
            // ('sort=asc' is a global writer flag; 'sortByField: Name:asc'
            // is the per-field readback). Mirrors how 'rows'/'cols'/'filters'
            // emit name csvs while Add/Set takes 'rows=' etc.
            var pivotFieldList = pivotFields.Elements<PivotField>().ToList();
            var sortParts = new List<string>();
            var collapsedFieldNames = new List<string>();
            var axisAsDataFieldNames = new List<string>();
            for (int pfIdx = 0; pfIdx < pivotFieldList.Count; pfIdx++)
            {
                var pf = pivotFieldList[pfIdx];
                // CONSISTENCY(enum-innertext): SortType uses InnerText, not
                // enum equality, for the same reason as ShowDataAsToCanonicalToken.
                var sortRaw = pf.SortType?.InnerText ?? "";
                if (sortRaw == "ascending" || sortRaw == "descending")
                {
                    var name = ResolveFieldName((uint)pfIdx);
                    sortParts.Add($"{name}:{(sortRaw == "ascending" ? "asc" : "desc")}");
                }

                // CONSISTENCY(collapsed-items-readonly): item-level sd="0"
                // (showDetail=false) is the OOXML encoding for a collapsed
                // pivot row. Add/Set does not yet write these, so readback
                // is purely informational. Emitted as a csv of field names
                // that have at least one collapsed item. NOTE: the OpenXML
                // SDK exposes this attribute as Item.HideDetails (named after
                // the "hide" semantic while the XML attribute is 'sd' which
                // is "showDetail") — so we read the raw attribute value via
                // GetAttribute to avoid depending on the SDK's potentially
                // surprising property-name translation.
                var items = pf.Items;
                if (items != null)
                {
                    bool hasCollapsed = false;
                    foreach (var it in items.Elements<Item>())
                    {
                        string sdVal;
                        try { sdVal = it.GetAttribute("sd", "").Value ?? ""; }
                        catch (KeyNotFoundException) { sdVal = ""; }
                        if (sdVal == "0" || sdVal.Equals("false", StringComparison.OrdinalIgnoreCase))
                        {
                            hasCollapsed = true;
                            break;
                        }
                    }
                    if (hasCollapsed)
                        collapsedFieldNames.Add(ResolveFieldName((uint)pfIdx));
                }

                // CONSISTENCY(axis-datafield-readonly): pivotField's
                // dataField="1" attribute by itself is the standard marker
                // for any field referenced in <dataFields>, so it alone is
                // NOT interesting. The dual-role case — the one worth
                // surfacing — is when the same pivotField is ALSO on an
                // axis (rows/cols), meaning it's used both as a row/col
                // label AND as a data aggregate. ECMA-376 § 18.10.1.69.
                // Pure readback; writer does not currently set this flag.
                if (pf.Axis != null && pf.DataField?.Value == true)
                    axisAsDataFieldNames.Add(ResolveFieldName((uint)pfIdx));
            }
            if (sortParts.Count > 0)
                node.Format["sortByField"] = string.Join(",", sortParts);
            if (collapsedFieldNames.Count > 0)
                node.Format["collapsedFields"] = string.Join(",", collapsedFieldNames);
            if (axisAsDataFieldNames.Count > 0)
                node.Format["axisAsDataField"] = string.Join(",", axisAsDataFieldNames);
        }
    }

    /// <summary>
    /// R10-1: refresh a pivot's cache definition + records from a new source
    /// range spec ("Sheet1!A1:C4" or "A1:C4" — same sheet as the existing
    /// CacheSource). Replaces CacheFields, updates WorksheetSource.Reference
    /// (and Sheet if changed), rewrites the PivotTableCacheRecordsPart, and
    /// resizes pivotDef.PivotFields to match the new column count. Existing
    /// PivotField Axis/DataField assignments are reset because indices may no
    /// longer line up — RebuildFieldAreas reapplies them after this returns.
    /// </summary>
    private static void RefreshPivotCacheFromSource(PivotTablePart pivotPart, string newSourceSpec,
        Dictionary<string, string>? pendingFieldAreaProps = null)
    {
        if (string.IsNullOrWhiteSpace(newSourceSpec))
            throw new ArgumentException("source must not be empty");
        newSourceSpec = newSourceSpec.Trim();
        if (newSourceSpec.StartsWith("["))
            throw new ArgumentException(
                "External workbook references are not supported in pivot source. "
                + "Use a local sheet name (e.g. Sheet1!A1:D10)");

        var cachePart = pivotPart.GetPartsOfType<PivotTableCacheDefinitionPart>().FirstOrDefault()
            ?? throw new InvalidOperationException("Pivot table has no cache definition part");
        var cacheDef = cachePart.PivotCacheDefinition
            ?? throw new InvalidOperationException("Pivot cache definition is missing");
        var existingWsSource = cacheDef.CacheSource?.WorksheetSource
            ?? throw new InvalidOperationException("Pivot cache source is not a worksheet source");

        // Parse the new source spec.
        string newSheetName;
        string newRef;
        if (newSourceSpec.Contains('!'))
        {
            var parts = newSourceSpec.Split('!', 2);
            newSheetName = parts[0].Trim().Trim('\'', '"').Trim();
            newRef = parts[1].Trim();
        }
        else
        {
            newSheetName = existingWsSource.Sheet?.Value ?? "";
            newRef = newSourceSpec;
        }

        // Locate the source worksheet via the workbook part.
        var workbookPart = pivotPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault()
            ?.GetParentParts().OfType<WorkbookPart>().FirstOrDefault()
            ?? throw new InvalidOperationException("Workbook part not reachable from pivot table part");
        var sheetEntry = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == newSheetName)
            ?? throw new ArgumentException($"Source sheet not found: {newSheetName}");
        if (sheetEntry.Id?.Value is not string srcRelId)
            throw new InvalidOperationException("Source sheet has no relationship id");
        var sourceWsPart = workbookPart.GetPartById(srcRelId) as WorksheetPart
            ?? throw new InvalidOperationException("Source sheet relationship does not resolve to a WorksheetPart");

        // Re-read source data from the new range.
        var (headers, columnData, _) = ReadSourceData(sourceWsPart, newRef);
        if (headers.Length == 0)
            throw new ArgumentException("Source range has no data");
        if (columnData.Count == 0 || columnData[0].Length == 0)
            throw new ArgumentException("Source range has no data rows");

        // R15-2: Before mutating any cache/pivot state, validate that existing
        // row/col/value/filter field references still fit inside the new
        // (possibly narrower) header list. A silent drop or index clamp here
        // would leave the DataFields pointing past the rendered columnData,
        // crashing RenderPivotIntoSheet with ArgumentOutOfRangeException.
        // Prefer strict error over data loss: user must explicitly restate the
        // affected axes in the same Set call if they intended to drop them.
        var newFieldCount = headers.Length;
        var existingPivotDef = pivotPart.PivotTableDefinition;
        if (existingPivotDef != null)
        {
            // Axes that the same Set call is explicitly overwriting are
            // excluded from validation — their new values will be parsed
            // against the fresh headers by RebuildFieldAreas.
            bool rowsOverwritten = pendingFieldAreaProps?.ContainsKey("rows") == true;
            bool colsOverwritten = pendingFieldAreaProps?.ContainsKey("cols") == true;
            bool valuesOverwritten = pendingFieldAreaProps?.ContainsKey("values") == true;
            bool filtersOverwritten = pendingFieldAreaProps?.ContainsKey("filters") == true;

            void ValidateIndex(int idx, string axis, string fieldRef)
            {
                if (idx >= newFieldCount)
                    throw new ArgumentException(
                        $"{axis} field '{fieldRef}' (index {idx}) is out of range " +
                        $"after source narrowing to {newFieldCount} column(s). " +
                        $"Restate {axis}= in the same Set call to drop or reassign it.");
            }
            if (!valuesOverwritten && existingPivotDef.DataFields != null)
            {
                foreach (var df in existingPivotDef.DataFields.Elements<DataField>())
                {
                    var fi = (int)(df.Field?.Value ?? 0);
                    ValidateIndex(fi, "value", df.Name?.Value ?? fi.ToString());
                }
            }
            if (!rowsOverwritten && existingPivotDef.RowFields != null)
            {
                foreach (var f in existingPivotDef.RowFields.Elements<Field>())
                {
                    var fi = f.Index?.Value ?? -1;
                    if (fi >= 0) ValidateIndex(fi, "row", fi.ToString());
                }
            }
            if (!colsOverwritten && existingPivotDef.ColumnFields != null)
            {
                foreach (var f in existingPivotDef.ColumnFields.Elements<Field>())
                {
                    var fi = f.Index?.Value ?? -1;
                    // -2 sentinel is the values pseudo-field; it is not a cache index.
                    if (fi >= 0) ValidateIndex(fi, "col", fi.ToString());
                }
            }
            if (!filtersOverwritten && existingPivotDef.PageFields != null)
            {
                foreach (var f in existingPivotDef.PageFields.Elements<PageField>())
                {
                    var fi = f.Field?.Value ?? -1;
                    if (fi >= 0) ValidateIndex(fi, "filter", fi.ToString());
                }
            }
        }

        // Build a fresh cache definition (just to harvest its CacheFields,
        // fieldNumeric, and fieldValueIndex). We do NOT swap the part — only
        // its child elements — so the workbook-level <pivotCache> registration
        // and the relationship id from PivotTablePart → PivotCacheDefinitionPart
        // stay intact.
        var (freshDef, fieldNumeric, fieldValueIndex) =
            BuildCacheDefinition(newSheetName, newRef, headers, columnData, axisFieldIndices: null, dateGroups: null);

        // Replace WorksheetSource attributes in place.
        existingWsSource.Reference = newRef;
        existingWsSource.Sheet = newSheetName;

        // Replace the CacheFields child wholesale.
        var oldCacheFields = cacheDef.GetFirstChild<CacheFields>();
        var freshCacheFields = freshDef.GetFirstChild<CacheFields>()
            ?? throw new InvalidOperationException("Fresh cache definition missing CacheFields");
        freshCacheFields.Remove();
        if (oldCacheFields != null)
            cacheDef.ReplaceChild(freshCacheFields, oldCacheFields);
        else
            cacheDef.AppendChild(freshCacheFields);

        // Update the record count attribute on the cache definition.
        var newRecordCount = (uint)columnData[0].Length;
        cacheDef.RecordCount = newRecordCount;

        // Rebuild the PivotTableCacheRecordsPart in place. Drop the old part
        // (if any) and add a fresh one so the records align with the new
        // CacheFields layout.
        var oldRecordsPart = cachePart.GetPartsOfType<PivotTableCacheRecordsPart>().FirstOrDefault();
        if (oldRecordsPart != null)
            cachePart.DeletePart(oldRecordsPart);
        var newRecordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        newRecordsPart.PivotCacheRecords = BuildCacheRecords(columnData, fieldNumeric, fieldValueIndex, skipFieldIndices: null);
        newRecordsPart.PivotCacheRecords.Save();
        cacheDef.Id = cachePart.GetIdOfPart(newRecordsPart);
        cacheDef.Save();

        // Resize pivotDef.PivotFields to match the new header count. Reset
        // axis/dataField on every retained PivotField — RebuildFieldAreas
        // (called immediately after this in SetPivotTableProperties) reads
        // the new headers and reapplies axis assignments.
        var pivotDef = pivotPart.PivotTableDefinition
            ?? throw new InvalidOperationException("Pivot table definition is missing");
        var pivotFields = pivotDef.PivotFields;
        if (pivotFields == null)
        {
            pivotFields = new PivotFields();
            pivotDef.PivotFields = pivotFields;
        }
        var existingPfList = pivotFields.Elements<PivotField>().ToList();
        // Drop trailing PivotFields beyond the new column count.
        while (existingPfList.Count > headers.Length)
        {
            existingPfList[existingPfList.Count - 1].Remove();
            existingPfList.RemoveAt(existingPfList.Count - 1);
        }
        // Append fresh PivotFields for any newly-added columns.
        while (existingPfList.Count < headers.Length)
        {
            var pf = new PivotField { ShowAll = false };
            pivotFields.AppendChild(pf);
            existingPfList.Add(pf);
        }
        // Items contents on retained PivotFields are stale (they were
        // generated from the old shared-items list). RebuildFieldAreas will
        // re-generate them from the fresh CacheFields, but it only resets
        // when the field is on an axis. Wipe them now so leftover entries
        // from non-axis fields cannot be read by Excel.
        foreach (var pf in existingPfList)
        {
            pf.RemoveAllChildren<Items>();
        }
        pivotFields.Count = (uint)headers.Length;

        // RowFields / ColumnFields / PageFields / DataFields are preserved
        // here so RebuildFieldAreas can read the current assignments and
        // carry over any axes the caller did not explicitly re-specify in
        // this Set call. RebuildFieldAreas resets PivotField.Axis/DataField
        // and rewrites the area lists from scratch.
        pivotDef.Save();
    }

}
