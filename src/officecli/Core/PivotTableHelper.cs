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
        // 1. Read source data to build cache
        var (headers, columnData) = ReadSourceData(sourceSheet, sourceRef);
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

        // Build cache definition
        var cacheDef = BuildCacheDefinition(sourceSheetName, sourceRef, headers, columnData);
        cachePart.PivotCacheDefinition = cacheDef;
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

        var pivotDef = BuildPivotTableDefinition(
            pivotName, cacheId, position, headers, columnData,
            rowFields, colFields, filterFields, valueFields, style);
        pivotPart.PivotTableDefinition = pivotDef;
        pivotPart.PivotTableDefinition.Save();

        // Return 1-based index
        return targetSheet.PivotTableParts.ToList().IndexOf(pivotPart) + 1;
    }

    // ==================== Source Data Reader ====================

    private static (string[] headers, List<string[]> columnData) ReadSourceData(
        WorksheetPart sourceSheet, string sourceRef)
    {
        var ws = sourceSheet.Worksheet ?? throw new InvalidOperationException("Worksheet missing");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null) return (Array.Empty<string>(), new List<string[]>());

        // Parse range "A1:D100"
        var parts = sourceRef.Replace("$", "").Split(':');
        if (parts.Length != 2) throw new ArgumentException($"Invalid source range: {sourceRef}");

        var (startCol, startRow) = ParseCellRef(parts[0]);
        var (endCol, endRow) = ParseCellRef(parts[1]);

        var startColIdx = ColToIndex(startCol);
        var endColIdx = ColToIndex(endCol);
        var colCount = endColIdx - startColIdx + 1;

        // Read all rows in range
        var rows = new List<string[]>();
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
            }
            rows.Add(values);
        }

        if (rows.Count == 0) return (Array.Empty<string>(), new List<string[]>());

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

        return (headers, columnDataList);
    }

    private static string GetCellText(Cell cell, SharedStringTablePart? sst)
    {
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

    private static PivotCacheDefinition BuildCacheDefinition(
        string sourceSheetName, string sourceRef,
        string[] headers, List<string[]> columnData)
    {
        var recordCount = columnData.Count > 0 ? columnData[0].Length : 0;

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

        // CacheFields
        var cacheFields = new CacheFields { Count = (uint)headers.Length };
        for (int i = 0; i < headers.Length; i++)
        {
            var fieldName = string.IsNullOrEmpty(headers[i]) ? $"Column{i + 1}" : headers[i];
            var values = i < columnData.Count ? columnData[i] : Array.Empty<string>();
            cacheFields.AppendChild(BuildCacheField(fieldName, values));
        }
        cacheDef.AppendChild(cacheFields);

        return cacheDef;
    }

    private static CacheField BuildCacheField(string name, string[] values)
    {
        var field = new CacheField { Name = name, NumberFormatId = 0u };
        var uniqueValues = values.Distinct().OrderBy(v => v).ToList();
        var allNumeric = values.Length > 0 && values.All(v =>
            string.IsNullOrEmpty(v) || double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _));

        var sharedItems = new SharedItems { Count = (uint)uniqueValues.Count };

        if (allNumeric && values.Any(v => !string.IsNullOrEmpty(v)))
        {
            // Numeric field — set metadata but don't enumerate all values
            var nums = values.Where(v => !string.IsNullOrEmpty(v))
                .Select(v => double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            sharedItems.ContainsSemiMixedTypes = false;
            sharedItems.ContainsString = false;
            sharedItems.ContainsNumber = true;
            sharedItems.MinValue = nums.Min();
            sharedItems.MaxValue = nums.Max();
            sharedItems.Count = 0;
        }
        else
        {
            // String field — enumerate shared items
            foreach (var v in uniqueValues)
                sharedItems.AppendChild(new StringItem { Val = v });
        }

        field.AppendChild(sharedItems);
        return field;
    }

    // ==================== Pivot Table Definition Builder ====================

    private static PivotTableDefinition BuildPivotTableDefinition(
        string name, uint cacheId, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<int> filterFieldIndices, List<(int idx, string func, string name)> valueFields,
        string styleName)
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
            Indent = 0u
        };

        // Use typed property setters to ensure correct schema order

        // Location
        pivotDef.Location = new Location
        {
            Reference = position,
            FirstHeaderRow = 1u,
            FirstDataRow = 1u,
            FirstDataColumn = (uint)rowFieldIndices.Count
        };

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

        // RowFields
        if (rowFieldIndices.Count > 0)
        {
            var rf = new RowFields { Count = (uint)rowFieldIndices.Count };
            foreach (var idx in rowFieldIndices)
                rf.AppendChild(new Field { Index = idx });
            if (valueFields.Count > 1)
                rf.AppendChild(new Field { Index = -2 });
            pivotDef.RowFields = rf;
        }

        // ColumnFields
        if (colFieldIndices.Count > 0)
        {
            var cf = new ColumnFields { Count = (uint)colFieldIndices.Count };
            foreach (var idx in colFieldIndices)
                cf.AppendChild(new Field { Index = idx });
            pivotDef.ColumnFields = cf;
        }

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
                df.AppendChild(new DataField
                {
                    Name = displayName,
                    Field = (uint)idx,
                    Subtotal = ParseSubtotal(func),
                    BaseField = 0,
                    BaseItem = 0u
                });
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

    private static void AppendFieldItems(PivotField pf, string[] values)
    {
        var unique = values.Where(v => !string.IsNullOrEmpty(v)).Distinct().OrderBy(v => v).ToList();
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
                var dfFunc = df.Subtotal?.Value.ToString().ToLowerInvariant() ?? "sum";
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
        var unsupported = new List<string>();
        var pivotDef = pivotPart.PivotTableDefinition;
        if (pivotDef == null) { unsupported.AddRange(properties.Keys); return unsupported; }

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    pivotDef.Name = value;
                    break;
                case "style":
                {
                    pivotDef.RemoveAllChildren<PivotTableStyle>();
                    pivotDef.AppendChild(new PivotTableStyle
                    {
                        Name = value,
                        ShowRowHeaders = true,
                        ShowColumnHeaders = true,
                        ShowRowStripes = false,
                        ShowColumnStripes = false,
                        ShowLastColumn = true
                    });
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        pivotDef.Save();
        return unsupported;
    }

    // ==================== Parse Helpers ====================

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
            "average" or "avg" => DataConsolidateFunctionValues.Average,
            "max" => DataConsolidateFunctionValues.Maximum,
            "min" => DataConsolidateFunctionValues.Minimum,
            "product" => DataConsolidateFunctionValues.Product,
            "stddev" => DataConsolidateFunctionValues.StandardDeviation,
            "var" => DataConsolidateFunctionValues.Variance,
            _ => DataConsolidateFunctionValues.Sum
        };
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
}
