// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Sle = DocumentFormat.OpenXml.Office2010.Drawing.Slicer;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Slicer (pivot-backed) ====================
    //
    // Slicers hang off an existing pivot table. The assembly involves six
    // distinct parts/elements that must all cross-reference consistently:
    //
    //   1. SlicerCachePart           (workbook-level)        — cache definition
    //   2. SlicerCacheDefinition     (root of #1)            — Name, SourceName
    //        └─ SlicerCachePivotTables/SlicerCachePivotTable  — TabId+Name ref
    //        └─ SlicerCacheData/TabularSlicerCache            — PivotCacheId ref
    //             └─ TabularSlicerCacheItems/TabularSlicerCacheItem × N
    //   3. SlicersPart               (worksheet-level)       — visual defs
    //        └─ Slicers/Slicer × 1                           — Name, Cache, RowHeight
    //   4. Workbook extLst           (WorkbookExtensionList) — registers cache
    //        uri "{BBE1A952-AA13-448e-AADC-164F8A28A991}"
    //        └─ X14.SlicerCaches/X14.SlicerCache { Id=slicerCachePartRelId }
    //   5. Worksheet extLst          (WorksheetExtensionList) — registers list
    //        uri "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}"
    //        └─ X14.SlicerList/X14.SlicerRef { Id=slicersPartRelId }
    //   6. Drawing anchor            (DrawingsPart/WorksheetDrawing)
    //        └─ AlternateContent
    //             ├─ Choice(a15) → GraphicFrame/Graphic/GraphicData(slicer uri)
    //             │                  └─ sle:slicer Name="..."
    //             └─ Fallback    → xdr:sp placeholder shape
    //
    // CONSISTENCY(pivot-dependency): slicers reference an EXISTING pivot table
    // by `pivotTable=/SheetName/pivottable[N]`. Unlike Excel's UI flow
    // (create pivot + slicer in one drag-drop), the CLI keeps these as two
    // separate operations so errors stay isolated. We mirror the pivot's
    // cache field set: the slicer's source field must match a pivotField name.

    private const string SlicerCachesExtUri   = "{BBE1A952-AA13-448e-AADC-164F8A28A991}";
    // Pivot-backed slicers use a DIFFERENT worksheet extLst URI than table-backed
    // slicers. The SDK conformance test uses {3A4CF648-...} for table-backed, but
    // Excel-generated pivot-backed files use {A8765BA9-...}. Wrong URI → Excel
    // silently strips the slicer parts on open with no schema error.
    private const string SlicerListExtUri     = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";
    private const string SlicerDrawingNsUri   = "http://schemas.microsoft.com/office/drawing/2010/slicer";
    private const string X14NsUri             = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private const string McNsUri              = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    // Pivot-backed slicer drawing uses a14 (2010/main), not a15 (2012/main).
    // Excel-generated reference files use a14; a15 gets the drawing removed.
    private const string A14NsUri             = "http://schemas.microsoft.com/office/drawing/2010/main";
    private const string XNsUri               = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Add a slicer bound to an existing pivot table field.
    /// Required props: pivotTable (path), field (field name in the pivot cache).
    /// Optional props: name, caption, columnCount, rowHeight, style, x, y, width, height.
    /// Returns the new slicer's path: /SheetName/slicer[N].
    /// </summary>
    private string AddSlicer(string parentPath, Dictionary<string, string> properties)
    {
        var segments = parentPath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var hostWorksheet = FindWorksheet(sheetName)
            ?? throw SheetNotFoundException(sheetName);

        // 1. Resolve pivot table reference ---------------------------------
        // R26-3: also accept `tableName=` as a user-friendly alias — when the
        // value isn't a path, resolve it as a pivot-table name on the host sheet.
        if (!properties.TryGetValue("pivotTable", out var pivotRef)
            && !properties.TryGetValue("pivot", out pivotRef)
            && !properties.TryGetValue("source", out pivotRef)
            && !properties.TryGetValue("tableName", out pivotRef))
        {
            throw new ArgumentException(
                "slicer requires 'pivotTable' property pointing to an existing pivot table " +
                "(e.g. pivotTable=/Sheet1/pivottable[1])");
        }
        if (!pivotRef.Contains('/') && !pivotRef.Contains('!') && !pivotRef.Contains('['))
        {
            // Bare name → search host sheet's pivot tables for a matching name.
            var hostPivots = hostWorksheet.PivotTableParts.ToList();
            int matchIdx = -1;
            for (int pi = 0; pi < hostPivots.Count; pi++)
            {
                var pn = hostPivots[pi].PivotTableDefinition?.Name?.Value;
                if (string.Equals(pn, pivotRef, StringComparison.OrdinalIgnoreCase))
                { matchIdx = pi; break; }
            }
            if (matchIdx < 0)
                throw new ArgumentException(
                    $"Pivot table named '{pivotRef}' not found on sheet '{sheetName}'.");
            pivotRef = $"/{sheetName}/pivottable[{matchIdx + 1}]";
        }

        var (pivotPart, pivotWorksheet, pivotSheetName) = ResolvePivotReference(pivotRef);
        var pivotDef = pivotPart.PivotTableDefinition
            ?? throw new ArgumentException($"Pivot table at '{pivotRef}' has no definition");
        var pivotCachePart = pivotPart.GetPartsOfType<PivotTableCacheDefinitionPart>().FirstOrDefault()
            ?? throw new ArgumentException($"Pivot table at '{pivotRef}' has no cache definition");
        var pivotCacheDef = pivotCachePart.PivotCacheDefinition
            ?? throw new ArgumentException($"Pivot table at '{pivotRef}' has no cache definition");

        // 2. Resolve field name → cacheField index -------------------------
        // R26-3: accept `column=` as an alias for `field=` (matches the
        // user-facing "filter by column" mental model).
        if ((!properties.TryGetValue("field", out var fieldName)
                && !properties.TryGetValue("column", out fieldName))
            || string.IsNullOrWhiteSpace(fieldName))
            throw new ArgumentException("slicer requires 'field' property naming a pivot field");

        var cacheFields = pivotCacheDef.GetFirstChild<CacheFields>()
            ?? throw new ArgumentException($"Pivot cache has no cacheFields");
        var cacheFieldList = cacheFields.Elements<CacheField>().ToList();
        int fieldIdx = -1;
        for (int i = 0; i < cacheFieldList.Count; i++)
        {
            if (string.Equals(cacheFieldList[i].Name?.Value, fieldName, StringComparison.OrdinalIgnoreCase))
            {
                fieldIdx = i;
                break;
            }
        }
        if (fieldIdx < 0)
        {
            var available = string.Join(", ", cacheFieldList.Select(f => f.Name?.Value ?? "?"));
            throw new ArgumentException(
                $"Field '{fieldName}' not found in pivot cache. Available: [{available}]");
        }
        // Use the real cacheField name for SourceName (exact match required by Excel)
        var sourceName = cacheFieldList[fieldIdx].Name?.Value ?? fieldName;

        // 3. Resolve slicer/cache names + collision check ------------------
        var slicerName = properties.GetValueOrDefault("name");
        if (string.IsNullOrWhiteSpace(slicerName))
            slicerName = $"Slicer_{sourceName}";
        slicerName = SanitizeSlicerName(slicerName);

        var cacheName = $"Slicer_{sourceName}";
        // Make both unique across the workbook
        var existingSlicerNames = CollectExistingSlicerNames();
        var existingCacheNames = CollectExistingSlicerCacheNames();
        slicerName = MakeUnique(slicerName, existingSlicerNames);
        cacheName = MakeUnique(cacheName, existingCacheNames);

        // 4. Pivot linkage metadata ----------------------------------------
        var pivotName = pivotDef.Name?.Value
            ?? throw new ArgumentException($"Pivot table at '{pivotRef}' has no name");
        var pivotCacheId = EnsurePivotCacheSlicerExtension(pivotCacheDef);
        var pivotTabId = GetSheetTabId(pivotWorksheet);

        // Enumerate shared items for the chosen field. Each distinct value
        // becomes one TabularSlicerCacheItem with s=true (selected=visible).
        var sharedItems = cacheFieldList[fieldIdx].SharedItems;
        int itemCount = sharedItems?.ChildElements.Count ?? 0;

        // 5. Create SlicerCachePart ---------------------------------------
        var workbookPart = _doc.WorkbookPart!;
        var slicerCachePart = workbookPart.AddNewPart<SlicerCachePart>();

        var slicerCacheDef = new X14.SlicerCacheDefinition
        {
            Name = cacheName,
            SourceName = sourceName,
            MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x" }
        };
        slicerCacheDef.AddNamespaceDeclaration("mc", McNsUri);
        slicerCacheDef.AddNamespaceDeclaration("x", XNsUri);

        var pivotTables = new X14.SlicerCachePivotTables();
        pivotTables.Append(new X14.SlicerCachePivotTable
        {
            TabId = pivotTabId,
            Name = pivotName
        });
        slicerCacheDef.Append(pivotTables);

        var tabularCache = new X14.TabularSlicerCache
        {
            PivotCacheId = pivotCacheId
        };
        var items = new X14.TabularSlicerCacheItems();
        for (int i = 0; i < itemCount; i++)
        {
            items.Append(new X14.TabularSlicerCacheItem
            {
                Atom = (uint)i,
                IsSelected = true
            });
        }
        tabularCache.Append(items);

        var slicerCacheData = new X14.SlicerCacheData();
        slicerCacheData.Append(tabularCache);
        slicerCacheDef.Append(slicerCacheData);

        slicerCachePart.SlicerCacheDefinition = slicerCacheDef;
        slicerCacheDef.Save(slicerCachePart);
        var slicerCacheRelId = workbookPart.GetIdOfPart(slicerCachePart);

        // 6. Register slicer cache in workbook extLst ---------------------
        RegisterSlicerCacheInWorkbook(workbookPart, slicerCacheRelId);

        // 6b. Register a workbook-level DefinedName placeholder for the
        // slicer. Excel expects each slicer name to have a matching
        // <definedName name="Slicer_Xxx">#N/A</definedName> entry — it's a
        // sentinel rather than a real named range, and Excel uses it to
        // guard the slicer identifier namespace.
        RegisterSlicerDefinedName(workbookPart, slicerName);

        // 7. Create SlicersPart + Slicer element on host worksheet ---------
        // If the host sheet already has a SlicersPart, reuse it so multiple
        // slicers on the same sheet share a single container (matches
        // Excel's on-disk layout).
        var slicersPart = hostWorksheet.GetPartsOfType<SlicersPart>().FirstOrDefault();
        X14.Slicers slicersContainer;
        string slicersPartRelId;
        if (slicersPart == null)
        {
            slicersPart = hostWorksheet.AddNewPart<SlicersPart>();
            slicersContainer = new X14.Slicers
            {
                MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x" }
            };
            slicersContainer.AddNamespaceDeclaration("mc", McNsUri);
            slicersContainer.AddNamespaceDeclaration("x", XNsUri);
            slicersPart.Slicers = slicersContainer;
            slicersPartRelId = hostWorksheet.GetIdOfPart(slicersPart);
            RegisterSlicerListInWorksheet(hostWorksheet, slicersPartRelId);
        }
        else
        {
            slicersContainer = slicersPart.Slicers
                ?? throw new InvalidOperationException("Existing SlicersPart has no Slicers element");
            slicersPartRelId = hostWorksheet.GetIdOfPart(slicersPart);
        }

        var rowHeight = properties.TryGetValue("rowHeight", out var rhStr)
            && uint.TryParse(rhStr, out var rh) ? rh : 225425U;
        var caption = properties.GetValueOrDefault("caption") ?? sourceName;
        // Strip XML control chars (\x00-\x08, \x0B-\x0C, \x0E-\x1F) — OOXML
        // rejects these in attribute values and Dispose() throws ArgumentException
        // on serialization. Keep the rest of the string verbatim.
        caption = StripXmlInvalidChars(caption);
        var slicerElement = new X14.Slicer
        {
            Name = slicerName,
            Cache = cacheName,
            Caption = caption,
            RowHeight = rowHeight
        };
        if (properties.TryGetValue("columnCount", out var ccStr)
            && uint.TryParse(ccStr, out var cc) && cc >= 1 && cc <= 20000)
            slicerElement.ColumnCount = cc;
        if (properties.TryGetValue("style", out var styleStr) && !string.IsNullOrWhiteSpace(styleStr))
            slicerElement.Style = styleStr;

        slicersContainer.Append(slicerElement);
        slicersContainer.Save(slicersPart);

        // 8. Add drawing anchor --------------------------------------------
        AddSlicerDrawingAnchor(hostWorksheet, slicerName, properties);

        SaveWorksheet(hostWorksheet);
        workbookPart.Workbook!.Save();

        // 9. Compute index for return path ---------------------------------
        var slicerIdx = slicersContainer.Elements<X14.Slicer>().Count();
        return $"/{sheetName}/slicer[{slicerIdx}]";
    }

    // ==================== Pivot reference resolution ====================

    private (PivotTablePart part, WorksheetPart worksheetPart, string sheetName)
        ResolvePivotReference(string pivotRef)
    {
        // Accepts: /SheetName/pivottable[N]  or  SheetName!pivottable[N]  or  just the name
        var normalized = NormalizeExcelPath(pivotRef.Trim());
        if (!normalized.StartsWith('/'))
            normalized = "/" + normalized;
        var parts = normalized.TrimStart('/').Split('/', 2);
        if (parts.Length != 2)
            throw new ArgumentException(
                $"Invalid pivotTable reference '{pivotRef}'. Expected /SheetName/pivottable[N]");
        var sheetName = parts[0];
        var worksheetPart = FindWorksheet(sheetName)
            ?? throw SheetNotFoundException(sheetName);
        var m = System.Text.RegularExpressions.Regex.Match(
            parts[1], @"^(?:pivottable|pivot)\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success)
            throw new ArgumentException(
                $"Invalid pivotTable reference '{pivotRef}'. Expected form /SheetName/pivottable[N]");
        var idx = int.Parse(m.Groups[1].Value);
        var pivotParts = worksheetPart.PivotTableParts.ToList();
        if (idx < 1 || idx > pivotParts.Count)
            throw new ArgumentException(
                $"pivottable[{idx}] out of range on sheet '{sheetName}' (have {pivotParts.Count})");
        return (pivotParts[idx - 1], worksheetPart, sheetName);
    }

    private uint GetSheetTabId(WorksheetPart worksheetPart)
    {
        var workbookPart = _doc.WorkbookPart!;
        var relId = workbookPart.GetIdOfPart(worksheetPart);
        var sheets = workbookPart.Workbook!.GetFirstChild<Sheets>()
            ?? throw new InvalidOperationException("Workbook has no Sheets element");
        var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Id?.Value == relId)
            ?? throw new InvalidOperationException(
                "Worksheet part is not referenced in workbook.sheets");
        return sheet.SheetId?.Value
            ?? throw new InvalidOperationException($"Sheet '{sheet.Name}' has no sheetId");
    }

    // ==================== Pivot cache 2010 extension ====================

    private const string PivotCache2010ExtUri = "{725AE2AE-9491-48be-B2B4-4EB974FC3084}";

    /// <summary>
    /// Ensure the pivot cache definition carries an Office 2010 pivot-cache
    /// extension carrying a random-looking uint32 as pivotCacheId. This is
    /// the ID that slicer caches reference via &lt;x14:tabular
    /// pivotCacheId="..."/&gt; — it is NOT the same as the workbook's
    /// &lt;pivotCache cacheId="..."&gt; attribute (which is an internal
    /// list index). Excel real reference files use a random 32-bit uint
    /// here. Returns the id so the caller can write it into the slicer
    /// cache. Idempotent — reuses the existing id on re-entry.
    /// </summary>
    private static uint EnsurePivotCacheSlicerExtension(PivotCacheDefinition pivotCacheDef)
    {
        // CONSISTENCY(strongly-typed-extLst): must use PivotCacheDefinitionExtensionList,
        // not the generic ExtensionList. The SDK has a distinct strongly-typed
        // class for each schema-location extLst, and on reload from disk the
        // parser produces exactly that typed instance. GetFirstChild<ExtensionList>()
        // returns null against a PivotCacheDefinitionExtensionList child — so in
        // direct-open mode (where every command re-reads the file), every slicer
        // add fails the "already exists?" check, allocates a fresh ExtensionList,
        // and appends a DUPLICATE `<extLst>` sibling. Excel then either silently
        // "repairs" the file (popping the "We found a problem" dialog) or drops
        // the cache extension entirely, breaking slicer ↔ pivot binding.
        //
        // Resident mode hid this bug: within a single handler lifetime the
        // originally-created ExtensionList stays in memory as ExtensionList (our
        // new-expression), so GetFirstChild<ExtensionList>() finds it and reuses
        // it — so single-process pipelines (like the dashboard script without an
        // intervening `close`) produced clean files while every direct-open-per-
        // command path (including the slicer-dashboard.py pattern once `close` is
        // interposed, and most external callers) produced broken files.
        //
        // Cleanup: also drop any stale ExtensionList siblings left behind by
        // older builds of this code, so re-opening an existing broken file
        // with a new write auto-heals it.
        var extList = pivotCacheDef.GetFirstChild<PivotCacheDefinitionExtensionList>();
        if (extList == null)
        {
            extList = new PivotCacheDefinitionExtensionList();
            pivotCacheDef.AppendChild(extList);
        }
        foreach (var stale in pivotCacheDef.Elements<ExtensionList>().ToList())
            stale.Remove();

        // Look for an existing x14:pivotCacheDefinition extension; reuse
        // its pivotCacheId so multiple slicers on the same pivot cache
        // all reference the same id.
        //
        // CONSISTENCY(strongly-typed-extLst): same trap as the extLst container
        // above — children of PivotCacheDefinitionExtensionList reload from
        // disk as PivotCacheDefinitionExtension (NOT the generic Extension),
        // so Elements<Extension>() misses them and we fall through to "append
        // a brand-new extension with a fresh random pivotCacheId" on every
        // second+ slicer. That leaves the pivotCache carrying multiple
        // x14:pivotCacheDefinition siblings each with its own id, while
        // individual slicerCache parts reference DIFFERENT ids — a bifurcated
        // structure Excel trips on at load time ("We found a problem ...",
        // even though the SDK validator treats each sibling as independently
        // valid). Use the strongly-typed Elements<PivotCacheDefinitionExtension>
        // so the lookup sees reloaded children.
        //
        // Also sweep any stale generic-Extension siblings produced by older
        // builds, for the same auto-heal reason as the container cleanup above.
        foreach (var staleGeneric in extList.Elements<Extension>().ToList())
            staleGeneric.Remove();

        foreach (var ext in extList.Elements<PivotCacheDefinitionExtension>())
        {
            if (ext.Uri?.Value != PivotCache2010ExtUri) continue;
            var existingDef = ext.GetFirstChild<X14.PivotCacheDefinition>();
            if (existingDef?.PivotCacheId?.HasValue == true)
                return existingDef.PivotCacheId.Value;
            // Extension exists but lacks the attribute — upgrade in place.
            var upgradeId = RandomPivotCacheId();
            if (existingDef == null)
            {
                existingDef = new X14.PivotCacheDefinition { PivotCacheId = upgradeId };
                ext.Append(existingDef);
            }
            else
            {
                existingDef.PivotCacheId = upgradeId;
            }
            return upgradeId;
        }

        var newId = RandomPivotCacheId();
        var newExt = new PivotCacheDefinitionExtension { Uri = PivotCache2010ExtUri };
        newExt.AddNamespaceDeclaration("x14", X14NsUri);
        newExt.Append(new X14.PivotCacheDefinition { PivotCacheId = newId });
        extList.Append(newExt);
        return newId;
    }

    /// <summary>
    /// Generate a random 32-bit unsigned integer in the range used by
    /// Excel-generated pivot cache ids (1 … int.MaxValue). Positive range
    /// avoids any theoretical signed-int interop issue with downstream
    /// consumers that may use Int32 internally.
    /// </summary>
    private static uint RandomPivotCacheId()
        => (uint)Random.Shared.Next(1, int.MaxValue);

    // ==================== Workbook / worksheet extLst registration ====================

    private void RegisterSlicerCacheInWorkbook(WorkbookPart workbookPart, string slicerCachePartRelId)
    {
        var workbook = workbookPart.Workbook!;
        var extList = workbook.GetFirstChild<WorkbookExtensionList>();
        if (extList == null)
        {
            extList = new WorkbookExtensionList();
            // WorkbookExtensionList must appear after most other workbook
            // children — AppendChild is correct since it's the last element.
            workbook.AppendChild(extList);
        }

        var ext = extList.Elements<WorkbookExtension>()
            .FirstOrDefault(e => e.Uri?.Value == SlicerCachesExtUri);
        X14.SlicerCaches caches;
        if (ext == null)
        {
            ext = new WorkbookExtension { Uri = SlicerCachesExtUri };
            ext.AddNamespaceDeclaration("x14", X14NsUri);
            caches = new X14.SlicerCaches();
            ext.Append(caches);
            extList.Append(ext);
        }
        else
        {
            caches = ext.GetFirstChild<X14.SlicerCaches>()
                ?? ext.AppendChild(new X14.SlicerCaches());
        }

        caches.Append(new X14.SlicerCache { Id = slicerCachePartRelId });
    }

    private static void RegisterSlicerDefinedName(WorkbookPart workbookPart, string slicerName)
    {
        var workbook = workbookPart.Workbook!;
        var definedNames = workbook.GetFirstChild<DefinedNames>();
        if (definedNames == null)
        {
            definedNames = new DefinedNames();
            // Schema order: per ECMA-376, DefinedNames appears AFTER sheets
            // / externalReferences and BEFORE calcPr / oleSize / pivotCaches
            // / extLst. Violating this order is what made Excel flag the
            // file as "corrupt and unrepairable" — Excel's workbook parser
            // aborts on out-of-order children without attempting recovery.
            // Walk the ordered list of "later" elements and insert before
            // the first one present.
            OpenXmlElement? insertBefore =
                workbook.GetFirstChild<CalculationProperties>() as OpenXmlElement
                ?? workbook.GetFirstChild<OleSize>() as OpenXmlElement
                ?? workbook.GetFirstChild<CustomWorkbookViews>() as OpenXmlElement
                ?? workbook.GetFirstChild<PivotCaches>() as OpenXmlElement
                ?? workbook.GetFirstChild<WebPublishing>() as OpenXmlElement
                ?? workbook.GetFirstChild<FileRecoveryProperties>() as OpenXmlElement
                ?? workbook.GetFirstChild<WebPublishObjects>() as OpenXmlElement
                ?? workbook.GetFirstChild<WorkbookExtensionList>() as OpenXmlElement;
            if (insertBefore != null)
                workbook.InsertBefore(definedNames, insertBefore);
            else
                workbook.AppendChild(definedNames);
        }

        // Skip if an identically-named entry already exists (idempotent).
        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (string.Equals(dn.Name?.Value, slicerName, StringComparison.Ordinal))
                return;
        }

        definedNames.Append(new DefinedName { Name = slicerName, Text = "#N/A" });
    }

    private void RegisterSlicerListInWorksheet(WorksheetPart worksheetPart, string slicersPartRelId)
    {
        var worksheet = GetSheet(worksheetPart);
        var extList = worksheet.GetFirstChild<WorksheetExtensionList>()
            ?? worksheet.AppendChild(new WorksheetExtensionList());

        var ext = extList.Elements<WorksheetExtension>()
            .FirstOrDefault(e => e.Uri?.Value == SlicerListExtUri);
        X14.SlicerList list;
        if (ext == null)
        {
            ext = new WorksheetExtension { Uri = SlicerListExtUri };
            ext.AddNamespaceDeclaration("x14", X14NsUri);
            list = new X14.SlicerList();
            ext.Append(list);
            extList.Append(ext);
        }
        else
        {
            list = ext.GetFirstChild<X14.SlicerList>()
                ?? ext.AppendChild(new X14.SlicerList());
        }

        list.Append(new X14.SlicerRef { Id = slicersPartRelId });
    }

    // ==================== Drawing anchor ====================

    private void AddSlicerDrawingAnchor(
        WorksheetPart worksheetPart, string slicerName, Dictionary<string, string> properties)
    {
        var worksheet = GetSheet(worksheetPart);
        var drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
        if (drawingsPart.WorksheetDrawing == null)
        {
            // Declare xmlns:a on the wsDr root so individual a:* elements
            // don't have to redeclare it per-element. Matches the format
            // Excel produces and avoids a theoretical renderer quirk where
            // scattered a: declarations might confuse the slicer pipeline.
            drawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
            drawingsPart.WorksheetDrawing.AddNamespaceDeclaration(
                "a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            drawingsPart.WorksheetDrawing.Save();
            if (worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
            {
                var drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
                worksheet.Append(
                    new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
            }
        }

        // Position: column/row indices like other Excel drawings. Default
        // anchor sits to the right of column D so a pivot at column A–B is
        // not covered. Width=3 cols × height=10 rows is Excel's rough
        // default slicer footprint.
        int fromCol, fromRow, toCol, toRow;
        // CONSISTENCY(ole-width-units): accept `anchor=B2:F7` as a cell
        // range (same grammar as shape/picture/chart/OLE), alongside the
        // legacy x/y/width/height form. When both are supplied, warn and
        // let anchor= win.
        if (properties.TryGetValue("anchor", out var slAnchorStr) && !string.IsNullOrWhiteSpace(slAnchorStr))
        {
            if (properties.ContainsKey("width") || properties.ContainsKey("height")
                || properties.ContainsKey("x") || properties.ContainsKey("y"))
                Console.Error.WriteLine(
                    "Warning: 'x'/'y'/'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
            if (!TryParseCellRangeAnchor(slAnchorStr, out var sxFrom, out var syFrom, out var sxTo, out var syTo))
                throw new ArgumentException($"Invalid anchor: '{slAnchorStr}'. Expected e.g. 'B2' or 'B2:F7'.");
            fromCol = sxFrom;
            fromRow = syFrom;
            if (sxTo < 0) { sxTo = fromCol + 3; syTo = fromRow + 10; }
            toCol = sxTo;
            toRow = syTo;
        }
        else
        {
            fromCol = properties.TryGetValue("x", out var xStr)
                ? ParseHelpers.SafeParseInt(xStr, "x") : 5;
            fromRow = properties.TryGetValue("y", out var yStr)
                ? ParseHelpers.SafeParseInt(yStr, "y") : 1;
            toCol = properties.TryGetValue("width", out var wStr)
                ? fromCol + ParseHelpers.SafeParseInt(wStr, "width") : fromCol + 3;
            toRow = properties.TryGetValue("height", out var hStr)
                ? fromRow + ParseHelpers.SafeParseInt(hStr, "height") : fromRow + 10;
        }

        // Reference Excel files use editAs="oneCell" for slicers (they
        // resize with the top-left cell but don't stretch). Absolute
        // positioning is valid but differs from what Excel writes.
        var anchor = new XDR.TwoCellAnchor { EditAs = XDR.EditAsValues.OneCell };
        anchor.Append(new XDR.FromMarker(
            new XDR.ColumnId(fromCol.ToString()),
            new XDR.ColumnOffset("0"),
            new XDR.RowId(fromRow.ToString()),
            new XDR.RowOffset("0")));
        anchor.Append(new XDR.ToMarker(
            new XDR.ColumnId(toCol.ToString()),
            new XDR.ColumnOffset("0"),
            new XDR.RowId(toRow.ToString()),
            new XDR.RowOffset("0")));

        // mc:AlternateContent lets older Excel clients render a fallback
        // rectangle while newer clients use the sle:slicer shape. Pivot-
        // backed slicer drawings require Choice Requires="a14" (Office
        // 2010 main) — Excel silently drops the drawing if a15 is used.
        // Namespace placement matches Excel reference files: `mc` on
        // AlternateContent, `a14` on Choice.
        var altContent = new AlternateContent();
        altContent.AddNamespaceDeclaration("mc", McNsUri);

        var choice = new AlternateContentChoice { Requires = "a14" };
        choice.AddNamespaceDeclaration("a14", A14NsUri);
        var graphicFrame = new XDR.GraphicFrame { Macro = string.Empty };

        // Allocate two unique cNvPr ids per slicer — one for the Choice
        // GraphicFrame (the one modern Excel actually renders) and one
        // for the Fallback Shape.
        //
        // Historical note: earlier code matched the reference-file
        // convention of `id="0" name=""` in the Fallback. That assumption
        // turned out to be WRONG in practice: Excel 2019+ on macOS runs
        // a drawing-wide ID-uniqueness integrity check at load time and
        // trips on duplicate `id="0"` whenever a sheet has ≥ 2 slicers
        // — the whole file pops the "We found a problem" repair dialog
        // even though the fallback shape itself is never rendered by
        // modern clients. The OOXML validator (SDK 3.x) also flagged it
        // as Sem_UniqueAttributeValue. Giving each Fallback shape its
        // own fresh id fixes both.
        //
        // The Max() scan includes Descendants of AlternateContentFallback,
        // so after adding slicer N, slicer N+1 sees the updated max and
        // keeps the monotonic allocation going.
        var nextId = drawingsPart.WorksheetDrawing
            .Descendants<XDR.NonVisualDrawingProperties>()
            .Select(p => (uint?)p.Id?.Value ?? 0u)
            .DefaultIfEmpty(1u)
            .Max() + 1;
        var fallbackId = nextId + 1;

        graphicFrame.NonVisualGraphicFrameProperties = new XDR.NonVisualGraphicFrameProperties(
            new XDR.NonVisualDrawingProperties { Id = nextId, Name = slicerName },
            new XDR.NonVisualGraphicFrameDrawingProperties());
        graphicFrame.Transform = new XDR.Transform(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = 0L, Cy = 0L });

        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = SlicerDrawingNsUri };
        var sleSlicer = new Sle.Slicer { Name = slicerName };
        sleSlicer.AddNamespaceDeclaration("sle", SlicerDrawingNsUri);
        graphicData.Append(sleSlicer);
        graphic.Append(graphicData);

        graphicFrame.Append(graphic);
        choice.Append(graphicFrame);

        var fallback = new AlternateContentFallback();
        fallback.Append(BuildSlicerFallbackShape(fallbackId, slicerName));

        altContent.Append(choice);
        altContent.Append(fallback);

        anchor.Append(altContent);
        anchor.Append(new XDR.ClientData());

        drawingsPart.WorksheetDrawing.Append(anchor);
        drawingsPart.WorksheetDrawing.Save();
    }

    private static XDR.Shape BuildSlicerFallbackShape(uint id, string slicerName)
    {
        var shape = new XDR.Shape { Macro = string.Empty, TextLink = string.Empty };

        var nvSp = new XDR.NonVisualShapeProperties();
        // The Fallback shape gets its own drawing-unique id even though
        // modern Excel never renders it — the load-time integrity check
        // walks AlternateContent/Fallback descendants too. See the
        // allocation comment at the Choice branch above for the full
        // rationale. `name` reuses the slicer name so the validator's
        // "empty name" heuristic also stays quiet; it has no visual
        // effect because the shape is schematic-only.
        nvSp.Append(new XDR.NonVisualDrawingProperties { Id = id, Name = slicerName });
        var nvSpDraw = new XDR.NonVisualShapeDrawingProperties();
        nvSpDraw.Append(new A.ShapeLocks { NoTextEdit = true });
        nvSp.Append(nvSpDraw);

        var sp = new XDR.ShapeProperties();
        var xfm = new A.Transform2D();
        xfm.Append(new A.Offset { X = 0L, Y = 0L });
        xfm.Append(new A.Extents { Cx = 1828800L, Cy = 2381250L });
        sp.Append(xfm);
        var geom = new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle };
        geom.Append(new A.AdjustValueList());
        sp.Append(geom);
        var fill = new A.SolidFill();
        fill.Append(new A.PresetColor { Val = A.PresetColorValues.White });
        sp.Append(fill);
        var outline = new A.Outline { Width = 1 };
        var outlineFill = new A.SolidFill();
        outlineFill.Append(new A.PresetColor { Val = A.PresetColorValues.Gray });
        outline.Append(outlineFill);
        sp.Append(outline);

        var tb = new XDR.TextBody();
        tb.Append(new A.BodyProperties
        {
            VerticalOverflow = A.TextVerticalOverflowValues.Clip,
            HorizontalOverflow = A.TextHorizontalOverflowValues.Clip
        });
        tb.Append(new A.ListStyle());
        var para = new A.Paragraph();
        var run = new A.Run();
        run.Append(new A.RunProperties { FontSize = 1100 });
        run.Append(new A.Text { Text = "Slicer (requires Excel 2010 or later)" });
        para.Append(run);
        tb.Append(para);

        shape.Append(nvSp);
        shape.Append(sp);
        shape.Append(tb);
        return shape;
    }

    // ==================== Name / uniqueness helpers ====================

    private static string SanitizeSlicerName(string name)
    {
        // Slicer names must be valid Excel defined-name-ish tokens: trim
        // whitespace and replace spaces with underscores so the x14:name
        // attribute passes Excel's length+character constraints.
        name = name.Trim().Replace(' ', '_');
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("slicer name cannot be empty");
        return name;
    }

    private static string MakeUnique(string baseName, HashSet<string> existing)
    {
        if (!existing.Contains(baseName))
        {
            existing.Add(baseName);
            return baseName;
        }
        for (int i = 2; ; i++)
        {
            var candidate = $"{baseName}{i}";
            if (!existing.Contains(candidate))
            {
                existing.Add(candidate);
                return candidate;
            }
        }
    }

    private HashSet<string> CollectExistingSlicerNames()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return names;
        foreach (var wsp in workbookPart.WorksheetParts)
        {
            foreach (var sp in wsp.GetPartsOfType<SlicersPart>())
            {
                if (sp.Slicers == null) continue;
                foreach (var sl in sp.Slicers.Elements<X14.Slicer>())
                    if (!string.IsNullOrEmpty(sl.Name?.Value))
                        names.Add(sl.Name!.Value!);
            }
        }
        return names;
    }

    private HashSet<string> CollectExistingSlicerCacheNames()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return names;
        foreach (var scp in workbookPart.GetPartsOfType<SlicerCachePart>())
        {
            var def = scp.SlicerCacheDefinition;
            if (def?.Name?.Value is { } n) names.Add(n);
        }
        return names;
    }

    // ==================== Readback ====================

    /// <summary>
    /// Locate a slicer by 1-based index on a sheet and resolve its backing
    /// cache definition. Returns false if the sheet has fewer slicers.
    /// </summary>
    internal bool TryFindSlicerByIndex(
        WorksheetPart worksheetPart, int index,
        out X14.Slicer? slicer, out X14.SlicerCacheDefinition? cacheDef)
    {
        slicer = null;
        cacheDef = null;
        var slicersPart = worksheetPart.GetPartsOfType<SlicersPart>().FirstOrDefault();
        if (slicersPart?.Slicers == null) return false;
        var list = slicersPart.Slicers.Elements<X14.Slicer>().ToList();
        if (index < 1 || index > list.Count) return false;
        slicer = list[index - 1];
        // Resolve the backing cache by matching Slicer.Cache → SlicerCacheDefinition.Name
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart != null && slicer.Cache?.Value is { } cacheName)
        {
            foreach (var scp in workbookPart.GetPartsOfType<SlicerCachePart>())
            {
                if (scp.SlicerCacheDefinition?.Name?.Value == cacheName)
                {
                    cacheDef = scp.SlicerCacheDefinition;
                    break;
                }
            }
        }
        return true;
    }

    internal static void ReadSlicerProperties(
        X14.Slicer slicer, X14.SlicerCacheDefinition? cacheDef, DocumentNode node)
    {
        if (slicer.Name?.Value is { } name) node.Format["name"] = name;
        if (slicer.Cache?.Value is { } cache) node.Format["cache"] = cache;
        if (slicer.Caption?.Value is { } cap) node.Format["caption"] = cap;
        if (slicer.RowHeight?.HasValue == true) node.Format["rowHeight"] = slicer.RowHeight.Value;
        if (slicer.ColumnCount?.HasValue == true) node.Format["columnCount"] = slicer.ColumnCount.Value;
        if (slicer.Style?.Value is { } style) node.Format["style"] = style;

        if (cacheDef?.SourceName?.Value is { } src) node.Format["field"] = src;
        var pivotTable = cacheDef?.SlicerCachePivotTables?
            .Elements<X14.SlicerCachePivotTable>().FirstOrDefault();
        // Schema canonical key is `pivotTable` (not `pivotTableName`).
        if (pivotTable?.Name?.Value is { } pt) node.Format["pivotTable"] = pt;
        var tabular = cacheDef?.SlicerCacheData?.GetFirstChild<X14.TabularSlicerCache>();
        if (tabular?.PivotCacheId?.HasValue == true)
            node.Format["pivotCacheId"] = tabular.PivotCacheId.Value;
        if (tabular?.TabularSlicerCacheItems != null)
            node.Format["itemCount"] = tabular.TabularSlicerCacheItems
                .Elements<X14.TabularSlicerCacheItem>().Count();
    }

    // Drop XML 1.0 illegal characters (\x00-\x08, \x0B-\x0C, \x0E-\x1F)
    // before assigning to OOXML attributes. Without this filter the value
    // round-trips through DOM but throws ArgumentException at serialize time
    // (Dispose), which surfaces as a confusing post-hoc crash.
    private static string StripXmlInvalidChars(string value)
    {
        if (string.IsNullOrEmpty(value)) return value;
        var sb = new System.Text.StringBuilder(value.Length);
        foreach (var ch in value)
            if (System.Xml.XmlConvert.IsXmlChar(ch)) sb.Append(ch);
        return sb.ToString();
    }
}
