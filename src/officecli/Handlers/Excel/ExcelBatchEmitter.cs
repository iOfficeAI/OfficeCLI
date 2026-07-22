// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

// CONSISTENCY(emit-X-mirror): scaffold mirrors WordBatchEmitter.cs /
// PptxBatchEmitter.cs — same public entry shape (full-doc + subtree
// overloads), same Get-driven transcription, same UnsupportedWarning
// contract.
//
// Excel's structural difference from docx/pptx: the document is a sparse
// grid with content addressing (/Sheet1/A1), not an ordered child list.
// That removes the positional-drift machinery (ordinal stubs, deferred
// connectors) but adds a volume problem — a 100k-cell sheet must not emit
// 100k `add cell` rows. The VALUE layer therefore rides `import` items
// (CSV baseline, replayed through ExcelHandler.Import), followed by:
//   1. corrective `set` rows for cells whose stored type import's
//      type-detection would get wrong (string "123", string "TRUE",
//      ISO-date-shaped strings, quote-prefixed text, error cells,
//      array formulas);
//   2. a FORMAT layer of `set` rows (cell styles run-length grouped into
//      ranges per row, merges, row heights, column widths, sheet-level
//      settings).
//
// PR1 scope: sheets + values + formulas + cell styles + merges + links +
// rows/cols + sheet & workbook settings + named ranges. Charts, tables,
// conditional formats, validations, comments, pivots, drawings, sparklines
// and slicers surface as unsupported_element warnings (see
// ExcelHandler.GetDumpUnsupportedFeatures).
public static partial class ExcelBatchEmitter
{
    /// <summary>
    /// Captured at emit time when a sheet carries content we cannot
    /// round-trip through the existing handler vocabulary. Mirrors
    /// PptxBatchEmitter.UnsupportedWarning.
    /// </summary>
    public sealed record UnsupportedWarning(string Element, string Path, string Reason);

    // Style keys CellToNode emits that Set (via ExcelStyleManager) accepts
    // verbatim. Anything else on a cell node is either handled specially
    // (formula/merge/link/...) or skipped via SkippedCellKeys.
    private static readonly HashSet<string> CellStyleKeyPrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        "font.", "border.", "alignment.", "protection."
    };
    private static readonly HashSet<string> CellStyleKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill", "fillPattern", "fillBg", "numberformat", "strike", "underline", "superscript", "subscript",
    };
    // Cell Format keys the emitter consumes through dedicated channels or
    // intentionally drops (derived/readonly state).
    private static readonly HashSet<string> HandledOrDerivedCellKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "type", "formula", "cachedValue", "computedValue", "evaluated", "empty",
        "merge", "link", "tooltip", "display", "arrayformula", "arrayref", "numFmtId",
        "quotePrefix", "phonetic", "image.contentType", "image.fileSize", "alt",
        "__raw", "__richruns", "__imageDataUri",
    };

    // Sheet-level Format(Get) key → Set key mapping. Only pairs verified on
    // both sides of the handler are listed; unlisted sheet keys are ignored.
    private static readonly (string GetKey, string SetKey)[] SheetSettingMap =
    {
        ("freeze", "freeze"),
        ("zoom", "zoom"),
        // Emit the canonical (Get-surface) casing — Set dispatch lowercases
        // keys, so both spellings replay identically, but dump output should
        // match what `get` shows (canonical-key convention).
        ("tabColor", "tabColor"),
        ("autoFilter", "autoFilter"),
        ("orientation", "orientation"),
        ("paperSize", "paperSize"),
        ("fitToPage", "fitToPage"),
        ("printArea", "printArea"),
        // _xlnm.Print_Titles repeating rows/cols. Mirrors the Print_Area path;
        // Set accepts the bare "1:2" / "A:A" grammar Get now surfaces.
        ("printTitleRows", "printTitleRows"),
        ("printTitleCols", "printTitleCols"),
        ("header", "header"),
        ("footer", "footer"),
        ("margin.top", "margin.top"),
        ("margin.bottom", "margin.bottom"),
        ("margin.left", "margin.left"),
        ("margin.right", "margin.right"),
        ("margin.header", "margin.header"),
        ("margin.footer", "margin.footer"),
        // sortState. Emitted LAST so it replays against a fully-populated
        // sheet (the sheet-level "sort" Set case sorts the used range). Get
        // surfaces only the sort keys ("A:desc"), not the sortState @ref
        // extent, so dump reconstructs keys-only — a partial-range sort
        // replays as a used-range sort. Faithful to the Get-surfaced canonical
        // state; strictly better than dropping sortState entirely.
        ("sort", "sort"),
    };

    // Workbook-level keys where Get key == Set case (Set.Workbook.cs).
    // calc.fullPrecision is emitted by Get even at its spec default (true),
    // so it is only carried when false.
    private static readonly string[] WorkbookSettingKeys =
    {
        "workbook.date1904", "workbook.codeName", "workbook.filterPrivacy",
        "workbook.showObjects", "workbook.backupFile", "workbook.dateCompatibility",
        "calc.mode", "calc.iterate", "calc.iterateCount", "calc.iterateDelta",
        "calc.fullCalcOnLoad", "calc.refMode",
        "activeTab", "firstSheet",
        "workbook.lockStructure", "workbook.lockWindows",
        // Legacy workbook password hash — carried verbatim (SetWorkbook's
        // workbook.passwordhash case writes it without re-hashing) so workbook
        // protection survives round-trip, mirroring sheet-level passwordHash.
        "workbook.passwordHash",
        // Core document properties. lastModifiedBy / timestamps excluded:
        // save-time stamping would flip them every replay cycle.
        "title", "author", "subject", "description", "keywords", "category",
    };

    /// <summary>Emit a full Excel workbook as a sequence of BatchItem rows.</summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitExcel(ExcelHandler xl)
    {
        var items = new List<BatchItem>();
        var warnings = new List<UnsupportedWarning>();

        EmitWorkbookSettings(xl, items, warnings);

        // Sheet-independent defined names (workbook-scoped constants/expressions
        // like TaxRate="=0.075", no sheet qualifier in the ref) are emitted
        // BEFORE any sheet so that formula cells referencing them resolve their
        // cached value on replay-import. Without this, `import`/`set` of
        // "=TaxRate*100" runs while TaxRate is undefined, so the replayed cell
        // keeps its formula but loses its cached <v> — a round-trip fidelity
        // loss that only Excel's on-open recalc (fullCalcOnLoad) masks, and not
        // for manual-calc workbooks or non-recalculating readers. Names that
        // reference a sheet range or carry a sheet scope stay in the post-sheet
        // pass (they need the sheet to exist first).
        EmitNamedRanges(xl, items, warnings, sheetIndependentOnly: true);

        var sheetNames = xl.GetDumpSheetNames();
        for (int i = 0; i < sheetNames.Count; i++)
            EmitSheet(xl, sheetNames[i], renameFirstSheet: i == 0, items, warnings);

        // Charts replay after every sheet's data baseline exists so that
        // cross-sheet series references (a chart on Sheet1 pointing at
        // Sheet2!$B$2:$B$5) resolve their numCache on replay instead of
        // keeping an empty numRef.
        foreach (var sheetName in sheetNames)
            EmitChartsPass(xl, sheetName, items, warnings);

        // Pivot tables replay LAST, after every sheet's data exists — a
        // pivot's source range routinely lives on a different sheet than the
        // pivot itself (including sheets emitted later).
        foreach (var sheetName in sheetNames)
            EmitPivotTables(xl, "/" + sheetName, xl.GetDumpPivotCount(sheetName), items, warnings);
        // Slicers bind to pivots by name — replay after the pivot pass.
        foreach (var sheetName in sheetNames)
            EmitSlicers(xl, "/" + sheetName, xl.GetDumpSlicerCount(sheetName), items, warnings);

        // Remaining names (sheet-scoped, or refs that qualify a sheet) after
        // every sheet exists. The pre-sheet pass above already emitted the
        // sheet-independent ones; this pass skips those to avoid duplicates.
        EmitNamedRanges(xl, items, warnings, sheetIndependentOnly: false, skipSheetIndependent: true);
        EmitDocPropsScan(xl, warnings);

        return (items, warnings);
    }

    /// <summary>
    /// Emit a subtree. Supported paths: `/` (full document), `/SheetName`,
    /// `/sheet[N]`. A single-sheet dump emits `add sheet` (not the
    /// rename-first-sheet form) so it can replay onto a workbook that
    /// already has content; workbook-level settings and named ranges are
    /// NOT included (they live at sibling paths — mirrors the docx/pptx
    /// subtree contract).
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitExcel(
        ExcelHandler xl, string path)
    {
        const string SupportedHint = "Supported: /, /SheetName, /sheet[N]";
        if (string.IsNullOrEmpty(path))
            throw new CliException($"dump path cannot be empty. Use '/' for the full document or a sheet path like /Sheet1. {SupportedHint}")
                { Code = "invalid_path" };
        if (path == "/") return EmitExcel(xl);

        var token = path.Trim('/');
        if (token.Length == 0 || token.Contains('/'))
            throw new CliException($"dump path not supported: {path}. {SupportedHint}")
                { Code = "unsupported_path" };

        var sheetName = xl.ResolveDumpSheetName(token)
            ?? throw new CliException($"dump path not found: {path} (no such sheet)")
                { Code = "path_not_found" };

        var items = new List<BatchItem>();
        var warnings = new List<UnsupportedWarning>();
        EmitSheet(xl, sheetName, renameFirstSheet: false, items, warnings, claimExistingSheet: true);
        EmitPivotTables(xl, "/" + sheetName, xl.GetDumpPivotCount(sheetName), items, warnings);
        EmitSlicers(xl, "/" + sheetName, xl.GetDumpSlicerCount(sheetName), items, warnings);
        return (items, warnings);
    }

    private static void EmitWorkbookSettings(ExcelHandler xl, List<BatchItem> items,
        List<UnsupportedWarning> warnings)
    {
        DocumentNode wb;
        try { wb = xl.GetDumpWorkbookNode(); }
        catch { return; }

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in WorkbookSettingKeys)
        {
            if (!wb.Format.TryGetValue(key, out var v) || v == null) continue;
            var s = FormatValue(v);
            if (s.Length > 0) props[key] = s;
        }
        if (wb.Format.TryGetValue("calc.fullPrecision", out var fp) && fp is bool fpB && !fpB)
            props["calc.fullPrecision"] = "false";
        // workbook.passwordHash (added above) now round-trips the protection
        // hash verbatim, so no lossy-password warning is needed. The
        // display-only workbook.password="***" mask is intentionally NOT
        // emitted (it is not in WorkbookSettingKeys) — replaying "***" would
        // re-hash the literal mask into a bogus password.

        if (props.Count == 0) return;
        items.Add(new BatchItem { Command = "set", Path = "/", Props = props });
    }

    // A defined name is "sheet-independent" when it is workbook-scoped and its
    // ref carries no sheet qualifier ('!') — i.e. a constant or expression such
    // as "=0.075" or a bare range on the implicit first sheet (which always
    // exists). Such names can safely be emitted before any sheet, so formula
    // cells referencing them resolve their cached value on replay-import.
    private static bool IsSheetIndependentName(DocumentNode nr, string refVal)
    {
        var scope = nr.Format.TryGetValue("scope", out var sc) ? sc as string : null;
        bool workbookScoped = string.IsNullOrEmpty(scope)
            || string.Equals(scope, "workbook", StringComparison.OrdinalIgnoreCase);
        return workbookScoped && !refVal.Contains('!');
    }

    private static void EmitNamedRanges(ExcelHandler xl, List<BatchItem> items,
        List<UnsupportedWarning> warnings,
        bool sheetIndependentOnly = false, bool skipSheetIndependent = false)
    {
        DocumentNode list;
        try { list = xl.Get("/namedrange"); }
        catch { return; }
        if (list.Children == null) return;
        // Some generators write the SAME defined name twice into workbook.xml
        // (identical name+ref+scope; schema validation tolerates it). Replay
        // would add the first and collide on the second ("already exists"),
        // failing the whole atomic batch over source rot — emit each
        // (name, ref, scope) triple once.
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var nr in list.Children)
        {
            var name = nr.Format.TryGetValue("name", out var n) ? n as string : null;
            var refVal = nr.Format.TryGetValue("ref", out var r) ? r as string : null;
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(refVal)) continue;
            // Two-pass partition: the pre-sheet pass takes only sheet-independent
            // names; the post-sheet pass takes the rest. Each name is emitted by
            // exactly one pass.
            bool independent = IsSheetIndependentName(nr, refVal!);
            if (sheetIndependentOnly && !independent) continue;
            if (skipSheetIndependent && independent) continue;
            // Excel's builtin names (_xlnm.Print_Area etc.) are carried by the
            // sheet-level printarea/printtitlerows emits — re-adding them here
            // would duplicate the defined name.
            if (name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) continue;
            // Slicer bookkeeping names (ref literally "#N/A") are created by
            // AddSlicer on replay; re-adding one collides with it.
            if (string.Equals(refVal, "#N/A", StringComparison.OrdinalIgnoreCase)) continue;
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = name!,
                ["ref"] = refVal!,
            };
            var scope = nr.Format.TryGetValue("scope", out var sc) ? sc as string : null;
            // AddNamedRange defaults a sheet-parent to sheet scope; pin the
            // scope explicitly so parent "/" + scope=<sheet> round-trips.
            props["scope"] = string.IsNullOrEmpty(scope) ? "workbook" : scope!;
            // Replay's duplicate check keys on name+scope, so a repeated
            // (name, scope) can only be emitted once regardless of ref. Warn
            // when a dropped duplicate carried a DIFFERENT ref — that one is
            // genuinely lost, not just redundant.
            var dedupKey = props["name"] + " " + props["scope"];
            if (!seenNames.Add(dedupKey))
            {
                var firstRef = items.LastOrDefault(it => it.Type == "namedrange"
                    && it.Props != null && it.Props["name"].Equals(name, StringComparison.OrdinalIgnoreCase)
                    && it.Props["scope"] == props["scope"])?.Props?["ref"];
                if (!string.Equals(firstRef, refVal, StringComparison.Ordinal))
                    warnings.Add(new UnsupportedWarning("namedrange", nr.Path ?? "/namedrange",
                        $"duplicate defined name '{name}' (same scope, different ref '{refVal}') dropped on dump — the first occurrence wins on replay"));
                continue;
            }
            if (nr.Format.TryGetValue("comment", out var cm) && cm is string cs && cs.Length > 0)
                props["comment"] = cs;
            // volatile (DefinedName.Function) — AddNamedRange consumes it.
            if (nr.Format.TryGetValue("volatile", out var vol) && IsTruthyFormatValue(vol))
                props["volatile"] = "true";
            try
            {
                items.Add(new BatchItem { Command = "add", Parent = "/", Type = "namedrange", Props = props });
            }
            catch
            {
                warnings.Add(new UnsupportedWarning("namedrange", nr.Path ?? "/namedrange",
                    $"defined name '{name}' could not be emitted"));
            }
        }
    }

    private static void EmitDocPropsScan(ExcelHandler xl, List<UnsupportedWarning> warnings)
    {
        // Core doc properties (title/author/...) are Set-able on "/" but Get
        // does not surface them on the workbook node yet — surface the gap
        // once per dump rather than silently dropping. Cheap probe via Get("/").
        // (Kept as a scan hook for parity with Word/Pptx aux-part scans.)
    }

    // ==================== Per-sheet emit ====================

    private static void EmitSheet(ExcelHandler xl, string sheetName, bool renameFirstSheet,
        List<BatchItem> items, List<UnsupportedWarning> warnings, bool claimExistingSheet = false)
    {
        var sheetPath = "/" + sheetName;
        DocumentNode sheetNode;
        try { sheetNode = xl.Get(sheetPath); }
        catch (Exception ex)
        {
            warnings.Add(new UnsupportedWarning("sheet", sheetPath, $"sheet read failed: {ex.Message}"));
            return;
        }

        // 1. Create / claim the sheet slot.
        if (renameFirstSheet)
        {
            // Replay target contract: a blank workbook (officecli create) whose
            // single placeholder sheet is claimed by renaming. `add sheet`
            // would refuse the duplicate name when the source's first sheet is
            // literally "Sheet1", and appending would leave the placeholder
            // dangling — rename sidesteps both. The rename is emitted even
            // when the name is already "Sheet1" (a no-op on replay): it makes
            // the dump self-describing, so a find/replace sheet rename over
            // the batch JSON (dump-as-template workflow) keeps working.
            items.Add(new BatchItem
            {
                Command = "set",
                Path = "/sheet[1]",
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["name"] = sheetName },
            });
        }
        else
        {
            var sheetProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["name"] = sheetName };
            // Subtree dumps replay both onto workbooks lacking the sheet
            // (clone) and back onto ones that already have it (merge) —
            // ifExists=use makes the add claim the existing sheet instead of
            // failing on the duplicate name. Full dumps stay strict: their
            // contract is a blank target, where a collision is a real error.
            if (claimExistingSheet) sheetProps["ifExists"] = "use";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/",
                Type = "sheet",
                Props = sheetProps,
            });
        }

        // 2. Value baseline: CSV import blocks + corrective typed sets.
        var rows = xl.GetDumpRowNodes(sheetName);

        // Cells whose shared-string index cannot be resolved (malformed
        // source: missing/truncated sharedStrings part) would emit their
        // INDEX as the value — warn and drop them instead of exporting
        // confidently-wrong data. Real Excel refuses such files outright.
        foreach (var rowNode in rows)
        {
            var unresolved = rowNode.Children?
                .Where(c => c.Format.ContainsKey("__unresolvedSst")).ToList();
            if (unresolved == null || unresolved.Count == 0) continue;
            foreach (var c in unresolved)
                warnings.Add(new UnsupportedWarning("cell", c.Path ?? sheetPath,
                    "shared-string index has no entry in the sharedStrings part (malformed source); the cell value cannot be resolved and was skipped"));
            rowNode.Children!.RemoveAll(c => c.Format.ContainsKey("__unresolvedSst"));
        }

        // Strip cells inside pivot-table locations from EVERY pass (values,
        // corrective sets, styles, links): they are derived render output
        // that `add pivottable` regenerates on replay — importing them as
        // static content would fight the rebuilt pivot.
        // Table totals rows are equally derived: `add table totalRow=true`
        // regenerates the label + SUBTOTAL formulas on replay. Leaving them
        // in the baseline imported them as literal data, and the replayed
        // AddTable then appended a SECOND totals row below — the ref grew a
        // row per dump→replay cycle and per-column functions reset.
        var pivotRects = xl.GetDumpPivotLocations(sheetName)
            .Concat(xl.GetDumpTableTotalRowRects(sheetName))
            .Select(ParseRangeRect)
            .Where(r => r != null)
            .Select(r => r!.Value)
            .ToList();
        if (pivotRects.Count > 0)
        {
            foreach (var rowNode in rows)
            {
                rowNode.Children?.RemoveAll(cell =>
                {
                    var cellRef = LastPathSegment(cell.Path);
                    return TryParseCellRef(cellRef, out var c, out var r)
                        && pivotRects.Any(rect => c >= rect.C1 && c <= rect.C2 && r >= rect.R1 && r <= rect.R2);
                });
            }
        }
        var corrective = new List<BatchItem>();
        var styleRows = new List<BatchItem>();

        EmitValueBaseline(sheetName, rows, items, corrective, warnings);

        // 3. Cell-level passes off the same row snapshot: corrective typed
        // sets first (they finalize VALUES), then styles (they only touch xf).
        items.AddRange(corrective);

        foreach (var rowNode in rows)
        {
            if (rowNode.Children == null) continue;
            EmitRowCellStyles(sheetName, rowNode, styleRows);
            foreach (var cell in rowNode.Children)
            {
                if (cell.Format.TryGetValue("link", out var lv) && lv is string link && link.Length > 0)
                {
                    var linkProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["link"] = link };
                    if (cell.Format.TryGetValue("tooltip", out var tt) && tt is string tts && tts.Length > 0)
                        linkProps["tooltip"] = tts;
                    if (cell.Format.TryGetValue("display", out var dp) && dp is string dps && dps.Length > 0)
                        linkProps["display"] = dps;
                    styleRows.Add(new BatchItem { Command = "set", Path = cell.Path, Props = linkProps });
                }
                if (cell.Format.TryGetValue("phonetic", out _))
                    warnings.Add(new UnsupportedWarning("phonetic", cell.Path ?? sheetPath,
                        "phonetic (furigana) guides are not round-tripped by dump"));
            }
        }

        items.AddRange(styleRows);

        // 4. Merges — read from the worksheet's MergeCells element (per-cell
        // Format["merge"] misses all-empty merges); one sheet-level set
        // carrying every range.
        var mergeRanges = xl.GetDumpMergeRanges(sheetName);
        if (mergeRanges.Count > 0)
            items.Add(new BatchItem
            {
                Command = "set",
                Path = sheetPath,
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["merge"] = string.Join(",", mergeRanges),
                },
            });

        // 5. Row heights / hidden / outline.
        foreach (var rowNode in rows)
        {
            var rp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (rowNode.Format.TryGetValue("height", out var h) && h is string hs) rp["height"] = hs;
            if (rowNode.Format.TryGetValue("hidden", out var hd) && hd is bool hb && hb) rp["hidden"] = "true";
            if (rowNode.Format.TryGetValue("outlineLevel", out var rolv)) rp["outline"] = FormatValue(rolv);
            if (rowNode.Format.TryGetValue("collapsed", out var rc) && rc is bool rcb && rcb) rp["collapsed"] = "true";
            // These flags are echoed by Get as truthy strings (or bools) but
            // were absent from this allowlist, so replay silently dropped them.
            foreach (var flagKey in new[] { "bestFit", "thickTop", "thickBot", "ph" })
                if (rowNode.Format.TryGetValue(flagKey, out var fv) && IsTruthyFormatValue(fv))
                    rp[flagKey] = "true";
            // Row-level default style (CT_Row @s + @customFormat). Get surfaces
            // both, but they were absent from this allowlist, so dump dropped
            // the row's default cellXf — the style governing the row's phantom
            // (never-materialized) cells. A persistent style attribute Excel
            // does NOT regenerate, so the loss is permanent once round-tripped.
            // customFormat is the on/off gate; s carries the style index.
            if (rowNode.Format.TryGetValue("s", out var rs))
            {
                var rsS = FormatValue(rs);
                if (rsS.Length > 0) rp["s"] = rsS;
            }
            if (rowNode.Format.TryGetValue("customFormat", out var rcf) && IsTruthyFormatValue(rcf))
                rp["customFormat"] = "1";
            if (rp.Count > 0)
                items.Add(new BatchItem { Command = "set", Path = rowNode.Path, Props = rp });
        }

        // 6. Column widths / hidden / outline.
        var colNodes = xl.GetDumpColumnNodes(sheetName, out var colsTruncated);
        if (colsTruncated)
            warnings.Add(new UnsupportedWarning("column", sheetPath,
                "a column definition spans more than 256 columns; the tail was dropped"));
        foreach (var colNode in colNodes)
        {
            var cp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // Raw whole-column style index (<col style="N">) when it carries
            // non-numberformat facets — GetDumpColumnNodes surfaces it as
            // `style`; SetColumn's default branch writes col@style. Otherwise
            // the numberformat code below carries a numberformat-only column.
            if (colNode.Format.TryGetValue("style", out var cstyle)) cp["style"] = FormatValue(cstyle);
            if (colNode.Format.TryGetValue("width", out var w)) cp["width"] = FormatValue(w);
            if (colNode.Format.TryGetValue("hidden", out var chd) && chd is bool chb && chb) cp["hidden"] = "true";
            if (colNode.Format.TryGetValue("outlineLevel", out var colv)) cp["outline"] = FormatValue(colv);
            if (colNode.Format.TryGetValue("collapsed", out var cc) && cc is bool ccb && ccb) cp["collapsed"] = "true";
            if (colNode.Format.TryGetValue("bestFit", out var cbf) && IsTruthyFormatValue(cbf)) cp["bestFit"] = "true";
            // Column-level number format (col @s -> cellXf -> numFmt). Set's
            // col[X] handler accepts numberformat and re-registers the style.
            if (colNode.Format.TryGetValue("numberformat", out var cnf) && cnf is string cnfS && cnfS.Length > 0)
                cp["numberformat"] = cnfS;
            if (cp.Count > 0)
                items.Add(new BatchItem { Command = "set", Path = colNode.Path, Props = cp });
        }

        // 7. Sheet-level settings (freeze/zoom/tab color/autofilter/print...).
        EmitSheetSettings(xl, sheetNode, sheetPath, items, warnings);

        // 8. Structured elements: tables, conditional formats, validations,
        // comments, charts, sparklines. After data + styles so referenced
        // ranges are populated. See ExcelBatchEmitter.Elements.cs.
        EmitSheetElements(xl, sheetName, items, warnings);

        // 9. Unsupported-content scan.
        foreach (var (element, reason) in xl.GetDumpUnsupportedFeatures(sheetName))
            warnings.Add(new UnsupportedWarning(element, sheetPath, reason));
    }

    private static void EmitSheetSettings(ExcelHandler xl, DocumentNode sheetNode, string sheetPath,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (getKey, setKey) in SheetSettingMap)
        {
            if (!sheetNode.Format.TryGetValue(getKey, out var v) || v == null) continue;
            var s = FormatValue(v);
            if (s.Length > 0) props[setKey] = s;
        }
        // Toggle-off keys: Get emits gridlines/headings only when false.
        if (sheetNode.Format.TryGetValue("gridlines", out var gl) && gl is bool glB && !glB)
            props["showgridlines"] = "false";
        if (sheetNode.Format.TryGetValue("headings", out var hdg) && hdg is bool hdB && !hdB)
            props["showrowcolheaders"] = "false";
        if (sheetNode.Format.TryGetValue("direction", out var dir) && dir is string ds
            && ds.Equals("rtl", StringComparison.OrdinalIgnoreCase))
            props["rtl"] = "true";
        if (props.Count > 0)
            items.Add(new BatchItem { Command = "set", Path = sheetPath, Props = props });

        // Visibility & protection ride separate rows: hiding a sheet mid-way
        // through its own settings row could interact with active-tab logic,
        // and protect must land LAST so earlier replay sets aren't blocked.
        if (sheetNode.Format.TryGetValue("visibility", out var vis) && vis is string visS && visS.Length > 0)
            items.Add(new BatchItem
            {
                Command = "set",
                Path = sheetPath,
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["visibility"] = visS },
            });
        if (sheetNode.Format.TryGetValue("protect", out var prot) && prot is bool pb && pb)
        {
            var protProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["protect"] = "true" };
            // Round-trip the stored hashes verbatim (legacy password attr and
            // the modern algorithm set) — previously the password was dropped
            // and replay silently degraded to passwordless protection.
            foreach (var hashKey in new[]
            {
                "passwordHash", "protection.algorithm", "protection.hash",
                "protection.salt", "protection.spinCount",
            })
            {
                if (sheetNode.Format.TryGetValue(hashKey, out var hv) && hv?.ToString() is { Length: > 0 } hvS)
                    protProps[hashKey] = hvS;
            }
            items.Add(new BatchItem
            {
                Command = "set",
                Path = sheetPath,
                Props = protProps,
            });
        }

        // Manual page breaks. Get exposes them as comma-joined index lists;
        // replay re-adds each one (rowbreak row=N / colbreak col=N).
        // Manual page breaks carry an optional restricted column/row span
        // (<brk min max>). The sheet-level list gives only the break index, so
        // fetch each break's node to recover min/max and replay a non-full
        // span (otherwise the break silently widened to sheet-default on
        // round-trip). Emitted inline on the add (AddRowBreak/AddColBreak now
        // accept min/max).
        EmitPageBreaks(xl, sheetNode, sheetPath, "rowBreaks", "rowbreak", "row", items);
        EmitPageBreaks(xl, sheetNode, sheetPath, "colBreaks", "colbreak", "col", items);
    }

    private static void EmitPageBreaks(ExcelHandler xl, DocumentNode sheetNode, string sheetPath,
        string getKey, string addType, string idKey, List<BatchItem> items)
    {
        if (!sheetNode.Format.TryGetValue(getKey, out var bkObj) || bkObj is not string bkS || bkS.Length == 0)
            return;
        var ids = bkS.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        for (int i = 0; i < ids.Length; i++)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { [idKey] = ids[i] };
            try
            {
                var brkNode = xl.Get($"{sheetPath}/{addType}[{i + 1}]");
                if (brkNode.Format.TryGetValue("min", out var mn)) props["min"] = FormatValue(mn);
                if (brkNode.Format.TryGetValue("max", out var mx)) props["max"] = FormatValue(mx);
            }
            catch { /* fall back to the bare index add */ }
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = addType, Props = props });
        }
    }

    // ==================== Value baseline (CSV import) ====================

    // Split the sheet into row blocks separated by gaps of more than
    // GapRowThreshold contentless rows, so a sparse sheet with data at A1
    // and A100000 doesn't produce a 100k-line CSV of empty rows.
    private const int GapRowThreshold = 50;
    // Rows per import item — bounds the CSV string a single BatchItem carries.
    private const int MaxRowsPerImport = 20_000;

    private sealed record CsvCell(int Col, string Text);

    private static void EmitValueBaseline(string sheetName, List<DocumentNode> rows,
        List<BatchItem> items, List<BatchItem> corrective, List<UnsupportedWarning> warnings)
    {
        // Collect CSV-safe payloads per row index.
        var csvRows = new SortedDictionary<uint, List<CsvCell>>();

        foreach (var rowNode in rows)
        {
            if (rowNode.Children == null) continue;
            foreach (var cell in rowNode.Children)
            {
                var cellRef = LastPathSegment(cell.Path);
                if (!TryParseCellRef(cellRef, out var colIdx, out var rowIdx)) continue;
                var payload = ClassifyCell(cell, corrective, warnings);
                if (payload == null) continue;
                if (!csvRows.TryGetValue(rowIdx, out var list))
                    csvRows[rowIdx] = list = new List<CsvCell>();
                list.Add(new CsvCell(colIdx, payload));
            }
        }

        if (csvRows.Count == 0) return;

        // Walk row indices, grouping into blocks split on large gaps.
        var block = new List<(uint Row, List<CsvCell> Cells)>();
        uint prevRow = 0;
        foreach (var (rowIdx, cells) in csvRows)
        {
            var gapExceeded = block.Count > 0 && rowIdx - prevRow - 1 > GapRowThreshold;
            if (gapExceeded || block.Count >= MaxRowsPerImport)
            {
                FlushImportBlock(sheetName, block, items);
                block.Clear();
            }
            block.Add((rowIdx, cells));
            prevRow = rowIdx;
        }
        FlushImportBlock(sheetName, block, items);
    }

    private static void FlushImportBlock(string sheetName,
        List<(uint Row, List<CsvCell> Cells)> block, List<BatchItem> items)
    {
        if (block.Count == 0) return;
        var minCol = block.Min(b => b.Cells.Min(c => c.Col));
        var startRow = block[0].Row;
        var sb = new StringBuilder();
        uint expectedRow = startRow;
        foreach (var (rowIdx, cells) in block)
        {
            // In-block gaps within the threshold are bridged with blank rows
            // so row alignment holds (import writes rows sequentially from
            // the start cell). The bridge line must be "," (two empty
            // fields), NOT an empty line — ParseCsv drops single-empty-field
            // rows entirely, which would shift every later row up by one.
            while (expectedRow < rowIdx) { sb.Append(",\n"); expectedRow++; }
            var byCol = cells.OrderBy(c => c.Col).ToList();
            int cursor = minCol;
            for (int i = 0; i < byCol.Count; i++)
            {
                while (cursor < byCol[i].Col) { sb.Append(','); cursor++; }
                sb.Append(EscapeCsvField(byCol[i].Text));
                // Field content written; the NEXT cell needs a separator.
                cursor++;
                if (i < byCol.Count - 1) sb.Append(',');
            }
            sb.Append('\n');
            expectedRow++;
        }
        items.Add(new BatchItem
        {
            Command = "import",
            Parent = "/" + sheetName,
            Text = sb.ToString(),
            Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["start-cell"] = $"{ColumnName(minCol)}{startRow}",
            },
        });
    }

    /// <summary>
    /// Decide how a cell's VALUE replays: return the CSV field text, or null
    /// when the cell has no value payload (style-only) or is handled through
    /// a corrective set row appended to <paramref name="corrective"/>.
    /// </summary>
    private static string? ClassifyCell(DocumentNode cell, List<BatchItem> corrective,
        List<UnsupportedWarning> warnings)
    {
        var raw = cell.Format.TryGetValue("__raw", out var rv) ? rv as string : null;
        var type = cell.Format.TryGetValue("type", out var tv) ? tv as string ?? "" : "";
        var formula = cell.Format.TryGetValue("formula", out var fv) ? fv as string : null;
        var quotePrefix = cell.Format.TryGetValue("quotePrefix", out var qp) && qp is bool qb && qb;

        // Rich-text cells replay through the set type=richtext runs=<json>
        // vocabulary (runs serialized by DumpSupport); the CSV baseline would
        // flatten the per-run formatting.
        if (cell.Format.TryGetValue("__richruns", out var rr) && rr is string runsJson && runsJson.Length > 2)
        {
            corrective.Add(new BatchItem
            {
                Command = "set",
                Path = cell.Path,
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["type"] = "richtext",
                    ["runs"] = runsJson,
                },
            });
            return null;
        }

        // In-cell images cannot ride the CSV value baseline. Replay their
        // package bytes through the existing set image=<data-uri> vocabulary.
        if (type == "Image")
        {
            if (cell.Format.TryGetValue("__imageDataUri", out var imageValue)
                && imageValue is string imageDataUri && imageDataUri.Length > 0)
            {
                var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["image"] = imageDataUri,
                };
                if (cell.Format.TryGetValue("alt", out var altValue)
                    && altValue is string alt && alt.Length > 0)
                    props["alt"] = alt;
                corrective.Add(new BatchItem { Command = "set", Path = cell.Path, Props = props });
            }
            else
            {
                warnings.Add(new UnsupportedWarning("cell.image", cell.Path ?? "",
                    "in-cell image part could not be read; skipped"));
            }
            return null;
        }

        // Array formulas replay through `set arrayformula=` on the anchor
        // cell only; interior cells of the array range carry the same formula
        // + a Reference on the anchor. Non-anchor cells emit nothing (the
        // spill fills them).
        if (formula != null && cell.Format.TryGetValue("arrayformula", out var af) && af is bool afb && afb)
        {
            var arrayRef = cell.Format.TryGetValue("arrayref", out var ar) ? ar as string : null;
            var cellRef = LastPathSegment(cell.Path);
            var anchor = arrayRef?.Split(':')[0];
            if (arrayRef == null || string.Equals(anchor, cellRef, StringComparison.OrdinalIgnoreCase))
            {
                var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["arrayformula"] = formula,
                };
                if (arrayRef != null) props["ref"] = arrayRef;
                corrective.Add(new BatchItem { Command = "set", Path = cell.Path, Props = props });
            }
            return null;
        }

        if (formula != null)
            return "=" + formula.TrimStart('=');

        var text = cell.Text ?? "";
        if (string.IsNullOrEmpty(raw) && text.Length == 0)
        {
            // Explicit empty-STRING cell (<c t="str"><v/></c>): COUNTA counts
            // it and ISBLANK returns FALSE, so dropping it changes formula
            // results on replay. There is no set vocabulary for it (set
            // value= clears the cell), so replay through add cell, which
            // writes the same t="str"/<v/> shape. A bare typeless cell
            // (no value, no string type) is semantically absent — Excel
            // itself prunes those on save — and stays dropped.
            if (type is "String" or "SharedString" or "InlineString"
                && cell.Path is { } emptyPath && emptyPath.LastIndexOf('/') > 0)
            {
                corrective.Add(new BatchItem
                {
                    Command = "add",
                    Parent = emptyPath[..emptyPath.LastIndexOf('/')],
                    Type = "cell",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["ref"] = LastPathSegment(emptyPath),
                        ["value"] = "",
                    },
                });
            }
            return null;
        }

        switch (type)
        {
            case "Number":
            case "Date":
                // Raw stored text (serial for dates) — bypasses display
                // formatting AND import's ISO-date detection; the date
                // numberformat rides the style layer.
                return raw ?? text;
            case "Boolean":
                return raw == "1" ? "TRUE" : "FALSE";
            case "Error":
                corrective.Add(new BatchItem
                {
                    Command = "set",
                    Path = cell.Path,
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["value"] = raw ?? text,
                        ["type"] = "error",
                    },
                });
                return null;
            default: // String / SharedString / InlineString
                if (quotePrefix)
                {
                    // Reproduce the leading-apostrophe idiom: value='<text>
                    // stores the literal + stamps quotePrefix on the xf.
                    corrective.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cell.Path,
                        Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["value"] = "'" + text,
                        },
                    });
                    return null;
                }
                if (text.StartsWith('='))
                {
                    // `set value==...` unconditionally coerces to formula and
                    // import does the same — no vocabulary reproduces a bare
                    // string "=..." without a quote prefix.
                    warnings.Add(new UnsupportedWarning("cell.literal-equals", cell.Path ?? "",
                        "string cell text starts with '='; replayed with a quote prefix (') to keep it literal"));
                    corrective.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cell.Path,
                        Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["value"] = "'" + text,
                        },
                    });
                    return null;
                }
                if (NeedsStringCorrection(text))
                {
                    // Import's type detection would store this as number /
                    // bool / date; pin the stored type explicitly instead.
                    corrective.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cell.Path,
                        Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["value"] = text,
                            ["type"] = "string",
                        },
                    });
                    return null;
                }
                return text;
        }
    }

    /// <summary>
    /// True when import's SetCellValueWithTypeDetection would store the text
    /// as anything other than a plain string (mirrors its detection order:
    /// number → ISO date → boolean; '=' is handled by the caller). Leading
    /// apostrophes must also be pinned — import stores them literally but
    /// `set value='x` strips them, so keep the CSV free of that ambiguity.
    /// </summary>
    private static bool NeedsStringCorrection(string text)
    {
        if (text.Length == 0) return false;
        if (text.StartsWith('\'')) return true;
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out _)) return true;
        if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
            || text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return true;
        // ISO-date shapes import converts to serials.
        string[] formats =
        {
            "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fff", "yyyy-MM-ddTHH:mm:ss.fffZ", "yyyy-MM-dd HH:mm:ss",
        };
        if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _)) return true;
        return false;
    }

    // ==================== Cell style layer ====================

    private static void EmitRowCellStyles(string sheetName, DocumentNode rowNode, List<BatchItem> styleRows)
    {
        if (rowNode.Children == null || rowNode.Children.Count == 0) return;

        // Run-length group consecutive cells (same row) with byte-identical
        // style prop dictionaries into A1:D1-style range sets.
        var runs = new List<(int StartCol, int EndCol, uint Row, Dictionary<string, string> Props)>();
        foreach (var cell in rowNode.Children)
        {
            var cellRef = LastPathSegment(cell.Path);
            if (!TryParseCellRef(cellRef, out var colIdx, out var rowIdx)) continue;
            var styleProps = ExtractStyleProps(cell);
            if (styleProps.Count == 0) continue;
            if (runs.Count > 0)
            {
                var (s, e, r, p) = runs[^1];
                if (r == rowIdx && colIdx == e + 1 && SamePropDict(p, styleProps))
                {
                    runs[^1] = (s, colIdx, r, p);
                    continue;
                }
            }
            runs.Add((colIdx, colIdx, rowIdx, styleProps));
        }

        foreach (var (startCol, endCol, row, props) in runs)
        {
            var path = startCol == endCol
                ? $"/{sheetName}/{ColumnName(startCol)}{row}"
                : $"/{sheetName}/{ColumnName(startCol)}{row}:{ColumnName(endCol)}{row}";
            styleRows.Add(new BatchItem { Command = "set", Path = path, Props = props });
        }
    }

    private static Dictionary<string, string> ExtractStyleProps(DocumentNode cell)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, value) in cell.Format)
        {
            if (value == null) continue;
            if (HandledOrDerivedCellKeys.Contains(key)) continue;
            var isStyle = CellStyleKeys.Contains(key)
                || CellStyleKeyPrefixes.Any(p => key.StartsWith(p, StringComparison.OrdinalIgnoreCase));
            // font.bold / font.italic Get keys map straight through; the
            // remaining top-level toggles (strike/underline/super/subscript)
            // are in CellStyleKeys.
            if (!isStyle) continue;
            var s = FormatValue(value);
            if (s.Length == 0) continue;
            props[key] = s;
        }
        return props;
    }

    private static bool SamePropDict(Dictionary<string, string> a, Dictionary<string, string> b)
    {
        if (a.Count != b.Count) return false;
        foreach (var (k, v) in a)
            if (!b.TryGetValue(k, out var bv) || !string.Equals(v, bv, StringComparison.Ordinal))
                return false;
        return true;
    }

    // ==================== Small helpers ====================

    // Row/col flag values arrive from Get as either typed bools or the raw
    // attribute string ("true"/"1").
    private static bool IsTruthyFormatValue(object? v) => v switch
    {
        bool b => b,
        string s => s.Equals("true", StringComparison.OrdinalIgnoreCase) || s == "1",
        _ => false,
    };

    private static string FormatValue(object? v) => v switch
    {
        bool b => b ? "true" : "false",
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        _ => Convert.ToString(v, CultureInfo.InvariantCulture) ?? "",
    };

    private static string LastPathSegment(string? path)
    {
        if (string.IsNullOrEmpty(path)) return "";
        var idx = path!.LastIndexOf('/');
        return idx >= 0 ? path[(idx + 1)..] : path;
    }

    private static (int C1, uint R1, int C2, uint R2)? ParseRangeRect(string range)
    {
        var parts = range.Split(':');
        if (!TryParseCellRef(parts[0].Trim(), out var c1, out var r1)) return null;
        if (parts.Length == 1) return (c1, r1, c1, r1);
        if (!TryParseCellRef(parts[1].Trim(), out var c2, out var r2)) return null;
        return (Math.Min(c1, c2), Math.Min(r1, r2), Math.Max(c1, c2), Math.Max(r1, r2));
    }

    private static bool TryParseCellRef(string cellRef, out int colIdx, out uint rowIdx)
    {
        colIdx = 0; rowIdx = 0;
        int i = 0;
        while (i < cellRef.Length && char.IsAsciiLetter(cellRef[i])) i++;
        if (i == 0 || i == cellRef.Length) return false;
        int col = 0;
        for (int k = 0; k < i; k++)
        {
            var c = char.ToUpperInvariant(cellRef[k]);
            if (c < 'A' || c > 'Z') return false;
            col = col * 26 + (c - 'A' + 1);
        }
        if (!uint.TryParse(cellRef[i..], NumberStyles.None, CultureInfo.InvariantCulture, out rowIdx)
            || rowIdx == 0) return false;
        colIdx = col;
        return true;
    }

    private static string ColumnName(int index)
    {
        var sb = new StringBuilder();
        while (index > 0)
        {
            index--;
            sb.Insert(0, (char)('A' + index % 26));
            index /= 26;
        }
        return sb.ToString();
    }

    private static string EscapeCsvField(string field)
    {
        if (field.Length == 0) return field;
        var needsQuote = field.Contains(',') || field.Contains('"')
            || field.Contains('\n') || field.Contains('\r');
        if (!needsQuote) return field;
        return "\"" + field.Replace("\"", "\"\"") + "\"";
    }
}
