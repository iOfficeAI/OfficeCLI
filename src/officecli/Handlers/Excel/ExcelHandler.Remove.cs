// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string? Remove(string path, Dictionary<string, string>? properties = null)
    {
        // Phase 4: trackChange.* is Word-only. Silently ignored here.
        Modified = true;
        // CONSISTENCY(container-remove-guard): reject removal of the
        // workbook root up front. Sheet-level removal has its own guard
        // (can't remove last sheet) further down and is a legitimate op;
        // /workbook is not.
        if (!string.IsNullOrEmpty(path)
            && path.TrimEnd('/').Equals("/workbook", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Cannot remove container element '{path}': it is a required structural element of the document.");

        // Batch Remove: a selector path (not starting with '/') → Query → Remove
        // each match, mirroring ExcelHandler.Set's selector branch. Row removals
        // are TRUE shift-deletes (rows below shift up — see "row[N] — true shift
        // delete"), so multiple matched rows MUST be removed in DESCENDING row
        // order: deleting /Sheet/row[2] first renumbers the old row[4] to row[3]
        // and the next delete would hit the wrong row. Non-row targets carry
        // index 0 and keep a stable relative order.
        if (!string.IsNullOrEmpty(path)
            && (!path.StartsWith("/") || Core.AttributeFilter.IsContentFilterPath(path)))
        {
            // Narrow via the shared engine (same as Set / query): pure-AND on the
            // legacy path, `or` selectors queried bracket-stripped then narrowed by
            // the boolean expression tree. The IsContentFilterPath arm routes a
            // `/`-scoped content filter (`/Sheet1/cell[value>5 or value<1]`) here
            // too, matching the Set dispatch — query, set and remove now agree on
            // every selector shape.
            var (targets, _) = Core.AttributeFilter.FilterSelector(path, Query, ResolveCellAttributeAlias);
            if (targets.Count == 0)
                // Empty selector result is not_found, not a crash — see Set.cs.
                throw new Core.CliException($"No elements matched selector: {path}") { Code = "not_found" };

            var ordered = targets.OrderByDescending(t => ExtractRowIndexForRemoval(t.Path)).ToList();
            string? lastWarning = null;
            foreach (var target in ordered)
            {
                var w = Remove(target.Path, properties);
                if (w != null) lastWarning = w;
            }
            var summary = $"{ordered.Count} element(s) removed by selector '{path}'";
            return lastWarning != null ? $"{summary}; {lastWarning}" : summary;
        }

        path = NormalizeExcelPath(path);
        path = ResolveSheetIndexInPath(path);
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        // Handle /namedrange[N] or /namedrange[Name] before sheet lookup
        var namedRangeRemoveMatch = Regex.Match(sheetName, @"^namedrange\[(.+?)\]$", RegexOptions.IgnoreCase);
        if (namedRangeRemoveMatch.Success)
        {
            var selector = namedRangeRemoveMatch.Groups[1].Value;
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames == null)
                throw new ArgumentException("No named ranges found in workbook");

            var allDefs = definedNames.Elements<DefinedName>().ToList();
            DefinedName? dn = null;

            if (int.TryParse(selector, out var dnIndex))
            {
                if (dnIndex < 1 || dnIndex > allDefs.Count)
                    throw new ArgumentException($"Named range index {dnIndex} out of range (1-{allDefs.Count})");
                dn = allDefs[dnIndex - 1];
            }
            else
            {
                dn = allDefs.FirstOrDefault(d =>
                    d.Name?.Value?.Equals(selector, StringComparison.OrdinalIgnoreCase) == true);
                if (dn == null)
                    throw new ArgumentException($"Named range '{selector}' not found");
            }

            dn.Remove();
            if (!definedNames.HasChildren) definedNames.Remove();
            workbook.Save();
            return null;
        }

        if (segments.Length == 1)
        {
            // Remove entire sheet
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var sheets = GetWorkbook().GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
            if (sheet == null)
                throw SheetNotFoundException(sheetName);

            var sheetCount = sheets!.Elements<Sheet>().Count();
            if (sheetCount <= 1)
                throw new InvalidOperationException($"Cannot remove the last sheet. A workbook must contain at least one sheet.");

            // CONSISTENCY(remove-sheet-chart-refs): a chart on another
            // sheet may carry <c:f>SheetName!$A$1:$B$2</c:f> references
            // pointing at the sheet about to disappear. The Open XML
            // SDK doesn't follow these into a dependency graph, so the
            // chart silently survives and Excel surfaces a confusing
            // "external links" warning when the file is reopened
            // (Excel reads the orphaned `SheetName!` prefix as a
            // pointer to a separate workbook). Refuse with a clear
            // message — named ranges referencing the sheet are
            // already cleaned up below as a passive cleanup, but a
            // chart series carries layout intent that the user almost
            // certainly wants to handle explicitly.
            var sheetIdForCheck = sheet.Id?.Value;
            var sheetWsPartForCheck = sheetIdForCheck != null
                ? workbookPart.GetPartById(sheetIdForCheck) as WorksheetPart
                : null;
            var refToken = sheetName + "!";
            var quotedRefToken = "'" + sheetName + "'!";
            foreach (var otherWsPart in workbookPart.WorksheetParts)
            {
                if (sheetWsPartForCheck != null && ReferenceEquals(otherWsPart, sheetWsPartForCheck)) continue;
                if (otherWsPart.DrawingsPart == null) continue;
                foreach (var dp in otherWsPart.DrawingsPart.ChartParts)
                {
                    var chartXml = dp.ChartSpace?.InnerXml;
                    if (chartXml == null) continue;
                    if (chartXml.Contains(refToken, StringComparison.OrdinalIgnoreCase)
                        || chartXml.Contains(quotedRefToken, StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException(
                            $"Cannot remove sheet '{sheetName}': it is referenced by a chart in this workbook. " +
                            $"Remove or repoint the chart first.");
                }
            }

            // CONSISTENCY(remove-sheet-refs): worksheet XML on other
            // sheets carries sheet-qualified formula text in three more
            // shapes that produce the same "external links" warning if
            // left dangling. Walk typed descendants per worksheet so we
            // don't false-positive on cell text or comments containing
            // the literal substring "Sheet1!".
            //   - sparkline data range  (<xne:f>SheetName!A1:A4</xne:f>)
            //   - data validation list  (<x:formula1>SheetName!...</x:formula1>)
            //   - conditional formatting (<x:formula>SheetName!...</x:formula>)
            // Cell formulas themselves (<x:f>) are intentionally not
            // guarded — Excel shows #REF! on open, which the existing
            // R9-1 cache invalidation already accommodates.
            foreach (var otherWsPart in workbookPart.WorksheetParts)
            {
                if (sheetWsPartForCheck != null && ReferenceEquals(otherWsPart, sheetWsPartForCheck)) continue;
                var wsRoot = otherWsPart.Worksheet;
                if (wsRoot == null) continue;

                bool MatchesRef(string? text) =>
                    text != null
                    && (text.Contains(refToken, StringComparison.OrdinalIgnoreCase)
                        || text.Contains(quotedRefToken, StringComparison.OrdinalIgnoreCase));

                foreach (var f in wsRoot.Descendants<DocumentFormat.OpenXml.Office.Excel.Formula>())
                    if (MatchesRef(f.Text))
                        throw new ArgumentException(
                            $"Cannot remove sheet '{sheetName}': it is referenced by a sparkline in this workbook. " +
                            $"Remove or repoint the sparkline first.");

                foreach (var f in wsRoot.Descendants<DocumentFormat.OpenXml.Spreadsheet.Formula1>())
                    if (MatchesRef(f.Text))
                        throw new ArgumentException(
                            $"Cannot remove sheet '{sheetName}': it is referenced by a data validation formula. " +
                            $"Remove or repoint the validation first.");

                foreach (var f in wsRoot.Descendants<DocumentFormat.OpenXml.Spreadsheet.Formula2>())
                    if (MatchesRef(f.Text))
                        throw new ArgumentException(
                            $"Cannot remove sheet '{sheetName}': it is referenced by a data validation formula. " +
                            $"Remove or repoint the validation first.");

                foreach (var f in wsRoot.Descendants<DocumentFormat.OpenXml.Spreadsheet.Formula>())
                    if (MatchesRef(f.Text))
                        throw new ArgumentException(
                            $"Cannot remove sheet '{sheetName}': it is referenced by a conditional formatting rule. " +
                            $"Remove or repoint the rule first.");

                // Internal hyperlinks: <x:hyperlink ref="A1"
                // location="SheetName!A1"/>. Same "external links"
                // class — Excel reads the orphan SheetName! as a
                // pointer to a separate workbook.
                foreach (var hl in wsRoot.Descendants<DocumentFormat.OpenXml.Spreadsheet.Hyperlink>())
                    if (MatchesRef(hl.Location?.Value))
                        throw new ArgumentException(
                            $"Cannot remove sheet '{sheetName}': it is referenced by a hyperlink in this workbook. " +
                            $"Remove or repoint the hyperlink first.");
            }

            // CONSISTENCY(remove-sheet-refs): pivotCacheDefinition parts live
            // at the workbook level; their <x:cacheSource><x:worksheetSource
            // sheet="SheetName" .../></x:cacheSource> binds the cache to a
            // source sheet. Removing that sheet leaves a dangling cache and
            // Excel surfaces the same "external links" / "found a problem"
            // dialog as the chart/sparkline/DV/hyperlink cases above.
            foreach (var cacheDefPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
            {
                var wsSource = cacheDefPart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                var srcSheet = wsSource?.Sheet?.Value;
                if (!string.IsNullOrEmpty(srcSheet)
                    && srcSheet.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException(
                        $"Cannot remove sheet '{sheetName}': it is referenced as the source of a pivot table in this workbook. " +
                        $"Remove or repoint the pivot table first.");
            }

            // CONSISTENCY(remove-sheet-refs): a pivot table hosted on the
            // sheet about to disappear may itself be named by a slicer cache
            // living on another sheet (SlicerCachePivotTables). Removing the
            // sheet orphans the pivot and leaves the slicer cache pointing at
            // a gone pivot — schema-valid but real Excel refuses (0x800A03EC).
            // Mirror the pivottable[N] guard: refuse and steer the user to
            // remove the slicer first. Note the mirror case (removing the
            // sheet that hosts the *slicer*, pivot elsewhere) is untouched and
            // remains a safe delete.
            {
                var relIdForSlicerCheck = sheet.Id?.Value;
                var wsForSlicerCheck = relIdForSlicerCheck != null
                    ? workbookPart.GetPartById(relIdForSlicerCheck) as WorksheetPart
                    : null;
                if (wsForSlicerCheck != null)
                    foreach (var pp in wsForSlicerCheck.PivotTableParts)
                        ThrowIfPivotReferencedBySlicer(
                            pp.PivotTableDefinition?.Name?.Value,
                            (pivotName, cacheName) =>
                                $"Cannot remove sheet '{sheetName}': it hosts pivot table '{pivotName}' " +
                                $"which is referenced by slicer cache '{cacheName}'. Remove the slicer first.");
            }

            // R10-2: capture pivot cache definitions referenced by this
            // sheet's pivot table parts BEFORE deleting the worksheet part,
            // so we can prune any caches that become orphaned by the
            // removal. Without this the workbook still carries pivotCaches
            // entries + cache parts whose owning pivot is gone, which
            // corrupts the file (Content_Types + workbook.xml.rels keep
            // references to unreachable parts). Mirrors the cleanup done
            // by the pivottable[N] branch below — both routes share the
            // same orphan prune helper.
            // localSheetId on <definedName> is a 0-based position into
            // <sheets>; capture the removed sheet's position before it is
            // detached so scoped names can be renumbered below.
            var removedSheetIndex = (uint)sheets.Elements<Sheet>()
                .TakeWhile(s => !ReferenceEquals(s, sheet)).Count();

            var relId = sheet.Id?.Value;
            var sheetWsPart = relId != null
                ? workbookPart.GetPartById(relId) as WorksheetPart
                : null;
            var cachePartsTouched = sheetWsPart != null
                ? sheetWsPart.PivotTableParts
                    .Select(pp => pp.PivotTableCacheDefinitionPart)
                    .Where(cp => cp != null)
                    .Cast<PivotTableCacheDefinitionPart>()
                    .Distinct()
                    .ToList()
                : new List<PivotTableCacheDefinitionPart>();

            // Evict the worksheet part from the row cache and dirty set BEFORE
            // DeletePart destroys it. FlushDirtyParts() calls GetSheet() on
            // every entry in _dirtyWorksheets; if the part is already destroyed
            // that call throws InvalidOperationException.
            if (sheetWsPart != null)
            {
                var removedSheetData = GetSheet(sheetWsPart).GetFirstChild<SheetData>();
                if (removedSheetData != null) InvalidateRowIndex(removedSheetData);
                _dirtyWorksheets.Remove(sheetWsPart);
            }

            sheet.Remove();
            if (relId != null)
                workbookPart.DeletePart(workbookPart.GetPartById(relId));

            // Prune orphan pivot caches now that the sheet (and its pivot
            // table parts) are gone. PrunePivotCacheIfOrphan walks every
            // remaining worksheet's pivot tables to confirm the cache is no
            // longer referenced, then drops the workbook-level pivotCache
            // entry and the cache part itself (which cascades to records,
            // _rels, and Content_Types).
            foreach (var cp in cachePartsTouched)
                PrunePivotCacheIfOrphan(workbookPart, cp);

            // CONSISTENCY(remove-sheet-refs): defined names that point into the
            // removed sheet are silently dropped (they would be orphaned).
            // BUT: if those defined names are referenced by formulas in *other*
            // sheets, dropping them silently leaves those formulas with #NAME?.
            // Mirror the DV / sparkline / pivot guards: throw if any other-sheet
            // formula uses one of the about-to-be-orphaned names.
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                var orphanNames = definedNames.Elements<DefinedName>()
                    .Where(dn => dn.Text?.Contains(sheetName + "!", StringComparison.OrdinalIgnoreCase) == true)
                    .Select(dn => dn.Name?.Value)
                    .Where(n => !string.IsNullOrEmpty(n))
                    .ToList();
                if (orphanNames.Count > 0)
                {
                    var refs = new List<string>();
                    foreach (var otherWsPart in workbookPart.WorksheetParts)
                    {
                        if (sheetWsPartForCheck != null && ReferenceEquals(otherWsPart, sheetWsPartForCheck)) continue;
                        var otherSheetName = workbook.Sheets!.Elements<Sheet>()
                            .FirstOrDefault(s => s.Id?.Value == workbookPart.GetIdOfPart(otherWsPart))?.Name?.Value ?? "?";
                        if (otherWsPart.Worksheet is null) continue;
                        foreach (var fcell in otherWsPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                        {
                            var f = fcell.CellFormula?.Text;
                            if (string.IsNullOrEmpty(f)) continue;
                            foreach (var n in orphanNames)
                            {
                                if (Regex.IsMatch(f, @"\b" + Regex.Escape(n!) + @"\b", RegexOptions.IgnoreCase))
                                {
                                    refs.Add($"{otherSheetName}!{fcell.CellReference?.Value ?? "?"} (uses '{n}')");
                                    break;
                                }
                            }
                        }
                    }
                    if (refs.Count > 0)
                        throw new ArgumentException(
                            $"Cannot remove sheet '{sheetName}': defined name(s) [{string.Join(", ", orphanNames)}] " +
                            $"are referenced by formulas in {string.Join(", ", refs)}. " +
                            $"Remove or repoint the formulas first.");
                }

                // No external usage — safe to drop the orphan names.
                var toRemove = definedNames.Elements<DefinedName>()
                    .Where(dn => dn.Text?.Contains(sheetName + "!", StringComparison.OrdinalIgnoreCase) == true)
                    .ToList();
                foreach (var dn in toRemove) dn.Remove();

                // Renumber sheet-scoped names: localSheetId is a 0-based
                // position into <sheets>, so removing a sheet shifts every
                // later sheet down by one. Names scoped to the removed sheet
                // itself lose their scope with it (the text-based drop above
                // only catches bodies that mention the removed sheet's name).
                // Excel refuses to open a workbook whose definedName carries
                // an out-of-range localSheetId (0x800A03EC).
                foreach (var dn in definedNames.Elements<DefinedName>().ToList())
                {
                    var lid = dn.LocalSheetId?.Value;
                    if (lid == null) continue;
                    if (lid == removedSheetIndex) dn.Remove();
                    else if (lid > removedSheetIndex) dn.LocalSheetId = lid.Value - 1;
                }
                if (!definedNames.HasChildren) definedNames.Remove();
            }

            // R9-1: invalidate stale cachedValue on formulas in other sheets
            // that referenced the removed sheet. Real Excel would recompute
            // to #REF! on open; our Get must not report the stale value.
            // Minimum viable: clear <x:v> so cachedValue drops out. We leave
            // the formula body alone — rewriting it to #REF! is what Excel
            // does on recalc and is hard to get right.
            InvalidateFormulaCacheReferencingSheet(workbookPart, sheetName);

            // Fix ActiveTab to prevent workbook corruption when deleting the last tab
            var remainingCount = sheets!.Elements<Sheet>().Count();
            var bookViews = workbook.GetFirstChild<BookViews>();
            if (bookViews != null)
            {
                foreach (var bv in bookViews.Elements<WorkbookView>())
                {
                    if (bv.ActiveTab?.Value >= (uint)remainingCount)
                        bv.ActiveTab = (uint)Math.Max(0, remainingCount - 1);
                }
            }

            workbook.Save();
            return null;
        }

        var cellRef = segments[1];
        var worksheet = FindWorksheet(sheetName)
            ?? throw SheetNotFoundException(sheetName);
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // row[N] — true shift delete
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = int.Parse(rowMatch.Groups[1].Value);
            sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx)
                ?.Remove();
            var affected = CollectFormulaCellsAffectedByRowDelete(worksheet, rowIdx);
            ShiftRowsUp(worksheet, rowIdx);
            DeleteCalcChainIfPresent();
            SaveWorksheet(worksheet);
            return FormatFormulaWarning(affected);
        }

        // col[X] — true shift delete
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colName = colMatch.Groups[1].Value.ToUpperInvariant();
            var deletedColIdx = ColumnNameToIndex(colName);
            var affected = CollectFormulaCellsAffectedByColDelete(worksheet, deletedColIdx);
            ShiftColumnsLeft(worksheet, colName);
            DeleteCalcChainIfPresent();
            SaveWorksheet(worksheet);
            return FormatFormulaWarning(affected);
        }

        // sparkline[N] — remove sparkline group
        var sparklineRemoveMatch = Regex.Match(cellRef, @"^sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (sparklineRemoveMatch.Success)
        {
            var spkIdx = int.Parse(sparklineRemoveMatch.Groups[1].Value);
            var spkGroup = GetSparklineGroup(worksheet, spkIdx)
                ?? throw new ArgumentException($"Sparkline[{spkIdx}] not found in sheet '{sheetName}'");
            var spkGroups = spkGroup.Parent!;
            spkGroup.Remove();
            // If no more sparkline groups, clean up empty extension
            if (!spkGroups.HasChildren)
            {
                var spkExt = spkGroups.Parent;
                spkGroups.Remove();
                if (spkExt != null && !spkExt.HasChildren)
                {
                    var extList = spkExt.Parent;
                    spkExt.Remove();
                    if (extList != null && !extList.HasChildren)
                        extList.Remove();
                }
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // rowbreak[N] / colbreak[N]
        var rbRemoveMatch = Regex.Match(cellRef, @"^rowbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (rbRemoveMatch.Success)
        {
            var rbIdx = int.Parse(rbRemoveMatch.Groups[1].Value);
            var rowBreaks = GetSheet(worksheet).GetFirstChild<RowBreaks>();
            var breaks = rowBreaks?.Elements<Break>().ToList() ?? new();
            if (rbIdx >= 1 && rbIdx <= breaks.Count)
            {
                breaks[rbIdx - 1].Remove();
                if (rowBreaks != null)
                {
                    rowBreaks.Count = (uint)rowBreaks.Elements<Break>().Count();
                    rowBreaks.ManualBreakCount = rowBreaks.Count;
                    if (rowBreaks.Count == 0) rowBreaks.Remove();
                }
            }
            SaveWorksheet(worksheet);
            return null;
        }
        var cbRemoveMatch = Regex.Match(cellRef, @"^colbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cbRemoveMatch.Success)
        {
            var cbIdx = int.Parse(cbRemoveMatch.Groups[1].Value);
            var colBreaks = GetSheet(worksheet).GetFirstChild<ColumnBreaks>();
            var breaks = colBreaks?.Elements<Break>().ToList() ?? new();
            if (cbIdx >= 1 && cbIdx <= breaks.Count)
            {
                breaks[cbIdx - 1].Remove();
                if (colBreaks != null)
                {
                    colBreaks.Count = (uint)colBreaks.Elements<Break>().Count();
                    colBreaks.ManualBreakCount = colBreaks.Count;
                    if (colBreaks.Count == 0) colBreaks.Remove();
                }
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // shape[N] — remove shape anchor from DrawingsPart
        var shapeRemoveMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (shapeRemoveMatch.Success)
        {
            var shpIdx = int.Parse(shapeRemoveMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/shapes");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/shapes");
            var shpAnchors = wsDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
                .Where(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().Any())
                .ToList();
            if (shpIdx < 1 || shpIdx > shpAnchors.Count)
                throw new ArgumentException($"Shape index {shpIdx} out of range (1..{shpAnchors.Count})");
            shpAnchors[shpIdx - 1].Remove();
            wsDrawing.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // picture[N] — remove picture anchor from DrawingsPart
        var picRemoveMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (picRemoveMatch.Success)
        {
            var picIdx = int.Parse(picRemoveMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/pictures");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/pictures");
            var picAnchors = EnumeratePictureAnchors(wsDrawing).ToList();
            if (picIdx < 1 || picIdx > picAnchors.Count)
                throw new ArgumentException($"Picture index {picIdx} out of range (1..{picAnchors.Count})");
            // Remove associated image part to avoid storage bloat
            var pic = picAnchors[picIdx - 1].Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().First();
            var blipFill = pic.BlipFill?.Blip?.Embed?.Value;
            picAnchors[picIdx - 1].Remove();
            if (blipFill != null)
            {
                try { drawingsPart.DeletePart(drawingsPart.GetPartById(blipFill)); } catch { }
            }
            wsDrawing.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // chart[N] — remove chart anchor from DrawingsPart
        var chartRemoveMatch = Regex.Match(cellRef, @"^chart\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (chartRemoveMatch.Success)
        {
            var chartIdx = int.Parse(chartRemoveMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart
                ?? throw new ArgumentException("Sheet has no drawings/charts");
            var wsDrawing = drawingsPart.WorksheetDrawing
                ?? throw new ArgumentException("Sheet has no drawings/charts");
            var chartAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
                .Where(a => a.Descendants<C.ChartReference>().Any())
                .ToList();
            if (chartIdx < 1 || chartIdx > chartAnchors.Count)
                throw new ArgumentException($"Chart index {chartIdx} out of range (1..{chartAnchors.Count})");
            var anchor = chartAnchors[chartIdx - 1];
            var chartRef = anchor.Descendants<C.ChartReference>().First();
            var relId = chartRef.Id?.Value;
            anchor.Remove();
            if (relId != null)
            {
                try { drawingsPart.DeletePart(drawingsPart.GetPartById(relId)); } catch { }
            }
            wsDrawing.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // table[N] — remove table (ListObject) from worksheet
        var tableRemoveMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableRemoveMatch.Success)
        {
            var tblIdx = int.Parse(tableRemoveMatch.Groups[1].Value);
            var tableParts = worksheet.TableDefinitionParts.ToList();
            if (tblIdx < 1 || tblIdx > tableParts.Count)
                throw new ArgumentException($"Table index {tblIdx} out of range (1..{tableParts.Count})");
            var tablePart = tableParts[tblIdx - 1];

            // CONSISTENCY(remove-refs): mirror sheet-remove DV / sparkline / pivot
            // guards. Removing a table referenced by structured-ref formulas
            // (Table1[Col], Table1[#All], or bare Table1) leaves stale formulas
            // that Excel surfaces as #REF!/#NAME?. Scan every sheet's cell
            // formulas; throw with the offending cell list.
            var tableName = tablePart.Table?.Name?.Value;
            if (!string.IsNullOrEmpty(tableName) && _doc.WorkbookPart != null)
            {
                var refs = new List<string>();
                foreach (var wsp in _doc.WorkbookPart.WorksheetParts)
                {
                    if (wsp.Worksheet is null) continue;
                    var wsName = _doc.WorkbookPart.Workbook?.Sheets?
                        .Elements<Sheet>()
                        .FirstOrDefault(s => s.Id?.Value == _doc.WorkbookPart.GetIdOfPart(wsp))?
                        .Name?.Value ?? "?";
                    foreach (var fcell in wsp.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                    {
                        var f = fcell.CellFormula?.Text;
                        if (string.IsNullOrEmpty(f)) continue;
                        // Match Table1[ ... ] (structured ref) or bare Table1 as a
                        // word boundary token. Case-insensitive per Excel norms.
                        var pattern = @"\b" + Regex.Escape(tableName) + @"(\[|\b)";
                        if (Regex.IsMatch(f, pattern, RegexOptions.IgnoreCase))
                            refs.Add($"{wsName}!{fcell.CellReference?.Value ?? "?"}");
                    }
                }
                if (refs.Count > 0)
                    throw new ArgumentException(
                        $"Cannot remove table '{tableName}': it is referenced by formulas in {string.Join(", ", refs)}. " +
                        $"Remove or repoint the formulas first.");
            }

            worksheet.DeletePart(tablePart);
            // Also remove the tablePart reference from the TableParts element
            var tblParts = worksheet.Worksheet?.GetFirstChild<TableParts>();
            if (tblParts != null)
            {
                var tblPartEntries = tblParts.Elements<TablePart>().ToList();
                if (tblIdx <= tblPartEntries.Count)
                    tblPartEntries[tblIdx - 1].Remove();
                tblParts.Count = (uint)tblParts.Elements<TablePart>().Count();
                if (tblParts.Count == 0)
                    tblParts.Remove();
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // comment[N] — remove comment from WorksheetCommentsPart
        var commentRemoveMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentRemoveMatch.Success)
        {
            var cmtIdx = int.Parse(commentRemoveMatch.Groups[1].Value);
            var commentsPart = worksheet.WorksheetCommentsPart;
            if (commentsPart?.Comments == null)
                throw new ArgumentException($"No comments found in sheet");
            var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
            var comments = cmtList?.Elements<Comment>().ToList() ?? new();
            if (cmtIdx < 1 || cmtIdx > comments.Count)
                throw new ArgumentException($"Comment index {cmtIdx} out of range (1..{comments.Count})");
            var removedCommentRef = comments[cmtIdx - 1].Reference?.Value;
            comments[cmtIdx - 1].Remove();
            if (cmtList != null && !cmtList.HasChildren)
            {
                worksheet.DeletePart(commentsPart);
                // Clean up VmlDrawingPart only if it contains no non-comment shapes (e.g. form controls)
                var vmlPart = worksheet.VmlDrawingParts.FirstOrDefault();
                if (vmlPart != null)
                {
                    bool hasNonCommentShapes = false;
                    try
                    {
                        using var stream = vmlPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
                        var vmlDoc = System.Xml.Linq.XDocument.Load(stream);
                        var vNs = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:vml";
                        var xNs = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:office:excel";
                        var shapes = vmlDoc.Descendants(vNs + "shape").ToList();
                        hasNonCommentShapes = shapes.Any(s =>
                        {
                            var clientData = s.Element(xNs + "ClientData");
                            return clientData == null ||
                                   clientData.Attribute("ObjectType")?.Value != "Note";
                        });
                    }
                    catch { }

                    if (!hasNonCommentShapes)
                    {
                        worksheet.DeletePart(vmlPart);
                        var legacyDrawing = GetSheet(worksheet).Elements<LegacyDrawing>().FirstOrDefault();
                        legacyDrawing?.Remove();
                    }
                    else
                    {
                        // Remove only comment shapes from VML, keep form controls
                        try
                        {
                            using var stream = vmlPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
                            var vmlDoc = System.Xml.Linq.XDocument.Load(stream);
                            var vNs2 = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:vml";
                            var xNs2 = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:office:excel";
                            var commentShapes = vmlDoc.Descendants(vNs2 + "shape")
                                .Where(s =>
                                {
                                    var cd = s.Element(xNs2 + "ClientData");
                                    return cd != null && cd.Attribute("ObjectType")?.Value == "Note";
                                }).ToList();
                            foreach (var cs in commentShapes) cs.Remove();
                            stream.SetLength(0);
                            vmlDoc.Save(stream);
                        }
                        catch { }
                    }
                }
            }
            else
            {
                commentsPart.Comments.Save();
                // Partial delete: remove the single orphaned VML Note shape for
                // the removed comment's cell. Without this the <v:shape> lingers
                // and Excel renders a ghost comment box (Bug family: partial
                // comment remove leaves orphan VML shape).
                if (!string.IsNullOrEmpty(removedCommentRef))
                    RemoveCommentVmlShapeByRef(worksheet, removedCommentRef);
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // dataValidation[N] (canonical) / validation[N] (legacy alias) —
        // remove data validation. R7-bt-6 CONSISTENCY(path-segment-naming).
        var validationRemoveMatch = Regex.Match(cellRef, @"^(?:dataValidation|validation)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationRemoveMatch.Success)
        {
            var dvIdx = int.Parse(validationRemoveMatch.Groups[1].Value);
            var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>();
            if (dvs == null)
                throw new ArgumentException("No data validations found in sheet");
            var dvList = dvs.Elements<DataValidation>().ToList();
            if (dvIdx < 1 || dvIdx > dvList.Count)
                throw new ArgumentException($"Validation index {dvIdx} out of range (1..{dvList.Count})");
            dvList[dvIdx - 1].Remove();
            if (!dvs.HasChildren)
                dvs.Remove();
            else
                dvs.Count = (uint)dvs.Elements<DataValidation>().Count();
            SaveWorksheet(worksheet);
            return null;
        }

        // cf[N] — remove conditional formatting
        var cfRemoveMatch = Regex.Match(cellRef, @"^cf\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cfRemoveMatch.Success)
        {
            var cfIdx = int.Parse(cfRemoveMatch.Groups[1].Value);
            var ws = GetSheet(worksheet);
            var cfElements = ws.Elements<ConditionalFormatting>().ToList();
            if (cfIdx < 1 || cfIdx > cfElements.Count)
                throw new ArgumentException($"Conditional formatting index {cfIdx} out of range (1..{cfElements.Count})");
            cfElements[cfIdx - 1].Remove();
            SaveWorksheet(worksheet);
            return null;
        }

        // pivottable[N] — remove pivot table (and its cache if no other pivot references it)
        var pivotRemoveMatch = Regex.Match(cellRef, @"^pivottable\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (pivotRemoveMatch.Success)
        {
            var ptIdx = int.Parse(pivotRemoveMatch.Groups[1].Value);
            var pivotParts = worksheet.PivotTableParts.ToList();
            if (ptIdx < 1 || ptIdx > pivotParts.Count)
                throw new ArgumentException($"PivotTable index {ptIdx} out of range (1..{pivotParts.Count})");
            var pivotPart = pivotParts[ptIdx - 1];

            // Referencing-slicer guard — a slicer cache names its pivot table
            // via SlicerCachePivotTables; deleting the pivot underneath it
            // leaves a dangling reference that passes schema validation but
            // real Excel refuses (0x800A03EC). Mirrors the sheet-remove
            // pivot-source protection.
            ThrowIfPivotReferencedBySlicer(
                pivotPart.PivotTableDefinition?.Name?.Value,
                (pivotName, cacheName) =>
                    $"Cannot remove pivottable '{pivotName}': it is referenced by slicer cache " +
                    $"'{cacheName}'. Remove the slicer first.");

            // Capture the cache-definition part (if any) so we can clean up
            // workbook-level PivotCache registration after removing the pivot.
            var cachePart = pivotPart.PivotTableCacheDefinitionPart;

            // Capture pivot location before deleting the part so we can erase
            // the rendered cell data from sheetData. Without this, add→remove
            // cycles leave orphaned rows in sheetData (duplicate row indices,
            // unbounded XML growth). CONSISTENCY(pivot-remove-cleanup)
            var pivotLocationRef = pivotPart.PivotTableDefinition
                ?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Location>()
                ?.Reference?.Value;

            // Remove the pivot table part itself.
            worksheet.DeletePart(pivotPart);

            // Erase the pivot's rendered cells from sheetData.
            if (!string.IsNullOrEmpty(pivotLocationRef))
            {
                var pivotSd = GetSheet(worksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                if (pivotSd != null)
                    OfficeCli.Core.PivotTableHelper.ClearPivotRangeCells(pivotSd, pivotLocationRef);
            }

            // If no other pivot table references this cache, drop the cache
            // definition (and its records) plus the workbook-level PivotCache
            // registration. Otherwise leave it alone — shared caches are valid.
            // Shared with the sheet-remove path above via PrunePivotCacheIfOrphan.
            if (cachePart != null)
                PrunePivotCacheIfOrphan(_doc.WorkbookPart!, cachePart);

            SaveWorksheet(worksheet);
            return null;
        }

        // slicer[N] — remove pivot-backed slicer and all six cross-referenced
        // parts/registrations created by AddSlicer (SlicersPart entry,
        // SlicerCachePart, workbook + worksheet extLst registrations, the
        // Slicer_ defined-name sentinel, and the drawing anchor). Mirrors the
        // pivottable[N] multi-part cleanup discipline above.
        var slicerRemoveMatch = Regex.Match(cellRef, @"^slicer\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (slicerRemoveMatch.Success)
        {
            var slIdx = int.Parse(slicerRemoveMatch.Groups[1].Value);
            var slicersPart = worksheet.GetPartsOfType<SlicersPart>().FirstOrDefault();
            var slicersContainer = slicersPart?.Slicers;
            var slicerList = slicersContainer?.Elements<X14.Slicer>().ToList()
                ?? new List<X14.Slicer>();
            if (slIdx < 1 || slIdx > slicerList.Count)
                throw new ArgumentException($"Slicer index {slIdx} out of range (1..{slicerList.Count})");
            var slicerElement = slicerList[slIdx - 1];
            var slicerDisplayName = slicerElement.Name?.Value;
            var slicerCacheName = slicerElement.Cache?.Value;

            var slicerWbPart = _doc.WorkbookPart!;

            // 1. Remove the Slicer element; delete the SlicersPart and its
            //    worksheet extLst registration if it was the last slicer.
            slicerElement.Remove();
            slicersContainer!.Save(slicersPart!);
            if (!slicersContainer.Elements<X14.Slicer>().Any())
            {
                var slicersRelId = worksheet.GetIdOfPart(slicersPart!);
                worksheet.DeletePart(slicersPart!);
                RemoveSlicerListFromWorksheet(worksheet, slicersRelId);
            }

            // 2. Remove the backing SlicerCachePart + workbook extLst entry.
            if (!string.IsNullOrEmpty(slicerCacheName))
            {
                foreach (var scp in slicerWbPart.GetPartsOfType<SlicerCachePart>().ToList())
                {
                    if (scp.SlicerCacheDefinition?.Name?.Value != slicerCacheName) continue;
                    var cacheRelId = slicerWbPart.GetIdOfPart(scp);
                    slicerWbPart.DeletePart(scp);
                    RemoveSlicerCacheFromWorkbook(slicerWbPart, cacheRelId);
                    break;
                }
            }

            // 3. Remove the Slicer_ defined-name sentinel (reverse of
            //    RegisterSlicerDefinedName).
            if (!string.IsNullOrEmpty(slicerDisplayName))
            {
                var slicerDefinedNames = slicerWbPart.Workbook!.GetFirstChild<DefinedNames>();
                if (slicerDefinedNames != null)
                {
                    slicerDefinedNames.Elements<DefinedName>()
                        .FirstOrDefault(d => string.Equals(d.Name?.Value, slicerDisplayName, StringComparison.Ordinal))
                        ?.Remove();
                    if (!slicerDefinedNames.HasChildren) slicerDefinedNames.Remove();
                }
            }

            // 4. Remove the drawing anchor bound to this slicer cache name.
            if (!string.IsNullOrEmpty(slicerCacheName))
                RemoveSlicerDrawingAnchor(worksheet, slicerCacheName);

            SaveWorksheet(worksheet);
            slicerWbPart.Workbook!.Save();
            return null;
        }

        // ole[N] — remove embedded OLE object (cleanup embedded payload +
        // icon image part). Same part-cleanup discipline as picture/chart
        // removal to avoid orphaned binaries bloating the package.
        var oleRemoveMatch = Regex.Match(cellRef, @"^(?:ole|object|embed)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (oleRemoveMatch.Success)
        {
            var oleIdx = int.Parse(oleRemoveMatch.Groups[1].Value);
            var ws = GetSheet(worksheet);
            var oleElements = ws.Descendants<OleObject>().ToList();
            if (oleIdx < 1 || oleIdx > oleElements.Count)
                throw new ArgumentException($"OLE object index {oleIdx} out of range (1..{oleElements.Count})");
            var oleToRemove = oleElements[oleIdx - 1];
            // Capture the shapeId before removal so we can prune the matching
            // legacy VML shape (see below).
            var oleShapeId = oleToRemove.ShapeId?.Value;
            // Delete backing embedded payload + icon image part by rel id.
            if (oleToRemove.Id?.Value is string oleRelId && !string.IsNullOrEmpty(oleRelId))
            {
                try { worksheet.DeletePart(oleRelId); } catch { }
            }
            var objectPr = oleToRemove.GetFirstChild<EmbeddedObjectProperties>();
            if (objectPr?.Id?.Value is string oleIconRelId && !string.IsNullOrEmpty(oleIconRelId))
            {
                try { worksheet.DeletePart(oleIconRelId); } catch { }
            }
            // Remove the OleObject element itself; if its parent OleObjects
            // becomes empty, remove that too so the worksheet XML stays clean.
            var oleParent = oleToRemove.Parent;
            oleToRemove.Remove();
            if (oleParent is OleObjects oleColl && !oleColl.HasChildren)
                oleColl.Remove();

            // Prune the companion legacy VML shape. Without this, add/remove
            // cycles leave ghost <v:shape> elements accumulating in the VML
            // part (and a dangling <legacyDrawing> when the VML empties out),
            // mirroring the comment-remove cleanup discipline above.
            var oleVmlPart = worksheet.VmlDrawingParts.FirstOrDefault();
            if (oleVmlPart != null && oleShapeId.HasValue)
            {
                bool anyShapesLeft = true;
                try
                {
                    System.Xml.Linq.XDocument vmlDoc;
                    using (var stream = oleVmlPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read))
                        vmlDoc = System.Xml.Linq.XDocument.Load(stream);
                    var vNs = (System.Xml.Linq.XNamespace)"urn:schemas-microsoft-com:vml";
                    var target = vmlDoc.Descendants(vNs + "shape")
                        .FirstOrDefault(s => (string?)s.Attribute("id") == $"_x0000_s{oleShapeId.Value}");
                    target?.Remove();
                    anyShapesLeft = vmlDoc.Descendants(vNs + "shape").Any();
                    if (anyShapesLeft)
                    {
                        using var wstream = oleVmlPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write);
                        vmlDoc.Save(wstream);
                    }
                }
                catch { anyShapesLeft = true; }

                if (!anyShapesLeft)
                {
                    worksheet.DeletePart(oleVmlPart);
                    GetSheet(worksheet).Elements<LegacyDrawing>().FirstOrDefault()?.Remove();
                }
            }
            SaveWorksheet(worksheet);
            return null;
        }

        // autofilter — remove AutoFilter from worksheet
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            var ws = GetSheet(worksheet);
            var autoFilter = ws.GetFirstChild<AutoFilter>();
            if (autoFilter != null)
            {
                autoFilter.Remove();
                SaveWorksheet(worksheet);
            }
            return null;
        }

        // run[N] — remove individual run from rich text cell
        var runRemoveMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runRemoveMatch.Success)
        {
            var runCellRef = runRemoveMatch.Groups[1].Value.ToUpperInvariant();
            var runIdx = int.Parse(runRemoveMatch.Groups[2].Value);

            var runCell = FindCell(sheetData, runCellRef)
                ?? throw new ArgumentException($"Cell {runCellRef} not found");

            if (runCell.DataType?.Value != CellValues.SharedString ||
                !int.TryParse(runCell.CellValue?.Text, out var sstIdx))
                throw new ArgumentException($"Cell {runCellRef} is not a rich text cell");

            var sstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi = sstPart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx);
            if (ssi == null) throw new ArgumentException($"SharedString entry {sstIdx} not found");

            var runs = ssi.Elements<Run>().ToList();
            if (runIdx < 1 || runIdx > runs.Count)
                throw new ArgumentException($"Run index {runIdx} out of range (1-{runs.Count})");

            runs[PathIndex.ToArrayIndex(runIdx)].Remove();

            // Convert back to plain text if appropriate
            var remainingRuns = ssi.Elements<Run>().ToList();
            if (remainingRuns.Count == 0)
            {
                // All runs removed — set empty plain text to avoid orphaned SSI
                ssi.RemoveAllChildren<Text>();
                ssi.AppendChild(new Text("") { Space = SpaceProcessingModeValues.Preserve });
            }
            else if (remainingRuns.Count == 1)
            {
                var lastRun = remainingRuns[0];
                var rProps = lastRun.RunProperties;
                bool hasFormatting = rProps != null && rProps.HasChildren;
                if (!hasFormatting)
                {
                    var plainText = lastRun.GetFirstChild<Text>()?.Text ?? "";
                    lastRun.Remove();
                    ssi.RemoveAllChildren<Text>();
                    ssi.AppendChild(new Text(plainText) { Space = SpaceProcessingModeValues.Preserve });
                }
            }

            sstPart!.SharedStringTable!.Save();
            SaveWorksheet(worksheet);
            return null;
        }

        // Element-looking paths that reach the cell fallthrough (e.g.
        // chart[1]/series[2] — series has no remove operation) used to die
        // with a nonsensical "Cell chart[1]/series[2] not found". Name the
        // real limitation instead.
        if (cellRef.Contains("/series[", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                "remove is not supported for chart series (operations: add/set/get). " +
                "Rebuild the chart without the series, or repoint its values via set.");

        // Single cell
        var cell = FindCell(sheetData, cellRef)
            ?? throw new ArgumentException($"Cell {cellRef} not found");
        cell.Remove();
        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return null;
    }

    // Referencing-slicer guard — a slicer cache names its pivot table via
    // SlicerCachePivotTables; deleting that pivot (directly, or by removing
    // the sheet that hosts it) leaves a dangling reference that passes schema
    // validation but real Excel refuses to open (0x800A03EC). Shared by the
    // pivottable[N] branch and the whole-sheet removal path.
    private void ThrowIfPivotReferencedBySlicer(string? pivotName, Func<string, string, string> message)
    {
        if (string.IsNullOrEmpty(pivotName) || _doc.WorkbookPart == null) return;
        foreach (var scPart in _doc.WorkbookPart.GetPartsOfType<SlicerCachePart>())
        {
            var refsPivot = scPart.SlicerCacheDefinition?
                .GetFirstChild<X14.SlicerCachePivotTables>()?
                .Elements<X14.SlicerCachePivotTable>()
                .Any(pt => string.Equals(pt.Name?.Value, pivotName, StringComparison.OrdinalIgnoreCase)) == true;
            if (refsPivot)
                throw new ArgumentException(
                    message(pivotName, scPart.SlicerCacheDefinition?.Name?.Value ?? "?"));
        }
    }

    // Trailing /row[N] index of a path, or 0 when the path is not a row. Used to
    // order selector-Remove deletions descending so a row shift-delete never
    // invalidates the indices of not-yet-deleted matches.
    private static int ExtractRowIndexForRemoval(string? path)
    {
        var m = Regex.Match(path ?? "", @"/row\[(\d+)\]$", RegexOptions.IgnoreCase);
        return m.Success ? int.Parse(m.Groups[1].Value) : 0;
    }

    /// <summary>
    /// Remove a single cell at the given path and shift the remaining cells
    /// in the same row (shift=left) or same column (shift=up) by one position
    /// to fill the gap. Mirrors Excel UI's "Delete Cells > Shift cells left /
    /// Shift cells up". For full row/column delete with all metadata
    /// adjustments use <c>Remove("/Sheet1/row[N]")</c> or
    /// <c>Remove("/Sheet1/col[X]")</c> instead — those handle merged cells,
    /// CF/DV/hyperlink/table refs, and formula refs across the entire sheet.
    ///
    /// <para>Limitation: only cell references inside the affected row (for
    /// shift=left) or column (for shift=up) are rewritten. Formula text in
    /// other rows/columns that references cells in the affected row/col is
    /// NOT adjusted — Excel will recalculate against the new values on open
    /// (fullCalcOnLoad), but a formula like A1=<c>=C5</c> after deleting B5
    /// with shift=left will still read literal C5, not the new B5. Mergeed
    /// cells and other range-based metadata that span the affected row/col
    /// are also not adjusted. If precise behavior matters, prefer the
    /// row/col-level remove.</para>
    /// </summary>
    public string? RemoveCellWithShift(string path, string shift)
    {
        Modified = true;
        if (string.IsNullOrEmpty(shift))
            throw new ArgumentException("--shift requires a value: left or up");
        var direction = shift.ToLowerInvariant();
        if (direction is not ("left" or "up"))
            throw new ArgumentException(
                $"--shift={shift} not valid for remove. Use 'left' or 'up'.");

        path = NormalizeExcelPath(path);
        path = ResolveSheetIndexInPath(path);
        var segments = path.TrimStart('/').Split('/', 2);
        if (segments.Length < 2)
            throw new ArgumentException(
                "--shift requires a cell path like /Sheet1/B5");
        var sheetName = segments[0];
        var cellRef = segments[1].ToUpperInvariant();
        if (!System.Text.RegularExpressions.Regex.IsMatch(cellRef, @"^[A-Z]+\d+$"))
            throw new ArgumentException(
                $"--shift requires a single-cell path; got {cellRef}");

        var worksheet = FindWorksheet(sheetName)
            ?? throw SheetNotFoundException(sheetName);
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        var (col, rowIdx) = ParseCellReference(cellRef);
        var colIdx = ColumnNameToIndex(col);

        if (direction == "left")
            ShiftCellsLeftInRow(sheetData, (uint)rowIdx, colIdx);
        else
            ShiftCellsUpInColumn(sheetData, col, rowIdx);

        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return null;
    }

    /// <summary>
    /// Remove the cell at (rowIdx, fromColIdx) and shift every cell with
    /// col &gt; fromColIdx in the same row left by one.
    /// </summary>
    private void ShiftCellsLeftInRow(SheetData sheetData, uint rowIdx, int fromColIdx)
    {
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null) return;

        foreach (var cell in row.Elements<Cell>().ToList())
        {
            if (cell.CellReference?.Value == null) continue;
            var (cCol, cRow) = ParseCellReference(cell.CellReference.Value);
            var cColIdx = ColumnNameToIndex(cCol);
            if (cColIdx == fromColIdx)
                cell.Remove();
            else if (cColIdx > fromColIdx)
                cell.CellReference = $"{IndexToColumnName(cColIdx - 1)}{cRow}";
        }
    }

    /// <summary>
    /// Shift every cell with col &gt;= fromColIdx in the given row right by
    /// one, opening a gap at (rowIdx, fromColIdx). Used by add cell with
    /// --prop shift=right.
    /// </summary>
    internal void ShiftCellsRightInRow(SheetData sheetData, uint rowIdx, int fromColIdx)
    {
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null) return;

        // Process in reverse-col order so we don't overwrite a not-yet-shifted ref.
        var cells = row.Elements<Cell>()
            .Where(c => c.CellReference?.Value != null)
            .Select(c => new { Cell = c, ColIdx = ColumnNameToIndex(ParseCellReference(c.CellReference!.Value!).Column) })
            .Where(t => t.ColIdx >= fromColIdx)
            .OrderByDescending(t => t.ColIdx)
            .ToList();
        foreach (var t in cells)
        {
            var pr = ParseCellReference(t.Cell.CellReference!.Value!);
            t.Cell.CellReference = $"{IndexToColumnName(t.ColIdx + 1)}{pr.Row}";
        }
    }

    /// <summary>
    /// Shift every cell with row &gt;= fromRow in the given column down by
    /// one, opening a gap at (fromRow, col). Used by add cell with
    /// --prop shift=down.
    /// </summary>
    internal void ShiftCellsDownInColumn(SheetData sheetData, string col, int fromRow)
    {
        // Reverse-row order to avoid collisions during rewrite.
        foreach (var row in sheetData.Elements<Row>().OrderByDescending(r => r.RowIndex?.Value ?? 0))
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < fromRow) continue;

            var cell = row.Elements<Cell>().FirstOrDefault(c =>
            {
                if (c.CellReference?.Value == null) return false;
                var (cCol, _) = ParseCellReference(c.CellReference.Value);
                return cCol.Equals(col, StringComparison.OrdinalIgnoreCase);
            });
            if (cell != null)
                cell.CellReference = $"{col}{rowIdx + 1}";
        }
    }

    /// <summary>
    /// Remove the cell at (fromRow, col) and shift every cell with row &gt;
    /// fromRow in the same column up by one.
    /// </summary>
    private void ShiftCellsUpInColumn(SheetData sheetData, string col, int fromRow)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < fromRow) continue;

            var cell = row.Elements<Cell>().FirstOrDefault(c =>
            {
                if (c.CellReference?.Value == null) return false;
                var (cCol, _) = ParseCellReference(c.CellReference.Value);
                return cCol.Equals(col, StringComparison.OrdinalIgnoreCase);
            });
            if (cell == null) continue;

            if (rowIdx == fromRow)
                cell.Remove();
            else
                cell.CellReference = $"{col}{rowIdx - 1}";
        }
    }

    // ==================== Row/Column insert shift ====================

    /// <summary>
    /// Shift all rows >= insertRow down by 1 to make room for a new row insert.
    /// Mirrors ShiftRowsUp but in the opposite direction.
    /// </summary>
    internal void ShiftRowsDown(WorksheetPart worksheet, int insertRow)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name ?? "";

        // 1. SheetData cellRef rewrite (axis-direction-specific reverse iter,
        //    stays in caller — walker doesn't handle row renumber).
        if (sheetData != null)
        {
            InvalidateRowIndex(sheetData);
            foreach (var row in sheetData.Elements<Row>().OrderByDescending(r => r.RowIndex?.Value ?? 0).ToList())
            {
                var rowIdx = (int)(row.RowIndex?.Value ?? 0);
                if (rowIdx < insertRow) continue;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value != null)
                    {
                        var (col, _) = ParseCellReference(cell.CellReference.Value);
                        cell.CellReference = $"{col}{rowIdx + 1}";
                    }
                }
                row.RowIndex = (uint)(rowIdx + 1);
            }
        }

        // 2. All sheet-level range-bearing structures + formulas + namedRanges.
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => ShiftRowInRefDown(r, insertRow),
            formulaTextMapper: f => Core.FormulaRefShifter.Shift(
                f, sheetName, sheetName, Core.FormulaShiftDirection.RowsDown, insertRow),
            rowMarkerShift: m => m >= insertRow - 1 ? m + 1 : m,
            crossSheetFormulaMapper: (other, f) => Core.FormulaRefShifter.Shift(
                f, other, sheetName, Core.FormulaShiftDirection.RowsDown, insertRow));
    }

    /// <summary>
    /// Shift all columns >= insertColIdx right by 1 to make room for a new column insert.
    /// </summary>
    internal void ShiftColumnsRight(WorksheetPart worksheet, int insertColIdx)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name ?? "";

        // 1. SheetData cellRef rewrite (col-shift, no reverse iter needed
        //    because we go by colIdx not row order).
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    if (cell.CellReference?.Value == null) continue;
                    var (col, rowIdx) = ParseCellReference(cell.CellReference.Value);
                    var colIdx = ColumnNameToIndex(col);
                    if (colIdx >= insertColIdx)
                        cell.CellReference = $"{IndexToColumnName(colIdx + 1)}{rowIdx}";
                }
            }
        }

        // 2. <Columns> width/style (col-only, op-asymmetric — kept out of walker).
        var columns = ws.GetFirstChild<Columns>();
        if (columns != null)
        {
            foreach (var col in columns.Elements<Column>().OrderByDescending(c => c.Min?.Value ?? 0).ToList())
            {
                var min = (int)(col.Min?.Value ?? 0);
                var max = (int)(col.Max?.Value ?? 0);
                if (min >= insertColIdx) { col.Min = (uint)(min + 1); col.Max = (uint)(max + 1); }
                else if (max >= insertColIdx) col.Max = (uint)(max + 1);
            }
        }

        // 3. All sheet-level range-bearing structures + formulas + namedRanges.
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => ShiftColInRefRight(r, insertColIdx),
            formulaTextMapper: f => Core.FormulaRefShifter.Shift(
                f, sheetName, sheetName, Core.FormulaShiftDirection.ColumnsRight, insertColIdx),
            colMarkerShift: m => m >= insertColIdx - 1 ? m + 1 : m,
            crossSheetFormulaMapper: (other, f) => Core.FormulaRefShifter.Shift(
                f, other, sheetName, Core.FormulaShiftDirection.ColumnsRight, insertColIdx));

        // A column inserted inside a table's span widened its ref above; sync
        // the tableColumns list so count matches the ref width (else 0x800A03EC).
        SyncTableColumnsAfterColInsert(worksheet, insertColIdx);
    }

    private static string? ShiftRowInRefDown(string? refStr, int insertRow)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                shifted.Add(row >= insertRow ? $"{col}{row + 1}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    // RewriteFormulaRefsInSheet was removed — its responsibility (rewriting
    // CellFormula.Text and the shared/array formula `ref` attribute) is now
    // section 7 of ApplySheetRangeMutations in ExcelHandler.SheetShift.cs.

    private static string? ShiftColInRefRight(string? refStr, int insertColIdx)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');
        var shifted = new List<string>(parts.Length);
        foreach (var part in parts)
        {
            try
            {
                var (col, row) = ParseCellReference(part);
                var colIdx = ColumnNameToIndex(col);
                shifted.Add(colIdx >= insertColIdx ? $"{IndexToColumnName(colIdx + 1)}{row}" : part);
            }
            catch { shifted.Add(part); }
        }
        return string.Join(":", shifted);
    }

    // ShiftNamedRangeRowsDown / ShiftNamedRangeColsRight removed — defined
    // names are now rewritten by section 8 of ApplySheetRangeMutations using
    // the proper FormulaRefShifter (which handles quoted sheet names, string
    // literals, and structured refs correctly, unlike the old regex helpers).

    // ==================== Row shift ====================

    private void ShiftRowsUp(WorksheetPart worksheet, int deletedRow)
    {
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name ?? "";

        // 1. SheetData cellRef rewrite (delete direction).
        if (sheetData != null)
        {
            InvalidateRowIndex(sheetData);
            foreach (var row in sheetData.Elements<Row>().ToList())
            {
                var rowIdx = (int)(row.RowIndex?.Value ?? 0);
                if (rowIdx <= deletedRow) continue;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value != null)
                    {
                        var (col, _) = ParseCellReference(cell.CellReference.Value);
                        cell.CellReference = $"{col}{rowIdx - 1}";
                    }
                }
                row.RowIndex = (uint)(rowIdx - 1);
            }
        }

        // 2. All sheet-level range-bearing structures + formulas + namedRanges.
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => ShiftRowInRef(r, deletedRow),
            formulaTextMapper: f => Core.FormulaRefShifter.Shift(
                f, sheetName, sheetName, Core.FormulaShiftDirection.RowsUp, deletedRow),
            rowMarkerShift: m => m > deletedRow - 1 ? m - 1 : m,
            crossSheetFormulaMapper: (other, f) => Core.FormulaRefShifter.Shift(
                f, other, sheetName, Core.FormulaShiftDirection.RowsUp, deletedRow));
    }

    // ==================== Column shift ====================

    private void ShiftColumnsLeft(WorksheetPart worksheet, string deletedColName)
    {
        var ws = GetSheet(worksheet);
        var deletedColIdx = ColumnNameToIndex(deletedColName);
        var sheetData = ws.GetFirstChild<SheetData>();
        var sheetName = GetWorksheets().FirstOrDefault(w => w.Part == worksheet).Name ?? "";

        // 1. SheetData cellRef rewrite: remove cells in deleted col, shift others left.
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    if (cell.CellReference?.Value == null) continue;
                    var (col, rowIdx) = ParseCellReference(cell.CellReference.Value);
                    var colIdx = ColumnNameToIndex(col);
                    if (colIdx == deletedColIdx) cell.Remove();
                    else if (colIdx > deletedColIdx)
                        cell.CellReference = $"{IndexToColumnName(colIdx - 1)}{rowIdx}";
                }
            }
        }

        // 2. <Columns> width/style (col-only, op-asymmetric — kept out of walker).
        var columns = ws.GetFirstChild<Columns>();
        if (columns != null)
        {
            foreach (var col in columns.Elements<Column>().ToList())
            {
                var min = (int)(col.Min?.Value ?? 0);
                var max = (int)(col.Max?.Value ?? 0);
                if (min == deletedColIdx && max == deletedColIdx) col.Remove();
                else if (min > deletedColIdx) { col.Min = (uint)(min - 1); col.Max = (uint)(max - 1); }
                else if (max >= deletedColIdx) col.Max = (uint)(max - 1);
            }
            if (!columns.HasChildren) columns.Remove();
        }

        // 2b. table (ListObject) column sync. Deleting a worksheet column that
        // falls inside a table's range shrinks the table ref (handled by the
        // walker below) but ALSO drops one table column — the <tableColumns
        // count=".."> and its <tableColumn> children must follow, or Excel
        // refuses to open (0x800A03EC) even though schema validation passes.
        // Mirrors Excel: deleting a sheet column narrows the table; deleting
        // the table's only column removes the table entirely. Done here (not in
        // the shared walker) because it is column-axis-specific and needs the
        // pre-shift ref to locate the column position.
        SyncTableColumnsAfterColDelete(worksheet, deletedColIdx);

        // 3. All sheet-level range-bearing structures + formulas + namedRanges.
        ApplySheetRangeMutations(
            worksheet, sheetName,
            refMapper: r => ShiftColInRef(r, deletedColIdx),
            formulaTextMapper: f => Core.FormulaRefShifter.Shift(
                f, sheetName, sheetName, Core.FormulaShiftDirection.ColumnsLeft, deletedColIdx),
            colMarkerShift: m => m > deletedColIdx - 1 ? m - 1 : m,
            crossSheetFormulaMapper: (other, f) => Core.FormulaRefShifter.Shift(
                f, other, sheetName, Core.FormulaShiftDirection.ColumnsLeft, deletedColIdx));
    }

    /// <summary>
    /// After a worksheet column delete, keep every table (ListObject) on the
    /// sheet structurally consistent: if the deleted column falls inside a
    /// table's range, remove the corresponding &lt;tableColumn&gt; child and
    /// decrement the count. If it was the table's only column, remove the whole
    /// table (part + TableParts entry), matching Excel's "delete the last
    /// column, the table disappears" behavior.
    /// </summary>
    private void SyncTableColumnsAfterColDelete(WorksheetPart worksheet, int deletedColIdx)
    {
        var tableParts = worksheet.TableDefinitionParts.ToList();
        for (int i = 0; i < tableParts.Count; i++)
        {
            var tablePart = tableParts[i];
            var tbl = tablePart.Table;
            var refStr = tbl?.Reference?.Value;
            if (tbl == null || string.IsNullOrEmpty(refStr)) continue;

            var rangeParts = refStr.Split(':');
            int startColIdx, endColIdx;
            try
            {
                startColIdx = ColumnNameToIndex(ParseCellReference(rangeParts[0]).Column);
                endColIdx = rangeParts.Length > 1
                    ? ColumnNameToIndex(ParseCellReference(rangeParts[1]).Column)
                    : startColIdx;
            }
            catch { continue; }

            // Deleted column outside the table span → nothing to sync (the
            // walker still shifts the ref if the table sits to the right).
            if (deletedColIdx < startColIdx || deletedColIdx > endColIdx) continue;

            // Last remaining column removed → the table disappears entirely.
            if (startColIdx == endColIdx)
            {
                var tblIndex = i + 1; // 1-based position among TableParts
                worksheet.DeletePart(tablePart);
                var tblParts = worksheet.Worksheet?.GetFirstChild<TableParts>();
                if (tblParts != null)
                {
                    var entries = tblParts.Elements<TablePart>().ToList();
                    if (tblIndex <= entries.Count) entries[tblIndex - 1].Remove();
                    tblParts.Count = (uint)tblParts.Elements<TablePart>().Count();
                    if (tblParts.Count == 0) tblParts.Remove();
                }
                continue;
            }

            // Drop the table column at the deleted position (0-based within the
            // table). The walker shrinks tbl.Reference; here we only sync the
            // column list + count.
            var tableColumns = tbl.TableColumns;
            if (tableColumns == null) continue;
            var cols = tableColumns.Elements<TableColumn>().ToList();
            var pos = deletedColIdx - startColIdx;
            if (pos >= 0 && pos < cols.Count)
            {
                cols[pos].Remove();
                tableColumns.Count = (uint)tableColumns.Elements<TableColumn>().Count();
                // Renumber ids and (for header-less tables) rename Column1..N —
                // a gap or out-of-order auto name makes Excel refuse (0x800A03EC).
                NormalizeTableColumns(tbl, tableColumns);
                tbl.Save();
            }
        }
    }

    /// <summary>Reassign tableColumn @id sequentially 1..N. Excel refuses a
    /// table whose column ids have gaps or don't start at 1.</summary>
    private static void RenumberTableColumnIds(TableColumns tableColumns)
    {
        uint id = 1;
        foreach (var tc in tableColumns.Elements<TableColumn>())
            tc.Id = id++;
    }

    /// <summary>
    /// After a column insert/delete resync, fix up the tableColumn ids (always)
    /// and, for a HEADER-LESS table (headerRowCount=0, auto-named columns),
    /// rename them Column1..N in order. A header-less table with out-of-order
    /// or gapped auto names (e.g. Column1, Column3) makes Excel refuse the file
    /// (0x800A03EC). Header tables are left alone — their column names must
    /// track the header-row cells (handled elsewhere).
    /// </summary>
    private static void NormalizeTableColumns(Table tbl, TableColumns tableColumns)
    {
        RenumberTableColumnIds(tableColumns);
        bool headerLess = tbl.HeaderRowCount != null && tbl.HeaderRowCount.Value == 0;
        if (!headerLess) return;
        int n = 1;
        foreach (var tc in tableColumns.Elements<TableColumn>())
            tc.Name = $"Column{n++}";
    }

    /// <summary>
    /// Mirror of SyncTableColumnsAfterColDelete for column INSERTION. When a
    /// column is inserted inside a table's span, ShiftColumnsRight widens the
    /// table ref but left tableColumns unchanged — count no longer matched the
    /// ref width, which real Excel refuses (0x800A03EC). Insert a matching
    /// tableColumn at the right position and renumber ids.
    /// </summary>
    internal void SyncTableColumnsAfterColInsert(WorksheetPart worksheet, int insertColIdx)
    {
        foreach (var tablePart in worksheet.TableDefinitionParts.ToList())
        {
            var tbl = tablePart.Table;
            var refStr = tbl?.Reference?.Value;
            if (tbl == null || string.IsNullOrEmpty(refStr)) continue;

            var rangeParts = refStr.Split(':');
            int startColIdx, endColIdx;
            try
            {
                startColIdx = ColumnNameToIndex(ParseCellReference(rangeParts[0]).Column);
                endColIdx = rangeParts.Length > 1
                    ? ColumnNameToIndex(ParseCellReference(rangeParts[1]).Column)
                    : startColIdx;
            }
            catch { continue; }

            var tableColumns = tbl.TableColumns;
            if (tableColumns == null) continue;
            var cols = tableColumns.Elements<TableColumn>().ToList();
            int width = endColIdx - startColIdx + 1;
            // Only tables whose ref actually WIDENED (insert landed inside the
            // span) need a new column; a table shifted wholesale to the right
            // keeps width == count.
            if (width <= cols.Count) continue;

            int pos = insertColIdx - startColIdx;
            if (pos < 0) pos = 0;
            if (pos > cols.Count) pos = cols.Count;

            var used = new HashSet<string>(
                cols.Select(tc => tc.Name?.Value ?? "").Where(n => n.Length > 0),
                StringComparer.OrdinalIgnoreCase);
            var baseName = $"Column{cols.Count + 1}";
            var colName = baseName;
            int dedupeIdx = 2;
            while (!used.Add(colName)) colName = $"{baseName}{dedupeIdx++}";

            var newCol = new TableColumn { Name = colName };
            if (pos == 0) tableColumns.PrependChild(newCol);
            else if (pos >= cols.Count) tableColumns.AppendChild(newCol);
            else cols[pos - 1].InsertAfterSelf(newCol);

            tableColumns.Count = (uint)tableColumns.Elements<TableColumn>().Count();
            NormalizeTableColumns(tbl, tableColumns);

            // Header tables: Excel requires the header-row cell text to match
            // the tableColumn name; an inserted column leaves its header cell
            // empty, which Excel refuses (0x800A03EC). Write the name into it.
            if ((tbl.HeaderRowCount?.Value ?? 1) != 0)
            {
                var (_, headerRow) = ParseCellReference(rangeParts[0]);
                var headerCellRef = $"{IndexToColumnName(insertColIdx)}{headerRow}";
                var hdrWs = GetSheet(worksheet);
                var hdrSheetData = hdrWs.GetFirstChild<SheetData>()
                    ?? hdrWs.AppendChild(new SheetData());
                var hdrCell = FindOrCreateCell(hdrSheetData, headerCellRef);
                hdrCell.CellValue = new CellValue(newCol.Name?.Value ?? colName);
                hdrCell.DataType = CellValues.String;
            }
            tbl.Save();
        }
    }

    // ==================== Shift helpers ====================

    /// <summary>
    /// Shift row numbers in a cell/range reference after a row deletion.
    /// Single cell: returns null when it sits on the deleted row. Range: shrinks
    /// — an endpoint on the deleted row collapses inward (start clamps to the
    /// deleted row, end drops by one); endpoints after the deleted row decrement
    /// by one. The range is dropped only when it collapses entirely (its whole
    /// span was the deleted row), mirroring how Excel keeps A1:A4 as A1:A3 after
    /// deleting row 1 instead of discarding the structure.
    /// </summary>
    private static string? ShiftRowInRef(string? refStr, int deletedRow)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');

        if (parts.Length == 1)
        {
            try
            {
                var (col, row) = ParseCellReference(parts[0]);
                if (row == deletedRow) return null;
                return row > deletedRow ? $"{col}{row - 1}" : parts[0];
            }
            catch { return refStr; }
        }

        try
        {
            var (startCol, startRow) = ParseCellReference(parts[0]);
            var (endCol, endRow) = ParseCellReference(parts[1]);
            // start clamps: only moves when strictly below the deleted row.
            int newStart = startRow > deletedRow ? startRow - 1 : startRow;
            // end clamps: moves when at or below the deleted row (loses that row).
            int newEnd = endRow >= deletedRow ? endRow - 1 : endRow;
            if (newStart > newEnd) return null; // whole span was the deleted row
            return $"{startCol}{newStart}:{endCol}{newEnd}";
        }
        catch { return refStr; }
    }

    /// <summary>
    /// Shift column letters in a cell/range reference after a column deletion.
    /// Single cell: returns null when it sits on the deleted column. Range:
    /// shrinks the same way <see cref="ShiftRowInRef"/> does for rows; dropped
    /// only when the whole span was the deleted column.
    /// </summary>
    private static string? ShiftColInRef(string? refStr, int deletedColIdx)
    {
        if (string.IsNullOrEmpty(refStr)) return null;
        var parts = refStr.Split(':');

        if (parts.Length == 1)
        {
            try
            {
                var (col, row) = ParseCellReference(parts[0]);
                var colIdx = ColumnNameToIndex(col);
                if (colIdx == deletedColIdx) return null;
                return colIdx > deletedColIdx ? $"{IndexToColumnName(colIdx - 1)}{row}" : parts[0];
            }
            catch { return refStr; }
        }

        try
        {
            var (startCol, startRow) = ParseCellReference(parts[0]);
            var (endCol, endRow) = ParseCellReference(parts[1]);
            int startIdx = ColumnNameToIndex(startCol);
            int endIdx = ColumnNameToIndex(endCol);
            int newStart = startIdx > deletedColIdx ? startIdx - 1 : startIdx;
            int newEnd = endIdx >= deletedColIdx ? endIdx - 1 : endIdx;
            if (newStart > newEnd) return null; // whole span was the deleted column
            return $"{IndexToColumnName(newStart)}{startRow}:{IndexToColumnName(newEnd)}{endRow}";
        }
        catch { return refStr; }
    }

    // ShiftNamedRangeRows / ShiftNamedRangeCols removed — see comment above
    // about ShiftNamedRangeRowsDown/ColsRight; same consolidation.

    // ==================== Formula impact detection ====================

    private record FormulaImpact(string CellRef, bool IsRefError);

    /// <summary>
    /// Find all surviving cells with formulas that reference the deleted row (→ #REF!) or rows after it (→ shifted).
    /// </summary>
    private List<FormulaImpact> CollectFormulaCellsAffectedByRowDelete(WorksheetPart worksheet, int deletedRow)
    {
        var affected = new List<FormulaImpact>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return affected;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var formula = cell.CellFormula?.Text;
                if (string.IsNullOrEmpty(formula)) continue;

                bool refError = FormulaReferencesExactRow(formula, deletedRow);
                bool shifted = !refError && FormulaReferencesRowAbove(formula, deletedRow);

                if (refError || shifted)
                    affected.Add(new FormulaImpact(cell.CellReference?.Value ?? "?", refError));
            }
        }
        return affected;
    }

    private static bool FormulaReferencesExactRow(string formula, int row)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?[A-Z]+\$?(\d+)", RegexOptions.IgnoreCase))
        {
            if (int.TryParse(m.Groups[1].Value, out var r) && r == row)
                return true;
        }
        return false;
    }

    private static bool FormulaReferencesRowAbove(string formula, int deletedRow)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?[A-Z]+\$?(\d+)", RegexOptions.IgnoreCase))
        {
            if (int.TryParse(m.Groups[1].Value, out var row) && row > deletedRow)
                return true;
        }
        return false;
    }

    /// <summary>
    /// Find all surviving cells with formulas that reference the deleted column (→ #REF!) or columns after it (→ shifted).
    /// </summary>
    private List<FormulaImpact> CollectFormulaCellsAffectedByColDelete(WorksheetPart worksheet, int deletedColIdx)
    {
        var affected = new List<FormulaImpact>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return affected;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var formula = cell.CellFormula?.Text;
                if (string.IsNullOrEmpty(formula)) continue;

                bool refError = FormulaReferencesExactCol(formula, deletedColIdx);
                bool shifted = !refError && FormulaReferencesColAbove(formula, deletedColIdx);

                if (refError || shifted)
                    affected.Add(new FormulaImpact(cell.CellReference?.Value ?? "?", refError));
            }
        }
        return affected;
    }

    private static bool FormulaReferencesExactCol(string formula, int colIdx)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?([A-Z]+)\$?\d+", RegexOptions.IgnoreCase))
        {
            if (ColumnNameToIndex(m.Groups[1].Value.ToUpperInvariant()) == colIdx)
                return true;
        }
        return false;
    }

    private static bool FormulaReferencesColAbove(string formula, int deletedColIdx)
    {
        foreach (Match m in Regex.Matches(formula, @"\$?([A-Z]+)\$?\d+", RegexOptions.IgnoreCase))
        {
            if (ColumnNameToIndex(m.Groups[1].Value.ToUpperInvariant()) > deletedColIdx)
                return true;
        }
        return false;
    }

    private static string? FormatFormulaWarning(List<FormulaImpact> affected)
    {
        if (affected.Count == 0) return null;

        var refErrors = affected.Where(a => a.IsRefError).Select(a => a.CellRef).ToList();
        var shifted = affected.Where(a => !a.IsRefError).Select(a => a.CellRef).ToList();

        var parts = new List<string>();
        if (refErrors.Count > 0)
            parts.Add($"{refErrors.Count} cell(s) will become #REF!: {string.Join(", ", refErrors)}");
        if (shifted.Count > 0)
            parts.Add($"{shifted.Count} cell(s) reference shifted rows/cols (formula text unchanged): {string.Join(", ", shifted)}");

        return $"Warning: {affected.Count} formula cell(s) affected — {string.Join("; ", parts)}";
    }

    /// <summary>
    // ShiftRowNumbersInText / ShiftColLettersInText removed — defined-name
    // text is now rewritten by section 8 of ApplySheetRangeMutations using
    // FormulaRefShifter, which correctly handles quoted sheet names, string
    // literals, and structured refs that the regex shifters mishandled.

    /// <summary>
    /// R9-1: after a sheet is removed, walk every remaining worksheet's
    /// formula cells and clear the CellValue on any formula that still
    /// references the removed sheet by name (bare or single-quote wrapped).
    /// We do not rewrite the formula body — that is Excel's job on recalc.
    /// Clearing the cached value keeps officecli's Get consistent with the
    /// state Real Excel presents when it opens the file.
    /// </summary>
    private void InvalidateFormulaCacheReferencingSheet(WorkbookPart workbookPart, string removedSheetName)
    {
        // Two literal match forms Excel uses for sheet-qualified refs:
        //   Sheet2!A1             (bare, no special chars)
        //   'My Data'!A1          (quoted when name has spaces/specials)
        // Internal single quotes in sheet names are escaped as '' inside
        // the quoted form, but creating such names is rare and the
        // Contains check below still handles the unescaped prefix.
        var bareToken = removedSheetName + "!";
        var quotedToken = "'" + removedSheetName.Replace("'", "''") + "'!";

        foreach (var wsPart in workbookPart.WorksheetParts)
        {
            var sheetData = GetSheet(wsPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            bool touched = false;
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var formula = cell.CellFormula?.Text;
                    if (string.IsNullOrEmpty(formula)) continue;
                    if (formula.IndexOf(bareToken, StringComparison.OrdinalIgnoreCase) < 0 &&
                        formula.IndexOf(quotedToken, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    // Clear the cached value. CellValue element removed so
                    // Get reports null/missing cachedValue, matching Excel's
                    // initial state on open (before recalc fills in #REF!).
                    cell.CellValue?.Remove();
                    touched = true;
                }
            }

            if (touched)
            {
                GetSheet(wsPart).Save();
            }
        }
    }

    /// <summary>
    /// R10-2 / R2-1 shared helper. Drops a PivotTableCacheDefinitionPart and
    /// its workbook-level &lt;pivotCache&gt; entry IF no remaining pivot
    /// table part references it. Used by both the sheet-remove and the
    /// pivottable[N]-remove code paths so the orphan-cleanup logic stays
    /// in one place.
    /// </summary>
    private static void PrunePivotCacheIfOrphan(WorkbookPart workbookPart, PivotTableCacheDefinitionPart cachePart)
    {
        bool stillReferenced = workbookPart.WorksheetParts
            .SelectMany(ws => ws.PivotTableParts)
            .Any(pp => pp.PivotTableCacheDefinitionPart == cachePart);
        if (stillReferenced) return;

        // Locate and remove the <pivotCache> entry in workbook.xml by
        // matching the relationship id from WorkbookPart → cachePart.
        string? cacheRelId = null;
        try { cacheRelId = workbookPart.GetIdOfPart(cachePart); } catch { }

        var wb = workbookPart.Workbook;
        if (wb != null)
        {
            var pivotCaches = wb.GetFirstChild<PivotCaches>();
            if (pivotCaches != null && cacheRelId != null)
            {
                var pcEntry = pivotCaches.Elements<PivotCache>()
                    .FirstOrDefault(pc => pc.Id?.Value == cacheRelId);
                pcEntry?.Remove();
                if (!pivotCaches.HasChildren)
                    pivotCaches.Remove();
            }
            try { workbookPart.DeletePart(cachePart); } catch { }
            wb.Save();
        }
        else
        {
            try { workbookPart.DeletePart(cachePart); } catch { }
        }
    }
}
