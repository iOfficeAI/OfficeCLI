// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ===== Chart-cache hygiene (stale numCache/strCache refresh at persist) =====
    //
    // A chart series references a cell range via <c:f> and carries a cached
    // snapshot of the referenced values in <c:numCache>/<c:strCache>. officecli
    // seeds that cache when the chart is created (see ParseDataRangeForChart /
    // BackfillSeriesRangeValues), but a later `set` on a referenced (or
    // formula-precedent) cell only refreshes the WORKSHEET, never the chart
    // cache. Real Excel recomputes charts on open, so the staleness is invisible
    // there — but offline consumers read the cache directly:
    //   - `dump --format batch` serializes the stale points (wrong series values);
    //   - dump→replay recomputes and produces DIFFERENT points, so d1/d2 are not
    //     idempotent;
    //   - `view html` and other static renderers draw the outdated chart.
    //
    // At persist (after RefreshStaleFormulaCaches, so formula precedents are
    // reconciled first) we re-seed every chart numRef/strRef whose <c:f> resolves
    // to a plain cell range against the CURRENT cell values. Formula cells are
    // evaluated with the same FormulaEvaluator the dump path uses — the on-disk
    // worksheet cache may have been dropped by the L2 formula-cache sweep, so
    // reading <v> alone would under-report; we must evaluate.
    //
    // Scope is deliberately coarse (every simple-range ref is reseeded, not a
    // precise cell-dependency graph). References we cannot resolve — named
    // ranges, multi-area refs, cross-workbook, missing sheets — are left
    // untouched, matching the chart-cache-stale scanner in ExcelHandler.View.cs.

    /// <summary>
    /// Re-seed chart numCache/strCache snapshots from current cell values.
    /// Called from <see cref="FlushDirtyParts"/> at persist, after the formula
    /// cache sweep. No-op unless this session changed something and the workbook
    /// has at least one standard ChartPart.
    /// </summary>
    private void RefreshStaleChartCaches()
    {
        if (_doc.WorkbookPart == null) return;
        if (!Modified && _dirtyWorksheets.Count == 0) return;

        var worksheets = GetWorksheets();
        var hasChart = worksheets.Any(w =>
            w.Part.DrawingsPart is { } dp && dp.ChartParts.Any());
        if (!hasChart) return;

        // Per-sheet evaluator + cell index, built lazily and shared across all
        // charts (charts on any sheet may reference any sheet by name).
        var sheetCtx = new Dictionary<string,
            (SheetData sd, Core.FormulaEvaluator ev, Dictionary<string, Cell> idx)>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var (_, wsPart) in worksheets)
        {
            if (wsPart.DrawingsPart is not { } dp) continue;
            foreach (var cp in dp.ChartParts)
            {
                if (cp.ChartSpace is null) continue;
                var changed = false;
                foreach (var numRef in cp.ChartSpace.Descendants<C.NumberReference>())
                {
                    var f = numRef.Formula?.Text;
                    if (string.IsNullOrWhiteSpace(f)) continue;
                    var cells = ResolveChartRangeCells(f, worksheets, sheetCtx);
                    if (cells == null) continue;
                    if (ReseedNumberingCache(numRef, cells)) changed = true;
                }
                foreach (var strRef in cp.ChartSpace.Descendants<C.StringReference>())
                {
                    var f = strRef.Formula?.Text;
                    if (string.IsNullOrWhiteSpace(f)) continue;
                    var cells = ResolveChartRangeCells(f, worksheets, sheetCtx);
                    if (cells == null) continue;
                    if (ReseedStringCache(strRef, cells)) changed = true;
                }
                if (changed) cp.ChartSpace.Save();
            }
        }
    }

    /// <summary>
    /// Resolve a chart <c:f> formula ("Sheet1!$A$1:$A$3" / "'Quoted'!B2") to the
    /// ordered (row-major) list of referenced cells, evaluating each into raw
    /// value text. Formula cells are computed via the sheet's FormulaEvaluator
    /// (the on-disk <v> may be absent after the L2 formula-cache sweep). Returns
    /// null for anything that is not a single-sheet plain cell/range reference.
    /// </summary>
    private List<string>? ResolveChartRangeCells(
        string formula,
        List<(string Name, WorksheetPart Part)> worksheets,
        Dictionary<string, (SheetData sd, Core.FormulaEvaluator ev, Dictionary<string, Cell> idx)> sheetCtx)
    {
        var bang = formula.IndexOf('!');
        if (bang <= 0) return null;
        var sheetPart = formula[..bang].Trim();
        if (sheetPart.StartsWith('\'') && sheetPart.EndsWith('\''))
            sheetPart = sheetPart[1..^1].Replace("''", "'");
        // Multi-area / 3D refs (contain a colon in the sheet portion or a comma)
        // are out of scope.
        if (sheetPart.Contains(':') || formula.Contains(',')) return null;
        var rangePart = formula[(bang + 1)..].Replace("$", "");

        var parts = rangePart.Split(':');
        var first = parts[0];
        var last = parts.Length > 1 ? parts[1] : parts[0];
        if (!System.Text.RegularExpressions.Regex.IsMatch(first, "^[A-Z]+\\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return null;
        if (!System.Text.RegularExpressions.Regex.IsMatch(last, "^[A-Z]+\\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return null;

        if (!sheetCtx.TryGetValue(sheetPart, out var ctx))
        {
            var wsPart = worksheets
                .FirstOrDefault(w => string.Equals(w.Name, sheetPart, StringComparison.OrdinalIgnoreCase))
                .Part;
            if (wsPart == null) return null;
            var sheetData = GetSheet(wsPart).GetFirstChild<SheetData>();
            if (sheetData == null) return null;
            var idx = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in sheetData.Elements<Row>())
                foreach (var cell in row.Elements<Cell>())
                    if (cell.CellReference?.Value is { } cr) idx[cr] = cell;
            ctx = (sheetData, new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart), idx);
            sheetCtx[sheetPart] = ctx;
        }

        var (col1Str, r1) = ParseCellReference(first.ToUpperInvariant());
        var (col2Str, r2) = ParseCellReference(last.ToUpperInvariant());
        int c1 = ColumnNameToIndex(col1Str), c2 = ColumnNameToIndex(col2Str);
        if (r2 < r1 || c2 < c1) return null;

        var values = new List<string>();
        for (int r = r1; r <= r2; r++)
            for (int c = c1; c <= c2; c++)
            {
                var addr = $"{IndexToColumnName(c)}{r}";
                if (!ctx.idx.TryGetValue(addr, out var cell)) { values.Add(""); continue; }
                values.Add(ResolveCellRawText(cell, ctx.ev));
            }
        return values;
    }

    /// <summary>
    /// Raw cell text for chart-cache seeding: formula cells are evaluated (their
    /// on-disk <v> may be gone after the L2 sweep), shared/inline strings are
    /// resolved to text, everything else returns its stored <v>. Unlike
    /// GetCellDisplayValue this does NOT apply the cell number format — the chart
    /// cache stores raw values, formatting lives on the chart.
    /// </summary>
    private string ResolveCellRawText(Cell cell, Core.FormulaEvaluator evaluator)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
            return RstTextWithoutPhonetic(cell.InlineString);
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var value = cell.CellValue?.Text ?? "";
            var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sst?.SharedStringTable != null && int.TryParse(value, out int sidx))
            {
                var ssi = SharedStringAt(sst.SharedStringTable, sidx);
                return ssi != null ? RstTextWithoutPhonetic(ssi) : value;
            }
            return value;
        }
        if (cell.CellFormula?.Text is { } ft)
        {
            if (!FormulaReferencesMissingSheet(ft))
            {
                var report = evaluator.EvaluateForReport(ft);
                if (report.Status == Core.EvalReportStatus.Evaluated)
                    return report.Result!.ToCellValueText();
                if (report.Status == Core.EvalReportStatus.Error)
                    return report.Result!.ErrorValue!;
            }
            // Fall back to whatever cache remains.
            return cell.CellValue?.Text ?? "";
        }
        return cell.CellValue?.Text ?? "";
    }

    /// <summary>
    /// Refresh a numRef's NumericPoint values in place from resolved cell values.
    /// The cache STRUCTURE (which indices carry a point, the PointCount, and the
    /// FormatCode) is preserved verbatim — only the numeric text of existing
    /// points is updated. This keeps whatever dispBlanksAs gap structure the
    /// builder baked (a missing index stays missing; a present blank stays 0)
    /// while correcting stale numbers. Non-numeric resolved text maps to 0,
    /// matching the builder. Returns true if any point text changed.
    /// </summary>
    private static bool ReseedNumberingCache(C.NumberReference numRef, List<string> values)
    {
        var cache = numRef.NumberingCache;
        if (cache == null) return false;
        var changed = false;
        // Drop phantom points beyond the (possibly shrunk) reference range —
        // a row delete that shrinks valuesRef leaves stale trailing points
        // that duplicate the last value in the live file. Remove any point
        // whose index is past the new range, and sync PointCount.
        foreach (var stale in cache.Elements<C.NumericPoint>()
                     .Where(p => (int)(p.Index?.Value ?? 0u) >= values.Count).ToList())
        {
            stale.Remove();
            changed = true;
        }
        if (cache.PointCount != null && cache.PointCount.Val?.Value != (uint)values.Count)
        {
            cache.PointCount.Val = (uint)values.Count;
            changed = true;
        }
        foreach (var pt in cache.Elements<C.NumericPoint>())
        {
            var idx = (int)(pt.Index?.Value ?? 0u);
            if (idx < 0 || idx >= values.Count) continue;
            var raw = values[idx];
            // Present-but-blank source under a non-gap cache resolves to 0
            // (the builder's default). Gap-mode blanks have no point here.
            if (!double.TryParse(raw, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var num))
                num = 0;
            var text = num.ToString("G", System.Globalization.CultureInfo.InvariantCulture);
            if ((pt.NumericValue?.Text ?? "") != text)
            {
                pt.NumericValue = new C.NumericValue(text);
                changed = true;
            }
        }
        return changed;
    }

    /// <summary>
    /// Refresh a strRef's StringPoint values in place from resolved cell values.
    /// Structure (indices, PointCount) is preserved; only the text of existing
    /// points is updated. Returns true if any point text changed.
    /// </summary>
    private static bool ReseedStringCache(C.StringReference strRef, List<string> values)
    {
        var cache = strRef.StringCache;
        if (cache == null) return false;
        var changed = false;
        // Drop phantom points beyond the shrunk reference range (see numeric
        // cache above) and sync PointCount.
        foreach (var stale in cache.Elements<C.StringPoint>()
                     .Where(p => (int)(p.Index?.Value ?? 0u) >= values.Count).ToList())
        {
            stale.Remove();
            changed = true;
        }
        if (cache.PointCount != null && cache.PointCount.Val?.Value != (uint)values.Count)
        {
            cache.PointCount.Val = (uint)values.Count;
            changed = true;
        }
        foreach (var pt in cache.Elements<C.StringPoint>())
        {
            var idx = (int)(pt.Index?.Value ?? 0u);
            if (idx < 0 || idx >= values.Count) continue;
            var text = values[idx] ?? "";
            if ((pt.NumericValue?.Text ?? "") != text)
            {
                pt.NumericValue = new C.NumericValue(text);
                changed = true;
            }
        }
        return changed;
    }
}
