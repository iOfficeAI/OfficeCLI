// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{

    // Map a table-column totals-row function token to its OOXML enum and the
    // SUBTOTAL function code Excel uses. Unknown tokens throw — the earlier
    // SUM fallback silently changed the aggregation the user asked for
    // (silent-accept enum-miss family). Every token the dump emitter can
    // produce (TotalsRowFunction InnerText, lowercased) is enumerated below,
    // so dump→batch replay never hits the throw.
    internal static (TotalsRowFunctionValues, int) MapTotalsRowFunction(string tok) => tok switch
    {
        "sum" => (TotalsRowFunctionValues.Sum, 109),
        "average" or "avg" => (TotalsRowFunctionValues.Average, 101),
        "count" => (TotalsRowFunctionValues.Count, 103),
        "countnums" or "countnumbers" => (TotalsRowFunctionValues.CountNumbers, 102),
        "max" or "maximum" => (TotalsRowFunctionValues.Maximum, 104),
        "min" or "minimum" => (TotalsRowFunctionValues.Minimum, 105),
        "stddev" or "stdev" => (TotalsRowFunctionValues.StandardDeviation, 107),
        "var" or "variance" => (TotalsRowFunctionValues.Variance, 110),
        "none" or "label" or "" => (TotalsRowFunctionValues.None, 0),
        "custom" => (TotalsRowFunctionValues.Custom, 109),
        _ => throw new ArgumentException(
            $"Unknown totals-row function '{tok}'. Valid: sum, average, count, countNums, max, min, stdDev, var, none, custom.")
    };

    private List<SharedStringItem>? _sharedStringItemCache;
    private List<string>? _sharedStringCache;

    private List<SharedStringItem> GetSharedStringItemCache()
    {
        if (_sharedStringItemCache != null) return _sharedStringItemCache;
        var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        if (sst?.SharedStringTable == null) return _sharedStringItemCache = new List<SharedStringItem>();

        var list = new List<SharedStringItem>();
        foreach (var item in sst.SharedStringTable.Elements<SharedStringItem>())
        {
            list.Add(item);
        }
        return _sharedStringItemCache = list;
    }

    private List<string> GetSharedStringCache()
    {
        if (_sharedStringCache != null) return _sharedStringCache;
        var items = GetSharedStringItemCache();
        var list = new List<string>(items.Count);
        foreach (var item in items)
        {
            list.Add(item.InnerText ?? "");
        }
        return _sharedStringCache = list;
    }

    /// <summary>
    /// Must be called whenever a new SharedStringItem is appended to or removed
    /// from the SharedStringTable so that the next GetCellDisplayValue call
    /// rebuilds the cache from the updated table.
    /// </summary>
    private void InvalidateSharedStringCache()
    {
        _sharedStringCache = null;
        _sharedStringItemCache = null;
    }



    private string GetCellDisplayValue(Cell cell, Core.FormulaEvaluator? evaluator = null)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? "";
        }

        var value = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (int.TryParse(value, out int idx))
            {
                var cache = GetSharedStringCache();
                if (idx >= 0 && idx < cache.Count)
                    return cache[idx];
            }
            return value;
        }

        // Boolean cells store 0/1 in <v> per the OOXML spec, but Excel displays
        // (and users expect) TRUE/FALSE. Decode it here so .Text matches what
        // Excel renders — the write side already accepts TRUE/FALSE, and the
        // dump emitter reads the raw 0/1 separately, so this is display-only.
        if (cell.DataType?.Value == CellValues.Boolean)
            return value == "1" ? "TRUE" : "FALSE";

        // Formula cells: if there's a cached value, return it.
        // If not, try to evaluate; otherwise emit a sentinel so callers can
        // distinguish "formula not evaluated" from "cell contains the literal
        // text `=FOO`". The sentinel matches Excel's `#…!` error-code shape
        // so it sorts visually next to #REF!/#VALUE!/etc.
        if (string.IsNullOrEmpty(value) && cell.CellFormula?.Text != null)
        {
            // Missing-sheet refs: ResolveSheetCellResult silently returns 0
            // and the error path surfaces a fake #REF!. Neither value is
            // trustworthy, so this branch matches the BuildCellNode contract
            // that suppresses computedValue and reports evaluated=false —
            // view text must emit the sentinel to keep the two readbacks
            // (view text vs Format["evaluated"]) consistent.
            if (FormulaReferencesMissingSheet(cell.CellFormula.Text))
                return "#OCLI_NOTEVAL!";
            if (evaluator != null)
            {
                var report = evaluator.EvaluateForReport(cell.CellFormula.Text);
                if (report.Status == Core.EvalReportStatus.Evaluated)
                    return report.Result!.ToCellValueText();
                // Error values (#DIV/0!, #VALUE!, …) surface directly so users
                // see what Excel would. Only NotEvaluated falls to the sentinel.
                if (report.Status == Core.EvalReportStatus.Error)
                    return report.Result!.ErrorValue!;
            }
            // Sentinel form #OCLI_NOTEVAL! (no formula tail) — OCLI_ prefix
            // keeps us out of Excel's reserved #XXX! error-code namespace
            // (#REF!, #VALUE!, …, and future #CALC!/#SPILL!). The formula
            // text stays available via Format["formula"] and `view issues`.
            return "#OCLI_NOTEVAL!";
        }

        // Apply number format to numeric cells (dates, percentages, etc.)
        // raw double + format code → display string
        if (cell.DataType == null && double.TryParse(value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
        {
            var (numFmtId, formatCode) = ExcelDataFormatter.GetCellFormat(cell, _doc.WorkbookPart);
            if (numFmtId > 0)
            {
                var is1904 = IsWorkbookDate1904();
                var formatted = ExcelDataFormatter.TryFormat(numVal, numFmtId, formatCode, is1904);
                if (formatted != null) return formatted;
            }
        }

        return value;
    }

    /// <summary>
    /// True when the workbook uses the 1904 date system (Date1904=true). The
    /// stored serial for a date differs from the 1900 system by 1462 days, so
    /// every date write and read must consult this flag to stay consistent.
    /// </summary>
    private bool IsWorkbookDate1904()
        => _doc.WorkbookPart?.Workbook?.WorkbookProperties?.Date1904?.Value == true;

    // Underlying stored value for COMPARISON, as opposed to GetCellDisplayValue's
    // formatted string. A percentage cell compares as 0.5 (not "50%") and a date
    // as its serial (not "2024-01-15"), so `row[Pct>0.3]` / `row[Hired>45000]`
    // evaluate against the real number instead of the display text. Non-numeric
    // cells (text, shared strings, formula results, sentinels) are identical to
    // the display path, so equality on a formatted literal still works.
    private string GetCellRawComparisonValue(Cell cell, Core.FormulaEvaluator? evaluator = null)
    {
        if (cell.DataType == null)
        {
            var raw = cell.CellValue?.Text;
            if (!string.IsNullOrEmpty(raw) && double.TryParse(raw,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out _))
                return raw;
        }
        return GetCellDisplayValue(cell, evaluator);
    }

    private static bool IsCellInMergeRange(string cellRef, string? rangeRef)
    {
        if (string.IsNullOrEmpty(rangeRef) || !rangeRef.Contains(':')) return false;
        var parts = rangeRef.Split(':');
        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);
        var (cellCol, cellRow) = ParseCellReference(cellRef);

        var cellColIdx = ColumnNameToIndex(cellCol);
        return cellRow >= startRow && cellRow <= endRow
            && cellColIdx >= ColumnNameToIndex(startCol) && cellColIdx <= ColumnNameToIndex(endCol);
    }

    // T4 — rectangle intersection over A1:B2 style ranges (case-insensitive).
    // Returns true if two inclusive cell ranges share at least one cell.
    private static bool RangesOverlap(string rangeA, string rangeB)
    {
        if (string.IsNullOrEmpty(rangeA) || string.IsNullOrEmpty(rangeB)) return false;
        // Whole-column (A:A) and whole-row (1:3) tokens are legal sqref
        // members (ValidateSqref admits them), but ParseCellReference below
        // rejects a bare "A"/"1" — a second dataValidation Add on a sheet
        // holding any whole-row/col range threw even when geometrically
        // disjoint. Expand them to explicit rectangles first.
        var expandedA = ExpandWholeRowColRange(rangeA);
        var expandedB = ExpandWholeRowColRange(rangeB);
        var (a1, a2) = SplitRange(expandedA);
        var (b1, b2) = SplitRange(expandedB);
        var (aSc, aSr) = ParseCellReference(a1);
        var (aEc, aEr) = ParseCellReference(a2);
        var (bSc, bSr) = ParseCellReference(b1);
        var (bEc, bEr) = ParseCellReference(b2);
        int aSci = ColumnNameToIndex(aSc), aEci = ColumnNameToIndex(aEc);
        int bSci = ColumnNameToIndex(bSc), bEci = ColumnNameToIndex(bEc);
        // Normalize (callers may pass B2:A1 theoretically)
        if (aSci > aEci) (aSci, aEci) = (aEci, aSci);
        if (bSci > bEci) (bSci, bEci) = (bEci, bSci);
        if (aSr > aEr) (aSr, aEr) = (aEr, aSr);
        if (bSr > bEr) (bSr, bEr) = (bEr, bSr);
        return aSci <= bEci && bSci <= aEci && aSr <= bEr && bSr <= aEr;
    }

    /// <summary>
    /// Canonicalize an inverted CELL:CELL range (D5:A1 → A1:D5, per axis).
    /// Well-formed input passes through untouched (keeps $ anchors);
    /// inverted input is rebuilt without $ (matching the printArea rule).
    /// </summary>
    internal static string NormalizeA1Range(string range)
    {
        var nm = System.Text.RegularExpressions.Regex.Match(range.Trim(),
            @"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!nm.Success) return range;
        var nc1 = ColumnNameToIndex(nm.Groups[1].Value.ToUpperInvariant());
        var nc2 = ColumnNameToIndex(nm.Groups[3].Value.ToUpperInvariant());
        var nr1 = long.Parse(nm.Groups[2].Value);
        var nr2 = long.Parse(nm.Groups[4].Value);
        if (nc1 <= nc2 && nr1 <= nr2) return range;
        var colA = nc1 <= nc2 ? nm.Groups[1].Value : nm.Groups[3].Value;
        var colB = nc1 <= nc2 ? nm.Groups[3].Value : nm.Groups[1].Value;
        return $"{colA}{Math.Min(nr1, nr2)}:{colB}{Math.Max(nr1, nr2)}";
    }

    // A:A → A1:A1048576, 1:3 → A1:XFD3. Non-whole ranges pass through.
    private static string ExpandWholeRowColRange(string range)
    {
        var wm = System.Text.RegularExpressions.Regex.Match(range.Trim(),
            @"^\$?([A-Z]+)\$?:\$?([A-Z]+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wm.Success)
            return $"{wm.Groups[1].Value}1:{wm.Groups[2].Value}1048576";
        var rm = System.Text.RegularExpressions.Regex.Match(range.Trim(),
            @"^\$?([0-9]+)\$?:\$?([0-9]+)$");
        if (rm.Success)
            return $"A{rm.Groups[1].Value}:XFD{rm.Groups[2].Value}";
        return range;
    }

    private static (string, string) SplitRange(string range)
    {
        if (!range.Contains(':')) return (range, range);
        var p = range.Split(':');
        return (p[0], p[1]);
    }

    // CONSISTENCY(merge-precision): list every existing <mergeCell> whose
    // ref lies entirely inside `outerRange` (inclusive rectangle containment).
    // Used by range-level unmerge to surface precise refs when the caller's
    // range covers sub-merges but does not equal one — see ExcelHandler.Set
    // SetRange merge=false branch.
    private static List<string> FindMergesContainedIn(MergeCells mergeCells, string outerRange)
    {
        var hits = new List<string>();
        var (o1, o2) = SplitRange(outerRange);
        var (oSc, oSr) = ParseCellReference(o1);
        var (oEc, oEr) = ParseCellReference(o2);
        int oSci = ColumnNameToIndex(oSc), oEci = ColumnNameToIndex(oEc);
        if (oSci > oEci) (oSci, oEci) = (oEci, oSci);
        if (oSr > oEr) (oSr, oEr) = (oEr, oSr);
        foreach (var mc in mergeCells.Elements<MergeCell>())
        {
            if (mc.Reference?.Value is not string r) continue;
            var (m1, m2) = SplitRange(r.ToUpperInvariant());
            var (mSc, mSr) = ParseCellReference(m1);
            var (mEc, mEr) = ParseCellReference(m2);
            int mSci = ColumnNameToIndex(mSc), mEci = ColumnNameToIndex(mEc);
            if (mSci > mEci) (mSci, mEci) = (mEci, mSci);
            if (mSr > mEr) (mSr, mEr) = (mEr, mSr);
            if (mSci >= oSci && mEci <= oEci && mSr >= oSr && mEr <= oEr)
                hits.Add(r);
        }
        return hits;
    }

    private static void InsertMergeCellChecked(MergeCells mergeCells, string newRangeRef, WorksheetPart? worksheetPart = null)
    {
        try
        {
            InsertMergeCellCheckedCore(mergeCells, newRangeRef, worksheetPart);
        }
        catch
        {
            // Callers create the <mergeCells> container before this check
            // runs. If the merge is rejected and the container is left
            // childless, an empty <x:mergeCells/> stays in the worksheet —
            // schema-INVALID (count >= 1 required), so a failed call
            // corrupted a previously-fine file. Remove the empty shell
            // before rethrowing.
            if (!mergeCells.Elements<MergeCell>().Any() && mergeCells.Parent != null)
                mergeCells.Remove();
            throw;
        }
    }

    private static void InsertMergeCellCheckedCore(MergeCells mergeCells, string newRangeRef, WorksheetPart? worksheetPart = null)
    {
        ValidateMergeRefLiteral(newRangeRef);
        var refUpper = newRangeRef.ToUpperInvariant();
        foreach (var existing in mergeCells.Elements<MergeCell>())
        {
            if (existing.Reference?.Value is not string er) continue;
            var erUpper = er.ToUpperInvariant();
            if (string.Equals(erUpper, refUpper, StringComparison.Ordinal)) return; // idempotent
            if (RangesOverlap(refUpper, erUpper))
                throw new ArgumentException(
                    $"Merge range '{refUpper}' overlaps existing merged range '{er}'. " +
                    $"Excel rejects overlapping mergeCell entries.");
        }
        // BUG-R2-table-merge BUG-5: Excel forbids mergeCell entries that
        // intersect a ListObject table range — files saved with such a
        // merge open with a "found a problem" repair dialog. Reject up
        // front so callers see a clear error instead of file corruption.
        if (worksheetPart != null)
        {
            foreach (var tdp in worksheetPart.TableDefinitionParts)
            {
                var tblRef = tdp.Table?.Reference?.Value;
                if (string.IsNullOrEmpty(tblRef)) continue;
                if (RangesOverlap(refUpper, tblRef.ToUpperInvariant()))
                {
                    var tblName = tdp.Table?.Name?.Value
                        ?? tdp.Table?.DisplayName?.Value
                        ?? "(unnamed)";
                    throw new ArgumentException(
                        $"Merge range '{refUpper}' overlaps ListObject table '{tblName}' (ref '{tblRef}'). " +
                        "Excel does not allow merging cells inside a table range. Convert the table to a normal range first.");
                }
            }
        }
        // Advisory only — a multi-million-cell merge is legal OOXML and
        // validates green, but real Excel stalls for minutes opening it
        // (empirically: 16384x1 renders in seconds; 1024x100000 hangs past
        // a 120s render timeout). Warn so the caller knows the file may be
        // unusable in practice; do not reject (lenient-input convention).
        if (TryParseRangeDims(refUpper, out var mergeRows, out var mergeCols)
            && (long)mergeRows * mergeCols > 10_000_000)
        {
            Console.Error.WriteLine(
                $"Warning: merge range '{refUpper}' spans {(long)mergeRows * mergeCols:N0} cells; " +
                "real Excel can hang for minutes opening merges this large.");
        }
        mergeCells.AppendChild(new MergeCell { Reference = refUpper });
    }

    private static bool TryParseRangeDims(string range, out int rows, out int cols)
    {
        rows = cols = 0;
        var dimParts = range.Split(':');
        if (dimParts.Length != 2) return false;
        var m1 = System.Text.RegularExpressions.Regex.Match(dimParts[0], @"^([A-Z]+)(\d+)$");
        var m2 = System.Text.RegularExpressions.Regex.Match(dimParts[1], @"^([A-Z]+)(\d+)$");
        if (!m1.Success || !m2.Success) return false;
        var c1 = ColumnNameToIndex(m1.Groups[1].Value);
        var c2 = ColumnNameToIndex(m2.Groups[1].Value);
        var r1 = int.Parse(m1.Groups[2].Value);
        var r2 = int.Parse(m2.Groups[2].Value);
        rows = Math.Abs(r2 - r1) + 1;
        cols = Math.Abs(c2 - c1) + 1;
        return true;
    }

    private DocumentNode GetCellRange(string sheetName, SheetData sheetData, string range, int depth, WorksheetPart? part = null)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid range: {range}");

        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{range}",
            Type = "range",
            Preview = range
        };

        // Build lookup of existing cells so we can fill empty stubs for missing positions
        var existingCells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value != null)
                    existingCells[cell.CellReference.Value] = cell;
            }
        }

        // Enumerate every position in the range in row-major order,
        // materializing empty stubs for positions that have no cell element.
        var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startColIdx; c <= endColIdx; c++)
            {
                var cellRef = $"{IndexToColumnName(c)}{r}";
                if (existingCells.TryGetValue(cellRef, out var existingCell))
                    node.Children.Add(CellToNode(sheetName, existingCell, part, eval));
                else
                    node.Children.Add(new DocumentNode
                    {
                        Path = $"/{sheetName}/{cellRef}",
                        Type = "cell",
                        Text = "",
                        Preview = cellRef,
                        Format = { ["type"] = "Number", ["empty"] = true }
                    });
            }
        }

        node.ChildCount = node.Children.Count;
        return node;
    }

    /// <summary>
    /// Parse a cell value for sorting: returns a tuple (rank, numVal, strVal) so that
    /// nulls/empties sort last, numbers sort before strings, and cross-type comparison never occurs.
    /// rank=0 for numbers, rank=1 for strings, rank=2 for empty/null.
    /// </summary>
    private static (int Rank, double NumVal, string StrVal) ParseSortValue(string value)
    {
        if (string.IsNullOrEmpty(value)) return (2, 0.0, "");
        // Excel treats NaN / Infinity / -Infinity as text, not numbers. double.TryParse
        // happily accepts them though, which would make sort order dependent on whether
        // the exact casing matched double.TryParse's spec vs not — classify explicitly.
        if (value.Equals("NaN", StringComparison.Ordinal)
            || value.Equals("Infinity", StringComparison.Ordinal)
            || value.Equals("-Infinity", StringComparison.Ordinal)
            || value.Equals("+Infinity", StringComparison.Ordinal))
            return (1, 0.0, value);
        if (double.TryParse(value, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var num))
        {
            // Defensive: even non-literal inputs can produce non-finite doubles
            // (e.g. "1e999" overflows to +Infinity). Keep those in the string bucket.
            if (!double.IsFinite(num)) return (1, 0.0, value);
            return (0, num, "");
        }
        return (1, 0.0, value);
    }

    private static Cell? FindCell(SheetData sheetData, string cellRef)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                    return cell;
            }
        }
        return null;
    }

    /// <summary>
    /// Find or create the Row for the given 1-based row index, using the per-SheetData
    /// row index cache to avoid O(n) linear scans. New rows are inserted in sorted order
    /// via binary search on the cache (O(log n)).
    /// </summary>
    private Row FindOrCreateRow(SheetData sheetData, uint rowIdx)
    {
        _rowIndex ??= new();
        if (!_rowIndex.TryGetValue(sheetData, out var rowMap))
        {
            rowMap = new SortedList<uint, Row>();
            foreach (var existingRow in sheetData.Elements<Row>())
                if (existingRow.RowIndex?.HasValue == true)
                    rowMap[existingRow.RowIndex.Value] = existingRow;
            _rowIndex[sheetData] = rowMap;
        }

        if (rowMap.TryGetValue(rowIdx, out var row))
            return row;

        row = new Row { RowIndex = rowIdx };
        // Binary search for predecessor in O(log n)
        var keys = rowMap.Keys;
        int lo = 0, hi = keys.Count - 1, predPos = -1;
        while (lo <= hi)
        {
            int mid = (lo + hi) / 2;
            if (keys[mid] < rowIdx) { predPos = mid; lo = mid + 1; }
            else hi = mid - 1;
        }
        if (predPos >= 0)
            rowMap.Values[predPos].InsertAfterSelf(row);
        else
            sheetData.InsertAt(row, 0);
        rowMap[rowIdx] = row;
        return row;
    }

    /// <summary>
    /// Invalidate the row index cache for a specific SheetData (or all sheets if null).
    /// Must be called whenever rows are structurally modified (removed, shifted).
    /// </summary>
    private void InvalidateRowIndex(SheetData? sheetData = null)
    {
        if (sheetData != null)
            _rowIndex?.Remove(sheetData);
        else
            _rowIndex = null;
    }

    private Cell FindOrCreateCell(SheetData sheetData, string cellRef)
    {
        var (colName, rowIdx) = ParseCellReference(cellRef);

        var row = FindOrCreateRow(sheetData, (uint)rowIdx);

        // Cell lookup within row — O(m) where m = cols per row (typically small)
        var cell = row.Elements<Cell>().FirstOrDefault(c =>
            c.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef.ToUpperInvariant() };
            // Insert in column order
            var afterCell = row.Elements<Cell>().LastOrDefault(c =>
            {
                var (cn, _) = ParseCellReference(c.CellReference?.Value ?? "A1");
                return ColumnNameToIndex(cn) < ColumnNameToIndex(colName);
            });
            if (afterCell != null)
                afterCell.InsertAfterSelf(cell);
            else
                row.InsertAt(cell, 0);
        }

        return cell;
    }

    // CONSISTENCY(xlsx/table-autoexpand): custom namespace marker stored on
    // the <x:table> root so `autoExpand=true` survives open/close cycles.
    // Real Excel ignores unknown-namespace attributes, so the file is still
    // opened cleanly on Windows — the flag only affects officecli's own
    // cell-write auto-grow behavior.
    private const string AutoExpandNamespaceUri = "https://officecli.ai/2025/autoexpand";
    private const string AutoExpandNamespacePrefix = "ae";
    private const string AutoExpandAttrName = "autoExpand";

    private static void SetTableAutoExpandMarker(Table table, bool enabled)
    {
        if (enabled)
        {
            table.AddNamespaceDeclaration(AutoExpandNamespacePrefix, AutoExpandNamespaceUri);
            table.SetAttribute(new OpenXmlAttribute(
                AutoExpandNamespacePrefix, AutoExpandAttrName, AutoExpandNamespaceUri, "1"));
        }
    }

    private static bool TableHasAutoExpand(Table? table)
    {
        if (table == null) return false;
        foreach (var attr in table.GetAttributes())
        {
            if (attr.NamespaceUri == AutoExpandNamespaceUri
                && attr.LocalName == AutoExpandAttrName
                && (attr.Value == "1" || string.Equals(attr.Value, "true", StringComparison.OrdinalIgnoreCase)))
                return true;
        }
        return false;
    }

    // Eager auto-grow on cell Add/Set. Called after writing `cellRef` on
    // `worksheet`. For each table on the sheet flagged with autoExpand:
    //   - if cell is in the row immediately below the table AND its column
    //     is within the table's column span → grow endRow by 1.
    //   - else if cell is in the column immediately right of the table AND
    //     its row is within the table's row span → grow endCol by 1 and
    //     append a blank tableColumn.
    // Both extensions are never applied at once (conservative).
    private void MaybeExpandTablesForCell(WorksheetPart worksheet, string cellRef)
    {
        var (cellCol, cellRow) = ParseCellReference(cellRef.ToUpperInvariant());
        var cellColIdx = ColumnNameToIndex(cellCol);

        foreach (var tdp in worksheet.TableDefinitionParts.ToList())
        {
            var table = tdp.Table;
            if (table == null) continue;
            if (!TableHasAutoExpand(table)) continue;
            if (table.Reference?.Value is not string rangeRef) continue;
            if (!rangeRef.Contains(':')) continue;

            var parts = rangeRef.Split(':');
            var (startColName, startRow) = ParseCellReference(parts[0]);
            var (endColName, endRow) = ParseCellReference(parts[1]);
            var startColIdx = ColumnNameToIndex(startColName);
            var endColIdx = ColumnNameToIndex(endColName);

            // Row below? (cell row == endRow + 1, within column span).
            if (cellRow == endRow + 1 && cellColIdx >= startColIdx && cellColIdx <= endColIdx)
            {
                endRow += 1;
                var newRef = $"{startColName}{startRow}:{endColName}{endRow}";
                table.Reference = newRef;
                var af = table.GetFirstChild<AutoFilter>();
                if (af != null) af.Reference = newRef;
                table.Save();
                continue;
            }

            // Column right? (cell col == endCol + 1, within row span).
            if (cellColIdx == endColIdx + 1 && cellRow >= startRow && cellRow <= endRow)
            {
                endColIdx += 1;
                var newEndColName = IndexToColumnName(endColIdx);
                var newRef = $"{startColName}{startRow}:{newEndColName}{endRow}";
                table.Reference = newRef;
                var af = table.GetFirstChild<AutoFilter>();
                if (af != null) af.Reference = newRef;

                var tableColumns = table.GetFirstChild<TableColumns>();
                if (tableColumns != null)
                {
                    var existing = tableColumns.Elements<TableColumn>().ToList();
                    var nextId = existing.Count == 0
                        ? 1u
                        : existing.Max(tc => tc.Id?.Value ?? 0u) + 1u;
                    var used = new HashSet<string>(
                        existing.Select(tc => tc.Name?.Value ?? "")
                                .Where(n => !string.IsNullOrEmpty(n)),
                        StringComparer.OrdinalIgnoreCase);
                    var baseName = $"Column{existing.Count + 1}";
                    var colName = baseName;
                    int dedupeIdx = 2;
                    while (!used.Add(colName))
                        colName = $"{baseName}{dedupeIdx++}";
                    tableColumns.AppendChild(new TableColumn
                    {
                        Id = nextId,
                        Name = colName
                    });
                    tableColumns.Count = (uint)tableColumns.Elements<TableColumn>().Count();
                }

                table.Save();
            }
        }
    }

    // DATA-CORRUPTION(xlsx/table-header-name): Excel requires each
    // <tableColumn name="..."> to EXACTLY match the visible text of its
    // header-row cell; a mismatch makes the file unopenable
    // (0x800A03EC). When `add table` runs over empty header cells the
    // columns get auto names (Column1, Column2, ...). If the user later
    // overwrites a header cell with a real value, the worksheet says
    // "Product" while the table still says "Column1". This pass — called
    // after every cell value write — re-syncs the matching tableColumn's
    // name to the header cell's text. Mirrors the naming/escaping logic
    // in AddTable (GetCellDisplayValue + Column{n} fallback + uniqueness).
    private void MaybeSyncTableHeaderName(WorksheetPart worksheet, string cellRef)
    {
        var (cellCol, cellRow) = ParseCellReference(cellRef.ToUpperInvariant());
        var cellColIdx = ColumnNameToIndex(cellCol);

        foreach (var tdp in worksheet.TableDefinitionParts.ToList())
        {
            var table = tdp.Table;
            if (table == null) continue;
            // Tables with no header row (HeaderRowCount=0) have no column
            // names tied to cells — skip them.
            if (table.HeaderRowCount != null && table.HeaderRowCount.Value == 0) continue;
            if (table.Reference?.Value is not string rangeRef || !rangeRef.Contains(':')) continue;

            var parts = rangeRef.Split(':');
            var (startColName, startRow) = ParseCellReference(parts[0]);
            var (endColName, _) = ParseCellReference(parts[1]);
            var startColIdx = ColumnNameToIndex(startColName);
            var endColIdx = ColumnNameToIndex(endColName);

            // Header row is the first row of the table reference. Only act
            // when the written cell is on that row and within the span.
            if (cellRow != (int)startRow) continue;
            if (cellColIdx < startColIdx || cellColIdx > endColIdx) continue;

            var tableColumns = table.GetFirstChild<TableColumns>();
            if (tableColumns == null) continue;
            var cols = tableColumns.Elements<TableColumn>().ToList();
            int colOffset = cellColIdx - startColIdx; // 0-based position in table
            if (colOffset < 0 || colOffset >= cols.Count) continue;
            var targetCol = cols[colOffset];

            // Read the header cell's current display text.
            var sheetData = worksheet.Worksheet?.GetFirstChild<SheetData>();
            var hdrRow = sheetData?.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == startRow);
            var headerCell = hdrRow?.Elements<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
            var text = headerCell != null ? GetCellDisplayValue(headerCell) : null;

            // Excel rejects empty/duplicate column names. If the header was
            // cleared, keep the existing (prior/auto) name to stay valid —
            // never write an empty name.
            if (string.IsNullOrEmpty(text)) continue;

            // Already in sync — nothing to do.
            if (string.Equals(targetCol.Name?.Value, text, StringComparison.Ordinal)) continue;

            // Uniqueness: if another column already uses this name, leave the
            // current name to avoid producing a duplicate (pre-existing edge);
            // at minimum this never produces a header/column mismatch worse
            // than before.
            var used = new HashSet<string>(
                cols.Where(c => !ReferenceEquals(c, targetCol))
                    .Select(c => c.Name?.Value ?? "")
                    .Where(n => !string.IsNullOrEmpty(n)),
                StringComparer.OrdinalIgnoreCase);
            if (used.Contains(text)) continue;

            targetCol.Name = text;
            table.Save();
        }
    }
}
