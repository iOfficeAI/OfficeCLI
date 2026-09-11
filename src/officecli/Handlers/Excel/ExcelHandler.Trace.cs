// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ===== Formula dependency trace (get --prop trace=precedents|dependents) =====
    //
    // Pure read-only query on top of the formula tokenizer. Precedents walk
    // the token stream of each visited cell's formula (FormulaEvaluator.
    // TryCollectRefs — defined names arrive pre-folded, so cell, range,
    // sheet-qualified and named refs all come out of one pass). Dependents
    // answer from a lazily-built in-process reverse index: every formula cell
    // in the workbook scanned once, its refs recorded either as exact single
    // cells or as clamped range entries, then cached until the document
    // changes. Invalidation rides the InvalidateRowIndex hook the row-index
    // cache already uses, so every mutation site that drops the row index
    // drops this one too; sheet rename/delete add explicit drops (the index
    // is keyed by sheet NAME, which those operations rewrite).

    /// <summary>
    /// Trace the formula dependency graph around <paramref name="path"/> and
    /// (optionally) attach the result to the cell node's Format as the
    /// "trace" key. JSON output carries it inside the envelope; text output
    /// renders the indented tree from the returned <see cref="Core.TraceOutput"/> —
    /// attach only in JSON mode so the one-line node record stays readable.
    /// Throws CliException with code invalid_argument (malformed path/depth),
    /// not_found (root cell missing), unsupported_element (precedents on a
    /// constant cell), or unsupported_type (non-Excel handler).
    /// </summary>
    internal Core.TraceOutput AttachTrace(Core.DocumentNode node, string path, string mode, int traceDepth, bool attachToFormat)
    {
        if (!string.Equals(mode, "precedents", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(mode, "dependents", StringComparison.OrdinalIgnoreCase))
            throw new Core.CliException($"Unknown trace mode: '{mode}'.")
            {
                Code = "invalid_argument",
                ValidValues = ["precedents", "dependents"],
                Suggestion = "trace=precedents walks what the cell depends on; trace=dependents walks who depends on it"
            };
        if (traceDepth < 1 || traceDepth > Core.FormulaTrace.MaxDepth)
            throw new Core.CliException($"trace depth must be an integer between 1 and {Core.FormulaTrace.MaxDepth} (got {traceDepth}).")
            {
                Code = "invalid_argument",
                Suggestion = "--prop depth=3 walks three hops of the dependency chain"
            };
        var segments = path.TrimStart('/').Split('/', 2);
        var cellRef = segments.Length == 2 ? segments[1].ToUpperInvariant() : "";
        if (cellRef.Contains('/') || !Regex.IsMatch(cellRef, @"^[A-Z]{1,3}\d+$"))
            throw new Core.CliException($"trace requires a single cell path like /Sheet1/F2 (got '{path}').")
            {
                Code = "invalid_argument",
                Suggestion = "get <file> '/Sheet1/F2' --prop trace=precedents [--prop depth=N]"
            };
        var sheetName = segments[0];
        var wsPart = FindWorksheet(sheetName);
        var sheetData = wsPart != null ? GetSheet(wsPart).GetFirstChild<SheetData>() : null;
        var cell = sheetData != null ? FindCell(sheetData, cellRef) : null;
        if (cell == null)
            throw new Core.CliException($"Cell /{sheetName}/{cellRef} not found — nothing to trace.")
            {
                Code = "not_found",
                Suggestion = $"set it first, or get /{sheetName}/{cellRef} to inspect the current state"
            };
        if (mode == "precedents" && string.IsNullOrEmpty(cell.CellFormula?.Text))
            throw new Core.CliException($"/{sheetName}/{cellRef} holds a constant — no formula to trace precedents from.")
            {
                Code = "unsupported_element",
                Suggestion = "trace=precedents needs a formula cell; trace=dependents works on any cell (it answers who references it), or get the cell --json and check the 'formula' key"
            };
        var source = new TraceSource(this);
        var output = Core.FormulaTrace.Trace(source, sheetName, cellRef,
            dependents: string.Equals(mode, "dependents", StringComparison.OrdinalIgnoreCase), traceDepth);
        if (attachToFormat) node.Format["trace"] = TraceToFormat(output);
        return output;
    }

    /// <summary>Envelope shape for the trace payload: only shapes the source-
    /// generated Format serializer has metadata for (Dictionary values resolve
    /// polymorphically — a List&lt;object?&gt; here fails at serialization time).</summary>
    internal static Dictionary<string, object?> TraceToFormat(Core.TraceOutput output)
    {
        var edgeList = new List<Dictionary<string, object?>>(output.Edges.Count);
        foreach (var e in output.Edges)
            edgeList.Add(new Dictionary<string, object?> { ["from"] = e.From, ["to"] = e.To });
        return new Dictionary<string, object?>
        {
            ["root"] = output.Root,
            ["edges"] = edgeList,
            ["truncated"] = output.Truncated,
        };
    }

    private sealed class DependentsIndex
    {
        /// <summary>"SHEET|CELL" (upper) → source keys "SHEET|CELL", in build order.</summary>
        public readonly Dictionary<string, List<string>> Exact = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>Range refs that cover >1 cell (bounds already resolved:
        /// open sides clamped to the used range at build time).</summary>
        public readonly List<(string Sheet, int R1, int C1, int R2, int C2, string Src)> Ranges = new();
        /// <summary>UPPER sheet name → canonical name (display uses real casing).</summary>
        public readonly Dictionary<string, string> CanonicalSheets = new(StringComparer.OrdinalIgnoreCase);
    }

    private DependentsIndex? _dependentsIndex;

    /// <summary>Drop the reverse index so the next dependents query rebuilds it
    /// from the current document. Called from InvalidateRowIndex (every cell /
    /// row / import mutation) and explicitly from sheet rename and remove.</summary>
    internal void InvalidateDependentsIndex() => _dependentsIndex = null;

    private DependentsIndex EnsureDependentsIndex()
    {
        if (_dependentsIndex != null) return _dependentsIndex;
        var idx = new DependentsIndex();
        var usedRanges = new Dictionary<string, (int MaxRow, int MaxCol)>(StringComparer.OrdinalIgnoreCase);
        foreach (var (sheetName, wsPart) in GetWorksheets())
        {
            idx.CanonicalSheets[sheetName.ToUpperInvariant()] = sheetName;
            var sheetData = GetSheet(wsPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;
            var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var formula = cell.CellFormula?.Text;
                    if (string.IsNullOrEmpty(formula) || cell.CellReference?.Value is not string cref) continue;
                    var src = $"{sheetName.ToUpperInvariant()}|{cref.ToUpperInvariant()}";
                    var refs = evaluator.TryCollectRefs(formula);
                    if (refs == null) continue;
                    foreach (var rf in refs)
                    {
                        var refSheet = CanonicalSheetName(rf.Sheet ?? sheetName);
                        var refUpper = refSheet.ToUpperInvariant();
                        idx.CanonicalSheets.TryAdd(refUpper, refSheet);
                        if (FindWorksheet(refSheet) == null) continue; // dangling sheet ref — nothing to attach to
                        int r2, c2;
                        if (rf.R2 == 0 || rf.C2 == 0)
                        {
                            var (maxRow, maxCol) = UsedRangeOf(usedRanges, refSheet);
                            r2 = rf.R2 == 0 ? Math.Max(1, maxRow) : rf.R2;
                            c2 = rf.C2 == 0 ? Math.Max(1, maxCol) : rf.C2;
                        }
                        else { r2 = rf.R2; c2 = rf.C2; }
                        var (r1, c1) = (Math.Max(1, rf.R1), Math.Max(1, rf.C1));
                        if (r2 < r1) r2 = r1;
                        if (c2 < c1) c2 = c1;
                        if (r1 == r2 && c1 == c2)
                        {
                            var key = $"{refUpper}|{Core.FormulaTrace.ColToName(c1)}{r1}";
                            if (!idx.Exact.TryGetValue(key, out var list))
                                idx.Exact[key] = list = new List<string>();
                            if (!list.Contains(src)) list.Add(src);
                        }
                        else
                        {
                            idx.Ranges.Add((refUpper, r1, c1, r2, c2, src));
                        }
                    }
                }
            }
        }
        _dependentsIndex = idx;
        return idx;
    }

    private string CanonicalSheetName(string name)
    {
        foreach (var (sheetName, _) in GetWorksheets())
            if (string.Equals(sheetName, name, StringComparison.OrdinalIgnoreCase))
                return sheetName;
        return name;
    }

    private (int MaxRow, int MaxCol) UsedRangeOf(Dictionary<string, (int, int)> cache, string sheetName)
    {
        if (cache.TryGetValue(sheetName, out var cached)) return cached;
        var (maxRow, maxCol) = (0, 0);
        if (FindWorksheet(sheetName) is { } wsPart &&
            GetSheet(wsPart).GetFirstChild<SheetData>() is { } sheetData)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                if (row.RowIndex?.Value is { } ri && ri > maxRow) maxRow = (int)ri;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value is not string cref) continue;
                    try
                    {
                        var col = ColumnNameToIndex(ParseCellReference(cref).Column);
                        if (col > maxCol) maxCol = col;
                    }
                    catch { /* malformed ref — skip the cell for bound purposes */ }
                }
            }
        }
        cache[sheetName] = (maxRow, maxCol);
        return (maxRow, maxCol);
    }

    /// <summary>Adapter the trace engine walks. Read-only: evaluators are
    /// constructed per sheet on demand and never asked to evaluate.</summary>
    private sealed class TraceSource(ExcelHandler owner) : Core.IFormulaTraceSource
    {
        private readonly Dictionary<string, SheetData?> _sheetData = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Core.FormulaEvaluator?> _evaluators = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, (int MaxRow, int MaxCol)> _usedRanges = new(StringComparer.OrdinalIgnoreCase);

        public List<Core.TraceRef>? RefsOf(Core.TraceCell cell)
        {
            var sheetData = SheetDataOf(cell.Sheet);
            if (sheetData == null) return null;
            var found = FindCell(sheetData, cell.CellRef);
            var formula = found?.CellFormula?.Text;
            if (string.IsNullOrEmpty(formula)) return null;
            var evaluator = EvaluatorOf(cell.Sheet, sheetData);
            if (evaluator == null) return null;
            var refs = evaluator.TryCollectRefs(formula);
            if (refs == null) return null;
            // Resolve raw token sheet names to canonical workbook names so the
            // graph keys line up across cells; owner-sheet refs stay null.
            for (var i = 0; i < refs.Count; i++)
            {
                if (refs[i].Sheet == null) continue;
                var canonical = owner.CanonicalSheetName(refs[i].Sheet!);
                refs[i] = refs[i] with { Sheet = canonical };
            }
            return refs;
        }

        public (int MaxRow, int MaxCol) UsedRange(string sheet)
        {
            if (_usedRanges.TryGetValue(sheet, out var cached)) return cached;
            var (maxRow, maxCol) = owner.UsedRangeOf(new(), sheet);
            _usedRanges[sheet] = (maxRow, maxCol);
            return (maxRow, maxCol);
        }

        public List<Core.TraceCell> DependentsOf(string sheet, int row, int col)
        {
            var idx = owner.EnsureDependentsIndex();
            var upper = sheet.ToUpperInvariant();
            var key = $"{upper}|{Core.FormulaTrace.ColToName(col)}{row}";
            var sources = new List<string>();
            if (idx.Exact.TryGetValue(key, out var exact))
                sources.AddRange(exact);
            foreach (var r in idx.Ranges)
            {
                if (!string.Equals(r.Sheet, upper, StringComparison.OrdinalIgnoreCase)) continue;
                if (row < r.R1 || row > r.R2 || col < r.C1 || col > r.C2) continue;
                if (!sources.Contains(r.Src)) sources.Add(r.Src);
            }
            var result = new List<Core.TraceCell>();
            foreach (var src in sources)
            {
                var pipe = src.IndexOf('|');
                if (pipe <= 0) continue;
                var srcUpper = src[..pipe];
                if (!idx.CanonicalSheets.TryGetValue(srcUpper, out var canonical)) continue;
                result.Add(new Core.TraceCell(canonical, src[(pipe + 1)..]));
            }
            return result;
        }

        private SheetData? SheetDataOf(string sheet)
        {
            if (_sheetData.TryGetValue(sheet, out var cached)) return cached;
            SheetData? data = null;
            if (owner.FindWorksheet(sheet) is { } wsPart)
                data = GetSheet(wsPart).GetFirstChild<SheetData>();
            _sheetData[sheet] = data;
            return data;
        }

        private Core.FormulaEvaluator? EvaluatorOf(string sheet, SheetData sheetData)
        {
            if (_evaluators.TryGetValue(sheet, out var cached)) return cached;
            var ev = new Core.FormulaEvaluator(sheetData, owner._doc.WorkbookPart);
            _evaluators[sheet] = ev;
            return ev;
        }
    }
}
