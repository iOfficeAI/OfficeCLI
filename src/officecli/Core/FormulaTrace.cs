// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// One reference collected from a formula's token stream. Single cells carry
/// R1==R2 and C1==C2; an unbounded side (whole-column / whole-row forms like
/// A:A, 1:5) marks the open bound with 0 and is resolved against the sheet's
/// used range by the consumer. <see cref="Sheet"/> is null for a same-sheet
/// reference — resolve against the formula's own sheet.
/// </summary>
/// <remarks>
/// There is no table-formula AST: the evaluator's recursive-descent parser
/// consumes the token stream directly, and defined names are folded by the
/// tokenizer itself (simple bodies emit their literal ref token; formula
/// bodies are inlined). The token stream is therefore the single source of
/// truth for dependency tracing.
/// </remarks>
internal sealed record TraceRef(string? Sheet, int R1, int C1, int R2, int C2);

/// <summary>A cell address inside a trace walk. Sheet is always the canonical
/// workbook sheet name by the time the engine sees it.</summary>
internal readonly record struct TraceCell(string Sheet, string CellRef);

/// <summary>One BFS hop: the visited cell, and the cells it depends on
/// (precedents) or that depend on it (dependents). Nodes are display paths in
/// the get path grammar: same-sheet "/Sheet1/B2", cross-sheet "/Sheet2!B2".</summary>
internal sealed record TraceEdge(string From, List<string> To);

internal sealed record TraceOutput(string Root, List<TraceEdge> Edges, bool Truncated);

/// <summary>
/// Read-only view over a workbook for <see cref="FormulaTrace.Trace"/>.
/// Implemented by the Excel handler; every member must be side-effect free.
/// </summary>
internal interface IFormulaTraceSource
{
    /// <summary>Token references of the formula at <paramref name="cell"/>, or
    /// null when the cell has no formula or its text does not tokenize.</summary>
    List<TraceRef>? RefsOf(TraceCell cell);

    /// <summary>Populated bounds (max row, max col; 0 when the sheet is empty),
    /// used to resolve unbounded range sides.</summary>
    (int MaxRow, int MaxCol) UsedRange(string sheet);

    /// <summary>Cells whose formulas reference the given cell, in deterministic
    /// build order. Empty when nothing references it.</summary>
    List<TraceCell> DependentsOf(string sheet, int row, int col);
}

/// <summary>
/// Formula dependency tracing over the evaluator's token stream (B2 task /
/// 03 §9). Precedents walk the refs of each visited cell's formula; dependents
/// answer through the source's reverse index. Pure graph walk — no evaluation,
/// no writes. Cycles terminate via a visited set while still being REPORTED:
/// edges list every ref of a cell, so A1=B1 / B1=A1 shows both hops.
/// </summary>
internal static class FormulaTrace
{
    internal const int DefaultDepth = 1;
    internal const int MaxDepth = 16;
    /// <summary>Leaf cap per range reference: a SUM over 100k cells must not
    /// materialize 100k edge nodes. Row-major from the range's top-left; the
    /// overflow sets <see cref="TraceOutput.Truncated"/>.</summary>
    internal const int MaxLeafCellsPerRange = 5000;
    /// <summary>Global leaf budget across the whole trace output.</summary>
    internal const int MaxTotalLeaves = 20000;
    /// <summary>Global edge budget (wide dependent fan-out protection).</summary>
    internal const int MaxEdges = 5000;

    public static TraceOutput Trace(IFormulaTraceSource src, string rootSheet, string rootCell,
        bool dependents, int depth)
    {
        depth = Math.Clamp(depth, 1, MaxDepth);
        var edges = new List<TraceEdge>();
        var truncated = false;
        long totalLeaves = 0;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var rootDisplay = FormatCell(new TraceCell(rootSheet, rootCell), rootSheet);
        var queue = new Queue<(TraceCell Cell, string Display)>();
        queue.Enqueue((new TraceCell(rootSheet, rootCell), rootDisplay));
        visited.Add(Key(rootSheet, rootCell));

        for (var level = 0; level < depth && queue.Count > 0 && edges.Count < MaxEdges; level++)
        {
            var levelSize = queue.Count;
            while (levelSize-- > 0 && edges.Count < MaxEdges)
            {
                var (cell, display) = queue.Dequeue();

                var to = new List<string>();
                var targets = new List<TraceCell>();
                if (dependents)
                {
                    var (row, col) = ParseCellRef(cell.CellRef);
                    targets = src.DependentsOf(cell.Sheet, row, col);
                }
                else
                {
                    var refs = src.RefsOf(cell);
                    if (refs != null)
                        foreach (var rf in refs)
                        {
                            if (totalLeaves >= MaxTotalLeaves) { truncated = true; break; }
                            foreach (var leaf in ExpandRef(src, rf, cell.Sheet, ref truncated))
                            {
                                if (totalLeaves >= MaxTotalLeaves) { truncated = true; break; }
                                targets.Add(leaf);
                                totalLeaves++;
                            }
                        }
                }

                foreach (var t in targets)
                    AddUnique(to, FormatCell(t, cell.Sheet));
                if (to.Count == 0) continue;

                edges.Add(new TraceEdge(display, to));
                if (edges.Count >= MaxEdges) { truncated = true; break; }
                foreach (var t in targets)
                {
                    if (visited.Add(Key(t.Sheet, t.CellRef)))
                        queue.Enqueue((t, FormatCell(t, cell.Sheet)));
                }
            }
        }
        return new TraceOutput(rootDisplay, edges, truncated);
    }

    /// <summary>
    /// Resolve one reference into cell leaves. Unbounded sides clamp to the
    /// sheet's used range; a range wider than <see cref="MaxLeafCellsPerRange"/>
    /// truncates row-major from the top-left and flags the output.
    /// </summary>
    private static List<TraceCell> ExpandRef(IFormulaTraceSource src, TraceRef rf, string ownerSheet, ref bool truncated)
    {
        var sheet = rf.Sheet ?? ownerSheet;
        int r2, c2;
        if (rf.R2 == 0 || rf.C2 == 0)
        {
            var (maxRow, maxCol) = src.UsedRange(sheet);
            // An unbounded range on an empty sheet still covers its first cell
            // (SUM(A:A) depends on A1 the moment A1 exists), so clamp to 1, not 0.
            r2 = rf.R2 == 0 ? Math.Max(1, maxRow) : rf.R2;
            c2 = rf.C2 == 0 ? Math.Max(1, maxCol) : rf.C2;
        }
        else { r2 = rf.R2; c2 = rf.C2; }
        var (r1, c1) = (Math.Max(1, rf.R1), Math.Max(1, rf.C1));
        if (r2 < r1) r2 = r1;
        if (c2 < c1) c2 = c1;

        var leaves = new List<TraceCell>();
        var count = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
        var emitted = 0L;
        for (var r = r1; r <= r2; r++)
        {
            if (emitted >= MaxLeafCellsPerRange) { truncated = true; break; }
            for (var c = c1; c <= c2; c++)
            {
                if (emitted >= MaxLeafCellsPerRange) { truncated = true; break; }
                leaves.Add(new TraceCell(sheet, $"{ColToName(c)}{r}"));
                emitted++;
            }
        }
        if (count > MaxLeafCellsPerRange) truncated = true;
        return leaves;
    }

    private static bool AddUnique(List<string> list, string item)
    {
        if (list.Contains(item)) return false;
        list.Add(item);
        return true;
    }

    private static string Key(string sheet, string cell) => $"{sheet.ToUpperInvariant()}!{cell.ToUpperInvariant()}";

    /// <summary>Same-sheet nodes render under the formula owner's sheet
    /// ("/Sheet1/B2"); cross-sheet nodes render fully qualified ("/Sheet2!B2").</summary>
    private static string FormatCell(TraceCell cell, string ownerSheet)
        => string.Equals(cell.Sheet, ownerSheet, StringComparison.OrdinalIgnoreCase)
            ? $"/{cell.Sheet}/{cell.CellRef}"
            : $"/{cell.Sheet}!{cell.CellRef}";

    /// <summary>"B2" → (row 2, col 2). Caller guarantees the cell shape.</summary>
    internal static (int Row, int Col) ParseCellRef(string cellRef)
    {
        var i = 0;
        while (i < cellRef.Length && char.IsAsciiLetter(cellRef[i])) i++;
        var col = 0;
        foreach (var ch in cellRef[..i])
            col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        return (int.Parse(cellRef[i..]), col);
    }

    internal static string ColToName(int col)
    {
        var sb = new StringBuilder();
        while (col > 0)
        {
            var rem = (col - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            col = (col - 1) / 26;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Text-mode rendering: indented tree, children under their formula cell.
    /// A node reachable through multiple parents renders under each; a node on
    /// its own ancestor path is a circular reference and is marked once.
    /// </summary>
    public static string RenderTree(TraceOutput t, bool dependents)
    {
        var byFrom = new Dictionary<string, TraceEdge>(StringComparer.Ordinal);
        foreach (var e in t.Edges) byFrom[e.From] = e;
        var sb = new StringBuilder();
        sb.AppendLine($"trace ({(dependents ? "dependents" : "precedents")}) of {t.Root}:");
        Render(sb, t.Root, byFrom, new HashSet<string>(StringComparer.Ordinal), 0, dependents);
        if (t.Truncated) sb.AppendLine("(truncated — output exceeded leaf/edge caps)");
        return sb.ToString().TrimEnd();
    }

    private static void Render(StringBuilder sb, string node, Dictionary<string, TraceEdge> byFrom,
        HashSet<string> path, int depth, bool dependents)
    {
        sb.Append(' ', depth * 2 + 2);
        if (depth > 0) sb.Append(dependents ? "-> " : "<- ");
        sb.AppendLine(node);
        if (!path.Add(node))
        {
            sb.Append(' ', depth * 2 + 4);
            sb.AppendLine("(circular reference)");
            return;
        }
        if (byFrom.TryGetValue(node, out var edge) && depth < MaxDepth)
            foreach (var to in edge.To)
                Render(sb, to, byFrom, path, depth + 1, dependents);
        path.Remove(node);
    }
}

/// <summary>
/// Trace side of the evaluator partial: expose the token stream without
/// evaluating. Lives here so the token types (private to the evaluator)
/// never leak into the trace engine above.
/// </summary>
internal partial class FormulaEvaluator
{
    /// <summary>
    /// Collect the reference tokens of a formula WITHOUT evaluating it.
    /// Defined names are already folded by the tokenizer (simple bodies emit
    /// their literal ref; formula bodies are inlined), so this covers cell,
    /// range, sheet-qualified, and named references in one pass. Returns null
    /// when the formula does not tokenize — trace degrades to "no collectable
    /// references" rather than failing the whole walk.
    /// </summary>
    internal List<TraceRef>? TryCollectRefs(string formula)
    {
        try
        {
            var tokens = Tokenize(ModernFunctionQualifier.Unqualify(formula));
            var refs = new List<TraceRef>();
            foreach (var t in tokens)
            {
                switch (t.Type)
                {
                    case TT.CellRef:
                        if (TraceRefParser.TryCellRef(t.Value, out var r1, out var c1))
                            refs.Add(new TraceRef(null, r1, c1, r1, c1));
                        break;
                    case TT.SheetCellRef:
                        SplitSheetToken(t.Value, out var sheet1, out var ref1);
                        if (TraceRefParser.TryCellRef(ref1, out var sr, out var sc))
                            refs.Add(new TraceRef(sheet1, sr, sc, sr, sc));
                        break;
                    case TT.Range:
                        if (TraceRefParser.TryRange(t.Value, out var rng))
                            refs.Add(rng with { Sheet = null });
                        break;
                    case TT.SheetRange:
                        SplitSheetToken(t.Value, out var sheet2, out var ref2);
                        if (TraceRefParser.TryRange(ref2, out var rng2))
                            refs.Add(rng2 with { Sheet = sheet2 });
                        break;
                }
            }
            return refs;
        }
        catch (NameResolutionException) { return null; }
        catch (NotSupportedException) { return null; }
        catch (RegexMatchTimeoutException) { return null; }
    }

    /// <summary>"'My Sheet'!B2" / "Sheet2!B2" → sheet part before the LAST '!'
    /// (a quoted sheet name may itself contain '!'), ref part after.</summary>
    private static void SplitSheetToken(string value, out string sheet, out string refPart)
    {
        var bang = value.LastIndexOf('!');
        if (bang < 0) { sheet = ""; refPart = value; return; }
        sheet = value[..bang];
        refPart = value[(bang + 1)..];
    }
}

/// <summary>Parses the ref-shaped token values the evaluator's tokenizer emits.</summary>
internal static class TraceRefParser
{
    /// <summary>"B2" → row 2, col 2.</summary>
    public static bool TryCellRef(string s, out int row, out int col)
    {
        row = col = 0;
        if (string.IsNullOrEmpty(s)) return false;
        var i = 0;
        while (i < s.Length && char.IsAsciiLetter(s[i])) i++;
        if (i == 0 || i > 3 || i == s.Length) return false;
        if (!int.TryParse(s[i..], out row) || row < 1) return false;
        col = ColFromLetters(s[..i]);
        return col > 0;
    }

    /// <summary>
    /// "A1:C3" → a bounded ref; "A:C" → rows 1..0 (open bottom); "1:5" →
    /// cols 1..0 (open right). 0 in an end bound always means unbounded.
    /// </summary>
    public static bool TryRange(string s, out TraceRef r)
    {
        r = new TraceRef(null, 0, 0, 0, 0);
        var parts = s.Split(':');
        if (parts.Length != 2) return false;
        var (a, b) = (parts[0], parts[1]);
        if (a.Length == 0 || b.Length == 0) return false;
        if (char.IsAsciiDigit(a[0]) && char.IsAsciiDigit(b[0]))
        {
            if (!int.TryParse(a, out var ra) || !int.TryParse(b, out var rb) || ra < 1 || rb < 1) return false;
            if (rb < ra) (ra, rb) = (rb, ra);
            r = new TraceRef(null, ra, 1, rb, 0);
            return true;
        }
        var aLetters = AllLetters(a);
        var bLetters = AllLetters(b);
        if (aLetters && bLetters)
        {
            var ca = ColFromLetters(a);
            var cb = ColFromLetters(b);
            if (ca <= 0 || cb <= 0) return false;
            if (cb < ca) (ca, cb) = (cb, ca);
            r = new TraceRef(null, 1, ca, 0, cb);
            return true;
        }
        if (!aLetters && !bLetters &&
            TryCellRef(a, out var r1, out var c1) && TryCellRef(b, out var r2, out var c2))
        {
            if (r2 < r1) (r1, r2) = (r2, r1);
            if (c2 < c1) (c1, c2) = (c2, c1);
            r = new TraceRef(null, r1, c1, r2, c2);
            return true;
        }
        return false;
    }

    private static bool AllLetters(string s)
    {
        foreach (var ch in s)
            if (!char.IsAsciiLetter(ch)) return false;
        return s.Length > 0;
    }

    private static int ColFromLetters(string letters)
    {
        var col = 0;
        foreach (var ch in letters)
        {
            var up = char.ToUpperInvariant(ch);
            if (up is < 'A' or > 'Z') return 0;
            col = col * 26 + (up - 'A' + 1);
        }
        return col;
    }
}
