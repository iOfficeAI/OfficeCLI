// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Result of a formula evaluation. Can be numeric, string, boolean, or error.
/// </summary>
internal record FormulaResult
{
    public double? NumericValue { get; init; }
    public string? StringValue { get; init; }
    public bool? BoolValue { get; init; }
    public string? ErrorValue { get; init; }
    public double[]? ArrayValue { get; init; }
    public RangeData? RangeValue { get; init; }
    // A LAMBDA(...) value: captured parameter names + unevaluated body tokens.
    // Carried as object to keep this record free of evaluator-internal types.
    public object? LambdaValue { get; init; }

    public bool IsLambda => LambdaValue != null;
    public bool IsNumeric => NumericValue.HasValue;
    public bool IsString => StringValue != null;
    public bool IsBool => BoolValue.HasValue;
    public bool IsError => ErrorValue != null;
    public bool IsArray => ArrayValue != null;
    public bool IsRange => RangeValue != null;
    // Blank carries no value of any kind: arithmetic coerces to 0, string
    // concat coerces to "". Used for empty/missing cells reached through
    // OFFSET / INDIRECT / direct ref so `=OFFSET(A1,5,0)&"x"` matches Excel's
    // "x" rather than emitting "0x".
    public bool IsBlank => !IsNumeric && !IsString && !IsBool && !IsError && !IsArray && !IsRange;

    public static FormulaResult Number(double v) => new() { NumericValue = v };
    public static FormulaResult Str(string v) => new() { StringValue = v };
    public static FormulaResult Bool(bool v) => new() { BoolValue = v };
    public static FormulaResult Error(string v) => new() { ErrorValue = v };
    public static FormulaResult Array(double[] v) => new() { ArrayValue = v };
    public static FormulaResult Area(RangeData v) => new() { RangeValue = v };
    public static FormulaResult Blank() => new();

    // Excel coerces numeric-looking text in arithmetic / scalar contexts:
    // ="1"*"4186"*0.03 → 125.58. Cells flagged t="str" (e.g. set under
    // numberformat="@") flow in here as IsString — without TryParse they'd
    // silently become 0 and pollute cachedValue. SUM/AVERAGE go through
    // RangeData.ToDoubleArray which gates on IsNumeric and is unaffected.
    public double AsNumber()
    {
        if (IsRange) return FirstCell()?.AsNumber() ?? 0;
        if (NumericValue.HasValue) return NumericValue.Value;
        if (BoolValue.HasValue) return BoolValue.Value ? 1 : 0;
        if (IsString && double.TryParse(StringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var s)) return s;
        return 0;
    }
    public string AsString() => IsRange ? (FirstCell()?.AsString() ?? "") :
        StringValue ?? NumericValue?.ToString(CultureInfo.InvariantCulture)
        ?? (BoolValue.HasValue ? (BoolValue.Value ? "TRUE" : "FALSE") : ErrorValue ?? "");

    private FormulaResult? FirstCell() =>
        RangeValue is { Rows: > 0, Cols: > 0 } rd ? rd.Cells[0, 0] : null;

    public string ToCellValueText()
    {
        // R3 BUG-5: errors must surface as their sentinel ("#REF!", "#VALUE!",
        // …) — not as the empty StringValue fallback which suppresses the
        // <v> write on the cell and leaves only the formula text. The Set
        // path also gates on IsError separately and writes t="e", so this
        // branch is the safety net for any caller (HtmlPreview, view) that
        // formats the value text directly.
        if (IsError) return ErrorValue!;
        // A blank (empty-cell ref) is 0 when it lands directly in a cell — Excel
        // displays `=A6` (A6 empty) as 0. The "" coercion is only the right
        // answer when blank is the right operand of a string concat (handled in
        // ParseConcat).
        if (IsBlank) return "0";
        // An Area placed into a single cell collapses to its top-left.
        // Excel does implicit-intersect; top-left is the simplest deterministic
        // choice (and matches FirstCell()).
        if (IsRange) return FirstCell()?.ToCellValueText() ?? "";
        if (NumericValue.HasValue)
        {
            var v = NumericValue.Value;
            // IEEE-754 ±Infinity / NaN have no OOXML representation; emitting
            // "Infinity" / "-Infinity" / "NaN" into <x:v> produces a file Excel
            // refuses to open. Surface them as the Excel error string so the
            // calling cell switches to t="e" (#NUM!) — matches what Excel does
            // for LOG(0), SQRT(-1), 0/0, etc.
            if (double.IsNaN(v) || double.IsInfinity(v))
                return "#NUM!";
            // Round to 15 significant digits to avoid floating point artifacts (e.g. 25300000.000000004)
            if (v != 0)
            {
                var digits = 15 - (int)Math.Floor(Math.Log10(Math.Abs(v))) - 1;
                if (digits is >= 0 and <= 15)
                    v = Math.Round(v, digits);
            }
            return v.ToString(CultureInfo.InvariantCulture);
        }
        return BoolValue.HasValue ? (BoolValue.Value ? "1" : "0") : StringValue ?? "";
    }
}

/// <summary>
/// Status returned by <see cref="FormulaEvaluator.EvaluateForReport"/>.
/// Distinguishes "evaluator gave up" (NotEvaluated) from "evaluator produced
/// an Excel-style error" (Error) — agents need both signals separately.
/// </summary>
internal enum EvalReportStatus { Evaluated, Error, NotEvaluated }

/// <summary>Single-source report from EvaluateForReport — feeds the
/// <c>evaluated</c> cell field, the <c>view text</c> sentinel, and the
/// <c>view issues</c> formula_not_evaluated warning from one decision.</summary>
internal sealed record EvalReport(EvalReportStatus Status, FormulaResult? Result);

/// <summary>
/// 2D range data for lookup functions (VLOOKUP, HLOOKUP, INDEX).
/// </summary>
internal class RangeData
{
    public FormulaResult?[,] Cells { get; }
    public int Rows { get; }
    public int Cols { get; }
    // Origin row/col of the top-left cell when this RangeData was produced by a
    // resolved reference (1-based). 0 means "not from a reference" (e.g. literal
    // array). Used by ROW() / COLUMN() / ADDRESS() so they can answer the
    // reference's origin even when given an OFFSET-returned Area instead of a
    // raw cell-ref string.
    public int BaseRow { get; init; }
    public int BaseCol { get; init; }
    // Sheet name when the area was produced by a cross-sheet reference (e.g.
    // OFFSET(Sheet2!A1, 0, 0)). Null/empty means same-sheet. Used by EvalOffset
    // when reconstructing a RefArg from an Area to preserve the origin sheet.
    public string? BaseSheet { get; init; }

    public RangeData(FormulaResult?[,] cells) { Cells = cells; Rows = cells.GetLength(0); Cols = cells.GetLength(1); }

    public double[] ToDoubleArray()
    {
        var values = new List<double>();
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Cols; c++)
            {
                var cell = Cells[r, c];
                if (cell?.IsNumeric == true) values.Add(cell.NumericValue!.Value);
                else if (cell?.IsBool == true) values.Add(cell.BoolValue!.Value ? 1 : 0);
            }
        return values.ToArray();
    }

    /// <summary>Flatten all cells into a flat list (preserving nulls for ISERROR etc.)</summary>
    public FormulaResult?[] ToFlatResults()
    {
        var results = new FormulaResult?[Rows * Cols];
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Cols; c++)
                results[r * Cols + c] = Cells[r, c];
        return results;
    }

    /// <summary>Returns the first error found in the range, or null if none.</summary>
    public FormulaResult? FirstError()
    {
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Cols; c++)
                if (Cells[r, c]?.IsError == true) return Cells[r, c];
        return null;
    }
}

/// <summary>
/// Excel formula evaluator supporting 350+ functions.
/// Split across partial class files:
///   FormulaEvaluator.cs          — core: tokenizer, parser, cell resolution
///   FormulaEvaluator.Functions.cs — function dispatch + implementations
///   FormulaEvaluator.Helpers.cs   — math utilities, comparison helpers
/// </summary>
/// <summary>
/// State shared by every <see cref="FormulaEvaluator"/> instance participating in
/// one evaluation session (a root evaluator plus the per-sheet children it spawns
/// for cross-sheet references). Without it, each cross-sheet cell dereference
/// created a fresh evaluator whose <c>_cellIndex</c> re-scanned the whole target
/// sheet, and every formula cell reached through a reference was re-evaluated from
/// scratch — O(n²)+ blowup on workbooks whose formulas reference other formula
/// cells (issue: resident pegs CPU and flush hangs on formula-heavy xlsx).
/// Callers that batch many evaluations (the save-time cache sweep) can pass one
/// session across per-sheet evaluators so memoized results survive sheet hops.
/// </summary>
internal sealed class FormulaEvalSession
{
    internal readonly HashSet<string> Visiting = new(StringComparer.OrdinalIgnoreCase);
    internal readonly Dictionary<string, FormulaEvaluator> SheetEvaluators = new(StringComparer.OrdinalIgnoreCase);
    internal readonly Dictionary<string, SheetData?> SheetDataByName = new(StringComparer.OrdinalIgnoreCase);
    // Memoized result per formula cell ("Sheet!A1" → result). Only completed,
    // cycle-free evaluations are stored; the seed-0 value a circular chain
    // returns depends on the entry point and must never be reused.
    internal readonly Dictionary<string, FormulaResult> CellMemo = new(StringComparer.OrdinalIgnoreCase);
    // Populated-extent cache for whole-column/row clamping ("A:A" → used rows).
    // Keyed by SheetData reference; evaluation never adds/removes rows or cells,
    // so the extent is stable for the session's lifetime.
    internal readonly Dictionary<SheetData, (int Min, int Max)> RowExtentBySheet = new();
    internal readonly Dictionary<SheetData, (int Min, int Max)> ColExtentBySheet = new();
    // Materialized-range cache ("Sheet|col,row,w,h" → RangeData). Many formulas
    // scan the same criteria/data columns (SUMIFS/COUNTIF over PLN!$F:$F etc.);
    // materializing the rect once per session instead of once per formula is the
    // difference between minutes and seconds on formula-heavy workbooks. Safe for
    // the same reason CellMemo is: cell results are stable within a session.
    internal readonly Dictionary<string, RangeData> RangeMemo = new(StringComparer.OrdinalIgnoreCase);
    internal int CircularHits;
    internal int CrossSheetDepth;
}

internal partial class FormulaEvaluator
{
    private readonly SheetData _sheetData;
    private readonly WorkbookPart? _workbookPart;
    // 1-based position of the cell currently being evaluated, for argument-less
    // ROW()/COLUMN(). 0 means the caller did not supply it.
    private int _ctxRow, _ctxCol;
    private readonly FormulaEvalSession _session;
    private HashSet<string> _visiting => _session.Visiting;
    private readonly HashSet<string> _expandingNames = new(StringComparer.OrdinalIgnoreCase);
    // LET / LAMBDA variable bindings (innermost scope). Names are case-insensitive.
    private readonly Dictionary<string, FormulaResult> _bindings = new(StringComparer.OrdinalIgnoreCase);

    // A LAMBDA value: parameter names + the body's token stream, re-evaluated per call.
    private sealed record Lambda(List<string> Parameters, List<Token> Body);
    private readonly int _depth;
    private readonly string _sheetKey; // used to qualify cell refs for circular detection

    // Same-sheet recursion guard. A long non-circular chain (B[N]=B[N-1]+A[N])
    // recurses ResolveCellResult→EvaluateFormula once per link; deep enough it
    // overflows the .NET stack, and since StackOverflowException is uncatchable
    // it kills the whole process — a fatal DoS for the resident server. A fixed
    // frame-count cap can't fully close this: complex nested formulas (e.g.
    // IF(SUM(...),VLOOKUP(...),...)) burn many more frames per link and would
    // overflow well below any simple-chain cap. So the PRIMARY guard is
    // RuntimeHelpers.TryEnsureSufficientExecutionStack() (the standard .NET
    // recursive-SOE defense), which adapts to the ACTUAL stack each formula
    // consumes — simple deep chains keep evaluating (no regression), complex
    // ones bail only when the stack is genuinely near the limit. MaxSameSheetDepth
    // is a high backstop (1000) for the pathological case where the probe
    // misjudges; it should rarely pre-empt a legitimate chain. _visiting handles
    // circular refs separately. Over either limit → visible #NUM! (propagates up
    // the arithmetic chain), never a silent 0 nor a crash.
    private const int MaxSameSheetDepth = 1000;
    private int _sameSheetDepth;

    // CONSISTENCY(dos-hardening): the expression parser is recursive-descent;
    // every "(" re-enters ParseConcat, so a single formula like =(((…1…)))
    // nested tens of thousands deep overflows the stack with an UNCATCHABLE
    // StackOverflowException (process abort), independent of the cross-cell
    // _sameSheetDepth guard above. Bound the per-formula parse recursion with a
    // hard cap plus a runtime stack probe; over either, surface a visible
    // #NUM! instead of crashing.
    private int _parseDepth;

    // Number-format engine hook. TEXT(value, format) must apply Excel format
    // codes (incl. date/time/percent/currency) identically to the cell renderer.
    // That engine (ApplyNumberFormat) lives in the Handlers layer, which Core
    // must not reference; ExcelHandler wires this delegate up on init so TEXT
    // routes through the same formatter. Null fallback = numeric-only path.
    internal static Func<double, string, string>? NumberFormatProvider;
    private Dictionary<string, Cell>? _cellIndex;
    private Dictionary<string, string>? _definedNames;

    /// <summary>Thrown when a defined name cannot be resolved — either it
    /// recursively references itself or its body fails to tokenize. Both
    /// surface to the user as <c>#NAME?</c>.</summary>
    private sealed class NameResolutionException : Exception
    {
        public NameResolutionException(string name) : base(name) { }
    }

    public FormulaEvaluator(SheetData sheetData, WorkbookPart? workbookPart = null)
        : this(sheetData, workbookPart, new FormulaEvalSession(), 0, "") { }

    /// <summary>
    /// Root evaluator bound to a caller-owned session, so several root evaluators
    /// (one per sheet in a sweep) share memoized cell results and per-sheet child
    /// evaluators. <paramref name="sheetKey"/> must be the real sheet name when a
    /// session is shared — it namespaces the memo/circular keys ("REP!A1").
    /// </summary>
    internal FormulaEvaluator(SheetData sheetData, WorkbookPart? workbookPart, FormulaEvalSession session, string sheetKey)
        : this(sheetData, workbookPart, session, 0, sheetKey)
    {
        if (!string.IsNullOrEmpty(sheetKey))
            session.SheetEvaluators[sheetKey] = this;
    }

    private FormulaEvaluator(SheetData sheetData, WorkbookPart? workbookPart, FormulaEvalSession session, int depth, string sheetKey)
    {
        _sheetData = sheetData;
        _workbookPart = workbookPart;
        _session = session;
        _depth = depth;
        _sheetKey = sheetKey;
    }

    public double? TryEvaluate(string formula)
    {
        var result = TryEvaluateFull(formula);
        return result?.NumericValue ?? (result?.BoolValue == true ? 1 : result?.BoolValue == false ? 0 : null);
    }

    public FormulaResult? TryEvaluateFull(string formula)
    {
        try
        {
            if (_depth == 0) { _visiting.Clear(); _expandingNames.Clear(); }
            // Accept both qualified (`_xlfn.SEQUENCE`) and bare (`SEQUENCE`)
            // forms. Stored XML uses the qualified form post-R11-2; user code
            // and tests still pass the canonical name.
            return EvaluateFormula(ModernFunctionQualifier.Unqualify(formula));
        }
        catch (NameResolutionException) { return FormulaResult.Error("#NAME?"); }
        catch { return null; }
    }

    /// <summary>
    /// Single-source report wrapper used by `view text` sentinel, `view issues`
    /// (formula_not_evaluated), and `get` (Format["evaluated"]). Routes all
    /// three signals through one decision so they cannot drift apart as the
    /// evaluator's coverage grows.
    /// </summary>
    internal EvalReport EvaluateForReport(string formula, string? cellRef = null)
    {
        SetCellContext(cellRef);
        var r = TryEvaluateFull(formula);
        if (r == null) return new EvalReport(EvalReportStatus.NotEvaluated, null);
        if (r.IsError) return new EvalReport(EvalReportStatus.Error, r);
        return new EvalReport(EvalReportStatus.Evaluated, r);
    }

    // Record the evaluating cell's 1-based row/column (from an A1 ref) so
    // argument-less ROW()/COLUMN() can answer. A null/unparsable ref clears it.
    private void SetCellContext(string? cellRef)
    {
        _ctxRow = 0; _ctxCol = 0;
        if (string.IsNullOrEmpty(cellRef)) return;
        var m = System.Text.RegularExpressions.Regex.Match(cellRef, @"^\$?([A-Za-z]{1,3})\$?(\d+)$");
        if (!m.Success) return;
        int col = 0;
        foreach (var ch in m.Groups[1].Value.ToUpperInvariant()) col = col * 26 + (ch - 'A' + 1);
        _ctxCol = col;
        _ctxRow = int.Parse(m.Groups[2].Value);
    }

    private FormulaResult? EvaluateFormula(string formula)
    {
        var tokens = Tokenize(formula);
        var pos = 0;
        var result = ParseExpression(tokens, ref pos);
        if (pos != tokens.Count) return null;
        // Top-level Array/Range collapse to scalar via implicit intersect
        // (Excel's pre-dynamic-array behavior). The first element is returned
        // so a cell holding `=B1:B3*1` shows the first row's product.
        if (result?.IsArray == true) return result.ArrayValue!.Length > 0 ? FormulaResult.Number(result.ArrayValue[0]) : FormulaResult.Number(0);
        if (result?.IsRange == true) { var rd = result.RangeValue!; return rd.Rows > 0 && rd.Cols > 0 ? rd.Cells[0, 0] ?? FormulaResult.Number(0) : FormulaResult.Number(0); }
        return result;
    }

    // ==================== Tokenizer ====================

    private enum TT { Number, String, CellRef, Range, Op, LParen, RParen, Comma, Func, Bool, Compare, SheetCellRef, SheetRange, ArrayLit, Error, Name }
    private record Token(TT Type, string Value);

    private Dictionary<string, string> GetDefinedNames()
    {
        if (_definedNames != null) return _definedNames;
        _definedNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var dns = _workbookPart?.Workbook?.Descendants<DefinedName>();
        if (dns != null)
        {
            foreach (var dn in dns)
            {
                var name = dn.Name?.Value;
                var value = dn.Text;
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
                    _definedNames[name] = value;
            }
        }
        return _definedNames;
    }

    private List<Token> Tokenize(string formula)
    {
        var tokens = new List<Token>();
        var i = 0;
        formula = formula.Trim();

        while (i < formula.Length)
        {
            var ch = formula[i];
            if (char.IsWhiteSpace(ch)) { i++; continue; }

            if (ch is '>' or '<' or '=')
            {
                if (ch == '=' && i == 0) { i++; continue; }
                if (i + 1 < formula.Length && formula[i + 1] is '=' or '>')
                { tokens.Add(new Token(TT.Compare, formula.Substring(i, 2))); i += 2; }
                else { tokens.Add(new Token(TT.Compare, ch.ToString())); i++; }
                continue;
            }

            if (ch is '+' or '-' or '*' or '/' or '^' or '%')
            {
                if ((ch is '-' or '+') && (tokens.Count == 0 ||
                    tokens[^1].Type is TT.Op or TT.LParen or TT.Comma or TT.Compare))
                { var ns = ParseNumber(formula, ref i); if (ns != null) { tokens.Add(new Token(TT.Number, ns)); continue; } }
                if (ch == '%') { tokens.Add(new Token(TT.Op, "%")); i++; continue; }
                tokens.Add(new Token(TT.Op, ch.ToString())); i++; continue;
            }

            if (ch == '(') { tokens.Add(new Token(TT.LParen, "(")); i++; continue; }
            if (ch == ')') { tokens.Add(new Token(TT.RParen, ")")); i++; continue; }
            if (ch == ',') { tokens.Add(new Token(TT.Comma, ",")); i++; continue; }
            if (ch == '&') { tokens.Add(new Token(TT.Op, "&")); i++; continue; }

            // Error-constant literal embedded in a formula (e.g. IFERROR(#N/A, x)).
            // Recognized here so the whole formula still tokenizes rather than
            // failing on the '#'.
            if (ch == '#')
            {
                var lit = MatchErrorLiteral(formula, i);
                if (lit != null) { tokens.Add(new Token(TT.Error, lit)); i += lit.Length; continue; }
            }

            // Array constant literal: {1,2,3} (row) or {1;2;3} (column) or
            // {1,2;3,4} (matrix). Per ECMA-376 §18.17.7.282 (array-constant),
            // comma separates columns, semicolon separates rows. Cells may be
            // numbers, quoted strings, or TRUE/FALSE. Nested {} is not allowed.
            if (ch == '{')
            {
                var start = i + 1;
                var end = formula.IndexOf('}', start);
                if (end < 0) throw new NotSupportedException("Unclosed { in array constant");
                tokens.Add(new Token(TT.ArrayLit, formula[start..end]));
                i = end + 1;
                continue;
            }

            if (ch == '"')
            {
                i++; var sb = new StringBuilder();
                while (i < formula.Length)
                {
                    if (formula[i] == '"') { if (i + 1 < formula.Length && formula[i + 1] == '"') { sb.Append('"'); i += 2; } else { i++; break; } }
                    else { sb.Append(formula[i]); i++; }
                }
                tokens.Add(new Token(TT.String, sb.ToString())); continue;
            }

            // Quoted sheet reference: 'Sheet Name'!CellRef or 'Sheet Name'!Range
            // ECMA-376 §18.17: an inner apostrophe inside a quoted sheet identifier
            // is escaped as '' (two consecutive apostrophes). The closing quote is
            // a single apostrophe NOT followed by another apostrophe.
            if (ch == '\'')
            {
                var si = i + 1;
                var ei = si;
                while (ei < formula.Length)
                {
                    if (formula[ei] == '\'')
                    {
                        if (ei + 1 < formula.Length && formula[ei + 1] == '\'') { ei += 2; continue; }
                        break;
                    }
                    ei++;
                }
                if (ei < formula.Length && ei > si && ei + 1 < formula.Length && formula[ei + 1] == '!')
                {
                    var sheetName = formula[si..ei].Replace("''", "'");
                    i = ei + 2; // skip closing ' and '!'
                    var refStart = i;
                    while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '$' || formula[i] == ':')) i++;
                    var refPart = StripDollar(formula[refStart..i]);
                    if (refPart.Contains(':'))
                        tokens.Add(new Token(TT.SheetRange, $"{sheetName}!{refPart}"));
                    else
                        tokens.Add(new Token(TT.SheetCellRef, $"{sheetName}!{refPart.ToUpperInvariant()}"));
                    continue;
                }
            }

            if (char.IsDigit(ch) || ch == '.')
            {
                var ns = ParseNumber(formula, ref i);
                if (ns != null)
                {
                    // Entire-row range like `1:1` or `2:5` — pure digits on both sides of the colon.
                    // Expand2DRange clamps these to the sheet's populated column range.
                    if (i < formula.Length && formula[i] == ':' && Regex.IsMatch(ns, @"^\d+$"))
                    {
                        var peek = i + 1;
                        while (peek < formula.Length && char.IsDigit(formula[peek])) peek++;
                        if (peek > i + 1)
                        {
                            var rhsRow = formula[(i + 1)..peek];
                            i = peek;
                            tokens.Add(new Token(TT.Range, $"{ns}:{rhsRow}"));
                            continue;
                        }
                    }
                    tokens.Add(new Token(TT.Number, ns));
                    continue;
                }
            }

            if (char.IsLetter(ch) || ch == '_' || ch == '$')
            {
                var start = i;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '$' or '.')) i++;
                var word = formula[start..i]; var stripped = StripDollar(word);

                // TRUE / FALSE are boolean literals, but the TRUE() / FALSE()
                // function forms are followed by '(' — let those fall through to
                // the function-call path below rather than emitting a bool token
                // that leaves a stray '()' the parser can't consume.
                bool boolFollowedByParen = i < formula.Length && formula[i] == '(';
                if (!boolFollowedByParen && stripped.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) { tokens.Add(new Token(TT.Bool, "TRUE")); continue; }
                if (!boolFollowedByParen && stripped.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) { tokens.Add(new Token(TT.Bool, "FALSE")); continue; }

                // Unquoted sheet reference: SheetName!CellRef or SheetName!Range
                if (i < formula.Length && formula[i] == '!')
                {
                    var sheetName = word;
                    i++; // skip '!'
                    var refStart = i;
                    while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '$' || formula[i] == ':')) i++;
                    var refPart = StripDollar(formula[refStart..i]);
                    if (refPart.Contains(':'))
                        tokens.Add(new Token(TT.SheetRange, $"{sheetName}!{refPart}"));
                    else
                        tokens.Add(new Token(TT.SheetCellRef, $"{sheetName}!{refPart.ToUpperInvariant()}"));
                    continue;
                }

                if (i < formula.Length && formula[i] == ':' && IsCellRef(stripped))
                { i++; var s2 = i; while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '$')) i++;
                  tokens.Add(new Token(TT.Range, $"{stripped}:{StripDollar(formula[s2..i])}")); continue; }

                // Entire-column range like `A:A` or `A:C` — left side is letters-only (no row number).
                // Expand2DRange clamps these to the sheet's populated row range.
                if (i < formula.Length && formula[i] == ':' && Regex.IsMatch(stripped, @"^[A-Z]+$", RegexOptions.IgnoreCase))
                { i++; var s2 = i; while (i < formula.Length && (char.IsLetter(formula[i]) || formula[i] == '$')) i++;
                  var rhs = StripDollar(formula[s2..i]);
                  if (Regex.IsMatch(rhs, @"^[A-Z]+$", RegexOptions.IgnoreCase))
                  { tokens.Add(new Token(TT.Range, $"{stripped}:{rhs}")); continue; }
                  throw new NotSupportedException($"Unknown: {stripped}:{rhs}"); }

                // A name immediately followed by '(' is always a function call — a cell
                // reference is never followed by '('. Without this, a function whose name
                // is shaped like a cell ref (e.g. LOG10 = column LOG + row 10, matching
                // IsCellRef ^[A-Z]{1,3}\d+$) was misclassified as a ref and never evaluated.
                // Mirrors the (?![\w(]) guard in FormulaRefShifter.CellRefPattern.
                if (i < formula.Length && formula[i] == '(')
                { tokens.Add(new Token(TT.Func, word.Replace(".", "_").ToUpperInvariant())); continue; }

                if (IsCellRef(stripped)) { tokens.Add(new Token(TT.CellRef, stripped.ToUpperInvariant())); continue; }

                // Defined name. Two flavors:
                //   1. Literal range/cellref body — emit a single ref token
                //      (e.g. `StageTable` → `Data!A2:B7`).
                //   2. Formula body (OFFSET(...), INDIRECT(...), arithmetic) —
                //      inline the body's tokens here so the parent expression
                //      evaluates them in place.
                var definedNames = GetDefinedNames();
                if (definedNames.TryGetValue(stripped, out var defRef))
                {
                    var body = defRef.TrimStart('=').Trim();
                    // Defined name pointing at an error literal (e.g. the
                    // target sheet was deleted and the workbook persisted
                    // `<definedName>#REF!</definedName>`) must surface as
                    // that exact error, not collapse to #NAME? via the
                    // tokenize-fail catch-all below.
                    if (body.Length >= 2 && body[0] == '#' && body[^1] == '!')
                    {
                        tokens.Add(new Token(TT.Error, body));
                        continue;
                    }
                    if (TryDefinedNameAsSimpleRef(body) is { } refToken)
                    {
                        tokens.Add(refToken);
                        continue;
                    }
                    if (string.IsNullOrEmpty(body))
                        throw new NameResolutionException(stripped);
                    if (!_expandingNames.Add(stripped))
                        throw new NameResolutionException(stripped);
                    try
                    {
                        var inner = Tokenize(body);
                        if (inner.Count == 0) throw new NameResolutionException(stripped);
                        // Wrap the inlined body in parentheses so a name like
                        // MyName=A1+B1 evaluates as `(A1+B1)*2 = 2*(A1+B1)`,
                        // not `A1+B1*2` (textual substitution would break
                        // operator precedence).
                        tokens.Add(new Token(TT.LParen, "("));
                        tokens.AddRange(inner);
                        tokens.Add(new Token(TT.RParen, ")"));
                    }
                    catch (NotSupportedException) { throw new NameResolutionException(stripped); }
                    finally { _expandingNames.Remove(stripped); }
                    continue;
                }

                // Not a function, cell ref, or defined name: a bare identifier.
                // Emit a Name token and defer resolution to evaluation time —
                // LET / LAMBDA bind these in scope; anything still unbound
                // surfaces #NAME? from ParseAtom, the same end result as before.
                tokens.Add(new Token(TT.Name, stripped));
                continue;
            }
            throw new NotSupportedException($"Unexpected: {ch}");
        }
        return tokens;
    }

    private static string? ParseNumber(string s, ref int i)
    {
        var start = i;
        if (i < s.Length && (s[i] == '-' || s[i] == '+')) i++;
        var hasDigits = false;
        while (i < s.Length && char.IsDigit(s[i])) { i++; hasDigits = true; }
        if (i < s.Length && s[i] == '.') { i++; while (i < s.Length && char.IsDigit(s[i])) { i++; hasDigits = true; } }
        if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
        { i++; if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++; while (i < s.Length && char.IsDigit(s[i])) i++; }
        if (!hasDigits) { i = start; return null; }
        return s[start..i];
    }

    private static bool IsCellRef(string s) => Regex.IsMatch(s, @"^[A-Z]{1,3}\d+$", RegexOptions.IgnoreCase);
    private static string StripDollar(string s) => s.Replace("$", "");

    /// <summary>
    /// If the defined-name body is a single literal cell or range (with optional
    /// sheet prefix), return the corresponding token; otherwise null so the
    /// caller falls back to inlining the body as a sub-formula.
    /// </summary>
    private static Token? TryDefinedNameAsSimpleRef(string body)
    {
        var cleaned = StripDollar(body).Trim();
        string? sheet = null;
        var cell = cleaned;
        var bang = cleaned.IndexOf('!');
        if (bang > 0)
        {
            sheet = cleaned[..bang].Trim('\'');
            cell = cleaned[(bang + 1)..];
        }
        if (cell.Contains(':'))
        {
            // Bare A1:B5 or A:A or 1:1 is a literal range; OFFSET(A:A,...) is not.
            if (cell.Contains('(') || cell.Contains(',') || cell.Contains(' '))
                return null;
            return new Token(sheet != null ? TT.SheetRange : TT.Range,
                sheet != null ? $"{sheet}!{cell}" : cell);
        }
        if (IsCellRef(cell))
            return new Token(sheet != null ? TT.SheetCellRef : TT.CellRef,
                sheet != null ? $"{sheet}!{cell.ToUpperInvariant()}" : cell.ToUpperInvariant());
        return null;
    }

    // ==================== Recursive Descent Parser ====================

    private FormulaResult? ParseExpression(List<Token> t, ref int p) => ParseComparison(t, ref p);

    private FormulaResult? ParseComparison(List<Token> t, ref int p)
    {
        var left = ParseConcat(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Compare)
        {
            var op = t[p].Value; p++;
            var right = ParseConcat(t, ref p); if (right == null) return null;
            if (left.IsError) return left; if (right.IsError) return right;
            // Element-wise comparison when either side is array/range — needed
            // by the SUMPRODUCT((A1:A3>0)*1) conditional-count idiom. Returns
            // 0/1 doubles (not Bool) so downstream `*1` stays in numeric domain.
            if (HasArrayShape(left) || HasArrayShape(right))
            {
                left = ApplyComparison(left, right, op);
                if (left == null) return null;
                continue;
            }
            var cmp = CompareValues(left, right);
            left = op switch { "=" => FormulaResult.Bool(cmp == 0), "<>" => FormulaResult.Bool(cmp != 0),
                "<" => FormulaResult.Bool(cmp < 0), ">" => FormulaResult.Bool(cmp > 0),
                "<=" => FormulaResult.Bool(cmp <= 0), ">=" => FormulaResult.Bool(cmp >= 0), _ => null };
            if (left == null) return null;
        }
        return left;
    }

    // Sibling of ApplyBinaryOp for comparison operators. Element-wise on
    // arrays/ranges, scalar fallback otherwise. Returns FormulaResult.Array
    // of 0/1 doubles (treating BoolEval as numeric, matching how SUMPRODUCT
    // / SUM / multiplication consume the result).
    private FormulaResult? ApplyComparison(FormulaResult left, FormulaResult right, string op)
    {
        // Preserve the operand's 2-D shape (a column stays a column) so the result
        // pairs element-wise with other arrays — a flat 1-D result would be read
        // as a row and broadcast into a matrix (breaking SUMPRODUCT((col>0)*col)).
        // 0/1 doubles keep the `*1` conditional-count idiom in the numeric domain.
        var lg = AsGrid(left); var rg = AsGrid(right);
        int rows = Math.Max(lg?.GetLength(0) ?? 1, rg?.GetLength(0) ?? 1);
        int cols = Math.Max(lg?.GetLength(1) ?? 1, rg?.GetLength(1) ?? 1);
        var grid = new FormulaResult?[rows, cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
            {
                var l = CellAt(lg, left, i, j);
                var r = CellAt(rg, right, i, j);
                if (l.IsError) { grid[i, j] = l; continue; }
                if (r.IsError) { grid[i, j] = r; continue; }
                var cmp = CompareValues(l, r);
                grid[i, j] = FormulaResult.Number(op switch
                {
                    "=" => cmp == 0 ? 1 : 0,
                    "<>" => cmp != 0 ? 1 : 0,
                    "<" => cmp < 0 ? 1 : 0,
                    ">" => cmp > 0 ? 1 : 0,
                    "<=" => cmp <= 0 ? 1 : 0,
                    ">=" => cmp >= 0 ? 1 : 0,
                    _ => 0
                });
            }
        return FormulaResult.Area(new RangeData(grid));
    }

    private static FormulaResult?[]? AsResultArray(FormulaResult r)
    {
        if (r.IsArray) return r.ArrayValue!.Select(x => (FormulaResult?)FormulaResult.Number(x)).ToArray();
        if (r.IsRange) return r.RangeValue!.ToFlatResults();
        return null;
    }

    private FormulaResult? ParseConcat(List<Token> t, ref int p)
    {
        // Recursion re-entry point for every parenthesised sub-expression.
        // Trip on excessive nesting / low stack before a StackOverflowException.
        if (_parseDepth >= OfficeCli.Core.DocumentLimits.MaxRecursionDepth
            || !System.Runtime.CompilerServices.RuntimeHelpers.TryEnsureSufficientExecutionStack())
            return FormulaResult.Error("#NUM!");
        _parseDepth++;
        try
        {
        var left = ParseAddSub(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "&")
        { p++; var right = ParseAddSub(t, ref p); if (right == null) return null;
          // An error propagates, but the rest of the operator chain must still be
          // consumed or the top-level "all tokens parsed" check fails and turns
          // the error into a NOTEVAL. Keep the leftmost error and keep scanning.
          if (left.IsError) continue; if (right.IsError) { left = right; continue; }
          left = FormulaResult.Str(left.AsString() + right.AsString()); }
        return left;
        }
        finally { _parseDepth--; }
    }

    private FormulaResult? ParseAddSub(List<Token> t, ref int p)
    {
        var left = ParseMulDiv(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value is "+" or "-")
        { var op = t[p].Value; p++; var r = ParseMulDiv(t, ref p); if (r == null) return null;
          if (left.IsError) continue; if (r.IsError) { left = r; continue; }
          Func<double, double, FormulaResult> f = op == "+"
              ? (a, b) => FormulaResult.Number(a + b)
              : (a, b) => FormulaResult.Number(a - b);
          left = ApplyBinaryOp(left, r, f); }
        return left;
    }

    private FormulaResult? ParseMulDiv(List<Token> t, ref int p)
    {
        var left = ParsePower(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value is "*" or "/")
        { var op = t[p].Value; p++; var r = ParsePower(t, ref p); if (r == null) return null;
          if (left.IsError) continue; if (r.IsError) { left = r; continue; }
          // Division by zero is #DIV/0! per element (scalar → the whole result;
          // array/range → only the zero-divisor cells, so aggregates can still
          // ignore them and SUM propagates via CheckRangeErrors).
          Func<double, double, FormulaResult> f = op == "/"
              ? (a, b) => b == 0 ? FormulaResult.Error("#DIV/0!") : FormulaResult.Number(a / b)
              : (a, b) => FormulaResult.Number(a * b);
          left = ApplyBinaryOp(left, r, f); }
        return left;
    }

    private FormulaResult? ParsePower(List<Token> t, ref int p)
    {
        var b = ParseUnary(t, ref p); if (b == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "^")
        { p++; var e = ParseUnary(t, ref p); if (e == null) return null;
          if (b.IsError) continue; if (e.IsError) { b = e; continue; }
          b = ApplyBinaryOp(b, e, (x, y) =>
          { var pr = ExcelPow(x, y); return double.IsNaN(pr) || double.IsInfinity(pr) ? FormulaResult.Error("#NUM!") : FormulaResult.Number(pr); }); }
        return b;
    }

    // Element-wise application of a binary numeric op. Handles scalar+scalar,
    // array+scalar, scalar+array, array+array. Range operands are flattened
    // row-major (empties treated as 0, matching Excel implicit-zero coercion).
    // Length mismatch in array+array uses Min(len) — Excel would emit #N/A, but
    // min-length is more lenient and only affects malformed inputs.
    // Element-wise binary op. Scalar+scalar returns a scalar; any array/range
    // operand yields a 2-D Area that preserves shape AND per-element errors, so
    // INDEX can address it, aggregates can ignore error cells, and SUM propagates
    // them via CheckRangeErrors. A singleton row/column broadcasts; out-of-range
    // positions in a mismatched pairing are #N/A.
    private static FormulaResult ApplyBinaryOp(FormulaResult left, FormulaResult right, Func<double, double, FormulaResult> op)
    {
        var lg = AsGrid(left); var rg = AsGrid(right);
        if (lg == null && rg == null) return ElemOp(left, right, op);
        int rows = Math.Max(lg?.GetLength(0) ?? 1, rg?.GetLength(0) ?? 1);
        int cols = Math.Max(lg?.GetLength(1) ?? 1, rg?.GetLength(1) ?? 1);
        var grid = new FormulaResult?[rows, cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
                grid[i, j] = ElemOp(CellAt(lg, left, i, j), CellAt(rg, right, i, j), op);
        return FormulaResult.Area(new RangeData(grid));
    }

    // 2-D cell grid of an operand, or null for a scalar. A 1-D array is treated
    // as a single row.
    private static FormulaResult?[,]? AsGrid(FormulaResult r)
    {
        if (r.IsRange) return r.RangeValue!.Cells;
        if (r.IsArray)
        {
            var a = r.ArrayValue!; var g = new FormulaResult?[1, a.Length];
            for (int j = 0; j < a.Length; j++) g[0, j] = FormulaResult.Number(a[j]);
            return g;
        }
        return null;
    }

    // Element at (i,j) with broadcasting: a null grid is the scalar; a singleton
    // row/column repeats; anything else out of range is #N/A. A blank cell is 0.
    private static FormulaResult CellAt(FormulaResult?[,]? g, FormulaResult scalar, int i, int j)
    {
        if (g == null) return scalar;
        int gr = g.GetLength(0), gc = g.GetLength(1);
        int ri = gr == 1 ? 0 : i, cj = gc == 1 ? 0 : j;
        if (ri >= gr || cj >= gc) return FormulaResult.Error("#N/A");
        return g[ri, cj] ?? FormulaResult.Number(0);
    }

    // Single-pair application: propagate an error operand, reject non-numeric
    // text (#VALUE!), else run the numeric op.
    private static FormulaResult ElemOp(FormulaResult a, FormulaResult b, Func<double, double, FormulaResult> op)
    {
        if (a.IsError) return a;
        if (b.IsError) return b;
        if (!TryCoerceArithmetic(a, out var av) || !TryCoerceArithmetic(b, out var bv))
            return FormulaResult.Error("#VALUE!");
        return op(av, bv);
    }

    private static bool HasArrayShape(FormulaResult r) => r.IsArray || r.IsRange;

    // Scalar arithmetic coercion. Numbers, booleans and blank cells (→0) always
    // coerce; text coerces only when numeric-looking. Non-numeric or empty text
    // is not coercible and the caller must surface #VALUE!.
    private static bool TryCoerceArithmetic(FormulaResult r, out double val)
    {
        if (r.IsBlank) { val = 0; return true; }
        if (r.IsString)
            return double.TryParse(r.StringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out val);
        val = r.AsNumber();
        return true;
    }

    // Parse the body of an array constant `{...}` (without the braces).
    // Rows are separated by ';', columns by ',' — per ECMA-376 §18.17.7.282.
    // Each cell is a number / "string" / TRUE / FALSE. Produces a RangeData
    // wrapped as Area so ApplyBinaryOp and aggregate functions handle it
    // identically to a real range. BaseRow/BaseCol stay 0 (not a workbook reference).
    // Split an array-constant body on a separator, ignoring separators that sit
    // inside a double-quoted string element (e.g. the comma in {",",";"}).
    private static List<string> SplitArrayConstant(string s, char sep)
    {
        var parts = new List<string>();
        bool inStr = false; int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '"') inStr = !inStr;
            else if (s[i] == sep && !inStr) { parts.Add(s[start..i]); start = i + 1; }
        }
        parts.Add(s[start..]);
        return parts;
    }

    private static FormulaResult ParseArrayConstant(string body)
    {
        var rows = SplitArrayConstant(body, ';');
        var rowCells = rows.Select(r => SplitArrayConstant(r, ',').Select(c => c.Trim()).ToArray()).ToArray();
        var cols = rowCells.Max(r => r.Length);
        var cells = new FormulaResult?[rowCells.Length, cols];
        for (int r = 0; r < rowCells.Length; r++)
            for (int c = 0; c < cols; c++)
            {
                var s = c < rowCells[r].Length ? rowCells[r][c] : "";
                cells[r, c] = ParseArrayConstantCell(s);
            }
        return FormulaResult.Area(new RangeData(cells));
    }

    private static FormulaResult? ParseArrayConstantCell(string s)
    {
        if (s.Length == 0) return null;
        if (s.Length >= 2 && s[0] == '"' && s[^1] == '"') return FormulaResult.Str(s[1..^1].Replace("\"\"", "\""));
        if (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return FormulaResult.Bool(true);
        if (s.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return FormulaResult.Bool(false);
        if (s.StartsWith('#') && s.EndsWith('!')) return FormulaResult.Error(s);
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var n)) return FormulaResult.Number(n);
        return FormulaResult.Str(s);
    }

    private static double[]? AsArrayLike(FormulaResult r)
    {
        if (r.IsArray) return r.ArrayValue;
        if (r.IsRange)
        {
            var rd = r.RangeValue!; var n = rd.Rows * rd.Cols; var a = new double[n];
            for (int rr = 0; rr < rd.Rows; rr++)
                for (int cc = 0; cc < rd.Cols; cc++)
                    a[rr * rd.Cols + cc] = rd.Cells[rr, cc]?.AsNumber() ?? 0;
            return a;
        }
        return null;
    }

    private FormulaResult? ParseUnary(List<Token> t, ref int p)
    {
        if (p < t.Count && t[p].Type == TT.Op)
        {
            if (t[p].Value == "-") { p++; var v = ParseUnary(t, ref p); if (v == null) return null;
                if (v.IsError) return v;
                // Element-wise negate for both Array and Range operands —
                // previously only IsArray was handled, so `-A1:A3` collapsed
                // via AsNumber to -FirstCell instead of producing an array.
                if (HasArrayShape(v))
                    return FormulaResult.Array(AsArrayLike(v)!.Select(x => -x).ToArray());
                if (!TryCoerceArithmetic(v, out var uv)) return FormulaResult.Error("#VALUE!");
                return FormulaResult.Number(-uv); }
            if (t[p].Value == "+") { p++; return ParseUnary(t, ref p); }
        }
        return ParsePostfix(t, ref p);
    }

    private FormulaResult? ParsePostfix(List<Token> t, ref int p)
    {
        var v = ParseAtom(t, ref p); if (v == null) return null;
        // Immediately-invoked LAMBDA: LAMBDA(x, x+1)(5).
        while (v.IsLambda && p < t.Count && t[p].Type == TT.LParen)
            v = InvokeLambda((Lambda)v.LambdaValue!, ParseCallArgs(t, ref p));
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "%") { p++; if (!TryCoerceArithmetic(v, out var pv)) return FormulaResult.Error("#VALUE!"); v = FormulaResult.Number(pv / 100.0); }
        return v;
    }

    private FormulaResult? ParseAtom(List<Token> t, ref int p)
    {
        if (p >= t.Count) return null;
        var tok = t[p];
        switch (tok.Type)
        {
            case TT.Number: p++; return double.TryParse(tok.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var n) ? FormulaResult.Number(n) : null;
            case TT.String: p++; return FormulaResult.Str(tok.Value);
            case TT.Bool: p++; return FormulaResult.Bool(tok.Value == "TRUE");
            case TT.CellRef: p++; return ResolveCellResult(tok.Value);
            case TT.SheetCellRef: p++; return ResolveSheetCellResult(tok.Value);
            // Range tokens that reach ParseAtom (e.g. inside an arithmetic expression
            // like B1:B3*1) become Area FormulaResults so ApplyBinaryOp can do
            // element-wise math. Range tokens appearing directly as function args
            // are intercepted earlier by ParseFunction and bypass this path.
            case TT.Range: p++; return FormulaResult.Area(Expand2DRange(tok.Value));
            case TT.SheetRange: p++; return FormulaResult.Area(Expand2DRange(tok.Value));
            case TT.ArrayLit: p++; return ParseArrayConstant(tok.Value);
            case TT.Error: p++; return FormulaResult.Error(tok.Value);
            case TT.Name:
                p++;
                if (_bindings.TryGetValue(tok.Value, out var bound)) return bound;
                throw new NameResolutionException(tok.Value);
            case TT.LParen: p++; var inner = ParseExpression(t, ref p); if (p < t.Count && t[p].Type == TT.RParen) p++; return inner;
            case TT.Func: return ParseFunction(t, ref p);
            default: return null;
        }
    }

    private FormulaResult? ParseFunction(List<Token> t, ref int p)
    {
        var name = t[p].Value; p++;
        if (p >= t.Count || t[p].Type != TT.LParen) return null; p++;

        // LET / LAMBDA capture their arguments as token ranges (binding names and
        // an unevaluated body) rather than eager evaluation.
        if (name == "LET") return EvalLet(t, ref p);
        if (name == "LAMBDA") return MakeLambda(t, ref p);
        // A LET/LAMBDA-bound name invoked as f(args).
        if (_bindings.TryGetValue(name, out var boundFn) && boundFn.IsLambda)
            return InvokeLambda((Lambda)boundFn.LambdaValue!, ParseCallArgs(t, ref p));

        var args = new List<object>();
        var argIdx = 0;
        if (p < t.Count && t[p].Type != TT.RParen)
        {
            while (true)
            {
                // Empty arg (immediate comma or close-paren after a comma) — Excel
                // treats omitted args as 0 for numeric-arg functions like OFFSET.
                if (p < t.Count && (t[p].Type == TT.Comma || t[p].Type == TT.RParen))
                { args.Add(FormulaResult.Number(0)); }
                else if (((argIdx == 0 && name is "OFFSET" or "ISREF" or "ISFORMULA" or "SHEET" or "ROW" or "COLUMN")
                          || (argIdx == 1 && name is "CELL"))
                         && TryParseRefArg(t, ref p) is { } refArg)
                { args.Add(refArg); }
                else if (p < t.Count && t[p].Type is TT.Range or TT.SheetRange
                         && (p + 1 >= t.Count || t[p + 1].Type is TT.Comma or TT.RParen))
                { args.Add(Expand2DRange(t[p].Value)); p++; }
                else { var expr = ParseExpression(t, ref p); if (expr == null) return null; args.Add(expr); }
                argIdx++;
                if (p >= t.Count || t[p].Type != TT.Comma) break; p++;
            }
        }
        if (p < t.Count && t[p].Type == TT.RParen) p++;
        return EvalFunction(name, args);
    }

    // Sentinel bound to a LAMBDA parameter that the caller did not supply, so
    // ISOMITTED can detect it.
    private static readonly FormulaResult OmittedArg = new() { StringValue = " __OCLI_OMITTED__" };
    private static bool IsOmittedArg(object? o) => ReferenceEquals(o, OmittedArg);

    // Consume `( a, b, … )` (p at the LParen already consumed by the caller) and
    // return the evaluated argument values for a lambda call.
    private List<FormulaResult> ParseCallArgs(List<Token> t, ref int p)
    {
        var argv = new List<FormulaResult>();
        if (p < t.Count && t[p].Type != TT.RParen)
            while (true)
            {
                // An explicit empty slot (f(10,)) is an omitted argument that
                // ISOMITTED can detect; a genuinely missing trailing argument is
                // caught by the parameter-count check in InvokeLambda.
                if (p < t.Count && (t[p].Type == TT.Comma || t[p].Type == TT.RParen))
                    argv.Add(OmittedArg);
                else { var a = ParseExpression(t, ref p); argv.Add(a ?? FormulaResult.Error("#VALUE!")); }
                if (p < t.Count && t[p].Type == TT.Comma) { p++; continue; }
                break;
            }
        if (p < t.Count && t[p].Type == TT.RParen) p++;
        return argv;
    }

    // Capture one top-level argument's tokens (balanced parens; stop at a
    // top-level comma or the closing paren) without evaluating them.
    private static List<Token> CaptureArg(List<Token> t, ref int p)
    {
        int depth = 0, start = p;
        while (p < t.Count)
        {
            var tt = t[p].Type;
            if (depth == 0 && (tt == TT.Comma || tt == TT.RParen)) break;
            if (tt == TT.LParen) depth++;
            else if (tt == TT.RParen) depth--;
            p++;
        }
        return t.GetRange(start, p - start);
    }

    // Capture every top-level argument as a token range; consume the closing paren.
    private static List<List<Token>> CaptureAllArgs(List<Token> t, ref int p)
    {
        var parts = new List<List<Token>>();
        if (p < t.Count && t[p].Type != TT.RParen)
            while (true)
            {
                parts.Add(CaptureArg(t, ref p));
                if (p < t.Count && t[p].Type == TT.Comma) { p++; continue; }
                break;
            }
        if (p < t.Count && t[p].Type == TT.RParen) p++;
        return parts;
    }

    private FormulaResult? EvalTokens(List<Token> body) { int q = 0; return ParseExpression(body, ref q); }

    // LET(name1, value1, …, calculation) — bind each name to its value (in order,
    // so later values can reference earlier names) then evaluate the calculation.
    private FormulaResult? EvalLet(List<Token> t, ref int p)
    {
        var parts = CaptureAllArgs(t, ref p);
        if (parts.Count < 3 || parts.Count % 2 == 0) return FormulaResult.Error("#VALUE!");
        var snapshot = new Dictionary<string, FormulaResult>(_bindings, StringComparer.OrdinalIgnoreCase);
        try
        {
            for (int i = 0; i < parts.Count - 1; i += 2)
            {
                if (parts[i].Count != 1 || parts[i][0].Type != TT.Name) return FormulaResult.Error("#VALUE!");
                var val = EvalTokens(parts[i + 1]);
                if (val == null) return FormulaResult.Error("#VALUE!");
                _bindings[parts[i][0].Value] = val;
            }
            return EvalTokens(parts[^1]);
        }
        finally { RestoreBindings(snapshot); }
    }

    // LAMBDA(param1, …, paramN, body) — a value capturing the parameter names and
    // the unevaluated body tokens.
    private FormulaResult? MakeLambda(List<Token> t, ref int p)
    {
        var parts = CaptureAllArgs(t, ref p);
        if (parts.Count < 1) return FormulaResult.Error("#VALUE!");
        var pars = new List<string>();
        for (int i = 0; i < parts.Count - 1; i++)
        {
            if (parts[i].Count != 1 || parts[i][0].Type != TT.Name) return FormulaResult.Error("#VALUE!");
            pars.Add(parts[i][0].Value);
        }
        return new FormulaResult { LambdaValue = new Lambda(pars, parts[^1]) };
    }

    // Bind the lambda's parameters to the supplied (or omitted) arguments, evaluate
    // the body, then restore the enclosing scope.
    private FormulaResult InvokeLambda(Lambda lam, List<FormulaResult> argv)
    {
        // Excel requires every parameter to have a slot; an under-supplied call
        // is #VALUE! (a supplied-but-empty slot binds OmittedArg for ISOMITTED).
        if (argv.Count < lam.Parameters.Count) return FormulaResult.Error("#VALUE!");
        var snapshot = new Dictionary<string, FormulaResult>(_bindings, StringComparer.OrdinalIgnoreCase);
        try
        {
            for (int i = 0; i < lam.Parameters.Count; i++)
                _bindings[lam.Parameters[i]] = argv[i];
            return EvalTokens(lam.Body) ?? FormulaResult.Error("#VALUE!");
        }
        finally { RestoreBindings(snapshot); }
    }

    private void RestoreBindings(Dictionary<string, FormulaResult> snapshot)
    {
        _bindings.Clear();
        foreach (var kv in snapshot) _bindings[kv.Key] = kv.Value;
    }

    /// <summary>
    /// Peek the next token; if it's a CellRef / SheetCellRef / Range / SheetRange,
    /// consume it and return a RefArg without dereferencing the cells. Used by
    /// reference-consuming functions (OFFSET) whose first argument must remain
    /// a reference instead of being eagerly evaluated to a scalar value.
    /// </summary>
    private RefArg? TryParseRefArg(List<Token> t, ref int p)
    {
        if (p >= t.Count) return null;
        var tok = t[p];
        switch (tok.Type)
        {
            case TT.CellRef:
            {
                var (col, row) = ParseRef(tok.Value);
                p++;
                return new RefArg(null, ColToIndex(col), row, 1, 1);
            }
            case TT.SheetCellRef:
            {
                var bang = tok.Value.IndexOf('!');
                var sheet = tok.Value[..bang];
                var (col, row) = ParseRef(tok.Value[(bang + 1)..]);
                p++;
                return new RefArg(sheet, ColToIndex(col), row, 1, 1);
            }
            case TT.Range:
                p++;
                return BuildRefFromRange(null, tok.Value);
            case TT.SheetRange:
            {
                var bang = tok.Value.IndexOf('!');
                var sheet = tok.Value[..bang];
                p++;
                return BuildRefFromRange(sheet, tok.Value[(bang + 1)..]);
            }
            default:
                return null;
        }
    }

    // ==================== Cell & Range Resolution ====================

    internal FormulaResult? ResolveCellResult(string cellRef)
    {
        cellRef = StripDollar(cellRef).ToUpperInvariant();
        var qualifiedRef = string.IsNullOrEmpty(_sheetKey) ? cellRef : $"{_sheetKey}!{cellRef}";
        if (!_visiting.Add(qualifiedRef))
        {
            // Circular ref: use 0 as initial value (matches Excel iterative calc).
            // Count the hit so in-flight evaluations know their result is
            // entry-point-dependent and must not be memoized.
            _session.CircularHits++;
            return FormulaResult.Number(0);
        }
        try
        {
            var cell = FindCell(cellRef);
            if (cell == null) return FormulaResult.Blank();

            // If cell has a formula, always evaluate it (cached values may be stale).
            // Guard recursive evaluation against an uncatchable StackOverflow that
            // would kill the resident process (DoS).
            if (cell.CellFormula?.Text != null)
            {
                // Memoized? Referenced formula cells are re-evaluated (their
                // cached <v> may be stale), but within one session the formula's
                // own result cannot change — reuse it. Skipped while LET/LAMBDA
                // bindings are live: the referenced cell's evaluation currently
                // sees the caller's bindings (pre-existing quirk), so a result
                // computed under bindings must not leak into other contexts.
                if (_bindings.Count == 0 && _session.CellMemo.TryGetValue(qualifiedRef, out var memoized))
                    return memoized;
                // Primary: probe the real remaining stack (adapts to formula
                // complexity, so complex nested formulas are covered too).
                // Secondary: a high fixed backstop. Over either, surface a
                // visible #NUM! that propagates up the chain (B[N-1]+A[N] returns
                // the error) — never a silent 0 or an uncatchable crash.
                if (_sameSheetDepth >= MaxSameSheetDepth
                    || !System.Runtime.CompilerServices.RuntimeHelpers.TryEnsureSufficientExecutionStack())
                    return FormulaResult.Error("#NUM!");
                _sameSheetDepth++;
                // _parseDepth bounds PER-FORMULA paren nesting only (the
                // dos-hardening cap in ParseConcat). The referenced cell's
                // formula re-enters ParseConcat while THIS formula's parse is
                // still on the stack, so without a reset the counter
                // accumulates one frame per chain link and a >MaxRecursionDepth
                // simple chain (B[N]=B[N-1]+A[N]) trips the cap mid-chain —
                // ParseConcat bails with pos=0, EvaluateFormula returns null,
                // and the link silently degrades to Blank()/0. Cross-link depth
                // is already guarded above by the stack probe + the
                // MaxSameSheetDepth backstop; each formula's parse recursion
                // must be counted from zero.
                var savedParseDepth = _parseDepth;
                _parseDepth = 0;
                try
                {
                    var circularBefore = _session.CircularHits;
                    var evaluated = EvaluateFormula(ModernFunctionQualifier.Unqualify(cell.CellFormula.Text));
                    if (evaluated != null)
                    {
                        // Memoize only clean results: no live bindings (see lookup
                        // guard above) and no circular fallback during this
                        // evaluation (a 0-seeded cycle result depends on where the
                        // cycle was entered). Lambdas capture evaluator state and
                        // are not safe to replay.
                        if (_bindings.Count == 0 && !evaluated.IsLambda
                            && circularBefore == _session.CircularHits)
                            _session.CellMemo[qualifiedRef] = evaluated;
                        return evaluated;
                    }
                }
                catch { /* fall through to cached value */ }
                finally { _sameSheetDepth--; _parseDepth = savedParseDepth; }
            }

            // InlineString cells store their text in <is><t>…</t></is>, NOT in
            // <v>. Reading CellValue?.Text returns null and the inline content
            // would silently degrade to 0 in any reference. Pull from
            // cell.InlineString.InnerText first when DataType says inlineStr.
            var cached = cell.DataType?.Value == CellValues.InlineString
                ? cell.InlineString?.InnerText
                : cell.CellValue?.Text;
            if (!string.IsNullOrEmpty(cached))
            {
                if (cell.DataType?.Value == CellValues.SharedString)
                {
                    var sst = _workbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (sst?.SharedStringTable != null && int.TryParse(cached, out int idx))
                        return FormulaResult.Str(sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx)?.InnerText ?? cached);
                    return FormulaResult.Str(cached);
                }
                if (cell.DataType?.Value == CellValues.Boolean) return FormulaResult.Bool(cached == "1");
                // BUG R4-4: error-typed cells (DataType=Error, e.g. cached "#REF!"
                // written by `Set value=#REF! type=error`) must propagate as an
                // Error FormulaResult so downstream formulas like =A1+1 return
                // #REF! instead of coercing the cached string to a number.
                if (cell.DataType?.Value == CellValues.Error) return FormulaResult.Error(cached);
                if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString) return FormulaResult.Str(cached);
                return double.TryParse(cached, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? FormulaResult.Number(v) : FormulaResult.Str(cached);
            }

            return FormulaResult.Blank();
        }
        finally { _visiting.Remove(qualifiedRef); }
    }

    /// <summary>
    /// Resolve a cross-sheet cell reference like "SheetName!A1".
    /// Creates a new evaluator for the target sheet and resolves the cell there.
    /// </summary>
    private FormulaResult? ResolveSheetCellResult(string sheetCellRef)
    {
        // Depth guard: over the cap, surface a visible #NUM! (propagates up the
        // chain) rather than Number(0), which leaked as a silent wrong value
        // (e.g. a 25-sheet chain reporting 22 instead of erroring). Matches the
        // same-sheet ResolveCellResult guard — depth exceeded → visible error,
        // never a silent numeric lie.
        // Chain depth lives on the session (not the instance) because child
        // evaluators are cached per sheet and reused at whatever depth the
        // current chain happens to be.
        if (_session.CrossSheetDepth > 20) return FormulaResult.Error("#NUM!"); // depth guard

        var bangIdx = sheetCellRef.IndexOf('!');
        if (bangIdx < 0) return FormulaResult.Number(0);

        var sheetName = sheetCellRef[..bangIdx];
        var cellRef = sheetCellRef[(bangIdx + 1)..];

        var sheetData = GetSheetDataFor(sheetName);
        // R3 BUG C: if the sheet name is non-empty and unresolved, the
        // reference itself is invalid (Excel: #REF!). The "0 fallback" was
        // historically applied here, but it's only correct for an existing
        // sheet with an empty cell — never for a missing sheet. INDIRECT,
        // direct cross-sheet refs (Sheet999!A1), and Expand2DRange all rely
        // on this path; surfacing #REF! here is Excel-correct in every case.
        if (sheetData == null)
        {
            if (!string.IsNullOrEmpty(sheetName)) return FormulaResult.Error("#REF!");
            return FormulaResult.Number(0);
        }

        // ResolveCellResult will handle circular detection using qualified ref
        // (sheetKey!cellRef). Reuse one child evaluator per sheet: a fresh
        // instance per dereference rebuilt _cellIndex (a full sheet scan) for
        // EVERY cell read through a cross-sheet range — the dominant cost on
        // SUMIFS/COUNTIF-heavy workbooks.
        if (!_session.SheetEvaluators.TryGetValue(sheetName, out var eval) || !ReferenceEquals(eval._sheetData, sheetData))
        {
            eval = new FormulaEvaluator(sheetData, _workbookPart, _session, _depth + 1, sheetName);
            _session.SheetEvaluators[sheetName] = eval;
        }
        _session.CrossSheetDepth++;
        try { return eval.ResolveCellResult(cellRef); }
        finally { _session.CrossSheetDepth--; }
    }

    /// <summary>
    /// Resolve a sheet name to its SheetData (or return _sheetData for null/empty name).
    /// </summary>
    private SheetData? GetSheetDataFor(string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName)) return _sheetData;
        if (_workbookPart == null) return null;
        if (_session.SheetDataByName.TryGetValue(sheetName, out var cachedSheet)) return cachedSheet;
        SheetData? resolved;
        try
        {
            var sheet = _workbookPart.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            var wsPart = sheet?.Id?.Value != null ? (WorksheetPart)_workbookPart.GetPartById(sheet.Id!.Value!) : null;
            resolved = wsPart?.Worksheet?.GetFirstChild<SheetData>();
        }
        catch { resolved = null; }
        _session.SheetDataByName[sheetName] = resolved;
        return resolved;
    }

    /// <summary>
    /// Scan a sheet's populated rows to find min/max row index. Returns (0,0) if empty.
    /// Used to clamp entire-column references like "A:A" to the actual data area.
    /// </summary>
    private (int minRow, int maxRow) GetPopulatedRowRange(SheetData sheetData)
    {
        if (_session.RowExtentBySheet.TryGetValue(sheetData, out var cached)) return cached;
        int minRow = int.MaxValue, maxRow = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            if (row.RowIndex?.Value is uint idx)
            {
                var i = (int)idx;
                if (i < minRow) minRow = i;
                if (i > maxRow) maxRow = i;
            }
        }
        var extent = maxRow == 0 ? (0, 0) : (minRow, maxRow);
        _session.RowExtentBySheet[sheetData] = extent;
        return extent;
    }

    /// <summary>
    /// Scan a sheet's populated cells to find min/max column index. Returns (0,0) if empty.
    /// Used to clamp entire-row references like "1:1" to the actual data area.
    /// </summary>
    private (int minCol, int maxCol) GetPopulatedColRange(SheetData sheetData)
    {
        if (_session.ColExtentBySheet.TryGetValue(sheetData, out var cached)) return cached;
        int minCol = int.MaxValue, maxCol = 0;
        foreach (var row in sheetData.Elements<Row>())
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is string cref)
                {
                    var m = Regex.Match(cref, @"^([A-Z]+)\d+$", RegexOptions.IgnoreCase);
                    if (m.Success)
                    {
                        var idx = ColToIndex(m.Groups[1].Value.ToUpperInvariant());
                        if (idx < minCol) minCol = idx;
                        if (idx > maxCol) maxCol = idx;
                    }
                }
            }
        var extent = maxCol == 0 ? (0, 0) : (minCol, maxCol);
        _session.ColExtentBySheet[sheetData] = extent;
        return extent;
    }

    private Cell? FindCell(string cellRef)
    {
        if (_cellIndex == null)
        {
            _cellIndex = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in _sheetData.Elements<Row>())
                foreach (var cell in row.Elements<Cell>())
                    if (cell.CellReference?.Value != null)
                        _cellIndex[cell.CellReference.Value] = cell;
        }
        return _cellIndex.TryGetValue(cellRef, out var found) ? found : null;
    }

    private RangeData Expand2DRange(string rangeExpr)
    {
        // Handle cross-sheet ranges like "SheetName!A1:B3"
        string? sheetPrefix = null;
        var expr = rangeExpr;
        var bangIdx = rangeExpr.IndexOf('!');
        if (bangIdx >= 0)
        {
            sheetPrefix = rangeExpr[..bangIdx];
            expr = rangeExpr[(bangIdx + 1)..];
        }

        var parts = expr.Split(':');
        if (parts.Length != 2) return new RangeData(new FormulaResult?[0, 0]);

        var left = StripDollar(parts[0]);
        var right = StripDollar(parts[1]);
        int r1, r2, cMin, cMax;

        // Entire-column reference like "A:A" or "A:C" — clamp to populated row range
        // of the target sheet (Excel would otherwise scan all 1,048,576 rows).
        var leftColOnly = Regex.IsMatch(left, @"^[A-Z]+$", RegexOptions.IgnoreCase);
        var rightColOnly = Regex.IsMatch(right, @"^[A-Z]+$", RegexOptions.IgnoreCase);
        // Entire-row reference like "1:1" or "2:5"
        var leftRowOnly = Regex.IsMatch(left, @"^\d+$");
        var rightRowOnly = Regex.IsMatch(right, @"^\d+$");

        if (leftColOnly && rightColOnly)
        {
            var c1 = ColToIndex(left.ToUpperInvariant());
            var c2 = ColToIndex(right.ToUpperInvariant());
            cMin = Math.Min(c1, c2); cMax = Math.Max(c1, c2);
            var targetSheet = GetSheetDataFor(sheetPrefix);
            if (targetSheet == null) return new RangeData(new FormulaResult?[0, 0]);
            var (minRow, maxRow) = GetPopulatedRowRange(targetSheet);
            if (maxRow == 0) return new RangeData(new FormulaResult?[0, 0]);
            r1 = minRow; r2 = maxRow;
        }
        else if (leftRowOnly && rightRowOnly)
        {
            r1 = Math.Min(int.Parse(left), int.Parse(right));
            r2 = Math.Max(int.Parse(left), int.Parse(right));
            var targetSheet = GetSheetDataFor(sheetPrefix);
            if (targetSheet == null) return new RangeData(new FormulaResult?[0, 0]);
            var (minCol, maxCol) = GetPopulatedColRange(targetSheet);
            if (maxCol == 0) return new RangeData(new FormulaResult?[0, 0]);
            cMin = minCol; cMax = maxCol;
        }
        else
        {
            var (col1, row1) = ParseRef(left);
            var (col2, row2) = ParseRef(right);
            var c1 = ColToIndex(col1); var c2 = ColToIndex(col2);
            r1 = Math.Min(row1, row2); r2 = Math.Max(row1, row2);
            cMin = Math.Min(c1, c2); cMax = Math.Max(c1, c2);
        }

        var rows = r2 - r1 + 1; var cols = cMax - cMin + 1;
        // Same rect-shaped memo as ResolveRef — literal range tokens ("A1:B3",
        // "Sheet!A:A") route here instead, and repeat just as often.
        var rangeMemoKey = rows * cols > 1
            ? $"{sheetPrefix ?? _sheetKey}|{cMin},{r1},{cols},{rows}"
            : null;
        if (rangeMemoKey != null && _session.RangeMemo.TryGetValue(rangeMemoKey, out var memoRange))
            return memoRange;
        var circularBefore = _session.CircularHits;
        var cells = new FormulaResult?[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                var cellRef = $"{IndexToCol(cMin + c)}{r1 + r}";
                cells[r, c] = sheetPrefix != null
                    ? ResolveSheetCellResult($"{sheetPrefix}!{cellRef}")
                    : ResolveCellResult(cellRef);
            }
        // R3-1: preserve the range's origin so ROW() / COLUMN() / ADDRESS() can
        // answer correctly when given a literal range token (`A1:B3`) — the
        // tokenizer routes those through Expand2DRange, bypassing ResolveRef
        // where Round 2 introduced BaseRow/BaseCol propagation.
        var range = new RangeData(cells) { BaseRow = r1, BaseCol = cMin, BaseSheet = sheetPrefix };
        if (rangeMemoKey != null && circularBefore == _session.CircularHits)
            _session.RangeMemo[rangeMemoKey] = range;
        return range;
    }

    private static (string col, int row) ParseRef(string r)
    {
        var m = Regex.Match(r, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        return m.Success ? (m.Groups[1].Value.ToUpperInvariant(), int.Parse(m.Groups[2].Value)) : ("A", 1);
    }

    private static int ColToIndex(string col) { int r = 0; foreach (var c in col.ToUpperInvariant()) r = r * 26 + (c - 'A' + 1); return r; }
    private static string IndexToCol(int i) { var r = ""; while (i > 0) { i--; r = (char)('A' + i % 26) + r; i /= 26; } return r; }
}
