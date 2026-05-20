// Copyright 2025 OfficeCLI (officecli.ai)
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

    public bool IsNumeric => NumericValue.HasValue;
    public bool IsString => StringValue != null;
    public bool IsBool => BoolValue.HasValue;
    public bool IsError => ErrorValue != null;
    public bool IsArray => ArrayValue != null;
    public bool IsRange => RangeValue != null;

    public static FormulaResult Number(double v) => new() { NumericValue = v };
    public static FormulaResult Str(string v) => new() { StringValue = v };
    public static FormulaResult Bool(bool v) => new() { BoolValue = v };
    public static FormulaResult Error(string v) => new() { ErrorValue = v };
    public static FormulaResult Array(double[] v) => new() { ArrayValue = v };
    public static FormulaResult Area(RangeData v) => new() { RangeValue = v };

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
/// Excel formula evaluator supporting 150+ functions.
/// Split across partial class files:
///   FormulaEvaluator.cs          — core: tokenizer, parser, cell resolution
///   FormulaEvaluator.Functions.cs — function dispatch + implementations
///   FormulaEvaluator.Helpers.cs   — math utilities, comparison helpers
/// </summary>
internal partial class FormulaEvaluator
{
    private readonly SheetData _sheetData;
    private readonly WorkbookPart? _workbookPart;
    private readonly HashSet<string> _visiting;
    private readonly HashSet<string> _expandingNames = new(StringComparer.OrdinalIgnoreCase);
    private readonly int _depth;
    private readonly string _sheetKey; // used to qualify cell refs for circular detection
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
        : this(sheetData, workbookPart, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0, "") { }

    private FormulaEvaluator(SheetData sheetData, WorkbookPart? workbookPart, HashSet<string> visiting, int depth, string sheetKey)
    {
        _sheetData = sheetData;
        _workbookPart = workbookPart;
        _visiting = visiting;
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
    /// evaluator's coverage grows. Inspired by POI BaseFormulaEvaluator and
    /// ScFormulaCell.MaybeInterpret single-entry pattern.
    /// </summary>
    internal EvalReport EvaluateForReport(string formula)
    {
        var r = TryEvaluateFull(formula);
        if (r == null) return new EvalReport(EvalReportStatus.NotEvaluated, null);
        if (r.IsError) return new EvalReport(EvalReportStatus.Error, r);
        return new EvalReport(EvalReportStatus.Evaluated, r);
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

    private enum TT { Number, String, CellRef, Range, Op, LParen, RParen, Comma, Func, Bool, Compare, SheetCellRef, SheetRange, ArrayLit, Error }
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

                if (stripped.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) { tokens.Add(new Token(TT.Bool, "TRUE")); continue; }
                if (stripped.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) { tokens.Add(new Token(TT.Bool, "FALSE")); continue; }

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

                if (i < formula.Length && formula[i] == '(' && !IsCellRef(stripped))
                { tokens.Add(new Token(TT.Func, word.Replace(".", "_").ToUpperInvariant())); continue; }

                if (IsCellRef(stripped)) { tokens.Add(new Token(TT.CellRef, stripped.ToUpperInvariant())); continue; }

                // Defined name. Two flavors:
                //   1. Literal range/cellref body — emit a single ref token
                //      (e.g. `StageTable` → `Data!A2:B7`).
                //   2. Formula body (OFFSET(...), INDIRECT(...), arithmetic) —
                //      inline the body's tokens here so the parent expression
                //      evaluates them in place, matching POI's
                //      `evaluateNameFormula` recursion.
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

                throw new NotSupportedException($"Unknown: {word}");
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
            // 0/1 doubles (not Bool) so downstream `*1` stays in numeric domain
            // matching POI's TwoOperandNumericOperation+ArrayFunction policy.
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
        // Lift to per-element FormulaResult arrays so CompareValues sees
        // proper typed cells (string vs number) instead of collapsed doubles.
        var la = AsResultArray(left); var ra = AsResultArray(right);
        int n = Math.Max(la?.Length ?? 1, ra?.Length ?? 1);
        var o = new double[n];
        for (int i = 0; i < n; i++)
        {
            var l = la != null ? (i < la.Length ? la[i] : null) : left;
            var r = ra != null ? (i < ra.Length ? ra[i] : null) : right;
            if (l == null || r == null) { o[i] = 0; continue; }
            var cmp = CompareValues(l, r);
            o[i] = op switch
            {
                "=" => cmp == 0 ? 1 : 0,
                "<>" => cmp != 0 ? 1 : 0,
                "<" => cmp < 0 ? 1 : 0,
                ">" => cmp > 0 ? 1 : 0,
                "<=" => cmp <= 0 ? 1 : 0,
                ">=" => cmp >= 0 ? 1 : 0,
                _ => 0
            };
        }
        return FormulaResult.Array(o);
    }

    private static FormulaResult?[]? AsResultArray(FormulaResult r)
    {
        if (r.IsArray) return r.ArrayValue!.Select(x => (FormulaResult?)FormulaResult.Number(x)).ToArray();
        if (r.IsRange) return r.RangeValue!.ToFlatResults();
        return null;
    }

    private FormulaResult? ParseConcat(List<Token> t, ref int p)
    {
        var left = ParseAddSub(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "&")
        { p++; var right = ParseAddSub(t, ref p); if (right == null) return null;
          if (left.IsError) return left; if (right.IsError) return right;
          left = FormulaResult.Str(left.AsString() + right.AsString()); }
        return left;
    }

    private FormulaResult? ParseAddSub(List<Token> t, ref int p)
    {
        var left = ParseMulDiv(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value is "+" or "-")
        { var op = t[p].Value; p++; var r = ParseMulDiv(t, ref p); if (r == null) return null;
          if (left.IsError) return left; if (r.IsError) return r;
          left = ApplyBinaryOp(left, r, op == "+" ? (a, b) => a + b : (a, b) => a - b); }
        return left;
    }

    private FormulaResult? ParseMulDiv(List<Token> t, ref int p)
    {
        var left = ParsePower(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value is "*" or "/")
        { var op = t[p].Value; p++; var r = ParsePower(t, ref p); if (r == null) return null;
          if (left.IsError) return left; if (r.IsError) return r;
          if (op == "/")
          {
              // Scalar-only div-by-zero gate. For array divisors, any zero produces
              // +Inf rather than #DIV/0! — acceptable degradation; tighten if needed.
              if (!HasArrayShape(r) && r.AsNumber() == 0) return FormulaResult.Error("#DIV/0!");
              left = ApplyBinaryOp(left, r, (a, b) => b == 0 ? double.PositiveInfinity : a / b);
          }
          else
              left = ApplyBinaryOp(left, r, (a, b) => a * b);
        }
        return left;
    }

    private FormulaResult? ParsePower(List<Token> t, ref int p)
    {
        var b = ParseUnary(t, ref p); if (b == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "^")
        { p++; var e = ParseUnary(t, ref p); if (e == null) return null;
          if (b.IsError) return b; if (e.IsError) return e;
          b = ApplyBinaryOp(b, e, Math.Pow); }
        return b;
    }

    // Element-wise application of a binary numeric op. Handles scalar+scalar,
    // array+scalar, scalar+array, array+array. Range operands are flattened
    // row-major (empties treated as 0, matching Excel implicit-zero coercion).
    // Length mismatch in array+array uses Min(len) — Excel would emit #N/A, but
    // min-length is more lenient and only affects malformed inputs.
    private static FormulaResult ApplyBinaryOp(FormulaResult left, FormulaResult right, Func<double, double, double> op)
    {
        var la = AsArrayLike(left); var ra = AsArrayLike(right);
        if (la == null && ra == null) return FormulaResult.Number(op(left.AsNumber(), right.AsNumber()));
        if (la != null && ra == null) { var rn = right.AsNumber(); var o = new double[la.Length]; for (int i = 0; i < la.Length; i++) o[i] = op(la[i], rn); return FormulaResult.Array(o); }
        if (la == null && ra != null) { var ln = left.AsNumber(); var o = new double[ra.Length]; for (int i = 0; i < ra.Length; i++) o[i] = op(ln, ra[i]); return FormulaResult.Array(o); }
        var n = Math.Min(la!.Length, ra!.Length); var oo = new double[n];
        for (int i = 0; i < n; i++) oo[i] = op(la[i], ra[i]);
        return FormulaResult.Array(oo);
    }

    private static bool HasArrayShape(FormulaResult r) => r.IsArray || r.IsRange;

    // Parse the body of an array constant `{...}` (without the braces).
    // Rows are separated by ';', columns by ',' — per ECMA-376 §18.17.7.282.
    // Each cell is a number / "string" / TRUE / FALSE. Produces a RangeData
    // wrapped as Area so ApplyBinaryOp and aggregate functions handle it
    // identically to a real range. BaseRow/BaseCol stay 0 (not a workbook
    // reference) — see POI CacheAreaEval-vs-AreaEval distinction.
    private static FormulaResult ParseArrayConstant(string body)
    {
        var rows = body.Split(';');
        var rowCells = rows.Select(r => r.Split(',').Select(c => c.Trim()).ToArray()).ToArray();
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
                return FormulaResult.Number(-v.AsNumber()); }
            if (t[p].Value == "+") { p++; return ParseUnary(t, ref p); }
        }
        return ParsePostfix(t, ref p);
    }

    private FormulaResult? ParsePostfix(List<Token> t, ref int p)
    {
        var v = ParseAtom(t, ref p); if (v == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "%") { p++; v = FormulaResult.Number(v.AsNumber() / 100.0); }
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
            case TT.LParen: p++; var inner = ParseExpression(t, ref p); if (p < t.Count && t[p].Type == TT.RParen) p++; return inner;
            case TT.Func: return ParseFunction(t, ref p);
            default: return null;
        }
    }

    private FormulaResult? ParseFunction(List<Token> t, ref int p)
    {
        var name = t[p].Value; p++;
        if (p >= t.Count || t[p].Type != TT.LParen) return null; p++;
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
                else if (argIdx == 0 && name == "OFFSET" && TryParseRefArg(t, ref p) is { } refArg)
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
        if (!_visiting.Add(qualifiedRef)) return FormulaResult.Number(0); // circular ref: use 0 as initial value (matches Excel iterative calc)
        try
        {
            var cell = FindCell(cellRef);
            if (cell == null) return FormulaResult.Number(0);

            // If cell has a formula, always evaluate it (cached values may be stale)
            if (cell.CellFormula?.Text != null)
            {
                try
                {
                    var evaluated = EvaluateFormula(ModernFunctionQualifier.Unqualify(cell.CellFormula.Text));
                    if (evaluated != null) return evaluated;
                }
                catch { /* fall through to cached value */ }
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

            return FormulaResult.Number(0);
        }
        finally { _visiting.Remove(qualifiedRef); }
    }

    /// <summary>
    /// Resolve a cross-sheet cell reference like "SheetName!A1".
    /// Creates a new evaluator for the target sheet and resolves the cell there.
    /// </summary>
    private FormulaResult? ResolveSheetCellResult(string sheetCellRef)
    {
        if (_depth > 20) return FormulaResult.Number(0); // depth guard

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

        // ResolveCellResult will handle circular detection using qualified ref (sheetKey!cellRef)
        var eval = new FormulaEvaluator(sheetData, _workbookPart, _visiting, _depth + 1, sheetName);
        return eval.ResolveCellResult(cellRef);
    }

    /// <summary>
    /// Resolve a sheet name to its SheetData (or return _sheetData for null/empty name).
    /// </summary>
    private SheetData? GetSheetDataFor(string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName)) return _sheetData;
        if (_workbookPart == null) return null;
        try
        {
            var sheet = _workbookPart.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet?.Id?.Value == null) return null;
            var wsPart = (WorksheetPart)_workbookPart.GetPartById(sheet.Id.Value);
            return wsPart.Worksheet?.GetFirstChild<SheetData>();
        }
        catch { return null; }
    }

    /// <summary>
    /// Scan a sheet's populated rows to find min/max row index. Returns (0,0) if empty.
    /// Used to clamp entire-column references like "A:A" to the actual data area.
    /// </summary>
    private static (int minRow, int maxRow) GetPopulatedRowRange(SheetData sheetData)
    {
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
        return maxRow == 0 ? (0, 0) : (minRow, maxRow);
    }

    /// <summary>
    /// Scan a sheet's populated cells to find min/max column index. Returns (0,0) if empty.
    /// Used to clamp entire-row references like "1:1" to the actual data area.
    /// </summary>
    private static (int minCol, int maxCol) GetPopulatedColRange(SheetData sheetData)
    {
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
        return maxCol == 0 ? (0, 0) : (minCol, maxCol);
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
        return new RangeData(cells) { BaseRow = r1, BaseCol = cMin, BaseSheet = sheetPrefix };
    }

    private static (string col, int row) ParseRef(string r)
    {
        var m = Regex.Match(r, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        return m.Success ? (m.Groups[1].Value.ToUpperInvariant(), int.Parse(m.Groups[2].Value)) : ("A", 1);
    }

    private static int ColToIndex(string col) { int r = 0; foreach (var c in col.ToUpperInvariant()) r = r * 26 + (c - 'A' + 1); return r; }
    private static string IndexToCol(int i) { var r = ""; while (i > 0) { i--; r = (char)('A' + i % 26) + r; i /= 26; } return r; }
}
