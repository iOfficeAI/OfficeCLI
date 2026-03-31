// Copyright 2025 OfficeCli (officecli.ai)
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

    public bool IsNumeric => NumericValue.HasValue;
    public bool IsString => StringValue != null;
    public bool IsBool => BoolValue.HasValue;
    public bool IsError => ErrorValue != null;

    public static FormulaResult Number(double v) => new() { NumericValue = v };
    public static FormulaResult Str(string v) => new() { StringValue = v };
    public static FormulaResult Bool(bool v) => new() { BoolValue = v };
    public static FormulaResult Error(string v) => new() { ErrorValue = v };

    public double AsNumber() => NumericValue ?? (BoolValue == true ? 1 : 0);
    public string AsString() => StringValue ?? NumericValue?.ToString(CultureInfo.InvariantCulture)
        ?? (BoolValue.HasValue ? (BoolValue.Value ? "TRUE" : "FALSE") : ErrorValue ?? "");

    public string ToCellValueText()
    {
        if (NumericValue.HasValue)
        {
            var v = NumericValue.Value;
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
/// 2D range data for lookup functions (VLOOKUP, HLOOKUP, INDEX).
/// </summary>
internal class RangeData
{
    public FormulaResult?[,] Cells { get; }
    public int Rows { get; }
    public int Cols { get; }

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
    private readonly HashSet<string> _visiting = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, Cell>? _cellIndex;

    public FormulaEvaluator(SheetData sheetData, WorkbookPart? workbookPart = null)
    {
        _sheetData = sheetData;
        _workbookPart = workbookPart;
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
            _visiting.Clear();
            return EvaluateFormula(formula);
        }
        catch { return null; }
    }

    private FormulaResult? EvaluateFormula(string formula)
    {
        var tokens = Tokenize(formula);
        var pos = 0;
        var result = ParseExpression(tokens, ref pos);
        return pos == tokens.Count ? result : null;
    }

    // ==================== Tokenizer ====================

    private enum TT { Number, String, CellRef, Range, Op, LParen, RParen, Comma, Func, Bool, Compare }
    private record Token(TT Type, string Value);

    private static List<Token> Tokenize(string formula)
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

            if (char.IsDigit(ch) || ch == '.')
            { var ns = ParseNumber(formula, ref i); if (ns != null) { tokens.Add(new Token(TT.Number, ns)); continue; } }

            if (char.IsLetter(ch) || ch == '_' || ch == '$')
            {
                var start = i;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '$' or '.')) i++;
                var word = formula[start..i]; var stripped = StripDollar(word);

                if (stripped.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) { tokens.Add(new Token(TT.Bool, "TRUE")); continue; }
                if (stripped.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) { tokens.Add(new Token(TT.Bool, "FALSE")); continue; }

                if (i < formula.Length && formula[i] == ':' && IsCellRef(stripped))
                { i++; var s2 = i; while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '$')) i++;
                  tokens.Add(new Token(TT.Range, $"{stripped}:{StripDollar(formula[s2..i])}")); continue; }

                if (i < formula.Length && formula[i] == '(' && !IsCellRef(stripped))
                { tokens.Add(new Token(TT.Func, word.Replace(".", "_").ToUpperInvariant())); continue; }

                if (IsCellRef(stripped)) { tokens.Add(new Token(TT.CellRef, stripped.ToUpperInvariant())); continue; }
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

    // ==================== Recursive Descent Parser ====================

    private FormulaResult? ParseExpression(List<Token> t, ref int p) => ParseComparison(t, ref p);

    private FormulaResult? ParseComparison(List<Token> t, ref int p)
    {
        var left = ParseConcat(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Compare)
        {
            var op = t[p].Value; p++;
            var right = ParseConcat(t, ref p); if (right == null) return null;
            var cmp = CompareValues(left, right);
            left = op switch { "=" => FormulaResult.Bool(cmp == 0), "<>" => FormulaResult.Bool(cmp != 0),
                "<" => FormulaResult.Bool(cmp < 0), ">" => FormulaResult.Bool(cmp > 0),
                "<=" => FormulaResult.Bool(cmp <= 0), ">=" => FormulaResult.Bool(cmp >= 0), _ => null };
            if (left == null) return null;
        }
        return left;
    }

    private FormulaResult? ParseConcat(List<Token> t, ref int p)
    {
        var left = ParseAddSub(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "&")
        { p++; var right = ParseAddSub(t, ref p); if (right == null) return null; left = FormulaResult.Str(left.AsString() + right.AsString()); }
        return left;
    }

    private FormulaResult? ParseAddSub(List<Token> t, ref int p)
    {
        var left = ParseMulDiv(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value is "+" or "-")
        { var op = t[p].Value; p++; var r = ParseMulDiv(t, ref p); if (r == null) return null;
          left = FormulaResult.Number(op == "+" ? left.AsNumber() + r.AsNumber() : left.AsNumber() - r.AsNumber()); }
        return left;
    }

    private FormulaResult? ParseMulDiv(List<Token> t, ref int p)
    {
        var left = ParsePower(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value is "*" or "/")
        { var op = t[p].Value; p++; var r = ParsePower(t, ref p); if (r == null) return null;
          if (op == "/" && r.AsNumber() == 0) return FormulaResult.Error("#DIV/0!");
          left = FormulaResult.Number(op == "*" ? left.AsNumber() * r.AsNumber() : left.AsNumber() / r.AsNumber()); }
        return left;
    }

    private FormulaResult? ParsePower(List<Token> t, ref int p)
    {
        var b = ParseUnary(t, ref p); if (b == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "^")
        { p++; var e = ParseUnary(t, ref p); if (e == null) return null; b = FormulaResult.Number(Math.Pow(b.AsNumber(), e.AsNumber())); }
        return b;
    }

    private FormulaResult? ParseUnary(List<Token> t, ref int p)
    {
        if (p < t.Count && t[p].Type == TT.Op)
        {
            if (t[p].Value == "-") { p++; var v = ParsePostfix(t, ref p); return v == null ? null : FormulaResult.Number(-v.AsNumber()); }
            if (t[p].Value == "+") { p++; return ParsePostfix(t, ref p); }
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
            case TT.Range: p++; return FormulaResult.Number(0);
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
        if (p < t.Count && t[p].Type != TT.RParen)
        {
            while (true)
            {
                if (p < t.Count && t[p].Type == TT.Range) { args.Add(Expand2DRange(t[p].Value)); p++; }
                else { var expr = ParseExpression(t, ref p); if (expr == null) return null; args.Add(expr); }
                if (p >= t.Count || t[p].Type != TT.Comma) break; p++;
            }
        }
        if (p < t.Count && t[p].Type == TT.RParen) p++;
        return EvalFunction(name, args);
    }

    // ==================== Cell & Range Resolution ====================

    private FormulaResult? ResolveCellResult(string cellRef)
    {
        cellRef = StripDollar(cellRef).ToUpperInvariant();
        if (!_visiting.Add(cellRef)) return FormulaResult.Error("#REF!");
        try
        {
            var cell = FindCell(cellRef);
            if (cell == null) return FormulaResult.Number(0);

            var cached = cell.CellValue?.Text;
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
                if (cell.DataType?.Value == CellValues.String || cell.DataType?.Value == CellValues.InlineString) return FormulaResult.Str(cached);
                return double.TryParse(cached, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? FormulaResult.Number(v) : FormulaResult.Str(cached);
            }

            if (cell.CellFormula?.Text != null) return EvaluateFormula(cell.CellFormula.Text);
            return FormulaResult.Number(0);
        }
        finally { _visiting.Remove(cellRef); }
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
        var parts = rangeExpr.Split(':');
        if (parts.Length != 2) return new RangeData(new FormulaResult?[0, 0]);
        var (col1, row1) = ParseRef(StripDollar(parts[0]));
        var (col2, row2) = ParseRef(StripDollar(parts[1]));
        var c1 = ColToIndex(col1); var c2 = ColToIndex(col2);
        var r1 = Math.Min(row1, row2); var r2 = Math.Max(row1, row2);
        var cMin = Math.Min(c1, c2); var cMax = Math.Max(c1, c2);
        var rows = r2 - r1 + 1; var cols = cMax - cMin + 1;
        var cells = new FormulaResult?[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                cells[r, c] = ResolveCellResult($"{IndexToCol(cMin + c)}{r1 + r}");
        return new RangeData(cells);
    }

    private static (string col, int row) ParseRef(string r)
    {
        var m = Regex.Match(r, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        return m.Success ? (m.Groups[1].Value.ToUpperInvariant(), int.Parse(m.Groups[2].Value)) : ("A", 1);
    }

    private static int ColToIndex(string col) { int r = 0; foreach (var c in col.ToUpperInvariant()) r = r * 26 + (c - 'A' + 1); return r; }
    private static string IndexToCol(int i) { var r = ""; while (i > 0) { i--; r = (char)('A' + i % 26) + r; i /= 26; } return r; }
}
