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
    public double[]? ArrayValue { get; init; }

    public bool IsNumeric => NumericValue.HasValue;
    public bool IsString => StringValue != null;
    public bool IsBool => BoolValue.HasValue;
    public bool IsError => ErrorValue != null;
    public bool IsArray => ArrayValue != null;

    public static FormulaResult Number(double v) => new() { NumericValue = v };
    public static FormulaResult Str(string v) => new() { StringValue = v };
    public static FormulaResult Bool(bool v) => new() { BoolValue = v };
    public static FormulaResult Error(string v) => new() { ErrorValue = v };
    public static FormulaResult Array(double[] v) => new() { ArrayValue = v };

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
    private readonly int _depth;
    private readonly string _sheetKey; // used to qualify cell refs for circular detection
    private Dictionary<string, Cell>? _cellIndex;
    private Dictionary<string, string>? _definedNames;

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
            if (_depth == 0) _visiting.Clear();
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

    private enum TT { Number, String, CellRef, Range, Op, LParen, RParen, Comma, Func, Bool, Compare, SheetCellRef, SheetRange }
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
            if (ch == '\'')
            {
                var si = i + 1;
                var ei = formula.IndexOf('\'', si);
                if (ei > si && ei + 1 < formula.Length && formula[ei + 1] == '!')
                {
                    var sheetName = formula[si..ei];
                    i = ei + 2; // skip '!'
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

                // Defined name (e.g. `StageTable` → `Data!A2:B7`).
                // Resolve to the target range/cell and emit the corresponding token.
                var definedNames = GetDefinedNames();
                if (definedNames.TryGetValue(stripped, out var defRef))
                {
                    var cleaned = StripDollar(defRef).Trim();
                    string? dnSheet = null;
                    var dnCell = cleaned;
                    var dnBang = cleaned.IndexOf('!');
                    if (dnBang > 0)
                    {
                        dnSheet = cleaned[..dnBang].Trim('\'');
                        dnCell = cleaned[(dnBang + 1)..];
                    }
                    if (dnCell.Contains(':'))
                        tokens.Add(new Token(dnSheet != null ? TT.SheetRange : TT.Range,
                            dnSheet != null ? $"{dnSheet}!{dnCell}" : dnCell));
                    else if (IsCellRef(dnCell))
                        tokens.Add(new Token(dnSheet != null ? TT.SheetCellRef : TT.CellRef,
                            dnSheet != null ? $"{dnSheet}!{dnCell.ToUpperInvariant()}" : dnCell.ToUpperInvariant()));
                    else
                        throw new NotSupportedException($"Unknown defined name target: {defRef}");
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
          left = FormulaResult.Number(op == "+" ? left.AsNumber() + r.AsNumber() : left.AsNumber() - r.AsNumber()); }
        return left;
    }

    private FormulaResult? ParseMulDiv(List<Token> t, ref int p)
    {
        var left = ParsePower(t, ref p); if (left == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value is "*" or "/")
        { var op = t[p].Value; p++; var r = ParsePower(t, ref p); if (r == null) return null;
          if (left.IsError) return left; if (r.IsError) return r;
          if (op == "/" && r.AsNumber() == 0) return FormulaResult.Error("#DIV/0!");
          left = FormulaResult.Number(op == "*" ? left.AsNumber() * r.AsNumber() : left.AsNumber() / r.AsNumber()); }
        return left;
    }

    private FormulaResult? ParsePower(List<Token> t, ref int p)
    {
        var b = ParseUnary(t, ref p); if (b == null) return null;
        while (p < t.Count && t[p].Type == TT.Op && t[p].Value == "^")
        { p++; var e = ParseUnary(t, ref p); if (e == null) return null;
          if (b.IsError) return b; if (e.IsError) return e;
          b = FormulaResult.Number(Math.Pow(b.AsNumber(), e.AsNumber())); }
        return b;
    }

    private FormulaResult? ParseUnary(List<Token> t, ref int p)
    {
        if (p < t.Count && t[p].Type == TT.Op)
        {
            if (t[p].Value == "-") { p++; var v = ParseUnary(t, ref p); if (v == null) return null;
                if (v.IsError) return v;
                return v.IsArray ? FormulaResult.Array(v.ArrayValue!.Select(x => -x).ToArray()) : FormulaResult.Number(-v.AsNumber()); }
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
            case TT.Range: p++; return FormulaResult.Number(0);
            case TT.SheetRange: p++; return FormulaResult.Number(0);
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
                if (p < t.Count && t[p].Type is TT.Range or TT.SheetRange) { args.Add(Expand2DRange(t[p].Value)); p++; }
                else { var expr = ParseExpression(t, ref p); if (expr == null) return null; args.Add(expr); }
                if (p >= t.Count || t[p].Type != TT.Comma) break; p++;
            }
        }
        if (p < t.Count && t[p].Type == TT.RParen) p++;
        return EvalFunction(name, args);
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
                    var evaluated = EvaluateFormula(cell.CellFormula.Text);
                    if (evaluated != null) return evaluated;
                }
                catch { /* fall through to cached value */ }
            }

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
        if (sheetData == null) return FormulaResult.Number(0);

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
