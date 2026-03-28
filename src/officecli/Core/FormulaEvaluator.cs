// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Lightweight Excel formula evaluator for computing cached values at write time.
/// Supports: arithmetic (+,-,*,/), parentheses, cell references (A1), range references (A1:A10),
/// and functions: SUM, AVERAGE, COUNT, MIN, MAX, IF, ROUND, ABS.
/// Unsupported formulas return null (caller should leave CellValue empty).
/// </summary>
internal class FormulaEvaluator
{
    private readonly SheetData _sheetData;
    private readonly WorkbookPart? _workbookPart;
    private readonly HashSet<string> _visiting = new(StringComparer.OrdinalIgnoreCase);

    public FormulaEvaluator(SheetData sheetData, WorkbookPart? workbookPart = null)
    {
        _sheetData = sheetData;
        _workbookPart = workbookPart;
    }

    /// <summary>
    /// Try to evaluate a formula string. Returns the numeric result, or null if unsupported.
    /// </summary>
    public double? TryEvaluate(string formula)
    {
        try
        {
            _visiting.Clear();
            var tokens = Tokenize(formula);
            var pos = 0;
            var result = ParseExpression(tokens, ref pos);
            return pos == tokens.Count ? result : null;
        }
        catch
        {
            return null;
        }
    }

    // ==================== Tokenizer ====================

    private enum TokenType { Number, CellRef, Range, Operator, LParen, RParen, Comma, Function }

    private record Token(TokenType Type, string Value);

    private static List<Token> Tokenize(string formula)
    {
        var tokens = new List<Token>();
        var i = 0;
        formula = formula.Trim();

        while (i < formula.Length)
        {
            var ch = formula[i];

            if (char.IsWhiteSpace(ch)) { i++; continue; }

            // Operators
            if (ch is '+' or '-' or '*' or '/' or '^')
            {
                // Unary minus/plus: at start, after operator, after '(' or ','
                if ((ch is '-' or '+') && (tokens.Count == 0 ||
                    tokens[^1].Type is TokenType.Operator or TokenType.LParen or TokenType.Comma))
                {
                    // Parse as part of a number
                    var numStr = ParseNumber(formula, ref i);
                    if (numStr != null) { tokens.Add(new Token(TokenType.Number, numStr)); continue; }
                }
                tokens.Add(new Token(TokenType.Operator, ch.ToString()));
                i++;
                continue;
            }

            if (ch == '(') { tokens.Add(new Token(TokenType.LParen, "(")); i++; continue; }
            if (ch == ')') { tokens.Add(new Token(TokenType.RParen, ")")); i++; continue; }
            if (ch == ',') { tokens.Add(new Token(TokenType.Comma, ",")); i++; continue; }

            // Number
            if (char.IsDigit(ch) || ch == '.')
            {
                var numStr = ParseNumber(formula, ref i);
                if (numStr != null) { tokens.Add(new Token(TokenType.Number, numStr)); continue; }
            }

            // Function or cell reference
            if (char.IsLetter(ch) || ch == '_')
            {
                var start = i;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '_' || formula[i] == '$'))
                    i++;

                var word = formula[start..i];

                // Check for range: A1:B10
                if (i < formula.Length && formula[i] == ':' && IsCellRef(word))
                {
                    i++; // skip ':'
                    var start2 = i;
                    while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '$'))
                        i++;
                    var word2 = formula[start2..i];
                    tokens.Add(new Token(TokenType.Range, $"{word}:{word2}"));
                    continue;
                }

                // Function: followed by '('
                if (i < formula.Length && formula[i] == '(' && !IsCellRef(word))
                {
                    tokens.Add(new Token(TokenType.Function, word.ToUpperInvariant()));
                    continue;
                }

                // Cell reference
                if (IsCellRef(word))
                {
                    tokens.Add(new Token(TokenType.CellRef, StripDollar(word)));
                    continue;
                }

                // Unknown identifier
                throw new NotSupportedException($"Unknown: {word}");
            }

            // String literal (skip)
            if (ch == '"')
            {
                i++;
                while (i < formula.Length && formula[i] != '"') i++;
                if (i < formula.Length) i++;
                throw new NotSupportedException("String formulas not supported");
            }

            throw new NotSupportedException($"Unexpected char: {ch}");
        }

        return tokens;
    }

    private static string? ParseNumber(string s, ref int i)
    {
        var start = i;
        if (i < s.Length && (s[i] == '-' || s[i] == '+')) i++;
        var hasDigits = false;
        while (i < s.Length && char.IsDigit(s[i])) { i++; hasDigits = true; }
        if (i < s.Length && s[i] == '.')
        {
            i++;
            while (i < s.Length && char.IsDigit(s[i])) { i++; hasDigits = true; }
        }
        if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
        {
            i++;
            if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
            while (i < s.Length && char.IsDigit(s[i])) i++;
        }
        if (!hasDigits) { i = start; return null; }
        return s[start..i];
    }

    private static bool IsCellRef(string s)
    {
        s = StripDollar(s);
        return Regex.IsMatch(s, @"^[A-Z]{1,3}\d+$", RegexOptions.IgnoreCase);
    }

    private static string StripDollar(string s) => s.Replace("$", "");

    // ==================== Recursive Descent Parser ====================
    // Grammar:
    //   expression  = term (('+' | '-') term)*
    //   term        = power (('*' | '/') power)*
    //   power       = unary ('^' unary)*
    //   unary       = ('-' | '+')? atom
    //   atom        = number | cellRef | range | function '(' args ')' | '(' expression ')'
    //   args        = (expression | range) (',' (expression | range))*

    private double? ParseExpression(List<Token> tokens, ref int pos)
    {
        var left = ParseTerm(tokens, ref pos);
        if (left == null) return null;

        while (pos < tokens.Count && tokens[pos].Type == TokenType.Operator
            && tokens[pos].Value is "+" or "-")
        {
            var op = tokens[pos].Value;
            pos++;
            var right = ParseTerm(tokens, ref pos);
            if (right == null) return null;
            left = op == "+" ? left + right : left - right;
        }
        return left;
    }

    private double? ParseTerm(List<Token> tokens, ref int pos)
    {
        var left = ParsePower(tokens, ref pos);
        if (left == null) return null;

        while (pos < tokens.Count && tokens[pos].Type == TokenType.Operator
            && tokens[pos].Value is "*" or "/")
        {
            var op = tokens[pos].Value;
            pos++;
            var right = ParsePower(tokens, ref pos);
            if (right == null) return null;
            left = op == "*" ? left * right : (right != 0 ? left / right : null);
            if (left == null) return null;
        }
        return left;
    }

    private double? ParsePower(List<Token> tokens, ref int pos)
    {
        var baseVal = ParseUnary(tokens, ref pos);
        if (baseVal == null) return null;

        while (pos < tokens.Count && tokens[pos].Type == TokenType.Operator
            && tokens[pos].Value == "^")
        {
            pos++;
            var exp = ParseUnary(tokens, ref pos);
            if (exp == null) return null;
            baseVal = Math.Pow(baseVal.Value, exp.Value);
        }
        return baseVal;
    }

    private double? ParseUnary(List<Token> tokens, ref int pos)
    {
        if (pos < tokens.Count && tokens[pos].Type == TokenType.Operator && tokens[pos].Value == "-")
        {
            pos++;
            var val = ParseAtom(tokens, ref pos);
            return val == null ? null : -val;
        }
        if (pos < tokens.Count && tokens[pos].Type == TokenType.Operator && tokens[pos].Value == "+")
        {
            pos++;
            return ParseAtom(tokens, ref pos);
        }
        return ParseAtom(tokens, ref pos);
    }

    private double? ParseAtom(List<Token> tokens, ref int pos)
    {
        if (pos >= tokens.Count) return null;
        var token = tokens[pos];

        switch (token.Type)
        {
            case TokenType.Number:
                pos++;
                return double.TryParse(token.Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var num) ? num : null;

            case TokenType.CellRef:
                pos++;
                return ResolveCellValue(token.Value);

            case TokenType.LParen:
                pos++; // skip '('
                var inner = ParseExpression(tokens, ref pos);
                if (pos >= tokens.Count || tokens[pos].Type != TokenType.RParen) return null;
                pos++; // skip ')'
                return inner;

            case TokenType.Function:
                return ParseFunction(tokens, ref pos);

            default:
                return null;
        }
    }

    // ==================== Function Evaluation ====================

    private double? ParseFunction(List<Token> tokens, ref int pos)
    {
        var funcName = tokens[pos].Value;
        pos++; // skip function name

        if (pos >= tokens.Count || tokens[pos].Type != TokenType.LParen) return null;
        pos++; // skip '('

        var args = ParseFunctionArgs(tokens, ref pos);
        if (args == null) return null;

        if (pos >= tokens.Count || tokens[pos].Type != TokenType.RParen) return null;
        pos++; // skip ')'

        return EvalFunction(funcName, args);
    }

    private List<object>? ParseFunctionArgs(List<Token> tokens, ref int pos)
    {
        var args = new List<object>();
        if (pos < tokens.Count && tokens[pos].Type == TokenType.RParen)
            return args; // no arguments

        while (true)
        {
            // Check for range argument
            if (pos < tokens.Count && tokens[pos].Type == TokenType.Range)
            {
                args.Add(ExpandRange(tokens[pos].Value));
                pos++;
            }
            else
            {
                var val = ParseExpression(tokens, ref pos);
                if (val == null) return null;
                args.Add(val.Value);
            }

            if (pos >= tokens.Count || tokens[pos].Type != TokenType.Comma) break;
            pos++; // skip ','
        }
        return args;
    }

    private double? EvalFunction(string name, List<object> args)
    {
        var values = FlattenArgs(args);

        return name switch
        {
            "SUM" => values.Length > 0 ? values.Sum() : 0,
            "AVERAGE" => values.Length > 0 ? values.Average() : null,
            "COUNT" => values.Length,
            "COUNTA" => values.Length,
            "MIN" => values.Length > 0 ? values.Min() : null,
            "MAX" => values.Length > 0 ? values.Max() : null,
            "ABS" => values.Length == 1 ? Math.Abs(values[0]) : null,
            "ROUND" => EvalRound(args),
            "ROUNDUP" => EvalRoundUp(args),
            "ROUNDDOWN" => EvalRoundDown(args),
            "INT" => values.Length == 1 ? Math.Floor(values[0]) : null,
            "MOD" => values.Length == 2 && values[1] != 0 ? values[0] % values[1] : null,
            "POWER" => values.Length == 2 ? Math.Pow(values[0], values[1]) : null,
            "SQRT" => values.Length == 1 && values[0] >= 0 ? Math.Sqrt(values[0]) : null,
            "IF" => EvalIf(args),
            _ => null // unsupported function
        };
    }

    private double? EvalIf(List<object> args)
    {
        if (args.Count < 2) return null;
        var condition = args[0] is double d ? d : null as double?;
        if (condition == null) return null;
        // Non-zero = true
        if (condition != 0)
            return args[1] is double t ? t : null;
        else
            return args.Count >= 3 && args[2] is double f ? f : 0;
    }

    private static double? EvalRound(List<object> args)
    {
        var vals = FlattenArgs(args);
        if (vals.Length < 2) return vals.Length == 1 ? Math.Round(vals[0]) : null;
        return Math.Round(vals[0], (int)vals[1]);
    }

    private static double? EvalRoundUp(List<object> args)
    {
        var vals = FlattenArgs(args);
        if (vals.Length < 2) return null;
        var factor = Math.Pow(10, (int)vals[1]);
        return Math.Ceiling(vals[0] * factor) / factor;
    }

    private static double? EvalRoundDown(List<object> args)
    {
        var vals = FlattenArgs(args);
        if (vals.Length < 2) return null;
        var factor = Math.Pow(10, (int)vals[1]);
        return Math.Floor(vals[0] * factor) / factor;
    }

    private static double[] FlattenArgs(List<object> args)
    {
        var result = new List<double>();
        foreach (var arg in args)
        {
            if (arg is double d) result.Add(d);
            else if (arg is double[] arr) result.AddRange(arr);
        }
        return result.ToArray();
    }

    // ==================== Cell & Range Resolution ====================

    private double? ResolveCellValue(string cellRef)
    {
        cellRef = StripDollar(cellRef).ToUpperInvariant();

        // Circular reference detection
        if (!_visiting.Add(cellRef)) return null;
        try
        {
            var cell = FindCell(cellRef);
            if (cell == null) return 0; // empty cell = 0

            // If cell has a cached value, use it
            var cachedText = cell.CellValue?.Text;
            if (!string.IsNullOrEmpty(cachedText))
            {
                // Shared string → not numeric
                if (cell.DataType?.Value == CellValues.SharedString) return null;
                return double.TryParse(cachedText, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;
            }

            // If cell has a formula but no cached value, evaluate recursively
            if (cell.CellFormula?.Text != null)
            {
                return TryEvaluate(cell.CellFormula.Text);
            }

            return 0; // empty cell
        }
        finally
        {
            _visiting.Remove(cellRef);
        }
    }

    private Cell? FindCell(string cellRef)
    {
        foreach (var row in _sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                    return cell;
            }
        }
        return null;
    }

    private double[] ExpandRange(string rangeExpr)
    {
        var parts = rangeExpr.Split(':');
        if (parts.Length != 2) return [];

        var (col1, row1) = ParseRef(StripDollar(parts[0]));
        var (col2, row2) = ParseRef(StripDollar(parts[1]));

        var c1 = ColToIndex(col1);
        var c2 = ColToIndex(col2);
        var r1 = Math.Min(row1, row2);
        var r2 = Math.Max(row1, row2);
        var cMin = Math.Min(c1, c2);
        var cMax = Math.Max(c1, c2);

        var values = new List<double>();
        for (int r = r1; r <= r2; r++)
        {
            for (int c = cMin; c <= cMax; c++)
            {
                var ref_ = $"{IndexToCol(c)}{r}";
                var val = ResolveCellValue(ref_);
                if (val.HasValue) values.Add(val.Value);
            }
        }
        return values.ToArray();
    }

    private static (string col, int row) ParseRef(string cellRef)
    {
        var match = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        return match.Success ? (match.Groups[1].Value.ToUpperInvariant(), int.Parse(match.Groups[2].Value)) : ("A", 1);
    }

    private static int ColToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    private static string IndexToCol(int index)
    {
        var result = "";
        while (index > 0) { index--; result = (char)('A' + index % 26) + result; index /= 26; }
        return result;
    }
}
