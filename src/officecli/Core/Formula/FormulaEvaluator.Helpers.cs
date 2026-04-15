// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Shorthand constructors ====================
    private static FormulaResult FR(double v) => FormulaResult.Number(v);
    private static FormulaResult FR_S(string v) => FormulaResult.Str(v);
    private static FormulaResult FR_B(bool v) => FormulaResult.Bool(v);

    // ==================== Comparison ====================

    private static int CompareValues(FormulaResult a, FormulaResult b)
    {
        if (a.IsNumeric && b.IsNumeric) return a.NumericValue!.Value.CompareTo(b.NumericValue!.Value);
        if (a.IsString && b.IsString) return string.Compare(a.StringValue, b.StringValue, StringComparison.OrdinalIgnoreCase);
        return a.AsNumber().CompareTo(b.AsNumber());
    }

    private static List<FormulaResult> AllArgs(List<object> args) =>
        args.SelectMany(a => a is RangeData rd ? Enumerable.Range(0, rd.Rows).SelectMany(r => Enumerable.Range(0, rd.Cols).Select(c => rd.Cells[r, c] ?? FormulaResult.Number(0)))
            : a is double[] arr ? arr.Select(v => FormulaResult.Number(v))
            : a is FormulaResult r ? [r] : Enumerable.Empty<FormulaResult>()).ToList();

    /// <summary>Returns the first error found in any RangeData or FormulaResult arg, or null.</summary>
    private static FormulaResult? CheckRangeErrors(List<object> args)
    {
        foreach (var a in args)
        {
            if (a is RangeData rd) { var err = rd.FirstError(); if (err != null) return err; }
            else if (a is FormulaResult { IsError: true } e) return e;
        }
        return null;
    }

    private static double[] FlattenNumbers(List<object> args)
    {
        var result = new List<double>();
        foreach (var a in args)
        {
            if (a is RangeData rd) result.AddRange(rd.ToDoubleArray());
            else if (a is double[] arr) result.AddRange(arr);
            else if (a is FormulaResult { IsNumeric: true } r) result.Add(r.NumericValue!.Value);
            else if (a is FormulaResult { IsBool: true } rb) result.Add(rb.BoolValue!.Value ? 1 : 0);
        }
        return result.ToArray();
    }

    // ==================== Criteria matching (for SUMIF, COUNTIF, etc.) ====================

    private static bool MatchesCriteria(double value, string criteria)
        => MatchesCriteria(FormulaResult.Number(value), criteria);

    private static bool MatchesCriteria(FormulaResult? cellValue, string criteria)
    {
        criteria = criteria.Trim();
        if (string.IsNullOrEmpty(criteria)) return true;

        // Numeric comparison operators
        double numVal = cellValue?.AsNumber() ?? 0;
        if (criteria.StartsWith(">=") && double.TryParse(criteria[2..], NumberStyles.Any, CultureInfo.InvariantCulture, out var ge)) return numVal >= ge;
        if (criteria.StartsWith("<=") && double.TryParse(criteria[2..], NumberStyles.Any, CultureInfo.InvariantCulture, out var le)) return numVal <= le;
        if (criteria.StartsWith("<>"))
        {
            var operand = criteria[2..];
            if (double.TryParse(operand, NumberStyles.Any, CultureInfo.InvariantCulture, out var ne)) return Math.Abs(numVal - ne) > 1e-10;
            // String not-equal
            return !string.Equals(cellValue?.AsString() ?? "", operand, StringComparison.OrdinalIgnoreCase);
        }
        if (criteria.StartsWith(">") && double.TryParse(criteria[1..], NumberStyles.Any, CultureInfo.InvariantCulture, out var gt)) return numVal > gt;
        if (criteria.StartsWith("<") && double.TryParse(criteria[1..], NumberStyles.Any, CultureInfo.InvariantCulture, out var lt)) return numVal < lt;
        if (criteria.StartsWith("="))
        {
            var operand = criteria[1..];
            if (double.TryParse(operand, NumberStyles.Any, CultureInfo.InvariantCulture, out var eq)) return Math.Abs(numVal - eq) < 1e-10;
            // String equality after =
            return string.Equals(cellValue?.AsString() ?? "", operand, StringComparison.OrdinalIgnoreCase);
        }
        if (double.TryParse(criteria, NumberStyles.Any, CultureInfo.InvariantCulture, out var plain)) return Math.Abs(numVal - plain) < 1e-10;

        // Wildcard / string matching
        string cellStr = cellValue?.AsString() ?? "";
        if (criteria.Contains('*') || criteria.Contains('?'))
        {
            // Convert Excel wildcards to regex: * -> .*, ? -> ., ~* -> literal *, ~? -> literal ?
            var pattern = Regex.Escape(criteria).Replace(@"\~\*", "\x01").Replace(@"\~\?", "\x02")
                .Replace(@"\*", ".*").Replace(@"\?", ".").Replace("\x01", @"\*").Replace("\x02", @"\?");
            return Regex.IsMatch(cellStr, "^" + pattern + "$", RegexOptions.IgnoreCase);
        }

        // Plain string equality
        return string.Equals(cellStr, criteria, StringComparison.OrdinalIgnoreCase);
    }

    // ==================== Math utilities ====================

    private static double RoundUp(double v, int d) { var f = Math.Pow(10, d); return Math.Ceiling(Math.Abs(v) * f) / f * Math.Sign(v); }
    private static double RoundDown(double v, int d) { var f = Math.Pow(10, d); return Math.Floor(Math.Abs(v) * f) / f * Math.Sign(v); }
    private static double CeilingF(double v, double s) => s == 0 ? 0 : Math.Ceiling(v / s) * s;
    private static double FloorF(double v, double s) => s == 0 ? 0 : Math.Floor(v / s) * s;
    private static double EvenF(double v) { var c = (int)Math.Ceiling(Math.Abs(v)); return (c % 2 == 0 ? c : c + 1) * Math.Sign(v); }
    private static double OddF(double v) { var c = (int)Math.Ceiling(Math.Abs(v)); return (c % 2 == 1 ? c : c + 1) * Math.Sign(v); }
    private static double Factorial(double n) { double r = 1; for (int i = 2; i <= (int)n; i++) r *= i; return r; }
    private static double Combin(int n, int k) => k < 0 || k > n ? 0 : Factorial(n) / (Factorial(k) * Factorial(n - k));
    private static double Permut(int n, int k) => k < 0 || k > n ? 0 : Factorial(n) / Factorial(n - k);
    private static long Gcd(long a, long b) { a = Math.Abs(a); b = Math.Abs(b); while (b != 0) { var t = b; b = a % b; a = t; } return a; }
    private static long Lcm(long a, long b) => a == 0 || b == 0 ? 0 : Math.Abs(a / Gcd(a, b) * b);

    private static string ToRoman(int n)
    {
        var vals = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var syms = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var sb = new StringBuilder();
        for (int i = 0; i < vals.Length; i++) while (n >= vals[i]) { sb.Append(syms[i]); n -= vals[i]; }
        return sb.ToString();
    }

    private static double FromRoman(string s)
    {
        var map = new Dictionary<char, int> { ['M'] = 1000, ['D'] = 500, ['C'] = 100, ['L'] = 50, ['X'] = 10, ['V'] = 5, ['I'] = 1 };
        double result = 0;
        for (int i = 0; i < s.Length; i++)
        {
            var val = map.GetValueOrDefault(char.ToUpper(s[i]));
            if (i + 1 < s.Length && val < map.GetValueOrDefault(char.ToUpper(s[i + 1]))) result -= val;
            else result += val;
        }
        return result;
    }
}
