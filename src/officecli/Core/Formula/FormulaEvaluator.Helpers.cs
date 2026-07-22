// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        if (a.IsBool && b.IsBool) return (a.BoolValue!.Value ? 1 : 0).CompareTo(b.BoolValue!.Value ? 1 : 0);
        // Excel cross-type ordering: Number < Text < FALSE < TRUE. Critically,
        // ="1"=1 is FALSE in Excel (text-vs-number never equal) — do NOT coerce
        // via AsNumber here. AsNumber's text→number coercion is for arithmetic
        // operators only; comparison operators preserve type identity.
        int Rank(FormulaResult r) => r.IsNumeric ? 0 : r.IsString ? 1 : r.IsBool ? (r.BoolValue!.Value ? 3 : 2) : 1;
        return Rank(a).CompareTo(Rank(b));
    }

    private static IEnumerable<FormulaResult> ExpandRange(RangeData rd) =>
        Enumerable.Range(0, rd.Rows).SelectMany(r =>
            Enumerable.Range(0, rd.Cols).Select(c => rd.Cells[r, c] ?? FormulaResult.Number(0)));

    // Functions exempt from automatic scalar-error propagation: they classify
    // errors, trap them, take lazy branches, or count/ignore them.
    private static readonly HashSet<string> ErrorTransparentFns = new(StringComparer.OrdinalIgnoreCase)
    {
        "IF", "IFS", "IFERROR", "IFNA", "CHOOSE", "SWITCH",
        "ISERROR", "ISERR", "ISNA", "ISBLANK", "ISNUMBER", "ISTEXT",
        "ISLOGICAL", "ISNONTEXT", "ISREF", "ISFORMULA",
        "ERROR_TYPE", "TYPE", "ISOMITTED", "NA",
        "COUNT", "COUNTA", "COUNTBLANK", "AGGREGATE", "SUBTOTAL",
    };

    private static List<FormulaResult> AllArgs(List<object> args) =>
        args.SelectMany(a => a is RangeData rd ? ExpandRange(rd)
            : a is FormulaResult { IsRange: true } fr ? ExpandRange(fr.RangeValue!)
            : a is FormulaResult { IsArray: true } fa ? fa.ArrayValue!.Select(FormulaResult.Number)
            : a is double[] arr ? arr.Select(v => FormulaResult.Number(v))
            : a is FormulaResult r ? [r] : Enumerable.Empty<FormulaResult>()).ToList();

    /// <summary>Returns the first error found in any RangeData or FormulaResult arg, or null.</summary>
    private static FormulaResult? CheckRangeErrors(List<object> args)
    {
        foreach (var a in args)
        {
            if (a is RangeData rd) { var err = rd.FirstError(); if (err != null) return err; }
            else if (a is FormulaResult { IsRange: true } fr) { var err = fr.RangeValue!.FirstError(); if (err != null) return err; }
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
            else if (a is FormulaResult { IsRange: true } fr) result.AddRange(fr.RangeValue!.ToDoubleArray());
            else if (a is FormulaResult { IsArray: true } fa) result.AddRange(fa.ArrayValue!);
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

    // Excel wildcard pattern -> unanchored .NET regex: * -> .*, ? -> ., with
    // ~*/~? as the literal characters. Shared by SEARCH. The ~ sentinels are
    // swapped out before Regex.Escape — Escape leaves ~ alone, so a
    // post-escape \~\* pattern never matches anything.
    private static string WildcardToRegex(string p) =>
        Regex.Escape(p.Replace("~*", "\x01").Replace("~?", "\x02"))
            .Replace(@"\*", ".*").Replace(@"\?", ".").Replace("\x01", @"\*").Replace("\x02", @"\?");

    // Excel PROPER capitalizes a letter after ANY non-letter (apostrophes and
    // digits included: o'brien -> O'Brien), unlike ToTitleCase's word breaks.
    private static string ProperCase(string s)
    {
        var chars = s.ToLowerInvariant().ToCharArray();
        bool boundary = true;
        for (int i = 0; i < chars.Length; i++)
        {
            if (char.IsLetter(chars[i]))
            {
                if (boundary) chars[i] = char.ToUpperInvariant(chars[i]);
                boundary = false;
            }
            else boundary = true;
        }
        return new string(chars);
    }

    private static double RoundUp(double v, int d) { var f = Math.Pow(10, d); return Math.Ceiling(Math.Abs(v) * f) / f * Math.Sign(v); }
    private static double RoundDown(double v, int d) { var f = Math.Pow(10, d); return Math.Floor(Math.Abs(v) * f) / f * Math.Sign(v); }
    // Excel extracts HOUR/MINUTE/SECOND from a time rounded to the nearest
    // second (so 12:14:58.56 reads 59s, not a truncated 58), carrying into the
    // minute/hour when it rolls over. Rounding the whole serial keeps the three
    // components consistent.
    private static DateTime TimeRoundedToSecond(double serial)
    {
        long ticks = DateTime.FromOADate(serial).Ticks;
        long half = TimeSpan.TicksPerSecond / 2;
        return new DateTime((ticks + half) / TimeSpan.TicksPerSecond * TimeSpan.TicksPerSecond);
    }

    private static double CeilingF(double v, double s) => s == 0 ? 0 : Math.Ceiling(v / s) * s;
    private static double FloorF(double v, double s) => s == 0 ? 0 : Math.Floor(v / s) * s;

    // CEILING.MATH / FLOOR.MATH take an optional mode; significance sign is
    // ignored (magnitude only). For a negative number a nonzero mode flips the
    // rounding direction: CEILING.MATH rounds away from zero, FLOOR.MATH toward
    // zero. Mode is irrelevant for non-negative numbers.
    private static double CeilingMath(double v, double s, double mode)
    {
        if (s == 0) return 0;
        s = Math.Abs(s);
        return v < 0 && mode != 0 ? Math.Floor(v / s) * s : Math.Ceiling(v / s) * s;
    }
    private static double FloorMath(double v, double s, double mode)
    {
        if (s == 0) return 0;
        s = Math.Abs(s);
        return v < 0 && mode != 0 ? Math.Ceiling(v / s) * s : Math.Floor(v / s) * s;
    }
    private static double EvenF(double v) { var c = (int)Math.Ceiling(Math.Abs(v)); return (c % 2 == 0 ? c : c + 1) * Math.Sign(v); }
    private static double OddF(double v) { if (v == 0) return 1; var c = (int)Math.Ceiling(Math.Abs(v)); return (c % 2 == 1 ? c : c + 1) * Math.Sign(v); }
    private static double Factorial(double n) { double r = 1; for (int i = 2; i <= (int)n; i++) r *= i; return r; }
    // n! overflows double past 170, so large arguments go through log-gamma:
    // C(n,k) = exp(lnΓ(n+1) − lnΓ(k+1) − lnΓ(n−k+1)); small ones keep the exact
    // factorial ratio (integer-precise, matches Excel digit-for-digit).
    private static double Combin(int n, int k) =>
        k < 0 || k > n ? 0 :
        n <= 170 ? Factorial(n) / (Factorial(k) * Factorial(n - k)) :
        Math.Exp(GammaLn(n + 1.0) - GammaLn(k + 1.0) - GammaLn(n - k + 1.0));
    private static double Permut(int n, int k) =>
        k < 0 || k > n ? 0 :
        n <= 170 ? Factorial(n) / Factorial(n - k) :
        Math.Exp(GammaLn(n + 1.0) - GammaLn(n - k + 1.0));
    private static long Gcd(long a, long b) { a = Math.Abs(a); b = Math.Abs(b); while (b != 0) { var t = b; b = a % b; a = t; } return a; }
    private static long Lcm(long a, long b) => a == 0 || b == 0 ? 0 : Math.Abs(a / Gcd(a, b) * b);

    // POWER/^: a negative base with a fractional exponent is real only when the
    // exponent is the reciprocal of an odd integer (e.g. a cube root); otherwise
    // NaN, which the serializer surfaces as #NUM!.
    private static double ExcelPow(double b, double e)
    {
        if (b < 0 && e != Math.Floor(e))
        {
            double inv = 1.0 / e, invR = Math.Round(inv);
            if (Math.Abs(inv - invR) < 1e-9 && ((long)invR) % 2 != 0) return -Math.Pow(-b, e);
            return double.NaN;
        }
        return Math.Pow(b, e);
    }

    // MROUND: number and multiple must share a sign (else #NUM!); a zero multiple
    // yields 0. Rounds half away from zero.
    private static FormulaResult MRound(double n, double m)
    {
        if (m == 0) return FR(0);
        if (n != 0 && Math.Sign(n) != Math.Sign(m)) return FormulaResult.Error("#NUM!");
        return FR(Math.Round(n / m, MidpointRounding.AwayFromZero) * m);
    }

    // LEFT/RIGHT: num_chars defaults to 1 (handled by caller); negative errors.
    private static FormulaResult EvalLeftRight(string s, int n, bool left)
    {
        if (n < 0) return FormulaResult.Error("#VALUE!");
        if (n >= s.Length) return FR_S(s);
        return FR_S(left ? s[..n] : s[^n..]);
    }

    // VALUE: parse numbers, thousands-grouped, percent, a leading currency symbol,
    // time-of-day (fraction of a day) and dates (serial).
    private static FormulaResult EvalValue(string s)
    {
        s = s.Trim();
        if (s.Length == 0) return FR(0);
        if (s.EndsWith("%") && double.TryParse(s[..^1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var p))
            return FR(p / 100.0);
        var body = s;
        foreach (var sym in new[] { "$", "¥", "€", "£", "₹" })
            if (body.StartsWith(sym)) { body = body[sym.Length..].Trim(); break; }
        if (double.TryParse(body, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return FR(v);
        if (Regex.IsMatch(s, @"^\d{1,2}:\d{2}(:\d{2})?$") && TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out var ts))
            return FR(ts.TotalDays);
        if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt)) return FR(dt.ToOADate());
        return FormulaResult.Error("#VALUE!");
    }

    // BIN2DEC/OCT2DEC/HEX2DEC: at most 10 digits; the top bit of the fixed field
    // width (10 binary / 30 octal / 40 hex bits) is a sign bit (two's complement).
    // An over-long string or an invalid digit is #NUM!.
    private static FormulaResult FromBase2Dec(string s, int radix)
    {
        s = s.Trim();
        if (s.Length == 0 || s.Length > 10) return FormulaResult.Error("#NUM!");
        long val;
        try { val = Convert.ToInt64(s, radix); }
        catch { return FormulaResult.Error("#NUM!"); }
        long field = radix == 2 ? (1L << 10) : radix == 8 ? (1L << 30) : (1L << 40);
        if (val >= field / 2) val -= field;
        return FR(val);
    }

    // For AGGREGATE ignore-error options: reduce a range/array arg to its
    // non-error numeric values so the delegated aggregate never sees an error.
    private static object StripErrorCells(object a)
    {
        if (a is FormulaResult { IsRange: true } fr)
            return fr.RangeValue!.ToFlatResults().Where(c => c is { IsError: false, IsNumeric: true }).Select(c => c!.AsNumber()).ToArray();
        if (a is RangeData rd)
            return rd.ToFlatResults().Where(c => c is { IsError: false, IsNumeric: true }).Select(c => c!.AsNumber()).ToArray();
        return a;
    }

    private static FormulaResult? EvalModeMult(double[] v)
    {
        if (v.Length == 0) return null;
        int maxCount = v.GroupBy(x => x).Max(g => g.Count());
        if (maxCount <= 1) return FormulaResult.Error("#N/A");
        var tied = v.GroupBy(x => x).Where(g => g.Count() == maxCount).Select(g => g.Key).ToHashSet();
        var seen = new HashSet<double>();
        var ordered = new List<double>();
        foreach (var x in v) if (tied.Contains(x) && seen.Add(x)) ordered.Add(x);
        var col = new FormulaResult?[ordered.Count, 1];
        for (int i = 0; i < ordered.Count; i++) col[i, 0] = FormulaResult.Number(ordered[i]);
        return MakeArea(col);
    }

    private const string BaseDigits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    private static FormulaResult EvalBase(long value, int radix, int minLen)
    {
        if (radix < 2 || radix > 36 || value < 0) return FormulaResult.Error("#NUM!");
        var sb = new System.Text.StringBuilder();
        if (value == 0) sb.Append('0');
        while (value > 0) { sb.Insert(0, BaseDigits[(int)(value % radix)]); value /= radix; }
        var s = sb.ToString();
        return FR_S(minLen > s.Length ? s.PadLeft(minLen, '0') : s);
    }

    private static FormulaResult EvalDecimal(string s, int radix)
    {
        if (radix < 2 || radix > 36) return FormulaResult.Error("#NUM!");
        s = s.Trim().ToUpperInvariant();
        long val = 0;
        foreach (var ch in s)
        {
            int d = BaseDigits.IndexOf(ch);
            if (d < 0 || d >= radix) return FormulaResult.Error("#NUM!");
            val = val * radix + d;
        }
        return FR(val);
    }

    // ---- workday / weekend helpers (NETWORKDAYS / WORKDAY and .INTL) ----
    private static readonly Func<DayOfWeek, bool> DefaultWeekend =
        d => d == DayOfWeek.Saturday || d == DayOfWeek.Sunday;

    // .INTL weekend descriptor: a 7-char string ('1' = weekend, Mon..Sun), a
    // two-day code 1..7, or a one-day code 11..17. Advances holidayArg past it.
    private static Func<DayOfWeek, bool> ParseWeekend(List<object> args, int idx, ref int holidayArg)
    {
        if (idx >= args.Count || args[idx] is not FormulaResult w) { holidayArg = idx; return DefaultWeekend; }
        holidayArg = idx + 1;
        var week = new[] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday };
        if (w.IsString)
        {
            var s = w.AsString();
            if (s.Length == 7)
            {
                var set = new HashSet<DayOfWeek>();
                for (int i = 0; i < 7; i++) if (s[i] == '1') set.Add(week[i]);
                return d => set.Contains(d);
            }
        }
        int code = (int)w.AsNumber();
        if (code >= 1 && code <= 7)
        {
            var first = week[(code + 4) % 7];   // code 1 → Saturday
            var second = week[(code + 5) % 7];  // code 1 → Sunday
            return d => d == first || d == second;
        }
        if (code >= 11 && code <= 17)
        {
            var single = week[(code - 11 + 6) % 7]; // code 11 → Sunday
            return d => d == single;
        }
        return DefaultWeekend;
    }

    private static HashSet<DateTime> CollectHolidays(List<object> args, int startArg)
    {
        var set = new HashSet<DateTime>();
        for (int i = startArg; i < args.Count; i++)
        {
            if (AsResults(args[i]) is { } rd)
            {
                foreach (var c in rd) if (c != null && c.IsNumeric) set.Add(DateTime.FromOADate(c.AsNumber()).Date);
            }
            else if (args[i] is FormulaResult r && !r.IsError) set.Add(DateTime.FromOADate(r.AsNumber()).Date);
        }
        return set;
    }

    // Double factorial n!! = n·(n-2)·…·(2 or 1). By definition 0!! = (-1)!! = 1.
    private static FormulaResult FactDouble(double nd)
    {
        int n = (int)nd;
        if (n < -1) return FormulaResult.Error("#NUM!");
        double r = 1; for (int i = n; i > 1; i -= 2) r *= i; return FR(r);
    }

    // MULTINOMIAL: (Σaᵢ)! / (a₁!·a₂!·…).
    private static double Multinomial(double[] v)
    {
        double sum = 0, denom = 1;
        foreach (var x in v) { sum += x; denom *= Factorial(x); }
        return Factorial(sum) / denom;
    }

    // SERIESSUM(x, n, m, coefs) = Σ coefsᵢ · x^(n + i·m).
    private static FormulaResult EvalSeriesSum(List<object> args)
    {
        if (args.Count < 4) return FormulaResult.Error("#VALUE!");
        double x = (args[0] as FormulaResult)?.AsNumber() ?? 0;
        double n = (args[1] as FormulaResult)?.AsNumber() ?? 0;
        double m = (args[2] as FormulaResult)?.AsNumber() ?? 0;
        var coefs = FlattenNumbers(args.Skip(3).ToList());
        double sum = 0;
        for (int i = 0; i < coefs.Length; i++) sum += coefs[i] * Math.Pow(x, n + i * m);
        return FR(sum);
    }

    private static FormulaResult EvalUnichar(int code)
    {
        if (code < 1 || code > 0x10FFFF || (code >= 0xD800 && code <= 0xDFFF))
            return FormulaResult.Error("#VALUE!");
        return FR_S(char.ConvertFromUtf32(code));
    }

    // Excel-style ROUND: round on the decimal value so decimal-literal midpoints
    // (1.005 → 1.01) go the way a user reading the decimal expects, not the way
    // the binary-float approximation (1.00499999…) would. Falls back to scaled
    // double rounding outside decimal's range or for negative digit counts.
    private static readonly string[] ErrorLiterals =
        { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA" };

    private static string? MatchErrorLiteral(string s, int i)
    {
        foreach (var lit in ErrorLiterals)
            if (string.CompareOrdinal(s, i, lit, 0, lit.Length) == 0) return lit;
        return null;
    }

    private static double ExcelRound(double value, int digits)
    {
        if (double.IsNaN(value) || double.IsInfinity(value)) return value;
        if (Math.Abs(value) < 7.9e28 && digits >= 0 && digits <= 28)
        {
            try { return (double)Math.Round((decimal)value, digits, MidpointRounding.AwayFromZero); }
            catch { }
        }
        double f = Math.Pow(10, digits);
        return Math.Round(value * f, MidpointRounding.AwayFromZero) / f;
    }

    // RANK.AVG — like RANK.EQ but ties share the average of the ranks they span.
    private static FormulaResult? EvalRankAvg(List<object> args)
    {
        if (args.Count < 2) return null;
        var val = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var arr = AsDoubles(args[1]); if (arr == null) return null;
        var order = args.Count > 2 && args[2] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        var sorted = order == 0 ? arr.OrderByDescending(x => x).ToArray() : arr.OrderBy(x => x).ToArray();
        var positions = new List<int>();
        for (int i = 0; i < sorted.Length; i++) if (Math.Abs(sorted[i] - val) < 1e-10) positions.Add(i + 1);
        return positions.Count == 0 ? FormulaResult.Error("#N/A") : FR(positions.Average());
    }

    // PROB — total probability mass whose x value falls in [lower, upper]
    // (upper defaults to lower for a single-point probability).
    private static FormulaResult? EvalProb(List<object> args)
    {
        if (args.Count < 3) return null;
        var x = AsDoubles(args[0]); var p = AsDoubles(args[1]);
        if (x == null || p == null || x.Length != p.Length) return FormulaResult.Error("#N/A");
        double lower = args[2] is FormulaResult r2 ? r2.AsNumber() : 0;
        double upper = args.Count > 3 && args[3] is FormulaResult r3 ? r3.AsNumber() : lower;
        double sum = 0;
        for (int i = 0; i < x.Length; i++) if (x[i] >= lower && x[i] <= upper) sum += p[i];
        return FR(sum);
    }

    private static double[,]? AsMatrix(object? a)
    {
        var rd = AsRangeData(a);
        if (rd == null) return null;
        var m = new double[rd.Rows, rd.Cols];
        for (int i = 0; i < rd.Rows; i++)
            for (int j = 0; j < rd.Cols; j++) m[i, j] = rd.Cells[i, j]?.AsNumber() ?? 0;
        return m;
    }

    private static FormulaResult EvalMdeterm(object? a)
    {
        var m = AsMatrix(a);
        if (m == null || m.GetLength(0) != m.GetLength(1)) return FormulaResult.Error("#VALUE!");
        int n = m.GetLength(0);
        var A = (double[,])m.Clone();
        double det = 1;
        for (int col = 0; col < n; col++)
        {
            int piv = col;
            for (int r = col + 1; r < n; r++) if (Math.Abs(A[r, col]) > Math.Abs(A[piv, col])) piv = r;
            if (Math.Abs(A[piv, col]) < 1e-300) return FR(0);
            if (piv != col) { for (int k = 0; k < n; k++) (A[piv, k], A[col, k]) = (A[col, k], A[piv, k]); det = -det; }
            det *= A[col, col];
            for (int r = col + 1; r < n; r++) { double f = A[r, col] / A[col, col]; for (int k = col; k < n; k++) A[r, k] -= f * A[col, k]; }
        }
        return FR(det);
    }

    private static FormulaResult EvalMmult(object? a, object? b)
    {
        var A = AsMatrix(a); var B = AsMatrix(b);
        if (A == null || B == null) return FormulaResult.Error("#VALUE!");
        int ar = A.GetLength(0), ac = A.GetLength(1), br = B.GetLength(0), bc = B.GetLength(1);
        if (ac != br) return FormulaResult.Error("#VALUE!");
        var outc = new FormulaResult?[ar, bc];
        for (int i = 0; i < ar; i++)
            for (int j = 0; j < bc; j++) { double s = 0; for (int k = 0; k < ac; k++) s += A[i, k] * B[k, j]; outc[i, j] = FormulaResult.Number(s); }
        return MakeArea(outc);
    }

    private static FormulaResult EvalMinverse(object? a)
    {
        var m = AsMatrix(a);
        if (m == null || m.GetLength(0) != m.GetLength(1)) return FormulaResult.Error("#VALUE!");
        int n = m.GetLength(0);
        var A = new double[n, 2 * n];
        for (int i = 0; i < n; i++) { for (int j = 0; j < n; j++) A[i, j] = m[i, j]; A[i, n + i] = 1; }
        for (int col = 0; col < n; col++)
        {
            int piv = col;
            for (int r = col + 1; r < n; r++) if (Math.Abs(A[r, col]) > Math.Abs(A[piv, col])) piv = r;
            if (Math.Abs(A[piv, col]) < 1e-300) return FormulaResult.Error("#NUM!");
            if (piv != col) for (int k = 0; k < 2 * n; k++) (A[piv, k], A[col, k]) = (A[col, k], A[piv, k]);
            double d = A[col, col]; for (int k = 0; k < 2 * n; k++) A[col, k] /= d;
            for (int r = 0; r < n; r++) { if (r == col) continue; double f = A[r, col]; for (int k = 0; k < 2 * n; k++) A[r, k] -= f * A[col, k]; }
        }
        var outc = new FormulaResult?[n, n];
        for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) outc[i, j] = FormulaResult.Number(A[i, n + j]);
        return MakeArea(outc);
    }

    // Wide (double-byte) character test for LENB — CJK, kana, Hangul and
    // fullwidth forms count as 2 bytes.
    private static bool IsWideChar(char c) =>
        (c >= 0x1100 && c <= 0x115F) || (c >= 0x2E80 && c <= 0xA4CF) || (c >= 0xAC00 && c <= 0xD7A3)
        || (c >= 0xF900 && c <= 0xFAFF) || (c >= 0xFE30 && c <= 0xFE4F)
        || (c >= 0xFF00 && c <= 0xFF60) || (c >= 0xFFE0 && c <= 0xFFE6);

    // ASC — fullwidth ASCII / ideographic space fold to their halfwidth forms.
    private static string ToHalfWidth(string s)
    {
        var sb = new System.Text.StringBuilder(s.Length);
        foreach (var c in s)
        {
            if (c >= 0xFF01 && c <= 0xFF5E) sb.Append((char)(c - 0xFEE0));
            else if (c == 0x3000) sb.Append(' ');
            else sb.Append(c);
        }
        return sb.ToString();
    }

    // DEC2BIN/OCT/HEX: minimal digits for non-negatives (optionally padded to
    // `places`); negatives use two's complement across the base's fixed field
    // width (10 binary / 30 octal / 40 hex bits → 10 digits each).
    private static FormulaResult Dec2Base(double numD, int radix, FormulaResult? placesArg)
    {
        long n = (long)numD;
        (long limit, int width, long field) = radix switch
        {
            2 => (512L, 10, 1L << 10),
            8 => (1L << 29, 10, 1L << 30),
            16 => (1L << 39, 10, 1L << 40),
            _ => (0L, 0, 0L)
        };
        if (n < -limit || n > limit - 1) return FormulaResult.Error("#NUM!");
        string s;
        if (n < 0)
        {
            s = Convert.ToString(field + n, radix).ToUpperInvariant();
            if (s.Length > width) s = s[^width..];
        }
        else
        {
            s = Convert.ToString(n, radix).ToUpperInvariant();
            if (placesArg != null) { int p = (int)placesArg.AsNumber(); if (p > s.Length) s = s.PadLeft(p, '0'); }
        }
        return FR_S(s);
    }

    // SUMX2MY2 (Σx²−y²) / SUMX2PY2 (Σx²+y²) / SUMXMY2 (Σ(x−y)²).
    private static FormulaResult EvalSumXY(List<object> args, int mode)
    {
        if (args.Count < 2) return FormulaResult.Error("#N/A");
        var x = FlattenNumbers(new List<object> { args[0] });
        var y = FlattenNumbers(new List<object> { args[1] });
        int n = Math.Min(x.Length, y.Length);
        double sum = 0;
        for (int i = 0; i < n; i++)
            sum += mode switch { 0 => x[i] * x[i] - y[i] * y[i], 1 => x[i] * x[i] + y[i] * y[i], _ => (x[i] - y[i]) * (x[i] - y[i]) };
        return FR(sum);
    }

    // CEILING.PRECISE / ISO.CEILING / FLOOR.PRECISE round to a multiple of the
    // significance's magnitude, always toward +∞ / −∞ regardless of sign.
    private static double CeilingPrecise(double x, double sig) { sig = Math.Abs(sig); return sig == 0 ? 0 : Math.Ceiling(x / sig) * sig; }
    private static double FloorPrecise(double x, double sig) { sig = Math.Abs(sig); return sig == 0 ? 0 : Math.Floor(x / sig) * sig; }

    // AVERAGEA / MAXA / MINA include text (as 0) and booleans (TRUE=1, FALSE=0);
    // truly-empty cells are skipped.
    private static double[] NumsA(List<object> args)
    {
        var list = new List<double>();
        foreach (var a in args)
        {
            var results = AsResults(a);
            if (results != null)
            {
                foreach (var c in results)
                {
                    if (c == null || c.IsBlank || c.IsError) continue;
                    list.Add(c.IsNumeric || c.IsBool ? c.AsNumber() : 0);
                }
            }
            else if (a is FormulaResult r && !r.IsError)
                list.Add(r.IsNumeric || r.IsBool ? r.AsNumber() : 0);
        }
        return list.ToArray();
    }

    // XMATCH — 1-based position of lookup value. matchMode 0 exact, -1/1 next
    // smaller/larger; searchMode -1 scans back-to-front.
    // Excel wildcard match (XLOOKUP/XMATCH match_mode 2, also SEARCH-style):
    // * = any run, ? = one char, ~ escapes the next wildcard. Case-insensitive.
    private static FormulaResult? EvalXmatch(List<object> args)
    {
        if (args.Count < 2 || args[0] is not FormulaResult lookupVal) return null;
        var arr = AsRangeData(args[1]);
        if (arr == null) return FormulaResult.Error("#N/A");
        int matchMode = args.Count >= 3 && args[2] is FormulaResult mm ? (int)mm.AsNumber() : 0;
        int searchMode = args.Count >= 4 && args[3] is FormulaResult sm ? (int)sm.AsNumber() : 1;
        bool isRow = arr.Rows == 1;
        int len = isRow ? arr.Cols : arr.Rows;
        int step = searchMode == -1 ? -1 : 1;
        int start = step == 1 ? 0 : len - 1, end = step == 1 ? len : -1;
        int found = -1, bestApprox = -1;
        double bestDelta = matchMode == -1 ? double.MinValue : double.MaxValue;
        for (int i = start; i != end; i += step)
        {
            var cell = isRow ? arr.Cells[0, i] : arr.Cells[i, 0];
            if (cell == null) continue;
            if (matchMode == 2)
            {
                if (Regex.IsMatch(cell.AsString(), "^" + WildcardToRegex(lookupVal.AsString()) + "$", RegexOptions.IgnoreCase))
                { found = i; break; }
                continue;
            }
            var cmp = CompareValues(cell, lookupVal);
            if (cmp == 0) { found = i; break; }
            if (matchMode == -1 && cmp < 0) { var d = cell.AsNumber() - lookupVal.AsNumber(); if (d > bestDelta) { bestDelta = d; bestApprox = i; } }
            else if (matchMode == 1 && cmp > 0) { var d = cell.AsNumber() - lookupVal.AsNumber(); if (d < bestDelta) { bestDelta = d; bestApprox = i; } }
        }
        if (found < 0) found = bestApprox;
        return found < 0 ? FormulaResult.Error("#N/A") : FR(found + 1);
    }

    // VDB — declining-balance depreciation accumulated over [start, end], with a
    // straight-line switch once that yields more (unless no_switch is set).
    private static FormulaResult? EvalVdb(List<object> args)
    {
        if (args.Count < 5) return null;
        double Num(int i, double d) => i < args.Count && args[i] is FormulaResult r ? r.AsNumber() : d;
        double cost = Num(0, 0), salvage = Num(1, 0), life = Num(2, 0), startP = Num(3, 0), endP = Num(4, 0), factor = Num(5, 2);
        bool noSwitch = args.Count > 6 && args[6] is FormulaResult r6 && r6.AsNumber() != 0;
        if (life <= 0) return FormulaResult.Error("#NUM!");
        double PeriodDep(int p, double book)
        {
            double ddb = Math.Min(book * factor / life, book - salvage);
            if (ddb < 0) ddb = 0;
            if (!noSwitch)
            {
                double remaining = life - (p - 1);
                double sl = remaining > 0 ? (book - salvage) / remaining : 0;
                if (sl > ddb) ddb = Math.Max(sl, 0);
            }
            return ddb;
        }
        double bookVal = cost, sum = 0;
        int lastWhole = (int)Math.Ceiling(endP);
        for (int p = 1; p <= lastWhole; p++)
        {
            double dep = PeriodDep(p, bookVal);
            double frac = Math.Min(endP, p) - Math.Max(startP, p - 1);
            if (frac > 0) sum += dep * frac;
            bookVal -= dep;
        }
        return FR(sum);
    }

    // CONVERT — ratio units share a category and convert through a common base;
    // temperature is affine and converts through Kelvin.
    private static readonly Dictionary<string, (string cat, double f)> _convUnits = new(StringComparer.Ordinal)
    {
        // mass (base gram)
        ["g"] = ("m", 1), ["kg"] = ("m", 1000), ["mg"] = ("m", 0.001), ["lbm"] = ("m", 453.59237),
        ["ozm"] = ("m", 28.349523125), ["u"] = ("m", 1.66053886e-24), ["sg"] = ("m", 14593.9), ["grain"] = ("m", 0.06479891),
        // distance (base metre)
        ["m"] = ("d", 1), ["km"] = ("d", 1000), ["cm"] = ("d", 0.01), ["mm"] = ("d", 0.001),
        ["mi"] = ("d", 1609.344), ["yd"] = ("d", 0.9144), ["ft"] = ("d", 0.3048), ["in"] = ("d", 0.0254),
        ["Nmi"] = ("d", 1852), ["ang"] = ("d", 1e-10), ["ly"] = ("d", 9.4607304725808e15), ["pc"] = ("d", 3.0856775814914e16),
        // time (base second)
        ["sec"] = ("t", 1), ["min"] = ("t", 60), ["mn"] = ("t", 60), ["hr"] = ("t", 3600),
        ["day"] = ("t", 86400), ["yr"] = ("t", 31557600),
        // pressure (base pascal)
        ["Pa"] = ("pr", 1), ["p"] = ("pr", 1), ["atm"] = ("pr", 101325), ["at"] = ("pr", 101325),
        ["mmHg"] = ("pr", 133.322), ["psi"] = ("pr", 6894.75729), ["Torr"] = ("pr", 133.322),
        // force (base newton)
        ["N"] = ("F", 1), ["dyn"] = ("F", 1e-5), ["dy"] = ("F", 1e-5), ["lbf"] = ("F", 4.4482216152605), ["pond"] = ("F", 9.80665e-3),
        // energy (base joule)
        ["J"] = ("e", 1), ["e"] = ("e", 1e-7), ["cal"] = ("e", 4.1868), ["c"] = ("e", 4.184),
        ["eV"] = ("e", 1.602176634e-19), ["ev"] = ("e", 1.602176634e-19), ["Wh"] = ("e", 3600), ["wh"] = ("e", 3600),
        ["BTU"] = ("e", 1055.05585262), ["btu"] = ("e", 1055.05585262), ["HPh"] = ("e", 2684519.537696),
        // power (base watt)
        ["W"] = ("pw", 1), ["w"] = ("pw", 1), ["HP"] = ("pw", 745.69987158227), ["h"] = ("pw", 745.69987158227), ["PS"] = ("pw", 735.49875),
        // volume (base litre)
        ["l"] = ("v", 1), ["L"] = ("v", 1), ["ml"] = ("v", 0.001), ["m3"] = ("v", 1000),
        ["gal"] = ("v", 3.785411784), ["qt"] = ("v", 0.946352946), ["pt"] = ("v", 0.473176473), ["us_pt"] = ("v", 0.473176473),
        ["cup"] = ("v", 0.2365882365), ["oz"] = ("v", 0.0295735295625), ["tsp"] = ("v", 0.00492892159375), ["tbs"] = ("v", 0.01478676478125),
    };

    private static FormulaResult EvalConvert(double value, string from, string to)
    {
        if (TempToKelvin(from, value, out var k) && KelvinToTemp(to, k, out var tv)) return FR(tv);
        if (_convUnits.TryGetValue(from, out var a) && _convUnits.TryGetValue(to, out var b) && a.cat == b.cat)
            return FR(value * a.f / b.f);
        return FormulaResult.Error("#N/A");
    }

    private static bool TempToKelvin(string u, double v, out double k)
    {
        switch (u) { case "C": case "Cel": k = v + 273.15; return true;
            case "F": case "Fah": k = (v - 32) * 5.0 / 9.0 + 273.15; return true;
            case "K": case "Kel": k = v; return true; default: k = 0; return false; }
    }
    private static bool KelvinToTemp(string u, double k, out double v)
    {
        switch (u) { case "C": case "Cel": v = k - 273.15; return true;
            case "F": case "Fah": v = (k - 273.15) * 9.0 / 5.0 + 32; return true;
            case "K": case "Kel": v = k; return true; default: v = 0; return false; }
    }

    // WEEKNUM return types: 1 (default, week starts Sunday), 2/11 (Monday),
    // 12–17 (Tue–Sun), 21 (ISO-8601). Week 1 is the week containing Jan 1.
    private static FormulaResult EvalWeekNum(double serial, int returnType)
    {
        var d = DateTime.FromOADate(serial);
        if (returnType == 21)
            return FR(CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
                d, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday));
        int startDow = returnType switch
        {
            1 => 0, 2 or 11 => 1, 12 => 2, 13 => 3, 14 => 4, 15 => 5, 16 => 6, 17 => 0, _ => 0
        };
        var jan1 = new DateTime(d.Year, 1, 1);
        int jan1Offset = (((int)jan1.DayOfWeek) - startDow + 7) % 7;
        int daysFromJan1 = (d - jan1).Days;
        return FR((daysFromJan1 + jan1Offset) / 7 + 1);
    }

    // DAYS360: day count on a 360-day year. US/NASD method by default; the
    // European flag adjusts day-31 handling on both endpoints.
    private static FormulaResult EvalDays360(double startSerial, double endSerial, bool european)
    {
        var s = DateTime.FromOADate(startSerial);
        var e = DateTime.FromOADate(endSerial);
        int d1 = s.Day, d2 = e.Day;
        if (european)
        {
            if (d1 == 31) d1 = 30;
            if (d2 == 31) d2 = 30;
        }
        else
        {
            if (d1 == 31) d1 = 30;
            if (d2 == 31 && d1 == 30) d2 = 30;
        }
        return FR((e.Year - s.Year) * 360 + (e.Month - s.Month) * 30 + (d2 - d1));
    }

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
