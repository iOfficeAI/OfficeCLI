// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Function Dispatch (150+ functions) ====================

    private FormulaResult? EvalFunction(string name, List<object> args)
    {
        double[] nums() => FlattenNumbers(args);
        FormulaResult? arg(int i) => i < args.Count && args[i] is FormulaResult r ? r : null;
        double num(int i) => arg(i)?.AsNumber() ?? 0;
        string str(int i) => arg(i)?.AsString() ?? "";

        return name switch
        {
            // ===== Math & Aggregation =====
            "SUM" => CheckRangeErrors(args) ?? FR(nums().Sum()),
            "SUMPRODUCT" => EvalSumProduct(args),
            "AVERAGE" => nums() is { Length: > 0 } a ? FR(a.Average()) : null,
            "COUNT" => FR(nums().Length),
            "COUNTA" => FR(args.Sum(a => a is RangeData rd ? rd.ToFlatResults().Count(c => c != null && !c.IsError && c.AsString() != "")
                : a is FormulaResult r && !r.IsError && r.AsString() != "" ? 1 : a is double[] arr ? arr.Length : 0)),
            "COUNTBLANK" => FR(0),
            "MIN" => nums() is { Length: > 0 } mn ? FR(mn.Min()) : FR(0),
            "MAX" => nums() is { Length: > 0 } mx ? FR(mx.Max()) : FR(0),
            "ABS" => FR(Math.Abs(num(0))),
            "SIGN" => FR(Math.Sign(num(0))),
            "INT" => FR(Math.Floor(num(0))),
            "TRUNC" => args.Count >= 2 ? FR(Math.Truncate(num(0) * Math.Pow(10, num(1))) / Math.Pow(10, num(1))) : FR(Math.Truncate(num(0))),
            "ROUND" => FR(Math.Round(num(0), (int)num(1), MidpointRounding.AwayFromZero)),
            "ROUNDUP" => FR(RoundUp(num(0), (int)num(1))),
            "ROUNDDOWN" => FR(RoundDown(num(0), (int)num(1))),
            "CEILING" or "CEILING_MATH" => FR(CeilingF(num(0), args.Count >= 2 ? num(1) : 1)),
            "FLOOR" or "FLOOR_MATH" => FR(FloorF(num(0), args.Count >= 2 ? num(1) : 1)),
            "MOD" => num(1) != 0 ? FR(num(0) - num(1) * Math.Floor(num(0) / num(1))) : FormulaResult.Error("#DIV/0!"),
            "POWER" => FR(Math.Pow(num(0), num(1))),
            "SQRT" => num(0) >= 0 ? FR(Math.Sqrt(num(0))) : FormulaResult.Error("#NUM!"),
            "FACT" => FR(Factorial(num(0))),
            "COMBIN" => FR(Combin((int)num(0), (int)num(1))),
            "PERMUT" => FR(Permut((int)num(0), (int)num(1))),
            "GCD" => FR(nums().Aggregate(0.0, (a, b) => Gcd((long)a, (long)b))),
            "LCM" => FR(nums().Aggregate(1.0, (a, b) => Lcm((long)a, (long)b))),
            "RAND" => FR(new Random().NextDouble()),
            "RANDBETWEEN" => FR(new Random().Next((int)num(0), (int)num(1) + 1)),
            "EVEN" => FR(EvenF(num(0))),
            "ODD" => FR(OddF(num(0))),
            "PRODUCT" => FR(nums().Aggregate(1.0, (a, b) => a * b)),
            "QUOTIENT" => num(1) != 0 ? FR(Math.Truncate(num(0) / num(1))) : FormulaResult.Error("#DIV/0!"),
            "MROUND" => num(1) != 0 ? FR(Math.Round(num(0) / num(1)) * num(1)) : FormulaResult.Error("#NUM!"),
            "ROMAN" => FR_S(ToRoman((int)num(0))),
            "ARABIC" => FR(FromRoman(str(0))),
            "BASE" => FR_S(Convert.ToString((long)num(0), (int)num(1)).ToUpperInvariant()),
            "DECIMAL" => FR(Convert.ToInt64(str(0), (int)num(1))),
            "LOG" => args.Count >= 2 ? FR(Math.Log(num(0), num(1))) : FR(Math.Log10(num(0))),
            "LOG10" => FR(Math.Log10(num(0))),
            "LN" => FR(Math.Log(num(0))),
            "EXP" => FR(Math.Exp(num(0))),

            // ===== Trigonometry =====
            "PI" => FR(Math.PI),
            "SIN" => FR(Math.Sin(num(0))), "COS" => FR(Math.Cos(num(0))), "TAN" => FR(Math.Tan(num(0))),
            "ASIN" => FR(Math.Asin(num(0))), "ACOS" => FR(Math.Acos(num(0))), "ATAN" => FR(Math.Atan(num(0))),
            "ATAN2" => FR(Math.Atan2(num(0), num(1))),
            "SINH" => FR(Math.Sinh(num(0))), "COSH" => FR(Math.Cosh(num(0))), "TANH" => FR(Math.Tanh(num(0))),
            "ASINH" => FR(Math.Asinh(num(0))), "ACOSH" => FR(Math.Acosh(num(0))), "ATANH" => FR(Math.Atanh(num(0))),
            "DEGREES" => FR(num(0) * 180.0 / Math.PI),
            "RADIANS" => FR(num(0) * Math.PI / 180.0),

            // ===== Statistical =====
            "MEDIAN" => EvalMedian(nums()),
            "MODE" or "MODE_SNGL" => EvalMode(nums()),
            "LARGE" => EvalLarge(args), "SMALL" => EvalSmall(args),
            "RANK" or "RANK_EQ" => EvalRank(args),
            "PERCENTILE" or "PERCENTILE_INC" => EvalPercentile(args),
            "PERCENTRANK" or "PERCENTRANK_INC" => EvalPercentRank(args),
            "STDEV" or "STDEV_S" => EvalStdev(nums(), true),
            "STDEVP" or "STDEV_P" => EvalStdev(nums(), false),
            "VAR" or "VAR_S" => EvalVar(nums(), true),
            "VARP" or "VAR_P" => EvalVar(nums(), false),
            "GEOMEAN" => nums() is { Length: > 0 } gm ? FR(Math.Pow(gm.Aggregate(1.0, (a, b) => a * b), 1.0 / gm.Length)) : null,
            "HARMEAN" => nums() is { Length: > 0 } hm ? FR(hm.Length / hm.Sum(x => 1.0 / x)) : null,

            // ===== Logical =====
            "IF" => EvalIf(args), "IFS" => EvalIfs(args),
            "AND" => FR_B(AllArgs(args).All(r => r.AsNumber() != 0)),
            "OR" => FR_B(AllArgs(args).Any(r => r.AsNumber() != 0)),
            "NOT" => FR_B(num(0) == 0),
            "XOR" => FR_B(AllArgs(args).Count(r => r.AsNumber() != 0) % 2 == 1),
            "TRUE" => FR_B(true), "FALSE" => FR_B(false),
            "IFERROR" or "IFNA" => arg(0) is { IsError: true } ? arg(1) : arg(0),
            "SWITCH" => EvalSwitch(args), "CHOOSE" => EvalChoose(args),

            // ===== Text =====
            "CONCATENATE" or "CONCAT" => FR_S(string.Concat(AllArgs(args).Select(r => r.AsString()))),
            "TEXTJOIN" => EvalTextJoin(args),
            "LEFT" => FR_S(str(0).Length >= (int)num(1) ? str(0)[..(int)num(1)] : str(0)),
            "RIGHT" => FR_S(str(0).Length >= (int)num(1) ? str(0)[^(int)num(1)..] : str(0)),
            "MID" => EvalMid(args),
            "LEN" => FR(str(0).Length),
            "TRIM" => FR_S(Regex.Replace(str(0).Trim(), @"\s+", " ")),
            "CLEAN" => FR_S(Regex.Replace(str(0), @"[\x00-\x1F]", "")),
            "UPPER" => FR_S(str(0).ToUpperInvariant()),
            "LOWER" => FR_S(str(0).ToLowerInvariant()),
            "PROPER" => FR_S(CultureInfo.InvariantCulture.TextInfo.ToTitleCase(str(0).ToLowerInvariant())),
            "REPT" => FR_S(string.Concat(Enumerable.Repeat(str(0), (int)num(1)))),
            "CHAR" => FR_S(((char)(int)num(0)).ToString()),
            "CODE" => FR(str(0).Length > 0 ? (int)str(0)[0] : 0),
            "FIND" => EvalFind(args, true), "SEARCH" => EvalFind(args, false),
            "REPLACE" => EvalReplace(args), "SUBSTITUTE" => EvalSubstitute(args),
            "EXACT" => FR_B(str(0) == str(1)),
            "VALUE" => double.TryParse(str(0), NumberStyles.Any, CultureInfo.InvariantCulture, out var pv) ? FR(pv) : FormulaResult.Error("#VALUE!"),
            "TEXT" => EvalText(args),
            "T" => arg(0) is { IsString: true } ? arg(0) : FR_S(""),
            "N" => FR(num(0)),
            "FIXED" => EvalFixed(args),
            "NUMBERVALUE" => EvalNumberValue(args),
            "DOLLAR" or "YEN" => FR_S(num(0).ToString("C", CultureInfo.InvariantCulture)),

            // ===== Lookup & Reference =====
            "INDEX" => EvalIndex(args), "MATCH" => EvalMatch(args),
            "ROW" => EvalRowCol(args, true), "COLUMN" => EvalRowCol(args, false),
            "ROWS" => EvalRowsCols(args, true), "COLUMNS" => EvalRowsCols(args, false),
            "ADDRESS" => EvalAddress(args),
            "VLOOKUP" => EvalVlookup(args),
            "HLOOKUP" => EvalHlookup(args),
            "LOOKUP" => EvalLookup(args),
            "XLOOKUP" => EvalXlookup(args),
            "HYPERLINK" => FR_S(args.Count >= 2 && args[1] is FormulaResult fn ? fn.AsString() : str(0)),
            "OFFSET" or "INDIRECT" => null, // unsupported (requires first-class reference values)

            // ===== Date & Time =====
            "TODAY" => FR(DateTime.Today.ToOADate()), "NOW" => FR(DateTime.Now.ToOADate()),
            "DATE" => FR(new DateTime((int)num(0), (int)num(1), (int)num(2)).ToOADate()),
            "YEAR" => FR(DateTime.FromOADate(num(0)).Year), "MONTH" => FR(DateTime.FromOADate(num(0)).Month),
            "DAY" => FR(DateTime.FromOADate(num(0)).Day), "HOUR" => FR(DateTime.FromOADate(num(0)).Hour),
            "MINUTE" => FR(DateTime.FromOADate(num(0)).Minute), "SECOND" => FR(DateTime.FromOADate(num(0)).Second),
            "WEEKDAY" => FR((int)DateTime.FromOADate(num(0)).DayOfWeek + 1),
            "DATEVALUE" => DateTime.TryParse(str(0), out var dv) ? FR(dv.ToOADate()) : FormulaResult.Error("#VALUE!"),
            "TIMEVALUE" => DateTime.TryParse(str(0), out var tv) ? FR(tv.TimeOfDay.TotalDays) : FormulaResult.Error("#VALUE!"),
            "EDATE" => FR(DateTime.FromOADate(num(0)).AddMonths((int)num(1)).ToOADate()),
            "EOMONTH" => EvalEomonth(args),
            "DAYS" => FR(num(0) - num(1)),
            "DATEDIF" => EvalDateDif(args),
            "NETWORKDAYS" or "NETWORKDAYS_INTL" => EvalNetworkDays(args),
            "WORKDAY" or "WORKDAY_INTL" => EvalWorkDay(args),
            "ISOWEEKNUM" => FR(CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.FromOADate(num(0)), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)),
            "YEARFRAC" => EvalYearFrac(args),

            // ===== Info =====
            "ISNUMBER" => FR_B(arg(0)?.IsNumeric == true),
            "ISTEXT" => FR_B(arg(0)?.IsString == true),
            "ISBLANK" => FR_B(arg(0) == null || (arg(0)?.AsString() == "" && !arg(0)!.IsNumeric)),
            "ISERROR" or "ISERR" => args.Count > 0 && args[0] is RangeData rd_err
                ? FormulaResult.Array(rd_err.ToFlatResults().Select(r => r?.IsError == true ? 1.0 : 0.0).ToArray())
                : FR_B(arg(0)?.IsError == true),
            "ISNA" => FR_B(arg(0)?.ErrorValue == "#N/A"),
            "ISLOGICAL" => FR_B(arg(0)?.IsBool == true),
            "ISEVEN" => FR_B((int)num(0) % 2 == 0), "ISODD" => FR_B((int)num(0) % 2 != 0),
            "ISNONTEXT" => FR_B(arg(0)?.IsString != true),
            "TYPE" => FR(arg(0) switch { { IsNumeric: true } => 1, { IsString: true } => 2, { IsBool: true } => 4, { IsError: true } => 16, _ => 1 }),
            "NA" => FormulaResult.Error("#N/A"),
            "ERROR_TYPE" => FR(arg(0)?.ErrorValue switch { "#NULL!" => 1, "#DIV/0!" => 2, "#VALUE!" => 3, "#REF!" => 4, "#NAME?" => 5, "#NUM!" => 6, "#N/A" => 7, _ => 0 }),

            // ===== Conditional Aggregation =====
            "SUMIF" => EvalSumIf(args), "SUMIFS" => EvalSumIfs(args),
            "COUNTIF" => EvalCountIf(args), "COUNTIFS" => EvalCountIfs(args),
            "AVERAGEIF" => EvalAverageIf(args), "AVERAGEIFS" => EvalAverageIfs(args),
            "MAXIFS" => EvalMaxMinIfs(args, true), "MINIFS" => EvalMaxMinIfs(args, false),

            // ===== Financial =====
            "PMT" => EvalPmt(args), "FV" => EvalFv(args), "PV" => EvalPv(args), "NPER" => EvalNper(args),
            "NPV" => EvalNpv(args), "IPMT" => EvalIpmt(args), "PPMT" => EvalPpmt(args),
            "SLN" => args.Count >= 3 ? FR((num(0) - num(1)) / num(2)) : null,
            "SYD" => EvalSyd(args), "DB" => EvalDb(args), "DDB" => EvalDdb(args),
            "RATE" or "IRR" => null, // iterative solvers — unsupported

            // ===== Conversion =====
            "BIN2DEC" => FR(Convert.ToInt64(str(0), 2)),
            "DEC2BIN" => FR_S(Convert.ToString((long)num(0), 2)),
            "HEX2DEC" => FR(Convert.ToInt64(str(0), 16)),
            "DEC2HEX" => FR_S(Convert.ToString((long)num(0), 16).ToUpperInvariant()),
            "OCT2DEC" => FR(Convert.ToInt64(str(0), 8)),
            "DEC2OCT" => FR_S(Convert.ToString((long)num(0), 8)),
            "BIN2HEX" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 2), 16).ToUpperInvariant()),
            "BIN2OCT" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 2), 8)),
            "HEX2BIN" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 16), 2)),
            "HEX2OCT" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 16), 8)),
            "OCT2BIN" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 8), 2)),
            "OCT2HEX" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 8), 16).ToUpperInvariant()),

            _ => null
        };
    }

    // ==================== Logical ====================

    private FormulaResult? EvalIf(List<object> args)
    {
        var c = args.Count > 0 && args[0] is FormulaResult r ? r : null; if (c == null) return null;
        var isTrue = c.IsNumeric ? c.NumericValue != 0 : c.BoolValue == true;
        if (isTrue) return args.Count > 1 && args[1] is FormulaResult t ? t : FR(0);
        return args.Count > 2 && args[2] is FormulaResult f ? f : FR_B(false);
    }

    private FormulaResult? EvalIfs(List<object> args)
    {
        for (int i = 0; i + 1 < args.Count; i += 2)
        { var c = args[i] is FormulaResult r ? r : null; if (c != null && c.AsNumber() != 0) return args[i + 1] is FormulaResult v ? v : null; }
        return FormulaResult.Error("#N/A");
    }

    private FormulaResult? EvalSwitch(List<object> args)
    {
        if (args.Count < 2) return null;
        var val = args[0] is FormulaResult r ? r : null; if (val == null) return null;
        for (int i = 1; i + 1 < args.Count; i += 2)
        { var cv = args[i] is FormulaResult c ? c : null; if (cv != null && CompareValues(val, cv) == 0) return args[i + 1] is FormulaResult res ? res : null; }
        return args.Count % 2 == 0 ? (args[^1] is FormulaResult def ? def : null) : FormulaResult.Error("#N/A");
    }

    private FormulaResult? EvalChoose(List<object> args)
    {
        if (args.Count < 2) return null;
        var idx = (int)(args[0] is FormulaResult r ? r.AsNumber() : 0);
        return idx >= 1 && idx < args.Count && args[idx] is FormulaResult v ? v : FormulaResult.Error("#VALUE!");
    }

    // ==================== Text ====================

    private FormulaResult? EvalMid(List<object> args)
    {
        var s = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        var start = args.Count > 1 && args[1] is FormulaResult r2 ? (int)r2.AsNumber() - 1 : 0;
        var len = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() : 0;
        if (start < 0 || start >= s.Length) return FR_S("");
        return FR_S(s.Substring(start, Math.Min(len, s.Length - start)));
    }

    private FormulaResult? EvalFind(List<object> args, bool caseSensitive)
    {
        var find = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        var within = args.Count > 1 && args[1] is FormulaResult r2 ? r2.AsString() : "";
        var startPos = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() - 1 : 0;
        var idx = within.IndexOf(find, startPos, caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
        return idx >= 0 ? FR(idx + 1) : FormulaResult.Error("#VALUE!");
    }

    private FormulaResult? EvalReplace(List<object> args)
    {
        var s = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        var start = args.Count > 1 && args[1] is FormulaResult r2 ? (int)r2.AsNumber() - 1 : 0;
        var len = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() : 0;
        var rep = args.Count > 3 && args[3] is FormulaResult r4 ? r4.AsString() : "";
        if (start < 0 || start > s.Length) return FormulaResult.Error("#VALUE!");
        return FR_S(s[..start] + rep + s[Math.Min(start + len, s.Length)..]);
    }

    private FormulaResult? EvalSubstitute(List<object> args)
    {
        var s = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        var old = args.Count > 1 && args[1] is FormulaResult r2 ? r2.AsString() : "";
        var neo = args.Count > 2 && args[2] is FormulaResult r3 ? r3.AsString() : "";
        if (args.Count > 3 && args[3] is FormulaResult r4)
        {
            var n = (int)r4.AsNumber(); var idx = -1;
            for (int i = 0; i < n; i++) { idx = s.IndexOf(old, idx + 1, StringComparison.Ordinal); if (idx < 0) return FR_S(s); }
            return FR_S(s[..idx] + neo + s[(idx + old.Length)..]);
        }
        return FR_S(s.Replace(old, neo));
    }

    private FormulaResult? EvalText(List<object> args)
    {
        var val = args.Count > 0 && args[0] is FormulaResult r ? r.AsNumber() : 0;
        var fmt = args.Count > 1 && args[1] is FormulaResult r2 ? r2.AsString() : "0";
        try { return FR_S(val.ToString(fmt.Replace("#", "0"), CultureInfo.InvariantCulture)); }
        catch { return FR_S(val.ToString(CultureInfo.InvariantCulture)); }
    }

    private static FormulaResult? EvalFixed(List<object> args)
    {
        var v = args.Count > 0 && args[0] is FormulaResult r ? r.AsNumber() : 0;
        var d = args.Count > 1 && args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 2;
        return FR_S(v.ToString($"N{d}", CultureInfo.InvariantCulture));
    }

    private static FormulaResult? EvalNumberValue(List<object> args)
    {
        var s = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        s = s.Replace(",", "").Replace(" ", "").Trim();
        return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? FR(v) : FormulaResult.Error("#VALUE!");
    }

    private FormulaResult? EvalTextJoin(List<object> args)
    {
        if (args.Count < 3) return null;
        var delim = args[0] is FormulaResult r ? r.AsString() : "";
        var ignoreEmpty = args[1] is FormulaResult r2 && r2.AsNumber() != 0;
        var parts = new List<string>();
        for (int i = 2; i < args.Count; i++)
        {
            if (args[i] is RangeData rd2)
            {
                for (int row = 0; row < rd2.Rows; row++)
                    for (int col = 0; col < rd2.Cols; col++)
                    {
                        var cv = rd2.Cells[row, col];
                        if (cv != null) { var s = cv.AsString(); if (!ignoreEmpty || s != "") parts.Add(s); }
                    }
            }
            else if (args[i] is double[] arr) foreach (var v in arr) parts.Add(v.ToString(CultureInfo.InvariantCulture));
            else if (args[i] is FormulaResult fr) { var s = fr.AsString(); if (!ignoreEmpty || s != "") parts.Add(s); }
        }
        return FR_S(string.Join(delim, parts));
    }

    // ==================== Lookup ====================

    private FormulaResult? EvalIndex(List<object> args)
    {
        if (args.Count < 2) return null;
        if (args[0] is RangeData rd)
        {
            var rowIdx = args[1] is FormulaResult r ? (int)r.AsNumber() : 0;
            var colIdx = args.Count > 2 && args[2] is FormulaResult c ? (int)c.AsNumber() : 1;
            if (rowIdx < 1 || rowIdx > rd.Rows || colIdx < 1 || colIdx > rd.Cols) return FormulaResult.Error("#REF!");
            return rd.Cells[rowIdx - 1, colIdx - 1] ?? FormulaResult.Number(0);
        }
        if (args[0] is double[] arr)
        {
            var idx = args[1] is FormulaResult r2 ? (int)r2.AsNumber() - 1 : 0;
            return idx >= 0 && idx < arr.Length ? FR(arr[idx]) : FormulaResult.Error("#REF!");
        }
        return null;
    }

    private FormulaResult? EvalMatch(List<object> args)
    {
        if (args.Count < 2) return null;
        var lookup = args[0] is FormulaResult r ? r : null; if (lookup == null) return null;
        if (args[1] is RangeData rd)
        {
            if (rd.Cols == 1) { for (int i = 0; i < rd.Rows; i++) { var cell = rd.Cells[i, 0]; if (cell != null && CompareValues(cell, lookup) == 0) return FR(i + 1); } }
            else if (rd.Rows == 1) { for (int i = 0; i < rd.Cols; i++) { var cell = rd.Cells[0, i]; if (cell != null && CompareValues(cell, lookup) == 0) return FR(i + 1); } }
        }
        if (args[1] is double[] arr)
        { for (int i = 0; i < arr.Length; i++) if (Math.Abs(arr[i] - lookup.AsNumber()) < 1e-10) return FR(i + 1); }
        return FormulaResult.Error("#N/A");
    }

    private FormulaResult? EvalRowCol(List<object> args, bool isRow)
    {
        if (args.Count == 0) return null;
        if (args[0] is FormulaResult r)
        { var m = Regex.Match(r.AsString(), @"([A-Z]+)(\d+)", RegexOptions.IgnoreCase);
          return m.Success ? FR(isRow ? int.Parse(m.Groups[2].Value) : ColToIndex(m.Groups[1].Value)) : null; }
        return null;
    }

    private static FormulaResult? EvalRowsCols(List<object> args, bool isRows)
    {
        if (args.Count > 0 && args[0] is RangeData rd) return FR(isRows ? rd.Rows : rd.Cols);
        if (args.Count > 0 && args[0] is double[] arr) return FR(arr.Length);
        return FR(1);
    }

    private FormulaResult? EvalVlookup(List<object> args)
    {
        if (args.Count < 3) return null;
        var lookupVal = args[0] is FormulaResult r ? r : null; if (lookupVal == null) return null;
        var table = args[1] is RangeData rd ? rd : null; if (table == null) return FormulaResult.Error("#N/A");
        var colIndex = args[2] is FormulaResult ci ? (int)ci.AsNumber() : 0;
        if (colIndex < 1 || colIndex > table.Cols) return FormulaResult.Error("#REF!");
        var exactMatch = args.Count > 3 && args[3] is FormulaResult rm && (rm.AsNumber() == 0 || rm.AsString().Equals("FALSE", StringComparison.OrdinalIgnoreCase));

        int foundRow = -1;
        if (exactMatch)
        { for (int i = 0; i < table.Rows; i++) { var cell = table.Cells[i, 0]; if (cell != null && CompareValues(cell, lookupVal) == 0) { foundRow = i; break; } } }
        else
        { for (int i = 0; i < table.Rows; i++) { var cell = table.Cells[i, 0]; if (cell == null) continue; if (CompareValues(cell, lookupVal) <= 0) foundRow = i; else break; } }

        return foundRow >= 0 ? (table.Cells[foundRow, colIndex - 1] ?? FormulaResult.Number(0)) : FormulaResult.Error("#N/A");
    }

    private FormulaResult? EvalHlookup(List<object> args)
    {
        if (args.Count < 3) return null;
        var lookupVal = args[0] is FormulaResult r ? r : null; if (lookupVal == null) return null;
        var table = args[1] is RangeData rd ? rd : null; if (table == null) return FormulaResult.Error("#N/A");
        var rowIndex = args[2] is FormulaResult ri ? (int)ri.AsNumber() : 0;
        if (rowIndex < 1 || rowIndex > table.Rows) return FormulaResult.Error("#REF!");
        var exactMatch = args.Count > 3 && args[3] is FormulaResult rm && (rm.AsNumber() == 0 || rm.AsString().Equals("FALSE", StringComparison.OrdinalIgnoreCase));

        int foundCol = -1;
        if (exactMatch)
        { for (int i = 0; i < table.Cols; i++) { var cell = table.Cells[0, i]; if (cell != null && CompareValues(cell, lookupVal) == 0) { foundCol = i; break; } } }
        else
        { for (int i = 0; i < table.Cols; i++) { var cell = table.Cells[0, i]; if (cell == null) continue; if (CompareValues(cell, lookupVal) <= 0) foundCol = i; else break; } }

        return foundCol >= 0 ? (table.Cells[rowIndex - 1, foundCol] ?? FormulaResult.Number(0)) : FormulaResult.Error("#N/A");
    }

    // LOOKUP(lookup_value, lookup_vector, [result_vector])
    // LOOKUP(lookup_value, array)
    // Legacy approximate-match lookup. Assumes lookup_vector is sorted ascending.
    // Array form: searches first row if wider than tall (HLOOKUP-like, returns last row);
    // otherwise searches first column (VLOOKUP-like, returns last column).
    private FormulaResult? EvalLookup(List<object> args)
    {
        if (args.Count < 2) return null;
        var lookupVal = args[0] is FormulaResult r ? r : null;
        if (lookupVal == null) return null;
        var lv = args[1] is RangeData rd ? rd : null;
        if (lv == null) return FormulaResult.Error("#N/A");

        // Vector form (1D): optionally with a parallel result_vector
        if (lv.Rows == 1 || lv.Cols == 1)
        {
            int found = ApproximateMatchVector(lv, lookupVal);
            if (found < 0) return FormulaResult.Error("#N/A");

            var resultVec = args.Count >= 3 && args[2] is RangeData rv ? rv : lv;
            if (resultVec.Rows == 1 && found < resultVec.Cols)
                return resultVec.Cells[0, found] ?? FormulaResult.Number(0);
            if (resultVec.Cols == 1 && found < resultVec.Rows)
                return resultVec.Cells[found, 0] ?? FormulaResult.Number(0);
            return FormulaResult.Error("#N/A");
        }

        // Array form: 2D — search first row or first column depending on orientation
        if (lv.Cols > lv.Rows)
        {
            int foundCol = -1;
            for (int c = 0; c < lv.Cols; c++)
            {
                var cell = lv.Cells[0, c];
                if (cell == null) continue;
                if (CompareValues(cell, lookupVal) <= 0) foundCol = c;
                else break;
            }
            return foundCol >= 0
                ? (lv.Cells[lv.Rows - 1, foundCol] ?? FormulaResult.Number(0))
                : FormulaResult.Error("#N/A");
        }
        else
        {
            int foundRow = -1;
            for (int rr = 0; rr < lv.Rows; rr++)
            {
                var cell = lv.Cells[rr, 0];
                if (cell == null) continue;
                if (CompareValues(cell, lookupVal) <= 0) foundRow = rr;
                else break;
            }
            return foundRow >= 0
                ? (lv.Cells[foundRow, lv.Cols - 1] ?? FormulaResult.Number(0))
                : FormulaResult.Error("#N/A");
        }
    }

    private int ApproximateMatchVector(RangeData rd, FormulaResult lookupVal)
    {
        int found = -1;
        if (rd.Rows == 1)
        {
            for (int c = 0; c < rd.Cols; c++)
            {
                var cell = rd.Cells[0, c];
                if (cell == null) continue;
                if (CompareValues(cell, lookupVal) <= 0) found = c;
                else break;
            }
        }
        else
        {
            for (int rr = 0; rr < rd.Rows; rr++)
            {
                var cell = rd.Cells[rr, 0];
                if (cell == null) continue;
                if (CompareValues(cell, lookupVal) <= 0) found = rr;
                else break;
            }
        }
        return found;
    }

    // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
    // match_mode: 0=exact (default), -1=exact or next smaller, 1=exact or next larger, 2=wildcard (NYI — treated as exact)
    // search_mode: 1=first to last (default), -1=last to first. Binary modes (2/-2) treated as linear.
    private FormulaResult? EvalXlookup(List<object> args)
    {
        if (args.Count < 3) return null;
        var lookupVal = args[0] is FormulaResult r ? r : null;
        if (lookupVal == null) return null;
        var lookupArr = args[1] is RangeData la ? la : null;
        var returnArr = args[2] is RangeData ra ? ra : null;
        if (lookupArr == null || returnArr == null) return FormulaResult.Error("#N/A");

        var ifNotFound = args.Count >= 4 && args[3] is FormulaResult inf ? inf : null;
        var matchMode = args.Count >= 5 && args[4] is FormulaResult mm ? (int)mm.AsNumber() : 0;
        var searchMode = args.Count >= 6 && args[5] is FormulaResult sm ? (int)sm.AsNumber() : 1;

        bool isRow = lookupArr.Rows == 1;
        int len = isRow ? lookupArr.Cols : lookupArr.Rows;
        int step = searchMode == -1 ? -1 : 1;
        int start = step == 1 ? 0 : len - 1;
        int end = step == 1 ? len : -1;

        int found = -1;
        int bestApprox = -1;
        double bestDelta = matchMode == -1 ? double.MinValue : double.MaxValue;

        for (int i = start; i != end; i += step)
        {
            var cell = isRow ? lookupArr.Cells[0, i] : lookupArr.Cells[i, 0];
            if (cell == null) continue;
            var cmp = CompareValues(cell, lookupVal);
            if (cmp == 0) { found = i; break; }
            if (matchMode == -1 && cmp < 0)
            {
                var delta = cell.AsNumber() - lookupVal.AsNumber();
                if (delta > bestDelta) { bestDelta = delta; bestApprox = i; }
            }
            else if (matchMode == 1 && cmp > 0)
            {
                var delta = cell.AsNumber() - lookupVal.AsNumber();
                if (delta < bestDelta) { bestDelta = delta; bestApprox = i; }
            }
        }

        if (found < 0) found = bestApprox;
        if (found < 0) return ifNotFound ?? FormulaResult.Error("#N/A");

        // Pull the value at `found` from return_array (same orientation as lookup_array).
        if (isRow)
        {
            if (found < returnArr.Cols) return returnArr.Cells[0, found] ?? FormulaResult.Number(0);
        }
        else
        {
            if (found < returnArr.Rows) return returnArr.Cells[found, 0] ?? FormulaResult.Number(0);
        }
        return FormulaResult.Error("#N/A");
    }

    private static FormulaResult? EvalAddress(List<object> args)
    {
        if (args.Count < 2) return null;
        var row = (int)(args[0] is FormulaResult r ? r.AsNumber() : 1);
        var col = (int)(args[1] is FormulaResult r2 ? r2.AsNumber() : 1);
        var abs = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() : 1;
        var cs = IndexToCol(col);
        return abs switch { 1 => FR_S($"${cs}${row}"), 2 => FR_S($"{cs}${row}"), 3 => FR_S($"${cs}{row}"), _ => FR_S($"{cs}{row}") };
    }

    // ==================== Statistical ====================

    private static FormulaResult? EvalMedian(double[] v)
    {
        if (v.Length == 0) return null;
        var s = v.OrderBy(x => x).ToArray();
        return FR(s.Length % 2 == 1 ? s[s.Length / 2] : (s[s.Length / 2 - 1] + s[s.Length / 2]) / 2.0);
    }

    private static FormulaResult? EvalMode(double[] v)
    {
        if (v.Length == 0) return null;
        var top = v.GroupBy(x => x).OrderByDescending(g => g.Count()).ThenBy(g => g.Key).First();
        return top.Count() > 1 ? FR(top.Key) : FormulaResult.Error("#N/A");
    }

    private static FormulaResult? EvalLarge(List<object> args)
    {
        var arr = args.Count > 0 && args[0] is double[] a ? a : null;
        var k = args.Count > 1 && args[1] is FormulaResult r ? (int)r.AsNumber() : 1;
        if (arr == null || k < 1 || k > arr.Length) return FormulaResult.Error("#NUM!");
        return FR(arr.OrderByDescending(x => x).ElementAt(k - 1));
    }

    private static FormulaResult? EvalSmall(List<object> args)
    {
        var arr = args.Count > 0 && args[0] is double[] a ? a : null;
        var k = args.Count > 1 && args[1] is FormulaResult r ? (int)r.AsNumber() : 1;
        if (arr == null || k < 1 || k > arr.Length) return FormulaResult.Error("#NUM!");
        return FR(arr.OrderBy(x => x).ElementAt(k - 1));
    }

    private static FormulaResult? EvalRank(List<object> args)
    {
        if (args.Count < 2) return null;
        var val = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var arr = args[1] is double[] a ? a : null; if (arr == null) return null;
        var order = args.Count > 2 && args[2] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        var sorted = order == 0 ? arr.OrderByDescending(x => x).ToArray() : arr.OrderBy(x => x).ToArray();
        for (int i = 0; i < sorted.Length; i++) if (Math.Abs(sorted[i] - val) < 1e-10) return FR(i + 1);
        return FormulaResult.Error("#N/A");
    }

    private static FormulaResult? EvalPercentile(List<object> args)
    {
        var arr = args.Count > 0 && args[0] is double[] a ? a : null;
        var k = args.Count > 1 && args[1] is FormulaResult r ? r.AsNumber() : 0;
        if (arr == null || arr.Length == 0 || k < 0 || k > 1) return FormulaResult.Error("#NUM!");
        var sorted = arr.OrderBy(x => x).ToArray();
        var idx = k * (sorted.Length - 1); var lower = (int)Math.Floor(idx); var upper = Math.Min(lower + 1, sorted.Length - 1);
        return FR(sorted[lower] + (idx - lower) * (sorted[upper] - sorted[lower]));
    }

    private static FormulaResult? EvalPercentRank(List<object> args)
    {
        var arr = args.Count > 0 && args[0] is double[] a ? a : null;
        var val = args.Count > 1 && args[1] is FormulaResult r ? r.AsNumber() : 0;
        if (arr == null || arr.Length == 0) return FormulaResult.Error("#NUM!");
        return FR((double)arr.Count(x => x < val) / (arr.Length - 1));
    }

    private static FormulaResult? EvalStdev(double[] v, bool sample)
    {
        if (v.Length < (sample ? 2 : 1)) return FormulaResult.Error("#DIV/0!");
        var mean = v.Average(); var sumSq = v.Sum(x => (x - mean) * (x - mean));
        return FR(Math.Sqrt(sumSq / (sample ? v.Length - 1 : v.Length)));
    }

    private static FormulaResult? EvalVar(double[] v, bool sample)
    {
        if (v.Length < (sample ? 2 : 1)) return FormulaResult.Error("#DIV/0!");
        var mean = v.Average(); return FR(v.Sum(x => (x - mean) * (x - mean)) / (sample ? v.Length - 1 : v.Length));
    }

    // ==================== Conditional Aggregation ====================

    // Helper: extract double[] from RangeData or double[]
    private static double[]? AsDoubles(object? a) => a is RangeData rd ? rd.ToDoubleArray() : a is double[] arr ? arr : null;

    // Helper: extract FormulaResult?[] from RangeData (preserves string values for criteria matching)
    private static FormulaResult?[]? AsResults(object? a) => a is RangeData rd ? rd.ToFlatResults() : null;

    // Helper: extract numeric value from a FormulaResult (null for non-numeric).
    // Used by conditional aggregation to keep value-range indices aligned with criteria-range indices
    // — AsDoubles/ToDoubleArray collapses non-numerics and shifts indices, which breaks SUMIF/AVERAGEIF alignment.
    private static double? AsNumeric(FormulaResult? v)
    {
        if (v?.IsNumeric == true) return v.NumericValue;
        if (v?.IsBool == true) return v.BoolValue!.Value ? 1 : 0;
        return null;
    }

    private FormulaResult? EvalSumIf(List<object> args)
    {
        if (args.Count < 2) return null;
        var range = AsResults(args[0]); var criteria = args[1] is FormulaResult c ? c.AsString() : "";
        var sumRange = args.Count > 2 ? AsResults(args[2]) : range;
        if (range == null || sumRange == null) return null;
        double sum = 0;
        for (int i = 0; i < range.Length && i < sumRange.Length; i++)
            if (MatchesCriteria(range[i], criteria))
            { var n = AsNumeric(sumRange[i]); if (n.HasValue) sum += n.Value; }
        return FR(sum);
    }

    private FormulaResult? EvalSumIfs(List<object> args)
    {
        if (args.Count < 3) return null;
        var sumRange = AsResults(args[0]); if (sumRange == null) return null;
        double sum = 0;
        for (int i = 0; i < sumRange.Length; i++)
        {
            var match = true;
            for (int c = 1; c + 1 < args.Count; c += 2)
            { var cr = AsResults(args[c]); var crit = args[c + 1] is FormulaResult cv ? cv.AsString() : "";
              if (cr == null || i >= cr.Length || !MatchesCriteria(cr[i], crit)) { match = false; break; } }
            if (match) { var n = AsNumeric(sumRange[i]); if (n.HasValue) sum += n.Value; }
        }
        return FR(sum);
    }

    private FormulaResult? EvalCountIf(List<object> args)
    {
        if (args.Count < 2) return null;
        var range = AsResults(args[0]); var criteria = args[1] is FormulaResult c ? c.AsString() : "";
        return range != null ? FR(range.Count(v => MatchesCriteria(v, criteria))) : null;
    }

    private FormulaResult? EvalCountIfs(List<object> args)
    {
        if (args.Count < 2) return null;
        var first = AsResults(args[0]); if (first == null) return null;
        int count = 0;
        for (int i = 0; i < first.Length; i++)
        {
            var match = true;
            for (int c = 0; c + 1 < args.Count; c += 2)
            { var cr = AsResults(args[c]); var crit = args[c + 1] is FormulaResult cv ? cv.AsString() : "";
              if (cr == null || i >= cr.Length || !MatchesCriteria(cr[i], crit)) { match = false; break; } }
            if (match) count++;
        }
        return FR(count);
    }

    private FormulaResult? EvalAverageIf(List<object> args)
    {
        if (args.Count < 2) return null;
        var range = AsResults(args[0]); var criteria = args[1] is FormulaResult c ? c.AsString() : "";
        var avgRange = args.Count > 2 ? AsResults(args[2]) : range;
        if (range == null || avgRange == null) return null;
        var vals = new List<double>();
        for (int i = 0; i < range.Length && i < avgRange.Length; i++)
            if (MatchesCriteria(range[i], criteria))
            { var n = AsNumeric(avgRange[i]); if (n.HasValue) vals.Add(n.Value); }
        return vals.Count > 0 ? FR(vals.Average()) : FormulaResult.Error("#DIV/0!");
    }

    private FormulaResult? EvalAverageIfs(List<object> args)
    {
        if (args.Count < 3) return null;
        var avgRange = AsResults(args[0]); if (avgRange == null) return null;
        var vals = new List<double>();
        for (int i = 0; i < avgRange.Length; i++)
        {
            var match = true;
            for (int c = 1; c + 1 < args.Count; c += 2)
            { var cr = AsResults(args[c]); var crit = args[c + 1] is FormulaResult cv ? cv.AsString() : "";
              if (cr == null || i >= cr.Length || !MatchesCriteria(cr[i], crit)) { match = false; break; } }
            if (match) { var n = AsNumeric(avgRange[i]); if (n.HasValue) vals.Add(n.Value); }
        }
        return vals.Count > 0 ? FR(vals.Average()) : FormulaResult.Error("#DIV/0!");
    }

    private FormulaResult? EvalMaxMinIfs(List<object> args, bool isMax)
    {
        if (args.Count < 3) return null;
        var valRange = AsResults(args[0]); if (valRange == null) return null;
        var vals = new List<double>();
        for (int i = 0; i < valRange.Length; i++)
        {
            var match = true;
            for (int c = 1; c + 1 < args.Count; c += 2)
            { var cr = AsResults(args[c]); var crit = args[c + 1] is FormulaResult cv ? cv.AsString() : "";
              if (cr == null || i >= cr.Length || !MatchesCriteria(cr[i], crit)) { match = false; break; } }
            if (match) { var n = AsNumeric(valRange[i]); if (n.HasValue) vals.Add(n.Value); }
        }
        return vals.Count > 0 ? FR(isMax ? vals.Max() : vals.Min()) : FR(0);
    }

    private FormulaResult? EvalSumProduct(List<object> args)
    {
        if (args.Count == 0) return FR(0);
        var arrays = args.Select(a =>
            a is RangeData rd ? rd.ToDoubleArray() :
            a is FormulaResult fr && fr.IsArray ? fr.ArrayValue :
            a is double[] arr ? arr : null).ToList();
        // Single numeric value: SUMPRODUCT(scalar) = scalar
        if (arrays.All(a => a == null) && args.Count == 1 && args[0] is FormulaResult single && single.IsNumeric)
            return single;
        if (arrays.Any(a => a == null)) return null;
        var len = arrays.Min(a => a!.Length); double sum = 0;
        for (int i = 0; i < len; i++) { double p = 1; foreach (var arr in arrays) p *= arr![i]; sum += p; }
        return FR(sum);
    }

    // ==================== Date ====================

    private static FormulaResult? EvalEomonth(List<object> args)
    {
        var d = args.Count > 0 && args[0] is FormulaResult r ? DateTime.FromOADate(r.AsNumber()) : DateTime.Today;
        var months = args.Count > 1 && args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        var t = d.AddMonths(months); return FR(new DateTime(t.Year, t.Month, DateTime.DaysInMonth(t.Year, t.Month)).ToOADate());
    }

    private static FormulaResult? EvalDateDif(List<object> args)
    {
        if (args.Count < 3) return null;
        var d1 = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var d2 = args[1] is FormulaResult r2 ? DateTime.FromOADate(r2.AsNumber()) : DateTime.Today;
        var unit = args[2] is FormulaResult r3 ? r3.AsString().ToUpperInvariant() : "D";
        return unit switch { "D" => FR((d2 - d1).Days), "M" => FR((d2.Year - d1.Year) * 12 + d2.Month - d1.Month), "Y" => FR(d2.Year - d1.Year), _ => null };
    }

    private static FormulaResult? EvalNetworkDays(List<object> args)
    {
        if (args.Count < 2) return null;
        var start = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var end = args[1] is FormulaResult r2 ? DateTime.FromOADate(r2.AsNumber()) : DateTime.Today;
        int count = 0; for (var d = start; d <= end; d = d.AddDays(1)) if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday) count++;
        return FR(count);
    }

    private static FormulaResult? EvalWorkDay(List<object> args)
    {
        if (args.Count < 2) return null;
        var start = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var days = args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        var d = start; var step = days > 0 ? 1 : -1; var rem = Math.Abs(days);
        while (rem > 0) { d = d.AddDays(step); if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday) rem--; }
        return FR(d.ToOADate());
    }

    private static FormulaResult? EvalYearFrac(List<object> args)
    {
        if (args.Count < 2) return null;
        var d1 = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var d2 = args[1] is FormulaResult r2 ? DateTime.FromOADate(r2.AsNumber()) : DateTime.Today;
        return FR(Math.Abs((d2 - d1).TotalDays / 365.25));
    }

    // ==================== Financial ====================

    private static FormulaResult? EvalPmt(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, nper = args[1] is FormulaResult r2 ? r2.AsNumber() : 0, pv = args[2] is FormulaResult r3 ? r3.AsNumber() : 0;
        var fv = args.Count > 3 && args[3] is FormulaResult r4 ? r4.AsNumber() : 0;
        if (rate == 0) return FR(-(pv + fv) / nper);
        return FR(-(rate * (pv * Math.Pow(1 + rate, nper) + fv) / (Math.Pow(1 + rate, nper) - 1)));
    }

    private static FormulaResult? EvalFv(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, nper = args[1] is FormulaResult r2 ? r2.AsNumber() : 0, pmt = args[2] is FormulaResult r3 ? r3.AsNumber() : 0;
        var pv = args.Count > 3 && args[3] is FormulaResult r4 ? r4.AsNumber() : 0;
        if (rate == 0) return FR(-(pv + pmt * nper));
        return FR(-(pv * Math.Pow(1 + rate, nper) + pmt * (Math.Pow(1 + rate, nper) - 1) / rate));
    }

    private static FormulaResult? EvalPv(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, nper = args[1] is FormulaResult r2 ? r2.AsNumber() : 0, pmt = args[2] is FormulaResult r3 ? r3.AsNumber() : 0;
        var fv = args.Count > 3 && args[3] is FormulaResult r4 ? r4.AsNumber() : 0;
        if (rate == 0) return FR(-(fv + pmt * nper));
        return FR(-(fv / Math.Pow(1 + rate, nper) + pmt * (1 - Math.Pow(1 + rate, -nper)) / rate));
    }

    private static FormulaResult? EvalNper(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, pmt = args[1] is FormulaResult r2 ? r2.AsNumber() : 0, pv = args[2] is FormulaResult r3 ? r3.AsNumber() : 0;
        var fv = args.Count > 3 && args[3] is FormulaResult r4 ? r4.AsNumber() : 0;
        if (rate == 0) return pmt != 0 ? FR(-(pv + fv) / pmt) : null;
        return FR(Math.Log((-fv * rate + pmt) / (pv * rate + pmt)) / Math.Log(1 + rate));
    }

    private static FormulaResult? EvalNpv(List<object> args)
    {
        if (args.Count < 2) return null;
        var rate = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var values = new List<double>();
        for (int i = 1; i < args.Count; i++) { if (args[i] is double[] arr) values.AddRange(arr); else if (args[i] is FormulaResult fr) values.Add(fr.AsNumber()); }
        double npv = 0; for (int i = 0; i < values.Count; i++) npv += values[i] / Math.Pow(1 + rate, i + 1);
        return FR(npv);
    }

    private static FormulaResult? EvalIpmt(List<object> args)
    {
        if (args.Count < 4) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, per = args[1] is FormulaResult r2 ? r2.AsNumber() : 0;
        double nper = args[2] is FormulaResult r3 ? r3.AsNumber() : 0, pv = args[3] is FormulaResult r4 ? r4.AsNumber() : 0;
        if (rate == 0) return FR(0);
        var pmt = rate * (pv * Math.Pow(1 + rate, nper)) / (Math.Pow(1 + rate, nper) - 1);
        var fvBefore = pv * Math.Pow(1 + rate, per - 1) + pmt * (Math.Pow(1 + rate, per - 1) - 1) / rate;
        return FR(-(fvBefore * rate));
    }

    private static FormulaResult? EvalPpmt(List<object> args)
    {
        if (args.Count < 4) return null;
        var pmt = EvalPmt(args)?.AsNumber() ?? 0;
        var ipmt = EvalIpmt(args)?.AsNumber() ?? 0;
        return FR(pmt - ipmt);
    }

    private static FormulaResult? EvalSyd(List<object> args)
    {
        if (args.Count < 4) return null;
        double cost = args[0] is FormulaResult r ? r.AsNumber() : 0, salvage = args[1] is FormulaResult r2 ? r2.AsNumber() : 0;
        double life = args[2] is FormulaResult r3 ? r3.AsNumber() : 0, per = args[3] is FormulaResult r4 ? r4.AsNumber() : 0;
        return FR((cost - salvage) * (life - per + 1) * 2 / (life * (life + 1)));
    }

    private static FormulaResult? EvalDb(List<object> args)
    {
        if (args.Count < 4) return null;
        double cost = args[0] is FormulaResult r ? r.AsNumber() : 0, salvage = args[1] is FormulaResult r2 ? r2.AsNumber() : 0;
        double life = args[2] is FormulaResult r3 ? r3.AsNumber() : 0; int period = args[3] is FormulaResult r4 ? (int)r4.AsNumber() : 1;
        var rate = Math.Round(1 - Math.Pow(salvage / cost, 1.0 / life), 3);
        double total = 0;
        for (int p = 1; p <= period; p++) { var dep = (cost - total) * rate; total += dep; if (p == period) return FR(dep); }
        return FR(0);
    }

    private static FormulaResult? EvalDdb(List<object> args)
    {
        if (args.Count < 4) return null;
        double cost = args[0] is FormulaResult r ? r.AsNumber() : 0, salvage = args[1] is FormulaResult r2 ? r2.AsNumber() : 0;
        double life = args[2] is FormulaResult r3 ? r3.AsNumber() : 0; int period = args[3] is FormulaResult r4 ? (int)r4.AsNumber() : 1;
        var factor = args.Count > 4 && args[4] is FormulaResult r5 ? r5.AsNumber() : 2;
        double bv = cost;
        for (int p = 1; p <= period; p++) { var dep = Math.Min(bv * factor / life, Math.Max(bv - salvage, 0)); bv -= dep; if (p == period) return FR(dep); }
        return FR(0);
    }
}
