// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Function Dispatch (350+ functions) ====================

    private FormulaResult? EvalFunction(string name, List<object> args)
    {
        double[] nums() => FlattenNumbers(args);
        FormulaResult? arg(int i) => i < args.Count && args[i] is FormulaResult r ? r : null;
        double num(int i) => arg(i)?.AsNumber() ?? 0;
        string str(int i) => arg(i)?.AsString() ?? "";

        // Error propagation: most functions return the first error-valued scalar
        // argument. Functions that classify/handle errors, that take lazy
        // branches (args are eagerly evaluated here), or that count/ignore
        // errors are exempt and inspect the error themselves. A whole-array error
        // argument is not a scalar error here — it flows on to array spilling.
        if (!ErrorTransparentFns.Contains(name))
            foreach (var a in args)
                if (a is FormulaResult { IsError: true } errArg) return errArg;

        // Implicit array spilling: an element-wise scalar function applied to an
        // array argument maps over each element and spills (Excel 365 behavior).
        // Only opt-in functions lift; array-consuming functions (SUM, SORT,
        // INDEX, lookups, ...) are excluded and handle arrays themselves.
        if (LiftableScalarFunctions.Contains(name) && TryLiftOverArrays(name, args) is { } lifted)
            return lifted;

        return name switch
        {
            // ===== Math & Aggregation =====
            "SUM" => CheckRangeErrors(args) ?? FR(nums().Sum()),
            "SUBTOTAL" => EvalSubtotal(args),
            "AGGREGATE" => EvalAggregate(args),
            "SUMPRODUCT" => EvalSumProduct(args),
            "AVERAGE" => CheckRangeErrors(args) ?? (nums() is { Length: > 0 } a ? FR(a.Average()) : null),
            "COUNT" => FR(nums().Length),
            // COUNTA counts every non-empty value, including error values; only a
            // genuinely blank cell is skipped.
            "COUNTA" => FR(args.Sum(a => AsRangeData(a) is { } rd ? rd.ToFlatResults().Count(c => c != null && !c.IsBlank)
                : a is FormulaResult r && !r.IsRange ? (r.IsBlank ? 0 : 1) : a is double[] arr ? arr.Length : 0)),
            // Count cells that are empty OR whose formula result is "" (Excel treats ="" as
            // blank in COUNTBLANK, unlike ISBLANK). A number 0 and non-empty text are not
            // blank. Range is fully materialized with null placeholders for truly-empty
            // cells, so a direct count is exact.
            "COUNTBLANK" => FR(args.Sum(a => AsResults(a) is { } rd
                ? rd.Count(c => c == null || c.IsBlank || (c.IsString && c.AsString() == ""))
                : 0)),
            "MIN" => CheckRangeErrors(args) ?? (nums() is { Length: > 0 } mn ? FR(mn.Min()) : FR(0)),
            "MAX" => CheckRangeErrors(args) ?? (nums() is { Length: > 0 } mx ? FR(mx.Max()) : FR(0)),
            "ABS" => FR(Math.Abs(num(0))),
            "SIGN" => FR(Math.Sign(num(0))),
            "INT" => FR(Math.Floor(num(0))),
            "TRUNC" => args.Count >= 2 ? FR(Math.Truncate(num(0) * Math.Pow(10, num(1))) / Math.Pow(10, num(1))) : FR(Math.Truncate(num(0))),
            "ROUND" => FR(ExcelRound(num(0), (int)num(1))),
            "ROUNDUP" => FR(RoundUp(num(0), (int)num(1))),
            "ROUNDDOWN" => FR(RoundDown(num(0), (int)num(1))),
            "CEILING_MATH" => FR(CeilingMath(num(0), args.Count >= 2 ? num(1) : 1, args.Count >= 3 ? num(2) : 0)),
            "FLOOR_MATH" => FR(FloorMath(num(0), args.Count >= 2 ? num(1) : 1, args.Count >= 3 ? num(2) : 0)),
            // Legacy CEILING/FLOOR: a positive number with negative significance
            // is #NUM! (other sign combinations round as usual).
            "CEILING" => num(0) > 0 && args.Count >= 2 && num(1) < 0 ? FormulaResult.Error("#NUM!") : FR(CeilingF(num(0), args.Count >= 2 ? num(1) : 1)),
            "FLOOR" => num(0) > 0 && args.Count >= 2 && num(1) < 0 ? FormulaResult.Error("#NUM!") : FR(FloorF(num(0), args.Count >= 2 ? num(1) : 1)),
            "MOD" => num(1) != 0 ? FR(num(0) - num(1) * Math.Floor(num(0) / num(1))) : FormulaResult.Error("#DIV/0!"),
            "POWER" => num(0) == 0 && num(1) == 0 ? FormulaResult.Error("#NUM!") : FR(ExcelPow(num(0), num(1))),
            "SQRT" => num(0) >= 0 ? FR(Math.Sqrt(num(0))) : FormulaResult.Error("#NUM!"),
            "FACT" => num(0) < 0 ? FormulaResult.Error("#NUM!") : FR(Factorial(num(0))),
            "COMBIN" => (int)num(0) < 0 || (int)num(1) < 0 || (int)num(1) > (int)num(0) ? FormulaResult.Error("#NUM!") : FR(Combin((int)num(0), (int)num(1))),
            "COMBINA" => (int)num(0) < 0 || (int)num(1) < 0 ? FormulaResult.Error("#NUM!") : FR(Combin((int)num(0) + (int)num(1) - 1, (int)num(1))),
            "PERMUT" => (int)num(0) < 0 || (int)num(1) < 0 || (int)num(1) > (int)num(0) ? FormulaResult.Error("#NUM!") : FR(Permut((int)num(0), (int)num(1))),
            "SUMSQ" => CheckRangeErrors(args) ?? FR(nums().Sum(x => x * x)),
            "SUMX2MY2" => EvalSumXY(args, 0),
            "SUMX2PY2" => EvalSumXY(args, 1),
            "SUMXMY2" => EvalSumXY(args, 2),
            "CEILING_PRECISE" or "ISO_CEILING" => FR(CeilingPrecise(num(0), args.Count >= 2 ? num(1) : 1)),
            "FLOOR_PRECISE" => FR(FloorPrecise(num(0), args.Count >= 2 ? num(1) : 1)),
            "SQRTPI" => num(0) >= 0 ? FR(Math.Sqrt(num(0) * Math.PI)) : FormulaResult.Error("#NUM!"),
            "FACTDOUBLE" => FactDouble(num(0)),
            "MULTINOMIAL" => CheckRangeErrors(args) ?? FR(Multinomial(nums())),
            "SERIESSUM" => EvalSeriesSum(args),
            "GCD" => CheckRangeErrors(args) ?? FR(nums().Aggregate(0.0, (a, b) => Gcd((long)a, (long)b))),
            "LCM" => CheckRangeErrors(args) ?? FR(nums().Aggregate(1.0, (a, b) => Lcm((long)a, (long)b))),
            "RAND" => FR(new Random().NextDouble()),
            "RANDBETWEEN" => FR(new Random().Next((int)num(0), (int)num(1) + 1)),
            "EVEN" => FR(EvenF(num(0))),
            "ODD" => FR(OddF(num(0))),
            "PRODUCT" => CheckRangeErrors(args) ?? FR(nums().Aggregate(1.0, (a, b) => a * b)),
            "QUOTIENT" => num(1) != 0 ? FR(Math.Truncate(num(0) / num(1))) : FormulaResult.Error("#DIV/0!"),
            "MROUND" => MRound(num(0), num(1)),
            "ROMAN" => FR_S(ToRoman((int)num(0))),
            "ARABIC" => FR(FromRoman(str(0))),
            "BASE" => EvalBase((long)num(0), (int)num(1), args.Count >= 3 ? (int)num(2) : 0),
            "DECIMAL" => EvalDecimal(str(0), (int)num(1)),
            "LOG" => args.Count >= 2 ? FR(Math.Log(num(0), num(1))) : FR(Math.Log10(num(0))),
            "LOG10" => FR(Math.Log10(num(0))),
            "LN" => FR(Math.Log(num(0))),
            "EXP" => FR(Math.Exp(num(0))),

            // ===== Trigonometry =====
            "PI" => FR(Math.PI),
            "SIN" => FR(Math.Sin(num(0))), "COS" => FR(Math.Cos(num(0))), "TAN" => FR(Math.Tan(num(0))),
            "ASIN" => FR(Math.Asin(num(0))), "ACOS" => FR(Math.Acos(num(0))), "ATAN" => FR(Math.Atan(num(0))),
            // Excel ATAN2(x, y) is the angle of point (x, y); .NET Math.Atan2(y, x)
            // takes y first. Pass them swapped so ATAN2(1,0)=0, not pi/2.
            "ATAN2" => FR(Math.Atan2(num(1), num(0))),
            "SINH" => FR(Math.Sinh(num(0))), "COSH" => FR(Math.Cosh(num(0))), "TANH" => FR(Math.Tanh(num(0))),
            "ASINH" => FR(Math.Asinh(num(0))), "ACOSH" => FR(Math.Acosh(num(0))), "ATANH" => FR(Math.Atanh(num(0))),
            "DEGREES" => FR(num(0) * 180.0 / Math.PI),
            "RADIANS" => FR(num(0) * Math.PI / 180.0),
            "SEC" => FR(1.0 / Math.Cos(num(0))), "CSC" => Math.Sin(num(0)) == 0 ? FormulaResult.Error("#DIV/0!") : FR(1.0 / Math.Sin(num(0))),
            "COT" => Math.Tan(num(0)) == 0 ? FormulaResult.Error("#DIV/0!") : FR(1.0 / Math.Tan(num(0))), "SECH" => FR(1.0 / Math.Cosh(num(0))),
            "CSCH" => Math.Sinh(num(0)) == 0 ? FormulaResult.Error("#DIV/0!") : FR(1.0 / Math.Sinh(num(0))), "COTH" => Math.Tanh(num(0)) == 0 ? FormulaResult.Error("#DIV/0!") : FR(1.0 / Math.Tanh(num(0))),
            "ACOT" => FR(Math.PI / 2 - Math.Atan(num(0))),
            "ACOTH" => FR(0.5 * Math.Log((num(0) + 1) / (num(0) - 1))),

            // ===== Statistical =====
            "MEDIAN" => CheckRangeErrors(args) ?? EvalMedian(nums()),
            "MODE" or "MODE_SNGL" => CheckRangeErrors(args) ?? EvalMode(nums()),
            "MODE_MULT" => CheckRangeErrors(args) ?? EvalModeMult(nums()),
            "LARGE" => CheckRangeErrors(args) ?? EvalLarge(args), "SMALL" => CheckRangeErrors(args) ?? EvalSmall(args),
            "RANK" or "RANK_EQ" => CheckRangeErrors(args) ?? EvalRank(args),
            "RANK_AVG" => CheckRangeErrors(args) ?? EvalRankAvg(args),
            "PROB" => EvalProb(args),
            "MDETERM" => EvalMdeterm(arg(0)),
            "MMULT" => EvalMmult(arg(0), arg(1)),
            "MINVERSE" => EvalMinverse(arg(0)),
            "PERCENTILE" or "PERCENTILE_INC" => CheckRangeErrors(args) ?? EvalPercentile(args),
            "PERCENTRANK" or "PERCENTRANK_INC" => CheckRangeErrors(args) ?? EvalPercentRank(args),
            "STDEV" or "STDEV_S" => CheckRangeErrors(args) ?? EvalStdev(nums(), true),
            "STDEVP" or "STDEV_P" => CheckRangeErrors(args) ?? EvalStdev(nums(), false),
            "VAR" or "VAR_S" => CheckRangeErrors(args) ?? EvalVar(nums(), true),
            "VARP" or "VAR_P" => CheckRangeErrors(args) ?? EvalVar(nums(), false),
            "AVERAGEA" => NumsA(args) is { Length: > 0 } aa ? FR(aa.Average()) : FormulaResult.Error("#DIV/0!"),
            "MAXA" => NumsA(args) is { Length: > 0 } mxa ? FR(mxa.Max()) : FR(0),
            "MINA" => NumsA(args) is { Length: > 0 } mna ? FR(mna.Min()) : FR(0),
            "GEOMEAN" => CheckRangeErrors(args) ?? (nums() is { Length: > 0 } gm ? FR(Math.Pow(gm.Aggregate(1.0, (a, b) => a * b), 1.0 / gm.Length)) : null),
            "HARMEAN" => CheckRangeErrors(args) ?? (nums() is { Length: > 0 } hm ? FR(hm.Length / hm.Sum(x => 1.0 / x)) : null),

            // ===== Statistical distributions (special-function based) =====
            "NORM_DIST" or "NORMDIST" => EvalNormDist(args),
            "NORM_S_DIST" => EvalNormSDist(args),
            "NORMSDIST" => FR(NormCdf(num(0))),
            "NORM_INV" or "NORMINV" => args.Count >= 3 ? FR(num(1) + num(2) * InvNormCdf(num(0))) : null,
            "NORM_S_INV" or "NORMSINV" => FR(InvNormCdf(num(0))),
            "STANDARDIZE" => num(2) > 0 ? FR((num(0) - num(1)) / num(2)) : FormulaResult.Error("#NUM!"),
            "GAUSS" => FR(NormCdf(num(0)) - 0.5),
            "PHI" => FR(NormPdf(num(0))),
            "CONFIDENCE" or "CONFIDENCE_NORM" => EvalConfidenceNorm(args),
            "ERF" => EvalErf(args),
            "ERFC" => FR(Erfc(num(0))),
            "ERF_PRECISE" => FR(Erf(num(0))),
            "ERFC_PRECISE" => FR(Erfc(num(0))),
            "GAMMALN" or "GAMMALN_PRECISE" => num(0) > 0 ? FR(GammaLn(num(0))) : FormulaResult.Error("#NUM!"),
            "GAMMA" => EvalGamma(args),
            "GAMMA_DIST" or "GAMMADIST" => EvalGammaDist(args),
            "GAMMA_INV" or "GAMMAINV" => args.Count >= 3 ? FR(num(2) * InvRegGammaP(num(1), num(0))) : null,
            "CHISQ_DIST" => EvalChisqDist(args),
            "CHISQ_DIST_RT" or "CHIDIST" => args.Count >= 2 ? FR(RegGammaQ(num(1) / 2, num(0) / 2)) : null,
            "CHISQ_INV" => args.Count >= 2 ? FR(2 * InvRegGammaP(num(1) / 2, num(0))) : null,
            "CHISQ_INV_RT" or "CHIINV" => args.Count >= 2 ? FR(2 * InvRegGammaP(num(1) / 2, 1 - num(0))) : null,
            "POISSON_DIST" or "POISSON" => EvalPoisson(args),
            "EXPON_DIST" or "EXPONDIST" => EvalExpon(args),
            "FISHER" => Math.Abs(num(0)) < 1 ? FR(0.5 * Math.Log((1 + num(0)) / (1 - num(0)))) : FormulaResult.Error("#NUM!"),
            "FISHERINV" => FR(Math.Tanh(num(0))),
            // ----- incomplete-beta family -----
            "BETA_DIST" or "BETADIST" => EvalBetaDist(args),
            "BETA_INV" or "BETAINV" => EvalBetaInv(args),
            "T_DIST" => EvalTDist(args),
            "T_DIST_2T" => EvalTDist2T(args),
            "T_DIST_RT" => args.Count >= 2 ? FR(TDistRightTail(num(0), num(1))) : null,
            "TDIST" => EvalTDistLegacy(args),
            "T_INV" => args.Count >= 2 ? FR(TInv(num(0), num(1))) : null,
            "T_INV_2T" or "TINV" => args.Count >= 2 ? FR(TInv2T(num(0), num(1))) : null,
            "F_DIST" => EvalFDist(args),
            "F_DIST_RT" or "FDIST" => args.Count >= 3 ? FR(FDistRightTail(num(0), num(1), num(2))) : null,
            "F_INV" => args.Count >= 3 ? FR(FInv(num(0), num(1), num(2))) : null,
            "F_INV_RT" or "FINV" => args.Count >= 3 ? FR(FInv(1 - num(0), num(1), num(2))) : null,
            "BINOM_DIST" or "BINOMDIST" => EvalBinomDist(args),
            "BINOM_INV" or "CRITBINOM" => EvalBinomInv(args),
            "NEGBINOM_DIST" or "NEGBINOMDIST" => EvalNegBinom(args),
            "WEIBULL_DIST" or "WEIBULL" => EvalWeibull(args),
            "LOGNORM_DIST" or "LOGNORMDIST" => EvalLognormDist(args),
            "LOGNORM_INV" or "LOGINV" => args.Count < 3 ? null
                : num(0) > 0 && num(0) < 1 ? FR(Math.Exp(num(1) + num(2) * InvNormCdf(num(0)))) : FormulaResult.Error("#NUM!"),
            "HYPGEOM_DIST" or "HYPGEOMDIST" => EvalHypgeom(args),
            // ----- descriptive & regression -----
            "SKEW" => EvalSkew(args, population: false),
            "SKEW_P" => EvalSkew(args, population: true),
            "KURT" => EvalKurt(args),
            "AVEDEV" => CheckRangeErrors(args) ?? (nums() is { Length: > 0 } ad ? FR(ad.Select(v => Math.Abs(v - ad.Average())).Average()) : FormulaResult.Error("#NUM!")),
            "DEVSQ" => CheckRangeErrors(args) ?? (nums() is { Length: > 0 } ds ? FR(ds.Sum(v => (v - ds.Average()) * (v - ds.Average()))) : FR(0)),
            "TRIMMEAN" => EvalTrimMean(args),
            "PERMUTATIONA" => FR(Math.Pow((int)num(0), (int)num(1))),
            "CORREL" or "PEARSON" => EvalCorrel(args),
            "COVARIANCE_P" or "COVAR" => EvalCovar(args, sample: false),
            "COVARIANCE_S" => EvalCovar(args, sample: true),
            "SLOPE" => EvalSlope(args), "INTERCEPT" => EvalIntercept(args),
            "RSQ" => EvalRsq(args), "STEYX" => EvalSteyx(args),
            "FORECAST" or "FORECAST_LINEAR" => EvalForecast(args),
            "QUARTILE" or "QUARTILE_INC" => EvalQuartile(args, exclusive: false),
            "QUARTILE_EXC" => EvalQuartile(args, exclusive: true),
            "PERCENTILE_EXC" => EvalPercentileExc(args),
            // Hypothesis tests (W4d) — dotted modern + legacy aliases.
            "T_TEST" or "TTEST" => EvalTTest(args),
            "CHISQ_TEST" or "CHITEST" => EvalChisqTest(args),
            "F_TEST" or "FTEST" => EvalFTest(args),
            "Z_TEST" or "ZTEST" => EvalZTest(args),
            // Array regression (W4d) — spill {coefficients}/{predictions}.
            "LINEST" => EvalLinest(args, log: false),
            "LOGEST" => EvalLinest(args, log: true),
            "TREND" => EvalTrend(args, log: false),
            "GROWTH" => EvalTrend(args, log: true),

            // ===== Logical =====
            "IF" => EvalIf(args), "IFS" => EvalIfs(args),
            "AND" => FR_B(AllArgs(args).All(r => r.AsNumber() != 0)),
            "OR" => FR_B(AllArgs(args).Any(r => r.AsNumber() != 0)),
            "NOT" => FR_B(num(0) == 0),
            "XOR" => FR_B(AllArgs(args).Count(r => r.AsNumber() != 0) % 2 == 1),
            "TRUE" => FR_B(true), "FALSE" => FR_B(false),
            "IFERROR" => arg(0) is { IsError: true } ? arg(1) : arg(0),
            "IFNA" => arg(0)?.ErrorValue == "#N/A" ? arg(1) : arg(0),
            "SWITCH" => EvalSwitch(args), "CHOOSE" => EvalChoose(args),
            "REDUCE" => EvalReduce(args),
            "ISOMITTED" => FR_B(args.Count > 0 && IsOmittedArg(args[0])),

            // ===== Text =====
            "CONCATENATE" or "CONCAT" => EvalConcat(args),
            "TEXTJOIN" => EvalTextJoin(args),
            "VALUETOTEXT" => EvalValueToText(args),
            "ARRAYTOTEXT" => EvalArrayToText(args),
            "LEFT" => EvalLeftRight(str(0), args.Count >= 2 ? (int)num(1) : 1, true),
            "RIGHT" => EvalLeftRight(str(0), args.Count >= 2 ? (int)num(1) : 1, false),
            "MID" => EvalMid(args),
            "LEN" => FR(str(0).Length),
            "TRIM" => FR_S(Regex.Replace(str(0).Trim(' '), " +", " ")),   // spaces (char 32) only; tabs/newlines are CLEAN's job
            "CLEAN" => FR_S(Regex.Replace(str(0), @"[\x00-\x1F\x7F]", "")),   // DEL included, matching Excel
            "UPPER" => FR_S(str(0).ToUpperInvariant()),
            "LOWER" => FR_S(str(0).ToLowerInvariant()),
            "PROPER" => FR_S(ProperCase(str(0))),
            "REPT" => (int)num(1) < 0 || (long)str(0).Length * (int)num(1) > 32767 ? FormulaResult.Error("#VALUE!") : FR_S(string.Concat(Enumerable.Repeat(str(0), (int)num(1)))),
            // CHAR is defined only for codes 1..255; out-of-range is #VALUE!.
            "CHAR" => (int)num(0) is >= 1 and <= 255 ? FR_S(((char)(int)num(0)).ToString()) : FormulaResult.Error("#VALUE!"),
            "CODE" => str(0).Length > 0 ? FR((int)str(0)[0]) : FormulaResult.Error("#VALUE!"),
            "LENB" => FR(str(0).Sum(c => IsWideChar(c) ? 2 : 1)),
            "ASC" => FR_S(ToHalfWidth(str(0))),
            // UNICHAR/UNICODE work on full Unicode code points (surrogate pairs),
            // unlike CHAR/CODE which are limited to a single UTF-16 unit.
            "UNICHAR" => EvalUnichar((int)num(0)),
            "UNICODE" => str(0).Length > 0 ? FR(char.ConvertToUtf32(str(0), 0)) : FormulaResult.Error("#VALUE!"),
            "FIND" => EvalFind(args, true), "SEARCH" => EvalFind(args, false),
            "REPLACE" => EvalReplace(args), "SUBSTITUTE" => EvalSubstitute(args),
            "EXACT" => FR_B(str(0) == str(1)),
            "VALUE" => EvalValue(str(0)),
            "TEXT" => EvalText(args),
            "TEXTBEFORE" => EvalTextBeforeAfter(args, before: true),
            "TEXTAFTER" => EvalTextBeforeAfter(args, before: false),
            "REGEXTEST" => EvalRegexTest(args),
            "REGEXEXTRACT" => EvalRegexExtract(args),
            "REGEXREPLACE" => EvalRegexReplace(args),
            "T" => arg(0) is { IsString: true } ? arg(0) : FR_S(""),
            "N" => arg(0) is { IsError: true } ? arg(0)! : arg(0) is { IsString: true } ? FR(0) : FR(num(0)),
            "FIXED" => EvalFixed(args),
            "NUMBERVALUE" => EvalNumberValue(args),
            "DOLLAR" => EvalCurrency(args, "$", 2),
            "YEN" => EvalCurrency(args, "¥", 0),

            // ===== Lookup & Reference =====
            "INDEX" => EvalIndex(args), "MATCH" => EvalMatch(args),
            "ROW" => EvalRowCol(args, true), "COLUMN" => EvalRowCol(args, false),
            "ROWS" => EvalRowsCols(args, true), "COLUMNS" => EvalRowsCols(args, false),
            "AREAS" => args.Count >= 1 ? FR(1) : FormulaResult.Error("#VALUE!"),
            "ADDRESS" => EvalAddress(args),
            "SHEET" => EvalSheet(args), "SHEETS" => EvalSheets(args),
            "CELL" => EvalCell(args),
            "VLOOKUP" => EvalVlookup(args),
            "HLOOKUP" => EvalHlookup(args),
            "LOOKUP" => EvalLookup(args),
            "XLOOKUP" => EvalXlookup(args),
            "XMATCH" => EvalXmatch(args),
            "HYPERLINK" => FR_S(args.Count >= 2 && args[1] is FormulaResult fn ? fn.AsString() : str(0)),
            "OFFSET" => EvalOffset(args),
            "INDIRECT" => EvalIndirect(args),

            // ===== Dynamic arrays / spill (W6) =====
            "SEQUENCE" => EvalSequence(args),
            "TRANSPOSE" => EvalTranspose(args),
            "SORT" => EvalSort(args),
            "SORTBY" => EvalSortBy(args),
            "UNIQUE" => EvalUnique(args),
            "FILTER" => EvalFilter(args),
            "TAKE" => EvalTake(args),
            "DROP" => EvalDrop(args),
            "CHOOSEROWS" => EvalChooseRowsCols(args, rowsMode: true),
            "CHOOSECOLS" => EvalChooseRowsCols(args, rowsMode: false),
            "TOCOL" => EvalToColRow(args, toCol: true),
            "TOROW" => EvalToColRow(args, toCol: false),
            "EXPAND" => EvalExpand(args),
            "HSTACK" => EvalStack(args, horizontal: true),
            "VSTACK" => EvalStack(args, horizontal: false),
            "WRAPROWS" => EvalWrap(args, byRows: true),
            "WRAPCOLS" => EvalWrap(args, byRows: false),
            "TEXTSPLIT" => EvalTextSplit(args),
            "MAP" => EvalMap(args),
            "BYROW" => EvalByRow(args),
            "BYCOL" => EvalByCol(args),
            "SCAN" => EvalScan(args),
            "MAKEARRAY" => EvalMakeArray(args),

            // ===== Date & Time =====
            "TODAY" => FR(DateTime.Today.ToOADate()), "NOW" => FR(DateTime.Now.ToOADate()),
            "DATE" => EvalDate(num(0), num(1), num(2)),
            "YEAR" => FR(DateTime.FromOADate(num(0)).Year), "MONTH" => FR(DateTime.FromOADate(num(0)).Month),
            "DAY" => FR(DateTime.FromOADate(num(0)).Day), "HOUR" => FR(TimeRoundedToSecond(num(0)).Hour),
            "MINUTE" => FR(TimeRoundedToSecond(num(0)).Minute), "SECOND" => FR(TimeRoundedToSecond(num(0)).Second),
            "WEEKDAY" => EvalWeekday(num(0), args.Count >= 2 ? (int)num(1) : 1),
            "DATEVALUE" => DateTime.TryParse(str(0), out var dv) ? FR(dv.ToOADate()) : FormulaResult.Error("#VALUE!"),
            "TIMEVALUE" => DateTime.TryParse(str(0), out var tv) ? FR(tv.TimeOfDay.TotalDays) : FormulaResult.Error("#VALUE!"),
            "EDATE" => FR(DateTime.FromOADate(num(0)).AddMonths((int)num(1)).ToOADate()),
            "EOMONTH" => EvalEomonth(args),
            "DAYS" => FR(num(0) - num(1)),
            "DATEDIF" => EvalDateDif(args),
            "NETWORKDAYS" => EvalNetworkDays(args, false),
            "NETWORKDAYS_INTL" => EvalNetworkDays(args, true),
            "WORKDAY" => EvalWorkDay(args, false),
            "WORKDAY_INTL" => EvalWorkDay(args, true),
            "ISOWEEKNUM" => FR(CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.FromOADate(num(0)), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)),
            "TIME" => FR((((long)num(0) * 3600 + (long)num(1) * 60 + (long)num(2)) % 86400 + 86400) % 86400 / 86400.0),
            "WEEKNUM" => EvalWeekNum(num(0), args.Count >= 2 ? (int)num(1) : 1),
            "DAYS360" => EvalDays360(num(0), num(1), args.Count >= 3 && arg(2)?.AsNumber() != 0),
            "YEARFRAC" => EvalYearFrac(args),

            // ===== Info =====
            "ISNUMBER" => FR_B(arg(0)?.IsNumeric == true),
            "ISTEXT" => FR_B(arg(0)?.IsString == true),
            "ISBLANK" => FR_B(arg(0) == null || arg(0)!.IsBlank),
            "ISERROR" => args.Count > 0 && AsRangeData(args[0]) is { } rd_err
                ? FormulaResult.Array(rd_err.ToFlatResults().Select(r => r?.IsError == true ? 1.0 : 0.0).ToArray())
                : FR_B(arg(0)?.IsError == true),
            // ISERR is ISERROR minus #N/A: a not-available result is reported as
            // "not an error" so callers can tell a lookup miss from a real fault.
            "ISERR" => args.Count > 0 && AsRangeData(args[0]) is { } rd_iserr
                ? FormulaResult.Array(rd_iserr.ToFlatResults().Select(r => r?.IsError == true && r.ErrorValue != "#N/A" ? 1.0 : 0.0).ToArray())
                : FR_B(arg(0)?.IsError == true && arg(0)?.ErrorValue != "#N/A"),
            "ISNA" => FR_B(arg(0)?.ErrorValue == "#N/A"),
            "ISLOGICAL" => FR_B(arg(0)?.IsBool == true),
            "ISEVEN" => FR_B((int)num(0) % 2 == 0), "ISODD" => FR_B((int)num(0) % 2 != 0),
            "ISNONTEXT" => FR_B(arg(0)?.IsString != true),
            "ISREF" => EvalIsRef(args),
            "ISFORMULA" => EvalIsFormula(args),
            "TYPE" => FR(arg(0) switch { { IsNumeric: true } => 1, { IsString: true } => 2, { IsBool: true } => 4, { IsError: true } => 16, _ => 1 }),
            "NA" => FormulaResult.Error("#N/A"),
            "ERROR_TYPE" => arg(0)?.ErrorValue switch { "#NULL!" => FR(1), "#DIV/0!" => FR(2), "#VALUE!" => FR(3), "#REF!" => FR(4), "#NAME?" => FR(5), "#NUM!" => FR(6), "#N/A" => FR(7), _ => FormulaResult.Error("#N/A") },

            // ===== Conditional Aggregation =====
            "SUMIF" => EvalSumIf(args), "SUMIFS" => EvalSumIfs(args),
            "COUNTIF" => EvalCountIf(args), "COUNTIFS" => EvalCountIfs(args),
            "AVERAGEIF" => EvalAverageIf(args), "AVERAGEIFS" => EvalAverageIfs(args),
            "MAXIFS" => EvalMaxMinIfs(args, true), "MINIFS" => EvalMaxMinIfs(args, false),

            // ===== Financial =====
            "PMT" => EvalPmt(args), "FV" => EvalFv(args), "PV" => EvalPv(args), "NPER" => EvalNper(args),
            "NPV" => EvalNpv(args), "IPMT" => EvalIpmt(args), "PPMT" => EvalPpmt(args),
            "SLN" => args.Count >= 3 ? FR((num(0) - num(1)) / num(2)) : null,
            "SYD" => EvalSyd(args), "DB" => EvalDb(args), "DDB" => EvalDdb(args), "VDB" => EvalVdb(args),
            "RATE" => EvalRate(args), "IRR" => EvalIrr(args), // via shared root solver
            "XNPV" => EvalXnpv(args), "XIRR" => EvalXirr(args), "MIRR" => EvalMirr(args),
            "CUMIPMT" => EvalCumulative(args, principal: false), "CUMPRINC" => EvalCumulative(args, principal: true),
            "FVSCHEDULE" => EvalFvSchedule(args),
            "PDURATION" => EvalPduration(args), "RRI" => EvalRri(args),
            "EFFECT" => EvalEffect(args), "NOMINAL" => EvalNominal(args),
            "DOLLARDE" => EvalDollar(args, toDecimal: true), "DOLLARFR" => EvalDollar(args, toDecimal: false),
            "ISPMT" => EvalIspmt(args),

            // ===== Securities / coupon bonds (W3b) =====
            "COUPDAYS" => EvalCoupDays(args), "COUPDAYBS" => EvalCoupDayBs(args),
            "COUPDAYSNC" => EvalCoupDaysNc(args), "COUPNCD" => EvalCoupNcd(args),
            "COUPPCD" => EvalCoupPcd(args), "COUPNUM" => EvalCoupNum(args),
            "ACCRINT" => EvalAccrInt(args), "ACCRINTM" => EvalAccrIntM(args),
            "DISC" => EvalDisc(args), "INTRATE" => EvalIntRate(args), "RECEIVED" => EvalReceived(args),
            "PRICEDISC" => EvalPriceDisc(args), "YIELDDISC" => EvalYieldDisc(args),
            "PRICEMAT" => EvalPriceMat(args), "YIELDMAT" => EvalYieldMat(args),
            "TBILLEQ" => EvalTBillEq(args), "TBILLPRICE" => EvalTBillPrice(args), "TBILLYIELD" => EvalTBillYield(args),
            "PRICE" => EvalPrice(args), "YIELD" => EvalYield(args),
            "DURATION" => EvalDuration(args, modified: false), "MDURATION" => EvalDuration(args, modified: true),

            // ===== Database (Dxxx) — aggregate a table column over criteria =====
            "DSUM" => EvalDatabase(args, DbAgg.Sum), "DCOUNT" => EvalDatabase(args, DbAgg.Count),
            "DCOUNTA" => EvalDatabase(args, DbAgg.CountA), "DAVERAGE" => EvalDatabase(args, DbAgg.Average),
            "DMAX" => EvalDatabase(args, DbAgg.Max), "DMIN" => EvalDatabase(args, DbAgg.Min),
            "DPRODUCT" => EvalDatabase(args, DbAgg.Product), "DGET" => EvalDatabase(args, DbAgg.Get),
            "DSTDEV" => EvalDatabase(args, DbAgg.StdevS), "DSTDEVP" => EvalDatabase(args, DbAgg.StdevP),
            "DVAR" => EvalDatabase(args, DbAgg.VarS), "DVARP" => EvalDatabase(args, DbAgg.VarP),

            // ===== Conversion =====
            "CONVERT" => EvalConvert(num(0), str(1), str(2)),
            "BIN2DEC" => FromBase2Dec(str(0), 2),
            "DEC2BIN" => Dec2Base(num(0), 2, arg(1)),
            "HEX2DEC" => FromBase2Dec(str(0), 16),
            "DEC2HEX" => Dec2Base(num(0), 16, arg(1)),
            "OCT2DEC" => FromBase2Dec(str(0), 8),
            "DEC2OCT" => Dec2Base(num(0), 8, arg(1)),
            "BIN2HEX" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 2), 16).ToUpperInvariant()),
            "BIN2OCT" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 2), 8)),
            "HEX2BIN" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 16), 2)),
            "HEX2OCT" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 16), 8)),
            "OCT2BIN" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 8), 2)),
            "OCT2HEX" => FR_S(Convert.ToString(Convert.ToInt64(str(0), 8), 16).ToUpperInvariant()),

            // ===== Engineering: complex numbers =====
            "COMPLEX" => EvalComplex(args),
            "IMABS" => ImScalar(args, c => c.Magnitude),
            "IMREAL" => ImScalar(args, c => c.Real),
            "IMAGINARY" => ImScalar(args, c => c.Imaginary),
            "IMARGUMENT" => args.Count >= 1 && ArgComplex(args[0], out var ca, out _) && ca == System.Numerics.Complex.Zero
                ? FormulaResult.Error("#DIV/0!") : ImScalar(args, c => c.Phase),
            "IMCONJUGATE" => ImUnary(args, System.Numerics.Complex.Conjugate),
            "IMSUM" => ImFold(args, System.Numerics.Complex.Zero, (a, b) => a + b),
            "IMSUB" => ImBinary(args, (a, b) => a - b),
            "IMPRODUCT" => ImFold(args, System.Numerics.Complex.One, (a, b) => a * b),
            "IMDIV" => ImBinary(args, (a, b) => a / b),
            "IMPOWER" => ImUnary(args, c => System.Numerics.Complex.Pow(c, args.Count > 1 && args[1] is FormulaResult pw ? pw.AsNumber() : 0)),
            "IMSQRT" => ImUnary(args, System.Numerics.Complex.Sqrt),
            "IMEXP" => ImUnary(args, System.Numerics.Complex.Exp),
            "IMLN" => ImUnary(args, System.Numerics.Complex.Log),
            "IMLOG10" => ImUnary(args, c => System.Numerics.Complex.Log10(c)),
            "IMLOG2" => ImUnary(args, c => System.Numerics.Complex.Log(c) / Math.Log(2)),
            "IMSIN" => ImUnary(args, System.Numerics.Complex.Sin),
            "IMCOS" => ImUnary(args, System.Numerics.Complex.Cos),
            "IMTAN" => ImUnary(args, System.Numerics.Complex.Tan),
            "IMSINH" => ImUnary(args, System.Numerics.Complex.Sinh),
            "IMCOSH" => ImUnary(args, System.Numerics.Complex.Cosh),
            "IMSEC" => ImUnary(args, c => System.Numerics.Complex.One / System.Numerics.Complex.Cos(c)),
            "IMCSC" => ImUnary(args, c => System.Numerics.Complex.One / System.Numerics.Complex.Sin(c)),
            "IMCOT" => ImUnary(args, c => System.Numerics.Complex.Cos(c) / System.Numerics.Complex.Sin(c)),
            "IMSECH" => ImUnary(args, c => System.Numerics.Complex.One / System.Numerics.Complex.Cosh(c)),
            "IMCSCH" => ImUnary(args, c => System.Numerics.Complex.One / System.Numerics.Complex.Sinh(c)),

            // ===== Engineering: bit operations & step =====
            "BITAND" => BitOp(args, (a, b) => a & b),
            "BITOR" => BitOp(args, (a, b) => a | b),
            "BITXOR" => BitOp(args, (a, b) => a ^ b),
            "BITLSHIFT" => BitShift(args, left: true),
            "BITRSHIFT" => BitShift(args, left: false),
            "DELTA" => FR_B(num(0) == (args.Count > 1 ? num(1) : 0)),
            "GESTEP" => FR_B(num(0) >= (args.Count > 1 ? num(1) : 0)),

            _ => null
        };
    }

    // BITAND/BITOR/BITXOR — operands are non-negative integers below 2^48.
    private static FormulaResult? BitOp(List<object> args, Func<long, long, long> op)
    {
        if (args.Count < 2) return null;
        double a = args[0] is FormulaResult x ? x.AsNumber() : 0, b = args[1] is FormulaResult y ? y.AsNumber() : 0;
        if (a < 0 || b < 0 || a != Math.Floor(a) || b != Math.Floor(b) || a >= 281474976710656d || b >= 281474976710656d)
            return FormulaResult.Error("#NUM!");
        return FR(op((long)a, (long)b));
    }

    // BITLSHIFT(n, shift) / BITRSHIFT(n, shift). A negative shift reverses
    // direction, as Excel defines it.
    private static FormulaResult? BitShift(List<object> args, bool left)
    {
        if (args.Count < 2) return null;
        double n = args[0] is FormulaResult x ? x.AsNumber() : 0, sh = args[1] is FormulaResult y ? y.AsNumber() : 0;
        if (n < 0 || n != Math.Floor(n) || n >= 281474976710656d || Math.Abs(sh) > 53) return FormulaResult.Error("#NUM!");
        int s = (int)sh * (left ? 1 : -1);
        double result = s >= 0 ? (long)n << s : (long)n >> -s;
        return FR(result);
    }

    // REDUCE(initial, array, lambda) — fold the array through a 2-parameter
    // LAMBDA(accumulator, value), returning the final scalar accumulator.
    private FormulaResult? EvalReduce(List<object> args)
    {
        if (args.Count < 3) return null;
        var acc = args[0] as FormulaResult ?? FormulaResult.Number(0);
        if (args[2] is not FormulaResult { IsLambda: true } lam) return FormulaResult.Error("#VALUE!");
        var lambda = (Lambda)lam.LambdaValue!;
        foreach (var el in EnumerateElements(args[1]))
            acc = InvokeLambda(lambda, new List<FormulaResult> { acc, el });
        return acc;
    }

    // Flatten a range/array/scalar argument into element FormulaResults.
    private static IEnumerable<FormulaResult> EnumerateElements(object? a)
    {
        if (AsRangeData(a) is { } rd)
        {
            for (int r = 0; r < rd.Rows; r++)
                for (int c = 0; c < rd.Cols; c++)
                    yield return rd.Cells[r, c] ?? FormulaResult.Blank();
        }
        else if (a is FormulaResult { IsArray: true } arr)
            foreach (var v in arr.ArrayValue!) yield return FormulaResult.Number(v);
        else if (a is FormulaResult fr)
            yield return fr;
    }

    // SUBTOTAL(function_num, ref1, ...): function_num 1-11 (and 101-111 = ignore-hidden, treated
    // identically for preview) map onto the matching aggregate function. Re-dispatches into the
    // existing aggregate implementations.
    private FormulaResult? EvalSubtotal(List<object> args)
    {
        if (args.Count < 2 || args[0] is not FormulaResult fn) return null;
        var code = (int)fn.AsNumber() % 100; // 101-111 -> 1-11 (ignore-hidden simplification)
        var name = code switch
        {
            1 => "AVERAGE", 2 => "COUNT", 3 => "COUNTA", 4 => "MAX", 5 => "MIN",
            6 => "PRODUCT", 7 => "STDEV", 8 => "STDEVP", 9 => "SUM", 10 => "VAR", 11 => "VARP",
            _ => null
        };
        return name == null ? null : EvalFunction(name, args.Skip(1).ToList());
    }

    // AGGREGATE(function_num, options, ref1, ...): function_num 1-19; common ones mapped, the
    // options arg is ignored for preview.
    private FormulaResult? EvalAggregate(List<object> args)
    {
        if (args.Count < 3 || args[0] is not FormulaResult fn) return null;
        var code = (int)fn.AsNumber();
        var name = code switch
        {
            1 => "AVERAGE", 2 => "COUNT", 3 => "COUNTA", 4 => "MAX", 5 => "MIN",
            6 => "PRODUCT", 7 => "STDEV", 8 => "STDEVP", 9 => "SUM", 10 => "VAR", 11 => "VARP",
            12 => "MEDIAN", 14 => "LARGE", 15 => "SMALL",
            _ => null
        };
        if (name == null) return null;
        // options 2/3/6/7 ignore error values; strip them from range/array args
        // before aggregating so a computed array with error cells still works.
        int options = args[1] is FormulaResult o ? (int)o.AsNumber() : 0;
        var rest = args.Skip(2).ToList();
        if (options is 2 or 3 or 6 or 7)
            rest = rest.Select(StripErrorCells).ToList();
        return EvalFunction(name, rest);
    }

    // WEEKDAY(serial, [return_type]). The return_type selects which day starts
    // the week and the numbering base; the old code ignored it and always used
    // type 1. Mirrors Excel's set: 1/17 Sun=1..Sat=7; 2/11 Mon=1..Sun=7; 3
    // Mon=0..Sun=6; 12..16 week starts Tue..Sat. Invalid types are #NUM!.
    private static FormulaResult EvalWeekday(double serial, int returnType)
    {
        int dow = (int)DateTime.FromOADate(serial).DayOfWeek; // Sun=0..Sat=6
        int iso = (dow + 6) % 7;                              // Mon=0..Sun=6
        return returnType switch
        {
            1 or 17 => FR(dow + 1),
            2 or 11 => FR(iso + 1),
            3 => FR(iso),
            >= 12 and <= 16 => FR(((iso - (returnType - 11) + 7) % 7) + 1),
            _ => FormulaResult.Error("#NUM!"),
        };
    }

    // ==================== Logical ====================

    private FormulaResult? EvalIf(List<object> args)
    {
        var c = args.Count > 0 && args[0] is FormulaResult r ? r : null; if (c == null) return null;
        if (c.IsError) return c;
        bool isTrue;
        if (c.IsNumeric) isTrue = c.NumericValue != 0;
        else if (c.IsBool) isTrue = c.BoolValue == true;
        else if (c.IsBlank) isTrue = false;
        else
        {
            // A text condition coerces only for the literal words TRUE/FALSE;
            // any other text (including numeric-looking "1" or "") is #VALUE!.
            var sv = c.AsString();
            if (sv.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) isTrue = true;
            else if (sv.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) isTrue = false;
            else return FormulaResult.Error("#VALUE!");
        }
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
        var startNum = args.Count > 1 && args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        var len = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() : 0;
        // Excel requires start_num >= 1 and num_chars >= 0; otherwise #VALUE!.
        if (startNum < 1 || len < 0) return FormulaResult.Error("#VALUE!");
        var start = startNum - 1;
        if (start >= s.Length) return FR_S("");
        return FR_S(s.Substring(start, Math.Min(len, s.Length - start)));
    }

    private FormulaResult? EvalFind(List<object> args, bool caseSensitive)
    {
        var find = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        var within = args.Count > 1 && args[1] is FormulaResult r2 ? r2.AsString() : "";
        var startPos = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() - 1 : 0;
        if (startPos < 0 || startPos > within.Length) return FormulaResult.Error("#VALUE!");
        // SEARCH (case-insensitive) honours the ? and * wildcards; FIND does not.
        if (!caseSensitive && (find.Contains('*') || find.Contains('?')))
        {
            var m = Regex.Match(within[startPos..], WildcardToRegex(find), RegexOptions.IgnoreCase);
            return m.Success ? FR(startPos + m.Index + 1) : FormulaResult.Error("#VALUE!");
        }
        var idx = within.IndexOf(find, startPos, caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
        return idx >= 0 ? FR(idx + 1) : FormulaResult.Error("#VALUE!");
    }

    private FormulaResult? EvalReplace(List<object> args)
    {
        var s = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        var start = args.Count > 1 && args[1] is FormulaResult r2 ? (int)r2.AsNumber() - 1 : 0;
        var len = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() : 0;
        var rep = args.Count > 3 && args[3] is FormulaResult r4 ? r4.AsString() : "";
        if (start < 0 || len < 0) return FormulaResult.Error("#VALUE!");
        // start_num past the end appends the replacement (Excel clamps rather than errors).
        if (start > s.Length) start = s.Length;
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
            if (n < 1) return FormulaResult.Error("#VALUE!");
            for (int i = 0; i < n; i++) { idx = s.IndexOf(old, idx + 1, StringComparison.Ordinal); if (idx < 0) return FR_S(s); }
            return FR_S(s[..idx] + neo + s[(idx + old.Length)..]);
        }
        return FR_S(s.Replace(old, neo));
    }

    // TEXTBEFORE / TEXTAFTER(text, delimiter, [instance_num=1], [match_mode=0],
    // [match_end=0], [if_not_found=#N/A]). instance_num<0 counts from the end;
    // match_mode=1 is case-insensitive; match_end=1 lets the start/end of text
    // act as a match when the instance runs past the delimiters.
    private FormulaResult? EvalTextBeforeAfter(List<object> args, bool before)
    {
        if (args.Count < 2) return null;
        string text = args[0] is FormulaResult t ? t.AsString() : "";
        string delim = args[1] is FormulaResult d ? d.AsString() : "";
        int instance = args.Count > 2 && args[2] is FormulaResult i ? (int)i.AsNumber() : 1;
        bool ci = args.Count > 3 && args[3] is FormulaResult m && m.AsNumber() != 0;
        bool matchEnd = args.Count > 4 && args[4] is FormulaResult e && e.AsNumber() != 0;
        bool hasNotFound = args.Count > 5 && args[5] is FormulaResult;
        FormulaResult notFound = hasNotFound ? (FormulaResult)args[5] : FormulaResult.Error("#N/A");

        if (instance == 0) return FormulaResult.Error("#VALUE!");
        if (delim.Length == 0)
            // empty delimiter: split point is the very start (before) / end (after).
            return FR_S(before ? "" : text);

        var cmp = ci ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        // Collect delimiter positions in order. match_end=1 also lets the very
        // end of the text act as a zero-length delimiter, so an instance count
        // can reach the text boundary.
        var hits = new List<(int pos, int len)>();
        for (int k = 0; (k = text.IndexOf(delim, k, cmp)) >= 0; k += delim.Length)
            hits.Add((k, delim.Length));
        if (matchEnd) hits.Add((text.Length, 0));

        // instance_num is 1-based; negative counts from the last delimiter.
        int chosen = instance > 0 ? instance - 1 : hits.Count + instance;
        if (chosen < 0 || chosen >= hits.Count) return notFound;

        var (pos, len) = hits[chosen];
        return FR_S(before ? text[..pos] : text[(pos + len)..]);
    }

    private static System.Text.RegularExpressions.Regex BuildExcelRegex(string pattern, bool ci)
        => new(pattern, ci ? RegexOptions.IgnoreCase : RegexOptions.None);

    // REGEXTEST(text, pattern, [case_sensitivity=0]) — 1=case-insensitive.
    private FormulaResult? EvalRegexTest(List<object> args)
    {
        if (args.Count < 2) return null;
        string text = args[0] is FormulaResult t ? t.AsString() : "";
        string pat = args[1] is FormulaResult p ? p.AsString() : "";
        bool ci = args.Count > 2 && args[2] is FormulaResult c && c.AsNumber() != 0;
        try { return FR_B(BuildExcelRegex(pat, ci).IsMatch(text)); }
        catch (ArgumentException) { return FormulaResult.Error("#VALUE!"); }
    }

    // REGEXEXTRACT(text, pattern, [return_mode=0], [case_sensitivity=0]).
    // return_mode 0 = first whole match (scalar). Modes 1 (all matches) and 2
    // (capture groups) spill an array — deferred to the dynamic-array wave; we
    // return the first match for mode 0 and #VALUE! for the array modes so the
    // cache is never silently mispredicted.
    private FormulaResult? EvalRegexExtract(List<object> args)
    {
        if (args.Count < 2) return null;
        string text = args[0] is FormulaResult t ? t.AsString() : "";
        string pat = args[1] is FormulaResult p ? p.AsString() : "";
        int mode = args.Count > 2 && args[2] is FormulaResult mm ? (int)mm.AsNumber() : 0;
        bool ci = args.Count > 3 && args[3] is FormulaResult c && c.AsNumber() != 0;
        if (mode != 0) return null;   // array modes spill — not yet supported
        try
        {
            var m = BuildExcelRegex(pat, ci).Match(text);
            return m.Success ? FR_S(m.Value) : FormulaResult.Error("#N/A");
        }
        catch (ArgumentException) { return FormulaResult.Error("#VALUE!"); }
    }

    // REGEXREPLACE(text, pattern, replacement, [occurrence=0], [case_sensitivity=0]).
    // occurrence 0 = replace all; n>0 = replace only the nth match.
    private FormulaResult? EvalRegexReplace(List<object> args)
    {
        if (args.Count < 3) return null;
        string text = args[0] is FormulaResult t ? t.AsString() : "";
        string pat = args[1] is FormulaResult p ? p.AsString() : "";
        string rep = args[2] is FormulaResult r ? r.AsString() : "";
        int occ = args.Count > 3 && args[3] is FormulaResult o ? (int)o.AsNumber() : 0;
        bool ci = args.Count > 4 && args[4] is FormulaResult c && c.AsNumber() != 0;
        try
        {
            var rx = BuildExcelRegex(pat, ci);
            if (occ <= 0) return FR_S(rx.Replace(text, rep));
            int n = 0;
            return FR_S(rx.Replace(text, mtch => (++n == occ) ? mtch.Result(rep) : mtch.Value));
        }
        catch (ArgumentException) { return FormulaResult.Error("#VALUE!"); }
    }

    // ISREF(value) — TRUE when the argument is a reference. The parser hands a
    // bare ref token through as a RefArg; OFFSET/INDIRECT resolve to an Area
    // whose RangeData carries a workbook origin (BaseRow>0).
    private static FormulaResult EvalIsRef(List<object> args)
    {
        if (args.Count == 0) return FR_B(false);
        if (args[0] is RefArg) return FR_B(true);
        if (args[0] is FormulaResult { IsRange: true } fr) return FR_B(fr.RangeValue!.BaseRow > 0);
        return FR_B(false);
    }

    // ISFORMULA(reference) — TRUE when the referenced cell holds a formula.
    // Same-sheet only; a cross-sheet reference (RefArg.Sheet set) yields #N/A
    // rather than a wrong answer, since the evaluator probes the current sheet.
    private FormulaResult EvalIsFormula(List<object> args)
    {
        string? sheet; int col, row;
        switch (args.Count > 0 ? args[0] : null)
        {
            case RefArg ra: sheet = ra.Sheet; col = ra.Col; row = ra.Row; break;
            case FormulaResult { IsRange: true } fr when fr.RangeValue!.BaseRow > 0:
                sheet = fr.RangeValue.BaseSheet; col = fr.RangeValue.BaseCol; row = fr.RangeValue.BaseRow; break;
            default: return FormulaResult.Error("#VALUE!");
        }
        if (!string.IsNullOrEmpty(sheet)) return FormulaResult.Error("#N/A");
        var cell = FindCell($"{IndexToCol(col)}{row}");
        return FR_B(cell?.CellFormula?.Text != null);
    }

    private FormulaResult? EvalText(List<object> args)
    {
        var val = args.Count > 0 && args[0] is FormulaResult r ? r.AsNumber() : 0;
        var fmt = args.Count > 1 && args[1] is FormulaResult r2 ? r2.AsString() : "0";
        // Route through the cell renderer's number-format engine when wired up so
        // date/time/percent/currency format codes apply identically to a cell with
        // that numFmt (e.g. TEXT(45580,"yyyy-mm-dd") -> "2024-10-15"). Falls back
        // to the numeric-only .NET ToString path when no provider is registered.
        if (NumberFormatProvider != null)
        {
            try { return FR_S(NumberFormatProvider(val, fmt)); }
            catch { /* fall through to numeric ToString below */ }
        }
        try { return FR_S(val.ToString(fmt.Replace("#", "0"), CultureInfo.InvariantCulture)); }
        catch { return FR_S(val.ToString(CultureInfo.InvariantCulture)); }
    }

    private static FormulaResult? EvalFixed(List<object> args)
    {
        var v = args.Count > 0 && args[0] is FormulaResult r ? r.AsNumber() : 0;
        var d = args.Count > 1 && args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 2;
        // Third arg suppresses the thousands separator when TRUE.
        bool noCommas = args.Count > 2 && args[2] is FormulaResult r3 && r3.AsNumber() != 0;
        // Negative decimals round to the left of the decimal point.
        double factor = Math.Pow(10, d);
        double rounded = Math.Round(v * factor, MidpointRounding.AwayFromZero) / factor;
        string fmt = (noCommas ? "F" : "N") + Math.Max(0, d);
        return FR_S(rounded.ToString(fmt, CultureInfo.InvariantCulture));
    }

    // DOLLAR(number, [decimals=2]) / YEN(number, [decimals=0]): round to the
    // given decimals (negative rounds left of the point), group by thousands,
    // prefix the currency symbol, and wrap negatives in parentheses, matching
    // Excel. The old code used .NET's "C" format, which emits the generic ¤
    // symbol and ignored the decimals argument.
    private static FormulaResult EvalCurrency(List<object> args, string symbol, int defaultDec)
    {
        double v = args.Count > 0 && args[0] is FormulaResult r ? r.AsNumber() : 0;
        int dec = args.Count > 1 && args[1] is FormulaResult d ? (int)d.AsNumber() : defaultDec;
        double factor = Math.Pow(10, dec);
        double rounded = Math.Round(v * factor, MidpointRounding.AwayFromZero) / factor;
        string body = symbol + Math.Abs(rounded).ToString("N" + Math.Max(0, dec), CultureInfo.InvariantCulture);
        return FR_S(rounded < 0 ? $"({body})" : body);
    }

    private static FormulaResult? EvalNumberValue(List<object> args)
    {
        var s = args.Count > 0 && args[0] is FormulaResult r ? r.AsString() : "";
        // Excel NUMBERVALUE(text, [decimal_sep], [group_sep]) uses the first
        // char of each separator; defaults match the en-US locale (. and ,).
        string dec = args.Count > 1 && args[1] is FormulaResult d && d.AsString() != "" ? d.AsString() : ".";
        string grp = args.Count > 2 && args[2] is FormulaResult g && g.AsString() != "" ? g.AsString() : ",";
        s = s.Trim().Replace(" ", "");
        // currency symbols parse in Excel ($5, -$5, 5€); InvariantCulture's
        // symbol is ¤, so strip the common ones around the number ourselves
        s = Regex.Replace(s, @"^([+-]?)[\$€¥£]+", "$1");
        s = Regex.Replace(s, @"[\$€¥£]+$", "");
        s = s.Replace(grp[0].ToString(), "");        // drop group separators
        s = s.Replace(dec[0].ToString(), ".");       // decimal separator -> '.'
        int pct = 0;
        while (s.EndsWith("%")) { pct++; s = s[..^1]; }   // trailing % scales by 1/100
        return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)
            ? FR(v / Math.Pow(100, pct))
            : FormulaResult.Error("#VALUE!");
    }

    private FormulaResult? EvalTextJoin(List<object> args)
    {
        if (args.Count < 3) return null;
        var delim = args[0] is FormulaResult r ? r.AsString() : "";
        var ignoreEmpty = args[1] is FormulaResult r2 && r2.AsNumber() != 0;
        var parts = new List<string>();
        for (int i = 2; i < args.Count; i++)
        {
            if (AsRangeData(args[i]) is { } rd2)
            {
                for (int row = 0; row < rd2.Rows; row++)
                    for (int col = 0; col < rd2.Cols; col++)
                    {
                        var cv = rd2.Cells[row, col];
                        if (cv == null) continue;
                        if (cv.IsError) return cv;   // any error in the text propagates (matches Excel)
                        var s = cv.AsString(); if (!ignoreEmpty || s != "") parts.Add(s);
                    }
            }
            else if (args[i] is double[] arr) foreach (var v in arr) parts.Add(v.ToString(CultureInfo.InvariantCulture));
            else if (args[i] is FormulaResult fr) { if (fr.IsError) return fr; var s = fr.AsString(); if (!ignoreEmpty || s != "") parts.Add(s); }
        }
        return FR_S(string.Join(delim, parts));
    }

    // CONCAT / CONCATENATE — join every argument's text, but any error among the
    // arguments propagates as that error (matches Excel; text functions are not
    // error-swallowing).
    private FormulaResult? EvalConcat(List<object> args)
    {
        var all = AllArgs(args);
        var err = all.FirstOrDefault(r => r.IsError);
        if (err != null) return err;
        return FR_S(string.Concat(all.Select(r => r.AsString())));
    }

    // VALUETOTEXT(value, [format]) — format 0 (concise, default) returns the
    // value as plain text; format 1 (strict) double-quotes text (doubling any
    // embedded quote) but leaves numbers/booleans/errors unquoted. An array arg
    // collapses to its top-left cell (the anchor path Excel spills from).
    private FormulaResult? EvalValueToText(List<object> args)
    {
        if (args.Count < 1 || args[0] is not FormulaResult v) return FormulaResult.Error("#VALUE!");
        if (v.IsRange) v = v.RangeValue is { Rows: > 0, Cols: > 0 } rd ? rd.Cells[0, 0] ?? FormulaResult.Blank() : FormulaResult.Blank();
        bool strict = args.Count > 1 && args[1] is FormulaResult f && (int)f.AsNumber() == 1;
        return FR_S(ValueToTextScalar(v, strict));
    }

    private static string ValueToTextScalar(FormulaResult? v, bool strict)
    {
        if (v == null || v.IsBlank) return "";
        if (v.IsError) return v.ErrorValue!;
        if (v.IsString) return strict ? "\"" + v.StringValue!.Replace("\"", "\"\"") + "\"" : v.StringValue!;
        return v.AsString();   // numbers, booleans — never quoted
    }

    // ARRAYTOTEXT(array, [format]) — format 0 (concise, default) lists the cells
    // row-major, comma-space separated, values unquoted; format 1 (strict) emits
    // the Excel array-literal "{a,b;c,d}" (comma = column sep, semicolon = row
    // sep, text double-quoted) which round-trips back to an array constant.
    private FormulaResult? EvalArrayToText(List<object> args)
    {
        if (ToGrid(args.Count > 0 ? args[0] : null) is not { } g) return FormulaResult.Error("#VALUE!");
        bool strict = args.Count > 1 && args[1] is FormulaResult f && (int)f.AsNumber() == 1;
        int rows = g.GetLength(0), cols = g.GetLength(1);
        if (strict)
        {
            var rowStrs = new List<string>();
            for (int r = 0; r < rows; r++)
            {
                var cells = new List<string>();
                for (int c = 0; c < cols; c++) cells.Add(ValueToTextScalar(g[r, c], strict: true));
                rowStrs.Add(string.Join(",", cells));
            }
            return FR_S("{" + string.Join(";", rowStrs) + "}");
        }
        var flat = new List<string>();
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++) flat.Add(ValueToTextScalar(g[r, c], strict: false));
        return FR_S(string.Join(", ", flat));
    }

    // Element-wise scalar functions that lift over array arguments (implicit
    // spilling). Every listed function takes only scalar arguments and returns a
    // scalar — array-consuming functions (aggregates, lookups, dynamic arrays)
    // are deliberately absent so they keep handling arrays themselves.
    private static readonly HashSet<string> LiftableScalarFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        // text
        "UPPER", "LOWER", "PROPER", "TRIM", "CLEAN", "LEN", "LENB", "LEFT", "RIGHT",
        "MID", "CHAR", "CODE", "UNICHAR", "UNICODE", "REPT", "VALUE", "NUMBERVALUE",
        "TEXT", "FIXED", "EXACT", "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "T",
        "ASC", "TEXTBEFORE", "TEXTAFTER", "DOLLAR", "YEN",
        // math
        "ABS", "SIGN", "INT", "TRUNC", "ROUND", "ROUNDUP", "ROUNDDOWN", "MROUND",
        "CEILING", "CEILING_MATH", "CEILING_PRECISE", "ISO_CEILING", "FLOOR",
        "FLOOR_MATH", "FLOOR_PRECISE", "EVEN", "ODD", "SQRT", "SQRTPI", "EXP", "LN",
        "LOG", "LOG10", "POWER", "MOD", "QUOTIENT", "FACT", "FACTDOUBLE", "DEGREES",
        "RADIANS", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "ATAN2", "SINH",
        "COSH", "TANH", "ASINH", "ACOSH", "ATANH", "SEC", "CSC", "COT", "SECH",
        "CSCH", "COTH", "ACOT", "ACOTH", "GAMMA", "GAMMALN", "ERF", "ERFC", "DELTA",
        "GESTEP", "ARABIC", "ROMAN", "COMBIN", "PERMUT",
        // logical / info
        "NOT", "ISNUMBER", "ISBLANK", "ISTEXT", "ISNONTEXT", "ISLOGICAL", "ISERR",
        "ISERROR", "ISNA", "ISEVEN", "ISODD", "ISREF", "ISFORMULA", "N", "TYPE",
        "ERROR_TYPE", "CHOOSE", "IFERROR", "IFNA", "IF",
        // date / time
        "DAY", "MONTH", "YEAR", "HOUR", "MINUTE", "SECOND", "WEEKDAY", "WEEKNUM",
        "ISOWEEKNUM", "EDATE", "EOMONTH", "DATEVALUE", "TIMEVALUE", "DATE", "TIME",
        "DATEDIF", "DAYS",
        // bitwise / base conversion
        "BITAND", "BITOR", "BITXOR", "BITLSHIFT", "BITRSHIFT", "DEC2BIN", "DEC2HEX",
        "DEC2OCT", "BIN2DEC", "BIN2HEX", "BIN2OCT", "HEX2BIN", "HEX2DEC", "HEX2OCT",
        "OCT2BIN", "OCT2DEC", "OCT2HEX",
    };

    // Map a liftable scalar function over its array arguments, broadcasting
    // scalars and vectors against the result shape. Returns null when no
    // argument is a multi-cell array (so the caller runs the normal scalar path).
    private FormulaResult? TryLiftOverArrays(string name, List<object> args)
    {
        var grids = new FormulaResult?[args.Count][,];
        int rows = 1, cols = 1;
        bool anyArray = false;
        for (int k = 0; k < args.Count; k++)
        {
            var g = ToGrid(args[k]);
            if (g != null && (g.GetLength(0) > 1 || g.GetLength(1) > 1))
            {
                grids[k] = g;
                anyArray = true;
                rows = Math.Max(rows, g.GetLength(0));
                cols = Math.Max(cols, g.GetLength(1));
            }
        }
        if (!anyArray) return null;

        var outc = new FormulaResult?[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                var elemArgs = new List<object>(args.Count);
                bool oob = false;
                for (int k = 0; k < args.Count; k++)
                {
                    if (grids[k] is { } g)
                    {
                        int gr = g.GetLength(0), gc = g.GetLength(1);
                        int rr = gr == 1 ? 0 : r, cc = gc == 1 ? 0 : c;   // vector broadcast
                        if (rr >= gr || cc >= gc) { oob = true; break; }  // shorter array pads #N/A
                        elemArgs.Add(g[rr, cc] ?? FormulaResult.Blank());
                    }
                    else elemArgs.Add(args[k]);
                }
                outc[r, c] = oob ? FormulaResult.Error("#N/A")
                                 : EvalFunction(name, elemArgs) ?? FormulaResult.Error("#VALUE!");
            }
        return MakeArea(outc);
    }

    // ==================== Lookup ====================

    private FormulaResult? EvalIndex(List<object> args)
    {
        if (args.Count < 2) return null;
        if (AsRangeData(args[0]) is { } rd)
        {
            var rowIdx = args[1] is FormulaResult r ? (int)r.AsNumber() : 0;
            var colIdx = args.Count > 2 && args[2] is FormulaResult c ? (int)c.AsNumber() : 1;
            // row_num / col_num = 0 selects the whole column / row as an array.
            if (rowIdx == 0 && colIdx == 0) return MakeArea(rd.Cells);
            if (rowIdx == 0)
            {
                if (colIdx < 1 || colIdx > rd.Cols) return FormulaResult.Error("#REF!");
                var col = new FormulaResult?[rd.Rows, 1];
                for (int i = 0; i < rd.Rows; i++) col[i, 0] = rd.Cells[i, colIdx - 1];
                return MakeArea(col);
            }
            if (colIdx == 0)
            {
                if (rowIdx < 1 || rowIdx > rd.Rows) return FormulaResult.Error("#REF!");
                var rowArr = new FormulaResult?[1, rd.Cols];
                for (int j = 0; j < rd.Cols; j++) rowArr[0, j] = rd.Cells[rowIdx - 1, j];
                return MakeArea(rowArr);
            }
            if (rowIdx < 1 || rowIdx > rd.Rows || colIdx < 1 || colIdx > rd.Cols) return FormulaResult.Error("#REF!");
            return rd.Cells[rowIdx - 1, colIdx - 1] ?? FormulaResult.Number(0);
        }
        if (AsDoubles(args[0]) is { } arr)
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
        int matchType = args.Count > 2 && args[2] is FormulaResult mt ? (int)mt.AsNumber() : 1;

        // Materialize the 1-D lookup vector (either range orientation, or array).
        var cells = new List<FormulaResult?>();
        if (AsRangeData(args[1]) is { } rd)
        {
            if (rd.Cols == 1) for (int i = 0; i < rd.Rows; i++) cells.Add(rd.Cells[i, 0]);
            else for (int i = 0; i < rd.Cols; i++) cells.Add(rd.Cells[0, i]);
        }
        else if (AsDoubles(args[1]) is { } arr)
            foreach (var v in arr) cells.Add(FormulaResult.Number(v));
        else return FormulaResult.Error("#N/A");

        if (matchType == 0)
        {
            // Exact match supports wildcards (* and ?) when the lookup is text.
            bool wild = lookup.IsString && (lookup.StringValue!.Contains('*') || lookup.StringValue.Contains('?'));
            var rx = wild ? "^" + WildcardToRegex(lookup.AsString()) + "$" : null;
            for (int i = 0; i < cells.Count; i++)
                if (cells[i] != null && (wild ? Regex.IsMatch(cells[i]!.AsString(), rx!, RegexOptions.IgnoreCase)
                                              : CompareValues(cells[i]!, lookup) == 0)) return FR(i + 1);
            return FormulaResult.Error("#N/A");
        }
        // Approximate: matchType 1 = largest value <= lookup (ascending data);
        // matchType -1 = smallest value >= lookup (descending data).
        int found = -1;
        for (int i = 0; i < cells.Count; i++)
        {
            if (cells[i] == null) continue;
            int cmp = CompareValues(cells[i]!, lookup);
            if (cmp == 0) return FR(i + 1);
            if (matchType == 1) { if (cmp < 0) found = i; else break; }
            else { if (cmp > 0) found = i; else break; }
        }
        return found >= 0 ? FR(found + 1) : FormulaResult.Error("#N/A");
    }

    private FormulaResult? EvalRowCol(List<object> args, bool isRow)
    {
        // No argument: the row/column of the cell holding the formula, when the
        // evaluating cell context is known.
        if (args.Count == 0) return _ctxRow > 0 ? FR(isRow ? _ctxRow : _ctxCol) : null;
        // A single-cell reference reaches here as a RefArg (see ParseFunction's
        // ref-preserving list) — read its row/column directly, 1-based.
        if (args[0] is RefArg ra) return FR(isRow ? ra.Row : ra.Col);
        // OFFSET / INDIRECT / ranges produce a FormulaResult.Area whose underlying
        // RangeData carries the resolved reference's top-left origin. Use that
        // when present so ROW(OFFSET(A1,2,0)) reports 3 (not the cell value's row).
        // For *computed* arrays (BaseRow=0 — array constants like {1,2,3} or
        // arithmetic results like A1:A3*2) there is no workbook origin. Return
        // #REF! rather than silently null: a transient array has no row/column identity.
        if (AsRangeData(args[0]) is { } rd)
        {
            if (rd.BaseRow > 0) return FR(isRow ? rd.BaseRow : rd.BaseCol);
            return FormulaResult.Error("#REF!");
        }
        if (args[0] is FormulaResult { IsArray: true }) return FormulaResult.Error("#REF!");
        if (args[0] is FormulaResult r)
        { var m = Regex.Match(r.AsString(), @"([A-Z]+)(\d+)", RegexOptions.IgnoreCase);
          return m.Success ? FR(isRow ? int.Parse(m.Groups[2].Value) : ColToIndex(m.Groups[1].Value)) : null; }
        return null;
    }

    private static FormulaResult? EvalRowsCols(List<object> args, bool isRows)
    {
        if (args.Count > 0 && args[0] is FormulaResult { IsError: true } e) return e;   // ROWS/COLUMNS(#N/A) → #N/A
        if (args.Count > 0 && AsRangeData(args[0]) is { } rd) return FR(isRows ? rd.Rows : rd.Cols);
        if (args.Count > 0 && AsDoubles(args[0]) is { } arr) return FR(arr.Length);
        return FR(1);
    }

    private FormulaResult? EvalVlookup(List<object> args)
    {
        if (args.Count < 3) return null;
        var lookupVal = args[0] is FormulaResult r ? r : null; if (lookupVal == null) return null;
        var table = AsRangeData(args[1]); if (table == null) return FormulaResult.Error("#N/A");
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
        var table = AsRangeData(args[1]); if (table == null) return FormulaResult.Error("#N/A");
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
        var lv = AsRangeData(args[1]);
        if (lv == null) return FormulaResult.Error("#N/A");

        // Vector form (1D): optionally with a parallel result_vector
        if (lv.Rows == 1 || lv.Cols == 1)
        {
            int found = ApproximateMatchVector(lv, lookupVal);
            if (found < 0) return FormulaResult.Error("#N/A");

            var resultVec = args.Count >= 3 && AsRangeData(args[2]) is { } rv ? rv : lv;
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
        var lookupArr = AsRangeData(args[1]);
        var returnArr = AsRangeData(args[2]);
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
            if (matchMode == 2)
            {
                // Wildcard match against the whole cell text (?/*).
                if (Regex.IsMatch(cell.AsString(), "^" + WildcardToRegex(lookupVal.AsString()) + "$", RegexOptions.IgnoreCase))
                { found = i; break; }
                continue;
            }
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

        // Pull from return_array at `found`. When return_array is 2D, XLOOKUP
        // returns the whole matching line (a vertical lookup yields the found
        // row across all columns; a horizontal lookup yields the found column
        // across all rows), not just its first cell.
        if (isRow)
        {
            if (found >= returnArr.Cols) return FormulaResult.Error("#N/A");
            if (returnArr.Rows == 1) return returnArr.Cells[0, found] ?? FormulaResult.Number(0);
            var colCells = new FormulaResult?[returnArr.Rows, 1];
            for (int rr = 0; rr < returnArr.Rows; rr++) colCells[rr, 0] = returnArr.Cells[rr, found];
            return MakeArea(colCells);
        }
        else
        {
            if (found >= returnArr.Rows) return FormulaResult.Error("#N/A");
            if (returnArr.Cols == 1) return returnArr.Cells[found, 0] ?? FormulaResult.Number(0);
            var rowCells = new FormulaResult?[1, returnArr.Cols];
            for (int cc = 0; cc < returnArr.Cols; cc++) rowCells[0, cc] = returnArr.Cells[found, cc];
            return MakeArea(rowCells);
        }
    }

    private static FormulaResult? EvalAddress(List<object> args)
    {
        if (args.Count < 2) return null;
        var row = (int)(args[0] is FormulaResult r ? r.AsNumber() : 1);
        var col = (int)(args[1] is FormulaResult r2 ? r2.AsNumber() : 1);
        var abs = args.Count > 2 && args[2] is FormulaResult r3 ? (int)r3.AsNumber() : 1;
        // 4th arg selects A1 (default/TRUE) vs R1C1 (FALSE); 5th is a sheet name.
        bool a1 = !(args.Count > 3 && args[3] is FormulaResult r4 && r4.AsNumber() == 0);
        string sheet = args.Count > 4 && args[4] is FormulaResult r5 ? r5.AsString() : "";
        string addr;
        if (a1)
        {
            var cs = IndexToCol(col);
            addr = abs switch { 1 => $"${cs}${row}", 2 => $"{cs}${row}", 3 => $"${cs}{row}", _ => $"{cs}{row}" };
        }
        else
        {
            string rPart = abs is 1 or 2 ? $"R{row}" : $"R[{row}]";
            string cPart = abs is 1 or 3 ? $"C{col}" : $"C[{col}]";
            addr = rPart + cPart;
        }
        return FR_S(string.IsNullOrEmpty(sheet) ? addr : sheet + "!" + addr);
    }

    // SHEET([value]) — 1-based position of a sheet in workbook tab order. No arg
    // = the sheet holding the formula; a reference = that ref's sheet; a text
    // name = the named sheet.
    private FormulaResult? EvalSheet(List<object> args)
    {
        var sheets = _workbookPart?.Workbook?
            .Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList();
        if (sheets == null) return args.Count == 0 ? FR(1) : null;

        FormulaResult? Current() { var i = CurrentSheetIndex(sheets); return FR(i > 0 ? i : 1); }
        if (args.Count == 0) return Current();

        string? name = args[0] switch
        {
            RefArg ra => string.IsNullOrEmpty(ra.Sheet) ? null : ra.Sheet,
            FormulaResult { IsRange: true } fr => string.IsNullOrEmpty(fr.RangeValue!.BaseSheet) ? null : fr.RangeValue.BaseSheet,
            FormulaResult r => r.AsString(),
            _ => null,
        };
        if (name == null) return Current();   // same-sheet ref → current sheet
        for (int i = 0; i < sheets.Count; i++)
            if (string.Equals(sheets[i].Name?.Value, name, StringComparison.OrdinalIgnoreCase)) return FR(i + 1);
        return FormulaResult.Error("#N/A");
    }

    // Match _sheetData (the sheet being evaluated) to its tab position. The same
    // SheetData instance hangs off the owning WorksheetPart, so reference
    // equality identifies the current sheet without a name being threaded in.
    private int CurrentSheetIndex(List<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets)
    {
        if (_workbookPart == null) return 0;
        for (int i = 0; i < sheets.Count; i++)
        {
            try
            {
                var wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)_workbookPart!.GetPartById(sheets[i].Id!.Value!);
                var sheetData = wsPart.Worksheet?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                if (sheetData != null && ReferenceEquals(sheetData, _sheetData))
                    return i + 1;
            }
            catch { /* malformed rel — skip */ }
        }
        return 0;
    }

    // SHEETS([reference]) — number of sheets. No arg = total in the workbook; a
    // single-area reference spans one sheet (3D references are not modeled).
    private FormulaResult? EvalSheets(List<object> args)
    {
        var count = _workbookPart?.Workbook?
            .Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() ?? 1;
        return args.Count == 0 ? FR(count) : FR(1);
    }

    // CELL(info_type, [reference]) — deterministic subtypes only. address/row/
    // col/contents/type are computed; format/color/protect/width/prefix/filename
    // depend on cell formatting or the file path the evaluator does not model and
    // return null (cache stays unverified rather than guessed). Reference is
    // required (the "last changed cell" default is non-deterministic).
    private FormulaResult? EvalCell(List<object> args)
    {
        if (args.Count == 0) return null;
        string info = (args[0] is FormulaResult t ? t.AsString() : "").ToLowerInvariant();

        string? sheet = null; int col = 0, row = 0; bool haveRef = false;
        if (args.Count > 1)
            switch (args[1])
            {
                case RefArg ra: sheet = ra.Sheet; col = ra.Col; row = ra.Row; haveRef = true; break;
                case FormulaResult { IsRange: true } fr when fr.RangeValue!.BaseRow > 0:
                    sheet = fr.RangeValue.BaseSheet; col = fr.RangeValue.BaseCol; row = fr.RangeValue.BaseRow; haveRef = true; break;
            }
        if (!haveRef) return null;

        FormulaResult? Inner()
        {
            var a = ResolveRef(new RefArg(sheet, col, row, 1, 1));
            return a is { IsRange: true } area ? area.RangeValue!.Cells[0, 0] : a;
        }
        switch (info)
        {
            case "address": return FR_S($"${IndexToCol(col)}${row}");
            case "row": return FR(row);
            case "col": return FR(col);
            case "contents": return Inner() ?? FR(0);
            case "type":
                var v = Inner();
                if (v == null || v.IsBlank || (v.IsString && v.AsString() == "")) return FR_S("b");
                return FR_S(v.IsString ? "l" : "v");
            default: return null;   // unsupported subtype — leave cache unverified
        }
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
        int maxCount = v.GroupBy(x => x).Max(g => g.Count());
        if (maxCount <= 1) return FormulaResult.Error("#N/A");
        // On a count tie Excel returns the value that appears first in the data,
        // not the numerically smallest.
        var tied = v.GroupBy(x => x).Where(g => g.Count() == maxCount).Select(g => g.Key).ToHashSet();
        foreach (var x in v) if (tied.Contains(x)) return FR(x);
        return FormulaResult.Error("#N/A");
    }

    private static FormulaResult? EvalLarge(List<object> args)
    {
        var arr = args.Count > 0 ? AsDoubles(args[0]) : null;
        var k = args.Count > 1 && args[1] is FormulaResult r ? (int)r.AsNumber() : 1;
        if (arr == null || k < 1 || k > arr.Length) return FormulaResult.Error("#NUM!");
        return FR(arr.OrderByDescending(x => x).ElementAt(k - 1));
    }

    private static FormulaResult? EvalSmall(List<object> args)
    {
        var arr = args.Count > 0 ? AsDoubles(args[0]) : null;
        var k = args.Count > 1 && args[1] is FormulaResult r ? (int)r.AsNumber() : 1;
        if (arr == null || k < 1 || k > arr.Length) return FormulaResult.Error("#NUM!");
        return FR(arr.OrderBy(x => x).ElementAt(k - 1));
    }

    private static FormulaResult? EvalRank(List<object> args)
    {
        if (args.Count < 2) return null;
        var val = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var arr = AsDoubles(args[1]); if (arr == null) return null;
        var order = args.Count > 2 && args[2] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        var sorted = order == 0 ? arr.OrderByDescending(x => x).ToArray() : arr.OrderBy(x => x).ToArray();
        for (int i = 0; i < sorted.Length; i++) if (Math.Abs(sorted[i] - val) < 1e-10) return FR(i + 1);
        return FormulaResult.Error("#N/A");
    }

    private static FormulaResult? EvalPercentile(List<object> args)
    {
        var arr = args.Count > 0 ? AsDoubles(args[0]) : null;
        var k = args.Count > 1 && args[1] is FormulaResult r ? r.AsNumber() : 0;
        if (arr == null || arr.Length == 0 || k < 0 || k > 1) return FormulaResult.Error("#NUM!");
        var sorted = arr.OrderBy(x => x).ToArray();
        var idx = k * (sorted.Length - 1); var lower = (int)Math.Floor(idx); var upper = Math.Min(lower + 1, sorted.Length - 1);
        return FR(sorted[lower] + (idx - lower) * (sorted[upper] - sorted[lower]));
    }

    private static FormulaResult? EvalPercentRank(List<object> args)
    {
        var arr = args.Count > 0 ? AsDoubles(args[0]) : null;
        var val = args.Count > 1 && args[1] is FormulaResult r ? r.AsNumber() : 0;
        int sig = args.Count > 2 && args[2] is FormulaResult s ? (int)s.AsNumber() : 3;
        if (arr == null || arr.Length == 0 || sig < 1) return FormulaResult.Error("#NUM!");
        var sorted = arr.OrderBy(x => x).ToArray();
        int n = sorted.Length;
        if (val < sorted[0] || val > sorted[n - 1]) return FormulaResult.Error("#N/A");
        int below = sorted.Count(x => x < val);
        double rank;
        if (below < n && sorted[below] == val)
            rank = (double)below / (n - 1);                       // exact match
        else                                                      // interpolate between neighbours
        {
            int i = below - 1;
            rank = (i + (val - sorted[i]) / (sorted[i + 1] - sorted[i])) / (n - 1);
        }
        // Excel truncates (not rounds) the percentage to `sig` significant digits.
        double result = rank;
        if (result != 0)
        {
            double mag = Math.Pow(10, sig - Math.Ceiling(Math.Log10(Math.Abs(result))));
            result = Math.Truncate(result * mag) / mag;
        }
        return FR(result);
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

    // Helper: accept a RangeData directly OR a FormulaResult.Area wrapping one.
    // OFFSET / INDIRECT return Area-typed FormulaResult for multi-cell results,
    // so any function that iterates cells must accept both forms.
    private static RangeData? AsRangeData(object? a)
    {
        if (a is RangeData rd) return rd;
        if (a is FormulaResult fr && fr.IsRange) return fr.RangeValue;
        return null;
    }

    // Helper: extract double[] from RangeData, FormulaResult.Area, FormulaResult.Array, or bare double[].
    // Area-aware so functions like LARGE/SMALL/RANK/PERCENTILE work over OFFSET/INDIRECT results.
    private static double[]? AsDoubles(object? a)
    {
        if (AsRangeData(a) is { } rd) return rd.ToDoubleArray();
        if (a is FormulaResult fr && fr.IsArray) return fr.ArrayValue;
        if (a is double[] arr) return arr;
        return null;
    }

    // Helper: extract FormulaResult?[] from RangeData OR FormulaResult.Area (preserves string values for criteria matching).
    private static FormulaResult?[]? AsResults(object? a)
    {
        if (AsRangeData(a) is { } rd) return rd.ToFlatResults();
        return null;
    }

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

    /// <summary>
    /// Pre-extract the (criteria-range, criteria) pairs of a *IFS call ONCE.
    /// Flattening inside the per-row loop re-copied the whole criteria range for
    /// every row scanned — O(rows²·criteria) element copies, the dominant cost on
    /// SUMIFS-heavy workbooks (issue #187). Returns null when any criteria range
    /// fails to resolve — the caller maps that to its own "no row matches" result,
    /// exactly as the per-row `cr == null → match=false` used to.
    /// </summary>
    private static List<(FormulaResult?[] Range, string Crit)>? ExtractCriteriaPairs(List<object> args, int start)
    {
        var pairs = new List<(FormulaResult?[], string)>();
        for (int c = start; c + 1 < args.Count; c += 2)
        {
            var cr = AsResults(args[c]);
            if (cr == null) return null;
            pairs.Add((cr, args[c + 1] is FormulaResult cv ? cv.AsString() : ""));
        }
        return pairs;
    }

    private static bool MatchesAllCriteria(List<(FormulaResult?[] Range, string Crit)> pairs, int i)
    {
        foreach (var (cr, crit) in pairs)
            if (i >= cr.Length || !MatchesCriteria(cr[i], crit)) return false;
        return true;
    }

    private FormulaResult? EvalSumIfs(List<object> args)
    {
        if (args.Count < 3) return null;
        var sumRange = AsResults(args[0]); if (sumRange == null) return null;
        var pairs = ExtractCriteriaPairs(args, 1);
        if (pairs == null) return FR(0); // unresolvable criteria range: no row matches
        double sum = 0;
        for (int i = 0; i < sumRange.Length; i++)
            if (MatchesAllCriteria(pairs, i))
            { var n = AsNumeric(sumRange[i]); if (n.HasValue) sum += n.Value; }
        return FR(sum);
    }

    private FormulaResult? EvalCountIf(List<object> args)
    {
        if (args.Count < 2) return null;
        var range = AsResults(args[0]);
        if (range == null) return null;
        // Array/range criterion — COUNTIF(A1:A5, A1:A5) returns one count per
        // criterion element, not a single scalar. This is the per-element form
        // Excel uses inside array math; e.g. the distinct-count idiom
        // SUMPRODUCT(1/COUNTIF(range,range)) needs it so `1/COUNTIF(...)`
        // broadcasts ([0.5,0.5,0.5,1,...]) instead of collapsing to 1/N. The
        // single-value criterion path below is unchanged (the common case).
        var critArr = AsResults(args[1]);
        if (critArr is { Length: > 1 })
        {
            var counts = critArr
                .Select(cv => (double)range.Count(v => MatchesCriteria(v, cv?.AsString() ?? "")))
                .ToArray();
            return FormulaResult.Array(counts);
        }
        var criteria = args[1] is FormulaResult c ? c.AsString() : "";
        return FR(range.Count(v => MatchesCriteria(v, criteria)));
    }

    private FormulaResult? EvalCountIfs(List<object> args)
    {
        if (args.Count < 2) return null;
        var first = AsResults(args[0]); if (first == null) return null;
        var pairs = ExtractCriteriaPairs(args, 0);
        if (pairs == null) return FR(0); // unresolvable criteria range: no row matches
        int count = 0;
        for (int i = 0; i < first.Length; i++)
            if (MatchesAllCriteria(pairs, i)) count++;
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
        var pairs = ExtractCriteriaPairs(args, 1);
        if (pairs == null) return FormulaResult.Error("#DIV/0!"); // unresolvable criteria range: no row matches
        var vals = new List<double>();
        for (int i = 0; i < avgRange.Length; i++)
            if (MatchesAllCriteria(pairs, i))
            { var n = AsNumeric(avgRange[i]); if (n.HasValue) vals.Add(n.Value); }
        return vals.Count > 0 ? FR(vals.Average()) : FormulaResult.Error("#DIV/0!");
    }

    private FormulaResult? EvalMaxMinIfs(List<object> args, bool isMax)
    {
        if (args.Count < 3) return null;
        var valRange = AsResults(args[0]); if (valRange == null) return null;
        var pairs = ExtractCriteriaPairs(args, 1);
        if (pairs == null) return FR(0); // unresolvable criteria range: no row matches
        var vals = new List<double>();
        for (int i = 0; i < valRange.Length; i++)
            if (MatchesAllCriteria(pairs, i))
            { var n = AsNumeric(valRange[i]); if (n.HasValue) vals.Add(n.Value); }
        return vals.Count > 0 ? FR(isMax ? vals.Max() : vals.Min()) : FR(0);
    }

    private FormulaResult? EvalSumProduct(List<object> args)
    {
        if (args.Count == 0) return FR(0);
        var arrays = args.Select(a => AsDoubles(a)).ToList();
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

    // DATE(year, month, day) with Excel's rollover: years 0..1899 map to 1900+year,
    // and out-of-range month/day roll into adjacent months/years rather than error.
    private static FormulaResult? EvalDate(double y, double m, double d)
    {
        int year = (int)y;
        if (year >= 0 && year < 1900) year += 1900;
        try { return FR(new DateTime(year, 1, 1).AddMonths((int)m - 1).AddDays((int)d - 1).ToOADate()); }
        catch { return FormulaResult.Error("#NUM!"); }
    }

    private static FormulaResult? EvalDateDif(List<object> args)
    {
        if (args.Count < 3) return null;
        var d1 = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var d2 = args[1] is FormulaResult r2 ? DateTime.FromOADate(r2.AsNumber()) : DateTime.Today;
        var unit = args[2] is FormulaResult r3 ? r3.AsString().ToUpperInvariant() : "D";
        if (d2 < d1) return FormulaResult.Error("#NUM!");
        // Decompose the span into complete years/months and a day remainder,
        // borrowing from the previous month when the day is negative.
        int years = d2.Year - d1.Year, months = d2.Month - d1.Month, days = d2.Day - d1.Day;
        if (days < 0)
        {
            months--;
            var prev = new DateTime(d2.Year, d2.Month, 1).AddDays(-1);
            days += DateTime.DaysInMonth(prev.Year, prev.Month);
        }
        if (months < 0) { months += 12; years--; }
        // YD: days ignoring whole years. Excel anchors in d1's year — d2's
        // month/day is placed into d1's year (or the following year when it
        // falls before d1's), so the leap day is counted exactly when the
        // start year is leap: DATEDIF(2020-01-15, 2022-03-20, "YD") = 65,
        // while a 2021 start gives 64.
        int ydYear = d2.Month > d1.Month || (d2.Month == d1.Month && d2.Day >= d1.Day) ? d1.Year : d1.Year + 1;
        int ydDay = Math.Min(d2.Day, DateTime.DaysInMonth(ydYear, d2.Month));
        var ydTarget = new DateTime(ydYear, d2.Month, ydDay);
        return unit switch
        {
            "Y" => FR(years),
            "M" => FR(years * 12 + months),
            "D" => FR((d2 - d1).Days),
            "MD" => FR(days),
            "YM" => FR(months),
            "YD" => FR((ydTarget - d1).Days),
            _ => FormulaResult.Error("#NUM!"),
        };
    }

    private static FormulaResult? EvalNetworkDays(List<object> args, bool intl)
    {
        if (args.Count < 2) return null;
        var start = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var end = args[1] is FormulaResult r2 ? DateTime.FromOADate(r2.AsNumber()) : DateTime.Today;
        // .INTL takes an optional weekend descriptor at arg 2; holidays follow.
        int holidayArg = 2;
        var isWeekend = intl ? ParseWeekend(args, 2, ref holidayArg) : DefaultWeekend;
        var holidays = CollectHolidays(args, holidayArg);
        int sign = end >= start ? 1 : -1;
        if (sign < 0) (start, end) = (end, start);
        int count = 0;
        for (var d = start; d <= end; d = d.AddDays(1))
            if (!isWeekend(d.DayOfWeek) && !holidays.Contains(d.Date)) count++;
        return FR(sign * count);
    }

    private static FormulaResult? EvalWorkDay(List<object> args, bool intl)
    {
        if (args.Count < 2) return null;
        var start = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var days = args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        int holidayArg = 2;
        var isWeekend = intl ? ParseWeekend(args, 2, ref holidayArg) : DefaultWeekend;
        var holidays = CollectHolidays(args, holidayArg);
        var d = start; var step = days > 0 ? 1 : -1; var rem = Math.Abs(days);
        while (rem > 0)
        {
            d = d.AddDays(step);
            if (!isWeekend(d.DayOfWeek) && !holidays.Contains(d.Date)) rem--;
        }
        return FR(d.ToOADate());
    }

    private static FormulaResult? EvalYearFrac(List<object> args)
    {
        if (args.Count < 2) return null;
        var d1 = args[0] is FormulaResult r1 ? DateTime.FromOADate(r1.AsNumber()) : DateTime.Today;
        var d2 = args[1] is FormulaResult r2 ? DateTime.FromOADate(r2.AsNumber()) : DateTime.Today;
        int basis = args.Count > 2 && args[2] is FormulaResult b ? (int)b.AsNumber() : 0;
        if (basis is < 0 or > 4) return FormulaResult.Error("#NUM!");
        return FR(Math.Abs(YearFracBasis(d1, d2, basis)));
    }

    // ==================== Financial ====================

    // Shared reader for the time-value-of-money annuity functions. `type` (0 =
    // payment at period end, 1 = at start) scales the payment leg by (1+rate).
    private static double TvmNum(List<object> args, int i, double def)
        => i < args.Count && args[i] is FormulaResult r ? r.AsNumber() : def;

    private static FormulaResult? EvalPmt(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = TvmNum(args, 0, 0), nper = TvmNum(args, 1, 0), pv = TvmNum(args, 2, 0),
               fv = TvmNum(args, 3, 0), type = TvmNum(args, 4, 0);
        if (rate == 0) return FR(-(pv + fv) / nper);
        double pow = Math.Pow(1 + rate, nper);
        return FR(-(pv * pow + fv) * rate / ((1 + rate * type) * (pow - 1)));
    }

    private static FormulaResult? EvalFv(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = TvmNum(args, 0, 0), nper = TvmNum(args, 1, 0), pmt = TvmNum(args, 2, 0),
               pv = TvmNum(args, 3, 0), type = TvmNum(args, 4, 0);
        if (rate == 0) return FR(-(pv + pmt * nper));
        double pow = Math.Pow(1 + rate, nper);
        return FR(-(pv * pow + pmt * (1 + rate * type) * (pow - 1) / rate));
    }

    private static FormulaResult? EvalPv(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = TvmNum(args, 0, 0), nper = TvmNum(args, 1, 0), pmt = TvmNum(args, 2, 0),
               fv = TvmNum(args, 3, 0), type = TvmNum(args, 4, 0);
        if (rate == 0) return FR(-(fv + pmt * nper));
        double pow = Math.Pow(1 + rate, nper);
        return FR(-(fv + pmt * (1 + rate * type) * (pow - 1) / rate) / pow);
    }

    private static FormulaResult? EvalNper(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = TvmNum(args, 0, 0), pmt = TvmNum(args, 1, 0), pv = TvmNum(args, 2, 0),
               fv = TvmNum(args, 3, 0), type = TvmNum(args, 4, 0);
        if (rate == 0) return pmt != 0 ? FR(-(pv + fv) / pmt) : null;
        double k = pmt * (1 + rate * type) / rate;
        return FR(Math.Log((k - fv) / (k + pv)) / Math.Log(1 + rate));
    }

    private static FormulaResult? EvalNpv(List<object> args)
    {
        if (args.Count < 2) return null;
        var rate = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var values = new List<double>();
        for (int i = 1; i < args.Count; i++) { if (AsDoubles(args[i]) is { } arr) values.AddRange(arr); else if (args[i] is FormulaResult fr) values.Add(fr.AsNumber()); }
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
        // Remaining balance before this period = principal grown by interest LESS
        // the payments already made. The payment term must be subtracted; adding
        // it (the old bug) only happened to be correct at per=1 where it is zero.
        var balanceBefore = pv * Math.Pow(1 + rate, per - 1) - pmt * (Math.Pow(1 + rate, per - 1) - 1) / rate;
        return FR(-(balanceBefore * rate));
    }

    private static FormulaResult? EvalPpmt(List<object> args)
    {
        if (args.Count < 4) return null;
        // PPMT(rate, per, nper, pv, ...) = PMT - IPMT, but PMT's signature is
        // (rate, nper, pv, ...) — drop the `per` argument when delegating to PMT
        // (the old code passed PPMT's args straight through, so PMT read `per`
        // as nper).
        var pmtArgs = new List<object> { args[0], args[2], args[3] };
        if (args.Count > 4) pmtArgs.Add(args[4]);
        if (args.Count > 5) pmtArgs.Add(args[5]);
        var pmt = EvalPmt(pmtArgs)?.AsNumber() ?? 0;
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
        double Num(int i, double def) => i < args.Count && args[i] is FormulaResult r ? r.AsNumber() : def;
        double cost = Num(0, 0), salvage = Num(1, 0), life = Num(2, 0);
        int period = (int)Num(3, 1);
        // month: number of months in the first year (default full 12). The first
        // and final (life+1) periods are prorated by month / (12 - month).
        double month = Num(4, 12);
        if (cost == 0) return FR(0);
        var rate = Math.Round(1 - Math.Pow(salvage / cost, 1.0 / life), 3);
        double accumulated = 0, dep = 0;
        for (int p = 1; p <= period; p++)
        {
            if (p == 1) dep = cost * rate * month / 12.0;
            else if (p == (int)life + 1) dep = (cost - accumulated) * rate * (12 - month) / 12.0;
            else dep = (cost - accumulated) * rate;
            accumulated += dep;
        }
        return FR(dep);
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

    // RATE(nper, pmt, pv, [fv], [type], [guess]) — periodic interest rate that
    // balances the time-value-of-money annuity equation. Solved via SolveRoot.
    private static FormulaResult? EvalRate(List<object> args)
    {
        if (args.Count < 3) return null;
        double Num(int i, double def) => i < args.Count && args[i] is FormulaResult r ? r.AsNumber() : def;
        double nper = Num(0, 0), pmt = Num(1, 0), pv = Num(2, 0), fv = Num(3, 0), type = Num(4, 0), guess = Num(5, 0.1);

        // f(r) = pv·(1+r)^n + pmt·(1+r·type)·((1+r)^n − 1)/r + fv. The /r term
        // has a removable singularity at r=0 whose limit is pmt·(1+r·type)·n;
        // use it near zero so the solver can pass cleanly through r=0.
        double F(double r)
        {
            double pow = Math.Pow(1 + r, nper);
            double annuity = Math.Abs(r) < 1e-12 ? nper : (pow - 1) / r;
            return pv * pow + pmt * (1 + r * type) * annuity + fv;
        }
        var root = SolveRoot(F, guess);
        return root.HasValue ? FR(root.Value) : FormulaResult.Error("#NUM!");
    }

    // IRR(values, [guess]) — rate making the NPV of the cashflow series zero.
    private static FormulaResult? EvalIrr(List<object> args)
    {
        if (args.Count < 1) return null;
        var cf = AsDoubles(args[0]);
        if (cf == null || cf.Length < 2) return FormulaResult.Error("#NUM!");
        double guess = args.Count > 1 && args[1] is FormulaResult g ? g.AsNumber() : 0.1;

        double F(double r)
        {
            double npv = 0;
            for (int i = 0; i < cf.Length; i++) npv += cf[i] / Math.Pow(1 + r, i);
            return npv;
        }
        var root = SolveRoot(F, guess);
        return root.HasValue ? FR(root.Value) : FormulaResult.Error("#NUM!");
    }

    // XNPV(rate, values, dates) — NPV over actual/365 day fractions from date[0].
    private static FormulaResult? EvalXnpv(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var values = AsDoubles(args[1]); var dates = AsDoubles(args[2]);
        if (values == null || dates == null || values.Length == 0 || values.Length != dates.Length)
            return FormulaResult.Error("#NUM!");
        double d0 = dates[0], npv = 0;
        for (int i = 0; i < values.Length; i++)
            npv += values[i] / Math.Pow(1 + rate, (dates[i] - d0) / 365.0);
        return FR(npv);
    }

    // XIRR(values, dates, [guess]) — rate making XNPV zero, via the shared solver.
    private static FormulaResult? EvalXirr(List<object> args)
    {
        if (args.Count < 2) return null;
        var values = AsDoubles(args[0]); var dates = AsDoubles(args[1]);
        if (values == null || dates == null || values.Length < 2 || values.Length != dates.Length)
            return FormulaResult.Error("#NUM!");
        double guess = args.Count > 2 && args[2] is FormulaResult g ? g.AsNumber() : 0.1;
        double d0 = dates[0];
        double F(double rate)
        {
            double npv = 0;
            for (int i = 0; i < values.Length; i++) npv += values[i] / Math.Pow(1 + rate, (dates[i] - d0) / 365.0);
            return npv;
        }
        var root = SolveRoot(F, guess);
        return root.HasValue ? FR(root.Value) : FormulaResult.Error("#NUM!");
    }

    // MIRR(values, finance_rate, reinvest_rate) — modified IRR.
    private static FormulaResult? EvalMirr(List<object> args)
    {
        if (args.Count < 3) return null;
        var cf = AsDoubles(args[0]);
        if (cf == null || cf.Length < 2) return FormulaResult.Error("#DIV/0!");
        double fin = args[1] is FormulaResult f ? f.AsNumber() : 0, rei = args[2] is FormulaResult r ? r.AsNumber() : 0;
        int n = cf.Length;
        double pvNeg = 0, fvPos = 0;
        for (int i = 0; i < n; i++)
        {
            if (cf[i] < 0) pvNeg += cf[i] / Math.Pow(1 + fin, i);
            else fvPos += cf[i] * Math.Pow(1 + rei, n - 1 - i);
        }
        if (pvNeg == 0 || fvPos == 0) return FormulaResult.Error("#DIV/0!");
        return FR(Math.Pow(-fvPos / pvNeg, 1.0 / (n - 1)) - 1);
    }

    // CUMIPMT / CUMPRINC(rate, nper, pv, start_period, end_period, type) — sum of
    // the per-period interest (or principal) over [start, end], reusing IPMT/PPMT.
    private static FormulaResult? EvalCumulative(List<object> args, bool principal)
    {
        if (args.Count < 6) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, nper = args[1] is FormulaResult r2 ? r2.AsNumber() : 0;
        double pv = args[2] is FormulaResult r3 ? r3.AsNumber() : 0;
        int start = args[3] is FormulaResult r4 ? (int)r4.AsNumber() : 0, end = args[4] is FormulaResult r5 ? (int)r5.AsNumber() : 0;
        double type = args[5] is FormulaResult r6 ? r6.AsNumber() : 0;
        if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper) return FormulaResult.Error("#NUM!");
        double total = 0;
        for (int per = start; per <= end; per++)
        {
            var perArgs = new List<object> { FR(rate), FR(per), FR(nper), FR(pv), FR(0), FR(type) };
            var v = principal ? EvalPpmt(perArgs) : EvalIpmt(perArgs);
            total += v?.AsNumber() ?? 0;
        }
        return FR(total);
    }

    // FVSCHEDULE(principal, schedule) — compound the principal by each rate.
    private static FormulaResult? EvalFvSchedule(List<object> args)
    {
        if (args.Count < 2) return null;
        double p = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var rates = AsDoubles(args[1]);
        if (rates == null) { if (args[1] is FormulaResult fr) rates = [fr.AsNumber()]; else return null; }
        foreach (var rate in rates) p *= 1 + rate;
        return FR(p);
    }

    // PDURATION(rate, pv, fv) — periods required to reach fv.
    private static FormulaResult? EvalPduration(List<object> args)
    {
        if (args.Count < 3) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, pv = args[1] is FormulaResult r2 ? r2.AsNumber() : 0, fv = args[2] is FormulaResult r3 ? r3.AsNumber() : 0;
        if (rate <= 0 || pv <= 0 || fv <= 0) return FormulaResult.Error("#NUM!");
        return FR((Math.Log(fv) - Math.Log(pv)) / Math.Log(1 + rate));
    }

    // RRI(nper, pv, fv) — equivalent periodic interest rate.
    private static FormulaResult? EvalRri(List<object> args)
    {
        if (args.Count < 3) return null;
        double nper = args[0] is FormulaResult r ? r.AsNumber() : 0, pv = args[1] is FormulaResult r2 ? r2.AsNumber() : 0, fv = args[2] is FormulaResult r3 ? r3.AsNumber() : 0;
        if (nper <= 0 || pv <= 0) return FormulaResult.Error("#NUM!");
        return FR(Math.Pow(fv / pv, 1.0 / nper) - 1);
    }

    // EFFECT(nominal_rate, npery) — effective annual interest rate.
    private static FormulaResult? EvalEffect(List<object> args)
    {
        if (args.Count < 2) return null;
        double nom = args[0] is FormulaResult r ? r.AsNumber() : 0; int npery = args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        if (nom <= 0 || npery < 1) return FormulaResult.Error("#NUM!");
        return FR(Math.Pow(1 + nom / npery, npery) - 1);
    }

    // NOMINAL(effect_rate, npery) — nominal annual interest rate.
    private static FormulaResult? EvalNominal(List<object> args)
    {
        if (args.Count < 2) return null;
        double eff = args[0] is FormulaResult r ? r.AsNumber() : 0; int npery = args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        if (eff <= 0 || npery < 1) return FormulaResult.Error("#NUM!");
        return FR(npery * (Math.Pow(1 + eff, 1.0 / npery) - 1));
    }

    // DOLLARDE / DOLLARFR — convert between a price quoted as a fraction and its
    // decimal form. The fractional part is read against the fraction's digit width.
    private static FormulaResult? EvalDollar(List<object> args, bool toDecimal)
    {
        if (args.Count < 2) return null;
        double dollar = args[0] is FormulaResult r ? r.AsNumber() : 0; int fraction = args[1] is FormulaResult r2 ? (int)r2.AsNumber() : 0;
        if (fraction < 0) return FormulaResult.Error("#NUM!");
        if (fraction == 0) return toDecimal ? FormulaResult.Error("#DIV/0!") : FR(dollar);
        double intPart = Math.Truncate(dollar);
        double frac = dollar - intPart;
        double pow = Math.Pow(10, Math.Ceiling(Math.Log10(fraction)));
        return toDecimal
            ? FR(intPart + frac * pow / fraction)             // fractional → decimal
            : FR(intPart + frac * fraction / pow);            // decimal → fractional
    }

    // ISPMT(rate, per, nper, pv) — interest for a period with even principal pay-down.
    private static FormulaResult? EvalIspmt(List<object> args)
    {
        if (args.Count < 4) return null;
        double rate = args[0] is FormulaResult r ? r.AsNumber() : 0, per = args[1] is FormulaResult r2 ? r2.AsNumber() : 0;
        double nper = args[2] is FormulaResult r3 ? r3.AsNumber() : 0, pv = args[3] is FormulaResult r4 ? r4.AsNumber() : 0;
        if (nper == 0) return FormulaResult.Error("#DIV/0!");
        return FR(pv * rate * (per / nper - 1));
    }

    // ==================== Database (Dxxx) ====================

    private enum DbAgg { Sum, Count, CountA, Average, Max, Min, Product, Get, StdevS, StdevP, VarS, VarP }

    // Dxxx(database, field, criteria): aggregate one column of a table over rows
    // matching a criteria block. database row 0 = field headers; criteria row 0 =
    // criteria field headers, rows 1..n = criteria sets (AND within a row, OR
    // across rows) — the standard Excel D-function contract.
    private static FormulaResult? EvalDatabase(List<object> args, DbAgg agg)
    {
        if (args.Count < 3) return null;
        var db = AsRangeData(args[0]);
        var crit = AsRangeData(args[2]);
        if (db == null || crit == null || db.Rows < 2) return null;
        // A criteria range must have at least one condition row below its header.
        if (crit.Rows < 2) return FormulaResult.Error("#VALUE!");

        // Resolve the aggregated column: numeric field = 1-based index, else
        // match a header (case-insensitive).
        int fieldCol = ResolveDbField(db, args[1]);
        if (fieldCol < 0) return FormulaResult.Error("#VALUE!");

        var matched = new List<FormulaResult?>();
        for (int r = 1; r < db.Rows; r++)
            if (DbRowMatches(db, r, crit))
                matched.Add(db.Cells[r, fieldCol]);

        var nums = matched.Where(c => c?.IsNumeric == true).Select(c => c!.NumericValue!.Value).ToList();

        switch (agg)
        {
            case DbAgg.Count: return FR(nums.Count);
            case DbAgg.CountA: return FR(matched.Count(c => c != null && !c.IsBlank && c.AsString() != ""));
            case DbAgg.Sum: return FR(nums.Sum());
            case DbAgg.Product: return FR(nums.Aggregate(1.0, (a, b) => a * b));
            case DbAgg.Average: return nums.Count > 0 ? FR(nums.Average()) : FormulaResult.Error("#DIV/0!");
            case DbAgg.Max: return nums.Count > 0 ? FR(nums.Max()) : FR(0);
            case DbAgg.Min: return nums.Count > 0 ? FR(nums.Min()) : FR(0);
            case DbAgg.Get:
                if (matched.Count == 0) return FormulaResult.Error("#VALUE!");
                if (matched.Count > 1) return FormulaResult.Error("#NUM!");
                return matched[0] ?? FR(0);
            // Reuse the shared sample/population helpers so D-stats stay
            // bit-identical to STDEV/STDEVP/VAR/VARP (same #DIV/0! guards).
            case DbAgg.StdevS: return EvalStdev(nums.ToArray(), sample: true);
            case DbAgg.StdevP: return EvalStdev(nums.ToArray(), sample: false);
            case DbAgg.VarS: return EvalVar(nums.ToArray(), sample: true);
            case DbAgg.VarP: return EvalVar(nums.ToArray(), sample: false);
        }
        return null;
    }

    // field can be a column header string or a 1-based column index.
    private static int ResolveDbField(RangeData db, object fieldArg)
    {
        var fr = fieldArg as FormulaResult;
        if (fr?.IsNumeric == true)
        {
            int idx = (int)fr.NumericValue!.Value - 1;
            return idx >= 0 && idx < db.Cols ? idx : -1;
        }
        string name = fr?.AsString() ?? "";
        for (int c = 0; c < db.Cols; c++)
            if (string.Equals(db.Cells[0, c]?.AsString() ?? "", name, StringComparison.OrdinalIgnoreCase))
                return c;
        return -1;
    }

    // A record matches if ANY criteria row is satisfied; a criteria row is
    // satisfied when EVERY non-empty criteria cell matches the same-named db column.
    private static bool DbRowMatches(RangeData db, int dbRow, RangeData crit)
    {
        for (int cr = 1; cr < crit.Rows; cr++)
        {
            bool rowOk = true;
            for (int cc = 0; cc < crit.Cols; cc++)
            {
                var critCell = crit.Cells[cr, cc];
                string critStr = critCell?.AsString() ?? "";
                if (critCell == null || critCell.IsBlank || critStr == "") continue;   // no constraint

                int dbCol = -1;
                string header = crit.Cells[0, cc]?.AsString() ?? "";
                for (int c = 0; c < db.Cols; c++)
                    if (string.Equals(db.Cells[0, c]?.AsString() ?? "", header, StringComparison.OrdinalIgnoreCase)) { dbCol = c; break; }
                if (dbCol < 0) { rowOk = false; break; }

                if (!MatchesCriteria(db.Cells[dbRow, dbCol], critStr)) { rowOk = false; break; }
            }
            if (rowOk) return true;
        }
        return false;
    }
}
