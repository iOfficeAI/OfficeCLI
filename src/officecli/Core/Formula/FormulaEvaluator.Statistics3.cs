// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

// W4d — statistical hypothesis tests (T.TEST / CHISQ.TEST / F.TEST / Z.TEST),
// each returning a p-value through the W4 distribution CDFs (RegIncBeta for the
// t/F tails, RegGammaP for chi-square, NormCdf for z). Verified against real
// Excel.
internal partial class FormulaEvaluator
{
    private static double Mean(double[] v) => v.Length == 0 ? 0 : v.Average();

    private static double SampleVar(double[] v)
    {
        if (v.Length < 2) return 0;
        double m = Mean(v);
        return v.Sum(x => (x - m) * (x - m)) / (v.Length - 1);
    }

    // T.TEST(array1, array2, tails, type) — paired (1), two-sample equal-var (2),
    // two-sample unequal-var/Welch (3).
    private FormulaResult? EvalTTest(List<object> args)
    {
        if (args.Count < 4) return FormulaResult.Error("#VALUE!");
        var a = AsDoubles(args[0]); var b = AsDoubles(args[1]);
        if (a is null || b is null) return FormulaResult.Error("#VALUE!");
        int tails = (int)SecNum(args, 2, 2), type = (int)SecNum(args, 3, 1);
        if (tails is not (1 or 2)) return FormulaResult.Error("#NUM!");

        double t, df;
        if (type == 1)
        {
            if (a.Length != b.Length || a.Length < 2) return FormulaResult.Error("#N/A");
            var d = a.Zip(b, (x, y) => x - y).ToArray();
            int n = d.Length;
            double sd = Math.Sqrt(SampleVar(d));
            if (sd == 0) return FormulaResult.Error("#DIV/0!");
            t = Mean(d) / (sd / Math.Sqrt(n));
            df = n - 1;
        }
        else
        {
            int n1 = a.Length, n2 = b.Length;
            if (n1 < 2 || n2 < 2) return FormulaResult.Error("#DIV/0!");
            double v1 = SampleVar(a), v2 = SampleVar(b), m1 = Mean(a), m2 = Mean(b);
            if (type == 2)
            {
                double sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2);
                t = (m1 - m2) / Math.Sqrt(sp * (1.0 / n1 + 1.0 / n2));
                df = n1 + n2 - 2;
            }
            else
            {
                double s = v1 / n1 + v2 / n2;
                t = (m1 - m2) / Math.Sqrt(s);
                df = s * s / (Math.Pow(v1 / n1, 2) / (n1 - 1) + Math.Pow(v2 / n2, 2) / (n2 - 1));
            }
        }
        double p = tails * TDistRightTail(Math.Abs(t), df);
        return FR(Math.Min(1.0, p));
    }

    // F.TEST(array1, array2) — two-tailed equality-of-variance p-value.
    private FormulaResult? EvalFTest(List<object> args)
    {
        var a = AsDoubles(args.Count > 0 ? args[0] : null);
        var b = AsDoubles(args.Count > 1 ? args[1] : null);
        if (a is null || b is null || a.Length < 2 || b.Length < 2) return FormulaResult.Error("#DIV/0!");
        double v1 = SampleVar(a), v2 = SampleVar(b);
        if (v1 == 0 || v2 == 0) return FormulaResult.Error("#DIV/0!");
        double f = v1 / v2;
        double rt = FDistRightTail(f, a.Length - 1, b.Length - 1);
        return FR(2 * Math.Min(rt, 1 - rt));
    }

    // CHISQ.TEST(actual_range, expected_range) — Pearson chi-square p-value.
    private FormulaResult? EvalChisqTest(List<object> args)
    {
        if (ToGrid(args.Count > 0 ? args[0] : null) is not { } act ||
            ToGrid(args.Count > 1 ? args[1] : null) is not { } exp)
            return FormulaResult.Error("#VALUE!");
        int rows = act.GetLength(0), cols = act.GetLength(1);
        if (exp.GetLength(0) != rows || exp.GetLength(1) != cols) return FormulaResult.Error("#N/A");
        double chi2 = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                double e = exp[r, c]?.AsNumber() ?? 0;
                if (e == 0) return FormulaResult.Error("#DIV/0!");
                double diff = (act[r, c]?.AsNumber() ?? 0) - e;
                chi2 += diff * diff / e;
            }
        int df = rows == 1 || cols == 1 ? rows * cols - 1 : (rows - 1) * (cols - 1);
        if (df < 1) return FormulaResult.Error("#N/A");
        return FR(RegGammaQ(df / 2.0, chi2 / 2.0));   // right tail
    }

    // Z.TEST(array, x, [sigma]) — one-tailed P(Z > z); sample stdev when sigma omitted.
    private FormulaResult? EvalZTest(List<object> args)
    {
        var a = AsDoubles(args.Count > 0 ? args[0] : null);
        if (a is null || a.Length < 1) return FormulaResult.Error("#N/A");
        double x = SecNum(args, 1, 0);
        double sigma = args.Count > 2 && args[2] is FormulaResult s && !s.IsBlank
            ? s.AsNumber() : Math.Sqrt(SampleVar(a));
        if (sigma == 0) return FormulaResult.Error("#DIV/0!");
        double z = (Mean(a) - x) / (sigma / Math.Sqrt(a.Length));
        return FR(NormCdf(-z));   // symmetry keeps the far tail (1-CDF cancels to 0)
    }
}
