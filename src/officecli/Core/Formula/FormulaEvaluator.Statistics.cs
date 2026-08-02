// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Statistical distribution wrappers ====================
    //
    // Thin wrappers over the special-function core (FormulaEvaluator.Special-
    // Functions.cs). Each takes the same arguments Excel does; the `cumulative`
    // flag selects CDF vs PDF. These are the normal- / gamma- / chi-squared- /
    // Poisson- / exponential-family functions whose engines are erf and the
    // regularized incomplete gamma integral.

    private static double N(List<object> a, int i, double def = 0) =>
        i < a.Count && a[i] is FormulaResult r ? r.AsNumber() : def;

    private static bool Cumulative(List<object> a, int i) =>
        i < a.Count && a[i] is FormulaResult r && r.AsNumber() != 0;

    // NORM.DIST(x, mean, sd, cumulative) / NORMDIST.
    private static FormulaResult? EvalNormDist(List<object> args)
    {
        if (args.Count < 4) return null;
        double x = N(args, 0), mean = N(args, 1), sd = N(args, 2);
        if (sd <= 0) return FormulaResult.Error("#NUM!");
        double z = (x - mean) / sd;
        return FR(Cumulative(args, 3) ? NormCdf(z) : NormPdf(z) / sd);
    }

    // NORM.S.DIST(z, cumulative).
    private static FormulaResult? EvalNormSDist(List<object> args)
    {
        if (args.Count < 1) return null;
        double z = N(args, 0);
        return FR(Cumulative(args, 1) ? NormCdf(z) : NormPdf(z));
    }

    // CONFIDENCE.NORM(alpha, sd, size) — half-width of the normal CI.
    private static FormulaResult? EvalConfidenceNorm(List<object> args)
    {
        if (args.Count < 3) return null;
        double alpha = N(args, 0), sd = N(args, 1), size = N(args, 2);
        if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) return FormulaResult.Error("#NUM!");
        return FR(InvNormCdf(1 - alpha / 2) * sd / Math.Sqrt(size));
    }

    // ERF(lower, [upper]) — error function, optionally over an interval.
    private static FormulaResult? EvalErf(List<object> args)
    {
        if (args.Count < 1) return null;
        double lower = N(args, 0);
        return args.Count >= 2 ? FR(Erf(N(args, 1)) - Erf(lower)) : FR(Erf(lower));
    }

    // GAMMA(x) — Γ(x); poles at 0 and the negative integers.
    private static FormulaResult? EvalGamma(List<object> args)
    {
        if (args.Count < 1) return null;
        double x = N(args, 0);
        if (x <= 0 && x == Math.Floor(x)) return FormulaResult.Error("#NUM!");
        var g = Gamma(x);
        return double.IsNaN(g) ? FormulaResult.Error("#NUM!") : FR(g);
    }

    // GAMMA.DIST(x, alpha, beta, cumulative) / GAMMADIST.
    private static FormulaResult? EvalGammaDist(List<object> args)
    {
        if (args.Count < 4) return null;
        double x = N(args, 0), alpha = N(args, 1), beta = N(args, 2);
        if (x < 0 || alpha <= 0 || beta <= 0) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 3)) return FR(RegGammaP(alpha, x / beta));
        double pdf = Math.Exp((alpha - 1) * Math.Log(x) - x / beta - alpha * Math.Log(beta) - GammaLn(alpha));
        return FR(pdf);
    }

    // CHISQ.DIST(x, df, cumulative) — chi-squared = gamma(df/2, 2).
    private static FormulaResult? EvalChisqDist(List<object> args)
    {
        if (args.Count < 3) return null;
        double x = N(args, 0), df = N(args, 1);
        if (x < 0 || df < 1) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 2)) return FR(RegGammaP(df / 2, x / 2));
        double pdf = Math.Exp((df / 2 - 1) * Math.Log(x) - x / 2 - (df / 2) * Math.Log(2) - GammaLn(df / 2));
        return FR(pdf);
    }

    // POISSON.DIST(x, mean, cumulative) / POISSON.
    private static FormulaResult? EvalPoisson(List<object> args)
    {
        if (args.Count < 3) return null;
        double k = Math.Floor(N(args, 0)), mean = N(args, 1);
        if (k < 0 || mean < 0) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 2)) return FR(RegGammaQ(k + 1, mean));   // CDF P(X≤k)
        return FR(Math.Exp(-mean + k * Math.Log(mean) - GammaLn(k + 1)));
    }

    // EXPON.DIST(x, lambda, cumulative) / EXPONDIST.
    private static FormulaResult? EvalExpon(List<object> args)
    {
        if (args.Count < 3) return null;
        double x = N(args, 0), lambda = N(args, 1);
        if (x < 0 || lambda <= 0) return FormulaResult.Error("#NUM!");
        return FR(Cumulative(args, 2) ? -Expm1(-lambda * x) : lambda * Math.Exp(-lambda * x));
    }

    // ==================== Incomplete-beta family ====================

    // BETA.DIST(x, alpha, beta, cumulative, [A], [B]) / BETADIST (always CDF).
    private static FormulaResult? EvalBetaDist(List<object> args)
    {
        if (args.Count < 3) return null;
        double x = N(args, 0), a = N(args, 1), b = N(args, 2);
        bool legacy = args.Count < 4 || args[3] is not FormulaResult;     // BETADIST has no cumulative flag
        bool cum = legacy || Cumulative(args, 3);
        double lo = N(args, legacy ? 3 : 4, 0), hi = N(args, legacy ? 4 : 5, 1);
        if (a <= 0 || b <= 0 || hi <= lo || x < lo || x > hi) return FormulaResult.Error("#NUM!");
        double z = (x - lo) / (hi - lo);
        if (cum) return FR(RegIncBeta(z, a, b));
        double pdf = Math.Exp((a - 1) * Math.Log(z) + (b - 1) * Math.Log(1 - z)
            - (GammaLn(a) + GammaLn(b) - GammaLn(a + b))) / (hi - lo);
        return FR(pdf);
    }

    // BETA.INV(p, alpha, beta, [A], [B]) / BETAINV.
    private static FormulaResult? EvalBetaInv(List<object> args)
    {
        if (args.Count < 3) return null;
        double p = N(args, 0), a = N(args, 1), b = N(args, 2), lo = N(args, 3, 0), hi = N(args, 4, 1);
        if (a <= 0 || b <= 0 || p < 0 || p > 1 || hi <= lo) return FormulaResult.Error("#NUM!");
        return FR(lo + (hi - lo) * InvRegIncBeta(p, a, b));
    }

    // Student-t CDF and tails via the incomplete beta.
    private static double TDistCdf(double t, double df)
    {
        double x = df / (df + t * t);
        double tail = 0.5 * RegIncBeta(x, df / 2, 0.5);
        return t >= 0 ? 1 - tail : tail;
    }
    // Right tail computed directly from the incomplete beta — 1 - TDistCdf
    // cancels to 0 once the tail drops below 1e-16.
    private static double TDistRightTail(double t, double df)
    {
        double tail = 0.5 * RegIncBeta(df / (df + t * t), df / 2, 0.5);
        return t >= 0 ? tail : 1 - tail;
    }

    // T.DIST(x, df, cumulative).
    private static FormulaResult? EvalTDist(List<object> args)
    {
        if (args.Count < 3) return null;
        double t = N(args, 0), df = N(args, 1);
        if (df < 1) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 2)) return FR(TDistCdf(t, df));
        double pdf = Math.Exp(GammaLn((df + 1) / 2) - GammaLn(df / 2)) / Math.Sqrt(df * Math.PI)
            * Math.Pow(1 + t * t / df, -(df + 1) / 2);
        return FR(pdf);
    }

    // T.DIST.2T(x, df) — two-tailed (x ≥ 0).
    private static FormulaResult? EvalTDist2T(List<object> args)
    {
        if (args.Count < 2) return null;
        double t = N(args, 0), df = N(args, 1);
        if (t < 0 || df < 1) return FormulaResult.Error("#NUM!");
        return FR(RegIncBeta(df / (df + t * t), df / 2, 0.5));
    }

    // TDIST(x, df, tails) — legacy; x ≥ 0, tails ∈ {1,2}.
    private static FormulaResult? EvalTDistLegacy(List<object> args)
    {
        if (args.Count < 3) return null;
        double t = N(args, 0), df = N(args, 1); int tails = (int)N(args, 2);
        if (t < 0 || df < 1 || (tails != 1 && tails != 2)) return FormulaResult.Error("#NUM!");
        double oneTail = TDistRightTail(t, df);
        return FR(tails == 1 ? oneTail : 2 * oneTail);
    }

    private static double TInv(double p, double df)
    {
        // invert the CDF; CDF is monotonic so bisect over a wide range
        double lo = -1e6, hi = 1e6;
        for (int i = 0; i < 200; i++)
        {
            double mid = 0.5 * (lo + hi);
            if (TDistCdf(mid, df) < p) lo = mid; else hi = mid;
            if (hi - lo < 1e-12) break;
        }
        return 0.5 * (lo + hi);
    }
    private static double TInv2T(double p, double df) => TInv(1 - p / 2, df);   // two-tailed prob → positive t

    // F-distribution CDF via the incomplete beta.
    private static double FDistCdf(double x, double d1, double d2)
    {
        if (x <= 0) return 0;
        return RegIncBeta(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2);
    }

    // Right tail via the beta symmetry I_x(a,b) = 1 - I_{1-x}(b,a) — evaluating
    // the complement directly keeps the far tail instead of cancelling to 0.
    private static double FDistRightTail(double x, double d1, double d2)
    {
        if (x <= 0) return 1;
        return RegIncBeta(d2 / (d1 * x + d2), d2 / 2, d1 / 2);
    }

    // F.DIST(x, df1, df2, cumulative).
    private static FormulaResult? EvalFDist(List<object> args)
    {
        if (args.Count < 4) return null;
        double x = N(args, 0), d1 = N(args, 1), d2 = N(args, 2);
        if (x < 0 || d1 < 1 || d2 < 1) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 3)) return FR(FDistCdf(x, d1, d2));
        double pdf = Math.Sqrt(Math.Pow(d1 * x, d1) * Math.Pow(d2, d2) / Math.Pow(d1 * x + d2, d1 + d2))
            / (x * Math.Exp(GammaLn(d1 / 2) + GammaLn(d2 / 2) - GammaLn((d1 + d2) / 2)));
        return FR(pdf);
    }

    private static double FInv(double p, double d1, double d2)
    {
        double lo = 0, hi = 1e7;
        for (int i = 0; i < 200; i++)
        {
            double mid = 0.5 * (lo + hi);
            if (FDistCdf(mid, d1, d2) < p) lo = mid; else hi = mid;
            if (hi - lo < 1e-10) break;
        }
        return 0.5 * (lo + hi);
    }

    // BINOM.DIST(k, n, p, cumulative) / BINOMDIST.
    private static FormulaResult? EvalBinomDist(List<object> args)
    {
        if (args.Count < 4) return null;
        double k = Math.Floor(N(args, 0)), n = Math.Floor(N(args, 1)), p = N(args, 2);
        if (k < 0 || k > n || p < 0 || p > 1) return FormulaResult.Error("#NUM!");
        if (Cumulative(args, 3))
        {
            double sum = 0;
            for (int i = 0; i <= k; i++) sum += BinomPmf(n, i, p);
            return FR(sum);
        }
        return FR(BinomPmf(n, k, p));
    }

    // BINOM.INV(n, p, alpha) / CRITBINOM — smallest k with CDF ≥ alpha.
    private static FormulaResult? EvalBinomInv(List<object> args)
    {
        if (args.Count < 3) return null;
        double n = Math.Floor(N(args, 0)), p = N(args, 1), alpha = N(args, 2);
        if (n < 0 || p < 0 || p > 1 || alpha <= 0 || alpha > 1) return FormulaResult.Error("#NUM!");
        double cdf = 0;
        for (int k = 0; k <= n; k++)
        {
            cdf += BinomPmf(n, k, p);
            if (cdf >= alpha) return FR(k);
        }
        return FR(n);
    }

    // NEGBINOM.DIST(f, s, p, [cumulative]) / NEGBINOMDIST (pmf only).
    private static FormulaResult? EvalNegBinom(List<object> args)
    {
        if (args.Count < 3) return null;
        double f = Math.Floor(N(args, 0)), s = Math.Floor(N(args, 1)), p = N(args, 2);
        if (f < 0 || s < 1 || p < 0 || p > 1) return FormulaResult.Error("#NUM!");
        // p^s · (1−p)^i with a C(i+s−1,i) that can overflow alone — log space
        // for large counts, direct product below (digit-exact) otherwise.
        double Pmf(double i) =>
            p >= 1 ? (i == 0 ? 1 : 0) :
            p <= 0 ? 0 :
            i + s - 1 <= 170 ? Binom(i + s - 1, i) * Math.Pow(p, s) * Math.Pow(1 - p, i)
                             : Math.Exp(BinomLn(i + s - 1, i) + s * Math.Log(p) + i * Math.Log(1 - p));
        bool cum = args.Count >= 4 && Cumulative(args, 3);
        if (cum)
        {
            double sum = 0;
            for (int i = 0; i <= f; i++) sum += Pmf(i);
            return FR(sum);
        }
        return FR(Pmf(f));
    }

    // WEIBULL.DIST(x, alpha, beta, cumulative) / WEIBULL.
    private static FormulaResult? EvalWeibull(List<object> args)
    {
        if (args.Count < 4) return null;
        double x = N(args, 0), a = N(args, 1), b = N(args, 2);
        if (x < 0 || a <= 0 || b <= 0) return FormulaResult.Error("#NUM!");
        double z = Math.Pow(x / b, a);
        return FR(Cumulative(args, 3) ? -Expm1(-z) : a / Math.Pow(b, a) * Math.Pow(x, a - 1) * Math.Exp(-z));
    }

    // LOGNORM.DIST(x, mean, sd, cumulative) / LOGNORMDIST (always CDF).
    private static FormulaResult? EvalLognormDist(List<object> args)
    {
        if (args.Count < 3) return null;
        double x = N(args, 0), mean = N(args, 1), sd = N(args, 2);
        if (x <= 0 || sd <= 0) return FormulaResult.Error("#NUM!");
        double z = (Math.Log(x) - mean) / sd;
        bool cum = args.Count < 4 || args[3] is not FormulaResult || Cumulative(args, 3);
        return FR(cum ? NormCdf(z) : NormPdf(z) / (x * sd));
    }

    // HYPGEOM.DIST(sample_s, sample_n, pop_s, pop_n, [cumulative]) / HYPGEOMDIST.
    private static FormulaResult? EvalHypgeom(List<object> args)
    {
        if (args.Count < 4) return null;
        double k = Math.Floor(N(args, 0)), n = Math.Floor(N(args, 1)), K = Math.Floor(N(args, 2)), Npop = Math.Floor(N(args, 3));
        if (k < 0 || k > n || n > Npop || K > Npop || k > K || n - k > Npop - K) return FormulaResult.Error("#NUM!");
        // Ratio of three coefficients — any of which can overflow alone; keep
        // the small-population direct ratio (digit-exact), log space beyond.
        double Pmf(double i) =>
            Npop <= 170 ? Binom(K, i) * Binom(Npop - K, n - i) / Binom(Npop, n)
                        : Math.Exp(BinomLn(K, i) + BinomLn(Npop - K, n - i) - BinomLn(Npop, n));
        bool cum = args.Count >= 5 && Cumulative(args, 4);
        if (cum) { double sum = 0; for (int i = 0; i <= k; i++) sum += Pmf(i); return FR(sum); }
        return FR(Pmf(k));
    }
}
