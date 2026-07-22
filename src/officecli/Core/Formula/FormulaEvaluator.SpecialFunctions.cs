// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Special functions ====================
    //
    // Shared numerical core for the statistical-distribution family: the log-
    // gamma function, the regularized incomplete gamma integrals and their
    // inverse, the error function, and the standard-normal CDF / inverse. The
    // distribution wrappers (NORM.DIST, GAMMA.DIST, CHISQ.DIST, POISSON.DIST, …)
    // are thin calls onto these. Algorithms are the standard Lanczos
    // approximation and the series / continued-fraction expansions for the
    // incomplete gamma integral (convergence to ~1e-12), which is what keeps the
    // predicted value within Excel's cache-staleness tolerance.

    private const double Sqrt2 = 1.4142135623730951;
    private const double Sqrt2Pi = 2.5066282746310002;

    // Lanczos log-gamma (g=7, n=9). Accurate to ~1e-15 for x > 0.
    private static readonly double[] LanczosG7 =
    {
        0.99999999999980993, 676.5203681218851, -1259.1392167224028,
        771.32342877765313, -176.61502916214059, 12.507343278686905,
        -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
    };

    internal static double GammaLn(double x)
    {
        if (x <= 0) return double.NaN;
        if (x < 0.5)
            // reflection: Γ(x)Γ(1-x) = π / sin(πx)
            return Math.Log(Math.PI / Math.Sin(Math.PI * x)) - GammaLn(1 - x);
        x -= 1;
        double a = LanczosG7[0];
        double t = x + 7.5;
        for (int i = 1; i < LanczosG7.Length; i++) a += LanczosG7[i] / (x + i);
        return 0.5 * Math.Log(2 * Math.PI) + (x + 0.5) * Math.Log(t) - t + Math.Log(a);
    }

    internal static double Gamma(double x)
    {
        if (x <= 0 && x == Math.Floor(x)) return double.NaN;   // poles at 0, -1, -2…
        if (x < 0.5) return Math.PI / (Math.Sin(Math.PI * x) * Gamma(1 - x));
        return Math.Exp(GammaLn(x));
    }

    // Lower regularized incomplete gamma P(a,x) = γ(a,x)/Γ(a). Q = 1 - P.
    internal static double RegGammaP(double a, double x)
    {
        if (x < 0 || a <= 0) return double.NaN;
        if (x == 0) return 0;
        if (x < a + 1) return GammaSeries(a, x);          // series converges fast here
        return 1.0 - GammaContinuedFraction(a, x);        // CF for the upper tail
    }

    internal static double RegGammaQ(double a, double x)
    {
        if (x < 0 || a <= 0) return double.NaN;
        if (x == 0) return 1;
        if (x < a + 1) return 1.0 - GammaSeries(a, x);
        return GammaContinuedFraction(a, x);
    }

    private static double GammaSeries(double a, double x)
    {
        double ap = a, sum = 1.0 / a, del = sum;
        for (int n = 0; n < 200; n++)
        {
            ap += 1; del *= x / ap; sum += del;
            if (Math.Abs(del) < Math.Abs(sum) * 1e-15) break;
        }
        return sum * Math.Exp(-x + a * Math.Log(x) - GammaLn(a));
    }

    private static double GammaContinuedFraction(double a, double x)
    {
        const double tiny = 1e-300;
        double b = x + 1 - a, c = 1 / tiny, d = 1 / b, h = d;
        for (int i = 1; i < 200; i++)
        {
            double an = -i * (i - a);
            b += 2;
            d = an * d + b; if (Math.Abs(d) < tiny) d = tiny;
            c = b + an / c; if (Math.Abs(c) < tiny) c = tiny;
            d = 1 / d;
            double del = d * c;
            h *= del;
            if (Math.Abs(del - 1) < 1e-15) break;
        }
        return Math.Exp(-x + a * Math.Log(x) - GammaLn(a)) * h;
    }

    // Inverse of P(a,x)=p in x. Newton on P with the gamma PDF as derivative,
    // seeded by a Wilson–Hilferty / tail estimate.
    internal static double InvRegGammaP(double a, double p)
    {
        if (p <= 0) return 0;
        if (p >= 1) return double.PositiveInfinity;
        double gln = GammaLn(a);
        double x;
        // initial guess
        if (a > 1)
        {
            double pp = p < 0.5 ? p : 1 - p;
            double tt = Math.Sqrt(-2 * Math.Log(pp));
            double xx = (2.30753 + tt * 0.27061) / (1 + tt * (0.99229 + tt * 0.04481)) - tt;
            if (p < 0.5) xx = -xx;
            x = Math.Max(1e-3, a * Math.Pow(1 - 1.0 / (9 * a) + xx / (3 * Math.Sqrt(a)), 3));
        }
        else
        {
            double t = 1 - a * (0.253 + a * 0.12);
            x = p < t ? Math.Pow(p / t, 1 / a) : 1 - Math.Log(1 - (p - t) / (1 - t));
        }
        for (int i = 0; i < 100; i++)
        {
            double err = RegGammaP(a, x) - p;
            double pdf = Math.Exp(-x + (a - 1) * Math.Log(x) - gln);   // d/dx P(a,x)
            double dx = err / (pdf > 1e-300 ? pdf : 1e-300);
            // damp so we never step below zero
            x -= (Math.Abs(dx) < 0.5 * x ? dx : 0.5 * x * Math.Sign(dx));
            if (x <= 0) x = 1e-12;
            if (Math.Abs(dx) < 1e-12 * x) break;
        }
        return x;
    }

    // erf / erfc via the incomplete gamma relation (high precision):
    // erf(x) = P(1/2, x²) for x ≥ 0.
    internal static double Erf(double x) => x < 0 ? -RegGammaP(0.5, x * x) : RegGammaP(0.5, x * x);
    // Q(1/2,x²) keeps the far tail (erfc(7)≈4e-23) instead of the 1-Erf form,
    // which cancels to exactly 0 once Erf rounds to 1 at x ≳ 6.
    internal static double Erfc(double x) => x < 0 ? 2.0 - RegGammaQ(0.5, x * x) : RegGammaQ(0.5, x * x);

    // e^x − 1 without cancellation near 0 (BCL has no Math.Expm1); the cubic
    // series term keeps full double precision over the small-|x| branch.
    internal static double Expm1(double x) =>
        Math.Abs(x) < 1e-4 ? x * (1 + x * (0.5 + x / 6)) : Math.Exp(x) - 1;

    internal static double NormPdf(double z) => Math.Exp(-0.5 * z * z) / Sqrt2Pi;
    internal static double NormCdf(double z) => 0.5 * Erfc(-z / Sqrt2);

    // Inverse standard-normal CDF (Acklam's rational approximation, then one
    // Halley refinement for full double precision).
    internal static double InvNormCdf(double p)
    {
        if (p <= 0) return double.NegativeInfinity;
        if (p >= 1) return double.PositiveInfinity;
        double[] a = { -3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00 };
        double[] b = { -5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01 };
        double[] c = { -7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00 };
        double[] d = { 7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00 };
        const double plow = 0.02425, phigh = 1 - 0.02425;
        double q, r, z;
        if (p < plow)
        {
            q = Math.Sqrt(-2 * Math.Log(p));
            z = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
                ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
        }
        else if (p <= phigh)
        {
            q = p - 0.5; r = q * q;
            z = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q /
                (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
        }
        else
        {
            q = Math.Sqrt(-2 * Math.Log(1 - p));
            z = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
                ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
        }
        // Halley step
        double e = NormCdf(z) - p;
        double u = e * Sqrt2Pi * Math.Exp(0.5 * z * z);
        z -= u / (1 + 0.5 * z * u);
        return z;
    }

    // Regularized incomplete beta I_x(a,b) and its inverse — base for the
    // T / F / BETA / BINOM distributions (added in a later wave).
    internal static double RegIncBeta(double x, double a, double b)
    {
        if (x <= 0) return 0;
        if (x >= 1) return 1;
        double lbeta = GammaLn(a) + GammaLn(b) - GammaLn(a + b);
        double front = Math.Exp(Math.Log(x) * a + Math.Log(1 - x) * b - lbeta) / a;
        // continued fraction (Lentz), use symmetry for fast convergence
        if (x < (a + 1) / (a + b + 2))
            return front * BetaCF(x, a, b);
        return 1 - Math.Exp(Math.Log(1 - x) * b + Math.Log(x) * a - lbeta) / b * BetaCF(1 - x, b, a);
    }

    // Inverse of I_x(a,b)=p in x, by bisection (monotonic in x on [0,1]).
    internal static double InvRegIncBeta(double p, double a, double b)
    {
        if (p <= 0) return 0;
        if (p >= 1) return 1;
        double lo = 0, hi = 1;
        for (int i = 0; i < 200; i++)
        {
            double mid = 0.5 * (lo + hi);
            if (RegIncBeta(mid, a, b) < p) lo = mid; else hi = mid;
            if (hi - lo < 1e-15) break;
        }
        return 0.5 * (lo + hi);
    }

    // Binomial coefficient C(n,k) via log-gamma (stable for large n).
    internal static double Binom(double n, double k)
    {
        if (k < 0 || k > n) return 0;
        return Math.Round(Math.Exp(BinomLn(n, k)));
    }

    // ln C(n,k) — kept in log space so pmf terms can cancel a huge coefficient
    // against underflowed p-powers before exponentiating (C(10000,5000) alone
    // overflows double, yet BINOM.DIST(5000,10000,0.5,FALSE) ≈ 0.008).
    internal static double BinomLn(double n, double k) =>
        k < 0 || k > n ? double.NegativeInfinity
                       : GammaLn(n + 1) - GammaLn(k + 1) - GammaLn(n - k + 1);

    // C(n,k)·p^k·(1−p)^(n−k). Small n keeps the direct product (digit-exact);
    // large n assembles the term in log space.
    internal static double BinomPmf(double n, double k, double p)
    {
        if (k < 0 || k > n) return 0;
        if (p <= 0) return k == 0 ? 1 : 0;
        if (p >= 1) return k == n ? 1 : 0;
        if (n <= 170) return Binom(n, k) * Math.Pow(p, k) * Math.Pow(1 - p, n - k);
        return Math.Exp(BinomLn(n, k) + k * Math.Log(p) + (n - k) * Math.Log(1 - p));
    }

    private static double BetaCF(double x, double a, double b)
    {
        const double tiny = 1e-300;
        double qab = a + b, qap = a + 1, qam = a - 1;
        double c = 1, d = 1 - qab * x / qap;
        if (Math.Abs(d) < tiny) d = tiny;
        d = 1 / d; double h = d;
        for (int m = 1; m < 300; m++)
        {
            int m2 = 2 * m;
            double aa = m * (b - m) * x / ((qam + m2) * (a + m2));
            d = 1 + aa * d; if (Math.Abs(d) < tiny) d = tiny;
            c = 1 + aa / c; if (Math.Abs(c) < tiny) c = tiny;
            d = 1 / d; h *= d * c;
            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
            d = 1 + aa * d; if (Math.Abs(d) < tiny) d = tiny;
            c = 1 + aa / c; if (Math.Abs(c) < tiny) c = tiny;
            d = 1 / d; double del = d * c; h *= del;
            if (Math.Abs(del - 1) < 1e-15) break;
        }
        return h;
    }
}
