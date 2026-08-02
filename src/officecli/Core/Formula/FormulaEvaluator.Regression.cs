// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

// W4d — the array-returning regression family (LINEST / LOGEST / TREND / GROWTH).
// Each spills: LINEST/LOGEST return the coefficient row (Excel order
// {m_k,…,m_1,b}) optionally with the 5-row stats block; TREND/GROWTH return the
// fitted values. The dynamic-array writeback (ExcelHandler.DynamicArray) makes
// Excel spill them; the anchor (top-left) is the cell's computedValue. Verified
// against real Excel.
internal partial class FormulaEvaluator
{
    // Ordinary least squares: solve (XᵀX)β = Xᵀy. Returns β (length p) or null
    // when the normal-equation system is singular.
    private static double[]? LeastSquares(double[,] x, double[] y, out double[,]? xtxInv)
    {
        xtxInv = null;
        int n = x.GetLength(0), p = x.GetLength(1);
        var a = new double[p, p];
        var g = new double[p];
        for (int i = 0; i < p; i++)
        {
            for (int j = 0; j < p; j++)
            {
                double s = 0;
                for (int k = 0; k < n; k++) s += x[k, i] * x[k, j];
                a[i, j] = s;
            }
            double gi = 0;
            for (int k = 0; k < n; k++) gi += x[k, i] * y[k];
            g[i] = gi;
        }
        var inv = Invert(a);
        if (inv == null) return null;
        xtxInv = inv;
        var beta = new double[p];
        for (int i = 0; i < p; i++)
        {
            double s = 0;
            for (int j = 0; j < p; j++) s += inv[i, j] * g[j];
            beta[i] = s;
        }
        return beta;
    }

    // Gauss-Jordan inverse with partial pivoting; null when singular.
    private static double[,]? Invert(double[,] m)
    {
        int n = m.GetLength(0);
        var a = new double[n, 2 * n];
        for (int i = 0; i < n; i++) { for (int j = 0; j < n; j++) a[i, j] = m[i, j]; a[i, n + i] = 1; }
        for (int col = 0; col < n; col++)
        {
            int piv = col;
            for (int r = col + 1; r < n; r++) if (Math.Abs(a[r, col]) > Math.Abs(a[piv, col])) piv = r;
            if (Math.Abs(a[piv, col]) < 1e-300) return null;
            if (piv != col) for (int j = 0; j < 2 * n; j++) (a[col, j], a[piv, j]) = (a[piv, j], a[col, j]);
            double d = a[col, col];
            for (int j = 0; j < 2 * n; j++) a[col, j] /= d;
            for (int r = 0; r < n; r++)
            {
                if (r == col) continue;
                double f = a[r, col];
                for (int j = 0; j < 2 * n; j++) a[r, j] -= f * a[col, j];
            }
        }
        var inv = new double[n, n];
        for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) inv[i, j] = a[i, n + j];
        return inv;
    }

    // A real range/array argument (known_x, new_x): arrives as a RangeData (from
    // the function-arg range interception) or an Area/Array FormulaResult — NOT a
    // plain FormulaResult. Bare scalars (e.g. an omitted-slot 0) are not grids.
    private static bool HasGridArg(List<object> args, int i)
        => i < args.Count && (args[i] is RangeData || args[i] is FormulaResult { IsRange: true } or FormulaResult { IsArray: true });

    // Pull (y vector, X matrix n×k of independent variables) from known_y /
    // optional known_x, normalizing orientation to n rows.
    private bool RegressionInputs(List<object> args, out double[] y, out double[][] xcols)
    {
        y = Array.Empty<double>(); xcols = Array.Empty<double[]>();
        if (ToGrid(args.Count > 0 ? args[0] : null) is not { } yg) return false;
        y = Flatten(yg).Select(c => c?.AsNumber() ?? 0).ToArray();
        int n = y.Length;
        if (n == 0) return false;

        if (HasGridArg(args, 1) && ToGrid(args[1]) is { } xg)
        {
            int xr = xg.GetLength(0), xc = xg.GetLength(1);
            // Variables run along the dimension that is NOT length n. When the
            // block is n×k, each column is a variable; when k×n, each row is.
            if (xr == n)
            {
                xcols = new double[xc][];
                for (int c = 0; c < xc; c++)
                {
                    xcols[c] = new double[n];
                    for (int r = 0; r < n; r++) xcols[c][r] = xg[r, c]?.AsNumber() ?? 0;
                }
            }
            else if (xc == n)
            {
                xcols = new double[xr][];
                for (int v = 0; v < xr; v++)
                {
                    xcols[v] = new double[n];
                    for (int r = 0; r < n; r++) xcols[v][r] = xg[v, r]?.AsNumber() ?? 0;
                }
            }
            else return false;
        }
        else
        {
            // Omitted known_x → the sequence {1,2,…,n}.
            var seq = new double[n];
            for (int i = 0; i < n; i++) seq[i] = i + 1;
            xcols = new[] { seq };
        }
        return true;
    }

    // Build the design matrix (independent columns + optional trailing intercept
    // column of ones) and fit. coeffs come back in NATURAL order [m1,…,mk,(b)].
    private double[]? FitLinear(double[] y, double[][] xcols, bool withConst, out double[,]? xtxInv, out int k)
    {
        k = xcols.Length;
        int n = y.Length;
        int p = k + (withConst ? 1 : 0);
        var x = new double[n, p];
        for (int r = 0; r < n; r++)
        {
            for (int c = 0; c < k; c++) x[r, c] = xcols[c][r];
            if (withConst) x[r, k] = 1.0;
        }
        return LeastSquares(x, y, out xtxInv);
    }

    private FormulaResult? EvalLinest(List<object> args, bool log)
    {
        if (!RegressionInputs(args, out var y, out var xcols)) return FormulaResult.Error("#VALUE!");
        bool withConst = args.Count <= 2 || args[2] is not FormulaResult c2 || c2.IsBlank || c2.AsNumber() != 0;
        bool stats = args.Count > 3 && args[3] is FormulaResult s3 && !s3.IsBlank && s3.AsNumber() != 0;

        double[] fitY = y;
        if (log)
        {
            if (y.Any(v => v <= 0)) return FormulaResult.Error("#NUM!");
            fitY = y.Select(v => Math.Log(v)).ToArray();
        }
        var beta = FitLinear(fitY, xcols, withConst, out var xtxInv, out int k);
        if (beta is null) return FormulaResult.Error("#NUM!");

        // Excel order: {m_k, …, m_1, b}. Natural order is [m1..mk,(b)].
        var slopes = new double[k];
        for (int i = 0; i < k; i++) slopes[i] = beta[i];
        double b0 = withConst ? beta[k] : 0.0;
        var coeff = new double[k + 1];
        for (int i = 0; i < k; i++) coeff[i] = slopes[k - 1 - i];   // reversed slopes
        coeff[k] = b0;
        var logCoeff = log ? coeff.Select((v, i) => i < k ? Math.Exp(v) : Math.Exp(v)).ToArray() : coeff;

        if (!stats)
        {
            var row = new FormulaResult?[1, k + 1];
            for (int i = 0; i <= k; i++) row[0, i] = FormulaResult.Number(logCoeff[i]);
            return MakeArea(row);
        }

        // Full 5-row stats block (coefficients, SEs, [r2, sey], [F, df], [ssreg, ssresid]).
        int n = fitY.Length, p = k + (withConst ? 1 : 0);
        var pred = new double[n];
        for (int r = 0; r < n; r++)
        {
            double v = withConst ? beta[k] : 0;
            for (int c = 0; c < k; c++) v += beta[c] * xcols[c][r];
            pred[r] = v;
        }
        double meanY = fitY.Average();
        double ssTot = withConst ? fitY.Sum(v => (v - meanY) * (v - meanY)) : fitY.Sum(v => v * v);
        double ssResid = 0; for (int r = 0; r < n; r++) ssResid += (fitY[r] - pred[r]) * (fitY[r] - pred[r]);
        double ssReg = ssTot - ssResid;
        int dfResid = n - p;
        double r2 = ssTot == 0 ? 1 : 1 - ssResid / ssTot;
        double sey = dfResid > 0 ? Math.Sqrt(ssResid / dfResid) : 0;
        double fStat = (dfResid > 0 && ssResid > 0) ? (ssReg / k) / (ssResid / dfResid) : double.PositiveInfinity;

        // Coefficient standard errors: sey * sqrt(diag((XᵀX)⁻¹)), same reversal.
        var se = new double[k + 1];
        if (xtxInv != null)
        {
            var seNat = new double[p];
            for (int i = 0; i < p; i++) seNat[i] = sey * Math.Sqrt(Math.Max(0, xtxInv[i, i]));
            for (int i = 0; i < k; i++) se[i] = seNat[k - 1 - i];
            se[k] = withConst ? seNat[k] : double.NaN;
        }

        var na = FormulaResult.Error("#N/A");
        var grid = new FormulaResult?[5, k + 1];
        for (int i = 0; i <= k; i++) grid[0, i] = FormulaResult.Number(logCoeff[i]);
        for (int i = 0; i <= k; i++) grid[1, i] = withConst || i < k ? FormulaResult.Number(se[i]) : na;
        grid[2, 0] = FormulaResult.Number(r2); grid[2, 1] = FormulaResult.Number(log ? Math.Exp(sey) : sey);
        for (int i = 2; i <= k; i++) grid[2, i] = na;
        grid[3, 0] = FormulaResult.Number(fStat); grid[3, 1] = FormulaResult.Number(dfResid);
        for (int i = 2; i <= k; i++) grid[3, i] = na;
        grid[4, 0] = FormulaResult.Number(ssReg); grid[4, 1] = FormulaResult.Number(ssResid);
        for (int i = 2; i <= k; i++) grid[4, i] = na;
        return MakeArea(grid);
    }

    // TREND / GROWTH(known_y, [known_x], [new_x], [const]) — fitted predictions.
    private FormulaResult? EvalTrend(List<object> args, bool log)
    {
        if (!RegressionInputs(args, out var y, out var xcols)) return FormulaResult.Error("#VALUE!");
        bool withConst = args.Count <= 3 || args[3] is not FormulaResult c3 || c3.IsBlank || c3.AsNumber() != 0;
        int k = xcols.Length, n = y.Length;

        double[] fitY = y;
        if (log)
        {
            if (y.Any(v => v <= 0)) return FormulaResult.Error("#NUM!");
            fitY = y.Select(v => Math.Log(v)).ToArray();
        }
        var beta = FitLinear(fitY, xcols, withConst, out _, out _);
        if (beta is null) return FormulaResult.Error("#NUM!");

        // new_x: a range/array OR a scalar; absent → predict over known_x.
        bool hasNewX = args.Count > 2 && (args[2] is RangeData || args[2] is FormulaResult { IsBlank: false });
        FormulaResult?[,]? newGrid = hasNewX && ToGrid(args[2]) is { } ng ? ng : null;
        double[][] predCols; int m; FormulaResult?[,] outShape;
        if (newGrid == null)
        {
            predCols = xcols; m = n;
            outShape = new FormulaResult?[n, 1];
        }
        else
        {
            int nr = newGrid.GetLength(0), nc = newGrid.GetLength(1);
            if (k == 1)
            {
                m = nr * nc;
                predCols = new[] { new double[m] };
                int idx = 0;
                for (int r = 0; r < nr; r++) for (int c = 0; c < nc; c++) predCols[0][idx++] = newGrid[r, c]?.AsNumber() ?? 0;
                outShape = new FormulaResult?[nr, nc];
            }
            else if (nr == k) { m = nc; predCols = ToCols(newGrid, k, m, byRow: true); outShape = new FormulaResult?[1, m]; }
            else if (nc == k) { m = nr; predCols = ToCols(newGrid, k, m, byRow: false); outShape = new FormulaResult?[m, 1]; }
            else return FormulaResult.Error("#REF!");
        }

        var preds = new double[m];
        for (int i = 0; i < m; i++)
        {
            double v = withConst ? beta[k] : 0;
            for (int c = 0; c < k; c++) v += beta[c] * predCols[c][i];
            preds[i] = log ? Math.Exp(v) : v;
        }
        // Fill outShape row-major with preds.
        int orows = outShape.GetLength(0), ocols = outShape.GetLength(1), pi = 0;
        for (int r = 0; r < orows; r++) for (int c = 0; c < ocols; c++) outShape[r, c] = FormulaResult.Number(preds[pi++]);
        return MakeArea(outShape);
    }

    private static double[][] ToCols(FormulaResult?[,] g, int k, int m, bool byRow)
    {
        var cols = new double[k][];
        for (int v = 0; v < k; v++)
        {
            cols[v] = new double[m];
            for (int i = 0; i < m; i++) cols[v][i] = (byRow ? g[v, i] : g[i, v])?.AsNumber() ?? 0;
        }
        return cols;
    }
}
