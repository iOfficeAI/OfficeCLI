// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

// W6b — the lambda-driven spill functions (MAP / BYROW / BYCOL / SCAN /
// MAKEARRAY). They build on the W5 LAMBDA machinery: the trailing argument
// arrives as an IsLambda FormulaResult and is applied per element/row/col/cell
// via InvokeLambda. Like the rest of W6 they return a FormulaResult.Area whose
// top-left becomes the anchor's computedValue while Excel spills the region.
internal partial class FormulaEvaluator
{
    private static Lambda? AsLambda(object? a)
        => a is FormulaResult { IsLambda: true } fr ? (Lambda)fr.LambdaValue! : null;

    // A single grid row / column as a 1×N / N×1 Area (so a lambda body like
    // SUM(r) sees a real range).
    private static FormulaResult RowArea(FormulaResult?[,] g, int r)
    {
        int cols = g.GetLength(1);
        var row = new FormulaResult?[1, cols];
        for (int c = 0; c < cols; c++) row[0, c] = g[r, c];
        return MakeArea(row);
    }

    private static FormulaResult ColArea(FormulaResult?[,] g, int c)
    {
        int rows = g.GetLength(0);
        var col = new FormulaResult?[rows, 1];
        for (int r = 0; r < rows; r++) col[r, 0] = g[r, c];
        return MakeArea(col);
    }

    // MAP(array1, [array2, …], lambda) — apply lambda elementwise across one or
    // more equally-shaped arrays.
    private FormulaResult? EvalMap(List<object> args)
    {
        if (args.Count < 2 || AsLambda(args[^1]) is not { } lam) return FormulaResult.Error("#VALUE!");
        var grids = new List<FormulaResult?[,]>();
        for (int i = 0; i < args.Count - 1; i++)
        {
            if (ToGrid(args[i]) is not { } g) return FormulaResult.Error("#VALUE!");
            grids.Add(g);
        }
        int rows = grids[0].GetLength(0), cols = grids[0].GetLength(1);
        if (grids.Any(g => g.GetLength(0) != rows || g.GetLength(1) != cols))
            return FormulaResult.Error("#VALUE!");
        var outc = new FormulaResult?[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                outc[r, c] = InvokeLambda(lam, grids.Select(g => Cell0(g[r, c])).ToList());
        return MakeArea(outc);
    }

    // BYROW(array, lambda) — lambda(row) → scalar; result is a column vector.
    private FormulaResult? EvalByRow(List<object> args)
    {
        if (ToGrid(args.Count > 0 ? args[0] : null) is not { } g || AsLambda(args.Count > 1 ? args[1] : null) is not { } lam)
            return FormulaResult.Error("#VALUE!");
        int rows = g.GetLength(0);
        var outc = new FormulaResult?[rows, 1];
        for (int r = 0; r < rows; r++)
        {
            var v = InvokeLambda(lam, new List<FormulaResult> { RowArea(g, r) });
            if (v is { IsRange: true } or { IsArray: true }) return FormulaResult.Error("#CALC!");   // lambda must yield a scalar
            outc[r, 0] = v;
        }
        return MakeArea(outc);
    }

    // BYCOL(array, lambda) — lambda(col) → scalar; result is a row vector.
    private FormulaResult? EvalByCol(List<object> args)
    {
        if (ToGrid(args.Count > 0 ? args[0] : null) is not { } g || AsLambda(args.Count > 1 ? args[1] : null) is not { } lam)
            return FormulaResult.Error("#VALUE!");
        int cols = g.GetLength(1);
        var outc = new FormulaResult?[1, cols];
        for (int c = 0; c < cols; c++)
        {
            var v = InvokeLambda(lam, new List<FormulaResult> { ColArea(g, c) });
            if (v is { IsRange: true } or { IsArray: true }) return FormulaResult.Error("#CALC!");   // lambda must yield a scalar
            outc[0, c] = v;
        }
        return MakeArea(outc);
    }

    // SCAN(init, array, lambda) — running fold; result mirrors array's shape with
    // each cell holding the accumulation up to and including that element.
    private FormulaResult? EvalScan(List<object> args)
    {
        if (args.Count < 3) return FormulaResult.Error("#VALUE!");
        var acc = args[0] as FormulaResult ?? FormulaResult.Number(0);
        if (ToGrid(args[1]) is not { } g || AsLambda(args[2]) is not { } lam) return FormulaResult.Error("#VALUE!");
        int rows = g.GetLength(0), cols = g.GetLength(1);
        var outc = new FormulaResult?[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                acc = InvokeLambda(lam, new List<FormulaResult> { acc, Cell0(g[r, c]) });
                outc[r, c] = acc;
            }
        return MakeArea(outc);
    }

    // MAKEARRAY(rows, cols, lambda) — lambda(r, c) over 1-based indices.
    private FormulaResult? EvalMakeArray(List<object> args)
    {
        if (args.Count < 3 || AsLambda(args[2]) is not { } lam) return FormulaResult.Error("#VALUE!");
        int rows = (int)Scalar(args, 0, 0), cols = (int)Scalar(args, 1, 0);
        if (rows < 1 || cols < 1) return FormulaResult.Error("#VALUE!");
        if ((long)rows * cols > 1_048_576) return FormulaResult.Error("#NUM!");
        var outc = new FormulaResult?[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                outc[r, c] = InvokeLambda(lam,
                    new List<FormulaResult> { FormulaResult.Number(r + 1), FormulaResult.Number(c + 1) });
        return MakeArea(outc);
    }
}
