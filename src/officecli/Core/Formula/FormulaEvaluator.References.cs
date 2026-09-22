// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Unresolved cell or area reference, kept first-class so that OFFSET / INDIRECT
/// can manipulate the reference itself instead of receiving a dereferenced value.
/// Single-cell refs use Width=Height=1.
/// </summary>
internal record RefArg(string? Sheet, int Col, int Row, int Width, int Height);

internal partial class FormulaEvaluator
{
    /// <summary>
    /// Convert a token-level range expression like "A1:B3" (or "A:A", "1:1")
    /// to a RefArg. Sheet-prefixed forms pass the sheet name in via parameter.
    /// </summary>
    private RefArg? BuildRefFromRange(string? sheet, string rangeExpr)
    {
        var parts = rangeExpr.Split(':');
        if (parts.Length != 2) return null;
        var left = StripDollar(parts[0]);
        var right = StripDollar(parts[1]);

        var leftColOnly = Regex.IsMatch(left, @"^[A-Z]+$", RegexOptions.IgnoreCase);
        var rightColOnly = Regex.IsMatch(right, @"^[A-Z]+$", RegexOptions.IgnoreCase);
        var leftRowOnly = Regex.IsMatch(left, @"^\d+$");
        var rightRowOnly = Regex.IsMatch(right, @"^\d+$");

        int r1, r2, c1, c2;
        if (leftColOnly && rightColOnly)
        {
            c1 = ColToIndex(left.ToUpperInvariant());
            c2 = ColToIndex(right.ToUpperInvariant());
            var target = GetSheetDataFor(sheet);
            if (target == null) return null;
            var (minRow, maxRow) = GetPopulatedRowRange(target);
            if (maxRow == 0) return null;
            r1 = minRow; r2 = maxRow;
        }
        else if (leftRowOnly && rightRowOnly)
        {
            r1 = int.Parse(left); r2 = int.Parse(right);
            var target = GetSheetDataFor(sheet);
            if (target == null) return null;
            var (minCol, maxCol) = GetPopulatedColRange(target);
            if (maxCol == 0) return null;
            c1 = minCol; c2 = maxCol;
        }
        else
        {
            var (col1, row1) = ParseRef(left);
            var (col2, row2) = ParseRef(right);
            c1 = ColToIndex(col1); c2 = ColToIndex(col2);
            r1 = row1; r2 = row2;
        }

        var colMin = Math.Min(c1, c2); var colMax = Math.Max(c1, c2);
        var rowMin = Math.Min(r1, r2); var rowMax = Math.Max(r1, r2);
        // Excel sheet limits: rows 1..1048576, cols 1..16384 (XFD).
        if (rowMin < 1 || rowMax > ExcelMaxRow) return null;
        if (colMin < 1 || colMax > ExcelMaxCol) return null;
        return new RefArg(sheet, colMin, rowMin, colMax - colMin + 1, rowMax - rowMin + 1);
    }

    /// <summary>
    /// Parse a reference string (e.g. "A1", "Sheet1!B2", "A1:C3") into a RefArg.
    /// Used by INDIRECT to convert its evaluated string argument into a reference.
    /// </summary>
    private RefArg? ParseRefString(string s)
    {
        // R3 BUG A: only trim ASCII space + tab. .Trim() (no args) strips ALL
        // Unicode whitespace including NBSP (U+00A0) — Excel does NOT trim NBSP
        // from INDIRECT's argument; an NBSP-padded ref must yield #REF!. We
        // keep ASCII-space lenience because Round 1 chose that as a deliberate
        // ergonomic deviation (`INDIRECT(" A1 ")` already worked and tests
        // depend on it); NBSP and other Unicode whitespace fall through to
        // IsCellRef, fail to match, and surface as #REF! naturally.
        s = s.Trim(' ', '\t');
        if (string.IsNullOrEmpty(s)) return null;
        string? sheet = null;
        var bang = s.IndexOf('!');
        if (bang > 0)
        {
            sheet = s[..bang].Trim('\'');
            s = s[(bang + 1)..];
        }
        s = StripDollar(s);
        if (s.Contains(':')) return BuildRefFromRange(sheet, s);
        if (IsCellRef(s))
        {
            var (col, row) = ParseRef(s);
            var colIdx = ColToIndex(col);
            // Excel sheet limits: row 1..1048576, col 1..16384 (XFD).
            if (row < 1 || row > ExcelMaxRow) return null;
            if (colIdx < 1 || colIdx > ExcelMaxCol) return null;
            return new RefArg(sheet, colIdx, row, 1, 1);
        }
        return null;
    }

    /// <summary>
    /// Resolve a RefArg to the actual cell values. Single-cell → scalar
    /// FormulaResult; multi-cell → FormulaResult.Area wrapping a RangeData.
    /// </summary>
    private FormulaResult? ResolveRef(RefArg r)
    {
        var rangeMemoKey = r.Height * r.Width > 1
            ? $"{r.Sheet ?? _sheetKey}|{r.Col},{r.Row},{r.Width},{r.Height}"
            : null;
        if (rangeMemoKey != null && _session.RangeMemo.TryGetValue(rangeMemoKey, out var memoRange))
            return FormulaResult.Area(memoRange);
        var circularBefore = _session.CircularHits;
        var cells = new FormulaResult?[r.Height, r.Width];
        for (int dr = 0; dr < r.Height; dr++)
            for (int dc = 0; dc < r.Width; dc++)
            {
                var cellRef = $"{IndexToCol(r.Col + dc)}{r.Row + dr}";
                cells[dr, dc] = r.Sheet != null
                    ? ResolveSheetCellResult($"{r.Sheet}!{cellRef}")
                    : ResolveCellResult(cellRef);
            }
        // R16-1: for a single-cell ref whose resolved value is an error (e.g. a
        // depth-guard #NUM! from a deep INDIRECT/OFFSET cross-sheet chain),
        // return the error UNWRAPPED. Wrapping it in an Area hides it —
        // Area.IsError is false and Area.AsNumber() coerces the error cell to 0
        // via FirstCell(), so arithmetic (INDIRECT(...)+1) silently produced a
        // wrong number instead of propagating the error. Returning it directly
        // lets ApplyBinaryOp see IsError=true and propagate #NUM! up the chain.
        if (r.Height == 1 && r.Width == 1 && cells[0, 0]?.IsError == true)
            return cells[0, 0];
        // Otherwise always return an Area, even for single-cell refs. This
        // preserves the origin row/col so ROW(OFFSET(...)) / COLUMN(OFFSET(...)) /
        // ADDRESS can answer correctly. Single-cell consumers (AsNumber, AsString)
        // transparently peek the lone cell via FirstCell() in FormulaResult.
        var range = new RangeData(cells) { BaseRow = r.Row, BaseCol = r.Col, BaseSheet = r.Sheet };
        // Same cycle-taint rule as CellMemo: a rect whose materialization tripped
        // the circular-ref fallback holds entry-point-dependent seed values.
        if (rangeMemoKey != null && circularBefore == _session.CircularHits)
            _session.RangeMemo[rangeMemoKey] = range;
        return FormulaResult.Area(range);
    }

    /// <summary>
    /// OFFSET(reference, rows, cols, [height], [width]).
    /// Returns the value at the offset position (single cell) or an Area result
    /// (multi-cell). Outer functions like SUM/AVERAGE consume the Area through
    /// the IsRange handling in helpers.
    /// </summary>
    private const int ExcelMaxRow = 1048576;
    private const int ExcelMaxCol = 16384;

    /// <summary>Coerce a FormulaResult to a number, accepting numeric strings ("1", "2.5").</summary>
    private static double CoerceToNumber(FormulaResult? r)
    {
        if (r == null) return 0;
        if (r.IsString && NumericText.TryParse(r.StringValue, out var v))
            return v;
        return r.AsNumber();
    }

    private FormulaResult? EvalOffset(List<object> args)
    {
        if (args.Count < 3 || args.Count > 5) return FormulaResult.Error("#VALUE!");
        // Accept either a RefArg (literal cell/range token captured by
        // TryParseRefArg) OR a FormulaResult.Area whose underlying RangeData
        // carries BaseRow/BaseCol — produced when a previous OFFSET / INDIRECT
        // returned an Area, or when a defined-name body inlined to such a call.
        // This lets nested OFFSET(OFFSET(...), ...) and three-level defined-name
        // OFFSET chains resolve.
        RefArg baseRef;
        if (args[0] is RefArg ra) baseRef = ra;
        else if (args[0] is FormulaResult fra && fra.IsRange &&
                 fra.RangeValue is { BaseRow: > 0, BaseCol: > 0 } rd)
            baseRef = new RefArg(rd.BaseSheet, rd.BaseCol, rd.BaseRow, rd.Cols, rd.Rows);
        else return FormulaResult.Error("#VALUE!");

        // Bug 1: propagate any error in row/col/height/width before consuming.
        for (int i = 1; i < args.Count; i++)
        {
            if (args[i] is FormulaResult fr && fr.IsError) return fr;
        }

        // Bug 7: numeric strings coerce to numbers.
        int rowOffset = (int)CoerceToNumber(args[1] as FormulaResult);
        int colOffset = (int)CoerceToNumber(args[2] as FormulaResult);
        int height = baseRef.Height;
        int width = baseRef.Width;
        if (args.Count >= 4 && args[3] is FormulaResult hArg) height = (int)CoerceToNumber(hArg);
        if (args.Count >= 5 && args[4] is FormulaResult wArg) width = (int)CoerceToNumber(wArg);
        if (height == 0 || width == 0) return FormulaResult.Error("#REF!");

        var newRow = baseRef.Row + rowOffset;
        var newCol = baseRef.Col + colOffset;
        if (height < 0) { newRow += height + 1; height = -height; }
        if (width < 0) { newCol += width + 1; width = -width; }
        if (newRow < 1 || newCol < 1) return FormulaResult.Error("#REF!");
        // Excel sheet limits: rows 1..1048576, cols 1..16384 (XFD).
        if (newRow > ExcelMaxRow || newCol > ExcelMaxCol) return FormulaResult.Error("#REF!");
        if (newRow + height - 1 > ExcelMaxRow || newCol + width - 1 > ExcelMaxCol) return FormulaResult.Error("#REF!");

        return ResolveRef(new RefArg(baseRef.Sheet, newCol, newRow, width, height));
    }

    /// <summary>
    /// INDIRECT(ref_text). Only the A1-style form is supported (the [a1] argument
    /// is accepted but ignored — R1C1 syntax is not implemented).
    /// </summary>
    private FormulaResult? EvalIndirect(List<object> args)
    {
        if (args.Count < 1) return FormulaResult.Error("#VALUE!");
        // Propagate the original error rather than treating its text as a ref.
        if (args[0] is FormulaResult { IsError: true } e) return e;
        var s = (args[0] as FormulaResult)?.AsString();
        if (string.IsNullOrEmpty(s)) return FormulaResult.Error("#REF!");
        // a1 = FALSE selects absolute R1C1 (R{row}C{col}); relative R[..]C[..]
        // needs the evaluating cell and is not supported.
        bool a1 = !(args.Count > 1 && args[1] is FormulaResult a && a.AsNumber() == 0);
        if (!a1)
        {
            var m = System.Text.RegularExpressions.Regex.Match(s.Trim(), @"^R(\d+)C(\d+)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (!m.Success) return FormulaResult.Error("#REF!");
            return ResolveRef(new RefArg(null, int.Parse(m.Groups[2].Value), int.Parse(m.Groups[1].Value), 1, 1));
        }
        var refArg = ParseRefString(s);
        if (refArg == null) return FormulaResult.Error("#REF!");
        return ResolveRef(refArg);
    }

    // ==================== Structured table references (B1) ====================
    // `Table[Col]`, `Table[[Col1]:[Col2]]`, `Table[#Data]/[#All]/[#Headers]`
    // fold to plain A1 range tokens at tokenize time from the workbook's
    // CURRENT table definition — every evaluation re-folds, so rows inserted
    // or deleted inside the table are picked up automatically (Excel recompute
    // parity). Row-context forms (`[@Col]`, `[#This Row]`) and compound item
    // strings (`[[#Data],[Col]]`) are explicitly rejected with #NAME? for v1.

    /// <summary>One workbook table: owning sheet name, "A1:D10"-style
    /// reference, ordered column names, header/totals row counts.</summary>
    private sealed record TableDef(string? Sheet, int StartCol, int StartRow, int EndCol, int EndRow,
        List<string> Columns, uint HeaderRows, uint TotalsRows);

    /// <summary>
    /// Scan the workbook's table definitions for <paramref name="displayName"/>
    /// (matches the Table @name or @displayName attribute, case-insensitive —
    /// Excel treats them as one namespace). Returns null when the workbook has
    /// no such table or no workbook context at all.
    /// </summary>
    private TableDef? FindTableDef(string displayName)
    {
        if (_workbookPart == null) return null;
        foreach (var sheet in _workbookPart.Workbook?.Descendants<Sheet>().ToList()
                 ?? new List<Sheet>())
        {
            if (sheet.Id?.Value == null) continue;
            WorksheetPart? wsPart;
            try { wsPart = (WorksheetPart?)_workbookPart.GetPartById(sheet.Id.Value!); }
            catch { continue; }
            if (wsPart == null) continue;
            foreach (var tdp in wsPart.TableDefinitionParts)
            {
                var t = tdp.Table;
                var name = t.Name?.Value;
                var display = t.DisplayName?.Value;
                if (!string.Equals(name, displayName, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(display, displayName, StringComparison.OrdinalIgnoreCase))
                    continue;
                var reference = t.Reference?.Value;
                if (string.IsNullOrEmpty(reference) || !reference.Contains(':'))
                    continue;
                var parts = reference.Split(':');
                var (startColLetter, startRow) = ParseRef(StripDollar(parts[0]));
                var (endColLetter, endRow) = ParseRef(StripDollar(parts[1]));
                var startCol = ColToIndex(startColLetter);
                var endCol = ColToIndex(endColLetter);
                var columns = t.TableColumns?.Elements<TableColumn>()
                    .Select(c => c.Name?.Value ?? "").ToList() ?? new List<string>();
                return new TableDef(sheet.Name?.Value,
                    Math.Min(startCol, endCol), Math.Min(startRow, endRow),
                    Math.Max(startCol, endCol), Math.Max(startRow, endRow),
                    columns, t.HeaderRowCount?.Value ?? 1U, t.TotalsRowCount?.Value ?? 0U);
            }
        }
        return null;
    }

    /// <summary>
    /// Fold a structured reference (cursor on the opening '[') into a Range /
    /// SheetRange token over the table's current definition. Returns null when
    /// no bracket follows; throws NameResolutionException for an unknown table,
    /// an unknown column, or an unsupported form — surfacing #NAME? in the cell,
    /// matching Excel's invalid-structured-ref behavior.
    /// </summary>
    private Token? TryFoldStructuredRef(string tableName, string formula, ref int i)
    {
        if (i >= formula.Length || formula[i] != '[') return null;
        var def = FindTableDef(tableName)
            ?? throw new NameResolutionException($"unknown table '{tableName}'");

        var pos = i + 1; // past the outer '['
        string? special = null;
        var columns = new List<string>();
        var isRowContext = false;
        // Only the double-bracket forms (`[[Col]]`, `[[Col1]:[Col2]]`) wrap the
        // item in an EXTRA bracket pair — the bare (`[Col]`) and special
        // (`[#Data]`) forms end at their first ']'.
        var expectOuterClose = false;

        if (pos < formula.Length && formula[pos] == '@')
        {
            isRowContext = true;
            ReadUntilClosingBracket(formula, ref pos, tableName);
        }
        else if (pos < formula.Length && formula[pos] == '#')
        {
            special = ReadUntilClosingBracket(formula, ref pos, tableName);
            if (pos < formula.Length && formula[pos] == ',')
                throw new NameResolutionException(
                    $"unsupported compound structured reference on table '{tableName}' (supported: [Col], [[Col1]:[Col2]], [#Data]/[#All]/[#Headers])");
        }
        else if (pos < formula.Length && formula[pos] == '[')
        {
            pos++;
            columns.Add(ReadUntilClosingBracket(formula, ref pos, tableName));
            expectOuterClose = true;
        }
        else
        {
            columns.Add(ReadUntilClosingBracket(formula, ref pos, tableName));
        }

        // Span form: [[Col1]:[Col2]].
        if (!isRowContext && special == null && expectOuterClose
            && pos < formula.Length && formula[pos] == ':'
            && pos + 1 < formula.Length && formula[pos + 1] == '[')
        {
            pos += 2;
            columns.Add(ReadUntilClosingBracket(formula, ref pos, tableName));
        }

        if (expectOuterClose)
        {
            if (pos >= formula.Length || formula[pos] != ']')
                throw new NameResolutionException($"unterminated structured reference on table '{tableName}'");
            pos++;
        }

        if (isRowContext || (special is not null && special.Equals("#This Row", StringComparison.OrdinalIgnoreCase)))
            throw new NameResolutionException(
                $"row-context structured reference (Table[@Col] / [#This Row]) is unsupported: '{tableName}'");

        // Data body excludes header and totals rows.
        var dataFirstRow = def.StartRow + (int)def.HeaderRows;
        var dataLastRow = def.EndRow - (int)def.TotalsRows;

        int colMin, colMax, rowMin, rowMax;
        if (special != null)
        {
            if (special.Equals("#Data", StringComparison.OrdinalIgnoreCase))
                (colMin, colMax, rowMin, rowMax) = (def.StartCol, def.EndCol, dataFirstRow, dataLastRow);
            else if (special.Equals("#All", StringComparison.OrdinalIgnoreCase))
                (colMin, colMax, rowMin, rowMax) = (def.StartCol, def.EndCol, def.StartRow, def.EndRow);
            else if (special.Equals("#Headers", StringComparison.OrdinalIgnoreCase))
            {
                if (def.HeaderRows == 0)
                    throw new NameResolutionException($"table '{tableName}' has no header row");
                (colMin, colMax, rowMin, rowMax) =
                    (def.StartCol, def.EndCol, def.StartRow, def.StartRow + (int)def.HeaderRows - 1);
            }
            else if (special.Equals("#Totals", StringComparison.OrdinalIgnoreCase))
            {
                if (def.TotalsRows == 0)
                    throw new NameResolutionException($"table '{tableName}' has no totals row");
                (colMin, colMax, rowMin, rowMax) =
                    (def.StartCol, def.EndCol, def.EndRow - (int)def.TotalsRows + 1, def.EndRow);
            }
            else
                throw new NameResolutionException(
                    $"unsupported special item '{special}' on table '{tableName}' (supported: #Data, #All, #Headers, #Totals)");
        }
        else
        {
            if (columns.Count == 0 || columns.Count > 2)
                throw new NameResolutionException(
                    $"unsupported structured reference form on table '{tableName}' (supported: [Col], [[Col1]:[Col2]], [#Data]/[#All]/[#Headers])");
            var idx = ResolveTableColumn(def, columns[0], tableName);
            var col1 = def.StartCol + idx;
            var col2 = columns.Count == 2 ? def.StartCol + ResolveTableColumn(def, columns[1], tableName) : col1;
            if (dataFirstRow > dataLastRow)
                throw new NameResolutionException($"table '{tableName}' has no data rows");
            (colMin, colMax) = (Math.Min(col1, col2), Math.Max(col1, col2));
            (rowMin, rowMax) = (dataFirstRow, dataLastRow);
        }

        i = pos;
        var rangeText = $"{IndexToCol(colMin)}{rowMin}:{IndexToCol(colMax)}{rowMax}";
        var sameSheet = string.IsNullOrEmpty(def.Sheet)
            || string.Equals(def.Sheet, _sheetKey, StringComparison.OrdinalIgnoreCase);
        return sameSheet
            ? new Token(TT.Range, rangeText)
            : new Token(TT.SheetRange, $"{def.Sheet}!{rangeText}");
    }

    /// <summary>Resolve a column name against the table's ordered columns
    /// (case-insensitive; the table's first column is offset 0).</summary>
    private static int ResolveTableColumn(TableDef def, string name, string tableName)
    {
        for (var k = 0; k < def.Columns.Count; k++)
            if (string.Equals(def.Columns[k], name, StringComparison.OrdinalIgnoreCase))
                return k;
        throw new NameResolutionException(
            $"unknown column '{name}' in table '{tableName}' (valid: {string.Join(", ", def.Columns)})");
    }

    /// <summary>Consume characters up to (and past) the next ']' and return the
    /// trimmed text. Escaped brackets inside column names are not supported in
    /// v1 — a column named with ']' cannot be referenced.</summary>
    private string ReadUntilClosingBracket(string formula, ref int pos, string tableName)
    {
        var end = formula.IndexOf(']', pos);
        if (end < 0)
            throw new NameResolutionException($"unterminated structured reference on table '{tableName}'");
        var text = formula[pos..end].Trim();
        pos = end + 1;
        return text;
    }
}
