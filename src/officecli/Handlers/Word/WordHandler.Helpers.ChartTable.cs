// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // Convention for `add --type chart --prop sourceTable=<tablePath>`:
    //   - The FIRST data row is the HEADER: each column (from index 1 on) becomes
    //     one series whose name is that column's header cell; index 0 is the
    //     categories label header and is not a series. Only the SERIES columns
    //     (1..) vote on header detection — a label-first data column (e.g. a
    //     headerless "Q1|100" table) must not be mistaken for a header row.
    //   - The FIRST data COLUMN (col 0) under the header is the categories axis
    //     (row labels). Every remaining column is a series of numeric values.
    //   - If no series column carries a header label, the table is treated as
    //     headerless: when col 0 is entirely non-numeric it becomes the
    //     categories axis and columns 1.. become "Series N"; otherwise every
    //     column is a series named "Series N" and categories fall back to 1..N.
    //   - Extraction is ROW-ATOMIC: a data row contributes its category label
    //     AND all of its series values, or nothing at all. Skipping a bad cell
    //     while keeping its label would shift categories against values, so a
    //     row with any non-numeric series cell is dropped whole with a warning
    //     — never aborts, so a mixed numeric/label table still charts the
    //     clean rows.
    private static (string[]? Categories, List<(string name, double[] values)> Series, List<string> Warnings)
        ExtractChartDataFromTable(Table tbl)
    {
        var warnings = new List<string>();

        // Materialize the table into a rectangular string grid, so the
        // header/column/row heuristics operate on stable positions.
        var rows = new List<List<string>>();
        foreach (var row in tbl.Elements<TableRow>())
        {
            var cells = new List<string>();
            foreach (var cell in row.Elements<TableCell>())
            {
                var para = cell.Elements<Paragraph>().FirstOrDefault();
                cells.Add(para != null ? GetParagraphText(para) : "");
            }
            if (cells.Count > 0 || rows.Count == 0)
                rows.Add(cells);
        }
        if (rows.Count == 0)
            return (null, new List<(string, double[])>{ }, warnings);

        var colCount = rows.Max(r => r.Count);
        if (colCount == 0)
            return (null, new List<(string, double[])>{ }, warnings);

        bool IsNumericCell(string s)
            => TryParseNumeric(s, out _);

        // Heuristic header: row 0 is a header only when at least one SERIES
        // column (index 1 on) carries a non-numeric label. Col 0 is the
        // categories corner and never votes — otherwise a headerless table
        // like "Q1|100 / Q2|150" would eat its first data row as a header.
        var firstRow = rows[0];
        var hasHeader = firstRow.Count > 1
            && firstRow.Skip(1).Any(c => !string.IsNullOrWhiteSpace(c) && !IsNumericCell(c));

        var series = new List<(string name, double[] values)>();
        List<string>? categories = null;

        // Ragged-row-safe cell read: a short row yields "" (non-numeric) for
        // missing cells instead of throwing IndexOutOfRangeException.
        static string CellAt(List<string> row, int ci)
            => ci < row.Count ? row[ci] : "";

        if (hasHeader)
        {
            // Series = columns 1.., named by the header row; categories = col 0.
            var names = new List<string>();
            for (int ci = 1; ci < colCount; ci++)
                names.Add(ci < firstRow.Count && !string.IsNullOrWhiteSpace(firstRow[ci])
                    ? firstRow[ci].Trim()
                    : $"Series {ci}");
            var valLists = names.Select(_ => new List<double>()).ToList();
            categories = new List<string>();
            for (int ri = 1; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                // Row-atomic: parse every series cell first; a single bad cell
                // drops the whole row (label included) with a warning, so the
                // categories array can never drift out of alignment with values.
                var parsed = new double[names.Count];
                var ok = true;
                for (int k = 0; k < names.Count; k++)
                {
                    var cellText = CellAt(row, k + 1);
                    if (!TryParseNumeric(cellText, out parsed[k]))
                    {
                        warnings.Add(
                            $"sourceTable: skipped row {ri + 1} " +
                            $"(column {k + 2} value '{cellText}' is not numeric); " +
                            "row dropped so categories stay aligned with values.");
                        ok = false;
                        break;
                    }
                }
                if (!ok) continue;
                categories.Add(CellAt(row, 0).Trim());
                for (int k = 0; k < names.Count; k++) valLists[k].Add(parsed[k]);
            }
            for (int k = 0; k < names.Count; k++) series.Add((names[k], valLists[k].ToArray()));
        }
        else
        {
            // Headerless: when col 0 is entirely non-numeric (label column) it
            // becomes the categories axis and columns 1.. become the series;
            // otherwise every column is a numeric series and categories stay
            // null (the chart builder falls back to 1..N).
            var col0IsLabels = colCount > 1
                && rows.All(r => r.Count == 0 || !IsNumericCell(r[0]));
            var firstDataCol = col0IsLabels ? 1 : 0;
            var names = new List<string>();
            for (int ci = firstDataCol; ci < colCount; ci++)
                names.Add($"Series {ci - firstDataCol + 1}");
            var valLists = names.Select(_ => new List<double>()).ToList();
            var headerlessCats = col0IsLabels ? new List<string>() : null;
            for (int ri = 0; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                var parsed = new double[names.Count];
                var ok = true;
                for (int k = 0; k < names.Count; k++)
                {
                    var cellText = CellAt(row, firstDataCol + k);
                    if (!TryParseNumeric(cellText, out parsed[k]))
                    {
                        warnings.Add(
                            $"sourceTable: skipped row {ri + 1} " +
                            $"(column {firstDataCol + k + 1} value '{cellText}' is not numeric); " +
                            "row dropped so categories stay aligned with values.");
                        ok = false;
                        break;
                    }
                }
                if (!ok) continue;
                headerlessCats?.Add(CellAt(row, 0).Trim());
                for (int k = 0; k < names.Count; k++) valLists[k].Add(parsed[k]);
            }
            // Drop series with no surviving points (e.g. an all-label column);
            // an empty series would render as a phantom legend entry.
            for (int k = 0; k < names.Count; k++)
            {
                if (valLists[k].Count == 0)
                    warnings.Add($"sourceTable: column {firstDataCol + k + 1} has no numeric values; dropped from the chart.");
                else
                    series.Add((names[k], valLists[k].ToArray()));
            }
            categories = headerlessCats;
        }

        return (categories?.Count > 0 ? categories.ToArray() : null, series, warnings);
    }

    // Lenient numeric parse: strips thousands commas, a leading currency symbol,
    // and a trailing '%' (keeping the raw magnitude), then parses as invariant
    // double. False for empty / non-numeric — callers warn+skip.
    private static bool TryParseNumeric(string raw, out double value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(raw)) return false;
        var s = raw.Trim();
        s = s.Replace(",", "").Replace("$", "").Replace("¥", "").Replace("%", "");
        if (s.Length == 0) return false;
        return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowCurrencySymbol,
            CultureInfo.InvariantCulture, out value);
    }

    // Resolve a document-rooted table path (`/body/t[1]`, `/body/t[2]`) to the
    // Table element, producing a human-readable context on failure (mirrors the
    // Set path's error shape). Returns null when unresolvable; `resolveErr`
    // carries the path-not-found detail.
    private Table? ResolveTableForChart(string tablePath, out string resolveErr)
    {
        resolveErr = "";
        if (string.IsNullOrWhiteSpace(tablePath)) return null;
        Table? table;
        try
        {
            var parts = ParsePath(tablePath);
            table = NavigateToElement(parts, out var ctx) as Table;
            if (table == null)
                resolveErr = ctx != null ? $"Path resolved but is not a table. {ctx}" : "";
        }
        catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException)
        {
            table = null;
            resolveErr = $" Path parse error: {ex.Message}";
        }
        return table;
    }
}