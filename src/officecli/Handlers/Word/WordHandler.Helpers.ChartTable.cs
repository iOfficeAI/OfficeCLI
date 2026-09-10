// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // Convention for `add --type chart --prop dataTable=<tablePath>`:
    //   - The FIRST data row is the HEADER: each column (from index 1 on) becomes
    //     one series whose name is that column's header cell; index 0 is the
    //     categories label header and is not a series.
    //   - The FIRST data COLUMN (col 0) under the header is the categories axis
    //     (row labels). Every remaining column is a series of numeric values.
    //   - If the first row is ENTIRELY numeric (no real header), the table is
    //     treated as headerless: each column is a series named "Series N", and
    //     the categories fall back to first column if non-numeric else 1..N.
    //   - Cells that cannot parse as a number are skipped with a warning — never
    //     aborts, so a mixed numeric/label table still yields the numeric points.
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

        // Heuristic header: row 0 is a header only when it contains at least one
        // non-numeric label — i.e. it is NOT entirely numeric, AND it has more
        // than the first cell filled. A fully-numeric first row is data.
        var firstRow = rows[0];
        var hasHeader = firstRow.Count > 1
            && firstRow.Any(c => !string.IsNullOrWhiteSpace(c) && !IsNumericCell(c));

        var series = new List<(string name, double[] values)>();
        List<string>? categories = null;

        if (hasHeader)
        {
            // Series = columns 1..; categories = column 0 beneath the header.
            categories = new List<string>();
            for (int ci = 1; ci < colCount; ci++)
            {
                var name = ci < firstRow.Count && !string.IsNullOrWhiteSpace(firstRow[ci])
                    ? firstRow[ci].Trim()
                    : $"Series {ci}";
                var vals = new List<double>();
                for (int ri = 1; ri < rows.Count; ri++)
                {
                    var cellText = ri < rows[ri].Count ? rows[ri][ci] : "";
                    if (TryParseNumeric(cellText, out var d)) vals.Add(d);
                    else if (ci == 0) { /* categories column */ }
                }
                series.Add((name, vals.ToArray()));
            }
            // Categories labels: column 0 of each data row (non-empty).
            for (int ri = 1; ri < rows.Count; ri++)
            {
                var label = ri < rows[ri].Count ? rows[ri][0] : "";
                if (!string.IsNullOrWhiteSpace(label)) categories.Add(label.Trim());
            }
        }
        else
        {
            // Headerless: every column is a series (default names); categories
            // derive from column 0 when it is non-numeric, else 1..N.
            for (int ci = 0; ci < colCount; ci++)
            {
                var vals = new List<double>();
                for (int ri = 0; ri < rows.Count; ri++)
                {
                    var cellText = ri < rows[ri].Count ? rows[ri][ci] : "";
                    if (TryParseNumeric(cellText, out var d)) vals.Add(d);
                }
                series.Add(($"Series {ci + 1}", vals.ToArray()));
            }
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