// Plan 84/99.9.H: Document Diff/Compare
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // H1: Block similarity threshold
    private const double BlockSimilarityThreshold = 0.4;

    // H2: Table similarity weights
    private const double TableDimWeight = 0.3;
    private const double TableContentWeight = 0.7;

    // H3: Max matrix cells for Levenshtein
    private const long MaxDiffMatrixCells = 10_000_000;

    /// <summary>Jaccard similarity between two strings based on word tokens.</summary>
    internal static double ComputeBlockSimilarity(string a, string b)
    {
        if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) return 1.0;
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;

        var tokensA = NormalizeForSimilarity(a).Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();
        var tokensB = NormalizeForSimilarity(b).Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();

        if (tokensA.Count == 0 && tokensB.Count == 0) return 1.0;

        int intersection = tokensA.Intersect(tokensB).Count();
        int union = tokensA.Union(tokensB).Count();

        return union == 0 ? 0.0 : (double)intersection / union;
    }

    /// <summary>Compare two documents block-by-block.</summary>
    public JsonNode CompareText(HwpxHandler other)
    {
        var linesA = ExtractTextLines();
        var linesB = other.ExtractTextLines();
        var result = ComputeLineDiff(linesA, linesB);

        return new JsonObject
        {
            ["mode"] = "text",
            ["linesA"] = linesA.Length,
            ["linesB"] = linesB.Length,
            ["changes"] = result
        };
    }

    /// <summary>
    /// Plan 99.9.I4: LCS-based line diff with fallback to linear scan for large inputs.
    /// Uses proper LCS DP alignment for accurate diff of insertions/deletions.
    /// </summary>
    internal JsonArray ComputeLineDiff(string[] linesA, string[] linesB)
    {
        long matrixCells = (long)(linesA.Length + 1) * (linesB.Length + 1);
        if (matrixCells > MaxDiffMatrixCells)
            return ComputeLineDiffLinear(linesA, linesB); // fallback for huge docs

        return ComputeLineDiffLcs(linesA, linesB);
    }

    /// <summary>LCS DP-based diff — accurate alignment for moderate-size documents.</summary>
    private JsonArray ComputeLineDiffLcs(string[] a, string[] b)
    {
        int n = a.Length, m = b.Length;

        // Build LCS table
        var dp = new int[n + 1, m + 1];
        for (int i = 1; i <= n; i++)
            for (int j = 1; j <= m; j++)
                dp[i, j] = a[i - 1] == b[j - 1]
                    ? dp[i - 1, j - 1] + 1
                    : Math.Max(dp[i - 1, j], dp[i, j - 1]);

        // Backtrack to produce diff
        var result = new List<(string type, int? la, int? lb, string? ta, string? tb)>();
        int ia = n, ib = m;
        while (ia > 0 && ib > 0)
        {
            if (a[ia - 1] == b[ib - 1])
            {
                result.Add(("unchanged", ia, ib, a[ia - 1], null));
                ia--; ib--;
            }
            else if (dp[ia - 1, ib] >= dp[ia, ib - 1])
            {
                result.Add(("removed", ia, null, a[ia - 1], null));
                ia--;
            }
            else
            {
                result.Add(("added", null, ib, null, b[ib - 1]));
                ib--;
            }
        }
        while (ia > 0) { result.Add(("removed", ia, null, a[ia - 1], null)); ia--; }
        while (ib > 0) { result.Add(("added", null, ib, null, b[ib - 1])); ib--; }

        result.Reverse();

        // Post-process: detect "modified" (adjacent removed+added with similar content)
        var output = new JsonArray();
        int idx = 0;
        while (idx < result.Count)
        {
            var (type, la, lb, ta, tb) = result[idx];
            if (type == "removed" && idx + 1 < result.Count && result[idx + 1].type == "added")
            {
                var next = result[idx + 1];
                var sim = ComputeBlockSimilarity(ta!, next.tb!);
                if (sim >= BlockSimilarityThreshold)
                {
                    output.Add(MakeDiffEntry("modified", la, next.lb, ta, next.tb));
                    idx += 2;
                    continue;
                }
            }
            output.Add(MakeDiffEntry(type, la, lb, ta, tb));
            idx++;
        }
        return output;
    }

    /// <summary>Linear scan fallback for very large documents (exceeds LCS matrix limit).</summary>
    private JsonArray ComputeLineDiffLinear(string[] linesA, string[] linesB)
    {
        var result = new JsonArray();
        int ia = 0, ib = 0;

        while (ia < linesA.Length && ib < linesB.Length)
        {
            if (linesA[ia] == linesB[ib])
            {
                result.Add(MakeDiffEntry("unchanged", ia + 1, ib + 1, linesA[ia], null));
                ia++; ib++;
                continue;
            }

            var sim = ComputeBlockSimilarity(linesA[ia], linesB[ib]);
            if (sim >= BlockSimilarityThreshold)
            {
                result.Add(MakeDiffEntry("modified", ia + 1, ib + 1, linesA[ia], linesB[ib]));
                ia++; ib++;
            }
            else
            {
                bool foundA = false, foundB = false;
                for (int lookahead = 1; lookahead <= 5; lookahead++)
                {
                    if (ib + lookahead < linesB.Length && linesA[ia] == linesB[ib + lookahead])
                    {
                        for (int j = 0; j < lookahead; j++)
                            result.Add(MakeDiffEntry("added", null, ib + j + 1, null, linesB[ib + j]));
                        ib += lookahead;
                        foundB = true;
                        break;
                    }
                    if (ia + lookahead < linesA.Length && linesA[ia + lookahead] == linesB[ib])
                    {
                        for (int j = 0; j < lookahead; j++)
                            result.Add(MakeDiffEntry("removed", ia + j + 1, null, linesA[ia + j], null));
                        ia += lookahead;
                        foundA = true;
                        break;
                    }
                }
                if (!foundA && !foundB)
                {
                    result.Add(MakeDiffEntry("modified", ia + 1, ib + 1, linesA[ia], linesB[ib]));
                    ia++; ib++;
                }
            }
        }

        while (ia < linesA.Length)
        { result.Add(MakeDiffEntry("removed", ia + 1, null, linesA[ia], null)); ia++; }
        while (ib < linesB.Length)
        { result.Add(MakeDiffEntry("added", null, ib + 1, null, linesB[ib])); ib++; }

        return result;
    }

    /// <summary>Extract text lines for diff.</summary>
    internal string[] ExtractTextLines()
        => ViewAsText()
            .Split('\n')
            .Select(l => Regex.Replace(l, @"^\d+\.\s*", "").Trim())
            .Where(l => !string.IsNullOrEmpty(l))
            .ToArray();

    private static JsonObject MakeDiffEntry(string type, int? lineA, int? lineB, string? textA, string? textB)
    {
        var obj = new JsonObject { ["type"] = type };
        if (lineA.HasValue) obj["lineA"] = lineA.Value;
        if (lineB.HasValue) obj["lineB"] = lineB.Value;
        if (textA != null) obj["textA"] = textA;
        if (textB != null) obj["textB"] = textB;
        return obj;
    }

    // --- H2: Table comparison ---

    /// <summary>Compute similarity between two tables (dimensions + content).</summary>
    internal static double ComputeTableSimilarity(string?[,] gridA, string?[,] gridB)
    {
        int rowsA = gridA.GetLength(0), colsA = gridA.GetLength(1);
        int rowsB = gridB.GetLength(0), colsB = gridB.GetLength(1);

        double dimSim = 0;
        if (rowsA + rowsB > 0)
            dimSim += (double)Math.Min(rowsA, rowsB) / Math.Max(rowsA, rowsB) * 0.5;
        if (colsA + colsB > 0)
            dimSim += (double)Math.Min(colsA, colsB) / Math.Max(colsA, colsB) * 0.5;

        int minRows = Math.Min(rowsA, rowsB), minCols = Math.Min(colsA, colsB);
        int matchCount = 0, totalCount = 0;
        for (int r = 0; r < minRows; r++)
            for (int c = 0; c < minCols; c++)
            {
                totalCount++;
                var cellA = gridA[r, c]?.Trim() ?? "";
                var cellB = gridB[r, c]?.Trim() ?? "";
                if (cellA == cellB) matchCount++;
            }
        totalCount += Math.Max(0, rowsA * colsA - minRows * minCols);
        totalCount += Math.Max(0, rowsB * colsB - minRows * minCols);

        double contentSim = totalCount == 0 ? 1.0 : (double)matchCount / totalCount;
        return TableDimWeight * dimSim + TableContentWeight * contentSim;
    }

    /// <summary>Compare tables between two documents by position index.</summary>
    public JsonNode CompareTables(HwpxHandler other)
    {
        var tablesA = ExtractAllTableGrids();
        var tablesB = other.ExtractAllTableGrids();

        var result = new JsonArray();
        int maxTables = Math.Max(tablesA.Count, tablesB.Count);

        for (int t = 0; t < maxTables; t++)
        {
            if (t >= tablesA.Count)
            {
                result.Add(new JsonObject { ["table"] = t + 1, ["type"] = "added" });
                continue;
            }
            if (t >= tablesB.Count)
            {
                result.Add(new JsonObject { ["table"] = t + 1, ["type"] = "removed" });
                continue;
            }

            var gridA = tablesA[t].Grid;
            var gridB = tablesB[t].Grid;
            var similarity = ComputeTableSimilarity(gridA, gridB);

            var cellDiffs = new JsonArray();
            int maxRows = Math.Max(gridA.GetLength(0), gridB.GetLength(0));
            int maxCols = Math.Max(gridA.GetLength(1), gridB.GetLength(1));

            for (int r = 0; r < maxRows; r++)
                for (int c = 0; c < maxCols; c++)
                {
                    var cellA = (r < gridA.GetLength(0) && c < gridA.GetLength(1)) ? gridA[r, c] : null;
                    var cellB = (r < gridB.GetLength(0) && c < gridB.GetLength(1)) ? gridB[r, c] : null;
                    if (cellA != cellB)
                    {
                        cellDiffs.Add(new JsonObject
                        {
                            ["row"] = r + 1, ["col"] = c + 1,
                            ["old"] = cellA, ["new"] = cellB
                        });
                    }
                }

            result.Add(new JsonObject
            {
                ["table"] = t + 1,
                ["type"] = cellDiffs.Count > 0 ? "modified" : "unchanged",
                ["similarity"] = Math.Round(similarity, 3),
                ["changes"] = cellDiffs
            });
        }

        return new JsonObject { ["mode"] = "table", ["tables"] = result };
    }

    /// <summary>Extract cell text grids for all tables.</summary>
    internal List<(string Path, string?[,] Grid)> ExtractAllTableGrids()
    {
        var result = new List<(string Path, string?[,] Grid)>();
        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            var (grid, _) = BuildTableGrid(tbl);
            var textGrid = new string?[grid.GetLength(0), grid.GetLength(1)];
            for (int r = 0; r < grid.GetLength(0); r++)
                for (int c = 0; c < grid.GetLength(1); c++)
                    textGrid[r, c] = grid[r, c] != null ? ExtractCellText(grid[r, c]!).Trim() : null;
            var path = $"/section[{sec.Index + 1}]/tbl[{localTblIdx + 1}]";
            result.Add((path, textGrid));
        }
        return result;
    }

    // --- H3: Levenshtein distance with fallback ---

    /// <summary>Levenshtein edit distance with matrix size limit.</summary>
    internal static int LevenshteinDistance(string[] a, string[] b)
    {
        long matrixSize = (long)(a.Length + 1) * (b.Length + 1);
        if (matrixSize > MaxDiffMatrixCells)
            return LevenshteinFallback(a, b);

        int n = a.Length, m = b.Length;
        var prev = new int[m + 1];
        var curr = new int[m + 1];

        for (int j = 0; j <= m; j++) prev[j] = j;

        for (int i = 1; i <= n; i++)
        {
            curr[0] = i;
            for (int j = 1; j <= m; j++)
            {
                int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                curr[j] = Math.Min(
                    Math.Min(prev[j] + 1, curr[j - 1] + 1),
                    prev[j - 1] + cost);
            }
            (prev, curr) = (curr, prev);
        }

        return prev[m];
    }

    private static int LevenshteinFallback(string[] a, string[] b)
    {
        const int sampleSize = 500;

        int headMatches = 0;
        int headLen = Math.Min(Math.Min(a.Length, b.Length), sampleSize);
        for (int i = 0; i < headLen; i++)
            if (a[i] == b[i]) headMatches++;

        int tailMatches = 0;
        int tailLen = Math.Min(Math.Min(a.Length, b.Length), sampleSize);
        for (int i = 0; i < tailLen; i++)
            if (a[a.Length - 1 - i] == b[b.Length - 1 - i]) tailMatches++;

        double matchRate = (headLen + tailLen) > 0
            ? (double)(headMatches + tailMatches) / (headLen + tailLen)
            : 0;
        int maxLen = Math.Max(a.Length, b.Length);

        return (int)((1 - matchRate) * maxLen) + Math.Abs(a.Length - b.Length);
    }

    // --- H4: Text normalization for similarity ---

    /// <summary>Normalize text for similarity comparison.</summary>
    internal static string NormalizeForSimilarity(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";

        text = HwpxKorean.Normalize(text);
        text = text.ToLowerInvariant();
        text = Regex.Replace(text, @"[^\p{L}\p{N}\s]", " ");
        text = Regex.Replace(text, @"\s+", " ").Trim();

        return text;
    }

    // --- H5: Page range compare ---

    /// <summary>Parse "1-3,5,7-9" into 1-based page numbers.</summary>
    internal static HashSet<int>? ParsePageRange(string? pageRange)
    {
        if (string.IsNullOrWhiteSpace(pageRange)) return null;

        var pages = new HashSet<int>();
        var parts = pageRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var part in parts)
        {
            var dashIdx = part.IndexOf('-');
            if (dashIdx > 0 && dashIdx < part.Length - 1)
            {
                if (int.TryParse(part[..dashIdx].Trim(), out var start)
                    && int.TryParse(part[(dashIdx + 1)..].Trim(), out var end))
                {
                    if (start > end) (start, end) = (end, start);
                    end = Math.Min(end, start + 999);
                    if (end > 100_000) end = start + 999; // absolute safety cap
                    for (int p = start; p <= end; p++)
                    {
                        pages.Add(p);
                        if (p == end) break; // prevent int overflow wrap
                    }
                }
            }
            else
            {
                if (int.TryParse(part.Trim(), out var page))
                    pages.Add(page);
            }
        }

        return pages.Count > 0 ? pages : null;
    }

    /// <summary>Compare text for specific page ranges.</summary>
    public JsonNode CompareTextRange(HwpxHandler other, string? pagesA, string? pagesB)
    {
        var rangeA = ParsePageRange(pagesA);
        var rangeB = ParsePageRange(pagesB);

        var linesA = ExtractTextLinesFiltered(rangeA);
        var linesB = other.ExtractTextLinesFiltered(rangeB);
        var changes = ComputeLineDiff(linesA, linesB);

        return new JsonObject
        {
            ["mode"] = "text",
            ["pagesA"] = pagesA ?? "all",
            ["pagesB"] = pagesB ?? "all",
            ["linesA"] = linesA.Length,
            ["linesB"] = linesB.Length,
            ["changes"] = changes
        };
    }

    /// <summary>Extract text lines filtered by section indices.</summary>
    internal string[] ExtractTextLinesFiltered(HashSet<int>? sectionFilter)
    {
        if (sectionFilter == null) return ExtractTextLines();

        var lines = new List<string>();
        foreach (var (section, para, path) in _doc.AllContentInOrder())
        {
            if (!sectionFilter.Contains(section.Index + 1)) continue;
            var text = HwpxKorean.Normalize(ExtractParagraphText(para)).Trim();
            if (!string.IsNullOrWhiteSpace(text))
                lines.Add(text);
        }
        return lines.ToArray();
    }
}
