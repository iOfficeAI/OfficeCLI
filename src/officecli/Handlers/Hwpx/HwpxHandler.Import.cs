// Plan 85: Markdown → HWPX Import
// Minimal GFM parser: headings, paragraphs, tables.
// Uses existing Add/Set infrastructure internally.

using System.Text.RegularExpressions;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    /// <summary>
    /// Import Markdown content into the current HWPX document.
    /// Supports: headings (#-######), paragraphs, GFM tables, bold, italic.
    /// </summary>
    public int ImportMarkdown(string markdown, string? align = null)
    {
        var lines = markdown.Split('\n');
        int blockCount = 0;

        int i = 0;

        // Skip YAML frontmatter (--- ... ---)
        if (i < lines.Length && lines[i].TrimEnd('\r') == "---")
        {
            i++;
            while (i < lines.Length && lines[i].TrimEnd('\r') != "---") i++;
            if (i < lines.Length) i++; // skip closing ---
        }
        while (i < lines.Length)
        {
            var line = lines[i].TrimEnd('\r');

            // Skip empty lines
            if (string.IsNullOrWhiteSpace(line)) { i++; continue; }

            // Skip code fence markers (``` or ~~~)
            if (Regex.IsMatch(line, @"^(`{3}|~{3})")) { i++; continue; }

            // Skip horizontal rules (--- or ***)
            if (Regex.IsMatch(line.Trim(), @"^[-*_]{3,}$")) { i++; continue; }

            // Skip image-only lines: ![alt](url)
            if (Regex.IsMatch(line.Trim(), @"^!\[.*\]\(.*\)$")) { i++; continue; }

            // G7: Blockquote: > text (preserve marker for round-trip)
            if (line.TrimStart().StartsWith('>'))
            {
                var text = StripInlineMarkdown(line.TrimStart().TrimStart('>').Trim());
                if (!string.IsNullOrEmpty(text))
                {
                    var props = new Dictionary<string, string> { ["text"] = $"> {text}" };
                    if (align != null) props["align"] = align.ToUpperInvariant();
                    Add("/section[1]", "paragraph", null, props);
                    blockCount++;
                }
                i++;
                continue;
            }

            // G7: Unordered list: - item, * item, + item
            var ulMatch = Regex.Match(line, @"^\s*[-*+]\s+(.+)$");
            if (ulMatch.Success)
            {
                var text = StripInlineMarkdown(ulMatch.Groups[1].Value);
                var props = new Dictionary<string, string> { ["text"] = $"  - {text}" };
                if (align != null) props["align"] = align.ToUpperInvariant();
                Add("/section[1]", "paragraph", null, props);
                blockCount++;
                i++;
                continue;
            }

            // G7: Ordered list: 1. item
            var olMatch = Regex.Match(line, @"^\s*(\d+)\.\s+(.+)$");
            if (olMatch.Success)
            {
                var num = olMatch.Groups[1].Value;
                var text = StripInlineMarkdown(olMatch.Groups[2].Value);
                var props = new Dictionary<string, string> { ["text"] = $"  {num}. {text}" };
                if (align != null) props["align"] = align.ToUpperInvariant();
                Add("/section[1]", "paragraph", null, props);
                blockCount++;
                i++;
                continue;
            }

            // Heading: # ... ######
            var headingMatch = Regex.Match(line, @"^(#{1,6})\s+(.+)$");
            if (headingMatch.Success)
            {
                var level = headingMatch.Groups[1].Value.Length;
                var text = StripInlineMarkdown(headingMatch.Groups[2].Value.Trim());
                var props = new Dictionary<string, string>
                {
                    ["text"] = text,
                    ["bold"] = "true",
                    ["fontsize"] = level switch { 1 => "22", 2 => "18", 3 => "14", _ => "12" }
                };
                if (level <= 3) props["styleidref"] = (level + 1).ToString();
                if (align != null) props["align"] = align.ToUpperInvariant();
                Add("/section[1]", "paragraph", null, props);
                blockCount++;
                i++;
                continue;
            }

            // GFM Table: starts with |
            if (line.TrimStart().StartsWith('|'))
            {
                var tableLines = new List<string>();
                while (i < lines.Length && lines[i].TrimEnd('\r').TrimStart().StartsWith('|'))
                {
                    tableLines.Add(lines[i].TrimEnd('\r'));
                    i++;
                }
                blockCount += ImportMarkdownTable(tableLines);
                continue;
            }

            // Bold/italic paragraph
            {
                var text = StripInlineMarkdown(line.Trim());
                if (!string.IsNullOrEmpty(text))
                {
                    var props = new Dictionary<string, string> { ["text"] = text };
                    if (line.Trim().StartsWith("**") && line.Trim().EndsWith("**"))
                        props["bold"] = "true";
                    if (align != null) props["align"] = align.ToUpperInvariant();
                    Add("/section[1]", "paragraph", null, props);
                    blockCount++;
                }
                i++;
            }
        }

        return blockCount;
    }

    private int ImportMarkdownTable(List<string> tableLines)
    {
        // Parse table rows, skipping separator line (| --- | --- |)
        var rows = new List<string[]>();
        foreach (var line in tableLines)
        {
            var trimmed = line.Trim();
            // Skip separator rows
            if (Regex.IsMatch(trimmed, @"^\|[\s\-:|]+\|$")) continue;

            var cells = trimmed.Split('|', StringSplitOptions.None)
                .Skip(1) // leading empty from first |
                .ToArray();
            // Remove trailing empty from last |
            if (cells.Length > 0 && string.IsNullOrWhiteSpace(cells[^1]))
                cells = cells[..^1];
            cells = cells.Select(c => StripInlineMarkdown(c.Trim())).ToArray();
            if (cells.Length > 0) rows.Add(cells);
        }

        if (rows.Count == 0) return 0;

        int rowCount = rows.Count;
        int colCount = rows.Max(r => r.Length);

        // Create table
        Add("/section[1]", "table", null, new Dictionary<string, string>
        {
            ["rows"] = rowCount.ToString(),
            ["cols"] = colCount.ToString()
        });

        // Find the table we just added — it's the last tbl in the document
        var lastTbl = _doc.Sections.SelectMany(s => s.Tables).LastOrDefault();
        if (lastTbl == null) return 0;

        // Find path to this table
        var tblPath = BuildPath(lastTbl);

        // Fill cells
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < rows[r].Length; c++)
            {
                var cellText = rows[r][c];
                if (!string.IsNullOrEmpty(cellText))
                {
                    var cellPath = $"{tblPath}/tr[{r + 1}]/tc[{c + 1}]";
                    try { Set(cellPath, new Dictionary<string, string> { ["text"] = cellText }); }
                    catch { /* skip cells that don't resolve */ }
                }
            }
        }

        return 1;
    }

    private static string StripInlineMarkdown(string text)
    {
        // Bold+italic (*** and ___)
        text = Regex.Replace(text, @"\*{3}(.+?)\*{3}", "$1");
        text = Regex.Replace(text, @"_{3}(.+?)_{3}", "$1");
        // Bold (** and __)
        text = Regex.Replace(text, @"\*{2}(.+?)\*{2}", "$1");
        text = Regex.Replace(text, @"_{2}(.+?)_{2}", "$1");
        // Italic (* and _)
        text = Regex.Replace(text, @"\*(.+?)\*", "$1");
        text = Regex.Replace(text, @"(?<!\w)_(.+?)_(?!\w)", "$1");
        // Strikethrough
        text = Regex.Replace(text, @"~~(.+?)~~", "$1");
        // Inline code
        text = Regex.Replace(text, @"`(.+?)`", "$1");
        // Images: ![alt](url) → alt
        text = Regex.Replace(text, @"!\[(.+?)\]\(.+?\)", "$1");
        // Links: [text](url) → text
        text = Regex.Replace(text, @"\[(.+?)\]\(.+?\)", "$1");
        // HTML entities
        text = text.Replace("&amp;", "&").Replace("&lt;", "<")
                   .Replace("&gt;", ">").Replace("&quot;", "\"")
                   .Replace("&#39;", "'");
        // Escaped markdown chars
        text = Regex.Replace(text, @"\\([*_~`\[\]\\|#>!\-])", "$1");
        // Escaped pipe
        text = text.Replace("\\|", "|");
        return text.Trim();
    }
}
