// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Regenerate TOC field entries from document headings.
///
/// Mirrors the 4-phase pipeline:
/// 1. Heading enumeration  — walk body, find paragraphs at requested levels
/// 2. Bookmark management  — ensure each heading has a stable anchor
/// 3. Entry generation     — emit TOC1/TOC2/TOC3 paragraphs with hyperlink + PAGEREF
/// 4. Page-number filling  — done externally via HTML pagination
/// </summary>
internal static class WordTocBuilder
{
    /// <summary>Regenerate every TOC field in the document. The TOC entries
    /// emit "0" as the page number; the caller resolves the real page numbers
    /// via HTML pagination and rewrites them into the PAGEREF result runs.</summary>
    public static void RegenerateAllTocs(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return;

        var headings = EnumerateHeadings(doc, body);

        foreach (var (tocPara, spec) in FindTocFields(body))
        {
            // A TOC whose switches are not all understood keeps its existing entries.
            if (spec == null) continue;
            var selected = SelectHeadings(headings, spec);
            EnsureHeadingBookmarks(body, selected.Select(s => s.Heading).ToList());
            var entries = GenerateEntries(selected, spec);
            ReplaceTocFieldContent(body, tocPara, entries);
        }
    }

    // ==================== Phase 1: Heading enumeration ====================

    public sealed class HeadingInfo
    {
        public Paragraph Para { get; }
        public string Text { get; }
        /// <summary>Style name as the TOC \t switch spells it ("Heading 3"), or "".</summary>
        public string StyleName { get; }
        /// <summary>N for a paragraph in built-in style HeadingN (style ID), else 0.</summary>
        public int HeadingNumber { get; }
        /// <summary>Applied outline level 1-9 (direct or from style), else -1.</summary>
        public int OutlineLevel { get; }
        public string BookmarkName { get; set; } = "";
        public HeadingInfo(Paragraph p, string text, string styleName, int headingNumber, int outlineLevel)
        {
            Para = p; Text = text; StyleName = styleName; HeadingNumber = headingNumber; OutlineLevel = outlineLevel;
        }
    }

    static List<HeadingInfo> EnumerateHeadings(WordprocessingDocument doc, Body body)
    {
        var styleLevels = ResolveHeadingStyleLevels(doc);
        var styleNames = ResolveStyleNames(doc);
        var list = new List<HeadingInfo>();
        foreach (var p in body.Descendants<Paragraph>())
        {
            // Skip paragraphs inside text boxes / floating frames — only
            // body-level headings should drive TOC generation, matching Word.
            if (p.Ancestors<TextBoxContent>().Any()) continue;

            var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
            var outline = ResolveOutlineLevel(p, styleLevels);
            // A paragraph can still be listed by \t through its style name, so keep any styled paragraph.
            if (outline < 1 && styleId == "") continue;
            var text = ExtractHeadingText(p);
            if (string.IsNullOrWhiteSpace(text)) continue;
            var headingMatch = Regex.Match(styleId, @"^Heading([1-9])$");
            var headingNumber = headingMatch.Success ? int.Parse(headingMatch.Groups[1].Value) : 0;
            styleNames.TryGetValue(styleId, out var styleName);
            list.Add(new HeadingInfo(p, text, styleName ?? "", headingNumber, outline));
        }
        return list;
    }

    static Dictionary<string, string> ResolveStyleNames(WordprocessingDocument doc)
    {
        var map = new Dictionary<string, string>();
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return map;
        foreach (var s in styles.Elements<Style>())
        {
            var id = s.StyleId?.Value;
            var name = s.StyleName?.Val?.Value;
            if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(name)) map[id] = name;
        }
        return map;
    }

    static Dictionary<string, int> ResolveHeadingStyleLevels(WordprocessingDocument doc)
    {
        var map = new Dictionary<string, int>();
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return map;
        foreach (var s in styles.Elements<Style>())
        {
            var id = s.StyleId?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            var lvl = s.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (lvl.HasValue && lvl.Value >= 0 && lvl.Value <= 8)
                map[id] = lvl.Value + 1;
        }
        return map;
    }

    static int ResolveOutlineLevel(Paragraph p, Dictionary<string, int> styleLevels)
    {
        var direct = p.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct.HasValue && direct.Value >= 0 && direct.Value <= 8) return direct.Value + 1;
        var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrEmpty(styleId))
        {
            if (styleLevels.TryGetValue(styleId, out var sl)) return sl;
            // Fallback: legacy Heading1-9 style names without explicit outline level.
            var m = Regex.Match(styleId, @"^Heading([1-9])$");
            if (m.Success) return int.Parse(m.Groups[1].Value);
        }
        return -1;
    }

    static string ExtractHeadingText(Paragraph p)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var t in p.Descendants<Text>())
            sb.Append(t.Text);
        return sb.ToString().Trim();
    }

    // ==================== Phase 2: Bookmark management ====================

    static void EnsureHeadingBookmarks(Body body, List<HeadingInfo> headings)
    {
        // Reuse bookmarks named _Toc* if already wrapped around the heading;
        // otherwise generate _Toc{16-hex} stable per-heading.
        int maxId = body.Descendants<BookmarkStart>()
            .Select(b => int.TryParse(b.Id?.Value, out var n) ? n : 0)
            .DefaultIfEmpty(0).Max();

        foreach (var h in headings)
        {
            var existing = h.Para.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name?.Value?.StartsWith("_Toc", StringComparison.Ordinal) == true);
            if (existing != null)
            {
                h.BookmarkName = existing.Name!.Value!;
                continue;
            }
            var name = $"_Toc{Guid.NewGuid().ToString("N")[..8]}";
            var bookmarkId = (++maxId).ToString();
            // Insert bookmarkStart at paragraph head (after pPr if present), end at tail.
            var pPr = h.Para.GetFirstChild<ParagraphProperties>();
            var bs = new BookmarkStart { Id = bookmarkId, Name = name };
            var be = new BookmarkEnd { Id = bookmarkId };
            if (pPr != null) pPr.InsertAfterSelf(bs);
            else h.Para.PrependChild(bs);
            h.Para.AppendChild(be);
            h.BookmarkName = name;
        }
    }

    // ==================== Phase 3: Entry generation ====================

    /// <summary>
    /// The switches of one TOC field, as ECMA-376 defines them (the TOC field):
    /// HeadingMin/HeadingMax is the \o range (1-9 when absent or without a range); HeadingStyles is set by \o;
    /// OutlineLevels by \u; StyleLevels maps a style name from \t to its TOC level; NoPageMin/NoPageMax is the
    /// level range that \n omits page numbers for (an empty range when \n is absent).
    /// </summary>
    public sealed record TocSpec(int HeadingMin, int HeadingMax, bool HeadingStyles, bool OutlineLevels,
        IReadOnlyDictionary<string, int> StyleLevels, bool Hyperlinks, int NoPageMin, int NoPageMax);

    static List<(Paragraph TocPara, TocSpec? Spec)> FindTocFields(Body body)
    {
        var list = new List<(Paragraph, TocSpec?)>();
        foreach (var p in body.Elements<Paragraph>())
        {
            var instrText = p.Descendants<FieldCode>()
                .FirstOrDefault(fc => fc.Text?.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase) == true);
            if (instrText == null) continue;
            list.Add((p, ParseTocSwitches(instrText.Text!)));
        }
        return list;
    }

    // Switches this builder implements. Any other switch can change which entries appear (\b, \c, \f, \l, ...),
    // so a TOC that uses one is left as it is rather than rebuilt wrongly.
    static readonly HashSet<string> HandledTocSwitches = new() { "o", "u", "t", "h", "z", "n" };

    /// <summary>Returns null when the TOC cannot be rebuilt faithfully.</summary>
    static TocSpec? ParseTocSwitches(string instr)
    {
        foreach (Match sw in Regex.Matches(instr, @"\\([a-zA-Z])(?![a-zA-Z])"))
            if (!HandledTocSwitches.Contains(sw.Groups[1].Value)) return null;

        // \o "a-b": built-in heading styles HeadingN. With no range, all levels are used.
        var headingMin = 1; var headingMax = 9;
        var o = Regex.Match(instr, @"\\o(?![a-zA-Z])(?:\s+""([^""]*)"")?");
        var headingStyles = o.Success;
        if (o.Groups[1].Success && !TryParseLevelRange(o.Groups[1].Value, out headingMin, out headingMax)) return null;

        // \u: paragraphs with an applied outline level.
        var outlineLevels = Regex.IsMatch(instr, @"\\u(?![a-zA-Z])");

        // \t "Style name,level,Style name,level": custom styles by name. The separator is a comma.
        var styleLevels = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var t = Regex.Match(instr, @"\\t(?![a-zA-Z])(?:\s+""([^""]*)"")?");
        if (t.Success)
        {
            if (!t.Groups[1].Success) return null;
            var parts = t.Groups[1].Value.Split(',');
            if (parts.Length % 2 != 0) return null;
            for (var i = 0; i < parts.Length; i += 2)
            {
                var name = parts[i].Trim();
                if (name == "" || !int.TryParse(parts[i + 1].Trim(), out var lvl) || lvl < 1 || lvl > 9) return null;
                styleLevels[name] = lvl;
            }
        }

        // \n omits page numbers; with a range only for those levels. \z only hides them in Web layout.
        var noPageMin = 1; var noPageMax = 0;
        var n = Regex.Match(instr, @"\\n(?![a-zA-Z])(?:\s+""([^""]*)"")?");
        if (n.Success)
        {
            noPageMin = 1; noPageMax = 9;
            if (n.Groups[1].Success && !TryParseLevelRange(n.Groups[1].Value, out noPageMin, out noPageMax)) return null;
        }

        if (!headingStyles && !outlineLevels && styleLevels.Count == 0) return null;
        var hyperlinks = Regex.IsMatch(instr, @"\\h(?![a-zA-Z])");
        return new TocSpec(headingMin, headingMax, headingStyles, outlineLevels, styleLevels, hyperlinks, noPageMin, noPageMax);
    }

    static bool TryParseLevelRange(string text, out int min, out int max)
    {
        var m = Regex.Match(text, @"^\s*(\d)\s*-\s*(\d)\s*$");
        min = 0; max = 0;
        if (!m.Success) return false;
        min = int.Parse(m.Groups[1].Value); max = int.Parse(m.Groups[2].Value);
        return min >= 1 && min <= max && max <= 9;
    }

    sealed record SelectedHeading(HeadingInfo Heading, int TocLevel);

    static List<SelectedHeading> SelectHeadings(List<HeadingInfo> headings, TocSpec spec)
    {
        var selected = new List<SelectedHeading>();
        foreach (var h in headings)
        {
            var level = 0;
            if (spec.StyleLevels.TryGetValue(h.StyleName, out var styleLevel)) level = styleLevel;
            else if (spec.HeadingStyles && InLevelRange(h.HeadingNumber, spec)) level = h.HeadingNumber;
            // Taking the \o range as also limiting \u is an inference (not stated in ECMA-376).
            else if (spec.OutlineLevels && InLevelRange(h.OutlineLevel, spec)) level = h.OutlineLevel;
            if (level > 0) selected.Add(new SelectedHeading(h, level));
        }
        return selected;
    }

    static bool InLevelRange(int level, TocSpec spec) => level >= spec.HeadingMin && level <= spec.HeadingMax;

    static List<Paragraph> GenerateEntries(List<SelectedHeading> selected, TocSpec spec)
    {
        return selected.Select(s => BuildEntryParagraph(s.Heading, s.TocLevel, spec)).ToList();
    }

    static Paragraph BuildEntryParagraph(HeadingInfo h, int tocLevel, TocSpec spec)
    {
        var styleId = $"TOC{tocLevel}";
        var noPageNum = tocLevel >= spec.NoPageMin && tocLevel <= spec.NoPageMax;
        var pPr = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
        var p = new Paragraph(pPr);

        // Heading text run; wrapped in hyperlink if \h.
        var textRun = new Run(new Text(h.Text) { Space = SpaceProcessingModeValues.Preserve });
        var tabRun = new Run(new TabChar());

        OpenXmlElement entryHost = p;
        if (spec.Hyperlinks)
        {
            var hyper = new Hyperlink { Anchor = h.BookmarkName, History = OnOffValue.FromBoolean(true) };
            p.AppendChild(hyper);
            entryHost = hyper;
        }
        entryHost.AppendChild(textRun);

        if (!noPageNum)
        {
            entryHost.AppendChild(tabRun);
            // PAGEREF field: { PAGEREF _TocXXX \h }. Result run starts as "0";
            // the caller rewrites it to the real page number after pagination.
            entryHost.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            entryHost.AppendChild(new Run(new FieldCode($" PAGEREF {h.BookmarkName} \\h ")
            { Space = SpaceProcessingModeValues.Preserve }));
            entryHost.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            entryHost.AppendChild(new Run(new Text("0") { Space = SpaceProcessingModeValues.Preserve }));
            entryHost.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }
        return p;
    }

    static void ReplaceTocFieldContent(Body body, Paragraph tocFieldPara, List<Paragraph> entries)
    {
        // The TOC field's begin / instr / sep / end can span multiple
        // paragraphs, and the body between sep and end typically contains
        // nested PAGEREF sub-fields per entry. Depth-track to find the
        // matching outer end fldChar.

        var sepRun = tocFieldPara.Descendants<Run>()
            .FirstOrDefault(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate);
        if (sepRun == null) return;

        Paragraph? endPara = null;
        int depth = 1;
        bool sawSep = false;
        foreach (var r in body.Descendants<Run>())
        {
            if (!sawSep) { if (r == sepRun) sawSep = true; continue; }
            var fc = r.GetFirstChild<FieldChar>();
            if (fc?.FieldCharType?.Value == FieldCharValues.Begin) depth++;
            else if (fc?.FieldCharType?.Value == FieldCharValues.End)
            {
                depth--;
                if (depth == 0) { endPara = r.Ancestors<Paragraph>().FirstOrDefault(); break; }
            }
        }
        if (endPara == null) return;

        // 1) Remove everything in tocFieldPara strictly after sepRun's outermost
        //    ancestor inside the paragraph. (sep is usually a direct child Run,
        //    but be defensive about wrapping containers.)
        OpenXmlElement sepRoot = sepRun;
        while (sepRoot.Parent != null && sepRoot.Parent != tocFieldPara) sepRoot = sepRoot.Parent;
        var afterSep = sepRoot.NextSibling();
        while (afterSep != null) { var n = afterSep.NextSibling(); afterSep.Remove(); afterSep = n; }

        // 2) Remove paragraphs from after tocFieldPara up to and including
        //    endPara (we'll synthesize a fresh end run in its own paragraph).
        if (endPara != tocFieldPara)
        {
            var p = tocFieldPara.NextSibling<Paragraph>();
            while (p != null)
            {
                var n = p.NextSibling<Paragraph>();
                bool wasEnd = (p == endPara);
                p.Remove();
                if (wasEnd) break;
                p = n;
            }
        }

        // 3) Insert generated entry paragraphs.
        OpenXmlElement insertAfter = tocFieldPara;
        foreach (var entry in entries)
        {
            insertAfter.InsertAfterSelf(entry);
            insertAfter = entry;
        }

        // 4) Append a synthetic end-fldChar paragraph closing the outer field.
        var endParaNew = new Paragraph(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        insertAfter.InsertAfterSelf(endParaNew);
    }

}
