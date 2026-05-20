// Copyright 2025 OfficeCLI (officecli.ai)
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
        EnsureHeadingBookmarks(body, headings);

        foreach (var (tocPara, spec) in FindTocFields(body))
        {
            var entries = GenerateEntries(headings, spec);
            ReplaceTocFieldContent(body, tocPara, entries);
        }
    }

    // ==================== Phase 1: Heading enumeration ====================

    public sealed class HeadingInfo
    {
        public Paragraph Para { get; }
        public int Level { get; }
        public string Text { get; }
        public string BookmarkName { get; set; } = "";
        public HeadingInfo(Paragraph p, int level, string text) { Para = p; Level = level; Text = text; }
    }

    static List<HeadingInfo> EnumerateHeadings(WordprocessingDocument doc, Body body)
    {
        var styleLevels = ResolveHeadingStyleLevels(doc);
        var list = new List<HeadingInfo>();
        foreach (var p in body.Descendants<Paragraph>())
        {
            // Skip paragraphs inside text boxes / floating frames — only
            // body-level headings should drive TOC generation, matching Word.
            if (p.Ancestors<TextBoxContent>().Any()) continue;

            var level = ResolveOutlineLevel(p, styleLevels);
            if (level < 1 || level > 9) continue;
            var text = ExtractHeadingText(p);
            if (string.IsNullOrWhiteSpace(text)) continue;
            list.Add(new HeadingInfo(p, level, text));
        }
        return list;
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

    public sealed record TocSpec(int MinLevel, int MaxLevel, bool Hyperlinks, bool NoPageNum);

    static List<(Paragraph TocPara, TocSpec Spec)> FindTocFields(Body body)
    {
        var list = new List<(Paragraph, TocSpec)>();
        foreach (var p in body.Elements<Paragraph>())
        {
            var instrText = p.Descendants<FieldCode>()
                .FirstOrDefault(fc => fc.Text?.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase) == true);
            if (instrText == null) continue;
            list.Add((p, ParseTocSwitches(instrText.Text!)));
        }
        return list;
    }

    static TocSpec ParseTocSwitches(string instr)
    {
        var min = 1; var max = 3;
        var m = Regex.Match(instr, @"\\o\s+""\s*(\d+)\s*-\s*(\d+)\s*""");
        if (m.Success) { min = int.Parse(m.Groups[1].Value); max = int.Parse(m.Groups[2].Value); }
        var hyperlinks = Regex.IsMatch(instr, @"\\h\b");
        var noPageNum = Regex.IsMatch(instr, @"\\z\b") || Regex.IsMatch(instr, @"\\n\b");
        return new TocSpec(min, max, hyperlinks, noPageNum);
    }

    static List<Paragraph> GenerateEntries(List<HeadingInfo> headings, TocSpec spec)
    {
        var paras = new List<Paragraph>();
        foreach (var h in headings)
        {
            if (h.Level < spec.MinLevel || h.Level > spec.MaxLevel) continue;
            paras.Add(BuildEntryParagraph(h, spec));
        }
        return paras;
    }

    static Paragraph BuildEntryParagraph(HeadingInfo h, TocSpec spec)
    {
        var styleId = $"TOC{h.Level}";
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

        if (!spec.NoPageNum)
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
