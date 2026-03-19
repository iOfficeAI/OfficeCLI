// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Private Helpers ====================

    private static bool IsTruthy(string value) =>
        ParseHelpers.IsTruthy(value);

    /// <summary>
    /// Sanitize a hex color for Word OOXML (ST_HexColorRGB = exactly 6-char RGB).
    /// Strips # prefix, uppercases, and handles 8-char AARRGGBB by extracting RGB portion.
    /// </summary>
    private static string SanitizeHex(string value) =>
        ParseHelpers.SanitizeColorForOoxml(value).Rgb;

    /// <summary>
    /// Warn if a value that should be a shading pattern name looks like a hex color instead.
    /// </summary>
    private static void WarnIfShadingOrderWrong(string patternSegment)
    {
        var trimmed = patternSegment.TrimStart('#');
        if (trimmed.Length >= 6 && trimmed.All(char.IsAsciiHexDigit))
            Console.Error.WriteLine($"Warning: '{patternSegment}' looks like a color, but is in the pattern position. "
                + "Shading format: FILL (single value) or PATTERN;FILL[;COLOR] e.g. clear;FF0000");
    }

    /// <summary>
    /// Append a child element to parent, but if parent is Body, insert before
    /// the final SectionProperties to maintain valid OOXML structure.
    /// </summary>
    private static void AppendToParent(OpenXmlElement parent, OpenXmlElement child)
    {
        if (parent is Body body)
        {
            var lastSectPr = body.GetFirstChild<SectionProperties>();
            if (lastSectPr != null)
            {
                body.InsertBefore(child, lastSectPr);
                return;
            }
        }
        parent.AppendChild(child);
    }

    private static double ParseFontSize(string value) =>
        ParseHelpers.ParseFontSize(value);

    private static string GetParagraphText(Paragraph para)
    {
        return string.Concat(para.Descendants<Text>().Select(t => t.Text));
    }

    /// <summary>
    /// Get paragraph text including inline math rendered as readable Unicode.
    /// </summary>
    private static string GetParagraphTextWithMath(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(GetRunText(run));
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
                sb.Append(FormulaParser.ToReadableText(child));
            else if (child is Hyperlink hyperlink)
                sb.Append(string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text)));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find math elements in a paragraph using both type and localName matching.
    /// </summary>
    private static List<OpenXmlElement> FindMathElements(Paragraph para)
    {
        return para.ChildElements
            .Where(e => e.LocalName == "oMath" || e is M.OfficeMath)
            .ToList();
    }

    /// <summary>
    /// Get all body-level elements, flattening SdtContent containers.
    /// This ensures paragraphs and tables inside w:sdt are not missed.
    /// </summary>
    private static IEnumerable<OpenXmlElement> GetBodyElements(Body body)
    {
        foreach (var element in body.ChildElements)
        {
            if (element is SdtBlock sdt)
            {
                var content = sdt.SdtContentBlock;
                if (content != null)
                {
                    foreach (var child in content.ChildElements)
                        yield return child;
                }
            }
            else
            {
                yield return element;
            }
        }
    }

    /// <summary>
    /// Checks if an element is a structural document element worth displaying
    /// (not inline markers like bookmarkStart, bookmarkEnd, proofErr, etc.)
    /// </summary>
    private static bool IsStructuralElement(OpenXmlElement element)
    {
        var name = element.LocalName;
        return name == "sectPr" || name == "altChunk" || name == "customXml";
    }

    /// <summary>
    /// Get all Run elements in a paragraph, including those nested inside
    /// Hyperlink and SdtContent containers.
    /// </summary>
    private static List<Run> GetAllRuns(Paragraph para)
    {
        return para.Descendants<Run>().ToList();
    }

    private static string GetRunText(Run run)
    {
        return string.Concat(run.Elements<Text>().Select(t => t.Text));
    }

    private string GetStyleName(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return "Normal";

        // Try to resolve display name from styles part
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles != null)
        {
            var style = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style?.StyleName?.Val?.Value != null)
                return style.StyleName.Val.Value;
        }

        return styleId;
    }

    private static string? GetRunFont(Run run)
    {
        var fonts = run.RunProperties?.RunFonts;
        return fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value ?? fonts?.EastAsia?.Value;
    }

    private static string? GetRunFontSize(Run run)
    {
        var size = run.RunProperties?.FontSize?.Val?.Value;
        if (size == null) return null;
        return $"{int.Parse(size) / 2.0:0.##}pt"; // stored as half-points
    }

    private string GetRunFormatDescription(Run run, Paragraph? para = null)
    {
        var parts = new List<string>();

        RunProperties? rProps;
        if (para != null)
        {
            rProps = ResolveEffectiveRunProperties(run, para);
        }
        else
        {
            rProps = run.RunProperties;
        }
        if (rProps == null) return "(default)";

        var font = GetFontFromProperties(rProps);
        if (font != null) parts.Add(font);

        var size = GetSizeFromProperties(rProps);
        if (size != null) parts.Add(size);

        if (rProps.Bold != null) parts.Add("bold");
        if (rProps.Italic != null) parts.Add("italic");
        if (rProps.Underline != null) parts.Add("underline");
        if (rProps.Strike != null) parts.Add("strikethrough");

        return parts.Count > 0 ? string.Join(" ", parts) : "(default)";
    }

    private static int GetHeadingLevel(string styleName)
    {
        // Heading 1, Heading 2, heading1, 标题 1, etc.
        foreach (var ch in styleName)
        {
            if (char.IsDigit(ch))
                return ch - '0';
        }
        if (styleName == "Title") return 0;
        if (styleName == "Subtitle") return 1;
        return 1;
    }

    private static bool IsNormalStyle(string styleName)
    {
        return styleName is "Normal" or "正文" or "Body Text" or "Body" or "a"
            || styleName.StartsWith("Normal");
    }

    private string? FindWatermark()
    {
        var headerParts = _doc.MainDocumentPart?.HeaderParts;
        if (headerParts == null) return null;

        foreach (var headerPart in headerParts)
        {
            var header = headerPart.Header;
            if (header == null) continue;

            // Search for VML shapes with watermark
            foreach (var pict in header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>())
            {
                foreach (var shape in pict.Descendants<Vml.Shape>())
                {
                    var id = shape.GetAttribute("id", "");
                    if (id.Value?.Contains("WaterMark", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        var textPath = shape.Descendants<Vml.TextPath>().FirstOrDefault();
                        return textPath?.String?.Value ?? "(image watermark)";
                    }
                }
            }

            // Also check for DrawingML watermarks
            foreach (var drawing in header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
            {
                // Simple detection: check if it looks like a watermark by inline/anchor properties
                var docProps = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
                if (docProps?.Name?.Value?.Contains("WaterMark", StringComparison.OrdinalIgnoreCase) == true)
                {
                    return "(image watermark)";
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Remove all header parts that contain watermark SDT elements.
    /// </summary>
    private void RemoveWatermarkHeaders()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return;

        var toRemove = new List<HeaderPart>();
        foreach (var hp in mainPart.HeaderParts)
        {
            if (hp.Header == null) continue;
            // Check for watermark: SDT with docPartGallery="Watermarks" or VML shape with "WaterMark" in id
            var hasSdt = hp.Header.Descendants<SdtProperties>()
                .Any(sp => sp.Descendants<DocPartGallery>().Any(g =>
                    g.Val?.Value?.Equals("Watermarks", StringComparison.OrdinalIgnoreCase) == true));
            if (hasSdt)
            {
                toRemove.Add(hp);
                continue;
            }
            foreach (var pict in hp.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>())
            {
                var hasWm = pict.InnerXml.Contains("WaterMark", StringComparison.OrdinalIgnoreCase);
                if (hasWm) { toRemove.Add(hp); break; }
            }
        }

        foreach (var hp in toRemove)
        {
            // Remove header references from section properties
            var relId = mainPart.GetIdOfPart(hp);
            foreach (var sectPr in mainPart.Document?.Body?.Elements<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
            {
                var refs = sectPr.Elements<HeaderReference>().Where(r => r.Id?.Value == relId).ToList();
                foreach (var r in refs) r.Remove();
            }
            mainPart.DeletePart(hp);
        }
    }

    private List<string> GetHeaderTexts()
    {
        var results = new List<string>();
        var headerParts = _doc.MainDocumentPart?.HeaderParts;
        if (headerParts == null) return results;

        foreach (var headerPart in headerParts)
        {
            var header = headerPart.Header;
            if (header == null) continue;
            var text = string.Concat(header.Descendants<Text>().Select(t => t.Text)).Trim();
            if (!string.IsNullOrEmpty(text))
                results.Add(text);
        }

        return results;
    }

    private List<string> GetFooterTexts()
    {
        var results = new List<string>();
        var footerParts = _doc.MainDocumentPart?.FooterParts;
        if (footerParts == null) return results;

        foreach (var footerPart in footerParts)
        {
            var footer = footerPart.Footer;
            if (footer == null) continue;
            var text = string.Concat(footer.Descendants<Text>().Select(t => t.Text)).Trim();
            if (!string.IsNullOrEmpty(text))
                results.Add(text);
            else
            {
                // Check for page numbers
                var fldChars = footer.Descendants<FieldCode>().Any();
                if (fldChars)
                    results.Add("(page number)");
            }
        }

        return results;
    }

    private static bool HasMixedPunctuation(string text)
    {
        var chinesePunct = "\uff0c\u3002\uff01\uff1f\u3001\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011";
        bool hasChinese = text.Any(c => chinesePunct.Contains(c));
        bool hasEnglish = text.Any(c => ",.!?;:\"'()[]".Contains(c));
        bool hasChineseChars = text.Any(c => c >= 0x4E00 && c <= 0x9FFF);
        return hasChinese && hasEnglish && hasChineseChars;
    }

    private static RunProperties EnsureRunProperties(Run run)
    {
        return run.RunProperties ?? run.PrependChild(new RunProperties());
    }

    /// <summary>
    /// Apply a run-level formatting property to either RunProperties or ParagraphMarkRunProperties.
    /// </summary>
    private static void ApplyRunFormatting(OpenXmlCompositeElement props, string key, string value)
    {
        switch (key.ToLowerInvariant())
        {
            case "size":
                var existingFs = props.GetFirstChild<FontSize>();
                if (existingFs != null) existingFs.Val = ((int)(ParseFontSize(value) * 2)).ToString();
                else props.AppendChild(new FontSize { Val = ((int)(ParseFontSize(value) * 2)).ToString() });
                break;
            case "font":
                var existingRf = props.GetFirstChild<RunFonts>();
                if (existingRf != null) { existingRf.Ascii = value; existingRf.HighAnsi = value; existingRf.EastAsia = value; }
                else props.AppendChild(new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value });
                break;
            case "bold":
                props.RemoveAllChildren<Bold>();
                if (IsTruthy(value)) props.AppendChild(new Bold());
                break;
            case "italic":
                props.RemoveAllChildren<Italic>();
                if (IsTruthy(value)) props.AppendChild(new Italic());
                break;
            case "color":
                props.RemoveAllChildren<Color>();
                props.AppendChild(new Color { Val = SanitizeHex(value) });
                break;
            case "highlight":
                props.RemoveAllChildren<Highlight>();
                props.AppendChild(new Highlight { Val = new HighlightColorValues(value) });
                break;
            case "underline":
                props.RemoveAllChildren<Underline>();
                var ulMapped = value.ToLowerInvariant() switch { "true" => "single", "false" or "none" => "none", _ => value };
                props.AppendChild(new Underline { Val = new UnderlineValues(ulMapped) });
                break;
            case "strike":
                props.RemoveAllChildren<Strike>();
                if (IsTruthy(value)) props.AppendChild(new Strike());
                break;
        }
    }

    private static string GetBookmarkText(BookmarkStart bkStart)
    {
        var bkId = bkStart.Id?.Value;
        if (bkId == null) return "";

        var sb = new System.Text.StringBuilder();
        var sibling = bkStart.NextSibling();
        while (sibling != null)
        {
            if (sibling is BookmarkEnd bkEnd && bkEnd.Id?.Value == bkId)
                break;
            if (sibling is Run run)
                sb.Append(string.Concat(run.Descendants<Text>().Select(t => t.Text)));
            sibling = sibling.NextSibling();
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find and replace text across the document. Returns the number of replacements made.
    /// Handles text split across multiple runs within a paragraph.
    /// </summary>
    private int FindAndReplace(string find, string replace, string scope = "all")
    {
        if (string.IsNullOrEmpty(find)) return 0;
        int totalCount = 0;

        // Collect all paragraphs to process based on scope
        var paragraphs = new List<Paragraph>();
        var mainPart = _doc.MainDocumentPart;

        if (scope is "all" or "body" or "")
        {
            if (mainPart?.Document?.Body != null)
                paragraphs.AddRange(mainPart.Document.Body.Descendants<Paragraph>());
        }
        if (scope is "all" or "headers")
        {
            foreach (var hp in mainPart?.HeaderParts ?? Enumerable.Empty<DocumentFormat.OpenXml.Packaging.HeaderPart>())
                if (hp.Header != null) paragraphs.AddRange(hp.Header.Descendants<Paragraph>());
        }
        if (scope is "all" or "footers")
        {
            foreach (var fp in mainPart?.FooterParts ?? Enumerable.Empty<DocumentFormat.OpenXml.Packaging.FooterPart>())
                if (fp.Footer != null) paragraphs.AddRange(fp.Footer.Descendants<Paragraph>());
        }

        foreach (var para in paragraphs)
        {
            totalCount += ReplaceInParagraph(para, find, replace);
        }

        return totalCount;
    }

    /// <summary>
    /// Replace text within a paragraph, handling text split across multiple runs.
    /// </summary>
    private static int ReplaceInParagraph(Paragraph para, string find, string replace)
    {
        var runs = para.Elements<Run>().ToList();
        if (runs.Count == 0) return 0;

        // Build concatenated text with run boundaries
        var runTexts = new List<(Run Run, Text TextElement, int Start, int End)>();
        int pos = 0;
        foreach (var run in runs)
        {
            foreach (var text in run.Elements<Text>())
            {
                var len = text.Text?.Length ?? 0;
                if (len > 0)
                    runTexts.Add((run, text, pos, pos + len));
                pos += len;
            }
        }

        if (runTexts.Count == 0) return 0;
        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));

        // Find all occurrences
        var indices = new List<int>();
        int idx = 0;
        while ((idx = fullText.IndexOf(find, idx, StringComparison.Ordinal)) >= 0)
        {
            indices.Add(idx);
            idx += find.Length;
        }

        if (indices.Count == 0) return 0;

        // Process replacements from end to start to preserve positions
        for (int i = indices.Count - 1; i >= 0; i--)
        {
            var matchStart = indices[i];
            var matchEnd = matchStart + find.Length;

            // Find which run-texts are affected
            bool first = true;
            foreach (var rt in runTexts)
            {
                if (rt.End <= matchStart || rt.Start >= matchEnd)
                    continue; // not affected

                var textStr = rt.TextElement.Text ?? "";
                var localStart = Math.Max(0, matchStart - rt.Start);
                var localEnd = Math.Min(textStr.Length, matchEnd - rt.Start);

                if (first)
                {
                    // First affected run: replace the matched portion with replacement text
                    rt.TextElement.Text = textStr[..localStart] + replace + textStr[localEnd..];
                    rt.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    first = false;
                }
                else
                {
                    // Subsequent runs: just remove the matched portion
                    rt.TextElement.Text = textStr[..Math.Max(0, matchStart - rt.Start)] + textStr[localEnd..];
                    rt.TextElement.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }

        return indices.Count;
    }

    /// <summary>
    /// Ensure Columns exists in SectionProperties in correct schema order.
    /// Schema order: ..., PageMargin, ..., Columns, ...
    /// </summary>
    private static Columns EnsureColumns(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<Columns>();
        if (existing != null) return existing;

        var cols = new Columns();
        var pm = sectPr.GetFirstChild<PageMargin>();
        if (pm != null)
            pm.InsertAfterSelf(cols);
        else
        {
            var pgSz = sectPr.GetFirstChild<PageSize>();
            if (pgSz != null)
                pgSz.InsertAfterSelf(cols);
            else
                sectPr.PrependChild(cols);
        }
        return cols;
    }

    /// <summary>
    /// Ensure PageSize exists in SectionProperties in correct schema order.
    /// Schema order: SectionType, PageSize, PageMargin, ...
    /// </summary>
    private static PageSize EnsureSectPrPageSize(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageSize>();
        if (existing != null) return existing;

        var ps = new PageSize();
        // Insert after SectionType if present, otherwise prepend
        var sectionType = sectPr.GetFirstChild<SectionType>();
        if (sectionType != null)
            sectionType.InsertAfterSelf(ps);
        else
            sectPr.PrependChild(ps);
        return ps;
    }

    /// <summary>
    /// Ensure PageMargin exists in SectionProperties in correct schema order.
    /// Schema order: SectionType, PageSize, PageMargin, ...
    /// </summary>
    private static PageMargin EnsureSectPrPageMargin(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageMargin>();
        if (existing != null) return existing;

        var pm = new PageMargin();
        // Insert after PageSize if present, after SectionType, or prepend
        var pageSize = sectPr.GetFirstChild<PageSize>();
        if (pageSize != null)
            pageSize.InsertAfterSelf(pm);
        else
        {
            var sectionType = sectPr.GetFirstChild<SectionType>();
            if (sectionType != null)
                sectionType.InsertAfterSelf(pm);
            else
                sectPr.PrependChild(pm);
        }
        return pm;
    }
}
