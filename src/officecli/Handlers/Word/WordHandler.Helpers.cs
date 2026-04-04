// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Private Helpers ====================

    /// <summary>
    /// Format twips as a human-readable cm string (e.g., "21cm").
    /// 1 inch = 1440 twips, 1 inch = 2.54 cm.
    /// </summary>
    private static string FormatTwipsToCm(uint twips)
    {
        var cm = twips * 2.54 / 1440.0;
        return $"{cm:0.##}cm";
    }

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    private static JustificationValues ParseJustification(string value) =>
        value.ToLowerInvariant() switch
        {
            "left" => JustificationValues.Left,
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "justify" or "both" => JustificationValues.Both,
            _ => throw new ArgumentException($"Invalid alignment value: '{value}'. Valid values: left, center, right, justify.")
        };

    /// <summary>
    /// Sanitize a hex color for Word OOXML (ST_HexColorRGB = exactly 6-char RGB).
    /// Strips # prefix, uppercases, and handles 8-char AARRGGBB by extracting RGB portion.
    /// </summary>
    private static string SanitizeHex(string value) =>
        ParseHelpers.SanitizeColorForOoxml(value).Rgb;

    /// <summary>
    /// Parse a highlight color name, throwing ArgumentException with valid options on failure.
    /// </summary>
    private static readonly HashSet<string> ValidHighlightColors = new(StringComparer.OrdinalIgnoreCase)
    {
        "yellow", "green", "cyan", "magenta", "blue", "red",
        "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
        "darkGray", "lightGray", "black", "white", "none"
    };

    private static HighlightColorValues ParseHighlightColor(string value)
    {
        if (!ValidHighlightColors.Contains(value))
            throw new ArgumentException(
                $"Invalid 'highlight' value '{value}'. Valid values: yellow, green, cyan, magenta, blue, red, " +
                $"darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, white, none.");
        return new HighlightColorValues(value);
    }

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
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(string.Concat(run.Elements<Text>().Select(t => t.Text)));
            else if (child is Hyperlink hyperlink)
                sb.Append(string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text)));
        }
        return sb.ToString();
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
        return para.Descendants<Run>()
            .Where(r => r.GetFirstChild<CommentReference>() == null)
            .ToList();
    }

    /// <summary>
    /// Find the paragraph path where a CommentRangeStart with the given ID is anchored.
    /// Returns "/body/p[N]" or null if not found.
    /// </summary>
    private string? FindCommentAnchorPath(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;

        var paragraphs = body.Elements<Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var hasRange = paragraphs[i].Descendants<CommentRangeStart>()
                .Any(rs => rs.Id?.Value == commentId);
            if (hasRange) return $"/body/p[{i + 1}]";
        }
        return null;
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
            foreach (var sectPr in mainPart.Document?.Body?.Descendants<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
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

            // Build footer text by processing paragraphs, resolving field codes
            var footerLines = new List<string>();
            foreach (var para in footer.Descendants<Paragraph>())
            {
                var sb = new System.Text.StringBuilder();
                bool inField = false;
                bool pastSeparator = false;

                foreach (var run in para.Elements<Run>())
                {
                    var fldChar = run.GetFirstChild<FieldChar>();
                    if (fldChar != null)
                    {
                        if (fldChar.FieldCharType! == FieldCharValues.Begin)
                        {
                            inField = true;
                            pastSeparator = false;
                        }
                        else if (fldChar.FieldCharType! == FieldCharValues.Separate)
                        {
                            pastSeparator = true;
                        }
                        else if (fldChar.FieldCharType! == FieldCharValues.End)
                        {
                            inField = false;
                            pastSeparator = false;
                        }
                        continue;
                    }

                    var fieldCode = run.GetFirstChild<FieldCode>();
                    if (fieldCode != null)
                    {
                        // Extract field type from instruction (e.g., " PAGE " -> "PAGE")
                        var instr = fieldCode.Text?.Trim() ?? "";
                        var fieldType = instr.Split(' ', System.StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? instr;
                        sb.Append($"[{fieldType.ToUpperInvariant()}]");
                        continue;
                    }

                    // Skip result runs inside a field (they contain stale/literal values)
                    if (inField && pastSeparator)
                        continue;

                    var text = run.GetFirstChild<Text>();
                    if (text != null)
                        sb.Append(text.Text);
                }

                var line = sb.ToString().Trim();
                if (!string.IsNullOrEmpty(line))
                    footerLines.Add(line);
            }

            if (footerLines.Count > 0)
                results.Add(string.Join(" ", footerLines));
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
    private static void ApplyRunFormatting(OpenXmlCompositeElement props, string key, string? value)
    {
        if (value is null) return;
        switch (key.ToLowerInvariant())
        {
            case "size":
                var existingFs = props.GetFirstChild<FontSize>();
                if (existingFs != null) existingFs.Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString();
                else props.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString() });
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
                props.AppendChild(new Highlight { Val = ParseHighlightColor(value) });
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
            var count = ReplaceInParagraph(para, find, replace);
            if (count > 0)
                para.TextId = GenerateParaId();
            totalCount += count;
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
            {
                // Insert after SectionType, or after last headerReference/footerReference
                var sectionType = sectPr.GetFirstChild<SectionType>();
                if (sectionType != null)
                    sectionType.InsertAfterSelf(cols);
                else
                {
                    OpenXmlElement? lastRef = null;
                    foreach (var child in sectPr.ChildElements)
                    {
                        if (child is HeaderReference || child is FooterReference)
                            lastRef = child;
                    }
                    if (lastRef != null)
                        lastRef.InsertAfterSelf(cols);
                    else
                        sectPr.PrependChild(cols);
                }
            }
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
        // Insert after SectionType if present, then after FooterReference/HeaderReference,
        // otherwise prepend. OOXML schema order: headerReference*, footerReference*, ..., sectType, pgSz, pgMar
        var sectionType = sectPr.GetFirstChild<SectionType>();
        if (sectionType != null)
        {
            sectionType.InsertAfterSelf(ps);
        }
        else
        {
            // Find the last HeaderReference or FooterReference to insert after
            OpenXmlElement? lastRef = null;
            foreach (var child in sectPr.ChildElements)
            {
                if (child is HeaderReference || child is FooterReference)
                    lastRef = child;
            }
            if (lastRef != null)
                lastRef.InsertAfterSelf(ps);
            else
                sectPr.PrependChild(ps);
        }
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
        // Insert after PageSize if present, after SectionType, after last headerRef/footerRef, or prepend
        var pageSize = sectPr.GetFirstChild<PageSize>();
        if (pageSize != null)
            pageSize.InsertAfterSelf(pm);
        else
        {
            var sectionType = sectPr.GetFirstChild<SectionType>();
            if (sectionType != null)
                sectionType.InsertAfterSelf(pm);
            else
            {
                OpenXmlElement? lastRef = null;
                foreach (var child in sectPr.ChildElements)
                {
                    if (child is HeaderReference || child is FooterReference)
                        lastRef = child;
                }
                if (lastRef != null)
                    lastRef.InsertAfterSelf(pm);
                else
                    sectPr.PrependChild(pm);
            }
        }
        return pm;
    }

    // ==================== w14 Text Effects ====================

    private const string W14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";

    /// <summary>
    /// Remove an existing w14 element from RunProperties by local name.
    /// </summary>
    private static void RemoveW14Element(RunProperties rPr, string localName)
    {
        var existing = rPr.ChildElements
            .Where(e => e.LocalName == localName && e.NamespaceUri == W14Ns)
            .ToList();
        foreach (var e in existing) e.Remove();
    }

    /// <summary>
    /// Split a w14 effect value string by ';' (preferred) or '-' (legacy fallback).
    /// ';' is unambiguous; '-' is only used as fallback when no ';' is present.
    /// </summary>
    private static string[] SplitEffectValue(string value) =>
        value.Contains(';') ? value.Split(';') : value.Split('-');

    /// <summary>
    /// Build w14:textOutline XML.
    /// Format: "WIDTH;COLOR" (e.g. "0.5pt;FF0000"), "WIDTH" (defaults to black), or "none"
    /// Width in pt, internally stored in EMU (1pt = 12700 EMU).
    /// Legacy: "WIDTH-COLOR" also accepted.
    /// </summary>
    internal static string BuildW14TextOutline(string value)
    {
        var parts = SplitEffectValue(value);
        var widthPt = ParseHelpers.SafeParseDouble(parts[0].Replace("pt", ""), "textOutline width");
        var widthEmu = (long)(widthPt * 12700);
        var color = parts.Length > 1 ? ParseHelpers.SanitizeColorForOoxml(parts[1]).Rgb : "000000";

        return $@"<w14:textOutline xmlns:w14=""{W14Ns}"" w14:w=""{widthEmu}"" w14:cap=""flat"" w14:cmpd=""sng"" w14:algn=""ctr""><w14:solidFill><w14:srgbClr w14:val=""{color}""/></w14:solidFill><w14:prstDash w14:val=""solid""/></w14:textOutline>";
    }

    /// <summary>
    /// Build w14:textFill XML.
    /// Format: "C1;C2[;ANGLE]" for linear gradient, "radial:C1;C2" for radial, or single color for solid.
    /// Legacy: '-' separator also accepted.
    /// </summary>
    internal static string BuildW14TextFill(string value)
    {
        if (value.StartsWith("radial:", StringComparison.OrdinalIgnoreCase))
        {
            var radParts = SplitEffectValue(value[7..]);
            var (c1, _) = ParseHelpers.SanitizeColorForOoxml(radParts[0]);
            var c2 = radParts.Length > 1 ? ParseHelpers.SanitizeColorForOoxml(radParts[1]).Rgb : c1;
            return $@"<w14:textFill xmlns:w14=""{W14Ns}""><w14:gradFill><w14:gsLst><w14:gs w14:pos=""0""><w14:srgbClr w14:val=""{c1}""/></w14:gs><w14:gs w14:pos=""100000""><w14:srgbClr w14:val=""{c2}""/></w14:gs></w14:gsLst><w14:path w14:path=""circle""><w14:fillToRect w14:l=""50000"" w14:t=""50000"" w14:r=""50000"" w14:b=""50000""/></w14:path></w14:gradFill></w14:textFill>";
        }

        var parts = SplitEffectValue(value);
        if (parts.Length == 1)
        {
            // Solid fill
            var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
            return $@"<w14:textFill xmlns:w14=""{W14Ns}""><w14:solidFill><w14:srgbClr w14:val=""{rgb}""/></w14:solidFill></w14:textFill>";
        }

        // Linear gradient: C1;C2[;angle]
        var (gc1, _a1) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
        var (gc2, _a2) = ParseHelpers.SanitizeColorForOoxml(parts[1]);
        var angle = parts.Length > 2 ? ParseHelpers.SafeParseInt(parts[2], "textFill angle") * 60000 : 0;
        return $@"<w14:textFill xmlns:w14=""{W14Ns}""><w14:gradFill><w14:gsLst><w14:gs w14:pos=""0""><w14:srgbClr w14:val=""{gc1}""/></w14:gs><w14:gs w14:pos=""100000""><w14:srgbClr w14:val=""{gc2}""/></w14:gs></w14:gsLst><w14:lin w14:ang=""{angle}"" w14:scaled=""1""/></w14:gradFill></w14:textFill>";
    }

    /// <summary>
    /// Build w14:shadow XML.
    /// Format: "COLOR[;BLUR[;ANGLE[;DIST[;OPACITY]]]]"
    /// Defaults: blur=4pt, angle=45°, dist=3pt, opacity=40%
    /// Legacy: '-' separator also accepted.
    /// </summary>
    internal static string BuildW14Shadow(string value)
    {
        var parts = SplitEffectValue(value);
        var (color, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
        var blurPt = parts.Length > 1 ? ParseHelpers.SafeParseDouble(parts[1], "shadow blur") : 4.0;
        var angleDeg = parts.Length > 2 ? ParseHelpers.SafeParseDouble(parts[2], "shadow angle") : 45.0;
        var distPt = parts.Length > 3 ? ParseHelpers.SafeParseDouble(parts[3], "shadow distance") : 3.0;
        var opacity = parts.Length > 4 ? ParseHelpers.SafeParseDouble(parts[4], "shadow opacity") : 40.0;

        var blurEmu = (long)(blurPt * 12700);
        var distEmu = (long)(distPt * 12700);
        var angleOoxml = (int)(angleDeg * 60000);
        var alphaVal = (int)(opacity * 1000);

        return $@"<w14:shadow xmlns:w14=""{W14Ns}"" w14:blurRad=""{blurEmu}"" w14:dist=""{distEmu}"" w14:dir=""{angleOoxml}"" w14:sx=""100000"" w14:sy=""100000"" w14:kx=""0"" w14:ky=""0"" w14:algn=""tl""><w14:srgbClr w14:val=""{color}""><w14:alpha w14:val=""{alphaVal}""/></w14:srgbClr></w14:shadow>";
    }

    /// <summary>
    /// Build w14:glow XML.
    /// Format: "COLOR[;RADIUS[;OPACITY]]"
    /// Defaults: radius=8pt, opacity=75%
    /// Legacy: '-' separator also accepted.
    /// </summary>
    internal static string BuildW14Glow(string value)
    {
        var parts = SplitEffectValue(value);
        var (color, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
        var radiusPt = parts.Length > 1 ? ParseHelpers.SafeParseDouble(parts[1], "glow radius") : 8.0;
        var opacity = parts.Length > 2 ? ParseHelpers.SafeParseDouble(parts[2], "glow opacity") : 75.0;

        var radiusEmu = (long)(radiusPt * 12700);
        var alphaVal = (int)(opacity * 1000);

        return $@"<w14:glow xmlns:w14=""{W14Ns}"" w14:rad=""{radiusEmu}""><w14:srgbClr w14:val=""{color}""><w14:alpha w14:val=""{alphaVal}""/></w14:srgbClr></w14:glow>";
    }

    /// <summary>
    /// Build w14:reflection XML.
    /// Values: "tight"/"small", "half"/"true", "full"
    /// </summary>
    internal static string BuildW14Reflection(string value)
    {
        var endPos = value.ToLowerInvariant() switch
        {
            "tight" or "small" => 55000,
            "true" or "half" => 90000,
            "full" => 100000,
            _ => int.TryParse(value, out var pct) ? (int)Math.Min((long)pct * 1000, 100000) : 90000
        };

        return $@"<w14:reflection xmlns:w14=""{W14Ns}"" w14:blurRad=""6350"" w14:stA=""52000"" w14:stPos=""0"" w14:endA=""300"" w14:endPos=""{endPos}"" w14:dist=""0"" w14:dir=""5400000"" w14:fadeDir=""5400000"" w14:sx=""100000"" w14:sy=""-100000"" w14:kx=""0"" w14:ky=""0"" w14:algn=""bl""/>";
    }

    /// <summary>
    /// Apply a w14 text effect to a run's RunProperties.
    /// Handles set and remove logic.
    /// </summary>
    internal static void ApplyW14TextEffect(Run run, string effectName, string value, Func<string, string> builder)
    {
        var rPr = EnsureRunProperties(run);
        RemoveW14Element(rPr, effectName);

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
            return;

        var xml = builder(value);
        var element = new OpenXmlUnknownElement("w14", "tmp", W14Ns);
        element.InnerXml = xml;
        var child = element.FirstChild;
        if (child != null)
        {
            child.Remove();
            rPr.AppendChild(child);
        }
    }

    /// <summary>
    /// Read w14 text effect values from RunProperties.
    /// Returns a dictionary of effect names to their parsed values.
    /// </summary>
    internal static void ReadW14TextEffects(RunProperties? rPr, DocumentNode node)
    {
        if (rPr == null) return;

        foreach (var child in rPr.ChildElements)
        {
            if (child.NamespaceUri != W14Ns) continue;

            switch (child.LocalName)
            {
                case "textOutline":
                {
                    var wAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
                    var widthEmu = long.TryParse(wAttr.Value, out var w) ? w : 0;
                    var widthPt = widthEmu / 12700.0;
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value) : "";
                    node.Format["textOutline"] = string.IsNullOrEmpty(color) ? $"{widthPt:0.##}pt" : $"{widthPt:0.##}pt;{color}";
                    break;
                }
                case "textFill":
                {
                    var innerXml = child.InnerXml;
                    if (innerXml.Contains("gradFill"))
                    {
                        var colors = new List<string>();
                        foreach (System.Text.RegularExpressions.Match m in
                            System.Text.RegularExpressions.Regex.Matches(innerXml, @"val=""([0-9A-Fa-f]{6})"""))
                            colors.Add(m.Groups[1].Value);

                        // Add # prefix to gradient colors
                        for (int ci = 0; ci < colors.Count; ci++)
                            colors[ci] = ParseHelpers.FormatHexColor(colors[ci]);

                        var isRadial = innerXml.Contains("<w14:path");
                        if (isRadial && colors.Count >= 2)
                            node.Format["textFill"] = $"radial:{colors[0]};{colors[1]}";
                        else if (colors.Count >= 2)
                        {
                            var angleMatch = System.Text.RegularExpressions.Regex.Match(innerXml, @"ang=""(\d+)""");
                            var angle = angleMatch.Success ? int.Parse(angleMatch.Groups[1].Value) / 60000.0 : 0.0;
                            var angleStr = angle % 1 == 0 ? $"{(int)angle}" : $"{angle:0.##}";
                            node.Format["textFill"] = $"{colors[0]};{colors[1]};{angleStr}";
                        }
                        else if (colors.Count == 1)
                            node.Format["textFill"] = colors[0];
                    }
                    else if (innerXml.Contains("solidFill"))
                    {
                        var colorMatch = System.Text.RegularExpressions.Regex.Match(
                            innerXml, @"val=""([0-9A-Fa-f]{6})""");
                        if (colorMatch.Success)
                            node.Format["textFill"] = ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value);
                    }
                    break;
                }
                case "shadow":
                {
                    var attrs = child.GetAttributes().ToDictionary(a => a.LocalName, a => a.Value);
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value) : "#000000";
                    var blurEmu = attrs.TryGetValue("blurRad", out var br) && long.TryParse(br, out var blurVal) ? blurVal : 0;
                    var blurPt = blurEmu / 12700.0;
                    var dirVal = attrs.TryGetValue("dir", out var dir) && long.TryParse(dir, out var dirLong) ? dirLong : 0;
                    var angleDeg = dirVal / 60000.0;
                    var distEmu = attrs.TryGetValue("dist", out var dist) && long.TryParse(dist, out var distLong) ? distLong : 0;
                    var distPt = distEmu / 12700.0;
                    // Read alpha (opacity) from inner srgbClr child
                    var alphaMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"alpha[^>]*val=""(\d+)""");
                    var opacity = alphaMatch.Success && double.TryParse(alphaMatch.Groups[1].Value, out var alphaVal) ? alphaVal / 1000.0 : 100.0;
                    node.Format["w14shadow"] = $"{color};{blurPt:0.##};{angleDeg:0.##};{distPt:0.##};{opacity:0.##}";
                    break;
                }
                case "glow":
                {
                    var radAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "rad");
                    var radiusEmu = long.TryParse(radAttr.Value, out var r) ? r : 0;
                    var radiusPt = radiusEmu / 12700.0;
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value) : "#000000";
                    // Read alpha (opacity) from inner srgbClr child
                    var alphaMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"alpha[^>]*val=""(\d+)""");
                    var opacity = alphaMatch.Success && double.TryParse(alphaMatch.Groups[1].Value, out var av) ? av / 1000.0 : 100.0;
                    node.Format["w14glow"] = $"{color};{radiusPt:0.##};{opacity:0.##}";
                    break;
                }
                case "reflection":
                {
                    var endPosAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "endPos");
                    var endPos = int.TryParse(endPosAttr.Value, out var ep) ? ep : 90000;
                    node.Format["w14reflection"] = endPos switch
                    {
                        <= 55000 => "tight",
                        <= 90000 => "half",
                        _ => "full"
                    };
                    break;
                }
            }
        }
    }

    // ==================== Extended Chart Helpers ====================

    private const string WordChartExUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    private const string WordChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>
    /// Count all charts (both standard ChartPart and ExtendedChartPart) in the document.
    /// </summary>
    private static int CountWordCharts(MainDocumentPart mainPart)
    {
        return mainPart.ChartParts.Count() + mainPart.ExtendedChartParts.Count();
    }

    /// <summary>
    /// Represents a chart part in Word that could be either a standard ChartPart or an ExtendedChartPart.
    /// </summary>
    private class WordChartInfo
    {
        public ChartPart? StandardPart { get; set; }
        public ExtendedChartPart? ExtendedPart { get; set; }
        public DW.DocProperties? DocProperties { get; set; }
        public bool IsExtended => ExtendedPart != null;
    }

    /// <summary>
    /// Get all chart parts (standard + extended) in document order by walking Drawing/Inline elements.
    /// </summary>
    private List<WordChartInfo> GetAllWordCharts()
    {
        var result = new List<WordChartInfo>();
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return result;

        foreach (var inline in mainPart.Document.Body.Descendants<DW.Inline>())
        {
            var graphicData = inline.Descendants<A.GraphicData>().FirstOrDefault();
            if (graphicData == null) continue;

            var docProps = inline.Descendants<DW.DocProperties>().FirstOrDefault();

            if (graphicData.Uri == WordChartUri)
            {
                // Standard chart
                var chartRef = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
                if (chartRef?.Id?.Value == null) continue;
                try
                {
                    var chartPart = (ChartPart)mainPart.GetPartById(chartRef.Id.Value);
                    result.Add(new WordChartInfo { StandardPart = chartPart, DocProperties = docProps });
                }
                catch { /* skip invalid references */ }
            }
            else if (graphicData.Uri == WordChartExUri)
            {
                // Extended chart (funnel, treemap, etc.)
                var relId = GetWordExtendedChartRelId(inline);
                if (relId == null) continue;
                try
                {
                    var extPart = (ExtendedChartPart)mainPart.GetPartById(relId);
                    result.Add(new WordChartInfo { ExtendedPart = extPart, DocProperties = docProps });
                }
                catch { /* skip invalid references */ }
            }
        }

        return result;
    }

    /// <summary>
    /// Get the relationship ID from an extended chart inline Drawing element.
    /// </summary>
    private static string? GetWordExtendedChartRelId(DW.Inline inline)
    {
        var gd = inline.Descendants<A.GraphicData>().FirstOrDefault(g => g.Uri == WordChartExUri);
        if (gd == null) return null;
        var typed = gd.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId>().FirstOrDefault();
        if (typed?.Id?.Value != null) return typed.Id.Value;
        foreach (var child in gd.ChildElements)
        {
            var rId = child.GetAttributes().FirstOrDefault(a =>
                a.LocalName == "id" && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (rId.Value != null) return rId.Value;
        }
        return null;
    }

    /// <summary>
    /// Get current document protection mode and enforcement status.
    /// </summary>
    private (string mode, bool enforced) GetDocumentProtection()
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var docProtection = settings?.GetFirstChild<DocumentProtection>();
        if (docProtection == null)
            return ("none", false);

        var mode = docProtection.Edit?.InnerText switch
        {
            "readOnly" => "readOnly",
            "comments" => "comments",
            "trackedChanges" => "trackedChanges",
            "forms" => "forms",
            _ => "none"
        };
        var enforced = docProtection.Enforcement?.Value == true
            || (docProtection.Enforcement?.Value == null && docProtection.Edit != null);
        return (mode, enforced);
    }

    /// <summary>
    /// Check if an SDT element is editable based on its lock attribute and the current document protection.
    /// </summary>
    private bool IsSdtEditable(SdtProperties? sdtProps)
    {
        var (mode, enforced) = GetDocumentProtection();

        // No protection or not enforced → all SDTs are editable
        if (!enforced || mode == "none")
            return true;

        // readOnly protection → SDTs are not editable (unless in permRange, P2)
        if (mode == "readOnly")
            return false;

        // forms protection → SDTs are editable unless content-locked
        if (mode == "forms")
        {
            var lockEl = sdtProps?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            var lockVal = lockEl?.Val?.Value;
            return lockVal != LockingValues.ContentLocked && lockVal != LockingValues.SdtContentLocked;
        }

        // comments/trackedChanges → not typically editable
        return false;
    }

    /// <summary>
    /// Generate a unique 8-character uppercase hex ID for w14:paraId / w14:textId.
    /// OOXML spec requires value &lt; 0x80000000 (MaxExclusive).
    /// </summary>
    private static string GenerateParaId()
    {
        return Random.Shared.Next(0, int.MaxValue).ToString("X8");
    }

    /// <summary>
    /// Assign paraId and textId to a paragraph if not already set.
    /// </summary>
    private static void AssignParaId(Paragraph para)
    {
        if (string.IsNullOrEmpty(para.ParagraphId?.Value))
            para.ParagraphId = GenerateParaId();
        if (string.IsNullOrEmpty(para.TextId?.Value))
            para.TextId = GenerateParaId();
    }

    /// <summary>
    /// Ensure all paragraphs in the document have w14:paraId and w14:textId.
    /// Called on document open.
    /// </summary>
    private void EnsureAllParaIds()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Collect all paragraphs from body + headers + footers
        var allParagraphs = mainPart.Document.Body.Descendants<Paragraph>().AsEnumerable();
        foreach (var headerPart in mainPart.HeaderParts)
            if (headerPart.Header != null)
                allParagraphs = allParagraphs.Concat(headerPart.Header.Descendants<Paragraph>());
        foreach (var footerPart in mainPart.FooterParts)
            if (footerPart.Footer != null)
                allParagraphs = allParagraphs.Concat(footerPart.Footer.Descendants<Paragraph>());

        var paragraphs = allParagraphs.ToList();

        // Collect existing IDs first to avoid collisions
        foreach (var para in paragraphs)
        {
            if (!string.IsNullOrEmpty(para.ParagraphId?.Value))
                usedIds.Add(para.ParagraphId.Value);
            if (!string.IsNullOrEmpty(para.TextId?.Value))
                usedIds.Add(para.TextId.Value);
        }

        // Assign IDs to paragraphs that don't have them
        foreach (var para in paragraphs)
        {
            if (string.IsNullOrEmpty(para.ParagraphId?.Value))
            {
                string newId;
                do { newId = GenerateParaId(); } while (!usedIds.Add(newId));
                para.ParagraphId = newId;
            }
            if (string.IsNullOrEmpty(para.TextId?.Value))
            {
                string newId;
                do { newId = GenerateParaId(); } while (!usedIds.Add(newId));
                para.TextId = newId;
            }
        }
    }

    // ==================== DocPr IDs (pictures, charts) ====================

    /// <summary>
    /// Ensure all DocProperties in the document have unique IDs.
    /// Called on document open.
    /// </summary>
    private void EnsureDocPropIds()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        var allDocProps = mainPart.Document.Body.Descendants<DW.DocProperties>().ToList();

        foreach (var headerPart in mainPart.HeaderParts)
            if (headerPart.Header != null)
                allDocProps.AddRange(headerPart.Header.Descendants<DW.DocProperties>());
        foreach (var footerPart in mainPart.FooterParts)
            if (footerPart.Footer != null)
                allDocProps.AddRange(footerPart.Footer.Descendants<DW.DocProperties>());

        var usedIds = new HashSet<uint>();
        var duplicates = new List<DW.DocProperties>();

        foreach (var dp in allDocProps)
        {
            if (dp.Id?.HasValue == true && !usedIds.Add(dp.Id.Value))
                duplicates.Add(dp);
            else if (dp.Id?.HasValue != true)
                duplicates.Add(dp);
        }

        foreach (var dp in duplicates)
        {
            uint newId = 1;
            while (!usedIds.Add(newId)) newId++;
            dp.Id = newId;
        }
    }
}
