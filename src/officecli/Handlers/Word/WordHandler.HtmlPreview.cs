// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>Rendering context passed through the HTML generation pipeline.</summary>
    private class HtmlRenderContext
    {
        public List<int> FootnoteRefs { get; } = new();
        public List<int> EndnoteRefs { get; } = new();
        public PageLayout? CachedPageLayout { get; set; }
    }

    /// <summary>Current render context — set during ViewAsHtml, used by all render methods.</summary>
    private HtmlRenderContext _ctx = null!;

    /// <summary>
    /// Generate a self-contained HTML file that previews the Word document
    /// with formatting, tables, images, and lists.
    /// </summary>
    public string ViewAsHtml()
    {
        _ctx = new HtmlRenderContext();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "<html><body><p>(empty document)</p></body></html>";

        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        var pgLayout = GetPageLayout();
        var docDef = ReadDocDefaults();
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateWordCss(pgLayout, docDef));
        sb.AppendLine("</style>");
        // KaTeX for math rendering
        sb.AppendLine("<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css\">");
        sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js\"></script>");
        sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/contrib/auto-render.min.js\"></script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        // Page container
        var maxW = $"max-width:{pgLayout.WidthCm:0.##}cm";

        sb.AppendLine($"<div class=\"page\" style=\"{maxW}\">");

        // Render header
        RenderHeaderFooterHtml(sb, isHeader: true);

        // Render body elements (footnote/endnote refs tracked in _ctx)
        RenderBodyHtml(sb, body);

        // Render footnotes section
        RenderFootnotesHtml(sb);

        // Render footer
        RenderHeaderFooterHtml(sb, isHeader: false);

        // Render endnotes section (at document end)
        RenderEndnotesHtml(sb);

        sb.AppendLine("</div>"); // page

        // KaTeX auto-render script
        sb.AppendLine("<script>");
        sb.AppendLine("document.addEventListener('DOMContentLoaded',function(){");
        sb.AppendLine("  if(typeof renderMathInElement!=='undefined'){");
        sb.AppendLine("    renderMathInElement(document.body,{delimiters:[");
        sb.AppendLine("      {left:'$$',right:'$$',display:true}");
        sb.AppendLine("    ],throwOnError:false});");
        sb.AppendLine("  }");
        sb.AppendLine("});");
        sb.AppendLine("</script>");

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }

    // ==================== Page Layout + Doc Defaults from OOXML ====================

    private record PageLayout(double WidthCm, double HeightCm,
        double MarginTopCm, double MarginBottomCm, double MarginLeftCm, double MarginRightCm);

    private PageLayout GetPageLayout()
    {
        if (_ctx?.CachedPageLayout != null) return _ctx.CachedPageLayout;
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        var pgSz = sectPr?.GetFirstChild<PageSize>();
        var pgMar = sectPr?.GetFirstChild<PageMargin>();
        const double c = 2.54 / 1440.0; // twips → cm
        var result = new PageLayout(
            (pgSz?.Width?.Value ?? 11906) * c,
            (pgSz?.Height?.Value ?? 16838) * c,
            (double)(pgMar?.Top?.Value ?? 1440) * c,
            (double)(pgMar?.Bottom?.Value ?? 1440) * c,
            (pgMar?.Left?.Value ?? 1440u) * c,
            (pgMar?.Right?.Value ?? 1440u) * c);
        if (_ctx != null) _ctx.CachedPageLayout = result;
        return result;
    }

    private record DocDef(string Font, double SizePt, double LineHeight, string Color);

    private DocDef ReadDocDefaults()
    {
        var defs = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
        var rPr = defs?.RunPropertiesDefault?.RunPropertiesBaseStyle;

        // Font: docDefaults rFonts → theme minor font → fallback
        var fonts = rPr?.RunFonts;
        var font = NonEmpty(fonts?.EastAsia?.Value) ?? NonEmpty(fonts?.Ascii?.Value) ?? NonEmpty(fonts?.HighAnsi?.Value);
        if (font == null)
        {
            var minor = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont;
            font = NonEmpty(minor?.EastAsianFont?.Typeface) ?? NonEmpty(minor?.LatinFont?.Typeface);
        }

        // Size: half-points → pt
        double sizePt = 10.5;
        if (rPr?.FontSize?.Val?.Value is string sz && int.TryParse(sz, out var hp))
            sizePt = hp / 2.0;

        // Line spacing from pPrDefault
        double lineH = 1.15;
        var sp = defs?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines;
        if (sp?.Line?.Value is string lv && int.TryParse(lv, out var lvi) && sp.LineRule?.InnerText is "auto" or null)
            lineH = lvi / 240.0;

        // Default text color: docDefaults → theme dk1
        var color = "#000000";
        var cv = rPr?.Color?.Val?.Value;
        if (cv != null && cv != "auto") color = $"#{cv}";
        else if (GetThemeColors().TryGetValue("dk1", out var dk1)) color = $"#{dk1}";

        return new DocDef(font ?? "Calibri", sizePt, lineH, color);
    }

    private static string? NonEmpty(string? s) => string.IsNullOrEmpty(s) ? null : s;

    /// <summary>Resolve shading fill color: direct hex or themeFill + themeFillTint/Shade.</summary>
    private string? ResolveShadingFill(Shading? shading)
    {
        if (shading == null) return null;
        var fill = shading.Fill?.Value;
        if (fill != null && fill != "auto") return $"#{fill}";
        // Check themeFill
        var themeFill = shading.GetAttributes().FirstOrDefault(a => a.LocalName == "themeFill").Value;
        if (themeFill != null)
        {
            var tc = GetThemeColors();
            if (tc.TryGetValue(themeFill, out var hex))
            {
                var tint = shading.GetAttributes().FirstOrDefault(a => a.LocalName == "themeFillTint").Value;
                var shade = shading.GetAttributes().FirstOrDefault(a => a.LocalName == "themeFillShade").Value;
                return ApplyTintShade(hex, tint, shade);
            }
        }
        return null;
    }

    /// <summary>Check if dimensions are ≥90% of the page size (full-page background element).</summary>
    private bool IsFullPageSize(long widthEmu, long heightEmu)
    {
        var pg = GetPageLayout();
        var pgW = (long)(pg.WidthCm / 2.54 * 914400);
        var pgH = (long)(pg.HeightCm / 2.54 * 914400);
        return widthEmu > pgW * 0.9 && heightEmu > pgH * 0.9;
    }

    /// <summary>Find embed attribute from a blip element anywhere in the element tree.</summary>
    private static string? FindEmbedInDescendants(OpenXmlElement el)
    {
        // Try SDK Descendants first
        foreach (var child in el.Descendants())
        {
            if (child.LocalName == "blip")
            {
                var embed = child.GetAttributes().FirstOrDefault(a => a.LocalName == "embed").Value;
                if (embed != null) return embed;
            }
        }
        // Fallback: parse outer XML for embed attribute (handles unknown elements)
        var xml = el.OuterXml;
        var match = Regex.Match(xml, @"r:embed=""(rId\d+)""");
        return match.Success ? match.Groups[1].Value : null;
    }

    // ==================== Header / Footer ====================

    private void RenderHeaderFooterHtml(StringBuilder sb, bool isHeader)
    {
        var cssClass = isHeader ? "doc-header" : "doc-footer";

        if (isHeader)
        {
            var headerParts = _doc.MainDocumentPart?.HeaderParts;
            if (headerParts == null) return;
            foreach (var hp in headerParts)
            {
                var paragraphs = hp.Header?.Elements<Paragraph>().ToList();
                if (paragraphs == null || paragraphs.Count == 0) continue;
                if (paragraphs.All(p => string.IsNullOrWhiteSpace(GetParagraphText(p)))) continue;
                sb.AppendLine($"<div class=\"{cssClass}\">");
                foreach (var para in paragraphs) RenderParagraphHtml(sb, para);
                sb.AppendLine("</div>");
                break;
            }
        }
        else
        {
            var footerParts = _doc.MainDocumentPart?.FooterParts;
            if (footerParts == null) return;
            foreach (var fp in footerParts)
            {
                var paragraphs = fp.Footer?.Elements<Paragraph>().ToList();
                if (paragraphs == null || paragraphs.Count == 0) continue;
                if (paragraphs.All(p => string.IsNullOrWhiteSpace(GetParagraphText(p)))) continue;
                sb.AppendLine($"<div class=\"{cssClass}\">");
                foreach (var para in paragraphs) RenderParagraphHtml(sb, para);
                sb.AppendLine("</div>");
                break;
            }
        }
    }

    // ==================== Body Rendering ====================

    private void RenderBodyHtml(StringBuilder sb, Body body)
    {
        var elements = GetBodyElements(body).ToList();
        // Track list state for proper HTML list rendering
        string? currentListType = null; // "bullet" or "ordered"
        int currentListLevel = 0;
        var listStack = new Stack<string>(); // track nested list tags

        foreach (var element in elements)
        {
            if (element is Paragraph para)
            {
                // Check for display equation
                var oMathPara = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathPara != null)
                {
                    CloseAllLists(sb, listStack, ref currentListType);
                    var latex = FormulaParser.ToLatex(oMathPara);
                    sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
                    continue;
                }

                // Check if this is a list item
                var listStyle = GetParagraphListStyle(para);
                if (listStyle != null)
                {
                    var ilvl = para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
                    var tag = listStyle == "bullet" ? "ul" : "ol";

                    // Adjust nesting (close deeper levels: </ol></li> or </ul></li>)
                    while (listStack.Count > ilvl + 1)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        sb.AppendLine("</li>");
                    }
                    while (listStack.Count < ilvl + 1)
                    {
                        if (listStack.Count > 0)
                            sb.Append("<li>"); // wrap nested list inside <li>
                        sb.AppendLine($"<{tag}>");
                        listStack.Push(tag);
                    }
                    // If same level but different list type, swap
                    if (listStack.Count > 0 && listStack.Peek() != tag)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        sb.AppendLine($"<{tag}>");
                        listStack.Push(tag);
                    }

                    currentListType = listStyle;
                    currentListLevel = ilvl;
                    sb.Append("<li");
                    var paraStyle = GetParagraphInlineCss(para, isListItem: true);
                    if (!string.IsNullOrEmpty(paraStyle))
                        sb.Append($" style=\"{paraStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine("</li>");
                    continue;
                }

                // Not a list — close any open lists
                CloseAllLists(sb, listStack, ref currentListType);

                // Check for heading
                var styleName = GetStyleName(para);
                var headingLevel = 0;
                if (styleName.Contains("Heading") || styleName.Contains("标题")
                    || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase))
                {
                    headingLevel = GetHeadingLevel(styleName);
                    if (headingLevel < 1) headingLevel = 1;
                    if (headingLevel > 6) headingLevel = 6;
                }
                else if (styleName == "Title")
                    headingLevel = 1;
                else if (styleName == "Subtitle")
                    headingLevel = 2;

                if (headingLevel > 0)
                {
                    sb.Append($"<h{headingLevel}");
                    var hStyle = GetParagraphInlineCss(para);
                    if (!string.IsNullOrEmpty(hStyle))
                        sb.Append($" style=\"{hStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine($"</h{headingLevel}>");
                }
                else
                {
                    // Normal paragraph
                    var text = GetParagraphText(para);
                    var runs = GetAllRuns(para);
                    var mathElements = FindMathElements(para);

                    // Empty paragraph = spacing break
                    if (runs.Count == 0 && mathElements.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        sb.AppendLine("<p class=\"empty\">&nbsp;</p>");
                        continue;
                    }

                    // Inline equation only
                    if (mathElements.Count > 0 && runs.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                        sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
                        continue;
                    }

                    sb.Append("<p");
                    var pStyle = GetParagraphInlineCss(para);
                    if (!string.IsNullOrEmpty(pStyle))
                        sb.Append($" style=\"{pStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine("</p>");
                }
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                CloseAllLists(sb, listStack, ref currentListType);
                var latex = FormulaParser.ToLatex(element);
                sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
            }
            else if (element is Table table)
            {
                CloseAllLists(sb, listStack, ref currentListType);
                RenderTableHtml(sb, table);
            }
            else if (element is SectionProperties)
            {
                // Skip — section properties are not visual content
            }
        }

        CloseAllLists(sb, listStack, ref currentListType);
    }

    private static void CloseAllLists(StringBuilder sb, Stack<string> listStack, ref string? currentListType)
    {
        bool first = true;
        while (listStack.Count > 0)
        {
            sb.AppendLine($"</{listStack.Pop()}>");
            // Close wrapper <li> for nested levels (not for the outermost list)
            if (!first || listStack.Count > 0)
            {
                if (listStack.Count > 0)
                    sb.AppendLine("</li>");
            }
            first = false;
        }
        currentListType = null;
    }
}
