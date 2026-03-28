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
    /// <summary>
    /// Generate a self-contained HTML file that previews the Word document
    /// with formatting, tables, images, and lists.
    /// </summary>
    public string ViewAsHtml()
    {
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

        // Render body elements
        RenderBodyHtml(sb, body);

        // Render footer
        RenderHeaderFooterHtml(sb, isHeader: false);

        sb.AppendLine("</div>"); // page

        // KaTeX auto-render script
        sb.AppendLine("<script>");
        sb.AppendLine("document.addEventListener('DOMContentLoaded',function(){");
        sb.AppendLine("  if(typeof renderMathInElement!=='undefined'){");
        sb.AppendLine("    renderMathInElement(document.body,{delimiters:[");
        sb.AppendLine("      {left:'$$',right:'$$',display:true},");
        sb.AppendLine("      {left:'$',right:'$',display:false}");
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
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        var pgSz = sectPr?.GetFirstChild<PageSize>();
        var pgMar = sectPr?.GetFirstChild<PageMargin>();
        const double c = 2.54 / 1440.0; // twips → cm
        return new PageLayout(
            (pgSz?.Width?.Value ?? 11906) * c,
            (pgSz?.Height?.Value ?? 16838) * c,
            (double)(pgMar?.Top?.Value ?? 1440) * c,
            (double)(pgMar?.Bottom?.Value ?? 1440) * c,
            (pgMar?.Left?.Value ?? 1440u) * c,
            (pgMar?.Right?.Value ?? 1440u) * c);
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

                    // Adjust nesting
                    while (listStack.Count > ilvl + 1)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                    }
                    while (listStack.Count < ilvl + 1)
                    {
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
        while (listStack.Count > 0)
            sb.AppendLine($"</{listStack.Pop()}>");
        currentListType = null;
    }

    // ==================== Paragraph Content ====================

    private void RenderParagraphHtml(StringBuilder sb, Paragraph para)
    {
        sb.Append("<p");
        var pStyle = GetParagraphInlineCss(para);
        if (!string.IsNullOrEmpty(pStyle))
            sb.Append($" style=\"{pStyle}\"");
        sb.Append(">");
        RenderParagraphContentHtml(sb, para);
        sb.AppendLine("</p>");
    }

    private void RenderParagraphContentHtml(StringBuilder sb, Paragraph para)
    {
        // Collect standalone images that precede a text box group (they overlay the group in Word)
        bool hasTextBoxGroup = HasTextBoxContent(para);
        var preGroupImages = hasTextBoxGroup ? new List<Drawing>() : null;
        bool textBoxSeen = false;

        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
            {
                // Find drawing (direct child or inside mc:AlternateContent Choice)
                // SDK's Descendants<Drawing>() naturally skips mc:Fallback (VML w:pict)
                var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();

                if (drawing != null && HasGroupOrShape(drawing))
                {
                    bool hasTextBox = HasTextBox(drawing);
                    if (hasTextBox && preGroupImages != null)
                    {
                        // Render group with preceding images overlaid into text box
                        RenderDrawingWithOverlaidImages(sb, drawing, preGroupImages);
                        preGroupImages.Clear();
                        textBoxSeen = true;
                    }
                    else
                    {
                        RenderDrawingHtml(sb, drawing);
                    }
                    continue;
                }

                // Collect standalone images before text box group for overlay
                if (hasTextBoxGroup && !textBoxSeen && drawing != null)
                {
                    preGroupImages!.Add(drawing);
                    continue;
                }

                RenderRunHtml(sb, run, para);
            }
            else if (child.LocalName is "ins" or "moveTo")
            {
                // Tracked insertions — render their child runs
                foreach (var insRun in child.Elements<Run>())
                    RenderRunHtml(sb, insRun, para);
            }
            else if (child.LocalName is "del" or "moveFrom")
            {
                // Tracked deletions — skip (deleted content should not be displayed)
            }
            else if (child is Hyperlink hyperlink)
            {
                var relId = hyperlink.Id?.Value;
                string? url = null;
                if (relId != null)
                {
                    try
                    {
                        url = _doc.MainDocumentPart?.HyperlinkRelationships
                            .FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                    }
                    catch { }
                    if (url == null)
                    {
                        try
                        {
                            url = _doc.MainDocumentPart?.ExternalRelationships
                                .FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                        }
                        catch { }
                    }
                }

                if (url != null)
                    sb.Append($"<a href=\"{HtmlEncode(url)}\" target=\"_blank\">");

                foreach (var hRun in hyperlink.Elements<Run>())
                    RenderRunHtml(sb, hRun, para);

                if (url != null)
                    sb.Append("</a>");
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                var latex = FormulaParser.ToLatex(child);
                sb.Append($"${HtmlEncode(latex)}$");
            }
        }
    }

    // ==================== Run Rendering ====================

    private void RenderRunHtml(StringBuilder sb, Run run, Paragraph para)
    {
        // Check for drawing (direct or inside mc:AlternateContent)
        var drawing = run.GetFirstChild<Drawing>()
            ?? run.Descendants<Drawing>().FirstOrDefault();
        if (drawing != null)
        {
            RenderDrawingHtml(sb, drawing);
            return;
        }

        // Check for break
        var br = run.GetFirstChild<Break>();
        if (br != null)
        {
            if (br.Type?.Value == BreakValues.Page)
                sb.Append("<hr class=\"page-break\">");
            else
                sb.Append("<br>");
        }

        // Check for tab
        var tab = run.GetFirstChild<TabChar>();

        var text = GetRunText(run);
        if (string.IsNullOrEmpty(text) && tab == null) return;

        var rProps = ResolveEffectiveRunProperties(run, para);
        var style = GetRunInlineCss(rProps);

        var needsSpan = !string.IsNullOrEmpty(style);
        if (needsSpan)
            sb.Append($"<span style=\"{style}\">");

        if (tab != null)
            sb.Append("&emsp;");

        sb.Append(HtmlEncode(text));

        if (needsSpan)
            sb.Append("</span>");
    }

    // ==================== Drawing with Overlaid Images ====================

    private void RenderDrawingWithOverlaidImages(StringBuilder sb, Drawing groupDrawing, List<Drawing> overlaidImages)
    {
        if (overlaidImages.Count == 0)
        {
            RenderDrawingHtml(sb, groupDrawing, null);
            return;
        }

        RenderDrawingHtml(sb, groupDrawing, overlaidImages);
    }

    // ==================== Drawing Rendering (images, groups, shapes) ====================

    /// <summary>Check if a paragraph contains drawings with actual text box content (txbxContent).</summary>
    private static bool HasTextBoxContent(Paragraph para)
    {
        foreach (var run in para.Elements<Run>())
        {
            var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();
            if (drawing != null && HasTextBox(drawing))
                return true;
        }
        return false;
    }

    /// <summary>Check if a drawing contains groups or shapes (for rendering).</summary>
    private static bool HasGroupOrShape(Drawing drawing)
    {
        return drawing.Descendants().Any(e => e.LocalName == "wgp" || e.LocalName == "wsp");
    }

    /// <summary>Check if a drawing contains actual text box content with text (not empty decorative shapes).</summary>
    private static bool HasTextBox(Drawing drawing)
    {
        foreach (var txbx in drawing.Descendants().Where(e => e.LocalName == "txbxContent"))
        {
            // Check if any paragraph inside has actual text
            if (txbx.Descendants<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text)))
                return true;
        }
        return false;
    }

    private void RenderDrawingHtml(StringBuilder sb, Drawing drawing, List<Drawing>? floatImages = null)
    {
        // Check for groups/shapes first (text boxes, decorated shapes)
        var group = drawing.Descendants().FirstOrDefault(e => e.LocalName == "wgp");
        if (group != null)
        {
            // Get overall extent from wp:inline or wp:anchor
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            long groupWidthEmu = extent?.Cx?.Value ?? 0;
            long groupHeightEmu = extent?.Cy?.Value ?? 0;

            if (groupWidthEmu > 0 && groupHeightEmu > 0)
            {
                RenderGroupHtml(sb, group, groupWidthEmu, groupHeightEmu, floatImages);
                return;
            }
        }

        // Check for standalone shape (wsp without group)
        var shape = drawing.Descendants().FirstOrDefault(e => e.LocalName == "wsp");
        if (shape != null)
        {
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            long shapeWidth = extent?.Cx?.Value ?? 0;
            long shapeHeight = extent?.Cy?.Value ?? 0;
            if (shapeWidth > 0 && shapeHeight > 0)
            {
                // Full-page shapes → render as background layer
                if (IsFullPageSize(shapeWidth, shapeHeight))
                {
                    var fillCss = ResolveShapeFillCss(shape.Elements().FirstOrDefault(e => e.LocalName == "spPr"));
                    if (!string.IsNullOrEmpty(fillCss))
                        sb.Append($"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;{fillCss}\"></div>");
                    return;
                }
                RenderShapeHtml(sb, shape, 0, 0, shapeWidth, shapeHeight, shapeWidth, shapeHeight, floatImages);
                return;
            }
        }

        // Fall back to image rendering
        RenderImageHtml(sb, drawing);
    }

    private void RenderImageHtml(StringBuilder sb, Drawing drawing)
    {
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value == null) return;

        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return;

        try
        {
            var imagePart = mainPart.GetPartById(blip.Embed.Value) as ImagePart;
            if (imagePart == null) return;

            var contentType = imagePart.ContentType;
            using var stream = imagePart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());

            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault()
                ?? drawing.Descendants<A.Extents>().FirstOrDefault() as OpenXmlElement;
            long imgCxEmu = 0, imgCyEmu = 0;
            if (extent is DW.Extent dwExt) { imgCxEmu = dwExt.Cx?.Value ?? 0; imgCyEmu = dwExt.Cy?.Value ?? 0; }
            else if (extent is A.Extents aExt) { imgCxEmu = aExt.Cx?.Value ?? 0; imgCyEmu = aExt.Cy?.Value ?? 0; }

            var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            var alt = docProps?.Description?.Value ?? docProps?.Name?.Value ?? "image";
            var dataUri = $"data:{contentType};base64,{base64}";

            // Detect full-page background images → render as absolute background
            if (IsFullPageSize(imgCxEmu, imgCyEmu))
            {
                sb.Append($"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;overflow:hidden\">");
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncode(alt)}\" style=\"width:100%;height:100%;object-fit:cover\">");
                sb.Append("</div>");
                return;
            }

            var widthPx = imgCxEmu / 9525;
            var heightPx = imgCyEmu / 9525;
            string widthAttr = widthPx > 0 ? $" width=\"{widthPx}\"" : "";
            string heightAttr = heightPx > 0 ? $" height=\"{heightPx}\"" : "";

            // Crop support: container-based cropping
            var crop = GetCropPercents(drawing);
            if (crop.HasValue)
            {
                RenderCroppedImage(sb, dataUri, widthPx, heightPx, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, HtmlEncode(alt));
            }
            else
            {
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncode(alt)}\"{widthAttr}{heightAttr} style=\"max-width:100%;height:auto\">");
            }
        }
        catch
        {
            sb.Append("<span class=\"img-error\">[Image]</span>");
        }
    }

    /// <summary>
    /// Get crop percentages from a:srcRect.
    /// Values are in 1/1000 of a percent (e.g., 25000 = 25%).
    /// Negative values mean extend (treated as 0).
    /// Returns (left, top, right, bottom) as CSS percentages, or null if no crop.
    /// </summary>
    private static (double l, double t, double r, double b)? GetCropPercents(OpenXmlElement container)
    {
        var srcRect = container.Descendants().FirstOrDefault(e => e.LocalName == "srcRect");
        if (srcRect == null) return null;

        var l = Math.Max(0, GetIntAttr(srcRect, "l") / 1000.0);
        var t = Math.Max(0, GetIntAttr(srcRect, "t") / 1000.0);
        var r = Math.Max(0, GetIntAttr(srcRect, "r") / 1000.0);
        var b = Math.Max(0, GetIntAttr(srcRect, "b") / 1000.0);

        if (l == 0 && t == 0 && r == 0 && b == 0) return null;
        return (l, t, r, b);
    }

    /// <summary>
    /// Render a cropped image using a container div with overflow:hidden.
    /// The image is scaled to its original size and positioned to show only the cropped region.
    /// </summary>
    private static void RenderCroppedImage(StringBuilder sb, string dataUri, long displayWidthPx, long displayHeightPx,
        double cropL, double cropT, double cropR, double cropB, string alt)
    {
        // The display size is the cropped result size.
        // Original image visible fraction: (1 - cropL/100 - cropR/100) horizontally, (1 - cropT/100 - cropB/100) vertically.
        var fracW = 1.0 - cropL / 100.0 - cropR / 100.0;
        var fracH = 1.0 - cropT / 100.0 - cropB / 100.0;
        if (fracW <= 0) fracW = 1; if (fracH <= 0) fracH = 1;

        // Original image size in CSS
        var imgW = displayWidthPx / fracW;
        var imgH = displayHeightPx / fracH;
        // Offset to show the cropped region
        var offsetX = -imgW * (cropL / 100.0);
        var offsetY = -imgH * (cropT / 100.0);

        sb.Append($"<div style=\"display:inline-block;width:{displayWidthPx}px;height:{displayHeightPx}px;overflow:hidden\">");
        sb.Append($"<img src=\"{dataUri}\" alt=\"{alt}\" style=\"width:{imgW:0}px;height:{imgH:0}px;margin-left:{offsetX:0}px;margin-top:{offsetY:0}px\">");
        sb.Append("</div>");
    }

    private static int GetIntAttr(OpenXmlElement el, string attrName)
    {
        var val = el.GetAttributes().FirstOrDefault(a => a.LocalName == attrName).Value;
        return val != null && int.TryParse(val, out var v) ? v : 0;
    }

    // ==================== Group / Shape Rendering ====================

    private void RenderGroupHtml(StringBuilder sb, OpenXmlElement group, long groupWidthEmu, long groupHeightEmu,
        List<Drawing>? floatImages = null)
    {
        var widthPx = groupWidthEmu / 9525;
        var heightPx = groupHeightEmu / 9525;

        // Get the group's child coordinate space from grpSpPr > xfrm
        long chOffX = 0, chOffY = 0, chExtCx = groupWidthEmu, chExtCy = groupHeightEmu;
        var grpSpPr = group.Elements().FirstOrDefault(e => e.LocalName == "grpSpPr");
        var grpXfrm = grpSpPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
        if (grpXfrm != null)
        {
            var chOff = grpXfrm.Elements().FirstOrDefault(e => e.LocalName == "chOff");
            var chExt = grpXfrm.Elements().FirstOrDefault(e => e.LocalName == "chExt");
            if (chOff != null)
            {
                chOffX = GetLongAttr(chOff, "x");
                chOffY = GetLongAttr(chOff, "y");
            }
            if (chExt != null)
            {
                chExtCx = GetLongAttr(chExt, "cx");
                chExtCy = GetLongAttr(chExt, "cy");
            }
        }

        sb.Append($"<div class=\"wg\" style=\"position:relative;width:{widthPx}px;height:{heightPx}px;display:inline-block;overflow:hidden\">");

        // Render each child element (shapes, pictures, nested groups)
        foreach (var child in group.Elements())
        {
            if (child.LocalName is "wsp" or "pic" or "grpSp")
            {
                // Get transform from xfrm (may be in spPr or grpSpPr)
                var xfrm = child.Descendants().FirstOrDefault(e => e.LocalName == "xfrm");
                long offX = 0, offY = 0, extCx = 0, extCy = 0;
                if (xfrm != null)
                {
                    var off = xfrm.Elements().FirstOrDefault(e => e.LocalName == "off");
                    var ext = xfrm.Elements().FirstOrDefault(e => e.LocalName == "ext");
                    if (off != null) { offX = GetLongAttr(off, "x"); offY = GetLongAttr(off, "y"); }
                    if (ext != null) { extCx = GetLongAttr(ext, "cx"); extCy = GetLongAttr(ext, "cy"); }
                }

                // Pass floatImages to first text box shape, then clear
                RenderShapeHtml(sb, child, offX - chOffX, offY - chOffY, extCx, extCy, chExtCx, chExtCy, floatImages);
                floatImages = null; // only inject into first shape
            }
        }

        sb.Append("</div>");
    }

    private void RenderShapeHtml(StringBuilder sb, OpenXmlElement shape, long offX, long offY,
        long extCx, long extCy, long coordSpaceCx, long coordSpaceCy,
        List<Drawing>? floatImages = null)
    {
        // Convert child coordinates to percentage of group
        double leftPct = coordSpaceCx > 0 ? (double)offX / coordSpaceCx * 100 : 0;
        double topPct = coordSpaceCy > 0 ? (double)offY / coordSpaceCy * 100 : 0;
        double widthPct = coordSpaceCx > 0 ? (double)extCx / coordSpaceCx * 100 : 100;
        double heightPct = coordSpaceCy > 0 ? (double)extCy / coordSpaceCy * 100 : 100;

        // Get fill color
        var spPr = shape.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        var fillCss = ResolveShapeFillCss(spPr);

        // Get border
        var borderCss = ResolveShapeBorderCss(spPr);

        // pic elements are always images, not text boxes
        var txbx = shape.LocalName == "pic" ? null
            : shape.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");

        // Rotation from xfrm rot attribute (60000ths of a degree)
        var xfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
        var rot = GetLongAttr(xfrm, "rot");
        var rotCss = rot != 0 ? $";transform:rotate({rot / 60000.0:0.##}deg)" : "";

        // Build style
        var style = $"position:absolute;left:{leftPct:0.##}%;top:{topPct:0.##}%;width:{widthPct:0.##}%;height:{heightPct:0.##}%";
        if (!string.IsNullOrEmpty(fillCss)) style += $";{fillCss}";
        if (!string.IsNullOrEmpty(borderCss)) style += $";{borderCss}";
        style += rotCss;

        // Get body properties for text layout
        var bodyPr = shape.Elements().FirstOrDefault(e => e.LocalName == "bodyPr");
        var vAnchor = bodyPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "anchor").Value;
        if (vAnchor == "ctr") style += ";display:flex;align-items:center";
        else if (vAnchor == "b") style += ";display:flex;align-items:flex-end";

        // Padding from bodyPr insets (EMU → px)
        var lIns = GetLongAttr(bodyPr, "lIns", 91440);
        var tIns = GetLongAttr(bodyPr, "tIns", 45720);
        var rIns = GetLongAttr(bodyPr, "rIns", 91440);
        var bIns = GetLongAttr(bodyPr, "bIns", 45720);
        style += $";padding:{tIns / 9525}px {rIns / 9525}px {bIns / 9525}px {lIns / 9525}px";

        sb.Append($"<div style=\"{style}\">");

        if (txbx != null)
        {
            // Render text box content (standard Word paragraphs)
            sb.Append("<div style=\"width:100%\">");

            // Inject pending float images into this text box
            if (floatImages != null && floatImages.Count > 0)
            {
                foreach (var imgDrawing in floatImages)
                {
                    var imgBlip = imgDrawing.Descendants<A.Blip>().FirstOrDefault();
                    if (imgBlip?.Embed?.Value == null) continue;
                    try
                    {
                        var imgPart = _doc.MainDocumentPart?.GetPartById(imgBlip.Embed.Value) as ImagePart;
                        if (imgPart == null) continue;
                        using var imgStream = imgPart.GetStream();
                        using var imgMs = new MemoryStream();
                        imgStream.CopyTo(imgMs);
                        var imgBase64 = Convert.ToBase64String(imgMs.ToArray());
                        var imgExtent = imgDrawing.Descendants<DW.Extent>().FirstOrDefault();
                        var imgW = imgExtent?.Cx?.Value > 0 ? imgExtent.Cx.Value / 9525 : 100;
                        var imgH = imgExtent?.Cy?.Value > 0 ? imgExtent.Cy.Value / 9525 : 100;
                        var imgDataUri = $"data:{imgPart.ContentType};base64,{imgBase64}";
                        // Read distT/distB/distL/distR for image margins (EMU)
                        var inline = imgDrawing.Descendants<DW.Inline>().FirstOrDefault();
                        var anchor = imgDrawing.Descendants<DW.Anchor>().FirstOrDefault();
                        long distT = 0, distB = 0, distL = 0, distR = 0;
                        if (inline != null)
                        {
                            distT = (long)(inline.DistanceFromTop?.Value ?? 0);
                            distB = (long)(inline.DistanceFromBottom?.Value ?? 0);
                            distL = (long)(inline.DistanceFromLeft?.Value ?? 0);
                            distR = (long)(inline.DistanceFromRight?.Value ?? 0);
                        }
                        else if (anchor != null)
                        {
                            distT = (long)(anchor.DistanceFromTop?.Value ?? 0);
                            distB = (long)(anchor.DistanceFromBottom?.Value ?? 0);
                            distL = (long)(anchor.DistanceFromLeft?.Value ?? 0);
                            distR = (long)(anchor.DistanceFromRight?.Value ?? 0);
                        }
                        var marginCss = $"margin:{distT/9525}px {distR/9525}px {distB/9525}px {distL/9525}px";
                        var crop = GetCropPercents(imgDrawing);
                        if (crop.HasValue)
                        {
                            sb.Append($"<div style=\"float:left;{marginCss}\">");
                            RenderCroppedImage(sb, imgDataUri, imgW, imgH, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, "");
                            sb.Append("</div>");
                        }
                        else
                        {
                            sb.Append($"<img src=\"{imgDataUri}\" style=\"float:left;width:{imgW}px;height:{imgH}px;object-fit:cover;{marginCss}\">");
                        }
                    }
                    catch { }
                }
                floatImages = null;
            }

            foreach (var para in txbx.Descendants<Paragraph>())
            {
                RenderParagraphHtml(sb, para);
            }
            sb.Append("</div>");
        }
        else
        {
            // Check for image inside shape (spPr blip or pic:blipFill blip)
            var embedAttr = FindEmbedInDescendants(shape);
            if (embedAttr != null)
            {
                try
                {
                    var mainPart = _doc.MainDocumentPart;
                    var imagePart = mainPart?.GetPartById(embedAttr) as ImagePart;
                    if (imagePart != null)
                    {
                        using var stream = imagePart.GetStream();
                        using var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        var base64 = Convert.ToBase64String(ms.ToArray());
                        sb.Append($"<img src=\"data:{imagePart.ContentType};base64,{base64}\" style=\"width:100%;height:100%;object-fit:contain\">");
                    }
                }
                catch { }
            }
        }

        sb.Append("</div>");
    }

    // ==================== Theme Color Resolution ====================

    private Dictionary<string, string>? _themeColors;

    private Dictionary<string, string> GetThemeColors()
    {
        if (_themeColors != null) return _themeColors;

        _themeColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var theme = _doc.MainDocumentPart?.ThemePart?.Theme;
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return _themeColors;

        void Add(string name, OpenXmlCompositeElement? color)
        {
            if (color == null) return;
            var rgb = color.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            var sys = color.GetFirstChild<A.SystemColor>();
            var srgb = sys?.LastColor?.Value;
            var hex = rgb ?? srgb;
            if (hex != null) _themeColors[name] = hex;
        }

        Add("dk1", colorScheme.Dark1Color);
        Add("dk2", colorScheme.Dark2Color);
        Add("lt1", colorScheme.Light1Color);
        Add("lt2", colorScheme.Light2Color);
        Add("accent1", colorScheme.Accent1Color);
        Add("accent2", colorScheme.Accent2Color);
        Add("accent3", colorScheme.Accent3Color);
        Add("accent4", colorScheme.Accent4Color);
        Add("accent5", colorScheme.Accent5Color);
        Add("accent6", colorScheme.Accent6Color);
        Add("hlink", colorScheme.Hyperlink);
        Add("folHlink", colorScheme.FollowedHyperlinkColor);

        // Aliases
        if (_themeColors.TryGetValue("dk1", out var dk1)) { _themeColors["tx1"] = dk1; _themeColors["dark1"] = dk1; }
        if (_themeColors.TryGetValue("lt1", out var lt1)) { _themeColors["bg1"] = lt1; _themeColors["light1"] = lt1; }
        if (_themeColors.TryGetValue("lt2", out var lt2)) { _themeColors["bg2"] = lt2; _themeColors["light2"] = lt2; }

        return _themeColors;
    }

    private string? ResolveSchemeColor(OpenXmlElement schemeColor)
    {
        var schemeName = schemeColor.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (schemeName == null) return null;

        var themeColors = GetThemeColors();
        if (!themeColors.TryGetValue(schemeName, out var hex)) return null;

        // Apply color transforms (lumMod, lumOff, tint, shade)
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        var lumMod = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "lumMod");
        var lumOff = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "lumOff");
        var tint = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "tint");
        var shade = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "shade");

        if (tint != null)
        {
            var t = GetLongAttr(tint, "val") / 100000.0;
            r = (int)(r + (255 - r) * (1 - t));
            g = (int)(g + (255 - g) * (1 - t));
            b = (int)(b + (255 - b) * (1 - t));
        }

        if (shade != null)
        {
            var s = GetLongAttr(shade, "val") / 100000.0;
            r = (int)(r * s);
            g = (int)(g * s);
            b = (int)(b * s);
        }

        if (lumMod != null || lumOff != null)
        {
            var mod = (lumMod != null ? GetLongAttr(lumMod, "val") : 100000) / 100000.0;
            var off = (lumOff != null ? GetLongAttr(lumOff, "val") : 0) / 100000.0;
            RgbToHsl(r, g, b, out var h, out var s, out var l);
            l = Math.Clamp(l * mod + off, 0, 1);
            HslToRgb(h, s, l, out r, out g, out b);
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private string ResolveShapeFillCss(OpenXmlElement? spPr)
    {
        if (spPr == null) return "";

        // No fill
        if (spPr.Elements().Any(e => e.LocalName == "noFill")) return "";

        // Solid fill
        var solidFill = spPr.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        if (solidFill != null)
        {
            var rgb = solidFill.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
            if (rgb != null)
            {
                var val = rgb.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                if (val != null) return $"background-color:#{val}";
            }
            var scheme = solidFill.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
            if (scheme != null)
            {
                var color = ResolveSchemeColor(scheme);
                if (color != null) return $"background-color:{color}";
            }
        }

        return "";
    }

    private string ResolveShapeBorderCss(OpenXmlElement? spPr)
    {
        if (spPr == null) return "";
        var ln = spPr.Elements().FirstOrDefault(e => e.LocalName == "ln");
        if (ln == null) return "";
        if (ln.Elements().Any(e => e.LocalName == "noFill")) return "border:none";

        var solidFill = ln.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        if (solidFill == null) return "";

        string? color = null;
        var rgb = solidFill.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        if (rgb != null) color = $"#{rgb.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value}";
        var scheme = solidFill.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
        if (scheme != null) color = ResolveSchemeColor(scheme);

        var w = ln.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
        var widthPx = w != null && long.TryParse(w, out var emu) ? Math.Max(1, emu / 12700.0) : 1;

        return $"border:{widthPx:0.#}px solid {color ?? "#000"}";
    }

    // ==================== Color Math Helpers ====================

    /// <summary>Apply themeTint/themeShade to a base theme color hex.</summary>
    private static string ApplyTintShade(string hex, string? tintHex, string? shadeHex)
    {
        if (hex.Length < 6) return $"#{hex}";
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        // themeTint: blend toward white (tint value is hex 00-FF)
        if (tintHex != null && int.TryParse(tintHex, System.Globalization.NumberStyles.HexNumber, null, out var tint))
        {
            var t = tint / 255.0;
            r = (int)(r * t + 255 * (1 - t));
            g = (int)(g * t + 255 * (1 - t));
            b = (int)(b * t + 255 * (1 - t));
        }

        // themeShade: blend toward black
        if (shadeHex != null && int.TryParse(shadeHex, System.Globalization.NumberStyles.HexNumber, null, out var shade))
        {
            var s = shade / 255.0;
            r = (int)(r * s);
            g = (int)(g * s);
            b = (int)(b * s);
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private static long GetLongAttr(OpenXmlElement? el, string attrName, long defaultVal = 0)
    {
        if (el == null) return defaultVal;
        var val = el.GetAttributes().FirstOrDefault(a => a.LocalName == attrName).Value;
        return val != null && long.TryParse(val, out var v) ? v : defaultVal;
    }

    private static void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
    {
        var rf = r / 255.0; var gf = g / 255.0; var bf = b / 255.0;
        var max = Math.Max(rf, Math.Max(gf, bf));
        var min = Math.Min(rf, Math.Min(gf, bf));
        var delta = max - min;
        l = (max + min) / 2.0;
        if (delta < 1e-10) { h = 0; s = 0; return; }
        s = l < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);
        if (Math.Abs(max - rf) < 1e-10) h = ((gf - bf) / delta + (gf < bf ? 6 : 0)) / 6.0;
        else if (Math.Abs(max - gf) < 1e-10) h = ((bf - rf) / delta + 2) / 6.0;
        else h = ((rf - gf) / delta + 4) / 6.0;
    }

    private static void HslToRgb(double h, double s, double l, out int r, out int g, out int b)
    {
        if (s < 1e-10) { r = g = b = (int)Math.Round(l * 255); return; }
        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        r = (int)Math.Round(HueToRgb(p, q, h + 1.0 / 3) * 255);
        g = (int)Math.Round(HueToRgb(p, q, h) * 255);
        b = (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3) * 255);
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1; if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    // ==================== Table Rendering ====================

    private void RenderTableHtml(StringBuilder sb, Table table)
    {
        // Check table-level borders to determine if this is a borderless layout table
        var tblBorders = table.GetFirstChild<TableProperties>()?.TableBorders;
        bool tableBordersNone = IsTableBorderless(tblBorders);

        var tableClass = tableBordersNone ? "borderless" : "";
        sb.AppendLine(string.IsNullOrEmpty(tableClass) ? "<table>" : $"<table class=\"{tableClass}\">");

        // Get column widths from grid
        var tblGrid = table.GetFirstChild<TableGrid>();
        if (tblGrid != null)
        {
            sb.Append("<colgroup>");
            foreach (var col in tblGrid.Elements<GridColumn>())
            {
                var w = col.Width?.Value;
                if (w != null)
                {
                    var px = (int)(double.Parse(w) / 1440.0 * 96); // twips to px
                    sb.Append($"<col style=\"width:{px}px\">");
                }
                else
                {
                    sb.Append("<col>");
                }
            }
            sb.AppendLine("</colgroup>");
        }

        foreach (var row in table.Elements<TableRow>())
        {
            var isHeader = row.TableRowProperties?.GetFirstChild<TableHeader>() != null;
            sb.AppendLine(isHeader ? "<tr class=\"header-row\">" : "<tr>");

            foreach (var cell in row.Elements<TableCell>())
            {
                var tag = isHeader ? "th" : "td";
                var cellStyle = GetTableCellInlineCss(cell, tableBordersNone);

                // Merge attributes
                var attrs = new StringBuilder();
                var gridSpan = cell.TableCellProperties?.GridSpan?.Val?.Value;
                if (gridSpan > 1) attrs.Append($" colspan=\"{gridSpan}\"");

                var vMerge = cell.TableCellProperties?.VerticalMerge;
                if (vMerge != null && vMerge.Val?.Value == MergedCellValues.Restart)
                {
                    // Count rowspan
                    var rowspan = CountRowSpan(table, row, cell);
                    if (rowspan > 1) attrs.Append($" rowspan=\"{rowspan}\"");
                }
                else if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
                {
                    continue; // Skip merged continuation cells
                }

                if (!string.IsNullOrEmpty(cellStyle))
                    attrs.Append($" style=\"{cellStyle}\"");

                sb.Append($"<{tag}{attrs}>");

                // Render cell content — use paragraph tags for multi-paragraph cells
                var cellParagraphs = cell.Elements<Paragraph>().ToList();
                for (int pi = 0; pi < cellParagraphs.Count; pi++)
                {
                    var cellPara = cellParagraphs[pi];
                    var text = GetParagraphText(cellPara);
                    var runs = GetAllRuns(cellPara);

                    if (runs.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        // empty cell paragraph — skip but preserve spacing between paragraphs
                        if (pi > 0 && pi < cellParagraphs.Count - 1)
                            sb.Append("<br>");
                    }
                    else
                    {
                        var pCss = GetParagraphInlineCss(cellPara);
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append($"<div style=\"{pCss}\">");
                        RenderParagraphContentHtml(sb, cellPara);
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append("</div>");
                        else if (pi < cellParagraphs.Count - 1)
                            sb.Append("<br>");
                    }
                }

                // Render nested tables
                foreach (var nestedTable in cell.Elements<Table>())
                    RenderTableHtml(sb, nestedTable);

                sb.AppendLine($"</{tag}>");
            }

            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</table>");
    }

    private static bool IsTableBorderless(TableBorders? borders)
    {
        if (borders == null) return false;
        // Check if all borders are none/nil
        return IsBorderNone(borders.TopBorder)
            && IsBorderNone(borders.BottomBorder)
            && IsBorderNone(borders.LeftBorder)
            && IsBorderNone(borders.RightBorder)
            && IsBorderNone(borders.InsideHorizontalBorder)
            && IsBorderNone(borders.InsideVerticalBorder);
    }

    private static bool IsBorderNone(OpenXmlElement? border)
    {
        if (border == null) return true;
        var val = border.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return val is null or "nil" or "none";
    }

    private static int CountRowSpan(Table table, TableRow startRow, TableCell startCell)
    {
        var rows = table.Elements<TableRow>().ToList();
        var startRowIdx = rows.IndexOf(startRow);
        var cellIdx = startRow.Elements<TableCell>().ToList().IndexOf(startCell);
        if (startRowIdx < 0 || cellIdx < 0) return 1;

        int span = 1;
        for (int i = startRowIdx + 1; i < rows.Count; i++)
        {
            var cells = rows[i].Elements<TableCell>().ToList();
            if (cellIdx >= cells.Count) break;

            var vm = cells[cellIdx].TableCellProperties?.VerticalMerge;
            if (vm != null && (vm.Val == null || vm.Val.Value == MergedCellValues.Continue))
                span++;
            else
                break;
        }
        return span;
    }

    // ==================== Inline CSS ====================

    private string GetParagraphInlineCss(Paragraph para, bool isListItem = false)
    {
        var parts = new List<string>();

        var pProps = para.ParagraphProperties;
        if (pProps == null) return ResolveParagraphStyleCss(para);

        // Alignment
        var jc = pProps.Justification?.Val;
        if (jc != null)
        {
            var align = jc.InnerText switch
            {
                "center" => "center",
                "right" or "end" => "right",
                "both" or "distribute" => "justify",
                _ => (string?)null
            };
            if (align != null) parts.Add($"text-align:{align}");
        }

        // Indentation (skip for list items — handled by list nesting)
        if (!isListItem)
        {
            var indent = pProps.Indentation;
            if (indent != null)
            {
                if (indent.Left?.Value is string leftTwips && leftTwips != "0")
                    parts.Add($"margin-left:{TwipsToPx(leftTwips)}px");
                if (indent.Right?.Value is string rightTwips && rightTwips != "0")
                    parts.Add($"margin-right:{TwipsToPx(rightTwips)}px");
                if (indent.FirstLine?.Value is string firstLineTwips && firstLineTwips != "0")
                    parts.Add($"text-indent:{TwipsToPx(firstLineTwips)}px");
                if (indent.Hanging?.Value is string hangTwips && hangTwips != "0")
                    parts.Add($"text-indent:-{TwipsToPx(hangTwips)}px");
            }
        }

        // Spacing
        var spacing = pProps.SpacingBetweenLines;
        if (spacing != null)
        {
            if (spacing.Before?.Value is string beforeTwips && beforeTwips != "0")
                parts.Add($"margin-top:{TwipsToPx(beforeTwips)}px");
            if (spacing.After?.Value is string afterTwips && afterTwips != "0")
                parts.Add($"margin-bottom:{TwipsToPx(afterTwips)}px");
            if (spacing.Line?.Value is string lineVal)
            {
                var rule = spacing.LineRule?.InnerText;
                if (rule == "auto" || rule == null)
                {
                    // Multiplier: value/240 = line spacing ratio
                    if (int.TryParse(lineVal, out var lv))
                        parts.Add($"line-height:{lv / 240.0:0.##}");
                }
                else if (rule == "exact" || rule == "atLeast")
                {
                    parts.Add($"line-height:{TwipsToPx(lineVal)}px");
                }
            }
        }

        // Shading / background (direct or from style)
        var shading = pProps.Shading;
        if (shading?.Fill?.Value is string fill && fill != "auto")
            parts.Add($"background-color:#{fill}");
        else
        {
            // Try to resolve from paragraph style
            var bgFromStyle = ResolveParagraphShadingFromStyle(para);
            if (bgFromStyle != null) parts.Add($"background-color:#{bgFromStyle}");
        }

        // Borders
        var pBdr = pProps.ParagraphBorders;
        if (pBdr != null)
        {
            RenderBorderCss(parts, pBdr.TopBorder, "border-top");
            RenderBorderCss(parts, pBdr.BottomBorder, "border-bottom");
            RenderBorderCss(parts, pBdr.LeftBorder, "border-left");
            RenderBorderCss(parts, pBdr.RightBorder, "border-right");
        }

        return string.Join(";", parts);
    }

    /// <summary>
    /// Resolve paragraph background shading from the style chain.
    /// </summary>
    private string? ResolveParagraphShadingFromStyle(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;

            var shading = style.StyleParagraphProperties?.Shading;
            if (shading?.Fill?.Value is string fill && fill != "auto")
                return fill;

            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>
    /// Resolve paragraph CSS from style chain when no direct paragraph properties.
    /// </summary>
    private string ResolveParagraphStyleCss(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return "";

        var parts = new List<string>();
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;

            var pPr = style.StyleParagraphProperties;
            if (pPr != null)
            {
                var jc = pPr.Justification?.Val;
                if (jc != null && !parts.Any(p => p.StartsWith("text-align")))
                {
                    var align = jc.InnerText switch { "center" => "center", "right" or "end" => "right", "both" => "justify", _ => (string?)null };
                    if (align != null) parts.Add($"text-align:{align}");
                }

                var spacing = pPr.SpacingBetweenLines;
                if (spacing != null)
                {
                    if (spacing.Before?.Value is string b && b != "0" && !parts.Any(p => p.StartsWith("margin-top")))
                        parts.Add($"margin-top:{TwipsToPx(b)}px");
                    if (spacing.After?.Value is string a && a != "0" && !parts.Any(p => p.StartsWith("margin-bottom")))
                        parts.Add($"margin-bottom:{TwipsToPx(a)}px");
                    if (spacing.Line?.Value is string lv && !parts.Any(p => p.StartsWith("line-height")))
                    {
                        var rule = spacing.LineRule?.InnerText;
                        if ((rule == "auto" || rule == null) && int.TryParse(lv, out var val))
                            parts.Add($"line-height:{val / 240.0:0.##}");
                    }
                }

                var shading = pPr.Shading;
                if (shading?.Fill?.Value is string fill && fill != "auto" && !parts.Any(p => p.StartsWith("background")))
                    parts.Add($"background-color:#{fill}");
            }

            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return string.Join(";", parts);
    }

    private string GetRunInlineCss(RunProperties? rProps)
    {
        if (rProps == null) return "";
        var parts = new List<string>();

        // Font
        var fonts = rProps.RunFonts;
        var font = fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
        if (font != null) parts.Add($"font-family:'{CssSanitize(font)}'");

        // Size (stored as half-points)
        var size = rProps.FontSize?.Val?.Value;
        if (size != null && int.TryParse(size, out var halfPts))
            parts.Add($"font-size:{halfPts / 2.0:0.##}pt");

        // Bold
        if (rProps.Bold != null)
            parts.Add("font-weight:bold");

        // Italic
        if (rProps.Italic != null)
            parts.Add("font-style:italic");

        // Underline
        if (rProps.Underline?.Val != null)
        {
            var ulVal = rProps.Underline.Val.InnerText;
            if (ulVal != "none")
                parts.Add("text-decoration:underline");
        }

        // Strikethrough
        if (rProps.Strike != null)
        {
            var existing = parts.FirstOrDefault(p => p.StartsWith("text-decoration:"));
            if (existing != null)
            {
                parts.Remove(existing);
                parts.Add(existing + " line-through");
            }
            else
            {
                parts.Add("text-decoration:line-through");
            }
        }

        // Color: w:color val is the pre-computed color (already has themeColor+themeTint applied).
        // Use val directly; only fall back to theme resolution if val is missing.
        var colorVal = rProps.Color?.Val?.Value;
        if (colorVal != null && colorVal != "auto")
        {
            parts.Add($"color:#{colorVal}");
        }
        else if (rProps.Color?.ThemeColor?.InnerText is string tcName)
        {
            var tc = GetThemeColors();
            if (tc.TryGetValue(tcName, out var tcHex))
            {
                var tint = rProps.Color?.GetAttributes().FirstOrDefault(a => a.LocalName == "themeTint").Value;
                var shade = rProps.Color?.GetAttributes().FirstOrDefault(a => a.LocalName == "themeShade").Value;
                parts.Add($"color:{ApplyTintShade(tcHex, tint, shade)}");
            }
        }

        // Highlight
        var highlight = rProps.Highlight?.Val?.InnerText;
        if (highlight != null)
        {
            var hlColor = HighlightToCssColor(highlight);
            if (hlColor != null) parts.Add($"background-color:{hlColor}");
        }

        // Superscript / Subscript
        var vertAlign = rProps.VerticalTextAlignment?.Val;
        if (vertAlign != null)
        {
            if (vertAlign.InnerText == "superscript")
                parts.Add("vertical-align:super;font-size:smaller");
            else if (vertAlign.InnerText == "subscript")
                parts.Add("vertical-align:sub;font-size:smaller");
        }

        return string.Join(";", parts);
    }

    private string GetTableCellInlineCss(TableCell cell, bool tableBordersNone)
    {
        var parts = new List<string>();
        var tcPr = cell.TableCellProperties;

        // If table-level borders are none, explicitly set border:none on cells
        if (tableBordersNone)
            parts.Add("border:none");

        if (tcPr == null) return string.Join(";", parts);

        // Shading / fill
        var shading = tcPr.Shading;
        if (shading?.Fill?.Value is string fill && fill != "auto")
            parts.Add($"background-color:#{fill}");

        // Vertical alignment
        var vAlign = tcPr.TableCellVerticalAlignment?.Val;
        if (vAlign != null)
        {
            var va = vAlign.InnerText switch
            {
                "center" => "middle",
                "bottom" => "bottom",
                _ => (string?)null
            };
            if (va != null) parts.Add($"vertical-align:{va}");
        }

        // Cell borders (override table-level setting if cell has its own)
        var tcBorders = tcPr.TableCellBorders;
        if (tcBorders != null)
        {
            // Remove the table-level border:none if cell has specific borders
            if (tableBordersNone)
                parts.Remove("border:none");
            RenderBorderCss(parts, tcBorders.TopBorder, "border-top");
            RenderBorderCss(parts, tcBorders.BottomBorder, "border-bottom");
            RenderBorderCss(parts, tcBorders.LeftBorder, "border-left");
            RenderBorderCss(parts, tcBorders.RightBorder, "border-right");
        }

        // Cell width
        var width = tcPr.TableCellWidth?.Width?.Value;
        if (width != null && int.TryParse(width, out var w))
        {
            var type = tcPr.TableCellWidth?.Type?.InnerText;
            if (type == "dxa")
                parts.Add($"width:{w / 1440.0 * 96:0}px");
            else if (type == "pct")
                parts.Add($"width:{w / 50.0:0.#}%");
        }

        // Padding
        var margins = tcPr.TableCellMargin;
        if (margins != null)
        {
            var padTop = margins.TopMargin?.Width?.Value;
            var padBot = margins.BottomMargin?.Width?.Value;
            var padLeft = margins.LeftMargin?.Width?.Value ?? margins.StartMargin?.Width?.Value;
            var padRight = margins.RightMargin?.Width?.Value ?? margins.EndMargin?.Width?.Value;
            if (padTop != null || padBot != null || padLeft != null || padRight != null)
            {
                parts.Add($"padding:{TwipsToPxStr(padTop ?? "0")} {TwipsToPxStr(padRight ?? "0")} {TwipsToPxStr(padBot ?? "0")} {TwipsToPxStr(padLeft ?? "0")}");
            }
        }

        return string.Join(";", parts);
    }

    // ==================== CSS Helpers ====================

    private static void RenderBorderCss(List<string> parts, OpenXmlElement? border, string cssProp)
    {
        if (border == null) return;
        var val = border.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (val == null || val == "nil" || val == "none") return;

        var sz = border.GetAttributes().FirstOrDefault(a => a.LocalName == "sz").Value;
        var color = border.GetAttributes().FirstOrDefault(a => a.LocalName == "color").Value;

        var width = sz != null && int.TryParse(sz, out var s) ? $"{Math.Max(1, s / 8.0):0.#}px" : "1px";
        var style = val switch
        {
            "single" => "solid",
            "double" => "double",
            "dashed" or "dashSmallGap" => "dashed",
            "dotted" => "dotted",
            _ => "solid"
        };
        var cssColor = (color != null && color != "auto") ? $"#{color}" : "#000";

        parts.Add($"{cssProp}:{width} {style} {cssColor}");
    }

    private static int TwipsToPx(string twipsStr)
    {
        if (!int.TryParse(twipsStr, out var twips)) return 0;
        return (int)(twips / 1440.0 * 96);
    }

    private static string TwipsToPxStr(string twipsStr)
    {
        return $"{TwipsToPx(twipsStr)}px";
    }

    private static string? HighlightToCssColor(string highlight) => highlight.ToLowerInvariant() switch
    {
        "yellow" => "#FFFF00",
        "green" => "#00FF00",
        "cyan" => "#00FFFF",
        "magenta" => "#FF00FF",
        "blue" => "#0000FF",
        "red" => "#FF0000",
        "darkblue" => "#00008B",
        "darkcyan" => "#008B8B",
        "darkgreen" => "#006400",
        "darkmagenta" => "#8B008B",
        "darkred" => "#8B0000",
        "darkyellow" => "#808000",
        "darkgray" => "#A9A9A9",
        "lightgray" => "#D3D3D3",
        "black" => "#000000",
        "white" => "#FFFFFF",
        _ => null
    };

    private static string CssSanitize(string value) =>
        Regex.Replace(value, @"[""'\\<>&;{}]", "");

    private static string HtmlEncode(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    // ==================== CSS Stylesheet ====================

    private static string GenerateWordCss(PageLayout pg, DocDef dd)
    {
        var mL = $"{pg.MarginLeftCm:0.##}cm";
        var mR = $"{pg.MarginRightCm:0.##}cm";
        var mT = $"{pg.MarginTopCm:0.##}cm";
        var mB = $"{pg.MarginBottomCm:0.##}cm";
        var lr = $"{pg.MarginLeftCm:0.##}cm {pg.MarginRightCm:0.##}cm";
        var font = $"\'{CssSanitize(dd.Font)}\', \'Microsoft YaHei\', \'Segoe UI\', -apple-system, \'PingFang SC\', sans-serif";
        var pageH = $"{pg.HeightCm:0.##}cm";
        var sz = $"{dd.SizePt:0.##}pt";
        var lh = $"{dd.LineHeight:0.##}";
        var tblW = $"calc(100% - {pg.MarginLeftCm + pg.MarginRightCm:0.##}cm)";

        return $@"
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ background: #f0f0f0; font-family: {font}; color: {dd.Color}; padding: 20px; }}
        .page {{ background: white; margin: 0 auto 40px; padding: {mT} {mR} {mB} {mL};
            box-shadow: 0 2px 8px rgba(0,0,0,0.15); border-radius: 4px;
            min-height: {pageH}; line-height: {lh}; font-size: {sz}; position: relative; overflow: hidden; }}
        .doc-header, .doc-footer {{ color: #888; font-size: 9pt;
            border-bottom: 1px solid #e0e0e0; margin-bottom: 1em; padding-bottom: 0.5em; }}
        .doc-footer {{ border-bottom: none; border-top: 1px solid #e0e0e0;
            margin-top: 1em; padding-top: 0.5em; margin-bottom: 0; }}
        h1, h2, h3, h4, h5, h6 {{ line-height: 1.4; }}
        h1 {{ margin-top: 0.5em; margin-bottom: 0.3em; }}
        h2 {{ margin-top: 0.4em; margin-bottom: 0.2em; }}
        h3 {{ margin-top: 0.3em; margin-bottom: 0.2em; }}
        h4 {{ margin-top: 0.2em; margin-bottom: 0.1em; }}
        h5, h6 {{ margin-top: 0.1em; margin-bottom: 0.1em; }}
        p {{ margin: 0.1em 0; }}
        p.empty {{ margin: 0; line-height: 0.8; font-size: 6pt; }}
        a {{ color: #2B579A; }} a:hover {{ color: #1a3c6e; }}
        ul, ol {{ padding-left: 2em; margin: 0.2em 0; }}
        li {{ margin: 0.1em 0; }}
        .equation {{ text-align: center; padding: 0.5em 0; overflow-x: auto; }}
        img {{ max-width: 100%; height: auto; }}
        .img-error {{ color: #999; font-style: italic; }}
        table {{ border-collapse: collapse; font-size: {sz}; width: 100%; }}
        .wg {{ margin: 0.3em 0; }}
        .wg p {{ padding: 0; margin: 0.05em 0; }}
        table.borderless {{ border: none; }}
        table.borderless td, table.borderless th {{ border: none; padding: 2px 6px; }}
        th, td {{ border: 1px solid #bbb; padding: 4px 8px; text-align: left; vertical-align: top; }}
        th {{ background: #f0f0f0; font-weight: 600; }}
        .header-row td, .header-row th {{ background: #f0f0f0; font-weight: 600; }}
        hr.page-break {{ border: none; border-top: 2px dashed #ccc; margin: 2em 0; }}
        @media print {{ body {{ background: white; padding: 0; }}
            .page {{ box-shadow: none; margin: 0; max-width: none; }}
            hr.page-break {{ page-break-after: always; border: none; margin: 0; }} }}";
    }
}
