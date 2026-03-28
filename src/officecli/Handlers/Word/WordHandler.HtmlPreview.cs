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
        // Check if paragraph has text box drawings — if so, skip fallback text runs
        bool hasTextBoxDrawing = HasTextBoxContent(para);
        bool textBoxRendered = false;

        // Collect standalone images that precede text box groups (they overlay the group in Word)
        var preGroupImages = new List<Drawing>();

        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
            {
                // If this run contains a text box drawing, render it
                var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();
                if (drawing != null && HasGroupOrShape(drawing))
                {
                    // Render group with any preceding images overlaid
                    RenderDrawingWithOverlaidImages(sb, drawing, preGroupImages);
                    preGroupImages.Clear();
                    textBoxRendered = true;
                    continue;
                }

                // Collect standalone images before text box group
                if (hasTextBoxDrawing && !textBoxRendered && drawing != null)
                {
                    preGroupImages.Add(drawing);
                    continue;
                }

                // Skip fallback text runs after text box has been rendered
                if (hasTextBoxDrawing && textBoxRendered)
                    continue;

                RenderRunHtml(sb, run, para);
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
            RenderDrawingHtml(sb, groupDrawing);
            return;
        }

        // Inject floating images into the group's first text box
        _pendingFloatImages = overlaidImages;
        RenderDrawingHtml(sb, groupDrawing);
        _pendingFloatImages = null;
    }

    /// <summary>Images to float-inject into the next text box rendered.</summary>
    private List<Drawing>? _pendingFloatImages;

    // ==================== Drawing Rendering (images, groups, shapes) ====================

    /// <summary>Check if a paragraph contains text box drawings.</summary>
    private static bool HasTextBoxContent(Paragraph para)
    {
        foreach (var run in para.Elements<Run>())
        {
            var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();
            if (drawing != null && HasGroupOrShape(drawing))
                return true;
        }
        return false;
    }

    /// <summary>Check if a drawing contains groups or shapes with text boxes.</summary>
    private static bool HasGroupOrShape(Drawing drawing)
    {
        return drawing.Descendants().Any(e => e.LocalName == "wgp" || e.LocalName == "wsp");
    }

    private void RenderDrawingHtml(StringBuilder sb, Drawing drawing)
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
                RenderGroupHtml(sb, group, groupWidthEmu, groupHeightEmu);
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
                RenderShapeHtml(sb, shape, 0, 0, shapeWidth, shapeHeight, shapeWidth, shapeHeight);
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
            string widthAttr = "", heightAttr = "";
            if (extent is DW.Extent dwExt)
            {
                if (dwExt.Cx?.Value > 0) widthAttr = $" width=\"{dwExt.Cx.Value / 9525}\"";
                if (dwExt.Cy?.Value > 0) heightAttr = $" height=\"{dwExt.Cy.Value / 9525}\"";
            }
            else if (extent is A.Extents aExt)
            {
                if (aExt.Cx?.Value > 0) widthAttr = $" width=\"{aExt.Cx.Value / 9525}\"";
                if (aExt.Cy?.Value > 0) heightAttr = $" height=\"{aExt.Cy.Value / 9525}\"";
            }

            var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            var alt = docProps?.Description?.Value ?? docProps?.Name?.Value ?? "image";

            // Crop support: a:srcRect on blipFill
            var cropCss = GetCropCss(drawing);
            var styleParts = new List<string> { "max-width:100%", "height:auto" };
            if (!string.IsNullOrEmpty(cropCss)) styleParts.Add(cropCss);

            sb.Append($"<img src=\"data:{contentType};base64,{base64}\" alt=\"{HtmlEncode(alt)}\"{widthAttr}{heightAttr} style=\"{string.Join(";", styleParts)}\">");
        }
        catch
        {
            sb.Append("<span class=\"img-error\">[Image]</span>");
        }
    }

    /// <summary>
    /// Extract CSS clip-path from a:srcRect crop data.
    /// srcRect l/t/r/b are in 1/1000 of a percent (e.g., 25000 = 25%).
    /// Negative values mean extend (no crop on that side).
    /// </summary>
    private static string GetCropCss(OpenXmlElement container)
    {
        // Look for srcRect in blipFill
        var srcRect = container.Descendants().FirstOrDefault(e => e.LocalName == "srcRect");
        if (srcRect == null) return "";

        var l = GetIntAttr(srcRect, "l");
        var t = GetIntAttr(srcRect, "t");
        var r = GetIntAttr(srcRect, "r");
        var b = GetIntAttr(srcRect, "b");

        // Skip if no positive crop values
        if (l <= 0 && t <= 0 && r <= 0 && b <= 0) return "";

        // Convert from 1/1000 percent to CSS percent
        var top = Math.Max(0, t / 1000.0);
        var right = Math.Max(0, r / 1000.0);
        var bottom = Math.Max(0, b / 1000.0);
        var left = Math.Max(0, l / 1000.0);

        return $"clip-path:inset({top:0.##}% {right:0.##}% {bottom:0.##}% {left:0.##}%)";
    }

    private static int GetIntAttr(OpenXmlElement el, string attrName)
    {
        var val = el.GetAttributes().FirstOrDefault(a => a.LocalName == attrName).Value;
        return val != null && int.TryParse(val, out var v) ? v : 0;
    }

    // ==================== Group / Shape Rendering ====================

    private void RenderGroupHtml(StringBuilder sb, OpenXmlElement group, long groupWidthEmu, long groupHeightEmu)
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

        // Render each child shape
        foreach (var child in group.Elements())
        {
            if (child.LocalName == "wsp")
            {
                // Get shape transform
                var spPr = child.Elements().FirstOrDefault(e => e.LocalName == "spPr");
                var xfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
                long offX = 0, offY = 0, extCx = 0, extCy = 0;
                if (xfrm != null)
                {
                    var off = xfrm.Elements().FirstOrDefault(e => e.LocalName == "off");
                    var ext = xfrm.Elements().FirstOrDefault(e => e.LocalName == "ext");
                    if (off != null) { offX = GetLongAttr(off, "x"); offY = GetLongAttr(off, "y"); }
                    if (ext != null) { extCx = GetLongAttr(ext, "cx"); extCy = GetLongAttr(ext, "cy"); }
                }

                RenderShapeHtml(sb, child, offX - chOffX, offY - chOffY, extCx, extCy, chExtCx, chExtCy);
            }
        }

        sb.Append("</div>");
    }

    private void RenderShapeHtml(StringBuilder sb, OpenXmlElement shape, long offX, long offY,
        long extCx, long extCy, long coordSpaceCx, long coordSpaceCy)
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

        // Check for text box content
        var txbx = shape.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");

        // Build style
        var style = $"position:absolute;left:{leftPct:0.##}%;top:{topPct:0.##}%;width:{widthPct:0.##}%;height:{heightPct:0.##}%";
        if (!string.IsNullOrEmpty(fillCss)) style += $";{fillCss}";
        if (!string.IsNullOrEmpty(borderCss)) style += $";{borderCss}";

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
            if (_pendingFloatImages != null && _pendingFloatImages.Count > 0)
            {
                foreach (var imgDrawing in _pendingFloatImages)
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
                        var cropCss = GetCropCss(imgDrawing);
                        var imgStyle = $"float:left;width:{imgW}px;height:{imgH}px;object-fit:cover;margin:5px 10px 5px 0";
                        if (!string.IsNullOrEmpty(cropCss)) imgStyle += $";{cropCss}";
                        sb.Append($"<img src=\"data:{imgPart.ContentType};base64,{imgBase64}\" style=\"{imgStyle}\">");
                    }
                    catch { }
                }
                _pendingFloatImages = null;
            }

            foreach (var para in txbx.Descendants<Paragraph>())
            {
                RenderParagraphHtml(sb, para);
            }
            sb.Append("</div>");
        }
        else
        {
            // Check for image inside shape
            var blipFill = spPr?.Descendants<A.Blip>().FirstOrDefault();
            if (blipFill?.Embed?.Value != null)
            {
                try
                {
                    var mainPart = _doc.MainDocumentPart;
                    var imagePart = mainPart?.GetPartById(blipFill.Embed.Value) as ImagePart;
                    if (imagePart != null)
                    {
                        using var stream = imagePart.GetStream();
                        using var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        var base64 = Convert.ToBase64String(ms.ToArray());
                        sb.Append($"<img src=\"data:{imagePart.ContentType};base64,{base64}\" style=\"width:100%;height:100%;object-fit:cover\">");
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

    private static string GetRunInlineCss(RunProperties? rProps)
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

        // Color
        var color = rProps.Color?.Val?.Value;
        if (color != null && color != "auto")
            parts.Add($"color:#{color}");

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
        .page {{ background: white; margin: 0 auto 40px; padding: {mT} 0 {mB} 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15); border-radius: 4px;
            min-height: {pageH}; line-height: {lh}; font-size: {sz}; }}
        .doc-header, .doc-footer {{ padding: 0 {lr}; color: #888; font-size: 9pt;
            border-bottom: 1px solid #e0e0e0; margin-bottom: 1em; padding-bottom: 0.5em; }}
        .doc-footer {{ border-bottom: none; border-top: 1px solid #e0e0e0;
            margin-top: 1em; padding-top: 0.5em; margin-bottom: 0; }}
        h1, h2, h3, h4, h5, h6 {{ padding: 0.3em {lr}; line-height: 1.4; }}
        h1 {{ font-size: 22pt; margin-top: 0.5em; margin-bottom: 0.3em; }}
        h2 {{ font-size: 16pt; margin-top: 0.4em; margin-bottom: 0.2em; }}
        h3 {{ font-size: 13pt; margin-top: 0.3em; margin-bottom: 0.2em; }}
        h4 {{ font-size: 11pt; margin-top: 0.2em; margin-bottom: 0.1em; }}
        h5 {{ font-size: 10pt; }} h6 {{ font-size: 9pt; }}
        p {{ padding: 0 {lr}; margin: 0.1em 0; }}
        p.empty {{ margin: 0; padding: 0 {lr}; line-height: 0.8; font-size: 6pt; }}
        a {{ color: #2B579A; }} a:hover {{ color: #1a3c6e; }}
        ul, ol {{ padding-left: 2em; margin: 0.2em 0 0.2em {mL}; }}
        li {{ margin: 0.1em 0; }}
        .equation {{ text-align: center; padding: 0.5em {lr}; overflow-x: auto; }}
        img {{ max-width: 100%; height: auto; }}
        .img-error {{ color: #999; font-style: italic; }}
        table {{ border-collapse: collapse; margin: 0.3em {lr}; font-size: {sz}; width: {tblW}; }}
        .wg {{ margin: 0.3em auto; }}
        .wg p {{ padding: 0; margin: 0.05em 0; }}
        table.borderless {{ border: none; }}
        table.borderless td, table.borderless th {{ border: none; padding: 2px 6px; }}
        th, td {{ border: 1px solid #bbb; padding: 4px 8px; text-align: left; vertical-align: top; }}
        th {{ background: #f0f0f0; font-weight: 600; }}
        .header-row td, .header-row th {{ background: #f0f0f0; font-weight: 600; }}
        hr.page-break {{ border: none; border-top: 2px dashed #ccc; margin: 2em {lr}; }}
        @media print {{ body {{ background: white; padding: 0; }}
            .page {{ box-shadow: none; margin: 0; max-width: none; }}
            hr.page-break {{ page-break-after: always; border: none; margin: 0; }} }}";
    }
}
