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
        public bool RenderingBody { get; set; }

        // CJK line-break tracking: accumulate character widths and insert <br> at Word-compatible positions
        public double LineWidthPt { get; set; }      // available width for current line
        public double LineAccumPt { get; set; }       // accumulated width on current line
        public bool LineBreakEnabled { get; set; }    // whether line-break tracking is active
        public double DefaultFontSizePt { get; set; } // default font size for width estimation

        public void ResetLineForParagraph(double contentWidthPt, double firstLineIndentPt, double defaultSizePt)
        {
            LineWidthPt = contentWidthPt - firstLineIndentPt;
            LineAccumPt = 0;
            LineBreakEnabled = true;
            DefaultFontSizePt = defaultSizePt;
        }

        public void NewLine(double contentWidthPt)
        {
            LineWidthPt = contentWidthPt;
            LineAccumPt = 0;
        }
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

        // Render body into temporary buffer, then split on page breaks
        var maxW = $"max-width:{pgLayout.WidthPt:0.#}pt";
        var bodySb = new StringBuilder();
        _ctx.RenderingBody = true;
        RenderBodyHtml(bodySb, body);
        _ctx.RenderingBody = false;

        // Render header/footer into reusable strings
        var headerSb = new StringBuilder();
        RenderHeaderFooterHtml(headerSb, isHeader: true);
        var headerHtml = headerSb.ToString();

        var footerSb = new StringBuilder();
        RenderHeaderFooterHtml(footerSb, isHeader: false);
        var footerHtml = footerSb.ToString();

        // Render footnotes/endnotes
        var footnotesSb = new StringBuilder();
        RenderFootnotesHtml(footnotesSb);
        var footnotesHtml = footnotesSb.ToString();

        var endnotesSb = new StringBuilder();
        RenderEndnotesHtml(endnotesSb);
        var endnotesHtml = endnotesSb.ToString();

        // Split body content on page breaks into pages
        var bodyContent = bodySb.ToString();
        var pages = bodyContent.Split("<!--PAGE_BREAK-->");

        // Filter out truly empty trailing page (empty string after final page break)
        var pageList = new List<string>();
        for (int i = 0; i < pages.Length; i++)
        {
            var pc = pages[i].Trim();
            if (string.IsNullOrEmpty(pc) && i == pages.Length - 1)
                continue; // Skip completely empty trailing split
            pageList.Add(pc);
        }

        // Detect PAGE field in footer and replace with placeholder
        // Footer typically contains: <span ...>1</span> where "1" is the cached PAGE field value
        // We replace single-digit page numbers in the footer with a placeholder for per-page substitution
        var footerHasPageNum = footerHtml.Contains("PAGE") || !string.IsNullOrEmpty(footerHtml);
        var pageNumPattern = new Regex(@"(<span[^>]*>)\s*\d+\s*(</span>)");
        var footerTemplate = pageNumPattern.Replace(footerHtml, "$1<!--PAGE_NUM-->$2", 1);

        for (int i = 0; i < pageList.Count; i++)
        {
            sb.AppendLine($"<div class=\"page\" style=\"{maxW}\">");
            if (i == 0) sb.Append(headerHtml);
            sb.Append("<div class=\"page-body\">");
            sb.Append(pageList[i]);
            // Place footnotes on the page that contains the footnote reference
            if (!string.IsNullOrEmpty(footnotesHtml) && pageList[i].Contains("fn-ref"))
                sb.Append(footnotesHtml);
            sb.Append("</div>");
            sb.Append(footerTemplate.Replace("<!--PAGE_NUM-->", (i + 1).ToString()));
            sb.AppendLine("</div>");
        }

        // Endnotes at document end (outside page divs)
        sb.Append(endnotesHtml);

        // Auto-pagination script: split overflowing pages and KaTeX rendering
        var bodyHeightPt = pgLayout.HeightPt - pgLayout.MarginTopPt - pgLayout.MarginBottomPt;
        sb.AppendLine("<script>");
        sb.AppendLine("document.addEventListener('DOMContentLoaded',function(){");
        sb.AppendLine("  if(typeof renderMathInElement!=='undefined'){");
        sb.AppendLine("    renderMathInElement(document.body,{delimiters:[");
        sb.AppendLine("      {left:'$$',right:'$$',display:true}");
        sb.AppendLine("    ],throwOnError:false});");
        sb.AppendLine("  }");
        // CJK punctuation compression (~25% per JIS X4051): negative margin on punctuation
        sb.AppendLine("  (function(){");
        sb.AppendLine("  var re=/([\\u3000-\\u303F\\uFF01-\\uFF60\\uFE30-\\uFE4F\\u2014\\u2015\\u2026\\u2018\\u2019\\u201C\\u201D])/;");
        sb.AppendLine("  document.querySelectorAll('.page-body').forEach(function(body){");
        sb.AppendLine("    var w=document.createTreeWalker(body,NodeFilter.SHOW_TEXT);");
        sb.AppendLine("    var nodes=[];while(w.nextNode())nodes.push(w.currentNode);");
        sb.AppendLine("    nodes.forEach(function(nd){");
        sb.AppendLine("      if(!re.test(nd.textContent))return;");
        sb.AppendLine("      var parts=nd.textContent.split(re);");
        sb.AppendLine("      if(parts.length<=1)return;");
        sb.AppendLine("      var frag=document.createDocumentFragment();");
        sb.AppendLine("      for(var i=0;i<parts.length;i++){");
        sb.AppendLine("        if(!parts[i])continue;");
        sb.AppendLine("        if(re.test(parts[i])){");
        sb.AppendLine("          var sp=document.createElement('span');");
        sb.AppendLine("          sp.textContent=parts[i];");
        sb.AppendLine("          sp.style.marginRight='-0.2em';");
        sb.AppendLine("          frag.appendChild(sp);");
        sb.AppendLine("        }else frag.appendChild(document.createTextNode(parts[i]));");
        sb.AppendLine("      }");
        sb.AppendLine("      nd.parentNode.replaceChild(frag,nd);");
        sb.AppendLine("    });");
        sb.AppendLine("  });");
        sb.AppendLine("  })();");
        // Auto-pagination: measure content and split overflowing pages
        sb.AppendLine($"  var maxBodyH={bodyHeightPt:0.#}*96/72;"); // pt to px (96dpi)
        sb.AppendLine("  var ftpl=" + JsStringLiteral(footerTemplate) + ";");
        sb.AppendLine(@"
  function paginate(){
    var pages=document.querySelectorAll('.page');
    for(var pi=0;pi<pages.length;pi++){
      var page=pages[pi];
      var body=page.querySelector('.page-body');
      if(!body)continue;
      if(body.scrollHeight<=maxBodyH+2)continue;
      // Find first child that overflows
      var children=Array.from(body.children);
      var splitIdx=-1;
      for(var ci=0;ci<children.length;ci++){
        var bot=children[ci].offsetTop+children[ci].offsetHeight-body.offsetTop;
        if(bot>maxBodyH){splitIdx=ci;break;}
      }
      if(splitIdx<=0)continue;
      // Create new page
      var np=document.createElement('div');
      np.className='page';
      np.style.cssText=page.style.cssText;
      var nb=document.createElement('div');
      nb.className='page-body';
      // Move overflow children to new page
      while(splitIdx<children.length){
        nb.appendChild(children[splitIdx]);
        splitIdx++;
      }
      np.appendChild(nb);
      // Clone footer into new page, replacing page number placeholder
      var nf=document.createElement('div');
      nf.innerHTML=ftpl.replace('<!--PAGE_NUM-->',(pi+2).toString());
      if(nf.firstChild)np.appendChild(nf.firstChild);
      page.after(np);
    }
    // Renumber pages
    var allPages=document.querySelectorAll('.page');
    var total=allPages.length;
    allPages.forEach(function(p,i){
      var nums=p.querySelectorAll('.page-num');
      nums.forEach(function(n){n.textContent=(i+1);});
      // Also handle PAGE_NUM placeholders in footer spans
      var footer=p.querySelector('.doc-footer');
      if(footer){
        var spans=footer.querySelectorAll('span');
        spans.forEach(function(s){
          if(s.textContent.trim().match(/^\d+$/)){
            s.textContent=(i+1);
          }
        });
      }
    });
    // Recurse in case new pages also overflow
    var again=false;
    document.querySelectorAll('.page').forEach(function(p){
      var b=p.querySelector('.page-body');
      if(b&&b.scrollHeight>maxBodyH+2)again=true;
    });
    if(again)setTimeout(paginate,0);
  }
  setTimeout(paginate,100);
");
        sb.AppendLine("});");
        sb.AppendLine("</script>");

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }

    // ==================== Page Layout + Doc Defaults from OOXML ====================

    private record PageLayout(double WidthCm, double HeightCm,
        double MarginTopCm, double MarginBottomCm, double MarginLeftCm, double MarginRightCm,
        double HeaderDistanceCm, double FooterDistanceCm,
        double WidthPt, double HeightPt,
        double MarginTopPt, double MarginBottomPt, double MarginLeftPt, double MarginRightPt,
        double HeaderDistancePt, double FooterDistancePt);

    private PageLayout GetPageLayout()
    {
        if (_ctx?.CachedPageLayout != null) return _ctx.CachedPageLayout;
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        var pgSz = sectPr?.GetFirstChild<PageSize>();
        var pgMar = sectPr?.GetFirstChild<PageMargin>();
        const double c = 2.54 / 1440.0; // twips → cm
        const double p = 1.0 / 20.0;    // twips → pt (exact)
        var wTwips = (double)(pgSz?.Width?.Value ?? 11906);
        var hTwips = (double)(pgSz?.Height?.Value ?? 16838);
        var tTwips = (double)(pgMar?.Top?.Value ?? 1440);
        var bTwips = (double)(pgMar?.Bottom?.Value ?? 1440);
        var lTwips = (double)(pgMar?.Left?.Value ?? 1440u);
        var rTwips = (double)(pgMar?.Right?.Value ?? 1440u);
        var hdTwips = (double)(pgMar?.Header?.Value ?? 851u);
        var fdTwips = (double)(pgMar?.Footer?.Value ?? 992u);
        var result = new PageLayout(
            wTwips * c, hTwips * c, tTwips * c, bTwips * c, lTwips * c, rTwips * c, hdTwips * c, fdTwips * c,
            wTwips * p, hTwips * p, tTwips * p, bTwips * p, lTwips * p, rTwips * p, hdTwips * p, fdTwips * p);
        if (_ctx != null) _ctx.CachedPageLayout = result;
        return result;
    }

    private record DocDef(string Font, double SizePt, double LineHeight, string Color, double GridLinePitchPt);

    private DocDef ReadDocDefaults()
    {
        var defs = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
        var rPr = defs?.RunPropertiesDefault?.RunPropertiesBaseStyle;

        // Find default paragraph style (Normal) for fallback
        var defaultStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
            ?.Elements<Style>().FirstOrDefault(s => s.Default?.Value == true && s.Type?.Value == StyleValues.Paragraph);
        var defaultRPr = defaultStyle?.StyleRunProperties;

        // Font: docDefaults rFonts → Normal style rFonts → theme minor font → fallback
        var fonts = rPr?.RunFonts;
        var font = NonEmpty(fonts?.EastAsia?.Value) ?? NonEmpty(fonts?.Ascii?.Value) ?? NonEmpty(fonts?.HighAnsi?.Value);
        if (font == null)
        {
            var nFonts = defaultRPr?.RunFonts;
            font = NonEmpty(nFonts?.EastAsia?.Value) ?? NonEmpty(nFonts?.Ascii?.Value) ?? NonEmpty(nFonts?.HighAnsi?.Value);
        }
        if (font == null)
        {
            var minor = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont;
            font = NonEmpty(minor?.EastAsianFont?.Typeface) ?? NonEmpty(minor?.LatinFont?.Typeface);
        }

        // Size: docDefaults → Normal style → fallback (half-points → pt)
        double sizePt = 0;
        if (rPr?.FontSize?.Val?.Value is string sz && int.TryParse(sz, out var hp))
            sizePt = hp / 2.0;
        if (sizePt == 0 && defaultRPr?.FontSize?.Val?.Value is string nsz && int.TryParse(nsz, out var nhp))
            sizePt = nhp / 2.0;
        if (sizePt == 0) sizePt = 10.5;

        // Line spacing: docDefaults pPrDefault → Normal style pPr → fallback
        double lineH = 0;
        var sp = defs?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines;
        if (sp?.Line?.Value is string lv && int.TryParse(lv, out var lvi) && sp.LineRule?.InnerText is "auto" or null)
            lineH = lvi / 240.0;
        if (lineH == 0)
        {
            var nsp = defaultStyle?.StyleParagraphProperties?.SpacingBetweenLines;
            if (nsp?.Line?.Value is string nlv && int.TryParse(nlv, out var nlvi) && nsp.LineRule?.InnerText is "auto" or null)
                lineH = nlvi / 240.0;
        }
        if (lineH == 0) lineH = 1.0; // Word default single-line spacing

        // docGrid linePitch — controls CJK snap-to-grid line spacing (twips → pt)
        double gridLinePitchPt = 0;
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        var docGrid = sectPr?.GetFirstChild<DocGrid>();
        if (docGrid?.Type?.Value == DocGridValues.Lines || docGrid?.Type?.Value == DocGridValues.LinesAndChars)
        {
            if (docGrid.LinePitch?.Value is int lp && lp > 0)
                gridLinePitchPt = lp / 20.0; // twips to pt
        }

        // Default text color: docDefaults → theme dk1
        var color = "#000000";
        var cv = rPr?.Color?.Val?.Value;
        if (cv != null && cv != "auto") color = $"#{cv}";
        else if (GetThemeColors().TryGetValue("dk1", out var dk1)) color = $"#{dk1}";

        return new DocDef(font ?? "Calibri", sizePt, lineH, color, gridLinePitchPt);
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
        int? currentNumId = null; // track numId for cross-numId nesting
        var numIdLevelOffset = new Dictionary<int, int>(); // numId → effective ilvl offset for cross-numId nesting
        var olCountPerLevel = new Dictionary<int, int>(); // ilvl → running <ol> item count for `start` attribute
        bool pendingLiClose = false; // defer </li> to allow nested lists inside
        bool inMultiColumn = false; // track whether we're inside a multi-column div

        // Pre-scan: build a map of section column counts from inline sectPr breaks
        // The last section's cols come from the body sectPr
        var bodySectPr = body.GetFirstChild<SectionProperties>();
        var bodyColCount = GetSectionColumnCount(bodySectPr);

        for (int ei = 0; ei < elements.Count; ei++)
        {
            var element = elements[ei];

            // Check for inline section break (sectPr inside paragraph pPr) — handle column changes
            if (element is Paragraph sectPara && sectPara.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
            {
                var nextCols = GetNextSectionColumnCount(elements, ei, bodyColCount);
                if (nextCols > 1 && !inMultiColumn)
                {
                    sb.AppendLine($"<div style=\"column-count:{nextCols};column-gap:36pt\">");
                    inMultiColumn = true;
                }
                else if (nextCols <= 1 && inMultiColumn)
                {
                    sb.AppendLine("</div>");
                    inMultiColumn = false;
                }
            }

            if (element is Paragraph para)
            {
                // Check for pageBreakBefore (direct or from style) — insert page break marker
                var pgBB = para.ParagraphProperties?.PageBreakBefore;
                if (pgBB == null)
                {
                    var sid = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    pgBB = ResolvePageBreakBeforeFromStyle(sid);
                }
                if (pgBB != null && pgBB.Val?.Value != false)
                    sb.Append("<!--PAGE_BREAK-->");

                // Check for display equation
                var oMathPara = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathPara != null)
                {
                    CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                    var latex = FormulaParser.ToLatex(oMathPara);
                    sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
                    continue;
                }

                // Check if this is a list item
                var listStyle = GetParagraphListStyle(para);
                if (listStyle != null)
                {
                    var ilvl = para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
                    var numId = para.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value ?? 0;
                    var tag = listStyle == "bullet" ? "ul" : "ol";

                    // Detect cross-numId nesting: when numId changes but we're already in a list,
                    // compare indentation to decide if the new list is nested or a sibling.
                    if (currentNumId != null && numId != currentNumId && listStack.Count > 0
                        && !numIdLevelOffset.ContainsKey(numId))
                    {
                        var curIndent = GetListLevelIndent(currentNumId.Value, currentListLevel);
                        var newIndent = GetListLevelIndent(numId, ilvl);
                        if (newIndent > curIndent)
                        {
                            numIdLevelOffset[numId] = currentListLevel + 1 - ilvl;
                        }
                    }
                    // Apply stored level offset for this numId
                    if (numIdLevelOffset.TryGetValue(numId, out var offset))
                        ilvl += offset;

                    // Close pending </li> from previous item — but only if NOT nesting deeper
                    if (pendingLiClose && ilvl + 1 <= listStack.Count)
                    {
                        sb.AppendLine("</li>");
                        pendingLiClose = false;
                    }

                    // Adjust nesting (close deeper levels: </ol></li> or </ul></li>)
                    while (listStack.Count > ilvl + 1)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        sb.AppendLine("</li>");
                    }
                    if (pendingLiClose)
                    {
                        // Still pending — we're going deeper, so leave parent <li> open
                        pendingLiClose = false;
                    }
                    while (listStack.Count < ilvl + 1)
                    {
                        // For <ol>, use start attribute to continue numbering when reopening
                        if (tag == "ol" && olCountPerLevel.TryGetValue(ilvl, out var prevCount) && prevCount > 0)
                            sb.AppendLine($"<{tag} start=\"{prevCount + 1}\">");
                        else
                            sb.AppendLine($"<{tag}>");
                        listStack.Push(tag);
                    }
                    // If same level but different list type, swap
                    if (listStack.Count > 0 && listStack.Peek() != tag)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        if (tag == "ol" && olCountPerLevel.TryGetValue(ilvl, out var pc) && pc > 0)
                            sb.AppendLine($"<{tag} start=\"{pc + 1}\">");
                        else
                            sb.AppendLine($"<{tag}>");
                        listStack.Push(tag);
                    }
                    // Track <ol> item count per level
                    if (tag == "ol")
                        olCountPerLevel[ilvl] = olCountPerLevel.GetValueOrDefault(ilvl, 0) + 1;

                    currentListType = listStyle;
                    currentListLevel = ilvl;
                    currentNumId = numId;
                    sb.Append("<li");
                    var paraStyle = GetParagraphInlineCss(para, isListItem: true);
                    if (!string.IsNullOrEmpty(paraStyle))
                        sb.Append($" style=\"{paraStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    pendingLiClose = true; // defer </li> in case next item nests
                    continue;
                }

                // Not a list — close any open lists
                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);

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

                    // Skip empty section-break paragraphs (they only carry sectPr, no visual content)
                    if (runs.Count == 0 && string.IsNullOrWhiteSpace(text)
                        && para.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
                    {
                        continue;
                    }

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
                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                var latex = FormulaParser.ToLatex(element);
                sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
            }
            else if (element is Table table)
            {
                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                RenderTableHtml(sb, table);
            }
            else if (element is SectionProperties)
            {
                // Skip — section properties are not visual content
            }

        }

        if (inMultiColumn) sb.AppendLine("</div>");
        CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
    }

    private static void CloseAllLists(StringBuilder sb, Stack<string> listStack, ref string? currentListType, ref bool pendingLiClose)
    {
        if (pendingLiClose) { sb.AppendLine("</li>"); pendingLiClose = false; }
        while (listStack.Count > 0)
        {
            sb.AppendLine($"</{listStack.Pop()}>");
            if (listStack.Count > 0)
                sb.AppendLine("</li>");
        }
        currentListType = null;
    }

    /// <summary>Get the column count from a section properties element.</summary>
    private static int GetSectionColumnCount(SectionProperties? sectPr)
    {
        var cols = sectPr?.GetFirstChild<Columns>();
        var num = cols?.ColumnCount?.Value;
        if (num != null && num > 1) return num.Value;
        return 1;
    }

    /// <summary>Get the column count for the next section after a given element index.</summary>
    private static int GetNextSectionColumnCount(List<OpenXmlElement> elements, int currentIdx, int bodyColCount)
    {
        // Look forward for the next inline sectPr; if none found, use body sectPr cols
        for (int i = currentIdx + 1; i < elements.Count; i++)
        {
            if (elements[i] is Paragraph p && p.ParagraphProperties?.GetFirstChild<SectionProperties>() is SectionProperties sect)
                return GetSectionColumnCount(sect);
        }
        return bodyColCount;
    }

    /// <summary>Get the left indent (in twips) for a numbering level definition.</summary>
    private int GetListLevelIndent(int numId, int ilvl)
    {
        var numPart = _doc.MainDocumentPart?.NumberingDefinitionsPart;
        if (numPart == null) return 0;
        var numbering = numPart.Numbering;
        var numInst = numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        var absId = numInst?.AbstractNumId?.Val?.Value;
        if (absId == null) return 0;
        var absDef = numbering?.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == absId);
        var lvl = absDef?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
        var indent = lvl?.PreviousParagraphProperties?.Indentation;
        if (indent?.Left?.Value is string left && int.TryParse(left, out var twips))
            return twips;
        return 0;
    }
}
