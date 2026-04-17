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
        public List<(string markerId, string imgHtml)> TopAnchoredImages { get; } = new();
        public PageLayout? CachedPageLayout { get; set; }
        public bool RenderingBody { get; set; }

        // CJK line-break tracking: accumulate character widths and insert <br> at Word-compatible positions
        public double LineWidthPt { get; set; }      // available width for current line
        public double LineAccumPt { get; set; }       // accumulated width on current line
        public bool LineBreakEnabled { get; set; }    // whether line-break tracking is active
        public double DefaultFontSizePt { get; set; } // default font size for width estimation

        // Tab positioning: count tabs seen in current paragraph to look up Nth tab stop.
        // Reset per paragraph in RenderParagraphContentHtml.
        public int CurrentParagraphTabIndex { get; set; }

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

    /// <summary>Cached EastAsia language from themeFontLang/docDefaults (e.g. "zh-CN", "ja-JP", "ko-KR").</summary>
    private string? _eastAsiaLang;

    /// <summary>CJK font resolved from theme's supplemental font list (e.g. "Microsoft YaHei" for Hans).</summary>
    private string? _themeCjkFont;

    /// <summary>
    /// Generate a self-contained HTML file that previews the Word document
    /// with formatting, tables, images, and lists.
    /// </summary>
    public string ViewAsHtml(string? pageFilter = null)
    {
        _ctx = new HtmlRenderContext();
        ResolveThemeCjkFont();
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
        // Load document fonts: @font-face with metric overrides for all fonts,
        // Google Fonts only for non-system fonts.
        var docFonts = CollectDocumentFonts();
        if (docFonts.Count > 0)
        {
            var fontFaces = ResolveLocalFontFaces(docFonts);
            if (fontFaces.Length > 0)
            {
                sb.AppendLine("<style>");
                sb.Append(fontFaces);
                sb.AppendLine("</style>");
            }
            // Filter out system fonts for Google Fonts loading (they're already local)
            var googleFonts = docFonts.Where(f =>
                !f.Equals("Arial", StringComparison.OrdinalIgnoreCase)
                && !f.Equals("Times New Roman", StringComparison.OrdinalIgnoreCase)
                && !f.Equals("Tahoma", StringComparison.OrdinalIgnoreCase)
                && !f.Equals("Courier New", StringComparison.OrdinalIgnoreCase)
                && !f.StartsWith("Symbol") && !f.StartsWith("Wingding")).ToList();
            if (googleFonts.Count > 0)
            {
                var families = string.Join("&", googleFonts.Select(f =>
                    $"family={f.Replace(' ', '+')}:ital,wght@0,400;0,700;1,400;1,700"));
                sb.AppendLine($"<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?{families}&display=swap\" onerror=\"this.remove()\">");
            }
        }
        // KaTeX for math rendering (graceful degradation: shows raw LaTeX when offline)
        sb.AppendLine("<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css\" onerror=\"this.remove()\">");
        sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js\" onerror=\"document.querySelectorAll('.katex-formula').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666'})\"></script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        // Render body into temporary buffer, then split on page breaks
        var maxW = $"width:{pgLayout.WidthPt:0.#}pt";
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

        var bodyContent = bodySb.ToString();

        // Split body content on page breaks into pages
        var pages = bodyContent.Split("<!--PAGE_BREAK-->");

        // Filter out truly empty trailing page (empty string after final page break)
        // Also relocate top-anchored images to the start of their page
        var markerMap = _ctx.TopAnchoredImages.ToDictionary(t => $"<!--{t.markerId}-->", t => t.imgHtml);
        var pageList = new List<string>();
        for (int i = 0; i < pages.Length; i++)
        {
            var pc = pages[i].Trim();
            if (string.IsNullOrEmpty(pc) && i == pages.Length - 1)
                continue; // Skip completely empty trailing split
            // Move top-anchored images to page start
            if (markerMap.Count > 0)
            {
                var prepend = new StringBuilder();
                foreach (var (marker, imgHtml) in markerMap)
                {
                    if (pc.Contains(marker))
                    {
                        prepend.Append(imgHtml);
                        pc = pc.Replace(marker, "");
                    }
                }
                if (prepend.Length > 0)
                    pc = prepend.ToString() + pc;
            }
            pageList.Add(pc);
        }

        // Parse page filter (e.g. "1", "2-5", "1,3,5", "2-4,7")
        HashSet<int>? requestedPages = null;
        int totalServerPages = pageList.Count;
        if (!string.IsNullOrWhiteSpace(pageFilter))
        {
            requestedPages = new HashSet<int>();
            foreach (var part in pageFilter.Split(','))
            {
                var trimmed = part.Trim();
                if (trimmed.Contains('-'))
                {
                    var range = trimmed.Split('-', 2);
                    if (int.TryParse(range[0].Trim(), out var from) && int.TryParse(range[1].Trim(), out var to))
                        for (int p = from; p <= to; p++) requestedPages.Add(p);
                }
                else if (int.TryParse(trimmed, out var num))
                    requestedPages.Add(num);
            }
        }

        // Detect PAGE field in footer and replace with placeholder
        // Footer typically contains: <span ...>1</span> where "1" is the cached PAGE field value
        // We replace single-digit page numbers in the footer with a placeholder for per-page substitution
        var footerHasPageNum = footerHtml.Contains("PAGE") || !string.IsNullOrEmpty(footerHtml);
        var pageNumPattern = new Regex(@"(<span[^>]*>)\s*\d+\s*(</span>)");
        var footerTemplate = pageNumPattern.Replace(footerHtml, "$1<!--PAGE_NUM-->$2", 1);
        // Second single-digit match = NUMPAGES in "Page X of Y"
        var footerTemplateWithTotal = pageNumPattern.Replace(footerTemplate, "$1<!--NUM_PAGES-->$2", 1);
        footerTemplate = footerTemplateWithTotal;

        // Section-level multi-column layout: w:cols num=N sep=true
        var sectCols = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>()?.GetFirstChild<Columns>();
        var colCount = sectCols?.ColumnCount?.Value ?? 1;
        var colSep = sectCols?.Separator?.Value == true;
        var colSpacing = sectCols?.Space?.Value;
        var colBodyStyle = colCount > 1
            ? $" style=\"column-count:{colCount}"
                + (colSep ? ";column-rule:1px solid #000" : "")
                + (int.TryParse(colSpacing, out var csp) && csp > 0 ? $";column-gap:{csp / 20.0:0.##}pt" : "")
                + "\""
            : "";

        for (int i = 0; i < pageList.Count; i++)
        {
            sb.AppendLine($"<div class=\"page-wrapper\" data-section=\"{i + 1}\">");
            sb.AppendLine($"<div class=\"page\" data-page=\"{i + 1}\" style=\"{maxW}\">");
            if (i == 0) sb.Append(headerHtml);
            sb.Append($"<div class=\"page-body\"{colBodyStyle}>");
            sb.Append(pageList[i]);
            // Place footnotes on the page that contains the footnote reference
            if (!string.IsNullOrEmpty(footnotesHtml) && pageList[i].Contains("fn-ref"))
                sb.Append(footnotesHtml);
            // Place endnotes on the last page
            if (i == pageList.Count - 1 && !string.IsNullOrEmpty(endnotesHtml))
                sb.Append(endnotesHtml);
            sb.Append("</div>");
            sb.Append(footerTemplate
                .Replace("<!--PAGE_NUM-->", (i + 1).ToString())
                .Replace("<!--NUM_PAGES-->", pageList.Count.ToString()));
            sb.AppendLine("</div>");
            sb.AppendLine("</div>");
        }

        // Auto-pagination script: split overflowing pages and KaTeX rendering
        var bodyHeightPt = pgLayout.HeightPt - pgLayout.MarginTopPt - pgLayout.MarginBottomPt;
        sb.AppendLine("<script>");
        sb.AppendLine("function _wordInit(){");
        sb.AppendLine("  if(typeof katex!=='undefined'){");
        sb.AppendLine("    document.querySelectorAll('.katex-formula:not(.katex-rendered)').forEach(function(el){");
        sb.AppendLine("      try{katex.render(el.dataset.formula,el,{throwOnError:false,displayMode:!!el.dataset.display});}catch(e){el.textContent=el.dataset.formula+' (Error: '+e.message+'. See https://katex.org/docs/supported.html for supported syntax.)';}");
        sb.AppendLine("      el.classList.add('katex-rendered');");
        sb.AppendLine("    });");
        sb.AppendLine("  }else{");
        sb.AppendLine("    document.querySelectorAll('.katex-formula:not(.katex-rendered)').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666';});");
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
      // Reserve space for footnotes at page bottom (like Word does)
      var fnEl=body.querySelector('.footnotes');
      var fnH=fnEl?fnEl.offsetHeight:0;
      var availH=maxBodyH-fnH;
      // Check if content (excluding footnotes) exceeds available space
      var contentH=0;
      Array.from(body.children).forEach(function(c){
        if(c.classList.contains('footnotes'))return;
        var b=c.offsetTop+c.offsetHeight-body.offsetTop;
        if(b>contentH)contentH=b;
      });
      if(contentH<=availH+2)continue;
      // Find first child that overflows available space
      var children=Array.from(body.children);
      var splitIdx=-1;
      for(var ci=0;ci<children.length;ci++){
        if(children[ci].classList.contains('footnotes'))continue;
        var bot=children[ci].offsetTop+children[ci].offsetHeight-body.offsetTop;
        if(bot>availH){splitIdx=ci;break;}
      }
      if(splitIdx<0)continue;
      // When the first child itself exceeds page height, keep it on this
      // page and split after, so the oversized element is not silently
      // dropped by being moved to a new (still-oversized) page.
      if(splitIdx===0)splitIdx=1;
      // Collect movable children from splitIdx onward (skip footnotes — they
      // stay on the reference page). If nothing is movable, the page is
      // irreducibly oversized and we let it overflow gracefully instead of
      // producing an empty follow-up page.
      var toMove=[];
      for(var mi=splitIdx;mi<children.length;mi++){
        if(!children[mi].classList.contains('footnotes'))toMove.push(children[mi]);
      }
      if(toMove.length===0)continue;
      // Create new page wrapped in page-wrapper
      var nw=document.createElement('div');
      nw.className='page-wrapper';
      var np=document.createElement('div');
      np.className='page';
      np.style.cssText=page.style.cssText;
      var nb=document.createElement('div');
      nb.className='page-body';
      for(var mi=0;mi<toMove.length;mi++){
        nb.appendChild(toMove[mi]);
      }
      np.appendChild(nb);
      // Clone footer into new page
      var nf=document.createElement('div');
      nf.innerHTML=ftpl.replace('<!--PAGE_NUM-->',(pi+2).toString());
      if(nf.firstChild)np.appendChild(nf.firstChild);
      nw.appendChild(np);
      var parentWrapper=page.closest('.page-wrapper');
      if(parentWrapper)parentWrapper.after(nw);
      else page.after(nw);
    }
    // Renumber pages
    var allPages=document.querySelectorAll('.page');
    allPages.forEach(function(p,i){
      var nums=p.querySelectorAll('.page-num');
      nums.forEach(function(n){n.textContent=(i+1);});
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
    // Recurse in case new pages also overflow. A page is only eligible for
    // another split when it has more than one visible child — otherwise the
    // single element is irreducible and we would recurse forever.
    var again=false;
    document.querySelectorAll('.page').forEach(function(p){
      var b=p.querySelector('.page-body');
      if(!b)return;
      var f=b.querySelector('.footnotes');
      var fh=f?f.offsetHeight:0;
      var ch=0;
      var visibleCount=0;
      Array.from(b.children).forEach(function(c){
        if(c.classList.contains('footnotes'))return;
        var bt=c.offsetTop+c.offsetHeight-b.offsetTop;
        if(bt>ch)ch=bt;
        if(c.offsetHeight>0)visibleCount++;
      });
      if(ch>maxBodyH-fh+2 && visibleCount>1)again=true;
    });
    if(again)setTimeout(paginate,0);
    else{setTimeout(positionFootnotes,0);setTimeout(applyPageFilter,0);setTimeout(function(){scalePages(false);},0);}
  }
  function positionFootnotes(){
    document.querySelectorAll('.page').forEach(function(page){
      var body=page.querySelector('.page-body');
      if(!body)return;
      var fn=body.querySelector('.footnotes');
      if(!fn)return;
      // Calculate space between last content element and page bottom
      var lastBot=0;
      Array.from(body.children).forEach(function(c){
        if(c===fn)return;
        var b=c.offsetTop+c.offsetHeight-body.offsetTop;
        if(b>lastBot)lastBot=b;
      });
      var gap=maxBodyH-lastBot-fn.offsetHeight;
      if(gap>0)fn.style.marginTop=gap+'px';
    });
  }
  function applyPageFilter(){
    var rf=window._requestedPages;
    if(!rf||rf.length===0)return;
    var rSet=new Set(rf);
    document.querySelectorAll('.page').forEach(function(p,i){
      if(!rSet.has(i+1))p.style.display='none';
    });
  }
  function renderNewContent(){
    if(typeof katex!=='undefined'){
      document.querySelectorAll('.katex-formula:not(.katex-rendered)').forEach(function(el){
        try{katex.render(el.dataset.formula,el,{throwOnError:false,displayMode:!!el.dataset.display});}catch(e){el.textContent=el.dataset.formula;}
        el.classList.add('katex-rendered');
      });
    }
    // CJK punctuation compression on new content
    var cjkRe=/([\u3000-\u303F\uFF01-\uFF60\uFE30-\uFE4F\u2014\u2015\u2026\u2018\u2019\u201C\u201D])/;
    document.querySelectorAll('.page-body').forEach(function(body){
      var tw=document.createTreeWalker(body,NodeFilter.SHOW_TEXT);
      var nodes=[];while(tw.nextNode()){var n=tw.currentNode;if(!n.parentNode||!n.parentNode.classList||!n.parentNode.classList.contains('cjk-done'))nodes.push(n);}
      nodes.forEach(function(nd){
        if(!cjkRe.test(nd.textContent))return;
        var parts=nd.textContent.split(cjkRe);
        if(parts.length<=1)return;
        var frag=document.createDocumentFragment();
        for(var i=0;i<parts.length;i++){
          if(!parts[i])continue;
          if(cjkRe.test(parts[i])){var sp=document.createElement('span');sp.textContent=parts[i];sp.style.marginRight='-0.2em';sp.classList.add('cjk-done');frag.appendChild(sp);}
          else frag.appendChild(document.createTextNode(parts[i]));
        }
        nd.parentNode.replaceChild(frag,nd);
      });
    });
  }
  window._wordPaginate=function(){renderNewContent();setTimeout(paginate,0);};
  setTimeout(paginate,100);
");
        // Responsive scaling: shrink pages to fit viewport (like PPT's scaleSlides)
        sb.AppendLine(@"  function scalePages(animate){
    var bs=getComputedStyle(document.body);
    var availW=document.body.clientWidth-parseFloat(bs.paddingLeft)-parseFloat(bs.paddingRight);
    if(!animate){
      document.querySelectorAll('.page-wrapper,.page').forEach(function(el){el.style.transition='none';});
    }
    document.querySelectorAll('.page-wrapper').forEach(function(wrapper){
      var page=wrapper.querySelector('.page');
      if(!page||page.style.display==='none')return;
      var pageW=page.offsetWidth;
      var pageH=page.offsetHeight;
      var s=Math.min(availW/pageW,1);
      page.style.transform='scale('+s+')';
      wrapper.style.height=(pageH*s)+'px';
      wrapper.style.width=(pageW*s)+'px';
    });
    if(!animate){
      document.body.offsetHeight;
      document.querySelectorAll('.page-wrapper,.page').forEach(function(el){el.style.transition='';});
    }
    if(window._pendingScrollTo){
      var _sel=window._pendingScrollTo;
      var _beh=window._pendingScrollBehavior||'smooth';
      window._pendingScrollTo=null;
      window._pendingScrollBehavior=null;
      var _t;
      if(_sel==='_last_page'){var _lb=document.querySelector('.page-wrapper:last-of-type .page-body');if(_lb){var _ck=Array.from(_lb.children).filter(function(c){return !c.classList.contains('footnotes')&&c.style.display!=='none'&&c.offsetHeight>0;});_t=_ck[_ck.length-1]||_lb;}if(!_t){var _ap=document.querySelectorAll('.page');_t=_ap[_ap.length-1];}}
      else{_t=document.querySelector(_sel);if(!_t){var _ap=document.querySelectorAll('.page');_t=_ap[_ap.length-1];}}
      if(_t)_t.scrollIntoView({behavior:_beh,block:'center'});
    }
    var _frz=document.getElementById('_sse_freeze');
    if(_frz)_frz.remove();
  }
  var _resizeTimer;
  window.addEventListener('resize',function(){
    clearTimeout(_resizeTimer);
    _resizeTimer=setTimeout(function(){scalePages(true);},100);
  });");
        // Pass requested pages to JS for post-pagination filtering
        if (requestedPages != null && requestedPages.Count > 0)
            sb.AppendLine($"  window._requestedPages=[{string.Join(",", requestedPages)}];");
        sb.AppendLine("}");
        sb.AppendLine("if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',_wordInit);");
        sb.AppendLine("else _wordInit();");
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

    private record DocDef(string Font, double SizePt, double LineHeight, string Color, double GridLinePitchPt,
        double SpaceAfterPt = 0, string DefaultAlign = "left");

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
        if (sizePt == 0) sizePt = 10.0; // OOXML spec default: 20 half-points = 10pt

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
        if (lineH == 0) lineH = 1.0; // OOXML default single-line spacing

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

        // Space after: Normal style pPr → docDefaults pPr → 0
        double spaceAfterPt = 0;
        var defSp = defaultStyle?.StyleParagraphProperties?.SpacingBetweenLines;
        var defSpAfter = defaultStyle?.StyleParagraphProperties?.GetFirstChild<SpacingBetweenLines>() != null
            ? defaultStyle.StyleParagraphProperties.SpacingBetweenLines?.After?.Value : null;
        if (defSpAfter == null)
            defSpAfter = defs?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines?.After?.Value;
        if (defSpAfter != null && int.TryParse(defSpAfter, out var saVal))
            spaceAfterPt = saVal / 20.0; // twips to pt

        // Default paragraph alignment: Normal style jc → left
        var defaultAlign = "left";
        var jc = defaultStyle?.StyleParagraphProperties?.Justification?.Val;
        if (jc != null)
        {
            defaultAlign = jc.InnerText switch
            {
                "center" => "center",
                "right" or "end" => "right",
                "both" or "distribute" => "justify",
                _ => "left"
            };
        }

        return new DocDef(font ?? "Calibri", sizePt, lineH, color, gridLinePitchPt, spaceAfterPt, defaultAlign);
    }

    /// <summary>Collect all distinct font names from document body, styles, and theme.</summary>
    private HashSet<string> CollectDocumentFonts()
    {
        var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // From styles
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles != null)
            foreach (var rf in styles.Descendants<RunFonts>())
            {
                if (!string.IsNullOrEmpty(rf.Ascii?.Value)) fonts.Add(rf.Ascii.Value);
                if (!string.IsNullOrEmpty(rf.HighAnsi?.Value)) fonts.Add(rf.HighAnsi.Value);
                if (!string.IsNullOrEmpty(rf.EastAsia?.Value)) fonts.Add(rf.EastAsia.Value);
            }
        // From document body
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body != null)
            foreach (var rf in body.Descendants<RunFonts>())
            {
                if (!string.IsNullOrEmpty(rf.Ascii?.Value)) fonts.Add(rf.Ascii.Value);
                if (!string.IsNullOrEmpty(rf.HighAnsi?.Value)) fonts.Add(rf.HighAnsi.Value);
            }
        // From theme
        var theme = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
        var majFont = theme?.MajorFont?.LatinFont?.Typeface?.Value;
        if (!string.IsNullOrEmpty(majFont)) fonts.Add(majFont);
        var minFont = theme?.MinorFont?.LatinFont?.Typeface?.Value;
        if (!string.IsNullOrEmpty(minFont)) fonts.Add(minFont);
        // Remove fonts that have no usable @font-face (symbols, wingdings)
        fonts.RemoveWhere(f => f.StartsWith("Symbol") || f.StartsWith("Wingding"));
        return fonts;
    }

    /// <summary>
    /// Resolve CJK font from theme supplemental font list (like libra's ThemeHandler).
    /// Also reads themeFontLang/eastAsia language for fallback.
    /// </summary>
    private void ResolveThemeCjkFont()
    {
        // 1. Read eastAsia language from settings (w:themeFontLang) or docDefaults (w:lang)
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var themeFontLang = settings?.Descendants<DocumentFormat.OpenXml.Wordprocessing.ThemeFontLanguages>().FirstOrDefault();
        _eastAsiaLang = themeFontLang?.EastAsia?.Value;

        // Also check docDefaults for w:lang eastAsia
        if (_eastAsiaLang == null)
        {
            var docDefLang = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle
                ?.Languages;
            _eastAsiaLang = docDefLang?.EastAsia?.Value;
        }

        // 2. Read CJK font from theme supplemental font list
        var fontScheme = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
        if (fontScheme == null) return;

        // Map eastAsia language to OOXML script tag
        var scriptTag = (_eastAsiaLang?.ToLowerInvariant()) switch
        {
            string l when l.StartsWith("ja") => "Jpan",
            string l when l.StartsWith("ko") => "Hang",
            string l when l.StartsWith("zh") && l.Contains("tw") => "Hant",
            string l when l.StartsWith("zh") && l.Contains("hk") => "Hant",
            _ => "Hans" // default to simplified Chinese
        };

        // Search supplemental font list in minorFont (body text), then majorFont (headings)
        foreach (var fontCollection in new OpenXmlElement?[] { fontScheme.MinorFont, fontScheme.MajorFont })
        {
            if (fontCollection == null) continue;
            foreach (var sf in fontCollection.Descendants<A.SupplementalFont>())
            {
                if (sf.Script?.Value == scriptTag && !string.IsNullOrEmpty(sf.Typeface?.Value))
                {
                    _themeCjkFont = sf.Typeface.Value;
                    return;
                }
            }
        }

        // Fallback: use EastAsianFont from theme
        var eaFont = fontScheme.MinorFont?.Descendants<A.EastAsianFont>().FirstOrDefault()?.Typeface?.Value
            ?? fontScheme.MajorFont?.Descendants<A.EastAsianFont>().FirstOrDefault()?.Typeface?.Value;
        if (!string.IsNullOrEmpty(eaFont))
            _themeCjkFont = eaFont;
    }

    /// <summary>Generate @font-face rules with local() for document fonts.
    /// Includes ascent-override/descent-override/line-gap-override to force
    /// the browser to use OS/2 winAscent+winDescent metrics instead of
    /// the browser's default (which may include hhea lineGap).</summary>
    private static string ResolveLocalFontFaces(HashSet<string> docFonts)
    {
        var sb = new StringBuilder();
        foreach (var font in docFonts)
        {
            var (ascentPct, descentPct) = FontMetricsReader.GetAscentDescentOverride(font);
            var overrides = ascentPct > 0
                ? $" ascent-override: {ascentPct:0.##}%; descent-override: {descentPct:0.##}%; line-gap-override: 0%;"
                : "";
            sb.AppendLine($"@font-face {{ font-family: '{font}'; src: local('{font}');{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{font}'; font-weight: bold; src: local('{font} Bold');{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{font}'; font-style: italic; src: local('{font} Italic');{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{font}'; font-weight: bold; font-style: italic; src: local('{font} Bold Italic');{overrides} }}");
        }
        return sb.ToString();
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
                if (hp.Header == null) continue;
                if (!HeaderFooterHasContent(hp.Header)) continue;
                sb.AppendLine($"<div class=\"{cssClass}\">");
                RenderHeaderFooterBody(sb, hp.Header);
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
                if (fp.Footer == null) continue;
                if (!HeaderFooterHasContent(fp.Footer)) continue;
                sb.AppendLine($"<div class=\"{cssClass}\">");
                RenderHeaderFooterBody(sb, fp.Footer);
                sb.AppendLine("</div>");
                break;
            }
        }
    }

    /// <summary>Returns true if the header/footer has any visible content:
    /// text, table, image/drawing, or field.</summary>
    private static bool HeaderFooterHasContent(OpenXmlElement hf)
    {
        foreach (var child in hf.ChildElements)
        {
            if (child is Table) return true;
            if (child is Paragraph p)
            {
                if (!string.IsNullOrWhiteSpace(p.InnerText)) return true;
                if (p.Descendants<Drawing>().Any()) return true;
                if (p.Descendants<FieldChar>().Any() || p.Descendants<SimpleField>().Any()) return true;
            }
        }
        return false;
    }

    /// <summary>Iterate header/footer children in order, rendering paragraphs
    /// and tables. Previously only paragraphs were emitted, dropping layout
    /// tables and image-only paragraphs.</summary>
    private void RenderHeaderFooterBody(StringBuilder sb, OpenXmlElement hf)
    {
        foreach (var child in hf.ChildElements)
        {
            if (child is Paragraph para)
                RenderParagraphHtml(sb, para);
            else if (child is Table tbl)
                RenderTableHtml(sb, tbl);
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
        var multiLevelCounters = new Dictionary<int, int>(); // ilvl → counter for multi-level numbering
        var headingCounters = new Dictionary<int, int>(); // ilvl → counter for heading auto-numbering from style numPr
        bool pendingLiClose = false; // defer </li> to allow nested lists inside
        bool inMultiColumn = false; // track whether we're inside a multi-column div

        // Pre-scan: build a map of section column counts from inline sectPr breaks
        // The last section's cols come from the body sectPr
        var bodySectPr = body.GetFirstChild<SectionProperties>();
        var bodyColCount = GetSectionColumnCount(bodySectPr);

        int wParaCount = 0, wTableCount = 0;
        int wBlockCount = 0;
        bool inList = false;
        int pendingBlockClose = 0; // block number that needs <!--wE:N--> before next block starts
        for (int ei = 0; ei < elements.Count; ei++)
        {
            var element = elements[ei];

            // Emit invisible anchors for watch scroll targeting
            if (element is Paragraph) { wParaCount++; sb.Append($"<a id=\"w-p-{wParaCount}\"></a>"); }
            else if (element is Table) { wTableCount++; sb.Append($"<a id=\"w-table-{wTableCount}\"></a>"); }

            // Block markers for server-side diff: each top-level block gets <!--wB:N--> / <!--wE:N-->
            // A "block" is: one paragraph, one table, one equation, OR an entire list (ul/ol group)
            // SectionProperties are skipped (not visual content, no block)
            if (element is SectionProperties) continue;
            var isListItem = element is Paragraph p2 && GetParagraphListStyle(p2) != null;
            if (!isListItem && inList)
            {
                // Leaving a list — close the list block
                sb.Append($"<span class=\"we\" data-block=\"{wBlockCount}\" style=\"display:none\"></span>");
                inList = false;
                pendingBlockClose = 0;
            }
            // Close previous non-list block if pending
            if (pendingBlockClose > 0)
            {
                sb.Append($"<span class=\"we\" data-block=\"{pendingBlockClose}\" style=\"display:none\"></span>");
                pendingBlockClose = 0;
            }
            if (isListItem && !inList)
            {
                // Entering a list — open a new block
                wBlockCount++;
                sb.Append($"<span class=\"wb\" data-block=\"{wBlockCount}\" style=\"display:none\"></span>");
                inList = true;
            }
            else if (!isListItem)
            {
                // Non-list element — each is its own block, close deferred to handle continue
                wBlockCount++;
                sb.Append($"<span class=\"wb\" data-block=\"{wBlockCount}\" style=\"display:none\"></span>");
                pendingBlockClose = wBlockCount;
            }

            // Check for inline section break (sectPr inside paragraph pPr) — handle page breaks and column changes
            if (element is Paragraph sectPara && sectPara.ParagraphProperties?.GetFirstChild<SectionProperties>() is SectionProperties inlineSectPr)
            {
                var sectType = inlineSectPr.GetFirstChild<SectionType>();
                if (sectType?.Val?.Value == SectionMarkValues.NextPage
                    || sectType?.Val?.Value == SectionMarkValues.EvenPage
                    || sectType?.Val?.Value == SectionMarkValues.OddPage)
                {
                    sb.Append("<!--PAGE_BREAK-->");
                }

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
                    sb.AppendLine($"<div class=\"equation\"><span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(latex)}\" data-display=\"true\"></span></div>");
                    continue;
                }

                // Check if this is a list item
                var listStyle = GetParagraphListStyle(para);
                if (listStyle != null)
                {
                    var ilvl = para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
                    var numId = para.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value ?? 0;
                    var numFmt = GetNumberingFormat(numId, ilvl);
                    var lvlText = GetLevelText(numId, ilvl);
                    var isMultiLevel = lvlText != null && System.Text.RegularExpressions.Regex.Matches(lvlText, @"%\d").Count > 1;
                    var picBulletUri = listStyle == "bullet" ? GetPicBulletDataUri(numId, ilvl) : null;
                    var tag = listStyle == "bullet" ? "ul" : "ol";

                    // When numId changes, decide: nesting or new list
                    if (currentNumId != null && numId != currentNumId)
                    {
                        if (listStack.Count > 0 && !numIdLevelOffset.ContainsKey(numId))
                        {
                            var curIndent = GetListLevelIndent(currentNumId.Value, currentListLevel);
                            var newIndent = GetListLevelIndent(numId, ilvl);
                            if (newIndent > curIndent)
                            {
                                numIdLevelOffset[numId] = currentListLevel + 1 - ilvl;
                            }
                            else
                            {
                                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                                olCountPerLevel.Clear();
                                multiLevelCounters.Clear();
                            }
                        }
                        else if (listStack.Count == 0)
                        {
                            // Previous list was closed by non-list content — reset counters for new list
                            olCountPerLevel.Clear();
                            multiLevelCounters.Clear();
                            numIdLevelOffset.Clear();
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

                    // Adjust nesting (close deeper levels)
                    while (listStack.Count > ilvl + 1)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        sb.AppendLine("</li>");
                    }
                    if (pendingLiClose)
                    {
                        pendingLiClose = false;
                    }

                    // Get indentation from numbering level definition
                    var (lvlLeft, lvlHanging) = GetListLevelIndentFull(numId, ilvl);
                    var parentLeft = ilvl > 0 ? GetListLevelIndent(numId, ilvl - 1) : 0;
                    double indentPt;
                    if (isMultiLevel)
                    {
                        // Multi-level: padding = number start position (left - hanging - parent)
                        indentPt = (lvlLeft - lvlHanging - parentLeft) / 20.0;
                    }
                    else
                    {
                        // Normal list: padding = relative indent from parent
                        indentPt = (lvlLeft - parentLeft) / 20.0;
                    }
                    if (indentPt < 18) indentPt = 18; // minimum indent
                    var hangingPt = lvlHanging / 20.0;
                    var listStyleParts = $"padding-left:{indentPt:0.#}pt;margin:0";
                    // CONSISTENCY(list-marker): every ordered list is rendered with
                    // list-style-type:none and a computed marker <span>. This lets
                    // WordNumFmtRenderer handle numFmt variants (chineseCounting,
                    // decimalZero, …) plus lvlText/suff/lvlJc that CSS `<ol type>`
                    // cannot express. See KNOWN_ISSUES.md #4.
                    if (tag == "ol") listStyleParts += ";list-style-type:none";
                    if (picBulletUri != null)
                        listStyleParts += $";list-style-image:url('{picBulletUri}')";
                    else if (tag == "ul")
                    {
                        listStyleParts += ";list-style-image:none"; // reset inherited picture bullet
                        // Map Word bullet character to CSS list-style-type
                        var bulletType = lvlText switch
                        {
                            "o" => "circle",
                            "\uf0a7" or "▪" or "\u25AA" => "square",
                            _ => "disc"
                        };
                        listStyleParts += $";list-style-type:{bulletType}";
                    }
                    var indentStyle = $" style=\"{listStyleParts}\"";

                    // Seed per-level counter from startOverride / level start
                    // when we're opening this level for the first time in the
                    // current list. Cross-list (different numId) continuation is
                    // preserved via olCountPerLevel survival.
                    int SeedStart(int forIlvl)
                    {
                        if (olCountPerLevel.TryGetValue(forIlvl, out var prev) && prev > 0)
                            return prev; // continuation
                        return (GetStartValue(numId, forIlvl) ?? 1) - 1;
                    }

                    while (listStack.Count < ilvl + 1)
                    {
                        sb.AppendLine($"<{tag}{indentStyle}>");
                        listStack.Push(tag);
                    }
                    // If same level but different list type, swap
                    if (listStack.Count > 0 && listStack.Peek() != tag)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        sb.AppendLine($"<{tag}{indentStyle}>");
                        listStack.Push(tag);
                    }

                    // Track counters
                    if (tag == "ol")
                    {
                        var seed = SeedStart(ilvl);
                        olCountPerLevel[ilvl] = olCountPerLevel.GetValueOrDefault(ilvl, seed) + 1;
                        multiLevelCounters[ilvl] = olCountPerLevel[ilvl];
                        // Reset deeper level counters
                        for (int lk = ilvl + 1; lk <= 8; lk++)
                        {
                            if (olCountPerLevel.ContainsKey(lk)) olCountPerLevel[lk] = 0;
                            if (multiLevelCounters.ContainsKey(lk)) multiLevelCounters[lk] = 0;
                        }
                    }

                    currentListType = listStyle;
                    currentListLevel = ilvl;
                    currentNumId = numId;
                    sb.Append("<li");
                    sb.Append($" data-path=\"/body/p[{wParaCount}]\"");
                    var paraStyle = GetParagraphInlineCss(para, isListItem: true);
                    if (!string.IsNullOrEmpty(paraStyle))
                        sb.Append($" style=\"{paraStyle}\"");
                    sb.Append(">");
                    // Computed marker for every ordered-list item (single or multi-level).
                    if (tag == "ol")
                    {
                        var template = string.IsNullOrEmpty(lvlText) ? $"%{ilvl + 1}" : lvlText!;
                        var marker = System.Text.RegularExpressions.Regex.Replace(template, @"%(\d)", m =>
                        {
                            var k = int.Parse(m.Groups[1].Value) - 1;
                            var lvlFmt = GetNumberingFormat(numId, k);
                            var counter = multiLevelCounters.GetValueOrDefault(k, 0);
                            return OfficeCli.Core.WordNumFmtRenderer.Render(counter, lvlFmt);
                        });
                        var suff = GetLevelSuffix(numId, ilvl);
                        var jc = GetLevelJustification(numId, ilvl);
                        var markerWidth = hangingPt > 0 ? $"{hangingPt:0.#}pt" : "3em";
                        var markerPadding = suff switch
                        {
                            "nothing" => "0",
                            "space" => "0.25em",
                            _ => "0.5em" // tab
                        };
                        var align = jc switch { "right" => "right", "center" => "center", _ => "left" };
                        sb.Append($"<span style=\"display:inline-block;min-width:{markerWidth};padding-right:{markerPadding};text-align:{align}\">{HtmlEncode(marker)}</span>");
                    }
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
                    var hasReflect = HasW14Reflection(para);
                    sb.Append($"<h{headingLevel}");
                    sb.Append($" data-path=\"/body/p[{wParaCount}]\"");
                    var hStyle = GetParagraphInlineCss(para);
                    // Remove bottom spacing when reflection follows immediately
                    if (hasReflect)
                        hStyle = string.IsNullOrEmpty(hStyle) ? "margin-bottom:0" : $"{hStyle};margin-bottom:0";
                    if (!string.IsNullOrEmpty(hStyle))
                        sb.Append($" style=\"{hStyle}\"");
                    sb.Append(">");

                    // Heading auto-numbering: if the heading's style chain
                    // carries a numPr, expand the level's lvlText ("%1.%2")
                    // against the running heading counters and prepend the
                    // result as a <span class="heading-num">.
                    //
                    // An explicit `<w:numPr><w:numId w:val="0"/></w:numPr>` on
                    // the paragraph suppresses this heading's number without
                    // disturbing the sibling counter (Word: …2→3→unnumbered→4).
                    var hNumPr = IsNumberingSuppressed(para) ? null : ResolveNumPrFromStyle(para);
                    if (hNumPr is { } hn)
                    {
                        headingCounters[hn.Ilvl] = headingCounters.GetValueOrDefault(hn.Ilvl, 0) + 1;
                        // Reset deeper level counters whenever a shallower heading ticks.
                        for (int lk = hn.Ilvl + 1; lk <= 8; lk++)
                            if (headingCounters.ContainsKey(lk)) headingCounters[lk] = 0;

                        var lvlText = GetLevelText(hn.NumId, hn.Ilvl);
                        if (!string.IsNullOrEmpty(lvlText))
                        {
                            var numStr = System.Text.RegularExpressions.Regex.Replace(lvlText, @"%(\d)", m =>
                            {
                                var lk = int.Parse(m.Groups[1].Value) - 1;
                                var lvlFmt = GetNumberingFormat(hn.NumId, lk);
                                var counter = headingCounters.GetValueOrDefault(lk, 0);
                                return OfficeCli.Core.WordNumFmtRenderer.Render(counter, lvlFmt);
                            });
                            // Skip the auto-num span when the paragraph text
                            // already begins with the computed number, so a
                            // user-typed "1. Overview" does not render as
                            // "1. 1. Overview".
                            var paraText = GetParagraphText(para).TrimStart();
                            if (!paraText.StartsWith(numStr, StringComparison.Ordinal))
                                sb.Append($"<span class=\"heading-num\" style=\"margin-right:0.5em\">{HtmlEncode(numStr)}</span>");
                        }
                    }

                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine($"</h{headingLevel}>");
                    if (hasReflect)
                        AppendW14ReflectionBlock(sb, para, $"h{headingLevel}", GetParagraphInlineCss(para));
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

                    // VML horizontal rule (w:pict > v:rect[o:hr="t"])
                    if (IsVmlHorizontalRule(para))
                    {
                        RenderVmlHorizontalRule(sb, para);
                        continue;
                    }

                    // Empty paragraph = spacing break
                    if (runs.Count == 0 && mathElements.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        sb.AppendLine($"<p class=\"empty\" data-path=\"/body/p[{wParaCount}]\">&nbsp;</p>");
                        continue;
                    }

                    // Inline equation only
                    if (mathElements.Count > 0 && runs.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                        sb.AppendLine($"<div class=\"equation\"><span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(latex)}\" data-display=\"true\"></span></div>");
                        continue;
                    }

                    sb.Append("<p");
                    sb.Append($" data-path=\"/body/p[{wParaCount}]\"");
                    // Add CSS class for TOC paragraphs (suppress hyperlink styling, enable dot leaders)
                    var paraStyleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    if (paraStyleId != null && paraStyleId.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                        sb.Append(" class=\"toc\"");
                    var pStyle = GetParagraphInlineCss(para);
                    if (!string.IsNullOrEmpty(pStyle))
                        sb.Append($" style=\"{pStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine("</p>");
                    AppendW14ReflectionBlock(sb, para, "p", pStyle);
                }
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                var latex = FormulaParser.ToLatex(element);
                sb.AppendLine($"<div class=\"equation\"><span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(latex)}\" data-display=\"true\"></span></div>");
            }
            else if (element is Table table)
            {
                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                RenderTableHtml(sb, table, dataPath: $"/body/table[{wTableCount}]");
            }
        }

        // Close any pending block (last element was non-list with continue, or last list block)
        if (pendingBlockClose > 0) sb.Append($"<span class=\"we\" data-block=\"{pendingBlockClose}\" style=\"display:none\"></span>");
        if (inList) sb.Append($"<span class=\"we\" data-block=\"{wBlockCount}\" style=\"display:none\"></span>");
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

    /// <summary>Get the left indent and hanging indent (in twips) for a numbering level definition.</summary>
    private (int left, int hanging) GetListLevelIndentFull(int numId, int ilvl)
    {
        var lvl = GetLevel(numId, ilvl);
        var indent = lvl?.PreviousParagraphProperties?.Indentation;
        int left = 0, hanging = 0;
        if (indent?.Left?.Value is string ls && int.TryParse(ls, out var lt))
            left = lt;
        if (indent?.Hanging?.Value is string hs && int.TryParse(hs, out var ht))
            hanging = ht;
        return (left, hanging);
    }

    private int GetListLevelIndent(int numId, int ilvl) => GetListLevelIndentFull(numId, ilvl).left;
}
