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

        // #8a: section-relative footnote numbering. When a section's
        // FootnoteProperties.NumberingRestart = eachSect, the fn counter
        // resets at that section boundary. FnLabels persists the displayed
        // label per fnId so the bottom-of-page <div class="footnotes">
        // list can emit the same number as the superscript ref.
        public int CurrentSectionIdx { get; set; }
        public int FnCountInSection { get; set; }
        public bool FnRestartEachSection { get; set; }
        public Dictionary<int, string> FnLabels { get; } = new();

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
        try
        {
            return ViewAsHtmlCore(pageFilter);
        }
        catch (System.Xml.XmlException)
        {
            // Any lazily-parsed subpart (styles/theme/numbering/footnotes/
            // header/footer/settings) can throw XmlException deep inside a
            // Render* callee if the backing XML is malformed. Treat the whole
            // preview as best-effort and degrade gracefully rather than
            // crashing the view command.
            return "<html><body><p>(document xml malformed)</p></body></html>";
        }
    }

    private string ViewAsHtmlCore(string? pageFilter)
    {
        _ctx = new HtmlRenderContext();
        ResolveThemeCjkFont();
        // Malformed docx (e.g. <!DOCTYPE> prolog, bogus encoding= attribute
        // on the XML declaration) makes accessing the lazily-parsed Document
        // throw XmlException. Tolerate it as an empty-body preview rather
        // than crashing the command.
        Body? body;
        try { body = _doc.MainDocumentPart?.Document?.Body; }
        catch (System.Xml.XmlException)
        {
            return "<html><body><p>(document xml malformed)</p></body></html>";
        }
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

        // Per-(numId, ilvl) marker CSS — picks up abstractNum level rPr
        // (color/font/size/bold/italic) and the actual lvlText glyph for
        // bullets. Without this every list marker rendered in the preview is
        // black, normal, and uses CSS's default disc/decimal — diverging from
        // what real Word renders.
        var markerCss = BuildListMarkerCss(body);
        if (!string.IsNullOrEmpty(markerCss))
        {
            sb.AppendLine("<style>");
            sb.AppendLine(markerCss);
            sb.AppendLine("</style>");
        }
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
                var families = string.Join("&", googleFonts
                    .Select(SanitizeFontName)
                    .Where(f => !string.IsNullOrEmpty(f))
                    .Select(f => $"family={f.Replace(' ', '+')}:ital,wght@0,400;0,700;1,400;1,700"));
                // media=print + onload swap → load asynchronously without blocking first paint
                // (Google Fonts is unreachable in many networks and would otherwise stall render until TCP timeout).
                sb.AppendLine($"<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?{families}&display=swap\" media=\"print\" onload=\"this.media='all'\" onerror=\"this.remove()\">");
            }
        }
        // KaTeX for math rendering — only include when the document actually has formulas.
        // Same non-blocking load trick so KaTeX CSS can never stall first paint.
        bool hasMathFormulas = body.Descendants<M.OfficeMath>().Any();
        if (hasMathFormulas)
        {
            sb.AppendLine("<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css\" media=\"print\" onload=\"this.media='all'\" onerror=\"this.remove()\">");
            sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js\" onerror=\"document.querySelectorAll('.katex-formula').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666'})\"></script>");
        }
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        // Render body into temporary buffer, then split on page breaks
        var maxW = $"width:{pgLayout.WidthPt:0.#}pt";
        var bodySb = new StringBuilder();
        _ctx.RenderingBody = true;
        RenderBodyHtml(bodySb, body);
        _ctx.RenderingBody = false;

        // #3: per-section header/footer bundles keyed by type. Resolved
        // at this stage so the page-emit loop can pick the right variant
        // per page (titlePg → first-page header; evenAndOddHeaders →
        // parity-based; default otherwise).
        var allSectionsForHf = CollectSections(body);
        var sectionHeaders = BuildSectionHfBundles(allSectionsForHf, isHeader: true);
        var sectionFooters = BuildSectionHfBundles(allSectionsForHf, isHeader: false);
        var evenAndOddGlobal = _doc.MainDocumentPart?.DocumentSettingsPart?
            .Settings?.GetFirstChild<EvenAndOddHeaders>() != null;
        // Legacy fallback for docs that didn't come through CollectSections'
        // per-section resolution path (e.g. no headers at body level).
        var fallbackHeaderSb = new StringBuilder();
        RenderHeaderFooterHtml(fallbackHeaderSb, isHeader: true);
        var fallbackHeaderHtml = fallbackHeaderSb.ToString();
        var fallbackFooterSb = new StringBuilder();
        RenderHeaderFooterHtml(fallbackFooterSb, isHeader: false);
        var footerHtml = fallbackFooterSb.ToString();

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
        // Match a single-digit-only run rendered as either <span> or <p>.
        // The footer's PAGE field is typically a single run; the tag name
        // depends on whether the run carries rPr styling.
        // Wrap the matched digit run in a sentinel span so the per-page
        // paginate JS can locate PAGE/NUMPAGES fields without clobbering
        // unrelated digit-only content (e.g. "2026", "5 USD", chapter ids).
        var pageNumPattern = new Regex(@"(<(?:span|p)[^>]*>)\s*\d+\s*(</(?:span|p)>)");
        var footerTemplate = pageNumPattern.Replace(footerHtml,
            "$1<span class=\"page-num-field\"><!--PAGE_NUM--></span>$2", 1);
        var footerTemplateWithTotal = pageNumPattern.Replace(footerTemplate,
            "$1<span class=\"num-pages-field\"><!--NUM_PAGES--></span>$2", 1);
        footerTemplate = footerTemplateWithTotal;

        // Section-level multi-column layout: w:cols num=N sep=true
        var sectCols = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>()?.GetFirstChild<Columns>();
        var colCount = sectCols?.ColumnCount?.Value ?? 1;
        var colSep = sectCols?.Separator?.Value == true;
        var colSpacing = sectCols?.Space?.Value;
        // CSS columns need a bounded height to balance — min-height alone
        // leaves the body unbounded so all content stacks in column 1 and
        // overflows the page. Use the doc-level pgLayout body height.
        var colBodyHeightPt = pgLayout.HeightPt - pgLayout.MarginTopPt - pgLayout.MarginBottomPt;
        var colBodyStyle = colCount > 1
            ? $" style=\"column-count:{colCount}"
                + $";height:{colBodyHeightPt.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture)}pt"
                + (colSep ? ";column-rule:1px solid #000" : "")
                + (int.TryParse(colSpacing, out var csp) && csp > 0 ? $";column-gap:{csp / 20.0:0.##}pt" : "")
                + "\""
            : "";

        // Per-section page layout (#7a00): each page carries one or more
        // <!--SECT:N--> markers inserted by RenderBodyHtml. The last marker
        // seen (inclusive of this page) decides the page's size/margins;
        // pages with no marker inherit from the previous page.
        var sections = CollectSections(body);
        var sectRegex = new Regex(@"<!--SECT:(\d+)-->");
        var activeLayout = pgLayout;
        // #10: per-section pgNumType — w:start resets the displayed page
        // counter at the section boundary; w:fmt swaps the number format
        // (decimalZero, upperRoman, …) applied to PAGE/NUMPAGES substitutions.
        int displayedPageNum = 0;
        string displayedFmt = "decimal";
        int activeSectionIdx = 0;
        int prevActiveSectionIdx = -1;
        for (int i = 0; i < pageList.Count; i++)
        {
            var pgContent = pageList[i];
            var sectMatches = sectRegex.Matches(pgContent);
            if (sectMatches.Count > 0)
            {
                var lastIdx = int.Parse(sectMatches[^1].Groups[1].Value);
                if (lastIdx >= 0 && lastIdx < sections.Count)
                {
                    activeLayout = GetPageLayoutFor(sections[lastIdx]);
                    activeSectionIdx = lastIdx;
                    var pgNumType = sections[lastIdx].GetFirstChild<PageNumberType>();
                    if (pgNumType?.Start?.Value is int startVal)
                        displayedPageNum = startVal - 1; // will ++ below
                    // Open XML SDK v3+: Enum.ToString() returns a
                    // debug string like "NumberFormatValues { }"; use
                    // InnerText to get the XML-level token ("decimalZero").
                    if (pgNumType?.Format?.InnerText is { Length: > 0 } fmtStr)
                        displayedFmt = fmtStr;
                }
                pgContent = sectRegex.Replace(pgContent, "");
                pageList[i] = pgContent;
            }
            displayedPageNum++;
            var isFirstPageOfSection = activeSectionIdx != prevActiveSectionIdx;
            prevActiveSectionIdx = activeSectionIdx;
            // Per-page inline style carries full geometry (width / min-height
            // / padding) so sections with different page sizes or margins
            // override the base .page CSS rules.
            var ci = System.Globalization.CultureInfo.InvariantCulture;
            var pageStyle =
                $"width:{activeLayout.WidthPt.ToString("0.#", ci)}pt;" +
                $"min-height:{activeLayout.HeightPt.ToString("0.#", ci)}pt;" +
                $"padding:{activeLayout.MarginTopPt.ToString("0.#", ci)}pt " +
                $"{activeLayout.MarginRightPt.ToString("0.#", ci)}pt " +
                $"{activeLayout.MarginBottomPt.ToString("0.#", ci)}pt " +
                $"{activeLayout.MarginLeftPt.ToString("0.#", ci)}pt";
            // #1: lnNumType — read per-section line-number settings and
            // expose them as data-* attributes so the JS paginator can
            // inject line numbers after layout settles. Only applies when
            // countBy > 0; absent element means "no line numbers".
            string lineNumAttrs = "";
            if (activeSectionIdx >= 0 && activeSectionIdx < sections.Count)
            {
                var ln = sections[activeSectionIdx].GetFirstChild<LineNumberType>();
                // LineNumberType fields are Int16Value — malformed raw docs
                // (huge/negative start, non-numeric countBy) throw on .Value
                // access. Parse the raw InnerText ourselves and swallow.
                short by = 0;
                if (ln?.CountBy != null)
                    short.TryParse(ln.CountBy.InnerText, out by);
                if (ln != null && by > 0)
                {
                    short startN = 1;
                    if (ln.Start != null) short.TryParse(ln.Start.InnerText, out startN);
                    int distTwips = 0;
                    if (ln.Distance != null) int.TryParse(ln.Distance.InnerText, out distTwips);
                    var distPt = distTwips / 20.0;
                    var restart = ln.Restart?.InnerText ?? "newPage";
                    lineNumAttrs =
                        $" data-line-num-by=\"{by}\"" +
                        $" data-line-num-start=\"{startN}\"" +
                        $" data-line-num-dist=\"{distPt.ToString("0.#", ci)}\"" +
                        $" data-line-num-restart=\"{restart}\"";
                }
            }
            sb.AppendLine($"<div class=\"page-wrapper\" data-section=\"{i + 1}\" data-section-idx=\"{activeSectionIdx}\"{lineNumAttrs}>");
            sb.AppendLine($"<div class=\"page\" data-page=\"{i + 1}\" style=\"{pageStyle}\">");
            // #3: per-page header/footer selection. titlePg → first-page
            // variant; evenAndOddHeaders + even-numbered page → even
            // variant; otherwise default. The per-page header lands on
            // every page (previously only page 0 got it).
            var pageIsEven = (i + 1) % 2 == 0;
            var hdrPageNumStr = OfficeCli.Core.WordNumFmtRenderer.Render(displayedPageNum, displayedFmt);
            var perPageHeader = PickHeaderFooter(
                sectionHeaders, sections, activeSectionIdx,
                isFirstPageOfSection, pageIsEven, evenAndOddGlobal, fallbackHeaderHtml);
            // Same PAGE/NUMPAGES substitution as the footer path so headers
            // with field=page / field=numpages update per page instead of
            // rendering the author-time cached literal "1".
            var phdr = new Regex(@"(<(?:span|p)[^>]*>)\s*\d+\s*(</(?:span|p)>)");
            var perPageHeaderTemplate = phdr.Replace(perPageHeader,
                "$1<span class=\"page-num-field\"><!--PAGE_NUM--></span>$2", 1);
            perPageHeaderTemplate = phdr.Replace(perPageHeaderTemplate,
                "$1<span class=\"num-pages-field\"><!--NUM_PAGES--></span>$2", 1);
            sb.Append(perPageHeaderTemplate
                .Replace("<!--PAGE_NUM-->", hdrPageNumStr)
                .Replace("<!--NUM_PAGES-->", pageList.Count.ToString()));
            sb.Append($"<div class=\"page-body\"{colBodyStyle}>");
            sb.Append(pageList[i]);
            // Place footnotes on the page that contains the footnote reference
            if (!string.IsNullOrEmpty(footnotesHtml) && pageList[i].Contains("fn-ref"))
                sb.Append(footnotesHtml);
            // Place endnotes on the last page
            if (i == pageList.Count - 1 && !string.IsNullOrEmpty(endnotesHtml))
                sb.Append(endnotesHtml);
            sb.Append("</div>");
            var pageNumStr = OfficeCli.Core.WordNumFmtRenderer.Render(displayedPageNum, displayedFmt);
            // #3: same picker as header — first/even/default footer variant.
            var perPageFooter = PickHeaderFooter(
                sectionFooters, sections, activeSectionIdx,
                isFirstPageOfSection, pageIsEven, evenAndOddGlobal, footerHtml);
            // Rebuild the PAGE field placeholder on the picked footer.
            var pf = new Regex(@"(<(?:span|p)[^>]*>)\s*\d+\s*(</(?:span|p)>)");
            var perPageFooterTemplate = pf.Replace(perPageFooter,
                "$1<span class=\"page-num-field\"><!--PAGE_NUM--></span>$2", 1);
            perPageFooterTemplate = pf.Replace(perPageFooterTemplate,
                "$1<span class=\"num-pages-field\"><!--NUM_PAGES--></span>$2", 1);
            sb.Append(perPageFooterTemplate
                .Replace("<!--PAGE_NUM-->", pageNumStr)
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
        // Header template cloned per paginated page. Capture the fallback
        // header's PAGE/NUMPAGES placeholders so field updates work on
        // every continuation page, not just page 1.
        var headerTemplate = pageNumPattern.Replace(fallbackHeaderHtml, "$1<!--PAGE_NUM-->$2", 1);
        headerTemplate = pageNumPattern.Replace(headerTemplate, "$1<!--NUM_PAGES-->$2", 1);
        sb.AppendLine("  var htpl=" + JsStringLiteral(headerTemplate) + ";");
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
      // #7b00: when the overflowing child is a <table>, split it at the
      // row boundary and clone any rows carrying data-tbl-header=""1""
      // onto the continuation so long tables have repeating headers
      // across pages the way Word renders them.
      var firstOverflow=children[splitIdx];
      if(firstOverflow&&firstOverflow.tagName==='TABLE'){
        var table=firstOverflow;
        var tableTop=table.offsetTop-body.offsetTop;
        // Only top-level rows — querySelectorAll('tr') would also pick up
        // nested subtable rows and mangle nested structures on page splits.
        var trs=Array.from(table.querySelectorAll('tr')).filter(function(tr){
          return tr.closest('table')===table;
        });
        var hdrRows=trs.filter(function(tr){return tr.getAttribute('data-tbl-header')==='1';});
        // Find first row whose bottom exceeds availH (relative to body).
        var rowSplit=-1;
        for(var ri=0;ri<trs.length;ri++){
          if(trs[ri].getAttribute('data-tbl-header')==='1')continue;
          var rowBot=trs[ri].offsetTop+trs[ri].offsetHeight-body.offsetTop;
          if(rowBot>availH){rowSplit=ri;break;}
        }
        if(rowSplit>0){
          // Build continuation table; clone attributes + header rows.
          var cont=table.cloneNode(false);
          var tbodies=table.querySelectorAll('tbody');
          var contBody=tbodies.length?document.createElement('tbody'):cont;
          if(tbodies.length)cont.appendChild(contBody);
          hdrRows.forEach(function(h){contBody.appendChild(h.cloneNode(true));});
          for(var rj=rowSplit;rj<trs.length;rj++){
            if(trs[rj].getAttribute('data-tbl-header')==='1')continue;
            contBody.appendChild(trs[rj]);
          }
          // Insert continuation as new sibling after the source table so
          // the split-point logic below moves it to a new page.
          table.parentNode.insertBefore(cont,table.nextSibling);
          children=Array.from(body.children);
          splitIdx=children.indexOf(cont);
        }
      }
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
      // Clone header into new page (prepended before page-body) so each
      // continuation page shows the same header tree as the source page.
      if(htpl){
        var nh=document.createElement('div');
        nh.innerHTML=htpl.replace('<!--PAGE_NUM-->',(pi+2).toString());
        if(nh.firstChild)np.appendChild(nh.firstChild);
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
      // Only touch explicit PAGE/NUMPAGES sentinel spans — scanning every
      // digit-only leaf silently rewrote years, prices, chapter ids etc.
      p.querySelectorAll('.page-num-field').forEach(function(s){s.textContent=(i+1);});
      p.querySelectorAll('.num-pages-field').forEach(function(s){s.textContent=allPages.length;});
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
    else{setTimeout(positionFootnotes,0);setTimeout(wrapFloats,0);setTimeout(applyLineNumbers,0);setTimeout(applyPageFilter,0);setTimeout(function(){scalePages(false);},0);}
  }
  // #2 / #7b light approximation: a floating table whose CSS has float:*
  // sits directly under .page-body (flex column) and has its float ignored.
  // Wrap it + following prose siblings in a non-flex BFC div until either
  // a heading, another table, or the wrap is tall enough for prose to
  // have cleared the table. Re-run is idempotent.
  function wrapFloats(){
    // Collect direct page-body children whose outer CSS or whose first
    // child <img> has float:*. Both cases need a BFC wrapper so the float
    // can push following prose sideways.
    var candidates=[];
    document.querySelectorAll('.page-body > *').forEach(function(el){
      if(el.parentElement && el.parentElement.classList.contains('float-wrap'))return;
      var ownFloat=(el.style&&el.style.cssFloat)||'';
      if(!ownFloat && el.getAttribute){
        var st=el.getAttribute('style')||'';
        if(/float\s*:\s*(left|right)/.test(st))ownFloat='y';
      }
      var innerImg=el.querySelector&&el.querySelector('img[style*=""float:""]');
      if(ownFloat||innerImg)candidates.push({el:el,anchor:innerImg||el});
    });
    candidates.forEach(function(c){
      var wrap=document.createElement('div');
      wrap.className='float-wrap';
      wrap.style.cssText='display:block;overflow:auto';
      c.el.parentNode.insertBefore(wrap,c.el);
      wrap.appendChild(c.el);
      var anchorH=c.anchor.offsetHeight||c.el.offsetHeight;
      // Absorb following siblings until a hard break or clearance.
      for(var guard=0;guard<50;guard++){
        var nxt=wrap.nextSibling;
        if(!nxt)break;
        if(nxt.nodeType===1){
          var tag=nxt.tagName;
          if(tag==='TABLE'||(tag&&tag.length===2&&tag[0]==='H'))break;
          if(nxt.classList&&nxt.classList.contains('footnotes'))break;
        }
        wrap.appendChild(nxt);
        if(wrap.offsetHeight>anchorH+16)break;
      }
    });
  }
  // #1: walk each page's text nodes, use Range.getClientRects() to find
  // visual line rectangles, and inject absolute-positioned <span> markers
  // in the left margin. Honors countBy (show every Nth line), start
  // (initial number), distance (offset from text), and restart semantics
  // (newPage resets per-page; continuous keeps running).
  function applyLineNumbers(){
    var wrappers=document.querySelectorAll('.page-wrapper[data-line-num-by]');
    if(!wrappers.length)return;
    var runningNum=null;  // continuous/newSection running counter across pages
    var prevSection=null;
    wrappers.forEach(function(wrap){
      var body=wrap.querySelector('.page-body');
      if(!body)return;
      // Clear any previous markers before re-applying (keeps idempotent).
      body.querySelectorAll('.line-number').forEach(function(m){m.remove();});
      var by=parseInt(wrap.dataset.lineNumBy||'1')||1;
      var start=parseInt(wrap.dataset.lineNumStart||'1')||1;
      var dist=parseFloat(wrap.dataset.lineNumDist||'0')||0;
      var restart=wrap.dataset.lineNumRestart||'newPage';
      var sectionIdx=wrap.dataset.sectionIdx||'-1';
      var sectionChanged=prevSection!==null && prevSection!==sectionIdx;
      var current;
      if(restart==='newPage'||runningNum===null) current=start;
      else if(restart==='newSection') current=sectionChanged?start:runningNum;
      else current=runningNum;  // continuous
      prevSection=sectionIdx;
      body.style.position='relative';
      var bodyRect=body.getBoundingClientRect();
      var seenY=Object.create(null);
      var lineTops=[];
      var walker=document.createTreeWalker(body,NodeFilter.SHOW_TEXT,{
        acceptNode:function(n){
          if(!n.textContent.trim())return NodeFilter.FILTER_REJECT;
          // Skip line numbers we just injected (idempotence), footers, etc.
          var el=n.parentElement;
          while(el && el!==body){
            if(el.classList && (el.classList.contains('line-number')
              ||el.classList.contains('footnotes')))return NodeFilter.FILTER_REJECT;
            el=el.parentElement;
          }
          return NodeFilter.FILTER_ACCEPT;
        }
      });
      var node;
      while((node=walker.nextNode())){
        var range=document.createRange();
        range.selectNodeContents(node);
        var rects=range.getClientRects();
        for(var i=0;i<rects.length;i++){
          var r=rects[i];
          var y=Math.round(r.top-bodyRect.top);
          if(!(y in seenY)){seenY[y]=true;lineTops.push(y);}
        }
      }
      lineTops.sort(function(a,b){return a-b;});
      var leftPt=-(dist+20);
      for(var li=0;li<lineTops.length;li++){
        var n=current+li;
        if(by>1 && n%by!==0)continue;
        var marker=document.createElement('span');
        marker.className='line-number';
        marker.textContent=n;
        marker.style.cssText='position:absolute;left:'+leftPt+'pt;'
          +'font-size:8pt;color:#888;user-select:none;pointer-events:none;';
        marker.style.top=lineTops[li]+'px';
        body.appendChild(marker);
      }
      runningNum=current+lineTops.length;
    });
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
  function _loadKatexLazy(cb){
    // Watch mode: doc may start formula-free (KaTeX tags omitted), then
    // gain a formula via SSE patch. Inject CSS + JS on demand; on load,
    // re-invoke the caller so the new formula renders.
    if(window._katexLoading){window._katexCallbacks=window._katexCallbacks||[];window._katexCallbacks.push(cb);return;}
    window._katexLoading=true;window._katexCallbacks=[cb];
    var link=document.createElement('link');link.rel='stylesheet';link.href='https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css';link.onerror=function(){this.remove();};document.head.appendChild(link);
    var s=document.createElement('script');s.src='https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js';
    s.onload=function(){(window._katexCallbacks||[]).forEach(function(f){try{f();}catch(e){}});window._katexCallbacks=[];};
    s.onerror=function(){document.querySelectorAll('.katex-formula:not(.katex-rendered)').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666';el.classList.add('katex-rendered');});};
    document.head.appendChild(s);
  }
  function renderNewContent(){
    var pending=document.querySelectorAll('.katex-formula:not(.katex-rendered)');
    if(typeof katex!=='undefined'){
      pending.forEach(function(el){
        try{katex.render(el.dataset.formula,el,{throwOnError:false,displayMode:!!el.dataset.display});}catch(e){el.textContent=el.dataset.formula;}
        el.classList.add('katex-rendered');
      });
    }else if(pending.length>0){
      _loadKatexLazy(renderNewContent);
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
        var result = GetPageLayoutFor(sectPr);
        if (_ctx != null) _ctx.CachedPageLayout = result;
        return result;
    }

    // OpenXML typed-value accessors throw on malformed raw attrs
    // (e.g. negative on UInt32Value, overflow on Int16Value, non-numeric).
    // These wrappers turn any access/parse exception into the fallback.
    private static double SafeUIntTwips(Func<uint?> read, double fallback)
    {
        try { return (double)(read() ?? (uint)fallback); }
        catch { return fallback; }
    }

    private static double SafeIntTwips(Func<int?> read, double fallback)
    {
        try { return (double)(read() ?? (int)fallback); }
        catch { return fallback; }
    }

    private static PageLayout GetPageLayoutFor(SectionProperties? sectPr)
    {
        var pgSz = sectPr?.GetFirstChild<PageSize>();
        var pgMar = sectPr?.GetFirstChild<PageMargin>();
        const double c = 2.54 / 1440.0; // twips → cm
        const double p = 1.0 / 20.0;    // twips → pt (exact)
        // OOXML schema types (UInt32Value) throw on .Value access when the
        // raw attribute is malformed (negative, non-numeric). Tolerate it.
        double wTwips = SafeUIntTwips(() => pgSz?.Width?.Value, WordPageDefaults.A4WidthTwips);
        double hTwips = SafeUIntTwips(() => pgSz?.Height?.Value, WordPageDefaults.A4HeightTwips);
        // Landscape: OOXML orient=landscape flips the width/height semantics.
        // w:w/w:h already reflect the orientation in most real-world docs,
        // but guard against the rare case where w:w < w:h but orient=landscape.
        if (pgSz?.Orient?.Value == PageOrientationValues.Landscape && wTwips < hTwips)
            (wTwips, hTwips) = (hTwips, wTwips);
        // pgMar Top/Bottom are Int32Value, Left/Right/Header/Footer are
        // UInt32Value — all throw on .Value access for malformed raw attrs.
        // Wrap in the same swallow-to-fallback helper as pgSz.
        double tTwips = SafeIntTwips(() => pgMar?.Top?.Value, 1440);
        double bTwips = SafeIntTwips(() => pgMar?.Bottom?.Value, 1440);
        double lTwips = SafeUIntTwips(() => pgMar?.Left?.Value, 1440);
        double rTwips = SafeUIntTwips(() => pgMar?.Right?.Value, 1440);
        double hdTwips = SafeUIntTwips(() => pgMar?.Header?.Value, 851);
        double fdTwips = SafeUIntTwips(() => pgMar?.Footer?.Value, 992);
        return new PageLayout(
            wTwips * c, hTwips * c, tTwips * c, bTwips * c, lTwips * c, rTwips * c, hdTwips * c, fdTwips * c,
            wTwips * p, hTwips * p, tTwips * p, bTwips * p, lTwips * p, rTwips * p, hdTwips * p, fdTwips * p);
    }

    /// <summary>
    /// Collect sectPrs in document order. Each paragraph's inline sectPr
    /// (held in its pPr) terminates a section; the body's trailing sectPr
    /// owns everything after the last inline one.
    /// </summary>
    private List<SectionProperties> CollectSections(Body body)
    {
        var list = new List<SectionProperties>();
        foreach (var p in body.Elements<Paragraph>())
        {
            var inline = p.ParagraphProperties?.GetFirstChild<SectionProperties>();
            if (inline != null) list.Add(inline);
        }
        var trailing = body.GetFirstChild<SectionProperties>();
        if (trailing != null) list.Add(trailing);
        return list;
    }

    private record DocDef(string Font, double SizePt, double LineHeight, string Color, double GridLinePitchPt,
        double SpaceAfterPt = 0, string DefaultAlign = "left");

    private DocDef ReadDocDefaults()
    {
        // Malformed styles.xml — same fallback policy as theme1.xml: the
        // preview should still render body content using system defaults
        // rather than rejecting the entire doc.
        DocDefaults? defs = null;
        Style? defaultStyle = null;
        try
        {
            defs = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
            defaultStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.Default?.Value == true && s.Type?.Value == StyleValues.Paragraph);
        }
        catch (System.Xml.XmlException) { }
        var rPr = defs?.RunPropertiesDefault?.RunPropertiesBaseStyle;
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
            try
            {
                var minor = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont;
                font = NonEmpty(minor?.EastAsianFont?.Typeface) ?? NonEmpty(minor?.LatinFont?.Typeface);
            }
            catch (System.Xml.XmlException) { }
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
        if (cv != null && cv != "auto" && IsHexColor(cv)) color = $"#{cv}";
        else if (GetThemeColors().TryGetValue("dk1", out var dk1) && IsHexColor(dk1)) color = $"#{dk1}";

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

        return new DocDef(font ?? GetThemeMinorLatinFont() ?? OfficeDefaultFonts.MinorLatin, sizePt, lineH, color, gridLinePitchPt, spaceAfterPt, defaultAlign);
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
        // From theme (malformed theme1.xml shouldn't taint the font set).
        try
        {
            var theme = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
            var majFont = theme?.MajorFont?.LatinFont?.Typeface?.Value;
            if (!string.IsNullOrEmpty(majFont)) fonts.Add(majFont);
            var minFont = theme?.MinorFont?.LatinFont?.Typeface?.Value;
            if (!string.IsNullOrEmpty(minFont)) fonts.Add(minFont);
        }
        catch (System.Xml.XmlException) { }
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
        // Any of the subpart accesses below (settings.xml, styles.xml,
        // theme1.xml) can throw XmlException if the corresponding part is
        // malformed. Catch at subpart granularity so the ViewAsHtml outer
        // guard doesn't collapse the whole preview to a malformed stub.
        try
        {
            var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
            var themeFontLang = settings?.Descendants<DocumentFormat.OpenXml.Wordprocessing.ThemeFontLanguages>().FirstOrDefault();
            _eastAsiaLang = themeFontLang?.EastAsia?.Value;
        }
        catch (System.Xml.XmlException) { }

        if (_eastAsiaLang == null)
        {
            try
            {
                var docDefLang = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                    ?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle
                    ?.Languages;
                _eastAsiaLang = docDefLang?.EastAsia?.Value;
            }
            catch (System.Xml.XmlException) { }
        }

        DocumentFormat.OpenXml.Drawing.FontScheme? fontScheme = null;
        try { fontScheme = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme; }
        catch (System.Xml.XmlException) { }
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
            // Font names come straight from w:rFonts@ascii/hAnsi/eastAsia and
            // theme.xml — attacker-controlled strings. Without sanitization,
            // a name like `x'; } body { background: url(javascript:...) } /*`
            // would inject arbitrary CSS rules into the stylesheet. Drop
            // anything not in the safe set (letters/digits/spaces/.-_).
            var safeFont = SanitizeFontName(font);
            if (string.IsNullOrEmpty(safeFont)) continue;
            var (ascentPct, descentPct) = FontMetricsReader.GetAscentDescentOverride(safeFont);
            var overrides = ascentPct > 0
                ? $" ascent-override: {ascentPct:0.##}%; descent-override: {descentPct:0.##}%; line-gap-override: 0%;"
                : "";
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; src: local('{safeFont}');{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; font-weight: bold; src: local('{safeFont} Bold');{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; font-style: italic; src: local('{safeFont} Italic');{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; font-weight: bold; font-style: italic; src: local('{safeFont} Bold Italic');{overrides} }}");
        }
        return sb.ToString();
    }

    private static string? NonEmpty(string? s) => string.IsNullOrEmpty(s) ? null : s;

    /// <summary>Resolve shading fill color: direct hex or themeFill + themeFillTint/Shade.</summary>
    // Strictly-hex check for OOXML color attrs that flow into inline style.
    // Unvalidated interpolation into `background-color:#{fill}` lets a
    // malicious fill attribute escape the style context and inject HTML.
    // Allowlist of URL schemes that are safe to emit as clickable <a href=...>.
    // javascript:, vbscript:, and data: are all XSS vectors via OOXML
    // hyperlink relationships (attacker-controlled Target in .rels).
    // Keep only CSS-safe characters in a font-family name.
    private static string SanitizeFontName(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
        {
            if (char.IsLetterOrDigit(c) || c == ' ' || c == '-' || c == '_' || c == '.')
                sb.Append(c);
        }
        return sb.ToString().Trim();
    }

    private static bool IsSafeLinkUrl(string url)
    {
        if (string.IsNullOrEmpty(url)) return false;
        if (url.StartsWith("#")) return true;
        var decoded = System.Net.WebUtility.HtmlDecode(url).TrimStart();
        var colon = decoded.IndexOf(':');
        if (colon < 0) return true; // relative URL (path, query)
        var scheme = decoded.Substring(0, colon).ToLowerInvariant().Trim();
        return scheme is "http" or "https" or "mailto" or "tel" or "ftp" or "ftps";
    }

    private static bool IsHexColor(string s)
        => s.Length is 3 or 6 or 8
           && s.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'));

    private string? ResolveShadingFill(Shading? shading)
    {
        if (shading == null) return null;
        var fill = shading.Fill?.Value;
        if (fill != null && fill != "auto" && IsHexColor(fill)) return $"#{fill}";
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
                // VML watermark (<v:pict>) is visible content even though
                // it carries no plain text and no DrawingML Drawing element.
                if (p.Descendants<Picture>().Any()) return true;
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
            {
                // Legacy VML watermark: a <v:shape> in a <w:pict> with
                // a <v:textpath> child carrying the watermark string
                // (DRAFT / CONFIDENTIAL / …). DrawingML text boxes are
                // already handled by the shape renderer; VML is a
                // parallel deprecated format we must detect by name.
                var watermarkText = ExtractVmlWatermarkText(para);
                if (watermarkText != null)
                {
                    sb.Append($"<span class=\"vml-watermark\" style=\"position:absolute;" +
                              "top:50%;left:50%;transform:translate(-50%,-50%) rotate(-45deg);" +
                              "color:#d0d0d0;font-size:7em;font-weight:bold;" +
                              "z-index:0;pointer-events:none;white-space:nowrap;" +
                              "user-select:none\">");
                    sb.Append(HtmlEncode(watermarkText));
                    sb.Append("</span>");
                    continue;
                }
                RenderParagraphHtml(sb, para);
            }
            else if (child is Table tbl)
                RenderTableHtml(sb, tbl);
        }
    }

    /// <summary>
    /// Return the watermark text from a legacy VML <c>w:pict &gt; v:shape &gt;
    /// v:textpath</c> structure, or null if the paragraph does not carry one.
    /// </summary>
    private static string? ExtractVmlWatermarkText(Paragraph para)
    {
        foreach (var pict in para.Descendants<Picture>())
        {
            var shape = pict.Descendants().FirstOrDefault(e => e.LocalName == "shape"
                && e.NamespaceUri == "urn:schemas-microsoft-com:vml");
            if (shape == null) continue;
            var textPath = shape.Descendants().FirstOrDefault(e => e.LocalName == "textpath"
                && e.NamespaceUri == "urn:schemas-microsoft-com:vml");
            if (textPath == null) continue;
            var str = textPath.GetAttributes().FirstOrDefault(a => a.LocalName == "string").Value;
            if (!string.IsNullOrWhiteSpace(str)) return str;
        }
        return null;
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
        // Per-(abstractNumId, ilvl) running counter. Persists across numId
        // changes so that two num instances pointing at the same abstractNum
        // share a counter (Word's "continue" behavior) UNLESS the new num
        // carries an explicit <w:lvlOverride><w:startOverride/></w:lvlOverride>,
        // in which case we reset to the override value.
        var absNumLevelCounters = new Dictionary<int, Dictionary<int, int>>();
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

        // Section tracking for per-section page layout (#7a00). The first
        // section owns page 1; each inline sectPr ends its section and
        // bumps the index so the next page can adopt the next section's
        // width/height/margins.
        int currentSectionIdx = 0;
        sb.Append($"<!--SECT:{currentSectionIdx}-->");
        var allSections = CollectSections(body);
        ApplySectionFnSettings(allSections, currentSectionIdx);

        // Drop cap wrapping (#7c): a framePr dropCap paragraph and the
        // paragraph that follows must sit inside a non-flex container so
        // `float:left` on the drop cap actually wraps the follow-on text.
        // The parent page-body is a flex column which would otherwise
        // stack them vertically. Counts down from 2 → 0.
        int dropCapWrapRemaining = 0;

        for (int ei = 0; ei < elements.Count; ei++)
        {
            var element = elements[ei];

            // Emit body-level <w:bookmarkStart> as a navigable <a id="...">.
            // Word places bookmarkStart directly under <w:body> when the
            // bookmark spans multiple paragraphs; the paragraph-level
            // emitter in RenderParagraphContentHtml only catches bookmarks
            // authored inside a <w:p>. Without this, TOC hyperlinks and
            // in-document #anchor hrefs resolve to nothing.
            if (element is BookmarkStart bmStart)
            {
                var bmName = bmStart.Name?.Value;
                if (!string.IsNullOrEmpty(bmName) && !bmName.StartsWith("_GoBack"))
                    sb.Append($"<a id=\"{HtmlEncodeAttr(bmName)}\"></a>");
                continue;
            }

            // #7c: close drop cap wrap once the follow-on paragraph has
            // emitted. If we hit a non-paragraph (table, SectionProperties)
            // before the follow-on, also close to keep HTML well-formed.
            if (dropCapWrapRemaining > 0 && ei > 0)
            {
                var prev = elements[ei - 1];
                if (prev is Paragraph)
                {
                    dropCapWrapRemaining--;
                    if (dropCapWrapRemaining == 0) sb.Append("</div>");
                }
                else if (prev is Table)
                {
                    sb.Append("</div>");
                    dropCapWrapRemaining = 0;
                }
            }

            // #8a / #7a00: a paragraph whose pPr carries an inline sectPr
            // is the *last* paragraph of that section — it still belongs to
            // the current section's context. So advance the section index
            // AFTER that paragraph emitted, i.e. at the top of the NEXT
            // iteration.
            if (ei > 0 && elements[ei - 1] is Paragraph prevP
                && prevP.ParagraphProperties?.GetFirstChild<SectionProperties>() is SectionProperties prevInlineSectPr)
            {
                var sectType = prevInlineSectPr.GetFirstChild<SectionType>();
                if (sectType?.Val?.Value == SectionMarkValues.NextPage
                    || sectType?.Val?.Value == SectionMarkValues.EvenPage
                    || sectType?.Val?.Value == SectionMarkValues.OddPage)
                {
                    sb.Append("<!--PAGE_BREAK-->");
                }
                currentSectionIdx++;
                sb.Append($"<!--SECT:{currentSectionIdx}-->");
                ApplySectionFnSettings(allSections, currentSectionIdx);
            }

            // Emit invisible anchors for watch scroll targeting. #6: a
            // paragraph that exists purely as an m:oMathPara wrapper is
            // emitted as a <div class="equation">, not a <p>. Skip it from
            // the wParaCount sequence so /body/p[N] in data-path attrs
            // lines up with Navigation.cs's path resolution.
            if (element is Paragraph wpara && !IsOMathParaWrapperParagraph(wpara))
            { wParaCount++; sb.Append($"<a id=\"w-p-{wParaCount}\"></a>"); }
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

            // Check for inline section break (sectPr inside paragraph pPr) — handle column changes.
            // PAGE_BREAK + SECT advance are emitted at the TOP of the next
            // iteration so the section-closing paragraph is still attributed
            // to the section it terminates.
            if (element is Paragraph sectPara && sectPara.ParagraphProperties?.GetFirstChild<SectionProperties>() is SectionProperties inlineSectPr)
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
                // Drop cap wrapping (#7c): open non-flex wrapper on the
                // dropCap paragraph; close after the paragraph that follows.
                // Skip wrapping when para is a list item, heading, or empty —
                // Word's drop cap only applies to body paragraphs.
                var paraFramePr = para.ParagraphProperties?.GetFirstChild<FrameProperties>();
                var paraIsDropCap = paraFramePr != null &&
                    paraFramePr.GetAttributes().FirstOrDefault(a => a.LocalName == "dropCap").Value
                        is "drop" or "margin";
                if (paraIsDropCap && dropCapWrapRemaining == 0)
                {
                    sb.Append("<div class=\"dropcap-wrap\" style=\"display:block;overflow:hidden\">");
                    dropCapWrapRemaining = 2;
                }

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
                    // Resolve numPr through the pStyle chain so style-borne
                    // numbering (the canonical Heading1..9 pattern) renders
                    // identically to direct-numPr paragraphs.
                    var resolvedNumPr = ResolveNumPrFromStyle(para);
                    var ilvl = resolvedNumPr?.Ilvl ?? 0;
                    var numId = resolvedNumPr?.NumId ?? 0;
                    // Clamp ilvl to the OOXML-legal range [0, 8]. Malformed
                    // docs with huge ilvl (observed via raw-zip fuzz: 10000
                    // or Int32.MaxValue) otherwise explode the nested <ul>
                    // stack — crash on stack pop, or inflate HTML by 50× per
                    // paragraph (DoS). Negative values snap to 0 as well.
                    if (ilvl < 0) ilvl = 0;
                    else if (ilvl > 8) ilvl = 8;
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

                    // Seed per-level counter. Three-way precedence:
                    //   1. olCountPerLevel survives within the current <ol> stack.
                    //   2. lvlOverride/startOverride on this num → restart from value.
                    //   3. abstractNum-level running counter → continuation across
                    //      sibling num instances on the same abstractNum (the
                    //      `continue=true` path through the API; matches Word's
                    //      default "list continues from previous list using the
                    //      same template" behavior).
                    //   4. Otherwise, abstractNum's level start (typically 1).
                    var seedAbsId = GetAbstractNumId(numId);
                    int SeedStart(int forIlvl)
                    {
                        if (olCountPerLevel.TryGetValue(forIlvl, out var prev) && prev > 0)
                            return prev;
                        var ovr = GetNumStartOverride(numId, forIlvl);
                        if (ovr.HasValue) return ovr.Value - 1;
                        if (seedAbsId.HasValue
                            && absNumLevelCounters.TryGetValue(seedAbsId.Value, out var byIlvl)
                            && byIlvl.TryGetValue(forIlvl, out var running) && running > 0)
                            return running;
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
                        // Mirror the running count into the per-abstractNum
                        // store so a later sibling num on the same template
                        // can pick it up (continuation). Reset the deeper
                        // levels there too — Word resets all sub-levels when
                        // a shallower level ticks.
                        if (seedAbsId.HasValue)
                        {
                            if (!absNumLevelCounters.TryGetValue(seedAbsId.Value, out var byIlvl))
                            {
                                byIlvl = new Dictionary<int, int>();
                                absNumLevelCounters[seedAbsId.Value] = byIlvl;
                            }
                            byIlvl[ilvl] = olCountPerLevel[ilvl];
                            for (int lk = ilvl + 1; lk <= 8; lk++)
                                if (byIlvl.ContainsKey(lk)) byIlvl[lk] = 0;
                        }
                    }

                    currentListType = listStyle;
                    currentListLevel = ilvl;
                    currentNumId = numId;
                    sb.Append("<li");
                    sb.Append($" data-path=\"/body/p[{wParaCount}]\"");
                    // Marker class wires up the ::marker rule emitted by
                    // BuildListMarkerCss so this <li> picks up the abstractNum
                    // level rPr (color/font/size/bold/italic) for ul, plus
                    // a custom list-style-type string when applicable.
                    sb.Append($" class=\"marker-{numId}-{ilvl}\"");
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
                        // Pull in marker-level rPr (color/font/size/bold/italic) so
                        // the ol marker span matches the styling emitted globally
                        // for ul ::marker. Word lets per-level rPr restyle markers
                        // independent of the body run; mirroring that here keeps
                        // sections like "red bold 1." parallel between ol/ul.
                        var inlineMarkerCss = GetMarkerInlineCss(numId, ilvl);
                        var markerStyle = $"display:inline-block;min-width:{markerWidth};padding-right:{markerPadding};text-align:{align}";
                        if (!string.IsNullOrEmpty(inlineMarkerCss))
                            markerStyle = inlineMarkerCss + ";" + markerStyle;
                        sb.Append($"<span style=\"{markerStyle}\">{HtmlEncode(marker)}</span>");
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
                    var classNames = new List<string>();
                    if (paraStyleId != null && paraStyleId.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                        classNames.Add("toc");
                    // CONSISTENCY(run-special-content): body-path render must
                    // also flag has-ptab so the paragraph becomes a flex
                    // container — without this, body and table-cell ptabs
                    // collapse into a single line (only the header/footer
                    // render path went through RenderParagraphHtml which had
                    // the class added in Round 2).
                    if (para.Descendants<PositionalTab>().Any())
                        classNames.Add("has-ptab");
                    if (classNames.Count > 0)
                        sb.Append($" class=\"{string.Join(" ", classNames)}\"");
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
            else if (element is AltChunk altChunk)
            {
                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                RenderAltChunkHtml(sb, altChunk);
            }
        }

        // Close any pending block (last element was non-list with continue, or last list block)
        if (pendingBlockClose > 0) sb.Append($"<span class=\"we\" data-block=\"{pendingBlockClose}\" style=\"display:none\"></span>");
        if (inList) sb.Append($"<span class=\"we\" data-block=\"{wBlockCount}\" style=\"display:none\"></span>");
        if (inMultiColumn) sb.AppendLine("</div>");
        if (dropCapWrapRemaining > 0) sb.Append("</div>");
        CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
    }

    /// <summary>
    /// #6: a <c>&lt;w:p&gt;</c> whose only non-pPr child is an
    /// <c>&lt;m:oMathPara&gt;</c> is semantically a display-math block,
    /// not a text paragraph. Both <c>data-path="/body/p[N]"</c>
    /// attribution and Navigation.cs path resolution skip such wrappers
    /// so <c>/body/p[N]</c> counts only real prose paragraphs, while
    /// <c>/body/oMathPara[M]</c> addresses the equations separately.
    /// </summary>
    internal static bool IsOMathParaWrapperParagraph(Paragraph p)
    {
        var kids = p.ChildElements.Where(c => c is not ParagraphProperties).ToList();
        if (kids.Count != 1) return false;
        var only = kids[0];
        return only.LocalName == "oMathPara" || only is M.Paragraph;
    }

    /// <summary>
    /// #3: per-section header/footer bundle. Missing types fall back to
    /// the default variant at lookup time; missing default returns null
    /// so the legacy fallback can kick in.
    /// </summary>
    private record HeaderFooterBundle(string? First, string? Default, string? Even);

    /// <summary>
    /// #3: walk each section's HeaderReference or FooterReference elements,
    /// resolve to the underlying part, pre-render to HTML, and bucket by
    /// type. Returns a dict keyed by section index.
    /// </summary>
    private Dictionary<int, HeaderFooterBundle> BuildSectionHfBundles(
        List<SectionProperties> sections, bool isHeader)
    {
        var result = new Dictionary<int, HeaderFooterBundle>();
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return result;
        for (int i = 0; i < sections.Count; i++)
        {
            string? first = null, def = null, even = null;
            var refs = isHeader
                ? sections[i].Elements<HeaderReference>().Cast<OpenXmlElement>()
                : sections[i].Elements<FooterReference>().Cast<OpenXmlElement>();
            foreach (var @ref in refs)
            {
                var rId = @ref.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value;
                var typeAttr = @ref.GetAttributes().FirstOrDefault(a => a.LocalName == "type").Value;
                if (string.IsNullOrEmpty(rId)) continue;
                string? html = null;
                try
                {
                    if (isHeader && mainPart.GetPartById(rId) is HeaderPart hp && hp.Header != null
                        && HeaderFooterHasContent(hp.Header))
                    {
                        var sb = new StringBuilder();
                        sb.Append("<div class=\"doc-header\">");
                        RenderHeaderFooterBody(sb, hp.Header);
                        sb.Append("</div>");
                        html = sb.ToString();
                    }
                    else if (!isHeader && mainPart.GetPartById(rId) is FooterPart fp && fp.Footer != null
                        && HeaderFooterHasContent(fp.Footer))
                    {
                        var sb = new StringBuilder();
                        sb.Append("<div class=\"doc-footer\">");
                        RenderHeaderFooterBody(sb, fp.Footer);
                        sb.Append("</div>");
                        html = sb.ToString();
                    }
                }
                catch { /* part missing; skip */ }
                if (html == null) continue;
                switch (typeAttr)
                {
                    case "first": first = html; break;
                    case "even":  even = html; break;
                    default:      def = html; break;
                }
            }
            result[i] = new HeaderFooterBundle(first, def, even);
        }
        return result;
    }

    /// <summary>#3: pick the right header/footer variant for a given page.</summary>
    private static string PickHeaderFooter(
        Dictionary<int, HeaderFooterBundle> bundles,
        List<SectionProperties> sections,
        int sectionIdx,
        bool isFirstPageOfSection,
        bool pageIsEven,
        bool evenAndOddGlobal,
        string fallbackHtml)
    {
        if (!bundles.TryGetValue(sectionIdx, out var bundle))
            return fallbackHtml;
        var sectHasTitlePg = sectionIdx >= 0 && sectionIdx < sections.Count
            && sections[sectionIdx].GetFirstChild<TitlePage>() != null;
        // BUG-R22-01: when titlePg is set on the section, the first page of
        // the section uses strictly the "first" variant. If no first-type
        // reference is defined (bundle.First == null), Word renders a blank
        // header/footer on page 1 — do NOT fall through to Default, which
        // would show the wrong content.
        if (isFirstPageOfSection && sectHasTitlePg)
            return bundle.First ?? string.Empty;
        if (evenAndOddGlobal && pageIsEven && bundle.Even != null)
            return bundle.Even;
        return bundle.Default ?? fallbackHtml;
    }

    /// <summary>
    /// #8a: update <see cref="HtmlRenderContext.FnRestartEachSection"/> and
    /// reset the per-section counter when a section with
    /// <c>&lt;w:footnotePr&gt;&lt;w:numRestart w:val="eachSect"/&gt;</c>
    /// begins. Called from RenderBodyHtml at every SECT marker emit.
    /// </summary>
    private void ApplySectionFnSettings(List<SectionProperties> sections, int idx)
    {
        _ctx.CurrentSectionIdx = idx;
        if (idx < 0 || idx >= sections.Count) return;
        var sectPr = sections[idx];
        var fnPr = sectPr.GetFirstChild<FootnoteProperties>();
        var restart = fnPr?.GetFirstChild<NumberingRestart>()?.Val?.InnerText;
        var eachSect = restart == "eachSect";
        if (eachSect)
        {
            _ctx.FnRestartEachSection = true;
            _ctx.FnCountInSection = 0;
        }
        else
        {
            _ctx.FnRestartEachSection = false;
        }
    }

    /// <summary>
    /// #8b: emit the alternate content referenced by a <c>&lt;w:altChunk&gt;</c>
    /// relationship. text/html is injected (with <c>&lt;script&gt;</c> tags
    /// stripped); text/plain is wrapped in <c>&lt;pre&gt;</c>; RTF and
    /// other binary-ish formats fall back to a stripped-text placeholder.
    /// Opens the door to rendering HTML fragments authors embed in Word
    /// via "Insert File → HTML" instead of rendering a blank gap.
    /// </summary>
    private void RenderAltChunkHtml(StringBuilder sb, AltChunk altChunk)
    {
        var rId = altChunk.Id?.Value;
        if (string.IsNullOrEmpty(rId)) return;
        try
        {
            var part = _doc.MainDocumentPart?.GetPartById(rId)
                       as AlternativeFormatImportPart;
            if (part == null) return;
            using var stream = part.GetStream();
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();
            var contentType = (part.ContentType ?? "").ToLowerInvariant();
            // Strip media-type parameters (e.g. "text/html; charset=utf-8")
            // before comparison: Pandoc/non-Word authors commonly emit them.
            var mediaType = contentType.Split(';', 2)[0].Trim();

            if (mediaType is "text/html" or "application/xhtml+xml"
                || mediaType.EndsWith("+xml") && mediaType.Contains("xhtml"))
            {
                // Regex-based HTML sanitization has too many bypasses:
                // unclosed <script>, HTML-entity-encoded javascript: URLs,
                // case-mangled <StYlE>, style="background:url(javascript:)"
                // etc. Since we can't guarantee safety against an
                // adversarial altChunk author, render the HTML payload as
                // escaped text instead so nothing ever enters the DOM as
                // live HTML. Callers that need rich inline HTML should use
                // Word's native insert-content features, not altChunk.
                var bodyMatch = Regex.Match(content,
                    @"<body[^>]*>(.*?)</body>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);
                var inner = bodyMatch.Success ? bodyMatch.Groups[1].Value : content;
                sb.AppendLine(
                    $"<pre class=\"alt-chunk-html-escaped\" " +
                    $"style=\"white-space:pre-wrap;background:#f7f7f7;padding:8px;border:1px dashed #bbb;\">" +
                    $"{HtmlEncode(inner)}</pre>");
            }
            else if (mediaType is "text/plain" or "text/css")
            {
                sb.AppendLine($"<pre class=\"alt-chunk-text\">{HtmlEncode(content)}</pre>");
            }
            else
            {
                // RTF etc.: strip control words and braces, emit as plain-text block.
                var plain = Regex.Replace(content, @"\\[a-zA-Z]+-?\d*\s?|[{}]", " ");
                plain = Regex.Replace(plain, @"\s+", " ").Trim();
                if (plain.Length > 1000) plain = plain[..1000] + "…";
                sb.AppendLine(
                    $"<div class=\"alt-chunk-fallback\" " +
                    $"style=\"border:1px dashed #bbb;padding:4px;font-style:italic;color:#555\">" +
                    $"{HtmlEncode(plain)}</div>");
            }
        }
        catch
        {
            // Silent skip: altChunk part missing / unreadable shouldn't break the whole preview.
        }
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
