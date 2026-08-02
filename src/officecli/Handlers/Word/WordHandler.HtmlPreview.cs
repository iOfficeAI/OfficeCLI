// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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

        // Image-relationship host part for the element currently being
        // rendered. Body content resolves r:embed against MainDocumentPart,
        // but a header/footer image's ImagePart + rel live on the
        // HeaderPart/FooterPart, so LoadImageAsDataUri must look there.
        // null → fall back to MainDocumentPart (body path).
        public DocumentFormat.OpenXml.Packaging.OpenXmlPart? ImageHostPart { get; set; }

        // Table-style run properties (base rPr + matching conditional-format
        // rPr) for the cell currently being rendered, ordered lowest→highest
        // priority. Per ECMA-376 §17.7.2 the run-property cascade is:
        // docDefaults → table styles → paragraph styles → … → run direct.
        // ResolveEffectiveRunPropertiesCore merges these layers in just after
        // docDefaults so a firstRow/band cell's <w:caps/> + white <w:color/>
        // reach the run (otherwise the run inherits only docDefaults color,
        // e.g. the Invoice "PAYMENT OPTIONS" header rendering lowercase grey).
        // null/empty → not inside a styled-table cell (body path). Saved and
        // restored around RenderCellChild like ImageHostPart. Each layer is an
        // rPr-shaped element (a style's <w:rPr> / a tblStylePr's <w:rPr>) merged
        // as a source via MergeRunProperties.
        public List<DocumentFormat.OpenXml.OpenXmlElement>? CurrentCellTableStyleRunProps { get; set; }

        // CJK line-break tracking: accumulate character widths and insert <br> at Word-compatible positions
        public double LineWidthPt { get; set; }      // available width for current line
        public double LineAccumPt { get; set; }       // accumulated width on current line
        public bool LineBreakEnabled { get; set; }    // whether line-break tracking is active
        public double DefaultFontSizePt { get; set; } // default font size for width estimation

        // Tab positioning: count tabs seen in current paragraph to look up Nth tab stop.
        // Reset per paragraph in RenderParagraphContentHtml.
        public int CurrentParagraphTabIndex { get; set; }

        // Tab absolute-alignment: sb offset where the current tab segment's
        // leading text begins (start of paragraph content, or just after the
        // previous positional tab's wrapper). A positional tab retro-wraps the
        // text from this offset in a fixed-width container so the following
        // text lands at the absolute tab position regardless of leading-text
        // length. -1 means "not tracking" (reset per paragraph).
        public int CurrentParagraphTabSegmentStart { get; set; } = -1;

        // Tab alignment band model: true while rendering a paragraph whose tab
        // stops include a center/right (no-leader) positional stop — the
        // three-part "Left \t Center \t Right" header shape. In this mode each
        // <w:tab/> closes the current leading text in a flex band and the
        // upcoming stop's Val decides the band's text-align, so segments stay
        // on one line and land left/centre/right (see has-aligned-tab CSS).
        // Reset per paragraph.
        public bool CurrentParagraphAlignedTab { get; set; }

        // CSS text-align for the CURRENT aligned-tab band (the text typed before
        // the next <w:tab/>). Band 0 (before the first tab) is left-aligned; each
        // tab sets this to its own Val (center/right) for the band it opens.
        public string CurrentAlignedTabAlign { get; set; } = "left";

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
    /// <param name="gridCols">When &gt; 0, render every page tiled into a
    /// thumbnail contact-sheet grid this many columns wide (screenshot mode).
    /// Mirrors pptx's HTML grid; <paramref name="pageFilter"/> is ignored
    /// (the grid always shows the whole document).</param>
    /// <param name="gridCellWpx">Exact thumbnail cell width in CSS px. The CLI
    /// computes this from the viewport width and column count so the C# height
    /// math and the in-browser layout agree exactly.</param>
    public string ViewAsHtml(string? pageFilter = null, int gridCols = 0, int gridCellWpx = 0)
    {
        using var _cul = InvariantCultureScope.Enter();
        try
        {
            return ViewAsHtmlCore(pageFilter, gridCols, gridCellWpx);
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

    private string ViewAsHtmlCore(string? pageFilter, int gridCols = 0, int gridCellWpx = 0)
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
        // i18n: emit lang from themeFontLang/docDefaults (ResolveThemeCjkFont
        // populates _eastAsiaLang) and dir="rtl" when any section carries
        // <w:bidi/>, so browsers activate the correct BiDi layout, default
        // text direction, and font/hyphenation heuristics. Falls back to
        // lang="en" with no dir for plain Latin documents. EastAsia covers
        // ja/zh/ko; Bidi covers ar/he/fa/ur/th/hi (read directly here
        // since _eastAsiaLang only carries the EA slot).
        string? htmlLangVal = _eastAsiaLang;
        if (string.IsNullOrEmpty(htmlLangVal))
        {
            try
            {
                var settingsForLang = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
                var tfl = settingsForLang?.Descendants<ThemeFontLanguages>().FirstOrDefault();
                htmlLangVal = tfl?.Bidi?.Value ?? tfl?.Val?.Value;
            }
            catch (System.Xml.XmlException) { }
        }
        var htmlLang = string.IsNullOrEmpty(htmlLangVal) ? "en" : htmlLangVal!;
        var docHasBidi = body.Descendants<SectionProperties>()
            .Any(sp => sp.GetFirstChild<BiDi>() != null);
        var dirAttr = docHasBidi ? " dir=\"rtl\"" : "";
        sb.AppendLine($"<html lang=\"{HtmlEncode(htmlLang)}\"{dirAttr}>");
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
            // Web-font loading: detect-then-inject, one request per family.
            //
            // Local-first: a font that resolves locally must NOT be replaced by
            // Google's metric-compatible substitute. A plain stylesheet <link>
            // can't express that — its @font-face parses after ours (same
            // family + descriptors → last wins) and Google's css2 dropped
            // local() fallbacks in 2020, so a successful load would shadow the
            // genuine local font. Instead a tiny inline script canvas-measures
            // each family against monospace AND serif fallbacks (dual-fallback
            // so a font that IS the platform default for one generic family is
            // still detected via the other) and only injects a css2 <link> for
            // families that are genuinely absent locally.
            //
            // Per-family requests: css2 is all-or-nothing — one unknown family
            // 400s the whole batch, so a 宋体+Open Sans document used to lose
            // the perfectly servable Open Sans too. Split links fail (or hit)
            // independently; a family Google doesn't carry costs one async 400
            // and nothing else. Known-never-on-Google names are deliberately
            // NOT filtered here: the skip would freeze today's Google catalog
            // into the binary, and a doomed request is harmless post-split.
            //
            // CONSISTENCY(katex-mirror): officecli's caching proxy first
            // (d.officecli.ai/fonts/css2 forwards to Google Fonts, rewrites the
            // font URLs to the /gstatic hop, and CF-edge-caches both), Google
            // direct as the onerror fallback, removal as the last resort — the
            // local metric-override @font-face block above keeps layout stable
            // without any web font (also the no-JS degradation).
            var webFontCandidates = docFonts.Where(f =>
                    !f.Equals("Arial", StringComparison.OrdinalIgnoreCase)
                    && !f.Equals("Times New Roman", StringComparison.OrdinalIgnoreCase)
                    && !f.Equals("Tahoma", StringComparison.OrdinalIgnoreCase)
                    && !f.Equals("Courier New", StringComparison.OrdinalIgnoreCase)
                    && !f.StartsWith("Symbol") && !f.StartsWith("Wingding"))
                .Select(SanitizeFontName)
                .Where(f => !string.IsNullOrEmpty(f))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (webFontCandidates.Count > 0)
            {
                // Sanitized names contain only letters/digits/space/.-_ so plain
                // quote-wrapping is JSON-safe (no escapable characters survive).
                var famJson = "[" + string.Join(",", webFontCandidates.Select(f => $"\"{f}\"")) + "]";
                sb.AppendLine("<script>(function(){");
                sb.AppendLine($"var fams={famJson};");
                sb.AppendLine("function w(f,fb){var c=w.c||(w.c=document.createElement('canvas').getContext('2d'));c.font='72px '+(f?'\"'+f+'\",':'')+fb;return c.measureText('mmmmmmmmmmlliWW@#\\u5e05\\u6c38\\u306e\\ud55c').width;}");
                sb.AppendLine("function has(f){try{return w(f,'monospace')!==w(null,'monospace')||w(f,'serif')!==w(null,'serif');}catch(e){return true;}}");
                sb.AppendLine("var ax=':ital,wght@0,400;0,700;1,400;1,700&display=swap';");
                sb.AppendLine("fams.forEach(function(f){if(has(f))return;");
                sb.AppendLine("var q='family='+encodeURIComponent(f).replace(/%20/g,'+')+ax;");
                sb.AppendLine("var l=document.createElement('link');l.rel='stylesheet';l.href='https://d.officecli.ai/fonts/css2?'+q;");
                sb.AppendLine("l.onerror=function(){if(!l.dataset.f){l.dataset.f=1;l.href='https://fonts.googleapis.com/css2?'+q;}else{l.remove();}};");
                sb.AppendLine("document.head.appendChild(l);});");
                sb.AppendLine("})();</script>");
            }
        }
        // KaTeX for math rendering — only include when the document actually has formulas.
        // Same non-blocking load trick so KaTeX CSS can never stall first paint.
        bool hasMathFormulas = body.Descendants<M.OfficeMath>().Any();
        if (hasMathFormulas)
        {
            // CONSISTENCY(katex-mirror): mirror-first with CDN fallback chain — see Core/KatexAssets.
            sb.AppendLine($"<link rel=\"stylesheet\" href=\"{Core.KatexAssets.CssUrl}\" media=\"print\" onload=\"this.media='all'\" onerror=\"{Core.KatexAssets.CssOnErrorJs}\">");
            sb.AppendLine($"<script defer src=\"{Core.KatexAssets.JsUrl}\" onerror=\"{Core.KatexAssets.JsOnErrorJs("document.querySelectorAll('.katex-formula').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666'})")}\"></script>");
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
        var sectionHeaders = BuildSectionHfBundles(allSectionsForHf, isHeader: true,
            out var sectionHeaderFields);
        var sectionFooters = BuildSectionHfBundles(allSectionsForHf, isHeader: false,
            out var sectionFooterFields);
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
        // PAGE/NUMPAGES presence for the fallback parts — RenderHeaderFooterHtml
        // emits only the FIRST content-bearing part, so probe the same one for
        // its field flags. The digit-rewrite below is gated on these so a
        // literal header/footer number is never mistaken for a page field.
        var fallbackHeaderFields = FirstContentHeaderFooterFlags(isHeader: true);
        var fallbackFooterFields = FirstContentHeaderFooterFlags(isHeader: false);

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

        // Replace the rendered PAGE / NUMPAGES result with a per-page
        // substitution placeholder — but ONLY when the source footer part truly
        // defines that field. Matching any digit-only run (the old behavior)
        // clobbered literal numbers (a year "2014", a phone number, a price)
        // that happened to be the first digit run, corrupting visible content.
        // The tag name (span vs p) depends on whether the run carries rPr
        // styling; ApplyPageNumFields handles both.
        var footerTemplate = ApplyPageNumFields(
            footerHtml, fallbackFooterFields.page, fallbackFooterFields.numPages);

        // Section-level multi-column layout: w:cols num=N sep=true.
        // BUG(first-section-cols): in a multi-section doc the page-body's initial
        // column layout belongs to SECTION 1, whose sectPr is the first inline
        // <w:sectPr> on a paragraph — NOT body.GetFirstChild<SectionProperties>(),
        // which is the LAST section's (trailing body) sectPr. Using the trailing
        // one rendered section 1 with the final section's column count.
        // CollectSections returns sections in document order (inline sectPr's
        // first, trailing body sectPr last), so [0] is section 1. Single-section
        // docs only have the trailing body sectPr → [0] is that, preserving the
        // original behavior.
        var firstSection = CollectSections(body).FirstOrDefault()
            ?? _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        // CSS columns need a height floor to balance — with no height the body
        // is unbounded so all content stacks in column 1 and overflows the page.
        // BuildColBodyStyle applies this as `min-height` (not fixed `height`) so
        // a section whose content exceeds two columns can grow the box downward
        // instead of overprinting a wrapped-back third column (BUG(cols-overprint,
        // R85)). Use the doc-level pgLayout body height as the floor.
        var colBodyHeightPt = pgLayout.HeightPt - pgLayout.MarginTopPt - pgLayout.MarginBottomPt;

        // Per-section page layout (#7a00): each page carries one or more
        // <!--SECT:N--> markers inserted by RenderBodyHtml. The last marker
        // seen (inclusive of this page) decides the page's size/margins;
        // pages with no marker inherit from the previous page.
        var sections = CollectSections(body);
        var sectRegex = new Regex(@"<!--SECT:(\d+)-->");
        var activeLayout = pgLayout;
        // Document-level page background (<w:background w:color="RRGGBB"/>).
        // Real Word fills the whole page area; emit background-color on the
        // .page div behind body/margins. Color is ST_HexColor (bare RRGGBB).
        string pageBgCss = "";
        var docBg = _doc.MainDocumentPart?.Document?.GetFirstChild<DocumentBackground>();
        if (docBg?.Color?.Value is { Length: > 0 } bgColor
            && !bgColor.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            pageBgCss = $"background-color:{ParseHelpers.FormatHexColor(bgColor)};";
        }
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
            // BUG(trailing-continuous-cols, R87): remember the FIRST section
            // marker on this page before the markers are stripped below — the
            // page-body column decision needs it to detect a fewer-col→multi-col
            // continuous transition within one page (trailing 2-col leaking onto
            // earlier 1-col content). -1 = no marker (inherits previous page).
            int firstSectIdxOnPage = sectMatches.Count > 0
                ? int.Parse(sectMatches[0].Groups[1].Value)
                : -1;
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
                    //
                    // Page number format does NOT inherit across sections:
                    // each section's w:fmt is independent, defaulting to
                    // "decimal" (ECMA-376 §17.6.12). A section with no
                    // explicit w:fmt (or no pgNumType) must RESET to decimal
                    // rather than keep the previous section's format —
                    // otherwise a sect1 lowerRoman leaks into a sect2 body
                    // that should number 4,5,6 (not iv,v,vi).
                    displayedFmt = pgNumType?.Format?.InnerText is { Length: > 0 } fmtStr
                        ? fmtStr
                        : "decimal";
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
                $"{activeLayout.MarginLeftPt.ToString("0.#", ci)}pt;" +
                pageBgCss;
            // Page border (<w:pgBorders> in the active section's sectPr).
            // Real Word boxes the whole page; emit per-side border on the
            // .page div so partial (some-sides-only) borders also work.
            if (activeSectionIdx >= 0 && activeSectionIdx < sections.Count)
                pageStyle += BuildPageBorderCss(sections[activeSectionIdx]);
            else if (sections.Count > 0)
                pageStyle += BuildPageBorderCss(sections[^1]);
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
            // rendering the author-time cached literal "1" — but only when the
            // picked header variant truly carries the field (gate against
            // literal header numbers).
            var hdrFlags = PickHeaderFooterFlags(
                sectionHeaders, sectionHeaderFields, sections, activeSectionIdx,
                isFirstPageOfSection, pageIsEven, evenAndOddGlobal, fallbackHeaderFields);
            var perPageHeaderTemplate = ApplyPageNumFields(
                perPageHeader, hdrFlags.page, hdrFlags.numPages);
            sb.Append(perPageHeaderTemplate
                .Replace("<!--PAGE_NUM-->", hdrPageNumStr)
                .Replace("<!--NUM_PAGES-->", pageList.Count.ToString()));
            // BUG(multi-section-cols): column style must come from the section
            // active on THIS page, not section 1. Earlier this used a single
            // pre-loop colBodyStyle from firstSection, so every page inherited
            // section 1's column count.
            var colSectionForPage = (activeSectionIdx >= 0 && activeSectionIdx < sections.Count)
                ? sections[activeSectionIdx]
                : firstSection;
            // BUG(trailing-continuous-cols, R87): when this page CONTAINS a
            // continuous section transition from a fewer-col section into a
            // multi-col one (e.g. a 1-col paper body whose trailing body sectPr
            // is 2-col References), the multi-col content is already wrapped in
            // its own scoped <div column-count> by RenderBodyHtml. Applying the
            // active (last) section's multi-col to the WHOLE page-body here would
            // reverse-leak it onto the earlier 1-col content. So when the page's
            // FIRST section has fewer columns than the active section, the
            // page-body keeps the FIRST section's column count and lets the inner
            // scoped wrapper own the multi-col region. Single-section pages (first
            // marker == active) and inline-multi-col pages (R85, active is already
            // the multi-col inline section appearing first on its page) are
            // unaffected.
            if (firstSectIdxOnPage >= 0 && firstSectIdxOnPage < sections.Count
                && firstSectIdxOnPage != activeSectionIdx
                && GetSectionColumnCount(sections[firstSectIdxOnPage])
                     < GetSectionColumnCount(colSectionForPage))
            {
                colSectionForPage = sections[firstSectIdxOnPage];
            }
            var colBodyStyle = BuildColBodyStyle(colSectionForPage, colBodyHeightPt);
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
            // Rebuild the PAGE field placeholder on the picked footer — gated to
            // footers that actually carry the field.
            var ftrFlags = PickHeaderFooterFlags(
                sectionFooters, sectionFooterFields, sections, activeSectionIdx,
                isFirstPageOfSection, pageIsEven, evenAndOddGlobal, fallbackFooterFields);
            var perPageFooterTemplate = ApplyPageNumFields(
                perPageFooter, ftrFlags.page, ftrFlags.numPages);
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
        // Top margin (= padding-top of .page, the w:top reserve below which the
        // body starts) and header distance, in px — used by adjustHeaderPadding
        // to detect when header content is taller than the top margin and must
        // push the body down (Word behaviour) rather than overlap it.
        sb.AppendLine($"  var topMarginPx={pgLayout.MarginTopPt:0.#}*96/72;");
        sb.AppendLine($"  var headerDistPx={pgLayout.HeaderDistancePt:0.#}*96/72;");
        sb.AppendLine("  var ftpl=" + JsStringLiteral(footerTemplate) + ";");
        // Header template cloned per paginated page. Continuation pages (2+)
        // never carry the first-page (titlePg) header — use the section's
        // DEFAULT header so a first-page header doesn't bleed onto later
        // pages. fallbackHeaderHtml takes arbitrary part order and may be the
        // first-page variant; prefer the resolved Default of section 0.
        var useSectionDefaultHeader = sectionHeaders.TryGetValue(0, out var hb0) && hb0.Default != null;
        var contHeaderHtml = useSectionDefaultHeader ? hb0!.Default! : fallbackHeaderHtml;
        // Flags for the continuation header — from section 0's Default when used,
        // else the fallback part. Keeps the rewrite gated to real PAGE fields.
        var contHeaderFlags = useSectionDefaultHeader
            && sectionHeaderFields.TryGetValue(0, out var hf0)
            ? hf0.Default
            : fallbackHeaderFields;
        // Mirror the footer (ftpl) template exactly: emit named
        // page-num-field / num-pages-field spans rather than bare comments.
        // The JS replaces <!--PAGE_NUM--> per page and the renumber loop fills
        // both .page-num-field and .num-pages-field spans — a bare
        // <!--NUM_PAGES--> comment in the header was never substituted, so
        // continuation-page headers showed an empty NUMPAGES count.
        var headerTemplate = ApplyPageNumFields(
            contHeaderHtml, contHeaderFlags.page, contHeaderFlags.numPages);
        sb.AppendLine("  var htpl=" + JsStringLiteral(headerTemplate) + ";");
        // Even-page header/footer continuation templates. When the document
        // has evenAndOddHeaders enabled AND section 0 defines an even-type
        // header/footer part, JS-created continuation pages whose final page
        // number is even must clone the EVEN template instead of the odd
        // (htpl/ftpl) one — mirroring the server-side PickHeaderFooter parity
        // logic for the first server-rendered page. Emitted as empty strings
        // otherwise so the JS picker transparently falls back to htpl/ftpl.
        string evenHeaderTemplate = "", evenFooterTemplate = "";
        if (evenAndOddGlobal)
        {
            if (sectionHeaders.TryGetValue(0, out var ehb) && ehb.Even != null)
            {
                var ef = sectionHeaderFields.TryGetValue(0, out var ehf) ? ehf.Even : (page: false, numPages: false);
                evenHeaderTemplate = ApplyPageNumFields(ehb.Even, ef.page, ef.numPages);
            }
            if (sectionFooters.TryGetValue(0, out var efb) && efb.Even != null)
            {
                var ef = sectionFooterFields.TryGetValue(0, out var eff) ? eff.Even : (page: false, numPages: false);
                evenFooterTemplate = ApplyPageNumFields(efb.Even, ef.page, ef.numPages);
            }
        }
        sb.AppendLine("  var etpl=" + JsStringLiteral(evenHeaderTemplate) + ";");
        sb.AppendLine("  var eftpl=" + JsStringLiteral(evenFooterTemplate) + ";");
        // Pick the header template for a 1-based page number: even pages use
        // the even template when present, all others use the odd/default htpl.
        sb.AppendLine("  function pickHtpl(n){return (etpl&&n%2===0)?etpl:htpl;}");
        sb.AppendLine("  function pickFtpl(n){return (eftpl&&n%2===0)?eftpl:ftpl;}");
        sb.AppendLine(@"
  // Out-of-flow children (position:absolute/fixed — full-page background
  // layers, behind-text watermarks, floating anchored drawings) are removed
  // from the normal document flow and do NOT occupy vertical space for the
  // content that follows them. A full-page background div
  // (position:absolute;height:100%) reports offsetHeight==page-height, so if
  // the pagination height math counted it as flow content it would think the
  // hosting paragraph is a full page tall and push the real content (title,
  // body) onto subsequent pages, splitting the cover and shifting the whole
  // document down. Skip these elements when measuring flow height / picking
  // split points so they paint where the renderer placed them without
  // displacing in-flow siblings.
  function isOutOfFlow(el){
    if(!el||el.nodeType!==1)return false;
    var pos=getComputedStyle(el).position;
    return pos==='absolute'||pos==='fixed';
  }
  // A header child is a full-page background / watermark layer (paint-behind,
  // must NOT push the body down — R47 full-bleed cover, behind-text watermark)
  // when it spans ~the whole page. An ordinary header logo (inline OR anchored
  // out-of-flow) is far smaller and MUST count toward header height so the body
  // clears it. Distinguishing by SIZE (not merely position:absolute) is the
  // R120 fix: the previous isOutOfFlow gate skipped EVERY absolute child, so an
  // anchored/floated header logo reported zero header height and the body
  // overlapped (then hid) it. We keep excluding the genuine full-page layers.
  function isFullPageHeaderLayer(c,page){
    if(c.classList&&(c.classList.contains('vml-watermark-layer')))return true;
    var pr=page.getBoundingClientRect();
    var cr=c.getBoundingClientRect();
    // ≥85% of page width AND height → a full-bleed background/cover layer.
    return cr.width>=pr.width*0.85 && cr.height>=pr.height*0.85;
  }
  // Tall-header reflow: Word pushes the body down when the header content is
  // taller than the top margin lets it sit above the body. Our .doc-header is
  // position:absolute (R47, so a full-bleed cover banner can paint behind the
  // body without reserving space), which means normal-flow header content
  // (logo + title, an address block) overlaps the body's first paragraph when
  // it exceeds the top margin. Measure each header's REQUIRED bottom from its
  // children, excluding only genuine full-page background / watermark layers
  // (isFullPageHeaderLayer) — NOT every absolute child, so an anchored logo
  // still counts (R120). The full-bleed cover stays behind the body (no push),
  // exactly the R47 design. When the header bottom exceeds the page's current
  // top padding, grow the page's padding-top so the flex .page-body starts
  // below the header. Geometry is read via getBoundingClientRect relative to
  // the page top so it is immune to offsetParent quirks. Idempotent: re-derives
  // the needed padding from the header geometry each call (base = topMarginPx).
  function adjustHeaderPadding(page){
    var hdr=page.querySelector('.doc-header');
    if(!hdr)return;
    var pageTop=page.getBoundingClientRect().top;
    var contentBottom=0;
    Array.from(hdr.children).forEach(function(c){
      if(isFullPageHeaderLayer(c,page))return; // skip full-page bg / watermark
      var b=c.getBoundingClientRect().bottom-pageTop; // bottom relative to page top
      if(b>contentBottom)contentBottom=b;
    });
    if(contentBottom<=0)return;
    // Small bottom gap so body doesn't butt against the header's last line.
    var needed=contentBottom+(headerDistPx*0.5);
    var base=topMarginPx;
    var pad=needed>base?needed:base;
    page.style.paddingTop=pad+'px';
  }
  function paginate(){
    // Reflow tall headers BEFORE measuring body overflow so the pushed-down
    // body height is accounted for when picking split points.
    document.querySelectorAll('.page').forEach(adjustHeaderPadding);
    var pages=document.querySelectorAll('.page');
    // Sync mode + page filter: bail once pages beyond max-requested exist
    // and pages 1..maxReq are stable. Avoids paginating 100-page docs to
    // completion when only the first few were asked for.
    var loopLim=pages.length;
    if(window._wpSync&&window._requestedPages&&window._requestedPages.length){
      var mxR=Math.max.apply(null,window._requestedPages);
      loopLim=Math.min(pages.length,mxR+1);
      if(pages.length>mxR){
        var settled=true;
        for(var sk=0;sk<mxR;sk++){
          var sp=pages[sk];var sb_=sp.querySelector('.page-body');
          if(!sb_)continue;
          var sf=sb_.querySelector('.footnotes');var sfH=sf?sf.offsetHeight:0;
          var sch=0;
          Array.from(sb_.children).forEach(function(c){
            if(c.classList.contains('footnotes'))return;
            if(isOutOfFlow(c))return;
            var bt=c.offsetTop+c.offsetHeight-sb_.offsetTop;
            if(bt>sch)sch=bt;
          });
          if(sch>maxBodyH-sfH+2){settled=false;break;}
        }
        if(settled){positionFootnotes();wrapFloats();applyLineNumbers();applyPageFilter();return;}
      }
    }
    for(var pi=0;pi<loopLim;pi++){
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
        if(isOutOfFlow(c))return;
        var b=c.offsetTop+c.offsetHeight-body.offsetTop;
        if(b>contentH)contentH=b;
      });
      if(contentH<=availH+2)continue;
      // Find first child that overflows available space. Use the same +2px
      // tolerance the outer check uses — Chrome rounds line-box height to
      // logical pixels at dsf>=2, accumulating ~0.13pt/line of drift, so
      // a paragraph straddling the boundary by <2px shouldn't get pushed
      // to the next page when the outer overflow check already accepted
      // the page.
      var children=Array.from(body.children);
      var splitIdx=-1;
      for(var ci=0;ci<children.length;ci++){
        if(children[ci].classList.contains('footnotes'))continue;
        if(isOutOfFlow(children[ci]))continue;
        var bot=children[ci].offsetTop+children[ci].offsetHeight-body.offsetTop;
        if(bot>availH+2){splitIdx=ci;break;}
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
      // Mirror the table case for <ol>/<ul>: split at the first <li>
      // that overflows so partial lists carry to the next page rather
      // than the whole list moving as one atomic unit. Multi-level
      // numbering is rendered via nested <ol>/<ul> wrapped inside the
      // parent <li>; recurse through that nesting to find the
      // shallowest list whose splitting keeps as much content on this
      // page as possible. Cumulative left-padding from the parent
      // chain is folded into the promoted continuation list so the
      // continuation keeps its visual indent on the new page.
      if(firstOverflow&&(firstOverflow.tagName==='OL'||firstOverflow.tagName==='UL')){
        var findListSplit=function(lst,parents){
          var lis=Array.from(lst.children).filter(function(c){return c.tagName==='LI';});
          for(var li=0;li<lis.length;li++){
            var liBot=lis[li].offsetTop+lis[li].offsetHeight-body.offsetTop;
            if(liBot>availH+2){
              if(li>0)return {parents:parents,list:lst,splitAt:li,moveAll:false};
              var nested=Array.from(lis[li].children).find(function(c){return c.tagName==='OL'||c.tagName==='UL';});
              if(nested){
                var nestedTop=nested.offsetTop-body.offsetTop;
                if(nestedTop<=availH+2){
                  var newParents=parents.concat([{list:lst,li:lis[li]}]);
                  var inner=findListSplit(nested,newParents);
                  if(inner)return inner;
                  return {parents:newParents,list:nested,splitAt:0,moveAll:true};
                }
              }
              return null;
            }
          }
          return null;
        };
        var sp=findListSplit(firstOverflow,[]);
        if(sp){
          var srcLis=Array.from(sp.list.children).filter(function(c){return c.tagName==='LI';});
          var cont=sp.list.cloneNode(false);
          for(var lj=sp.splitAt;lj<srcLis.length;lj++){
            cont.appendChild(srcLis[lj]);
          }
          // Promote `cont` to a body-level sibling. Each parent <li>
          // would have contributed its own list's padding-left, so
          // sum the live padding-left of every list in the parent
          // chain (plus the split list itself) and apply that to the
          // promoted wrapper. Reading getComputedStyle keeps us in
          // sync with whatever the renderer emits per-level — no
          // assumed indent constant.
          if(sp.parents.length>0){
            var cumPadPx=0;
            for(var pi2=0;pi2<sp.parents.length;pi2++){
              cumPadPx+=parseFloat(getComputedStyle(sp.parents[pi2].list).paddingLeft)||0;
            }
            cumPadPx+=parseFloat(getComputedStyle(sp.list).paddingLeft)||0;
            cont.style.paddingLeft=cumPadPx+'px';
            // If the split list is now empty (moveAll case), drop it
            // so the source page doesn't render an empty bullet gutter.
            if(sp.moveAll&&sp.list.children.length===0){
              sp.list.parentNode.removeChild(sp.list);
            }
          }
          firstOverflow.parentNode.insertBefore(cont,firstOverflow.nextSibling);
          children=Array.from(body.children);
          splitIdx=children.indexOf(cont);
        }
      }
      // When the first child itself exceeds page height, keep it on this
      // page and split after, so the oversized element is not silently
      // dropped by being moved to a new (still-oversized) page.
      if(splitIdx===0)splitIdx=1;
      // Sync mode: greedy multi-split — walk children once, find ALL break
      // points based on cumulative offsetTop/Height, and create N new pages
      // in a single pass. Avoids the O(N) layout reflows the async path
      // pays via recursive single-splits. Async path retains the original
      // single-split behaviour (setTimeout chain stays cheap).
      if(window._wpSync){
        // table-row-split above may have refreshed children — re-collect.
        children=Array.from(body.children);
        var bodyT=body.offsetTop;
        var movable=[];
        for(var mi=splitIdx;mi<children.length;mi++){
          if(children[mi].classList.contains('footnotes'))continue;
          // Out-of-flow layers (full-page background, watermark, floating
          // drawing) are positioned relative to THIS .page; leave them on the
          // source page and don't let them participate in split height math.
          if(isOutOfFlow(children[mi]))continue;
          movable.push({el:children[mi],top:children[mi].offsetTop-bodyT,h:children[mi].offsetHeight});
        }
        if(movable.length===0)continue;
        // Greedy bin-pack: each segment-leading element becomes :first-child
        // of its new page-body. New bodies get class=page-body-cont so the
        // CSS rule
        //   .page-body-cont > :first-child { margin-top: 0 !important; }
        // zeroes the leader's margin-top contribution. (The original page-body
        // is class=page-body without -cont, so its first paragraph keeps its
        // server-rendered spaceBefore — matching Word.) Since the leader's
        // old offsetTop already absorbs its own margin-top (it wasn't
        // first-child of source body), and inter-sibling distances are
        // preserved across the move
        // (flex items don't collapse), bot-segStartTop with
        // segStartTop=leader.oldTop directly equals the leader's newBottom in
        // the new page-body. No margin-top shift needed.
        var splits=[0];
        var segStartTop=movable[0].top;
        for(var i=1;i<movable.length;i++){
          var bot=movable[i].top+movable[i].h;
          if(bot-segStartTop>availH+2){
            splits.push(i);
            segStartTop=movable[i].top;
          }
        }
        // splits = [0, k1, k2, ...]: segment s holds movable[splits[s]..splits[s+1]-1]
        var prevWrapper=page.closest('.page-wrapper')||page;
        for(var s=0;s<splits.length;s++){
          var segStart=splits[s];
          var segEnd=s+1<splits.length?splits[s+1]:movable.length;
          var nw=document.createElement('div');
          nw.className='page-wrapper';
          var np=document.createElement('div');
          np.className='page';
          np.style.cssText=page.style.cssText;
          var nb=document.createElement('div');
          nb.className='page-body page-body-cont';
          for(var mi=segStart;mi<segEnd;mi++){
            nb.appendChild(movable[mi].el);
          }
          var _hT=pickHtpl(pi+s+2);
          if(_hT){
            var nh=document.createElement('div');
            nh.innerHTML=_hT.replace('<!--PAGE_NUM-->',(pi+s+2).toString());
            if(nh.firstChild)np.appendChild(nh.firstChild);
          }
          np.appendChild(nb);
          var nf=document.createElement('div');
          nf.innerHTML=pickFtpl(pi+s+2).replace('<!--PAGE_NUM-->',(pi+s+2).toString());
          if(nf.firstChild)np.appendChild(nf.firstChild);
          nw.appendChild(np);
          prevWrapper.after(nw);
          prevWrapper=nw;
        }
      }
      else{
        // Async path: single split per iter (setTimeout-driven, layout cost
        // amortized across event loop turns).
        var toMove=[];
        for(var mi=splitIdx;mi<children.length;mi++){
          if(children[mi].classList.contains('footnotes'))continue;
          if(isOutOfFlow(children[mi]))continue;
          toMove.push(children[mi]);
        }
        if(toMove.length===0)continue;
        var nw=document.createElement('div');
        nw.className='page-wrapper';
        var np=document.createElement('div');
        np.className='page';
        np.style.cssText=page.style.cssText;
        var nb=document.createElement('div');
        nb.className='page-body page-body-cont';
        for(var mi=0;mi<toMove.length;mi++){
          nb.appendChild(toMove[mi]);
        }
        var _hT=pickHtpl(pi+2);
        if(_hT){
          var nh=document.createElement('div');
          nh.innerHTML=_hT.replace('<!--PAGE_NUM-->',(pi+2).toString());
          if(nh.firstChild)np.appendChild(nh.firstChild);
        }
        np.appendChild(nb);
        var nf=document.createElement('div');
        nf.innerHTML=pickFtpl(pi+2).replace('<!--PAGE_NUM-->',(pi+2).toString());
        if(nf.firstChild)np.appendChild(nf.firstChild);
        nw.appendChild(np);
        var parentWrapper=page.closest('.page-wrapper');
        if(parentWrapper)parentWrapper.after(nw);
        else page.after(nw);
      }
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
    var rcAll=document.querySelectorAll('.page');
    var rcLim=rcAll.length;
    if(window._wpSync&&window._requestedPages&&window._requestedPages.length){
      rcLim=Math.min(rcAll.length,Math.max.apply(null,window._requestedPages)+1);
    }
    for(var rci=0;rci<rcLim;rci++){
      var p=rcAll[rci];
      var b=p.querySelector('.page-body');
      if(!b)continue;
      var f=b.querySelector('.footnotes');
      var fh=f?f.offsetHeight:0;
      var ch=0;
      var visibleCount=0;
      Array.from(b.children).forEach(function(c){
        if(c.classList.contains('footnotes'))return;
        if(isOutOfFlow(c))return;
        var bt=c.offsetTop+c.offsetHeight-b.offsetTop;
        if(bt>ch)ch=bt;
        if(c.offsetHeight>0)visibleCount++;
      });
      if(ch>maxBodyH-fh+2 && visibleCount>1){again=true;break;}
    }
    // Pages created in this (terminal, non-recursing) pass still need their
    // cloned header reflowed — adjustHeaderPadding is idempotent.
    if(!again)document.querySelectorAll('.page').forEach(adjustHeaderPadding);
    if(again){if(window._wpSync)paginate();else setTimeout(paginate,0);}
    else if(window._wpSync){
      positionFootnotes();wrapFloats();applyLineNumbers();
      // Build heading→page map BEFORE applyPageFilter (it sets display:none).
      var pgs=Array.from(document.querySelectorAll('.page'));
      var pmap=[];
      document.querySelectorAll('a[id^=""_Toc""]').forEach(function(a){
        for(var i=0;i<pgs.length;i++) if(pgs[i].contains(a)){pmap.push(a.id+'='+(i+1));break;}
      });
      if(window._gridCols>0){layoutGrid(window._gridCols);}
      else{applyPageFilter();flushScreenshotPage();}
      document.title='PAGES:'+pgs.length+(pmap.length?'|MAP:'+pmap.join(','):'');
    }
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
          +'font-size:inherit;color:#000;vertical-align:baseline;'
          +'user-select:none;pointer-events:none;';
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
  // Single-page screenshot: clip the one visible page to its page box and drop
  // the chrome (gray body padding, drop-shadow, rounded corners, wrapper margin)
  // so the capture is flush, for ANY page (not just page 1) — matching pptx.
  // The screenshot viewport is sized to the page's native pixels.
  function flushScreenshotPage(){
    if(!SCREENSHOT)return;
    // offsetParent is null when the page OR its wrapper is display:none, so this
    // counts only genuinely-visible pages (page-1 path hides wrappers; page-N
    // path hides .page elements via applyPageFilter).
    var vis=Array.prototype.filter.call(document.querySelectorAll('.page'),function(p){return p.offsetParent!==null;});
    if(vis.length!==1)return;
    var page=vis[0];
    page.style.height=page.style.minHeight;page.style.overflow='hidden';
    page.style.boxShadow='none';page.style.borderRadius='0';
    document.body.style.padding='0';
    var w=page.closest('.page-wrapper');if(w)w.style.margin='0';
    document.querySelectorAll('.page-wrapper').forEach(function(pw){if(pw!==w)pw.style.display='none';});
  }
  // Contact-sheet grid: scale every page down to a fixed cell width and let the
  // body flex-wrap tile them into _gridCols columns. cellW comes from the CLI
  // (window._gridCellW) so the captured viewport height (computed C#-side from
  // page count) matches the laid-out grid exactly. Mirrors flushScreenshotPage's
  // chrome-stripping but for ALL pages instead of clipping to one.
  function layoutGrid(cols){
    if(!cols||cols<1)return;
    var gap=12,pad=12;
    // Derive cellW from the live content width (excludes any scrollbar) so the
    // requested column count always fits — trusting a CLI-passed width risks the
    // scrollbar/rounding pushing the last column to a new row. Floor for safety.
    var avail=document.body.clientWidth-2*pad;
    var cellW=Math.floor((avail-(cols-1)*gap)/cols);
    if(cellW<1)cellW=window._gridCellW||220;
    var wraps=Array.prototype.filter.call(document.querySelectorAll('.page-wrapper'),function(w){return w.offsetParent!==null;});
    wraps.forEach(function(wrapper){
      var page=wrapper.querySelector('.page');if(!page)return;
      var pageW=page.offsetWidth,pageH=page.offsetHeight;
      var s=pageW>0?cellW/pageW:1;
      page.style.transition='none';page.style.transformOrigin='top left';
      page.style.transform='scale('+s+')';
      page.style.boxShadow='none';page.style.borderRadius='0';
      wrapper.style.transition='none';wrapper.style.margin='0';
      wrapper.style.width=cellW+'px';wrapper.style.height=Math.round(pageH*s)+'px';
      wrapper.style.overflow='hidden';
    });
    var b=document.body;
    b.style.display='flex';b.style.flexWrap='wrap';
    b.style.alignContent='flex-start';b.style.justifyContent='center';b.style.alignItems='flex-start';
    b.style.gap=gap+'px';b.style.padding=pad+'px';
  }
  function _loadKatexLazy(cb){
    // Watch mode: doc may start formula-free (KaTeX tags omitted), then
    // gain a formula via SSE patch. Inject CSS + JS on demand; on load,
    // re-invoke the caller so the new formula renders.
    if(window._katexLoading){window._katexCallbacks=window._katexCallbacks||[];window._katexCallbacks.push(cb);return;}
    window._katexLoading=true;window._katexCallbacks=[cb];
    // CONSISTENCY(katex-mirror): mirror-first, CDN retry, then plain-text fallback.
    var link=document.createElement('link');link.rel='stylesheet';link.href='{{KATEX_CSS}}';link.onerror=function(){if(!this.dataset.f){this.dataset.f=1;this.href='{{KATEX_CSS_CDN}}';}else{this.remove();}};document.head.appendChild(link);
    var _kOk=function(){(window._katexCallbacks||[]).forEach(function(f){try{f();}catch(e){}});window._katexCallbacks=[];};
    var _kFail=function(){document.querySelectorAll('.katex-formula:not(.katex-rendered)').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666';el.classList.add('katex-rendered');});};
    var s=document.createElement('script');s.src='{{KATEX_JS}}';s.onload=_kOk;
    s.onerror=function(){var s2=document.createElement('script');s2.src='{{KATEX_JS_CDN}}';s2.onload=_kOk;s2.onerror=_kFail;document.head.appendChild(s2);};
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
"
            // CONSISTENCY(katex-mirror): the verbatim block above can't interpolate;
            // substitute the KaTeX asset URLs (mirror + CDN fallback) afterwards.
            .Replace("{{KATEX_CSS}}", Core.KatexAssets.CssUrl)
            .Replace("{{KATEX_CSS_CDN}}", Core.KatexAssets.CdnCssUrl)
            .Replace("{{KATEX_JS}}", Core.KatexAssets.JsUrl)
            .Replace("{{KATEX_JS_CDN}}", Core.KatexAssets.CdnJsUrl));
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
        // Contact-sheet grid: tile every page into gridCols columns. cellW is
        // supplied by the CLI (not derived from clientWidth) so the C# viewport
        // height math and the in-browser layout use the identical cell size.
        if (gridCols > 0)
        {
            sb.AppendLine($"  window._gridCols={gridCols};");
            if (gridCellWpx > 0) sb.AppendLine($"  window._gridCellW={gridCellWpx};");
        }
        sb.AppendLine(@"  var SCREENSHOT=location.hash.indexOf('screenshot')>=0||navigator.webdriver||/HeadlessChrome/.test(navigator.userAgent);
  window._wpSync=SCREENSHOT;
  if(SCREENSHOT){
    var rp=window._requestedPages;
    if(window._gridCols>0){paginate();}
    else if(rp&&rp.length===1&&rp[0]===1){
      // Page 1 is the first .page before pagination — flush it directly, skipping
      // the full paginate pass. Other single pages go through paginate and flush
      // at its sync completion. flushScreenshotPage clips + drops the chrome.
      document.querySelectorAll('.page-wrapper:not(:first-of-type)').forEach(function(w){w.style.display='none';});
      var _p1=document.querySelector('.page');if(_p1)adjustHeaderPadding(_p1);
      positionFootnotes();wrapFloats();applyLineNumbers();
      flushScreenshotPage();
    }else{paginate();}
  }else{setTimeout(paginate,100);}");
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

    /// <summary>
    /// First-section page size in CSS pixels at 96 DPI (page pt × 96/72) — the
    /// size a page renders at in the HTML preview. The screenshot path sizes a
    /// single-page viewport to this so the PNG is the page, with no letterbox
    /// padding. Mirrors PowerPointHandler.GetSlideNativePixels. Falls back to
    /// US-Letter (816×1056).
    /// </summary>
    internal (int width, int height) GetPageNativePixels()
    {
        var pg = GetPageLayout();
        int w = (int)Math.Round(pg.WidthPt * 96.0 / 72.0);
        int h = (int)Math.Round(pg.HeightPt * 96.0 / 72.0);
        return w > 0 && h > 0 ? (w, h) : (816, 1056);
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

    /// <summary>
    /// Per-section column style for the page-body div. Sections with no
    /// &lt;w:cols&gt; (or 1 column) → empty (single column). Computed per page
    /// from the active section so later sections don't inherit section 1's
    /// column count.
    /// </summary>
    private static string BuildColBodyStyle(SectionProperties? section, double colBodyHeightPt)
    {
        // w:vAlign on the section vertically aligns the page content block.
        // The .page-body is already a flex column (flex:1), so justify-content
        // positions the whole content block along the cross page axis. Word's
        // "both"/"justify" stretch the block top-to-bottom; we render those as
        // top-aligned (no vertical justification of paragraph gaps) which is
        // visually closest without per-line distribution.
        var vAlign = section?.GetFirstChild<VerticalTextAlignmentOnPage>()?.Val?.InnerText;
        var justify = vAlign switch
        {
            "center" => "center",
            "bottom" => "flex-end",
            _ => null, // top / both / justify / null → default flex-start
        };

        var sectCols = section?.GetFirstChild<Columns>();
        var colCount = sectCols?.ColumnCount?.Value ?? 1;
        if (colCount <= 1)
            return justify != null ? $" style=\"justify-content:{justify}\"" : "";
        var colSep = sectCols?.Separator?.Value == true;
        var colSpacing = sectCols?.Space?.Value;
        // BUG(cols-overprint, R85): use min-height (not fixed height) so the
        // multi-column box can GROW past the page body height when the section's
        // content exceeds two columns worth. With a fixed `height`, CSS
        // multi-column has no page break to flow into — once both columns fill
        // to that height, the overflow opens a THIRD column that wraps back to
        // the top of the SAME bounded box and visually overprints the existing
        // columns (right ~40% became unreadable). min-height keeps the same
        // balanced height for short content (columns still balance at the
        // minimum) but lets tall content extend the box downward so later
        // columns stack below instead of overlapping. HTML has no true
        // page break, so this multi-column block renders taller than Word's
        // real paginated layout — accepted (content-visible > exact height,
        // same no-clip principle as the rest of this campaign).
        return $" style=\"column-count:{colCount}"
            + $";min-height:{colBodyHeightPt.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture)}pt"
            + (colSep ? ";column-rule:1px solid #000" : "")
            + (int.TryParse(colSpacing, out var csp) && csp > 0 ? $";column-gap:{csp / 20.0:0.##}pt" : "")
            + "\"";
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
    /// Build the per-side CSS for a section's page border (<w:pgBorders>).
    /// Real Word draws a box around the whole page; we map each present
    /// side element (top/left/bottom/right) to a border-&lt;side&gt; rule.
    ///   sz   — 1/8 pt → pt
    ///   color — ST_HexColor (auto → black) → #hex
    ///   val  — single/thick/etc. → solid; the *Thick*Gap / double-ish
    ///          variants → double (best-effort, OOXML art borders collapse).
    /// offsetFrom is approximated: page-relative borders (the common case)
    /// hug the page edge, so we leave the border on the .page box itself
    /// (which sits at the page edge). text-relative would inset to the
    /// margin, but the box-on-page rendering is the close-enough default.
    /// Returns "" when no pgBorders / no sides present.
    /// </summary>
    private static string BuildPageBorderCss(SectionProperties? sectPr)
    {
        var pb = sectPr?.GetFirstChild<PageBorders>();
        if (pb == null) return "";
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        var sb = new System.Text.StringBuilder();
        void Emit(BorderType? b, string side)
        {
            if (b == null) return;
            // val: "none"/"nil" means no border on that side.
            var val = b.Val?.InnerText ?? "single";
            if (val.Equals("none", StringComparison.OrdinalIgnoreCase)
                || val.Equals("nil", StringComparison.OrdinalIgnoreCase))
                return;
            // sz is in eighths of a point; default ~4 (0.5pt) when absent.
            double szEighths = 4;
            if (b.Size?.Value is { } szv) szEighths = szv;
            else if (b.Size?.InnerText is { Length: > 0 } szt
                     && double.TryParse(szt, System.Globalization.NumberStyles.Any, ci, out var szp))
                szEighths = szp;
            var widthPt = szEighths / 8.0;
            if (widthPt <= 0) widthPt = 0.5;
            // color: ST_HexColor; "auto" or empty → black.
            var rawColor = b.Color?.InnerText;
            string color = string.IsNullOrEmpty(rawColor)
                || rawColor.Equals("auto", StringComparison.OrdinalIgnoreCase)
                ? "#000000"
                : ParseHelpers.FormatHexColor(rawColor);
            // style: the "...Gap" double-line art variants → double; the
            // explicit double → double; everything else → solid.
            var style = val.IndexOf("double", StringComparison.OrdinalIgnoreCase) >= 0
                || val.IndexOf("Gap", StringComparison.OrdinalIgnoreCase) >= 0
                ? "double"
                : "solid";
            // CSS double needs >=3px to show two lines; bump the doubled
            // art borders so both rules render instead of collapsing.
            if (style == "double" && widthPt < 3) widthPt = 3;
            sb.Append($"border-{side}:{widthPt.ToString("0.##", ci)}pt {style} {color};");
        }
        Emit(pb.TopBorder, "top");
        Emit(pb.LeftBorder, "left");
        Emit(pb.BottomBorder, "bottom");
        Emit(pb.RightBorder, "right");
        return sb.ToString();
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

        // Font: Normal style rFonts → docDefaults rFonts → theme minor font → fallback.
        // OOXML cascade: the default paragraph (Normal) style overrides docDefaults,
        // so Normal's rFonts must win when present (matches ResolveEffectiveRunPropertiesCore).
        var fonts = rPr?.RunFonts;
        var nFonts = defaultRPr?.RunFonts;
        var font = NonEmpty(nFonts?.EastAsia?.Value) ?? NonEmpty(nFonts?.Ascii?.Value) ?? NonEmpty(nFonts?.HighAnsi?.Value);
        if (font == null)
            font = NonEmpty(fonts?.EastAsia?.Value) ?? NonEmpty(fonts?.Ascii?.Value) ?? NonEmpty(fonts?.HighAnsi?.Value);
        if (font == null)
        {
            try
            {
                var minor = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont;
                font = NonEmpty(minor?.EastAsianFont?.Typeface) ?? NonEmpty(minor?.LatinFont?.Typeface);
            }
            catch (System.Xml.XmlException) { }
        }

        // Size: Normal style → docDefaults → fallback (half-points → pt).
        // Same cascade rationale as font above — Normal's sz overrides docDefaults.
        double sizePt = 0;
        if (defaultRPr?.FontSize?.Val?.Value is string nsz && int.TryParse(nsz, out var nhp))
            sizePt = nhp / 2.0;
        if (sizePt == 0 && rPr?.FontSize?.Val?.Value is string sz && int.TryParse(sz, out var hp))
            sizePt = hp / 2.0;
        // OOXML §17.7.4.5 default: 20 half-points = 10pt when neither
        // rPrDefault nor Normal carries a size.
        if (sizePt == 0) sizePt = 10.0;

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

        // Default text color: explicit docDefaults w:color only; otherwise black.
        // Word does NOT derive the default run color from theme dk1 — an
        // uncolored run is "auto" (rendered black on a light backdrop). A
        // brand-colored dk1 (e.g. srgbClr B82326) must not bleed onto every
        // run that lacks an explicit w:color / themeColor. Runs that DO carry
        // w:themeColor="text1"/"dark1" still resolve to dk1 via ResolveRunColor;
        // auto-on-dark-bg reverse-out (R39) is handled separately in
        // ResolveEffectiveBackgroundForRun.
        var color = "#000000";
        var cv = rPr?.Color?.Val?.Value;
        if (cv != null && cv != "auto" && IsHexColor(cv)) color = $"#{cv}";

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

        // Theme fonts, read once — both added to the set and used to resolve
        // theme-token literals below. Malformed theme1.xml shouldn't taint
        // the font set.
        string? majLatin = null, minLatin = null, majEa = null, minEa = null, majCs = null, minCs = null;
        try
        {
            var theme = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
            majLatin = theme?.MajorFont?.LatinFont?.Typeface?.Value;
            minLatin = theme?.MinorFont?.LatinFont?.Typeface?.Value;
            majEa = theme?.MajorFont?.EastAsianFont?.Typeface?.Value;
            minEa = theme?.MinorFont?.EastAsianFont?.Typeface?.Value;
            majCs = theme?.MajorFont?.ComplexScriptFont?.Typeface?.Value;
            minCs = theme?.MinorFont?.ComplexScriptFont?.Typeface?.Value;
        }
        catch (System.Xml.XmlException) { }

        // Wild documents park theme-token ENUM literals in the plain
        // ascii/eastAsia attributes (where only asciiTheme/eastAsiaTheme
        // should carry them). Taken verbatim they leak to @font-face and the
        // web-font requests as fake family names ("minorHAnsi"), and the
        // local metric-override lookup misses. Resolve to the theme font;
        // an unresolvable token is dropped, never emitted verbatim.
        void Add(string? raw)
        {
            var mapped = raw switch
            {
                "majorAscii" or "majorHAnsi" => majLatin,
                "minorAscii" or "minorHAnsi" => minLatin,
                "majorEastAsia" => majEa,
                "minorEastAsia" => minEa,
                "majorBidi" => majCs,
                "minorBidi" => minCs,
                _ => raw,
            };
            var name = NormalizeHarvestedFontName(mapped);
            if (name != null) fonts.Add(name);
        }

        // From styles
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles != null)
            foreach (var rf in styles.Descendants<RunFonts>())
            {
                Add(rf.Ascii?.Value);
                Add(rf.HighAnsi?.Value);
                Add(rf.EastAsia?.Value);
            }
        // From document body
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body != null)
            foreach (var rf in body.Descendants<RunFonts>())
            {
                Add(rf.Ascii?.Value);
                Add(rf.HighAnsi?.Value);
            }
        Add(majLatin);
        Add(minLatin);
        // Remove fonts that have no usable @font-face (symbols, wingdings)
        fonts.RemoveWhere(f => f.StartsWith("Symbol") || f.StartsWith("Wingding"));
        return fonts;
    }

    /// <summary>
    /// Junk gate for font names harvested from wild documents. rFonts values
    /// in real-world files carry OOXML control-char escapes (_x000B_), CSS
    /// keywords ("sans-serif", "inherit"), CSS fragments ("var(--x)",
    /// "font-weight 400", "auto 16px 1.4"), and mojibake — all observed being
    /// sent upstream as font families. Returns the cleaned name, or null when
    /// the value cannot be a font name.
    /// </summary>
    internal static string? NormalizeHarvestedFontName(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        // Bad producers park escaped control chars in rFonts (_x000B_ = VT).
        var s = System.Text.RegularExpressions.Regex.Replace(raw, "_x[0-9A-Fa-f]{4}_", "").Trim();
        // Real family names are short; longer strings are style blobs.
        if (s.Length == 0 || s.Length > 64) return null;
        // CSS-fragment characters never appear in genuine family names.
        if (s.IndexOfAny(CssFragmentChars) >= 0) return null;
        if (s.Contains("--")) return null; // var(--custom-prop) residue
        if (!s.Any(char.IsLetter)) return null;
        // CSS-wide keywords / generic families / non-names.
        if (CssKeywordNonFonts.Contains(s)) return null;
        // Style-shorthand residue: "font-weight 400", "auto 16px 1.4".
        if (s.StartsWith("font-", StringComparison.OrdinalIgnoreCase)) return null;
        if (System.Text.RegularExpressions.Regex.IsMatch(s, @"\d+(\.\d+)?(px|pt|em|rem)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return null;
        return s;
    }

    private static readonly char[] CssFragmentChars = { '(', ')', ':', ';', ',', '{', '}', '"', '\'', '<', '>', '/', '\\' };

    private static readonly HashSet<string> CssKeywordNonFonts = new(StringComparer.OrdinalIgnoreCase)
    {
        // NOTE: "fangsong" is deliberately ABSENT although CSS4 lists it as a
        // generic family — it is also the real ASCII name of the FangSong
        // (仿宋) font, and killing it would drop the font's metric-override
        // and web-font fallback. A generic-keyword leak of that spelling is
        // far less likely than the genuine font name.
        "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui",
        "ui-serif", "ui-sans-serif", "ui-monospace", "ui-rounded", "math", "emoji",
        "inherit", "initial", "unset", "revert", "revert-layer",
        "auto", "normal", "none", "undefined", "null", "inline",
    };

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
            // "Calibri Light"-style names are separate families in Word but
            // rarely separate installed fonts; fall back to the base family
            // locally so a machine with plain Calibri still renders close.
            var baseFont = safeFont.EndsWith(" Light", StringComparison.OrdinalIgnoreCase) && safeFont.Length > 6
                ? safeFont[..^6].TrimEnd()
                : null;
            string Src(string suffix) => baseFont != null
                ? $"local('{safeFont}{suffix}'), local('{baseFont}{suffix}')"
                : $"local('{safeFont}{suffix}')";
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; src: {Src("")};{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; font-weight: bold; src: {Src(" Bold")};{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; font-style: italic; src: {Src(" Italic")};{overrides} }}");
            sb.AppendLine($"@font-face {{ font-family: '{safeFont}'; font-weight: bold; font-style: italic; src: {Src(" Bold Italic")};{overrides} }}");
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
        // val="solid": pattern is 100% of the FOREGROUND (w:color), so the
        // background is the color (black when color is absent/auto), NOT the fill.
        // val="clear" (and pct/named patterns approximated elsewhere): background is the fill.
        if (shading.Val != null && shading.Val.Value == ShadingPatternValues.Solid)
        {
            var color = shading.Color?.Value;
            if (color != null && color != "auto" && IsHexColor(color)) return $"#{color}";
            return "#000000";
        }
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
        var pgW = (long)(pg.WidthCm / 2.54 * EmuConverter.EmuPerInch);
        var pgH = (long)(pg.HeightCm / 2.54 * EmuConverter.EmuPerInch);
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
                var savedHost = _ctx.ImageHostPart;
                _ctx.ImageHostPart = hp;
                // Watermark spans are collected separately so they can be
                // emitted OUTSIDE the .doc-header div — the watermark must
                // be centered on the whole page (.page is its positioning
                // ancestor), not pinned inside the narrow header band.
                var headerBodySb = new StringBuilder();
                var watermarkSb = new StringBuilder();
                RenderHeaderFooterBody(headerBodySb, hp.Header, watermarkSb);
                _ctx.ImageHostPart = savedHost;
                sb.Append(watermarkSb);
                sb.AppendLine($"<div class=\"{cssClass}\">");
                sb.Append(headerBodySb);
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
                var savedHost = _ctx.ImageHostPart;
                _ctx.ImageHostPart = fp;
                var footerBodySb = new StringBuilder();
                var watermarkSb = new StringBuilder();
                RenderHeaderFooterBody(footerBodySb, fp.Footer, watermarkSb);
                _ctx.ImageHostPart = savedHost;
                sb.Append(watermarkSb);
                sb.AppendLine($"<div class=\"{cssClass}\">");
                sb.Append(footerBodySb);
                sb.AppendLine("</div>");
                break;
            }
        }
    }

    /// <summary>
    /// Detect whether a header/footer part actually contains a PAGE and/or
    /// NUMPAGES field. Only when the source part carries the real field may the
    /// digit-only run it rendered be rewritten into a dynamic .page-num-field /
    /// .num-pages-field span. Without this gate the digit-matching regex would
    /// clobber literal numbers (a year "2014", a phone number, a price) that
    /// happen to be the first digit run in the header/footer.
    ///
    /// Both <c>&lt;w:fldSimple w:instr="...PAGE..."&gt;</c> and the
    /// <c>&lt;w:instrText&gt;PAGE&lt;/w:instrText&gt;</c> (fldChar) forms are
    /// matched. NUMPAGES is checked first so a "NUMPAGES" instruction is not
    /// miscounted as a plain "PAGE".
    /// </summary>
    private static (bool hasPage, bool hasNumPages) HeaderFooterFieldFlags(OpenXmlElement? hf)
    {
        bool hasPage = false, hasNumPages = false;
        if (hf == null) return (false, false);
        // fldSimple: instruction lives in the w:instr attribute.
        foreach (var fld in hf.Descendants<SimpleField>())
            Classify(fld.Instruction?.Value);
        // fldChar field: instruction lives in <w:instrText> runs.
        foreach (var instr in hf.Descendants<FieldCode>())
            Classify(instr.Text);
        return (hasPage, hasNumPages);

        void Classify(string? instr)
        {
            if (string.IsNullOrEmpty(instr)) return;
            // Word field instructions are case-insensitive; tokens are
            // whitespace-delimited (e.g. "PAGE \* MERGEFORMAT").
            var tokens = instr.Split(new[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);
            foreach (var tok in tokens)
            {
                if (tok.Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase)) hasNumPages = true;
                else if (tok.Equals("PAGE", StringComparison.OrdinalIgnoreCase)) hasPage = true;
            }
        }
    }

    /// <summary>Classify a Word field instruction string: 1 = PAGE,
    /// 2 = NUMPAGES, 0 = neither. NUMPAGES is tested first so a NUMPAGES
    /// instruction is never miscounted as a plain PAGE. Tokens are
    /// whitespace-delimited and case-insensitive
    /// (e.g. "PAGE \* MERGEFORMAT").</summary>
    private static int ClassifyPageFieldInstruction(string? instr)
    {
        if (string.IsNullOrEmpty(instr)) return 0;
        var tokens = instr.Split(new[] { ' ', '\t', '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries);
        foreach (var tok in tokens)
        {
            if (tok.Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase)) return 2;
            if (tok.Equals("PAGE", StringComparison.OrdinalIgnoreCase)) return 1;
        }
        return 0;
    }

    /// <summary>
    /// Rewrite the rendered PAGE / NUMPAGES field-result run into a dynamic
    /// page-number / total-pages span. The result run was tagged at render time
    /// with a <c>page-num-result</c> / <c>num-pages-result</c> class
    /// (<see cref="RenderParagraphContentHtml"/>), so the rewrite targets the
    /// real field scope — never a literal header/footer number (a date "01", a
    /// course code "001", a year "2021") that merely happens to be the first
    /// digit run in document order.
    ///
    /// <paramref name="hasPage"/> / <paramref name="hasNumPages"/> still gate the
    /// rewrite: only field kinds the source part actually carries are rewritten.
    /// A legacy fallback regex handles parts whose result run carries no tag
    /// (e.g. an empty/uncached field result, or a field whose result digits were
    /// rendered outside a tagged run) — but it is now scoped to a tagged result
    /// span first, and only falls through to the first-digit-run heuristic when
    /// no tagged result span exists, so the date-corruption case never reaches it.
    /// </summary>
    private static string ApplyPageNumFields(string html, bool hasPage, bool hasNumPages)
    {
        if (string.IsNullOrEmpty(html) || (!hasPage && !hasNumPages)) return html;

        // Primary path: the result run was tagged with its field-kind class.
        // Replace the whole tagged span's inner content with the dynamic
        // placeholder span (keeping the renumber-loop class hooks intact).
        //
        // The body alternation `(?:[^<]|<span[^>]*>[^<]*</span>)*` is required
        // because RenderRunHtml wraps a formatted result run (e.g. the cached
        // PAGE digit carrying the Footer style's color) in its OWN inner
        // `<span style=...>…</span>`. A naive `.*?</span>` (non-greedy) would
        // stop at that INNER `</span>`, replace only up to it, and leak the
        // cached digit out as a stray span — rendering "Page 2" as "Page 22"
        // (live number + leftover cache). Matching the wrapper's balanced
        // close (allowing zero-or-more nested formatting spans) consumes the
        // whole result run so the cached digit never survives.
        var pageTag = new Regex(
            @"<span class=""page-num-result"">(?:[^<]|<span[^>]*>[^<]*</span>)*</span>",
            RegexOptions.Singleline);
        var numPagesTag = new Regex(
            @"<span class=""num-pages-result"">(?:[^<]|<span[^>]*>[^<]*</span>)*</span>",
            RegexOptions.Singleline);
        bool didPage = false, didNumPages = false;
        if (hasPage && pageTag.IsMatch(html))
        {
            html = pageTag.Replace(html,
                "<span class=\"page-num-field\"><!--PAGE_NUM--></span>", 1);
            didPage = true;
        }
        if (hasNumPages && numPagesTag.IsMatch(html))
        {
            html = numPagesTag.Replace(html,
                "<span class=\"num-pages-field\"><!--NUM_PAGES--></span>", 1);
            didNumPages = true;
        }
        // Strip any leftover (unmatched) result tags so they don't leak the
        // raw class into the preview; the inner digits stay visible.
        html = pageTag.Replace(html, m =>
            m.Value.Substring("<span class=\"page-num-result\">".Length,
                m.Value.Length - "<span class=\"page-num-result\">".Length - "</span>".Length));
        html = numPagesTag.Replace(html, m =>
            m.Value.Substring("<span class=\"num-pages-result\">".Length,
                m.Value.Length - "<span class=\"num-pages-result\">".Length - "</span>".Length));

        // Legacy fallback: only when the field carried no tagged result run
        // (e.g. fldSimple with a cached <w:t>, where the result isn't a
        // begin/separate/end run sequence). Scoped exactly as before.
        var pat = new Regex(@"(<(?:span|p)[^>]*>)\s*\d+\s*(</(?:span|p)>)");
        if (hasPage && !didPage)
            html = pat.Replace(html,
                "$1<span class=\"page-num-field\"><!--PAGE_NUM--></span>$2", 1);
        if (hasNumPages && !didNumPages)
            html = pat.Replace(html,
                "$1<span class=\"num-pages-field\"><!--NUM_PAGES--></span>$2", 1);
        return html;
    }

    /// <summary>PAGE/NUMPAGES flags for the FIRST content-bearing header/footer
    /// part — the same one <see cref="RenderHeaderFooterHtml"/> renders into the
    /// fallback HTML. Keeps the fallback digit-rewrite gated to real fields.</summary>
    private (bool page, bool numPages) FirstContentHeaderFooterFlags(bool isHeader)
    {
        if (isHeader)
        {
            var headerParts = _doc.MainDocumentPart?.HeaderParts;
            if (headerParts != null)
                foreach (var hp in headerParts)
                    if (hp.Header != null && HeaderFooterHasContent(hp.Header))
                        return HeaderFooterFieldFlags(hp.Header);
        }
        else
        {
            var footerParts = _doc.MainDocumentPart?.FooterParts;
            if (footerParts != null)
                foreach (var fp in footerParts)
                    if (fp.Footer != null && HeaderFooterHasContent(fp.Footer))
                        return HeaderFooterFieldFlags(fp.Footer);
        }
        return (false, false);
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
            // Watermarks live inside <w:sdt><w:sdtContent><w:p><w:pict>…
            // descend through SDT block wrappers so they aren't treated as
            // empty headers.
            if (child is SdtBlock sdt && sdt.SdtContentBlock is { } content
                && HeaderFooterHasContent(content)) return true;
        }
        return false;
    }

    /// <summary>Iterate header/footer children in order, rendering paragraphs
    /// and tables. Previously only paragraphs were emitted, dropping layout
    /// tables and image-only paragraphs.</summary>
    private void RenderHeaderFooterBody(StringBuilder sb, OpenXmlElement hf, StringBuilder? watermarkSb = null)
    {
        // List items in a header/footer previously rendered as bare paragraphs
        // (RenderParagraphHtml carries no list machinery) — every bullet/number
        // marker was silently dropped while the body and table-cell paths
        // rendered them. Reuse the cell path's per-container list state: same
        // container-matrix defect class as the text-box display-math gap.
        string? hfListTag = null;
        var hfOlState = new OrderedListNumberingState();
        void CloseHfList()
        {
            if (hfListTag != null) { sb.Append($"</{hfListTag}>"); hfListTag = null; }
        }
        foreach (var child in hf.ChildElements)
        {
            if (TryEmitContainerBookmarkAnchor(sb, child)) continue;
            // Watermark paragraphs are nested inside <w:sdt><w:sdtContent>;
            // recurse so they render the same as direct paragraph children.
            if (child is SdtBlock sdt && sdt.SdtContentBlock is { } content)
            {
                RenderHeaderFooterBody(sb, content, watermarkSb);
                continue;
            }
            if (child is Paragraph para)
            {
                // Legacy VML watermark: a <v:shape> in a <w:pict> with
                // a <v:textpath> child carrying the watermark string
                // (DRAFT / CONFIDENTIAL / …). DrawingML text boxes are
                // already handled by the shape renderer; VML is a
                // parallel deprecated format we must detect by name.
                var watermark = ExtractVmlWatermark(para);
                if (watermark is (var watermarkText, var watermarkColor))
                {
                    var colorCss = string.IsNullOrWhiteSpace(watermarkColor)
                        ? "#d0d0d0"
                        : NormalizeVmlColorCss(watermarkColor!);
                    // The watermark must be centered over the whole page, so it
                    // is wrapped in a full-page layer (position:absolute;inset:0)
                    // that the caller emits as a direct child of .page — NOT
                    // inside .doc-header (whose narrow band would pin the
                    // top:50%/left:50% anchor to the header strip, leaving the
                    // watermark in the upper-left quadrant). The inner span's
                    // 50%/50% then resolves against this full-page layer.
                    var watermarkHtml =
                        "<div class=\"vml-watermark-layer\" style=\"position:absolute;inset:0;" +
                        "z-index:0;pointer-events:none\">" +
                        $"<span class=\"vml-watermark\" style=\"position:absolute;" +
                        "top:50%;left:50%;transform:translate(-50%,-50%) rotate(-45deg);" +
                        $"color:{colorCss};font-size:7em;font-weight:bold;" +
                        "pointer-events:none;white-space:nowrap;user-select:none\">" +
                        HtmlEncode(watermarkText) +
                        "</span></div>";
                    // Prefer the dedicated watermark buffer (emitted outside
                    // .doc-header); fall back to inline if a caller didn't
                    // provide one (keeps older call paths working).
                    (watermarkSb ?? sb).Append(watermarkHtml);
                    // The same paragraph may carry other in-flow content
                    // alongside the VML watermark — most notably a DrawingML
                    // logo (<a:blip r:embed>) in a first-page header that packs
                    // the brand mark and the "EXAMPLE"/"DRAFT" watermark into a
                    // single <w:p>. Render that residual content (it walks runs
                    // via Descendants<Drawing>(), which skips the VML
                    // <w:pict>/<v:textpath>, so the watermark is not
                    // double-rendered). Skip the paragraph entirely when the
                    // watermark is its only content, to avoid emitting a blank
                    // <p> in the header band (prior behaviour: bare `continue`).
                    if (ParagraphHasNonWatermarkContent(para))
                    {
                        CloseHfList();
                        RenderParagraphHtml(sb, para);
                    }
                    continue;
                }
                var hfListStyle = GetParagraphListStyle(para);
                if (hfListStyle != null)
                {
                    RenderCellListItem(sb, para, hfListStyle, ref hfListTag, hfOlState);
                    continue;
                }
                CloseHfList();
                RenderParagraphHtml(sb, para);
            }
            else if (child is Table tbl)
            {
                CloseHfList();
                RenderTableHtmlPaged(sb, tbl);
            }
        }
        CloseHfList();
    }

    /// <summary>
    /// Render a top-level (body) table, then neutralize any page-break markers
    /// that originated inside its cells.
    ///
    /// R130: a <c>&lt;w:br w:type="page"/&gt;</c> inside a table cell emits the
    /// <c>&lt;!--PAGE_BREAK--&gt;</c> marker via <see cref="RenderParagraphContentHtml"/>
    /// (it can't know it's inside a cell). When the body buffer is later split on
    /// that marker (<see cref="ViewAsHtmlCore"/>), the page-wrapper/page/page-body
    /// boundary divs would be injected while the <c>&lt;td&gt;</c>/<c>&lt;table&gt;</c>
    /// is still open — producing invalid nested HTML AND two bogus blank A4 pages.
    /// Real Word renders an in-cell page break as a soft line/paragraph break and
    /// keeps the whole table on one page (a table can't be split mid-cell across a
    /// hard page boundary the way the body flow can).
    ///
    /// Every PAGE_BREAK marker produced during a table render is necessarily
    /// in-cell (tables themselves never emit the marker — only the per-paragraph
    /// content renderer does, and inside a table that only runs for cell
    /// paragraphs). The in-cell emit shape is <c>&lt;/p&gt;&lt;!--PAGE_BREAK--&gt;&lt;p…&gt;</c>,
    /// so dropping just the marker leaves a valid <c>&lt;/p&gt;&lt;p…&gt;</c>
    /// paragraph boundary inside the cell. Body-level page breaks are unaffected:
    /// they are emitted directly into the body buffer outside any table wrapper.
    /// </summary>
    private void RenderTableHtmlPaged(StringBuilder sb, Table table, string? dataPath = null, OrderedListNumberingState? olState = null)
    {
        int start = sb.Length;
        RenderTableHtml(sb, table, dataPath: dataPath, olState: olState);
        // Strip in-cell page-break markers from just-appended table fragment.
        var fragment = sb.ToString(start, sb.Length - start);
        if (fragment.Contains("<!--PAGE_BREAK-->"))
        {
            sb.Length = start;
            sb.Append(fragment.Replace("<!--PAGE_BREAK-->", ""));
        }
    }

    /// <summary>
    /// Return the watermark text from a legacy VML <c>w:pict &gt; v:shape &gt;
    /// v:textpath</c> structure, or null if the paragraph does not carry one.
    /// </summary>
    private static (string text, string? color)? ExtractVmlWatermark(Paragraph para)
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
            if (string.IsNullOrWhiteSpace(str)) continue;
            // fillcolor on v:shape lives in the default (no-prefix) namespace.
            var fillcolor = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "fillcolor").Value;
            return (str, string.IsNullOrWhiteSpace(fillcolor) ? null : fillcolor);
        }
        return null;
    }

    /// <summary>
    /// Turn a VML color attribute value into a valid CSS color. VML
    /// <c>fillcolor</c> accepts both hex (<c>#C0C0C0</c>, <c>C0C0C0</c>,
    /// <c>#ccc</c>) and CSS named colors (<c>silver</c>, <c>red</c>, …).
    /// Only bare 3/6-digit hex gets a <c>#</c> prepended; named colors pass
    /// through unchanged. Blindly prepending <c>#</c> turned <c>silver</c>
    /// into the invalid <c>#silver</c>, which browsers drop → solid black.
    /// </summary>
    private static string NormalizeVmlColorCss(string color)
    {
        var c = color.Trim();
        if (c.StartsWith("#")) return c;
        bool isHex = (c.Length == 3 || c.Length == 6)
            && c.All(ch => (ch >= '0' && ch <= '9')
                || (ch >= 'a' && ch <= 'f')
                || (ch >= 'A' && ch <= 'F'));
        return isHex ? "#" + c : c;
    }

    /// <summary>
    /// True when a paragraph that carries a VML watermark (<c>v:textpath</c>)
    /// ALSO carries renderable in-flow content — a DrawingML drawing
    /// (<c>&lt;w:drawing&gt;</c>, e.g. a header logo) or any visible text.
    /// Used to decide whether to fall through to <c>RenderParagraphHtml</c>
    /// after emitting the watermark span: a watermark-only paragraph is
    /// skipped to avoid a blank <c>&lt;p&gt;</c>, while a paragraph mixing the
    /// watermark with a logo still renders the logo.
    /// </summary>
    private static bool ParagraphHasNonWatermarkContent(Paragraph para)
    {
        // DrawingML (real picture/shape) — SDK Descendants<Drawing>() does not
        // see the VML <w:pict>, so any hit here is non-watermark content.
        if (para.Descendants<Drawing>().Any()) return true;
        // Any visible text run.
        if (para.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text))) return true;
        return false;
    }

    /// <summary>
    /// True when a paragraph carries renderable content (visible text or a
    /// DrawingML drawing). Used by the R116 pre-break multi-column scan to tell
    /// a content-bearing section (wrap its preceding content in columns) apart
    /// from an empty leading section-delimiter paragraph (handled after-break).
    /// </summary>
    private static bool ParagraphHasRenderableContent(Paragraph para)
    {
        if (para.Descendants<Drawing>().Any()) return true;
        return para.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text));
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
        int prevOoxmlIlvl = 0; // previous list item's RAW (pre-offset) ilvl — nesting depth follows ilvl, not numId indent
        var numIdLevelOffset = new Dictionary<int, int>(); // numId → effective ilvl offset for cross-numId nesting
        // Ordered-list counters shared with the plain-text walker via
        // SeedOrderedStart / AdvanceOrderedCounter / RenderOrderedMarker
        // (CONSISTENCY(list-marker)). olState.OlCountPerLevel holds the running
        // <ol> item count per ilvl; .MultiLevelCounters feeds the lvlText
        // template; .AbsNumLevelCounters persists across numId changes so two
        // num instances on the same abstractNum continue numbering (Word's
        // "continue" behavior) unless a <w:lvlOverride><w:startOverride/>
        // resets it. The HTML path keeps currentNumId locally for nesting.
        var olState = new OrderedListNumberingState();
        var headingCounters = new Dictionary<int, int>(); // ilvl → counter for heading auto-numbering from style numPr
        bool pendingLiClose = false; // defer </li> to allow nested lists inside
        bool inMultiColumn = false; // track whether we're inside a multi-column div

        // Pre-scan: build a map of section column counts from inline sectPr breaks
        // The last section's cols come from the body sectPr
        var bodySectPr = body.GetFirstChild<SectionProperties>();
        var bodyColCount = GetSectionColumnCount(bodySectPr);

        // BUG(content-before-multicol-break, R116): an inline `continuous`
        // multi-column sectPr applies its column layout to the content that
        // PRECEDES it (the section ending AT that paragraph). The mechanism-2
        // branch below (and the page-body BuildColBodyStyle) only cover the
        // opposite shape — the multi-col section's content FOLLOWING an empty
        // leading delimiter paragraph (vacancy, R85), or a multi-col section
        // that owns a whole page. When the multi-col section has real content
        // BEFORE its break AND a fewer-col section follows on the SAME server
        // page (an IEEE 2-col body whose 1-col References tail shares the page),
        // the page-body picks the LAST (1-col) section so the whole body renders
        // single-column, and the after-break div wraps the wrong (later) content
        // — emitting an empty `<div column-count:2>` while the 2-col body escapes
        // it. Detect those sections here and wrap exactly their paragraph range
        // (section start → break paragraph) in a scoped column-count div.
        //
        // Discriminator vs the vacancy (after-break) shape: a "pre-break" section
        // has renderable content in the elements BEFORE its closing sectPr-bearing
        // paragraph. The vacancy 2-col delimiter paragraph is the ONLY element of
        // its section (no prior content) → not matched here, still handled by the
        // existing after-break branch.
        var preBreakColOpenAt = new Dictionary<int, (int cols, string gap, bool sep)>();
        var preBreakColCloseAfter = new HashSet<int>();
        {
            int sectionStartIdx = 0; // first element index of the current section
            for (int ei = 0; ei < elements.Count; ei++)
            {
                if (elements[ei] is Paragraph bp
                    && bp.ParagraphProperties?.GetFirstChild<SectionProperties>() is SectionProperties bpSect)
                {
                    var bpType = bpSect.GetFirstChild<SectionType>()?.Val?.Value;
                    var bpCols = GetSectionColumnCount(bpSect);
                    // BUG(first-section-double-col): skip the FIRST section
                    // (sectionStartIdx == 0). Page 1's page-body already adopts
                    // section 0's column layout (firstSection / activeSectionIdx
                    // starts at 0), so emitting a scoped pre-break column div for
                    // section 0 too nests `column-count:2` inside `column-count:2`,
                    // halving the usable width again — a 44pt title in a leading
                    // 2-col continuous section (e.g. a cover table) then renders
                    // into a ~110pt sub-column and wraps mid-word. The R116
                    // pre-break wrapper exists only for LATER multi-col sections
                    // whose columns the page-body does NOT own (it picked a
                    // different/fewer-col active section); section 0 is always
                    // owned by the page-body, so its wrapper is pure redundancy.
                    if (bpType == SectionMarkValues.Continuous && bpCols > 1
                        && sectionStartIdx > 0)
                    {
                        // Does this section carry real content BEFORE its break?
                        bool hasContentBefore = false;
                        for (int k = sectionStartIdx; k < ei; k++)
                        {
                            if (elements[k] is Table) { hasContentBefore = true; break; }
                            if (elements[k] is Paragraph kp && ParagraphHasRenderableContent(kp))
                            { hasContentBefore = true; break; }
                        }
                        if (hasContentBefore)
                        {
                            var sc = bpSect.GetFirstChild<Columns>();
                            var sep = sc?.Separator?.Value == true;
                            var space = sc?.Space?.Value;
                            var gap = int.TryParse(space, out var sp) && sp > 0
                                ? (sp / 20.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)
                                : "36";
                            preBreakColOpenAt[sectionStartIdx] = (bpCols, gap, sep);
                            preBreakColCloseAfter.Add(ei);
                        }
                    }
                    sectionStartIdx = ei + 1; // next section starts after this break
                }
            }
        }
        bool inPreBreakMultiCol = false;

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

            // BUG(content-before-multicol-break, R116): close the scoped
            // pre-break multi-column div once its section-closing paragraph
            // (emitted in the PREVIOUS iteration) has rendered. Mirrors the
            // section-advance timing below so the break paragraph stays inside
            // the columns it belongs to.
            if (inPreBreakMultiCol && ei > 0 && preBreakColCloseAfter.Contains(ei - 1))
            {
                sb.AppendLine("</div>");
                inPreBreakMultiCol = false;
            }
            // (open emitted below, AFTER the section-advance PAGE_BREAK so a
            // section starting on a fresh page keeps the <div> on the new page)

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
                // ECMA-376 §17.6.22: an omitted <w:type> defaults to
                // "nextPage" — the section break still starts a new page.
                // Only "continuous" and the same-page "nextColumn" variant
                // stay on the current page. Treating a missing type as
                // "no break" merged consecutive sections onto one page, so the
                // page's LAST <!--SECT:N--> marker (picked by sectMatches[^1])
                // pointed at a later section than the page's real owner —
                // flipping page-1 header/footer selection to the wrong
                // section's variant.
                var sectTypeVal = prevInlineSectPr.GetFirstChild<SectionType>()?.Val?.Value;
                if (sectTypeVal != SectionMarkValues.Continuous
                    && sectTypeVal != SectionMarkValues.NextColumn)
                {
                    sb.Append("<!--PAGE_BREAK-->");
                }
                currentSectionIdx++;
                sb.Append($"<!--SECT:{currentSectionIdx}-->");
                ApplySectionFnSettings(allSections, currentSectionIdx);

                // BUG(trailing-continuous-cols, R87): the TRAILING body sectPr
                // (body.GetFirstChild<SectionProperties>(), the LAST entry of
                // CollectSections) is never inline on a paragraph, so it never
                // hits the inline-break multi-column branch below — that branch
                // only fires for an inline sectPr ON a section-closing paragraph.
                // When that trailing section is `continuous` AND multi-col (a
                // 2-col References tail under a 1-col paper body), its multi-col
                // layout otherwise falls through ONLY to the page-body-level
                // BuildColBodyStyle, which applies it to the WHOLE page —
                // reverse-leaking 2 columns onto the earlier 1-col body. Scope it
                // here with its own column-count div so only the trailing content
                // is multi-col; the page-body decision (above) keeps the earlier
                // section's column count for the rest of the page.
                // Narrowly gated to: (a) entering the LAST section, (b) which is
                // the trailing body sectPr (not inline → not handled below),
                // (c) a true column INCREASE vs the section we left. This cannot
                // touch the inline-multi-col path (R85, handled by the branch
                // below) or single-section docs (no advance happens).
                var trailingBodySectPr = body.GetFirstChild<SectionProperties>();
                if (!inMultiColumn
                    && currentSectionIdx == allSections.Count - 1
                    && allSections.Count >= 2
                    && ReferenceEquals(allSections[currentSectionIdx], trailingBodySectPr))
                {
                    var enteredCols = GetSectionColumnCount(allSections[currentSectionIdx]);
                    var leftCols = GetSectionColumnCount(allSections[currentSectionIdx - 1]);
                    if (enteredCols > 1 && enteredCols > leftCols)
                    {
                        var sc = allSections[currentSectionIdx].GetFirstChild<Columns>();
                        var sep = sc?.Separator?.Value == true;
                        var space = sc?.Space?.Value;
                        var gap = int.TryParse(space, out var sp) && sp > 0
                            ? (sp / 20.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)
                            : "36";
                        sb.AppendLine($"<div style=\"column-count:{enteredCols};column-gap:{gap}pt"
                            + (sep ? ";column-rule:1px solid #000" : "") + "\">");
                        inMultiColumn = true;
                    }
                }
            }

            // BUG(content-before-multicol-break, R116): open the scoped pre-break
            // multi-column div at the START of a section whose content precedes
            // its multi-col continuous break. Emitted AFTER the section-advance
            // PAGE_BREAK above so a section starting on a fresh page keeps the
            // <div> on the new page (never split across the page boundary).
            if (!inPreBreakMultiCol && !inMultiColumn
                && preBreakColOpenAt.TryGetValue(ei, out var pbc))
            {
                sb.AppendLine($"<div style=\"column-count:{pbc.cols};column-gap:{pbc.gap}pt"
                    + (pbc.sep ? ";column-rule:1px solid #000" : "") + "\">");
                inPreBreakMultiCol = true;
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
                // A `continuous` section break's own w:cols defines the column
                // layout for content FOLLOWING the break in the same page. The
                // next-sectPr scan returns the body sectPr fallback (usually 1
                // column), missing the change. Read this sectPr's own w:cols.
                //
                // R116 exception: when THIS break's multi-col section is the one
                // whose CONTENT PRECEDED it (handled by the scoped pre-break div
                // above), its own w:cols describes the content already wrapped,
                // NOT the content that follows. Reading it here would re-open a
                // multi-col div over the next (fewer-col) section. Fall through to
                // the next-section scan so the following content gets its real
                // column count.
                var sectType = inlineSectPr.GetFirstChild<SectionType>()?.Val?.Value;
                var nextCols = (sectType == SectionMarkValues.Continuous
                                && !preBreakColCloseAfter.Contains(ei))
                    ? GetSectionColumnCount(inlineSectPr)
                    : GetNextSectionColumnCount(elements, ei, bodyColCount);
                // BUG(content-before-multicol-break, R116): GetNextSectionColumnCount
                // looks PAST this break to the next inline sectPr's w:cols. When the
                // next section is a "pre-break" multi-col one (its content precedes
                // its OWN closing break, e.g. an IEEE 2-col body that follows a
                // nextPage title section), that scan returns the 2-col count and this
                // branch would open a multi-col div HERE — on the wrong content, and
                // set inMultiColumn=true so the scoped pre-break div (which actually
                // wraps that section) is blocked from opening (its open guard is
                // !inMultiColumn). The pre-break branch owns that section, so skip
                // this inline open/close entirely when the section START immediately
                // following this break (ei+1) is a pre-break key.
                if (!preBreakColOpenAt.ContainsKey(ei + 1))
                {
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
                    // Reserve the drop cap paragraph's own spacing-after below
                    // the floated frame so consecutive drop cap structures
                    // stack with the same gap as ordinary paragraphs.
                    var dcAfterPt = ResolveParaAfterSpacingPt(para);
                    var marginCss = dcAfterPt > 0 ? $";margin-bottom:{dcAfterPt:0.##}pt" : "";
                    sb.Append($"<div class=\"dropcap-wrap\" style=\"display:block;overflow:hidden{marginCss}\">");
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

                    // Re-arm per-instance startOverrides whenever the active num
                    // changes (including the first list): the override fires on
                    // the first DIRECT advance of its level under the new num,
                    // even when a deeper item carried the level over. Mirrors the
                    // text path (GetListPrefix), which arms on every CurrentNumId
                    // change — both feed the same AdvanceOrderedCounter.
                    if (numId != currentNumId)
                    {
                        olState.PendingInstanceOverride.Clear();
                        // Skip lvlRestart="0" levels: a never-restart continued
                        // level keeps counting across the switch rather than
                        // jumping to the override value (see GetListPrefix).
                        for (int lv = 0; lv <= 8; lv++)
                            if (GetNumInstanceOverrideStart(numId, lv).HasValue
                                && GetEffectiveLvlRestart(numId, lv) != 0)
                                olState.PendingInstanceOverride.Add(lv);
                    }

                    // When numId changes, decide: nesting or new list
                    if (currentNumId != null && numId != currentNumId)
                    {
                        if (listStack.Count > 0 && !numIdLevelOffset.ContainsKey(numId))
                        {
                            var curIndent = GetListLevelIndent(currentNumId.Value, currentListLevel);
                            var newIndent = GetListLevelIndent(numId, ilvl);
                            // Nesting DEPTH must follow the paragraph's ilvl, not the
                            // numbering-definition indent. A new-numId item only nests
                            // DEEPER when its raw ilvl actually descends below the
                            // previous item's ilvl; if the ilvl is the same or shallower
                            // it is a return/un-nest, so fall through to CloseAllLists +
                            // the regular ilvl-driven close loop regardless of indent.
                            // (Without this guard, a returning L0 item on a different
                            // numId whose lvl0 indent happens to exceed the previous
                            // item's lvlN indent was pushed two levels too deep.)
                            if (newIndent > curIndent && ilvl > prevOoxmlIlvl)
                            {
                                numIdLevelOffset[numId] = currentListLevel + 1 - ilvl;
                            }
                            else
                            {
                                CloseAllLists(sb, listStack, ref currentListType, ref pendingLiClose);
                                olState.OlCountPerLevel.Clear();
                                olState.MultiLevelCounters.Clear();
                            }
                        }
                        else if (listStack.Count == 0)
                        {
                            // Previous list was closed by non-list content — reset counters for new list
                            olState.OlCountPerLevel.Clear();
                            olState.MultiLevelCounters.Clear();
                            numIdLevelOffset.Clear();
                        }
                    }
                    // Preserve the paragraph's OOXML ilvl (the index into
                    // <w:abstractNum><w:lvl>) for level-rPr / level-pPr lookups
                    // that must read the level definition the doc actually
                    // references. The HTML rendering may stack this list a
                    // step deeper for visual nesting across numIds; that
                    // structural depth uses the bumped value.
                    var ilvlOoxml = ilvl;
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

                    // Get indentation from numbering level definition, then let
                    // the paragraph's own <w:ind> override it (BUG-R105:
                    // paragraph-direct indentation supersedes the level value).
                    var (lvlLeft, lvlHanging) = ResolveListIndent(para, numId, ilvl);
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
                        // Map Word bullet character to CSS list-style-type.
                        // CONSISTENCY(bullet-glyph-map): shared with table-cell
                        // path and GetCustomListStyleString; default disc.
                        // Symbol-font bullets resolve to the custom glyph string
                        // so an inline keyword doesn't override the ::marker.
                        var bulletType = GetUlListStyleTypeCss(numId, ilvl, lvlText);
                        listStyleParts += $";list-style-type:{bulletType}";
                        // CONSISTENCY(bullet-text-indent-reset): a bullet uses the
                        // native ::marker (outside the content box), so it must NOT
                        // carry a hanging text-indent. When a bullet list nests
                        // under an ordered item, the parent <li>'s negative
                        // text-indent (calc(-markerWidth - padding)) inherits into
                        // this <ul> and pulls the first line left OVER the disc,
                        // hiding the marker behind the text. Reset to 0 here so the
                        // disc sits in the padding where Word draws it. Ordered
                        // children re-establish their own hanging indent per item.
                        listStyleParts += ";text-indent:0";
                    }
                    var indentStyle = $" style=\"{listStyleParts}\"";

                    // Counter seeding precedence (in-run → startOverride →
                    // abstractNum continuation → level start) lives in
                    // SeedOrderedStart, shared with the text walker.
                    var seedAbsId = GetAbstractNumId(numId);

                    while (listStack.Count < ilvl + 1)
                    {
                        // Nested-deeper open: the previous list item (parent) is
                        // about to host this <ol>/<ul> as its child, so the
                        // parent <li>'s margin-bottom applies AFTER the nested
                        // list ends, not between the parent text and the first
                        // nested item. OOXML §17.3.1.4 spaceAfter applies per
                        // paragraph regardless of list nesting; promote the
                        // parent's after-spacing to the nested list's margin-top
                        // so consecutive level transitions get the same vertical
                        // gap as same-level siblings.
                        var nestedStyle = indentStyle;
                        if (listStack.Count > 0)
                        {
                            var parentPara = para.PreviousSibling<Paragraph>();
                            if (parentPara != null)
                            {
                                var parentAfterPt = ResolveParaAfterSpacingPt(parentPara);
                                if (parentAfterPt > 0)
                                {
                                    var styleAttr = nestedStyle.TrimEnd('"');
                                    nestedStyle = styleAttr + $";margin-top:{parentAfterPt:0.##}pt\"";
                                }
                            }
                        }
                        sb.AppendLine($"<{tag}{nestedStyle}>");
                        listStack.Push(tag);
                    }
                    // If same level but different list type, swap
                    if (listStack.Count > 0 && listStack.Peek() != tag)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        sb.AppendLine($"<{tag}{indentStyle}>");
                        listStack.Push(tag);
                    }

                    // Advance the ordered counter (increment level, reset
                    // deeper levels, mirror to the abstractNum store) via the
                    // shared state machine — identical to the text walker.
                    if (tag == "ol")
                        AdvanceOrderedCounter(olState, numId, seedAbsId, ilvl);

                    currentListType = listStyle;
                    currentListLevel = ilvl;
                    prevOoxmlIlvl = ilvlOoxml;
                    currentNumId = numId;
                    sb.Append("<li");
                    sb.Append($" data-path=\"/body/p[{wParaCount}]\"");
                    // Marker class wires up the ::marker rule emitted by
                    // BuildListMarkerCss so this <li> picks up the abstractNum
                    // level rPr (color/font/size/bold/italic) for ul, plus
                    // a custom list-style-type string when applicable.
                    sb.Append($" class=\"marker-{numId}-{ilvlOoxml}\"");
                    var paraStyle = GetParagraphInlineCss(para, isListItem: true);
                    // ul markers render via ::marker pseudo, which sits outside
                    // the line box and can't inflate it. ol markers render via
                    // an inline-block <span> that already contributes its full
                    // height — the precise line-height there is enough.
                    if (tag == "ul")
                    {
                        var liLh = GetListItemLineHeightOverride(numId, ilvlOoxml, para);
                        if (liLh.HasValue)
                        {
                            var rx = new System.Text.RegularExpressions.Regex(@"line-height:[^;]+");
                            var replacement = $"line-height:{liLh.Value:0.##}pt";
                            paraStyle = rx.IsMatch(paraStyle)
                                ? rx.Replace(paraStyle, replacement)
                                : (string.IsNullOrEmpty(paraStyle) ? replacement : paraStyle + ";" + replacement);
                        }
                    }
                    // Compute the ordered-list marker (single or multi-level)
                    // and its box width up front so the <li> can host a
                    // HANGING-INDENT layout: the marker span sits to the LEFT
                    // of the text (negative text-indent of the marker width)
                    // rather than pushing the text right. This makes numbered
                    // text land at the same padding-left boundary as bullet
                    // text — bullet markers render via ::marker (outside the
                    // content box), so their text starts at padding-left; the
                    // ol marker is an inline-block <span> INSIDE the box, which
                    // without the hanging indent shoved the text right by the
                    // marker width (~18pt). Real Word aligns the two. The
                    // padding-right is pt-based (not em) so the negative
                    // text-indent cancels the marker width exactly.
                    string? olMarkerSpan = null;
                    if (tag == "ol")
                    {
                        var marker = RenderOrderedMarker(olState, numId, ilvl, lvlText);
                        var suff = GetLevelSuffix(numId, ilvl);
                        var jc = GetLevelJustification(numId, ilvl);
                        var markerWidth = hangingPt > 0 ? $"{hangingPt:0.#}pt" : "3em";
                        var markerPadding = suff switch
                        {
                            "nothing" => "0",
                            "space" => "0.25em",
                            _ => "0.5em" // tab
                        };
                        // The marker span is a HANGING box pulled fully left of the
                        // text (text-indent below). Its RIGHT edge sits at the text
                        // start (padding-left), so the number must be RIGHT-aligned
                        // inside the box to HUG the text — landing where the bullet
                        // ::marker disc sits (just left of the content box). A
                        // left-aligned number would float to the far-left of the
                        // hanging box, opening a wide gap to the text and landing the
                        // marker well left of the bullet glyph.
                        //
                        // The bullet (ul) reference ignores lvlJc entirely — the CSS
                        // ::marker hugs the content box regardless. To keep ol markers
                        // visually parallel with bullets, the default lvlJc=left (the
                        // value Word's template writes for ordinary lists) ALSO hugs
                        // via right-align. Only an explicit lvlJc=center keeps a
                        // distinct centered layout. (Explicit right is right too.)
                        var align = jc switch { "center" => "center", _ => "right" };
                        // Pull in marker-level rPr (color/font/size/bold/italic) so
                        // the ol marker span matches the styling emitted globally
                        // for ul ::marker. Word lets per-level rPr restyle markers
                        // independent of the body run; mirroring that here keeps
                        // sections like "red bold 1." parallel between ol/ul.
                        var inlineMarkerCss = GetMarkerInlineCss(numId, ilvlOoxml, para);
                        var markerStyle = $"display:inline-block;min-width:{markerWidth};padding-right:{markerPadding};text-align:{align}";
                        if (!string.IsNullOrEmpty(inlineMarkerCss))
                            markerStyle = inlineMarkerCss + ";" + markerStyle;
                        olMarkerSpan = $"<span style=\"{markerStyle}\">{HtmlEncode(marker)}</span>";
                        // Hanging indent: pull the marker (min-width + padding-right)
                        // into the padding so the text resumes at padding-left
                        // (aligned with the bullet text). calc() lets us cancel the
                        // mixed pt + em marker box exactly regardless of font size.
                        var hangCss = $"text-indent:calc(-{markerWidth} - {markerPadding})";
                        paraStyle = string.IsNullOrEmpty(paraStyle) ? hangCss : paraStyle + ";" + hangCss;
                    }
                    if (!string.IsNullOrEmpty(paraStyle))
                        sb.Append($" style=\"{paraStyle}\"");
                    sb.Append(">");
                    if (olMarkerSpan != null)
                        sb.Append(olMarkerSpan);
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
                    // Browser default `<hN>{font-weight:bold}` forces every heading
                    // bold, but Word styles like `Title` deliberately render thin —
                    // their pStyle chain has no <w:b/> and inherits from Normal
                    // which also isn't bold. Emit `font-weight:normal` whenever
                    // the resolved chain doesn't EXPLICITLY say bold (true).
                    // Heading 1 etc. carry <w:b/> in their style → keep h1's
                    // browser-default bold.
                    var pStyleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    // Only force normal when the style chain EXPLICITLY resolves to
                    // non-bold (false). null = style not in doc → defer to browser
                    // <hN> bold default instead of stomping it.
                    if (ResolveStyleBold(pStyleId) == false)
                        hStyle = string.IsNullOrEmpty(hStyle) ? "font-weight:normal" : $"{hStyle};font-weight:normal";
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
                            // Only %1..%9 are valid Word level placeholders.
                            // Match the RenderOrderedMarker hardening: restrict
                            // to %1-%9 (so %0 / %x stays literal) and emit ""
                            // for a placeholder whose fmt resolves to "bullet"
                            // or an undefined level — otherwise a heading number
                            // gets polluted with a • glyph and diverges from
                            // view text. See WordHandler.StyleList.cs
                            // RenderOrderedMarker.
                            var numStr = System.Text.RegularExpressions.Regex.Replace(lvlText, @"%([1-9])", m =>
                            {
                                var lk = int.Parse(m.Groups[1].Value) - 1;
                                var lvlFmt = GetNumberingFormat(hn.NumId, lk);
                                if (lvlFmt.Equals("bullet", StringComparison.OrdinalIgnoreCase))
                                    return "";
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

                    // Inline equation only
                    if (mathElements.Count > 0 && runs.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                        sb.AppendLine($"<div class=\"equation\"><span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(latex)}\" data-display=\"true\"></span></div>");
                        continue;
                    }

                    // Block-level drawings (anchored textboxes / charts / shapes)
                    // emit float:left <div>s. A <div> inside <p> is invalid HTML
                    // and browsers auto-close the <p> before the <div>, breaking
                    // float layout. Promote to <div> when the paragraph contains
                    // any block-level drawing. CONSISTENCY: mirrors RenderParagraphHtml.
                    var pTag = HasBlockLevelDrawing(para) ? "div" : "p";
                    sb.Append("<").Append(pTag);
                    sb.Append($" data-path=\"/body/p[{wParaCount}]\"");
                    // Add CSS class for TOC paragraphs (suppress hyperlink styling, enable dot leaders).
                    // Match by resolved style NAME too, not just styleId prefix: WPS / localized
                    // Word emit numeric styleIds (e.g. "28") whose display name is "toc 1", so a
                    // styleId-only test silently misses their TOC entries and leaks the Hyperlink
                    // character-style color. See IsTocParagraphStyle in HtmlPreview.Text.cs.
                    var paraStyleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    var classNames = new List<string>();
                    if (IsTocParagraphStyle(paraStyleId, GetStyleName(para)))
                        classNames.Add("toc");
                    // CONSISTENCY(run-special-content): body-path render must
                    // also flag has-ptab so the paragraph becomes a flex
                    // container — without this, body and table-cell ptabs
                    // collapse into a single line (only the header/footer
                    // render path went through RenderParagraphHtml which had
                    // the class added in Round 2).
                    if (para.Descendants<PositionalTab>().Any())
                        classNames.Add("has-ptab");
                    if (ParagraphHasLeaderTab(para))
                        classNames.Add("has-leader-tab");
                    else if (ParagraphHasAlignedTab(para))
                        classNames.Add("has-aligned-tab");
                    if (classNames.Count > 0)
                        sb.Append($" class=\"{string.Join(" ", classNames)}\"");
                    var pStyle = GetParagraphInlineCss(para);
                    // A body paragraph that anchors a wrapNone shape positioned
                    // relative to the column/paragraph (see
                    // ComputeParagraphAnchorAbsoluteCss) is the position:relative
                    // containing block for those absolutely positioned shapes —
                    // mirrors BuildParagraphOpenTag (header/footer path) and the
                    // table-cell <td> path. Without it the shapes resolve against
                    // the .page box and lose their per-shape posOffset.
                    if (ParagraphAnchorsSubParagraphShape(para))
                        pStyle = string.IsNullOrEmpty(pStyle) ? "position:relative" : pStyle + ";position:relative";
                    if (!string.IsNullOrEmpty(pStyle))
                        sb.Append($" style=\"{pStyle}\"");
                    sb.Append(">");
                    // Use rendered-output length as the source of truth: a
                    // paragraph might have <w:r> with empty <w:t> (counts as
                    // a run but produces zero visible content). Anything that
                    // emits nothing collapses the line box in the browser, so
                    // a placeholder &nbsp; is needed to preserve line-height.
                    var lenBefore = sb.Length;
                    RenderParagraphContentHtml(sb, para);
                    if (sb.Length == lenBefore) sb.Append("&nbsp;");
                    sb.Append("</").Append(pTag).AppendLine(">");
                    AppendW14ReflectionBlock(sb, para, pTag, pStyle);
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
                // Thread the body walk's ordered-list counter so a table cell's
                // <ol> continues document-flow numbering (Word advances the
                // counter through table paragraphs too). (CONSISTENCY(list-marker))
                RenderTableHtmlPaged(sb, table, dataPath: $"/body/table[{wTableCount}]", olState: olState);
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
        if (inPreBreakMultiCol) sb.AppendLine("</div>"); // R116: section ran to doc end
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
    /// <paramref name="EvenExplicit"/> records that an even-type
    /// header/footer reference exists for this section even when its
    /// content is empty (so <see cref="Even"/> stays null). An explicitly
    /// referenced-but-empty even header means Word renders a BLANK even
    /// page — it must NOT inherit the odd/default content.
    /// </summary>
    private record HeaderFooterBundle(
        string? First, string? Default, string? Even, bool EvenExplicit = false);

    /// <summary>Per-variant PAGE / NUMPAGES presence flags, mirroring the three
    /// HTML slots in <see cref="HeaderFooterBundle"/>. A slot is set only when
    /// the corresponding source part actually defines that field, so the
    /// digit-rewrite is gated and never clobbers literal numbers.</summary>
    private record HeaderFooterFieldBundle(
        (bool page, bool numPages) First,
        (bool page, bool numPages) Default,
        (bool page, bool numPages) Even);

    /// <summary>
    /// #3: walk each section's HeaderReference or FooterReference elements,
    /// resolve to the underlying part, pre-render to HTML, and bucket by
    /// type. Returns a dict keyed by section index.
    /// </summary>
    private Dictionary<int, HeaderFooterBundle> BuildSectionHfBundles(
        List<SectionProperties> sections, bool isHeader,
        out Dictionary<int, HeaderFooterFieldBundle> fieldFlags)
    {
        var result = new Dictionary<int, HeaderFooterBundle>();
        fieldFlags = new Dictionary<int, HeaderFooterFieldBundle>();
        var noField = (false, false);
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return result;
        for (int i = 0; i < sections.Count; i++)
        {
            string? first = null, def = null, even = null;
            bool evenExplicit = false;
            (bool, bool) firstF = noField, defF = noField, evenF = noField;
            var refs = isHeader
                ? sections[i].Elements<HeaderReference>().Cast<OpenXmlElement>()
                : sections[i].Elements<FooterReference>().Cast<OpenXmlElement>();
            foreach (var @ref in refs)
            {
                var rId = @ref.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value;
                var typeAttr = @ref.GetAttributes().FirstOrDefault(a => a.LocalName == "type").Value;
                if (string.IsNullOrEmpty(rId)) continue;
                // BUG-evenheader: an even-type reference that resolves to a
                // real part marks the section as explicitly defining its even
                // header/footer EVEN when the part is empty (no content →
                // html stays null below). Word renders a blank even page in
                // that case; the explicit flag suppresses odd/default bleed.
                if (typeAttr == "even")
                {
                    try
                    {
                        if (isHeader && mainPart.GetPartById(rId) is HeaderPart ehp && ehp.Header != null)
                            evenExplicit = true;
                        else if (!isHeader && mainPart.GetPartById(rId) is FooterPart efp && efp.Footer != null)
                            evenExplicit = true;
                    }
                    catch { /* part missing; not explicit */ }
                }
                string? html = null;
                (bool, bool) flags = noField;
                try
                {
                    if (isHeader && mainPart.GetPartById(rId) is HeaderPart hp && hp.Header != null
                        && HeaderFooterHasContent(hp.Header))
                    {
                        var sb = new StringBuilder();
                        var savedHost = _ctx.ImageHostPart;
                        _ctx.ImageHostPart = hp;
                        // Watermark lifted out of .doc-header so it centers on
                        // the whole page (see RenderHeaderFooterHtml).
                        var bodySb = new StringBuilder();
                        var watermarkSb = new StringBuilder();
                        RenderHeaderFooterBody(bodySb, hp.Header, watermarkSb);
                        _ctx.ImageHostPart = savedHost;
                        sb.Append(watermarkSb);
                        sb.Append("<div class=\"doc-header\">");
                        sb.Append(bodySb);
                        sb.Append("</div>");
                        html = sb.ToString();
                        flags = HeaderFooterFieldFlags(hp.Header);
                    }
                    else if (!isHeader && mainPart.GetPartById(rId) is FooterPart fp && fp.Footer != null
                        && HeaderFooterHasContent(fp.Footer))
                    {
                        var sb = new StringBuilder();
                        var savedHost = _ctx.ImageHostPart;
                        _ctx.ImageHostPart = fp;
                        var bodySb = new StringBuilder();
                        var watermarkSb = new StringBuilder();
                        RenderHeaderFooterBody(bodySb, fp.Footer, watermarkSb);
                        _ctx.ImageHostPart = savedHost;
                        sb.Append(watermarkSb);
                        sb.Append("<div class=\"doc-footer\">");
                        sb.Append(bodySb);
                        sb.Append("</div>");
                        html = sb.ToString();
                        flags = HeaderFooterFieldFlags(fp.Footer);
                    }
                }
                catch { /* part missing; skip */ }
                if (html == null) continue;
                switch (typeAttr)
                {
                    case "first": first = html; firstF = flags; break;
                    case "even":  even = html; evenF = flags; break;
                    default:      def = html; defF = flags; break;
                }
            }
            result[i] = new HeaderFooterBundle(first, def, even, evenExplicit);
            fieldFlags[i] = new HeaderFooterFieldBundle(firstF, defF, evenF);
        }
        // ECMA-376 §17.10.1: a section that does not define its own
        // header/footer reference of a given type inherits. In practice
        // (and in Word's rendering) a doc with refs on only one section
        // applies them to all sections — propagate forward, then backward,
        // so sections lacking own refs pick up whatever the document
        // actually defines instead of falling through to fallbackHtml.
        for (int i = 1; i < sections.Count; i++)
        {
            var prev = result[i - 1];
            var cur = result[i];
            // Inherit the field flags alongside the HTML they describe: when a
            // slot's HTML is inherited from a neighbour, its PAGE/NUMPAGES flags
            // must come from the same neighbour, not stay at this section's
            // (empty) default.
            var prevF = fieldFlags[i - 1];
            var curF = fieldFlags[i];
            result[i] = new HeaderFooterBundle(
                cur.First ?? prev.First,
                cur.Default ?? prev.Default,
                cur.Even ?? prev.Even,
                // A section with no even definition of its own (neither content
                // nor an explicit empty reference) inherits the neighbour's
                // explicit-even flag along with its Even html.
                cur.EvenExplicit || (cur.Even == null && prev.EvenExplicit));
            fieldFlags[i] = new HeaderFooterFieldBundle(
                cur.First != null ? curF.First : prevF.First,
                cur.Default != null ? curF.Default : prevF.Default,
                cur.Even != null ? curF.Even : prevF.Even);
        }
        for (int i = sections.Count - 2; i >= 0; i--)
        {
            var next = result[i + 1];
            var cur = result[i];
            var nextF = fieldFlags[i + 1];
            var curF = fieldFlags[i];
            result[i] = new HeaderFooterBundle(
                cur.First ?? next.First,
                cur.Default ?? next.Default,
                cur.Even ?? next.Even,
                cur.EvenExplicit || (cur.Even == null && next.EvenExplicit));
            fieldFlags[i] = new HeaderFooterFieldBundle(
                cur.First != null ? curF.First : nextF.First,
                cur.Default != null ? curF.Default : nextF.Default,
                cur.Even != null ? curF.Even : nextF.Even);
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
        // BUG-R101: titlePg OFF (or absent). Per ECMA-376 §17.10.6, the
        // "first" variant is used ONLY when titlePg is set; with titlePg off
        // the first page uses the DEFAULT header/footer (or nothing when no
        // default is defined). Resolve to bundle.Default ?? "" here — never to
        // fallbackHtml, which is the first content-bearing part in arbitrary
        // part order and may be the "first" variant (e.g. a DRAFT watermark
        // header referenced only as type="first"). Letting it leak put the
        // first-page header on page 1 even though titlePg was off.
        if (isFirstPageOfSection)
            return bundle.Default ?? string.Empty;
        if (evenAndOddGlobal && pageIsEven)
        {
            if (bundle.Even != null) return bundle.Even;
            // Even header/footer explicitly referenced but empty → Word
            // renders a BLANK even page. Do NOT fall through to Default,
            // which would bleed the odd-page content onto even pages.
            if (bundle.EvenExplicit) return string.Empty;
        }
        return bundle.Default ?? fallbackHtml;
    }

    /// <summary>PAGE/NUMPAGES flags for the variant that
    /// <see cref="PickHeaderFooter"/> would return on the same page — used to
    /// gate the digit-rewrite to parts that truly carry the field.
    /// <paramref name="fallbackFlags"/> applies when the section has no bundle
    /// (mirroring the fallbackHtml return).</summary>
    private static (bool page, bool numPages) PickHeaderFooterFlags(
        Dictionary<int, HeaderFooterBundle> htmlBundles,
        Dictionary<int, HeaderFooterFieldBundle> flagBundles,
        List<SectionProperties> sections,
        int sectionIdx,
        bool isFirstPageOfSection,
        bool pageIsEven,
        bool evenAndOddGlobal,
        (bool page, bool numPages) fallbackFlags)
    {
        if (!flagBundles.TryGetValue(sectionIdx, out var flags)
            || !htmlBundles.TryGetValue(sectionIdx, out var html))
            return fallbackFlags;
        var sectHasTitlePg = sectionIdx >= 0 && sectionIdx < sections.Count
            && sections[sectionIdx].GetFirstChild<TitlePage>() != null;
        if (isFirstPageOfSection && sectHasTitlePg)
            return flags.First;
        // BUG-R101 mirror: titlePg off → first page uses Default (or none).
        // Match PickHeaderFooter's "return bundle.Default ?? string.Empty":
        // no fallback-part leak. When Default html is absent the empty string
        // carries no field, so flags are (false, false).
        if (isFirstPageOfSection)
            return html.Default != null ? flags.Default : (false, false);
        if (evenAndOddGlobal && pageIsEven)
        {
            if (html.Even != null) return flags.Even;
            // Explicit-but-empty even page renders blank (see PickHeaderFooter)
            // → no field carrier, mirror the empty-string return with no flags.
            if (html.EvenExplicit) return (false, false);
        }
        // bundle.Default ?? fallbackHtml — when Default html is absent the
        // fallback HTML is used, so the fallback's flags apply too.
        return html.Default != null ? flags.Default : fallbackFlags;
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
            var contentType = (part.ContentType ?? "").ToLowerInvariant();
            // Strip media-type parameters (e.g. "text/html; charset=utf-8")
            // before comparison: Pandoc/non-Word authors commonly emit them.
            var mediaType = contentType.Split(';', 2)[0].Trim();
            // An altChunk stores the source bytes verbatim, so the payload is
            // whatever the authoring tool wrote — windows-1252 and friends are
            // routine, which is why the content type carries a charset at all.
            // Decoding everything as UTF-8 turned `café` and em-dashes into
            // U+FFFD in the preview; honour the declaration (and any BOM or
            // <meta charset> inside the payload) instead.
            using var stream = part.GetStream();
            var content = OfficeCli.Core.CharsetDecoder.Decode(
                stream, OfficeCli.Core.CharsetDecoder.ParseCharset(contentType));

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
                plain = OfficeCli.Core.DisplayText.Truncate(plain, 1000);
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

    /// <summary>BUG-R105: a list paragraph's own &lt;w:ind&gt; OVERRIDES the
    /// numbering-level indentation (ECMA-376 §17.3.1.12 — paragraph-direct
    /// indentation supersedes the value inherited from the referenced
    /// &lt;w:lvl&gt;&lt;w:pPr&gt;&lt;w:ind&gt;). Override is per-attribute: a
    /// paragraph that specifies only w:left keeps the level's hanging, etc.
    /// Returns (left, hanging) in twips, starting from the numbering-level
    /// values and replacing each slot the paragraph defines.</summary>
    private (int left, int hanging) ResolveListIndent(
        Paragraph para, int numId, int ilvl)
    {
        var (left, hanging) = GetListLevelIndentFull(numId, ilvl);
        var ind = para.ParagraphProperties?.Indentation;
        if (ind == null) return (left, hanging);
        if (ind.Left?.Value is string ls && int.TryParse(ls, out var lt))
            left = lt;
        // hanging and firstLine are mutually exclusive in OOXML. A direct
        // w:hanging replaces the level hanging; a direct w:firstLine clears
        // any hanging (first-line indent is the negative-hanging counterpart).
        if (ind.Hanging?.Value is string hs && int.TryParse(hs, out var ht))
            hanging = ht;
        else if (ind.FirstLine?.Value is string fs && int.TryParse(fs, out _))
            hanging = 0;
        return (left, hanging);
    }
}
