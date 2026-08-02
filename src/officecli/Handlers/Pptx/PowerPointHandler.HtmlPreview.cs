// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string? _cachedDocCjkFallback;

    /// <summary>
    /// Resolve a CSS CJK font-family fallback fragment for the whole document,
    /// based on the theme's MinorFont/EastAsianFont declaration. Instance
    /// wrapper around <see cref="ResolveDocCjkFallbackStatic"/>; caches the
    /// result because every shape's font-family CSS string may need it.
    /// </summary>
    private string ResolveDocCjkFallback()
        => _cachedDocCjkFallback ??= ResolveDocCjkFallbackStatic(_doc);

    /// <summary>
    /// Static counterpart of <see cref="ResolveDocCjkFallback"/> — accepts
    /// the document directly so it can be invoked from static SVG render
    /// helpers that don't carry a handler instance reference.
    ///
    /// Returns a comma-separated, individually-quoted CSS font-family
    /// fragment (no leading comma). When the document declares no CJK
    /// font in the theme — i.e. it's locale-neutral — returns a wide,
    /// language-agnostic CJK chain so any CJK glyphs in the slides still
    /// render reliably, without privileging one script's typography.
    /// </summary>
    internal static string ResolveDocCjkFallbackStatic(PresentationDocument doc)
    {
        string? themeEa = null;
        try
        {
            var masters = doc.PresentationPart?.SlideMasterParts;
            if (masters != null)
            {
                foreach (var m in masters)
                {
                    var ea = m.ThemePart?.Theme?.ThemeElements?.FontScheme?
                        .MinorFont?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                    if (!string.IsNullOrEmpty(ea)) { themeEa = ea; break; }
                }
            }
        }
        catch (System.Xml.XmlException) { }

        var locale = LocaleFontRegistry.DetectLocaleFromCjkFontName(themeEa);
        var chain = LocaleFontRegistry.GetCjkCssFallback(locale);

        // Locale-neutral fallback: when the document carries no script signal,
        // emit a broad CJK chain covering zh/ja/ko on macOS/Windows/Linux
        // without favoring one. Slides containing CJK content still render;
        // pure-Latin documents are unaffected (browsers ignore unused fonts).
        return string.IsNullOrEmpty(chain)
            ? "'PingFang SC', 'Hiragino Sans', 'Yu Gothic', 'Apple SD Gothic Neo', 'Microsoft YaHei', 'Noto Sans CJK SC'"
            : chain;
    }

    /// <summary>
    /// Generate a self-contained HTML file that previews all slides.
    /// Each slide is rendered as an absolutely-positioned div with CSS styling.
    /// Images are embedded as base64 data URIs.
    /// </summary>
    public string ViewAsHtml(int? startSlide = null, int? endSlide = null, int gridCols = 0, int viewportPx = 1600)
    {
        // CSS demands '.' as the decimal separator; under comma-decimal locales
        // (de-DE, fr-FR, …) every `$"{double}pt"` interpolation deep in the
        // renderer would emit `141,73pt`, producing invalid CSS. Switching the
        // thread culture once at the entry point covers every nested helper
        // without auditing each interpolation.
        using var _cul = InvariantCultureScope.Enter();
        ResetModel3DRenderState();
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        // Get slide dimensions
        var (slideWidthEmu, slideHeightEmu) = GetSlideSize();
        double slideWidthPt = Units.EmuToPt(slideWidthEmu);
        double slideHeightPt = Units.EmuToPt(slideHeightEmu);

        // Resolve theme colors once for the whole presentation
        var themeColors = ResolveThemeColorMap();

        sb.AppendLine("<!DOCTYPE html>");
        // i18n: emit lang from the first run's <a:rPr lang=...> when present
        // (PPT carries no presentation-level language tag analogous to Word's
        // themeFontLang; per-run lang is the closest signal). RTL containers are
        // emitted PER-SHAPE / PER-PARAGRAPH (direction:rtl on shape-text and
        // unicode-bidi:embed on the para); document-wide dir="rtl" is NOT set
        // because it forces every LTR shape's default text-align to right.
        string presLang = "en";
        foreach (var sp in slideParts)
        {
            var slide = sp.Slide;
            if (slide == null) continue;
            var firstRunLang = slide.Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>()
                .Select(rp => rp.Language?.Value)
                .FirstOrDefault(l => !string.IsNullOrEmpty(l));
            if (!string.IsNullOrEmpty(firstRunLang)) { presLang = firstRunLang!; break; }
        }
        sb.AppendLine($"<html lang=\"{HtmlEncode(presLang)}\">");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        // KaTeX for math rendering — only include when any slide actually has formulas.
        // media=print + onload swap makes the CSS non-blocking so it can never stall first paint.
        bool hasMathFormulas = slideParts.Any(sp => sp.Slide?.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() == true);
        if (hasMathFormulas)
        {
            // CONSISTENCY(katex-mirror): mirror-first with CDN fallback chain — see Core/KatexAssets.
            sb.AppendLine($"<link rel=\"stylesheet\" href=\"{Core.KatexAssets.CssUrl}\" media=\"print\" onload=\"this.media='all'\" onerror=\"{Core.KatexAssets.CssOnErrorJs}\">");
            sb.AppendLine($"<script defer src=\"{Core.KatexAssets.JsUrl}\" onerror=\"{Core.KatexAssets.JsOnErrorJs("document.querySelectorAll('.katex-formula').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666'})")}\"></script>");
        }
        // Three.js for 3D model rendering (graceful degradation: shows placeholder when offline).
        // CONSISTENCY(katex-mirror): mirror-first importmap; CDN fallback lives at the
        // dynamic import() site in HtmlPreview.Shapes.cs — see Core/ThreeAssets.
        sb.AppendLine($"<script type=\"importmap\">{Core.ThreeAssets.ImportMapJson}</script>");
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateCss(slideWidthPt, slideHeightPt));
        sb.AppendLine("</style>");
        if (gridCols > 0)
        {
            // Grid override for thumbnail-style screenshot. 1pt = 4/3 px;
            // each cell gets viewportPx/cols width; scale slides to fit.
            double slideNativePx = slideWidthPt * 4.0 / 3.0;
            double padding = 24.0;
            double gap = 12.0;
            double cellPx = (viewportPx - padding - (gridCols - 1) * gap) / gridCols;
            double scale = cellPx / slideNativePx;
            sb.AppendLine("<style>");
            sb.AppendLine(".sidebar,.sidebar-toggle,.toggle-zone,.slide-label,.slide-notes,.file-title{display:none !important}");
            sb.AppendLine($".main{{display:grid !important;grid-template-columns:repeat({gridCols},1fr) !important;gap:{gap}px !important;padding:{padding / 2}px !important;margin-left:0 !important;align-items:start !important;justify-items:center !important;flex-direction:unset !important}}");
            sb.AppendLine($".slide-container{{width:100% !important;align-items:flex-start !important}}");
            sb.AppendLine($".slide-wrapper{{width:{cellPx:0.##}px !important;height:{cellPx / (slideWidthPt / slideHeightPt):0.##}px !important;overflow:hidden !important;display:block !important;position:relative !important}}");
            sb.AppendLine($".slide{{transform:scale({scale:0.######}) !important;transform-origin:top left !important;position:absolute !important;top:0 !important;left:0 !important}}");
            sb.AppendLine("</style>");
        }
        else if (startSlide.HasValue && endSlide.HasValue && startSlide.Value == endSlide.Value)
        {
            // Single-slide screenshot: drop the .main page padding/gap so the slide
            // renders flush to the captured viewport (which the screenshot path sizes
            // to the slide's native pixels). Scoped to headless so interactive
            // `view html` keeps its breathing room.
            sb.AppendLine("<style>html.headless .main{padding:0 !important;gap:0 !important}html.headless .slide{box-shadow:none !important}</style>");
        }
        // Auto-hide sidebar in headless/automated browsers (screenshot, Playwright, etc.)
        // Screenshot/automated render → flush mode. Lead with the explicit
        // '#screenshot' fragment that HtmlScreenshot appends to every capture URL
        // (deterministic, we control it); fall back to webdriver/UA sniffing so
        // external headless tools (html-screenshot.py, visual-regression) flush too.
        // Same trigger as the docx preview's SCREENSHOT flag.
        sb.AppendLine("<script>if(location.hash.indexOf('screenshot')>=0||navigator.webdriver||/HeadlessChrome/.test(navigator.userAgent))document.documentElement.classList.add('headless')</script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class=\"toggle-zone\"></div><button class=\"sidebar-toggle\" onclick=\"toggleSidebar()\">\u2630</button>");

        // ===== Sidebar (thumbnails populated by JS cloneNode to avoid duplicating base64 images) =====
        sb.AppendLine("<div class=\"sidebar\">");
        sb.AppendLine($"  <div class=\"sidebar-title\">{HtmlEncode(Path.GetFileName(_filePath))}</div>");
        // Empty thumb containers — JS will clone slide content into them
        int thumbNum = 0;
        foreach (var slidePart in slideParts)
        {
            thumbNum++;
            if (startSlide.HasValue && thumbNum < startSlide.Value) continue;
            if (endSlide.HasValue && thumbNum > endSlide.Value) break;

            sb.AppendLine($"  <div class=\"thumb\" data-slide=\"{thumbNum}\">");
            sb.AppendLine("    <div class=\"thumb-inner\"></div>");
            sb.AppendLine($"    <span class=\"thumb-num\">{thumbNum}</span>");
            sb.AppendLine("  </div>");
        }
        sb.AppendLine("</div>");

        // ===== Main content area =====
        sb.AppendLine("<div class=\"main\">");
        sb.AppendLine($"<h1 class=\"file-title\">{HtmlEncode(Path.GetFileName(_filePath))}</h1>");

        int slideNum = 0;
        foreach (var slidePart in slideParts)
        {
            slideNum++;
            if (startSlide.HasValue && slideNum < startSlide.Value) continue;
            if (endSlide.HasValue && slideNum > endSlide.Value) break;

            // R9-2: per-slide color map honoring any p:clrMapOvr on this slide.
            var slideColors = ApplySlideColorMapOverride(slidePart, themeColors);

            sb.AppendLine($"<div class=\"slide-container\" data-slide=\"{slideNum}\">");
            sb.AppendLine($"  <div class=\"slide-label\">Slide {slideNum}</div>");
            sb.AppendLine("  <div class=\"slide-wrapper\">");
            sb.Append($"    <div class=\"slide\"");

            // Slide background + inherited text defaults from master/layout/theme
            var slideStyles = new List<string>();
            var bgStyle = GetSlideBackgroundCss(slidePart, slideColors);
            if (!string.IsNullOrEmpty(bgStyle))
                slideStyles.Add(bgStyle);
            var textDefaults = GetTextDefaults(slidePart, slideColors);
            if (!string.IsNullOrEmpty(textDefaults))
                slideStyles.Add(textDefaults);
            if (slideStyles.Count > 0)
                sb.Append($" style=\"{string.Join("", slideStyles)}\"");
            sb.AppendLine(">");

            // Render slide elements + inherited layout placeholders
            RenderLayoutPlaceholders(sb, slidePart, slideColors, slideNum);
            RenderSlideElements(sb, slidePart, slideNum, slideWidthEmu, slideHeightEmu, slideColors);

            sb.AppendLine("    </div>");
            sb.AppendLine("  </div>");
            RenderSpeakerNotes(sb, slidePart);
            sb.AppendLine("</div>");
        }

        sb.AppendLine("</div>"); // main

        // Page counter
        sb.AppendLine($"<div class=\"page-counter\">1 / {slideParts.Count}</div>");

        // Navigation script
        sb.AppendLine("<script>");
        sb.AppendLine(GenerateScript());
        sb.AppendLine("</script>");
        sb.AppendLine("<script>");
        sb.AppendLine(@"(function() {
    var _katexRetries = 0;
    function fallbackKatex() {
        document.querySelectorAll('.katex-formula:not(.katex-rendered)').forEach(function(el) {
            el.textContent = el.dataset.formula;
            el.style.fontFamily = 'monospace';
            el.style.color = '#666';
            el.classList.add('katex-rendered');
        });
    }
    function renderKatex() {
        var pending = document.querySelectorAll('.katex-formula:not(.katex-rendered)');
        if (pending.length === 0) return;
        if (typeof katex === 'undefined') {
            // Lazy-load on first demand — handles watch mode where the initial
            // doc had no formulas (KaTeX tags omitted from head), then a
            // formula arrived via SSE patch.
            if (!window._katexLoading) {
                window._katexLoading = true;
                // CONSISTENCY(katex-mirror): mirror-first, CDN retry, then fallback.
                var link = document.createElement('link');
                link.rel = 'stylesheet';
                link.href = '{{KATEX_CSS}}';
                link.onerror = function() { if (!this.dataset.f) { this.dataset.f = 1; this.href = '{{KATEX_CSS_CDN}}'; } else { this.remove(); } };
                document.head.appendChild(link);
                var script = document.createElement('script');
                script.src = '{{KATEX_JS}}';
                script.onload = renderKatex;
                script.onerror = function() { var s2 = document.createElement('script'); s2.src = '{{KATEX_JS_CDN}}'; s2.onload = renderKatex; s2.onerror = fallbackKatex; document.head.appendChild(s2); };
                document.head.appendChild(script);
                return;
            }
            if (++_katexRetries > 20) { fallbackKatex(); return; }
            setTimeout(renderKatex, 100); return;
        }
        pending.forEach(function(el) {
            try {
                katex.render(el.dataset.formula, el, { throwOnError: false, displayMode: true });
                el.classList.add('katex-rendered');
            } catch(e) { el.textContent = el.dataset.formula + ' (Error: ' + e.message + '. See https://katex.org/docs/supported.html for supported syntax.)'; }
        });
    }
    // Initial render
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', renderKatex);
    else renderKatex();
    // Re-render when DOM changes (watch mode incremental updates)
    new MutationObserver(function() { renderKatex(); }).observe(document.body, { childList: true, subtree: true });
})();"
            // CONSISTENCY(katex-mirror): the verbatim block above can't interpolate;
            // substitute the KaTeX asset URLs (mirror + CDN fallback) afterwards.
            .Replace("{{KATEX_CSS}}", Core.KatexAssets.CssUrl)
            .Replace("{{KATEX_CSS_CDN}}", Core.KatexAssets.CdnCssUrl)
            .Replace("{{KATEX_JS}}", Core.KatexAssets.JsUrl)
            .Replace("{{KATEX_JS_CDN}}", Core.KatexAssets.CdnJsUrl));
        sb.AppendLine("</script>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        return sb.ToString();
    }

    /// <summary>
    /// Render a single slide's HTML fragment (slide-container div) for incremental updates.
    /// Returns null if the slide number is out of range.
    /// </summary>
    public string? RenderSlideHtml(int slideNum)
    {
        // Each slide-render call must be self-contained: the receiver (watch
        // SSE replace) has no other source for the GLB data scripts.
        ResetModel3DRenderState();
        var slideParts = GetSlideParts().ToList();
        if (slideNum < 1 || slideNum > slideParts.Count) return null;

        var (slideWidthEmu, slideHeightEmu) = GetSlideSize();
        var themeColors = ResolveThemeColorMap();
        var slidePart = slideParts[slideNum - 1];
        // R9-2: per-slide color map honoring any p:clrMapOvr on this slide.
        var slideColors = ApplySlideColorMapOverride(slidePart, themeColors);

        var sb = new StringBuilder();
        sb.AppendLine($"<div class=\"slide-container\" data-slide=\"{slideNum}\">");
        sb.AppendLine($"  <div class=\"slide-label\">Slide {slideNum}</div>");
        sb.AppendLine("  <div class=\"slide-wrapper\">");
        sb.Append($"    <div class=\"slide\"");

        var slideStyles = new List<string>();
        var bgStyle = GetSlideBackgroundCss(slidePart, slideColors);
        if (!string.IsNullOrEmpty(bgStyle))
            slideStyles.Add(bgStyle);
        var textDefaults = GetTextDefaults(slidePart, slideColors);
        if (!string.IsNullOrEmpty(textDefaults))
            slideStyles.Add(textDefaults);
        if (slideStyles.Count > 0)
            sb.Append($" style=\"{string.Join("", slideStyles)}\"");
        sb.AppendLine(">");

        RenderLayoutPlaceholders(sb, slidePart, slideColors, slideNum);
        RenderSlideElements(sb, slidePart, slideNum, slideWidthEmu, slideHeightEmu, slideColors);

        sb.AppendLine("    </div>");
        sb.AppendLine("  </div>");
        RenderSpeakerNotes(sb, slidePart);
        sb.AppendLine("</div>");

        return sb.ToString();
    }

    /// <summary>
    /// Get total slide count.
    /// </summary>
    public int GetSlideCount()
    {
        return GetSlideParts().Count();
    }

    // ==================== Speaker Notes ====================

    /// <summary>
    /// Render the slide's speaker notes (if any) as a sibling block under the
    /// slide-wrapper. R8-bt-3: prior to this, ViewAsHtml silently dropped
    /// notes — Arabic / Hebrew authors reviewing notes saw nothing.
    /// Direction is propagated from the notes body shape's first paragraph
    /// rtl flag so RTL notes render right-aligned.
    /// </summary>
    private static void RenderSpeakerNotes(StringBuilder sb, SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;
        var spTree = notesPart?.NotesSlide?.CommonSlideData?.ShapeTree;
        if (spTree == null) return;

        Shape? notesShape = null;
        foreach (var shape in spTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Index?.Value == 1)
            {
                notesShape = shape;
                break;
            }
        }
        if (notesShape == null) return;

        var paragraphs = notesShape.TextBody?.Elements<Drawing.Paragraph>().ToList()
            ?? new List<Drawing.Paragraph>();
        if (paragraphs.Count == 0) return;

        // Reduce to plain-text lines; bail if every paragraph is empty.
        var lines = paragraphs
            .Select(p => string.Concat(p.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? "")))
            .ToList();
        if (lines.All(string.IsNullOrEmpty)) return;

        // Inherit direction from the first paragraph's rtl flag (notes-level
        // direction is uniform — ApplyNotesDirection stamps every paragraph).
        bool rtl = paragraphs.FirstOrDefault()?.ParagraphProperties?.RightToLeft?.Value == true;
        var dirAttr = rtl ? " dir=\"rtl\"" : "";

        sb.AppendLine($"  <div class=\"slide-notes\"{dirAttr}>");
        sb.AppendLine("    <div class=\"slide-notes-label\">Notes</div>");
        sb.AppendLine("    <div class=\"slide-notes-body\">");
        foreach (var line in lines)
        {
            // System.Net.WebUtility.HtmlEncode is the canonical escape used
            // elsewhere in the preview — empty paragraphs render as <br/>.
            if (string.IsNullOrEmpty(line))
                sb.AppendLine("      <br/>");
            else
                sb.AppendLine($"      <div>{System.Net.WebUtility.HtmlEncode(line)}</div>");
        }
        sb.AppendLine("    </div>");
        sb.AppendLine("  </div>");
    }

    // ==================== CSS ====================

    private static string GenerateCss(double slideWidthPt, double slideHeightPt)
    {
        var aspect = slideWidthPt / slideHeightPt;
        // Dynamic CSS variables + static CSS from embedded resource
        var dynamicVars = $":root{{--slide-design-w:{slideWidthPt:0.##}pt;--slide-design-h:{slideHeightPt:0.##}pt;--slide-aspect:{aspect:0.####};}}\n";
        return dynamicVars + LoadEmbeddedResource("Resources.preview.css");
    }

    private static string GenerateScript()
    {
        return LoadEmbeddedResource("Resources.preview.js");
    }

    private static string LoadEmbeddedResource(string name)
    {
        var assembly = typeof(PowerPointHandler).Assembly;
        var fullName = $"OfficeCli.{name}";
        using var stream = assembly.GetManifestResourceStream(fullName);
        if (stream == null) return $"/* Resource not found: {fullName} */";
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    // ==================== Slide Background ====================

    private string GetSlideBackgroundCss(SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        var slide = GetSlide(slidePart);

        // R40-BG2: per OOXML, a slide's OWN <p:bg> (whether <p:bgPr> or
        // <p:bgRef>) always wins over inherited layout/master backgrounds.
        // Resolve each level top-down; at every level "bgPr OR bgRef present
        // wins before descending". Previously bgPr was collected across all
        // three levels first, so a master bgPr could shadow the slide's own
        // bgRef.
        var slideCss = LevelBackgroundCss(slide.CommonSlideData?.Background, slidePart, themeColors);
        if (slideCss != null) return slideCss;

        // Image/blip backgrounds inherited from the layout/master register their
        // r:embed relationship in the LAYOUT/MASTER part, not the slide part, so
        // the blip must be resolved against the owning part (else GetPartById throws
        // and the background is silently dropped).
        var layoutCss = LevelBackgroundCss(
            slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Background,
            (OpenXmlPart?)slidePart.SlideLayoutPart ?? slidePart, themeColors);
        if (layoutCss != null) return layoutCss;

        var masterCss = LevelBackgroundCss(
            slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background,
            (OpenXmlPart?)slidePart.SlideLayoutPart?.SlideMasterPart ?? slidePart, themeColors);
        if (masterCss != null) return masterCss;

        return "";
    }

    // R40-BG2: resolve a single level's background. At each level the explicit
    // <p:bgPr> wins; otherwise a <p:bgRef> (theme background-fill-style index +
    // schemeClr) is resolved against the theme map. Returns null when this level
    // has no background at all, so the caller can descend to the next level.
    private string? LevelBackgroundCss(Background? bg, OpenXmlPart part, Dictionary<string, string> themeColors)
    {
        if (bg == null) return null;

        var bgPr = bg.BackgroundProperties;
        if (bgPr != null)
            return BackgroundPropertiesToCss(bgPr, part, themeColors);

        // R4-3: a level can style its background via <p:bgRef> instead of
        // explicit bgPr. Resolve the bgRef's scheme color against the theme map.
        var bgRef = bg.GetFirstChild<BackgroundStyleReference>();
        if (bgRef != null)
        {
            // The bgRef idx selects an entry in the theme's <a:bgFillStyleLst>
            // (idx 1001..1003 -> entries 0..2). When that entry is a GRADIENT whose
            // stops reference phClr, the background is a tinted theme gradient — not a
            // flat color. Inject phClr = the bgRef's resolved color and emit the
            // gradient, exactly like shape fillRef (GetStyleFillRefCss). Previously the
            // idx was ignored, so every themed gradient background rendered as a solid.
            var idx = (int)(bgRef.Index?.Value ?? 0);
            var bgFill = idx >= 1001
                ? ResolveFormatScheme(part)?.BackgroundFillStyleList?.ChildElements
                    .OfType<OpenXmlElement>().ElementAtOrDefault(idx - 1001)
                : null;
            var bgRefColor = ResolveStyleMatrixRefColor(bgRef, themeColors);
            if (bgFill is Drawing.GradientFill gf)
            {
                var phHex = bgRefColor != null && bgRefColor.StartsWith('#') ? bgRefColor[1..] : null;
                var patched = phHex != null
                    ? new Dictionary<string, string>(themeColors) { ["phClr"] = phHex }
                    : themeColors;
                var css = GradientToCss(gf, patched);
                if (!string.IsNullOrEmpty(css) && css != "transparent")
                    return $"background:{css};";
            }
            if (bgRefColor != null) return $"background:{bgRefColor};";
        }
        return null;
    }

    // R4-3: resolve a <p:bgRef>/<a:*Ref> style-matrix reference's color. The
    // reference carries a direct schemeClr (or srgbClr) child; resolve it
    // through the theme map exactly like a solidFill body. The idx (which theme
    // bgFillStyle to use) is not modelled here — we surface the explicit color
    // override, which is what PowerPoint paints when present.
    private static string? ResolveStyleMatrixRefColor(OpenXmlElement styleRef, Dictionary<string, string> themeColors)
    {
        var schemeColor = styleRef.GetFirstChild<Drawing.SchemeColor>();
        if (schemeColor?.Val?.HasValue == true)
        {
            var schemeName = schemeColor.Val!.InnerText;
            // R40-BG1: in a bgRef context, <a:schemeClr val="phClr"/> means
            // "the theme's background anchor" = lt1 (bg1). phClr is never in the
            // theme color map, so map it to lt1 before lookup (invisible on the
            // default white-bg1 theme, wrong color on non-white-bg1 themes).
            if (schemeName == "phClr") schemeName = "lt1";
            if (schemeName != null && themeColors.TryGetValue(schemeName, out var themeHex))
                return ApplyColorTransforms(themeHex, schemeColor);
        }
        var srgb = styleRef.GetFirstChild<Drawing.RgbColorModelHex>();
        if (srgb?.Val?.Value != null)
            return $"#{srgb.Val.Value}";
        return null;
    }

    private static string BackgroundPropertiesToCss(BackgroundProperties bgPr, OpenXmlPart part, Dictionary<string, string> themeColors)
    {
        var solidFill = bgPr.GetFirstChild<Drawing.SolidFill>();
        if (solidFill != null)
        {
            var color = ResolveFillColor(solidFill, themeColors);
            if (color != null) return $"background:{color};";
        }

        var gradFill = bgPr.GetFirstChild<Drawing.GradientFill>();
        if (gradFill != null)
            return $"background:{GradientToCss(gradFill, themeColors)};";

        var blipFill = bgPr.GetFirstChild<Drawing.BlipFill>();
        if (blipFill != null)
        {
            var dataUri = BlipToDataUri(blipFill, part);
            if (dataUri != null)
            {
                // <a:alphaModFix amt="..."/>: PowerPoint composites the background
                // image at this alpha over the slide's (white) base — fading ONLY the
                // background. amt is 0..100000 (100000 = opaque). Emitting `opacity` on
                // the slide div faded EVERY shape/text on the slide (CSS opacity applies
                // to the whole subtree). Reproduce the blend with a translucent-white
                // overlay layer painted over the image: (1-alpha) white over the image ==
                // alpha*image + (1-alpha)*white, exactly PowerPoint's compositing, while
                // leaving all shapes fully opaque. (.slide's base is white.)
                var alphaMod = blipFill.GetFirstChild<Drawing.Blip>()?.GetFirstChild<Drawing.AlphaModulationFixed>();
                var overlay = "";
                if (alphaMod?.Amount?.HasValue == true && alphaMod.Amount.Value < 100000)
                {
                    var ov = 1.0 - alphaMod.Amount.Value / 100000.0;
                    overlay = $"linear-gradient(rgba(255,255,255,{ov:0.##}),rgba(255,255,255,{ov:0.##})),";
                }
                // R4-4: honor <a:tile> — repeat at native size rather than cover.
                return blipFill.GetFirstChild<Drawing.Tile>() != null
                    ? $"background:{overlay}url('{dataUri}') repeat;background-size:auto;"
                    : $"background:{overlay}url('{dataUri}') center/cover no-repeat;";
            }
        }

        // Pattern slide backgrounds (third-party files) — mirror shape pattFill.
        var pattFill = bgPr.GetFirstChild<Drawing.PatternFill>();
        if (pattFill != null)
            return PatternFillToCss(pattFill, themeColors) + ";";

        return "";
    }

    // ==================== Text Default Inheritance ====================

    /// <summary>
    /// Read default text styles from theme → slide master → slide layout chain.
    /// Returns CSS properties (font-family, font-size, color) that apply to all text on this slide
    /// unless overridden by individual shape/run formatting.
    ///
    /// Inheritance chain per OOXML spec:
    ///   Theme fonts → Presentation defaultTextStyle → SlideMaster bodyStyle/otherStyle
    ///   → SlideLayout → Shape TextBody defaults → Paragraph → Run
    /// </summary>
    private string GetTextDefaults(SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        var styles = new List<string>();

        // 1. Theme fonts (major = headings, minor = body)
        var theme = slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme;
        var fontScheme = theme?.ThemeElements?.FontScheme;
        var minorLatin = fontScheme?.MinorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
        var minorEa = fontScheme?.MinorFont?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;

        // Build font-family with fallbacks including CJK fonts. The CJK chain
        // is locale-driven (read from theme's east-asian font name); when the
        // document carries no script signal, ResolveDocCjkFallback returns a
        // broad cross-script chain so slides still render reliably.
        var fonts = new List<string>();
        if (!string.IsNullOrEmpty(minorLatin)) fonts.Add($"'{CssSanitize(minorLatin)}'");
        if (!string.IsNullOrEmpty(minorEa)) fonts.Add($"'{CssSanitize(minorEa)}'");
        fonts.Add(ResolveDocCjkFallback());
        fonts.Add("sans-serif");
        styles.Add($"font-family:{string.Join(",", fonts)};");

        // 2. Default text size from presentation defaultTextStyle or slide master otherStyle
        int? defaultSizeHundredths = null;
        string? defaultColorHex = null;

        // Check presentation-level defaultTextStyle
        var presDefStyle = _doc.PresentationPart?.Presentation?.DefaultTextStyle;
        if (presDefStyle != null)
        {
            var level1 = (OpenXmlCompositeElement?)presDefStyle.GetFirstChild<Drawing.DefaultParagraphProperties>()
                ?? presDefStyle.GetFirstChild<Drawing.Level1ParagraphProperties>();
            var defRp = level1?.GetFirstChild<Drawing.DefaultRunProperties>();
            if (defRp?.FontSize?.HasValue == true)
                defaultSizeHundredths = defRp.FontSize.Value;
            var defColor = ResolveFillColor(defRp?.GetFirstChild<Drawing.SolidFill>(), themeColors);
            if (defColor != null) defaultColorHex = defColor;
        }

        // Check slide master otherStyle (higher priority for body text)
        var masterTxStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        var otherStyle = masterTxStyles?.OtherStyle;
        if (otherStyle != null)
        {
            var masterLevel1 = otherStyle.GetFirstChild<Drawing.Level1ParagraphProperties>();
            var masterDefRp = masterLevel1?.GetFirstChild<Drawing.DefaultRunProperties>();
            if (masterDefRp?.FontSize?.HasValue == true)
                defaultSizeHundredths = masterDefRp.FontSize.Value;
            var masterColor = ResolveFillColor(masterDefRp?.GetFirstChild<Drawing.SolidFill>(), themeColors);
            if (masterColor != null) defaultColorHex = masterColor;

            // Font override from master
            var masterFont = masterDefRp?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrEmpty(masterFont) && !masterFont.StartsWith("+", StringComparison.Ordinal))
            {
                fonts.Insert(0, $"'{CssSanitize(masterFont)}'");
                styles[0] = $"font-family:{string.Join(",", fonts)};";
            }
        }

        if (defaultSizeHundredths.HasValue)
            styles.Add($"font-size:{defaultSizeHundredths.Value / 100.0:0.##}pt;");

        // Default text color — if not set, derive from theme dk1 (standard dark text on light bg)
        if (defaultColorHex != null)
            styles.Add($"color:{defaultColorHex};");
        else if (themeColors.TryGetValue("dk1", out var dk1))
            styles.Add($"color:#{dk1};");

        return string.Join("", styles);
    }

    // ==================== Render Slide Elements ====================

    private void RenderSlideElements(StringBuilder sb, SlidePart slidePart, int slideNum,
        long slideWidthEmu, long slideHeightEmu, Dictionary<string, string> themeColors)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return;

        // Per-element-type positional counters used to build the data-path of each
        // top-level element. We prefer @id= when the element has a cNvPr id (stable
        // across edits), and fall back to positional [N] otherwise.
        int shapeIdx = 0, picIdx = 0, tableIdx = 0, chartIdx = 0, cxnIdx = 0, groupIdx = 0, oleIdx = 0, model3dIdx = 0, smartartIdx = 0;
        string PathFor(string typeName, OpenXmlElement el, int positional)
            => $"/slide[{slideNum}]/{BuildElementPathSegment(typeName, el, positional)}";

        // Collect all content elements in z-order (as they appear in XML)
        foreach (var element in shapeTree.ChildElements)
        {
            switch (element)
            {
                case Shape shape:
                    shapeIdx++;
                    RenderShape(sb, shape, slidePart, themeColors, dataPath: PathFor("shape", shape, shapeIdx), slideNumber: slideNum);
                    break;
                case Picture pic:
                    picIdx++;
                    RenderPicture(sb, pic, slidePart, themeColors, dataPath: PathFor("picture", pic, picIdx));
                    break;
                case GraphicFrame gf:
                    if (gf.Descendants<Drawing.Table>().Any())
                    {
                        tableIdx++;
                        RenderTable(sb, gf, themeColors, dataPath: PathFor("table", gf, tableIdx), part: slidePart);
                    }
                    else if (gf.Descendants().Any(e => e.LocalName == "chart" && e.NamespaceUri.Contains("chart")))
                    {
                        chartIdx++;
                        RenderChart(sb, gf, slidePart, themeColors, dataPath: PathFor("chart", gf, chartIdx));
                    }
                    else if (gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any())
                    {
                        oleIdx++;
                        RenderOlePlaceholder(sb, gf, slidePart, dataPath: PathFor("ole", gf, oleIdx));
                    }
                    else if (gf.Descendants().Any(e =>
                                 (e.LocalName == "graphicData" && (e.GetAttributes().Any(a => a.LocalName == "uri" && a.Value != null && a.Value.Contains("diagram"))))
                                 || (e.LocalName == "relIds" && e.NamespaceUri.Contains("diagram"))))
                    {
                        smartartIdx++;
                        RenderSmartArt(sb, gf, slidePart, themeColors, dataPath: PathFor("smartart", gf, smartartIdx));
                    }
                    break;
                case ConnectionShape cxn:
                    cxnIdx++;
                    RenderConnector(sb, cxn, themeColors, dataPath: PathFor("connector", cxn, cxnIdx), part: slidePart);
                    break;
                case GroupShape grp:
                    groupIdx++;
                    RenderGroup(sb, grp, slidePart, themeColors, dataPath: PathFor("group", grp, groupIdx));
                    break;
                default:
                    // mc:AlternateContent — render 3D models, zoom, etc.
                    if (element.LocalName == "AlternateContent")
                    {
                        string? acDataPath = null;
                        if (element.Descendants().Any(d => d.LocalName == "model3d"))
                        {
                            model3dIdx++;
                            acDataPath = $"/slide[{slideNum}]/model3d[{model3dIdx}]";
                        }
                        RenderAlternateContent(sb, element, slidePart, themeColors, acDataPath);
                    }
                    break;
            }
        }
    }

    // ==================== Layout/Master Placeholder Rendering ====================

    /// <summary>
    /// Render visible placeholders from SlideLayout and SlideMaster that are not
    /// overridden by the slide itself. This includes footers, slide numbers,
    /// date/time, logos, and decorative shapes from the layout/master.
    /// </summary>
    private void RenderLayoutPlaceholders(StringBuilder sb, SlidePart slidePart, Dictionary<string, string> themeColors, int slideNum = 1)
    {
        // Collect placeholder identifiers already present on the slide
        var slidePlaceholders = new HashSet<string>();
        var slideShapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (slideShapeTree != null)
        {
            foreach (var shape in slideShapeTree.Elements<Shape>())
            {
                var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                if (ph?.Index?.HasValue == true) slidePlaceholders.Add($"idx:{ph.Index.Value}");
                if (ph?.Type?.HasValue == true) slidePlaceholders.Add($"type:{ph.Type.InnerText}");
            }
        }

        // Render shapes from SlideLayout (higher priority)
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart != null)
            RenderInheritedShapes(sb, layoutPart.SlideLayout?.CommonSlideData?.ShapeTree, layoutPart, slidePlaceholders, themeColors, slideNum);

        // Render shapes from SlideMaster (lower priority, only if not in layout).
        // R12-2: <p:sld showMasterSp="0"> suppresses master-level decoration.
        // (Layout placeholders above still render — matches PowerPoint, where
        // the flag only governs the master's own shapes.)
        var showMasterSp = GetSlide(slidePart).ShowMasterShapes?.Value ?? true;
        var masterPart = layoutPart?.SlideMasterPart;
        if (masterPart != null && showMasterSp)
            RenderInheritedShapes(sb, masterPart.SlideMaster?.CommonSlideData?.ShapeTree, masterPart, slidePlaceholders, themeColors, slideNum);
    }

    // RenderInheritedShapes — render the layout/master shapes that the slide
    // doesn't override. Two rules:
    //
    //   1. Layout/master placeholders never contribute TEXT — what's in their
    //      <p:txBody> is edit-prompt boilerplate ("Click to add title", "单击
    //      此处添加正文"). Real content always lives on the slide. The only
    //      placeholders whose text IS legitimately layout/master-supplied are
    //      the four metadata slots (date/footer/header/slide number); keep
    //      those.
    //
    //   2. ECMA-376 §19.3.1.36: a <p:ph> with no `type` attribute defaults to
    //      `obj`. Open XML SDK exposes this as `Type.HasValue == false`, so
    //      type-based logic that hinges on HasValue silently misses these
    //      shapes — that was the bug behind issue #79: a layout body
    //      placeholder authored without an explicit type leaked its prompt
    //      text onto the slide.
    private void RenderInheritedShapes(StringBuilder sb, ShapeTree? shapeTree, OpenXmlPart part,
        HashSet<string> skipIndices, Dictionary<string, string> themeColors, int slideNum = 1)
    {
        if (shapeTree == null) return;

        foreach (var element in shapeTree.ChildElements)
        {
            switch (element)
            {
                case Shape shape:
                    RenderInheritedShape(sb, shape, part, skipIndices, themeColors, slideNum);
                    break;
                // R12-1: PowerPoint renders group/connector/graphic-frame
                // decoration from the layout/master tree too. The old code
                // (`if (element is not Shape shape) continue;`) dropped them.
                // These are never placeholders, so no skip-index logic applies.
                case GroupShape grp:
                    RenderGroup(sb, grp, part, themeColors);
                    break;
                case ConnectionShape cxn:
                    RenderConnector(sb, cxn, themeColors, part: part);
                    break;
                case GraphicFrame gf:
                    // Only tables are cheap to inherit here; RenderChart needs a
                    // SlidePart (chart-part relationship lookup) which the
                    // layout/master tree doesn't provide. Layout/master charts
                    // are rare, so leave them out (R12-1 scope).
                    if (gf.Descendants<Drawing.Table>().Any())
                        RenderTable(sb, gf, themeColors, part: part);
                    break;
            }
        }

        // Also render pictures from layout/master (logos, decorative images)
        foreach (var pic in shapeTree.Elements<Picture>())
        {
            RenderPicture(sb, pic, part, themeColors);
        }
    }

    private void RenderInheritedShape(StringBuilder sb, Shape shape, OpenXmlPart part,
        HashSet<string> skipIndices, Dictionary<string, string> themeColors, int slideNum = 1)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        bool suppressText = false;
        if (ph != null)
        {
            // Slide already supplies this slot — slide content wins.
            if (ph.Index?.HasValue == true && skipIndices.Contains($"idx:{ph.Index.Value}"))
                return;
            if (ph.Type?.HasValue == true && skipIndices.Contains($"type:{ph.Type.InnerText}"))
                return;

            // ECMA-376 default: absent type == obj. Without this, a body
            // placeholder authored without an explicit type sneaks past
            // every type-based check.
            var type = ph.Type?.HasValue == true ? ph.Type.Value : PlaceholderValues.Object;
            suppressText = !IsLayoutSuppliedTextPlaceholder(type);
        }

        // Skip shapes with no visual content. When text is suppressed, treat
        // it as empty: a content placeholder with only prompt text and no
        // fill/outline isn't worth an empty box on the slide.
        var text = suppressText ? "" : GetShapeText(shape);
        var spPr = shape.ShapeProperties;
        var hasFill = spPr?.GetFirstChild<Drawing.SolidFill>() != null
            || spPr?.GetFirstChild<Drawing.GradientFill>() != null
            || spPr?.GetFirstChild<Drawing.BlipFill>() != null
            || spPr?.GetFirstChild<Drawing.PatternFill>() != null;
        // A visible outline needs a fill (solid OR gradient) — a width-only <a:ln w="X"/>
        // with no fill child renders NOTHING in PowerPoint (verified), so it must NOT
        // count as a line here. Previously only SolidFill was checked, so a layout/master
        // decoration whose only outline was a GRADIENT was silently dropped while
        // RenderShape/ParseOutline would have drawn it.
        var ln = spPr?.GetFirstChild<Drawing.Outline>();
        var hasLine = ln != null && ln.GetFirstChild<Drawing.NoFill>() == null
            && (ln.GetFirstChild<Drawing.SolidFill>() != null
                || ln.GetFirstChild<Drawing.GradientFill>() != null);

        // Style-matrix fill/line (<p:style>/<a:fillRef>/<a:lnRef>) also make the shape
        // visible — RenderShape falls back to GetStyleFillRefCss/GetStyleLineRefCss when
        // spPr carries no fill/outline. A layout/master decoration styled only via the
        // theme shape gallery (fillRef accent, empty spPr) was dropped by this guard
        // while RenderShape would have drawn the themed fill. Mirror that fallback here.
        if (!hasFill && !string.IsNullOrEmpty(GetStyleFillRefCss(shape.ShapeStyle, part, themeColors)))
            hasFill = true;
        if (!hasLine && !string.IsNullOrEmpty(GetStyleLineRefCss(shape.ShapeStyle, part, themeColors)))
            hasLine = true;

        if (string.IsNullOrWhiteSpace(text) && !hasFill && !hasLine)
            return;

        RenderShape(sb, shape, part, themeColors, suppressText: suppressText, slideNumber: slideNum);
    }

    private static bool IsLayoutSuppliedTextPlaceholder(PlaceholderValues type) =>
        type == PlaceholderValues.DateAndTime
        || type == PlaceholderValues.Footer
        || type == PlaceholderValues.Header
        || type == PlaceholderValues.SlideNumber;

}
