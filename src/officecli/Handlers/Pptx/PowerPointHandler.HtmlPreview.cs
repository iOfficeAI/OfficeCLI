// Copyright 2025 OfficeCli (officecli.ai)
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
    public string ViewAsHtml(int? startSlide = null, int? endSlide = null)
    {
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
        // themeFontLang; per-run lang is the closest signal). Emit dir="rtl"
        // when any shape carries <a:bodyPr rtlCol="1"/> or any paragraph
        // <a:pPr rtl="1"/>, so browsers activate BiDi layout document-wide.
        string presLang = "en";
        bool presHasRtl = false;
        foreach (var sp in slideParts)
        {
            var slide = sp.Slide;
            if (slide == null) continue;
            if (presLang == "en")
            {
                var firstRunLang = slide.Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>()
                    .Select(rp => rp.Language?.Value)
                    .FirstOrDefault(l => !string.IsNullOrEmpty(l));
                if (!string.IsNullOrEmpty(firstRunLang)) presLang = firstRunLang!;
            }
            if (!presHasRtl)
            {
                if (slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>()
                        .Any(p => p.ParagraphProperties?.RightToLeft?.Value == true))
                {
                    presHasRtl = true;
                }
                else
                {
                    foreach (var bp in slide.Descendants<DocumentFormat.OpenXml.Drawing.BodyProperties>())
                    {
                        foreach (var attr in bp.GetAttributes())
                        {
                            if (attr.LocalName == "rtlCol"
                                && (attr.Value == "1" || string.Equals(attr.Value, "true", StringComparison.OrdinalIgnoreCase)))
                            {
                                presHasRtl = true; break;
                            }
                        }
                        if (presHasRtl) break;
                    }
                }
            }
            if (presLang != "en" && presHasRtl) break;
        }
        var presDirAttr = presHasRtl ? " dir=\"rtl\"" : "";
        sb.AppendLine($"<html lang=\"{HtmlEncode(presLang)}\"{presDirAttr}>");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        // KaTeX for math rendering — only include when any slide actually has formulas.
        // media=print + onload swap makes the CSS non-blocking so it can never stall first paint.
        bool hasMathFormulas = slideParts.Any(sp => sp.Slide?.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() == true);
        if (hasMathFormulas)
        {
            sb.AppendLine("<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css\" media=\"print\" onload=\"this.media='all'\" onerror=\"this.remove()\">");
            sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js\" onerror=\"document.querySelectorAll('.katex-formula').forEach(function(el){el.textContent=el.dataset.formula;el.style.fontFamily='monospace';el.style.color='#666'})\"></script>");
        }
        // Three.js for 3D model rendering (graceful degradation: shows placeholder when offline)
        sb.AppendLine(@"<script type=""importmap"">{""imports"":{""three"":""https://cdn.jsdelivr.net/npm/three@0.170.0/build/three.module.js"",""three/addons/"":""https://cdn.jsdelivr.net/npm/three@0.170.0/examples/jsm/""}}</script>");
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateCss(slideWidthPt, slideHeightPt));
        sb.AppendLine("</style>");
        // Auto-hide sidebar in headless/automated browsers (screenshot, Playwright, etc.)
        sb.AppendLine("<script>if(navigator.webdriver||/HeadlessChrome/.test(navigator.userAgent))document.documentElement.classList.add('headless')</script>");
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

            sb.AppendLine($"<div class=\"slide-container\" data-slide=\"{slideNum}\">");
            sb.AppendLine($"  <div class=\"slide-label\">Slide {slideNum}</div>");
            sb.AppendLine("  <div class=\"slide-wrapper\">");
            sb.Append($"    <div class=\"slide\"");

            // Slide background + inherited text defaults from master/layout/theme
            var slideStyles = new List<string>();
            var bgStyle = GetSlideBackgroundCss(slidePart, themeColors);
            if (!string.IsNullOrEmpty(bgStyle))
                slideStyles.Add(bgStyle);
            var textDefaults = GetTextDefaults(slidePart, themeColors);
            if (!string.IsNullOrEmpty(textDefaults))
                slideStyles.Add(textDefaults);
            if (slideStyles.Count > 0)
                sb.Append($" style=\"{string.Join("", slideStyles)}\"");
            sb.AppendLine(">");

            // Render slide elements + inherited layout placeholders
            RenderLayoutPlaceholders(sb, slidePart, themeColors);
            RenderSlideElements(sb, slidePart, slideNum, slideWidthEmu, slideHeightEmu, themeColors);

            sb.AppendLine("    </div>");
            sb.AppendLine("  </div>");
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
            // Lazy-load on first demand — covers watch mode where the initial
            // doc had no formulas (KaTeX tags omitted from head), then a
            // formula arrived via SSE patch.
            if (!window._katexLoading) {
                window._katexLoading = true;
                var link = document.createElement('link');
                link.rel = 'stylesheet';
                link.href = 'https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css';
                link.onerror = function() { this.remove(); };
                document.head.appendChild(link);
                var script = document.createElement('script');
                script.src = 'https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js';
                script.onload = renderKatex;
                script.onerror = fallbackKatex;
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
})();");
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

        var sb = new StringBuilder();
        sb.AppendLine($"<div class=\"slide-container\" data-slide=\"{slideNum}\">");
        sb.AppendLine($"  <div class=\"slide-label\">Slide {slideNum}</div>");
        sb.AppendLine("  <div class=\"slide-wrapper\">");
        sb.Append($"    <div class=\"slide\"");

        var slideStyles = new List<string>();
        var bgStyle = GetSlideBackgroundCss(slidePart, themeColors);
        if (!string.IsNullOrEmpty(bgStyle))
            slideStyles.Add(bgStyle);
        var textDefaults = GetTextDefaults(slidePart, themeColors);
        if (!string.IsNullOrEmpty(textDefaults))
            slideStyles.Add(textDefaults);
        if (slideStyles.Count > 0)
            sb.Append($" style=\"{string.Join("", slideStyles)}\"");
        sb.AppendLine(">");

        RenderLayoutPlaceholders(sb, slidePart, themeColors);
        RenderSlideElements(sb, slidePart, slideNum, slideWidthEmu, slideHeightEmu, themeColors);

        sb.AppendLine("    </div>");
        sb.AppendLine("  </div>");
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
        var bgPr = slide.CommonSlideData?.Background?.BackgroundProperties;
        if (bgPr == null)
        {
            // Check slide layout and master for inherited background
            var layoutBg = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Background?.BackgroundProperties;
            var masterBg = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background?.BackgroundProperties;
            bgPr = layoutBg ?? masterBg;
        }
        if (bgPr == null) return "";

        return BackgroundPropertiesToCss(bgPr, slidePart, themeColors);
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
                return $"background:url('{dataUri}') center/cover no-repeat;";
        }

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
        int shapeIdx = 0, picIdx = 0, tableIdx = 0, chartIdx = 0, cxnIdx = 0, groupIdx = 0;
        string PathFor(string typeName, OpenXmlElement el, int positional)
            => $"/slide[{slideNum}]/{BuildElementPathSegment(typeName, el, positional)}";

        // Collect all content elements in z-order (as they appear in XML)
        foreach (var element in shapeTree.ChildElements)
        {
            switch (element)
            {
                case Shape shape:
                    shapeIdx++;
                    RenderShape(sb, shape, slidePart, themeColors, dataPath: PathFor("shape", shape, shapeIdx));
                    break;
                case Picture pic:
                    picIdx++;
                    RenderPicture(sb, pic, slidePart, themeColors, dataPath: PathFor("picture", pic, picIdx));
                    break;
                case GraphicFrame gf:
                    if (gf.Descendants<Drawing.Table>().Any())
                    {
                        tableIdx++;
                        RenderTable(sb, gf, themeColors, dataPath: PathFor("table", gf, tableIdx));
                    }
                    else if (gf.Descendants().Any(e => e.LocalName == "chart" && e.NamespaceUri.Contains("chart")))
                    {
                        chartIdx++;
                        RenderChart(sb, gf, slidePart, themeColors, dataPath: PathFor("chart", gf, chartIdx));
                    }
                    break;
                case ConnectionShape cxn:
                    cxnIdx++;
                    RenderConnector(sb, cxn, themeColors, dataPath: PathFor("connector", cxn, cxnIdx));
                    break;
                case GroupShape grp:
                    groupIdx++;
                    RenderGroup(sb, grp, slidePart, themeColors, dataPath: PathFor("group", grp, groupIdx));
                    break;
                default:
                    // mc:AlternateContent — render 3D models, zoom, etc.
                    if (element.LocalName == "AlternateContent")
                        RenderAlternateContent(sb, element, slidePart, themeColors);
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
    private void RenderLayoutPlaceholders(StringBuilder sb, SlidePart slidePart, Dictionary<string, string> themeColors)
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
            RenderInheritedShapes(sb, layoutPart.SlideLayout?.CommonSlideData?.ShapeTree, layoutPart, slidePlaceholders, themeColors);

        // Render shapes from SlideMaster (lower priority, only if not in layout)
        var masterPart = layoutPart?.SlideMasterPart;
        if (masterPart != null)
            RenderInheritedShapes(sb, masterPart.SlideMaster?.CommonSlideData?.ShapeTree, masterPart, slidePlaceholders, themeColors);
    }

    // RenderInheritedShapes — render the layout/master shapes that the slide
    // doesn't override. Two rules borrowed from Apache POI:
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
    //
    // Compare: POI's SlideShowExtractor.java:179-183 ("Ignoring boiler plate
    // (placeholder) text on slide master") and XSLFShape.java:369-370 (the
    // explicit `if (!ph.isSetType()) return INT_BODY;` default).
    private void RenderInheritedShapes(StringBuilder sb, ShapeTree? shapeTree, OpenXmlPart part,
        HashSet<string> skipIndices, Dictionary<string, string> themeColors)
    {
        if (shapeTree == null) return;

        foreach (var element in shapeTree.ChildElements)
        {
            if (element is not Shape shape) continue;

            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();

            bool suppressText = false;
            if (ph != null)
            {
                // Slide already supplies this slot — slide content wins.
                if (ph.Index?.HasValue == true && skipIndices.Contains($"idx:{ph.Index.Value}"))
                    continue;
                if (ph.Type?.HasValue == true && skipIndices.Contains($"type:{ph.Type.InnerText}"))
                    continue;

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
            var hasFill = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>() != null
                || shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null
                || shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>() != null;
            var hasLine = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>()?.GetFirstChild<Drawing.SolidFill>() != null;

            if (string.IsNullOrWhiteSpace(text) && !hasFill && !hasLine)
                continue;

            RenderShape(sb, shape, part, themeColors, suppressText: suppressText);
        }

        // Also render pictures from layout/master (logos, decorative images)
        foreach (var pic in shapeTree.Elements<Picture>())
        {
            RenderPicture(sb, pic, part, themeColors);
        }
    }

    private static bool IsLayoutSuppliedTextPlaceholder(PlaceholderValues type) =>
        type == PlaceholderValues.DateAndTime
        || type == PlaceholderValues.Footer
        || type == PlaceholderValues.Header
        || type == PlaceholderValues.SlideNumber;

}
