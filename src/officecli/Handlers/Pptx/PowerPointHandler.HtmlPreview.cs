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
    /// <summary>
    /// Generate a self-contained HTML file that previews all slides.
    /// Each slide is rendered as an absolutely-positioned div with CSS styling.
    /// Images are embedded as base64 data URIs.
    /// </summary>
    public string ViewAsHtml(int? startSlide = null, int? endSlide = null)
    {
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        // Get slide dimensions
        var (slideWidthEmu, slideHeightEmu) = GetSlideSize();
        double slideWidthCm = slideWidthEmu / 360000.0;
        double slideHeightCm = slideHeightEmu / 360000.0;

        // Resolve theme colors once for the whole presentation
        var themeColors = ResolveThemeColorMap();

        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        sb.AppendLine("<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css\">");
        sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js\"></script>");
        // Three.js for 3D model rendering (importmap for ES module support)
        sb.AppendLine(@"<script type=""importmap"">{""imports"":{""three"":""https://cdn.jsdelivr.net/npm/three@0.170.0/build/three.module.js"",""three/addons/"":""https://cdn.jsdelivr.net/npm/three@0.170.0/examples/jsm/""}}</script>");
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateCss(slideWidthCm, slideHeightCm));
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
    function renderKatex() {
        if (typeof katex === 'undefined') { setTimeout(renderKatex, 100); return; }
        document.querySelectorAll('.katex-formula:not(.katex-rendered)').forEach(function(el) {
            try {
                katex.render(el.dataset.formula, el, { throwOnError: false, displayMode: true });
                el.classList.add('katex-rendered');
            } catch(e) { el.textContent = el.dataset.formula; }
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

    private static string GenerateCss(double slideWidthCm, double slideHeightCm)
    {
        var aspect = slideWidthCm / slideHeightCm;
        // Dynamic CSS variables + static CSS from embedded resource
        var dynamicVars = $":root{{--slide-design-w:{slideWidthCm:0.###}cm;--slide-design-h:{slideHeightCm:0.###}cm;--slide-aspect:{aspect:0.####};}}\n";
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

        // Build font-family with fallbacks including CJK fonts
        var fonts = new List<string>();
        if (!string.IsNullOrEmpty(minorLatin)) fonts.Add($"'{CssSanitize(minorLatin)}'");
        if (!string.IsNullOrEmpty(minorEa)) fonts.Add($"'{CssSanitize(minorEa)}'");
        // CJK fallback chain: macOS → Windows → Linux
        fonts.AddRange(new[] { "'PingFang SC'", "'Microsoft YaHei'", "'Noto Sans CJK SC'", "'Hiragino Sans GB'", "sans-serif" });
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

        // Collect all content elements in z-order (as they appear in XML)
        foreach (var element in shapeTree.ChildElements)
        {
            switch (element)
            {
                case Shape shape:
                    RenderShape(sb, shape, slidePart, themeColors);
                    break;
                case Picture pic:
                    RenderPicture(sb, pic, slidePart, themeColors);
                    break;
                case GraphicFrame gf:
                    if (gf.Descendants<Drawing.Table>().Any())
                        RenderTable(sb, gf, themeColors);
                    else if (gf.Descendants().Any(e => e.LocalName == "chart" && e.NamespaceUri.Contains("chart")))
                        RenderChart(sb, gf, slidePart, themeColors);
                    break;
                case ConnectionShape cxn:
                    RenderConnector(sb, cxn, themeColors);
                    break;
                case GroupShape grp:
                    RenderGroup(sb, grp, slidePart, themeColors);
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

    private void RenderInheritedShapes(StringBuilder sb, ShapeTree? shapeTree, OpenXmlPart part,
        HashSet<string> skipIndices, Dictionary<string, string> themeColors)
    {
        if (shapeTree == null) return;

        foreach (var element in shapeTree.ChildElements)
        {
            if (element is not Shape shape) continue;

            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();

            // Skip title/body content placeholders (these are structural, not decorative)
            if (ph?.Type?.HasValue == true)
            {
                var t = ph.Type.Value;
                if (t == PlaceholderValues.Title || t == PlaceholderValues.CenteredTitle ||
                    t == PlaceholderValues.SubTitle || t == PlaceholderValues.Body ||
                    t == PlaceholderValues.Object)
                    continue;

                // Skip if slide already has this placeholder type
                if (skipIndices.Contains($"type:{ph.Type.InnerText}")) continue;
            }

            // Skip if slide already has a shape with this placeholder index
            if (ph?.Index?.HasValue == true && skipIndices.Contains($"idx:{ph.Index.Value}"))
                continue;

            // Skip shapes with no visual content (empty text, no fill, no picture)
            var text = GetShapeText(shape);
            var hasFill = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>() != null
                || shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null
                || shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>() != null;
            var hasLine = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>()?.GetFirstChild<Drawing.SolidFill>() != null;

            if (string.IsNullOrWhiteSpace(text) && !hasFill && !hasLine)
                continue;

            // Render this inherited shape
            RenderShape(sb, shape, part, themeColors);
        }

        // Also render pictures from layout/master (logos, decorative images)
        foreach (var pic in shapeTree.Elements<Picture>())
        {
            if (part is SlidePart sp)
                RenderPicture(sb, pic, sp, themeColors);
        }
    }

}
