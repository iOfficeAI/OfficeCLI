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

        // Get slide dimensions (default: standard 16:9 = 33.867cm x 19.05cm)
        var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
        long slideWidthEmu = sldSz?.Cx?.Value ?? 12192000;
        long slideHeightEmu = sldSz?.Cy?.Value ?? 6858000;
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
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateCss(slideWidthCm, slideHeightCm));
        sb.AppendLine("</style>");
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

        var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
        long slideWidthEmu = sldSz?.Cx?.Value ?? 12192000;
        long slideHeightEmu = sldSz?.Cy?.Value ?? 6858000;
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

    // ==================== Shape Rendering ====================

    /// <summary>
    /// Render a shape element to HTML. When called from a group, pass overridePos
    /// with the adjusted coordinates — the original element is NEVER modified.
    /// </summary>
    private static void RenderShape(StringBuilder sb, Shape shape, OpenXmlPart part,
        Dictionary<string, string> themeColors, (long x, long y, long cx, long cy)? overridePos = null)
    {
        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm?.Offset == null || xfrm?.Extents == null) return;

        var x = overridePos?.x ?? xfrm.Offset.X?.Value ?? 0;
        var y = overridePos?.y ?? xfrm.Offset.Y?.Value ?? 0;
        var cx = overridePos?.cx ?? xfrm.Extents.Cx?.Value ?? 0;
        var cy = overridePos?.cy ?? xfrm.Extents.Cy?.Value ?? 0;

        var styles = new List<string>
        {
            $"left:{EmuToCm(x)}cm",
            $"top:{EmuToCm(y)}cm",
            $"width:{EmuToCm(cx)}cm",
            $"height:{EmuToCm(cy)}cm"
        };

        // Fill
        var fillCss = GetShapeFillCss(shape.ShapeProperties, part, themeColors);
        if (!string.IsNullOrEmpty(fillCss))
            styles.Add(fillCss);

        // Border/outline
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        if (outline != null)
        {
            var borderCss = OutlineToCss(outline, themeColors);
            if (!string.IsNullOrEmpty(borderCss))
                styles.Add(borderCss);
        }

        // Build transform chain (must be combined into one transform property)
        var transforms = new List<string>();

        // 2D rotation
        if (xfrm.Rotation != null && xfrm.Rotation.Value != 0)
        {
            var deg = xfrm.Rotation.Value / 60000.0;
            transforms.Add($"rotate({deg:0.##}deg)");
        }

        // Flip
        if (xfrm.HorizontalFlip?.Value == true && xfrm.VerticalFlip?.Value == true)
            transforms.Add("scale(-1,-1)");
        else if (xfrm.HorizontalFlip?.Value == true)
            transforms.Add("scaleX(-1)");
        else if (xfrm.VerticalFlip?.Value == true)
            transforms.Add("scaleY(-1)");

        // 3D rotation (scene3d camera rotation) → CSS perspective transform
        var scene3d = shape.ShapeProperties?.GetFirstChild<Drawing.Scene3DType>();
        var cam = scene3d?.Camera;
        var rot3d = cam?.Rotation;
        if (rot3d != null)
        {
            var rx = (rot3d.Latitude?.Value ?? 0) / 60000.0;
            var ry = (rot3d.Longitude?.Value ?? 0) / 60000.0;
            var rz = (rot3d.Revolution?.Value ?? 0) / 60000.0;
            if (rx != 0 || ry != 0 || rz != 0)
            {
                styles.Add("perspective:800px");
                if (rx != 0) transforms.Add($"rotateX({rx:0.##}deg)");
                if (ry != 0) transforms.Add($"rotateY({ry:0.##}deg)");
                if (rz != 0) transforms.Add($"rotateZ({rz:0.##}deg)");
            }
        }

        if (transforms.Count > 0)
            styles.Add($"transform:{string.Join(" ", transforms)}");

        // Geometry: preset or custom — track clip-path separately to avoid clipping text
        string clipPathCss = "";
        string borderRadiusCss = "";
        var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
        {
            var geomCss = PresetGeometryToCss(presetGeom.Preset!.InnerText!);
            if (!string.IsNullOrEmpty(geomCss))
            {
                if (geomCss.StartsWith("clip-path:"))
                    clipPathCss = geomCss;
                else
                {
                    styles.Add(geomCss);
                    borderRadiusCss = geomCss;
                }
            }
        }
        else
        {
            // Custom geometry (custGeom) → SVG clip-path
            var custGeom = shape.ShapeProperties?.GetFirstChild<Drawing.CustomGeometry>();
            if (custGeom != null)
            {
                var clipPath = CustomGeometryToClipPath(custGeom);
                if (!string.IsNullOrEmpty(clipPath))
                    clipPathCss = clipPath;
            }
        }

        // Shadow
        var effectList = shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        var shadowCss = EffectListToShadowCss(effectList, themeColors);
        if (!string.IsNullOrEmpty(shadowCss))
            styles.Add(shadowCss);

        // Soft edge → fade out at edges using CSS mask-image
        // Unlike filter:blur() which blurs the entire element,
        // mask-image with edge gradients only affects the border region.
        var softEdge = effectList?.GetFirstChild<Drawing.SoftEdge>()
            ?? shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>()?.GetFirstChild<Drawing.SoftEdge>();
        if (softEdge == null)
        {
            softEdge = shape.TextBody?.Descendants<Drawing.RunProperties>()
                .Select(rp => rp.GetFirstChild<Drawing.EffectList>()?.GetFirstChild<Drawing.SoftEdge>())
                .FirstOrDefault(se => se != null);
        }
        if (softEdge?.Radius?.HasValue == true)
        {
            var edgePx = Math.Max(2, softEdge.Radius.Value / 12700.0 * 0.8);
            // Use linear-gradient masks on all 4 edges to create edge fade-out
            styles.Add($"-webkit-mask-image:linear-gradient(to right,transparent 0,black {edgePx:0.#}px,black calc(100% - {edgePx:0.#}px),transparent 100%)," +
                       $"linear-gradient(to bottom,transparent 0,black {edgePx:0.#}px,black calc(100% - {edgePx:0.#}px),transparent 100%)");
            styles.Add("-webkit-mask-composite:source-in;mask-composite:intersect");
        }

        // Bevel → approximate with inset box-shadow for a subtle 3D appearance
        var sp3d = shape.ShapeProperties?.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d?.BevelTop != null)
        {
            var bevelW = sp3d.BevelTop.Width?.HasValue == true ? sp3d.BevelTop.Width.Value / 12700.0 : 4;
            var bW = Math.Max(1, bevelW * 0.5);
            styles.Add($"box-shadow:inset {bW:0.#}px {bW:0.#}px {bW * 1.5:0.#}px rgba(255,255,255,0.25),inset -{bW:0.#}px -{bW:0.#}px {bW * 1.5:0.#}px rgba(0,0,0,0.15)");
        }

        // Note: fill opacity (alpha) is already baked into rgba() by ResolveFillColor.
        // Do NOT add a separate CSS opacity here — it would double-apply.

        // Text margins
        var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        var lIns = bodyPr?.LeftInset?.Value ?? 91440;
        var tIns = bodyPr?.TopInset?.Value ?? 45720;
        var rIns = bodyPr?.RightInset?.Value ?? 91440;
        var bIns = bodyPr?.BottomInset?.Value ?? 45720;
        styles.Add($"padding:{EmuToCm(tIns)}cm {EmuToCm(rIns)}cm {EmuToCm(bIns)}cm {EmuToCm(lIns)}cm");

        // Vertical alignment class
        var valign = "top";
        if (bodyPr?.Anchor?.HasValue == true)
        {
            valign = bodyPr.Anchor.InnerText switch
            {
                "ctr" => "center",
                "b" => "bottom",
                _ => "top"
            };
        }

        // Add has-fill class to clip overflow when shape has a visible background
        var hasFillBg = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>() != null
            || shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null
            || shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>() != null;
        var shapeClass = hasFillBg ? "shape has-fill" : "shape";

        if (!string.IsNullOrEmpty(clipPathCss))
        {
            // For clip-path shapes: move fill to a clipped background layer, keep text unclipped
            // Extract fill-related styles for the clipped background layer
            var fillStyles = new List<string>();
            var outerStyles = new List<string>();
            foreach (var s in styles)
            {
                if (s.StartsWith("background:") || s.StartsWith("background-image:"))
                    fillStyles.Add(s);
                else
                    outerStyles.Add(s);
            }
            sb.Append($"    <div class=\"{shapeClass}\" style=\"{string.Join(";", outerStyles)}\">");
            if (fillStyles.Count > 0)
                sb.Append($"<div style=\"position:absolute;inset:0;{clipPathCss};{string.Join(";", fillStyles)}\"></div>");
        }
        else
        {
            sb.Append($"    <div class=\"{shapeClass}\" style=\"{string.Join(";", styles)}\">");
        }

        // Text content
        if (shape.TextBody != null)
        {
            // Counter-flip text so it remains readable when shape is flipped
            var flipStyle = "";
            var isFlipH = xfrm?.HorizontalFlip?.Value == true;
            var isFlipV = xfrm?.VerticalFlip?.Value == true;
            if (isFlipH && isFlipV)
                flipStyle = "transform:scale(-1,-1);";
            else if (isFlipH)
                flipStyle = "transform:scaleX(-1);";
            else if (isFlipV)
                flipStyle = "transform:scaleY(-1);";

            var textStyle = !string.IsNullOrEmpty(flipStyle) || !string.IsNullOrEmpty(clipPathCss)
                ? $" style=\"{flipStyle}{(string.IsNullOrEmpty(clipPathCss) ? "" : "position:relative;")}\""
                : "";
            sb.Append($"<div class=\"shape-text valign-{valign}\"{textStyle}>");
            RenderTextBody(sb, shape.TextBody, themeColors);
            sb.Append("</div>");
        }

        sb.AppendLine("</div>");
    }

    // ==================== Text Rendering ====================

    private static void RenderTextBody(StringBuilder sb, OpenXmlElement textBody, Dictionary<string, string> themeColors)
    {
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            var paraStyles = new List<string>();

            var pProps = para.ParagraphProperties;
            if (pProps?.Alignment?.HasValue == true)
            {
                var align = pProps.Alignment.InnerText switch
                {
                    "l" => "left",
                    "ctr" => "center",
                    "r" => "right",
                    "just" => "justify",
                    _ => "left"
                };
                paraStyles.Add($"text-align:{align}");
            }

            // Paragraph spacing
            var sbPts = pProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sbPts.HasValue) paraStyles.Add($"margin-top:{sbPts.Value / 100.0:0.##}pt");
            var saPts = pProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (saPts.HasValue) paraStyles.Add($"margin-bottom:{saPts.Value / 100.0:0.##}pt");

            // Line spacing
            var lsPct = pProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (lsPct.HasValue) paraStyles.Add($"line-height:{lsPct.Value / 100000.0:0.##}");
            var lsPts = pProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (lsPts.HasValue) paraStyles.Add($"line-height:{lsPts.Value / 100.0:0.##}pt");

            // Indent
            if (pProps?.Indent?.HasValue == true)
                paraStyles.Add($"text-indent:{EmuToCm(pProps.Indent.Value)}cm");
            if (pProps?.LeftMargin?.HasValue == true)
                paraStyles.Add($"margin-left:{EmuToCm(pProps.LeftMargin.Value)}cm");

            // Bullet
            var bulletChar = pProps?.GetFirstChild<Drawing.CharacterBullet>()?.Char?.Value;
            var bulletAuto = pProps?.GetFirstChild<Drawing.AutoNumberedBullet>();
            var hasBullet = bulletChar != null || bulletAuto != null;

            sb.Append($"<div class=\"para\" style=\"{string.Join(";", paraStyles)}\">");

            if (hasBullet)
            {
                var bullet = bulletChar ?? "\u2022";
                sb.Append($"<span class=\"bullet\">{HtmlEncode(bullet)} </span>");
            }

            // Check for OfficeMath (a14:m inside mc:AlternateContent) in paragraph XML
            var paraXml = para.OuterXml;
            if (paraXml.Contains("oMath"))
            {
                // AlternateContent is opaque to Descendants() — parse from XML
                var mathMatch = System.Text.RegularExpressions.Regex.Match(paraXml,
                    @"<m:oMathPara[^>]*>.*?</m:oMathPara>|<m:oMath[^>]*>.*?</m:oMath>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (mathMatch.Success)
                {
                    var mathXml = $"<wrapper xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">{mathMatch.Value}</wrapper>";
                    try
                    {
                        var wrapper = new OpenXmlUnknownElement("wrapper");
                        wrapper.InnerXml = mathMatch.Value;
                        var oMath = wrapper.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                        if (oMath != null)
                        {
                            var latex = FormulaParser.ToLatex(oMath);
                            sb.Append($"<span class=\"katex-formula\" data-formula=\"{HtmlEncode(latex)}\"></span>");
                        }
                    }
                    catch { }
                }
            }

            var hasMath = paraXml.Contains("oMath");
            var runs = para.Elements<Drawing.Run>().ToList();
            if (runs.Count == 0 && !hasMath)
            {
                // Empty paragraph (line break)
                sb.Append("&nbsp;");
            }
            else
            {
                foreach (var run in runs)
                {
                    RenderRun(sb, run, themeColors);
                }
            }

            // Line breaks within paragraph
            foreach (var br in para.Elements<Drawing.Break>())
                sb.Append("<br>");

            sb.AppendLine("</div>");
        }
    }

    private static void RenderRun(StringBuilder sb, Drawing.Run run, Dictionary<string, string> themeColors)
    {
        var text = run.Text?.Text ?? "";
        if (string.IsNullOrEmpty(text)) return;

        var styles = new List<string>();
        var rp = run.RunProperties;

        if (rp != null)
        {
            // Font
            var font = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
            if (font != null && !font.StartsWith("+", StringComparison.Ordinal))
                styles.Add($"font-family:'{CssSanitize(font)}'");

            // Size
            if (rp.FontSize?.HasValue == true)
                styles.Add($"font-size:{rp.FontSize.Value / 100.0:0.##}pt");

            // Bold
            if (rp.Bold?.Value == true)
                styles.Add("font-weight:bold");

            // Italic
            if (rp.Italic?.Value == true)
                styles.Add("font-style:italic");

            // Underline
            if (rp.Underline?.HasValue == true && rp.Underline.Value != Drawing.TextUnderlineValues.None)
                styles.Add("text-decoration:underline");

            // Strikethrough
            if (rp.Strike?.HasValue == true && rp.Strike.Value != Drawing.TextStrikeValues.NoStrike)
                styles.Add("text-decoration:line-through");

            // Color
            var solidFill = rp.GetFirstChild<Drawing.SolidFill>();
            var color = ResolveFillColor(solidFill, themeColors);
            if (color != null)
                styles.Add($"color:{color}");

            // Character spacing
            if (rp.Spacing?.HasValue == true)
                styles.Add($"letter-spacing:{rp.Spacing.Value / 100.0:0.##}pt");

            // Superscript/subscript
            if (rp.Baseline?.HasValue == true && rp.Baseline.Value != 0)
            {
                if (rp.Baseline.Value > 0)
                    styles.Add("vertical-align:super;font-size:smaller");
                else
                    styles.Add("vertical-align:sub;font-size:smaller");
            }
        }

        // Hyperlink
        var hlinkClick = run.Parent?.Elements<Drawing.Run>()
            .Where(r => r == run)
            .Select(_ => run.Parent)
            .FirstOrDefault()
            ?.GetFirstChild<Drawing.HyperlinkOnClick>();
        // Actually check run's parent paragraph for hyperlinks on this run
        // Not critical for preview, skip for simplicity

        if (styles.Count > 0)
            sb.Append($"<span style=\"{string.Join(";", styles)}\">{HtmlEncode(text)}</span>");
        else
            sb.Append(HtmlEncode(text));
    }

    // ==================== Picture Rendering ====================

    /// <summary>
    /// Render a picture element to HTML. When called from a group, pass overridePos
    /// with the adjusted coordinates — the original element is NEVER modified.
    /// </summary>
    private static void RenderPicture(StringBuilder sb, Picture pic, SlidePart slidePart,
        Dictionary<string, string> themeColors, (long x, long y, long cx, long cy)? overridePos = null)
    {
        var xfrm = pic.ShapeProperties?.Transform2D;
        if (xfrm?.Offset == null || xfrm?.Extents == null) return;

        var x = overridePos?.x ?? xfrm.Offset.X?.Value ?? 0;
        var y = overridePos?.y ?? xfrm.Offset.Y?.Value ?? 0;
        var cx = overridePos?.cx ?? xfrm.Extents.Cx?.Value ?? 0;
        var cy = overridePos?.cy ?? xfrm.Extents.Cy?.Value ?? 0;

        var styles = new List<string>
        {
            $"left:{EmuToCm(x)}cm",
            $"top:{EmuToCm(y)}cm",
            $"width:{EmuToCm(cx)}cm",
            $"height:{EmuToCm(cy)}cm"
        };

        // Rotation
        if (xfrm.Rotation != null && xfrm.Rotation.Value != 0)
            styles.Add($"transform:rotate({xfrm.Rotation.Value / 60000.0:0.##}deg)");

        // Border
        var outline = pic.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        if (outline != null)
        {
            var borderCss = OutlineToCss(outline, themeColors);
            if (!string.IsNullOrEmpty(borderCss))
                styles.Add(borderCss);
        }

        // Shadow
        var effectList = pic.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        var shadowCss = EffectListToShadowCss(effectList, themeColors);
        if (!string.IsNullOrEmpty(shadowCss))
            styles.Add(shadowCss);

        // Geometry (rounded corners)
        var presetGeom = pic.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
        {
            var geomCss = PresetGeometryToCss(presetGeom.Preset!.InnerText!);
            if (!string.IsNullOrEmpty(geomCss))
                styles.Add(geomCss);
        }

        sb.Append($"    <div class=\"picture\" style=\"{string.Join(";", styles)}\">");

        // Extract image data
        var blipFill = pic.BlipFill;
        var blip = blipFill?.GetFirstChild<Drawing.Blip>();
        if (blip?.Embed?.HasValue == true)
        {
            try
            {
                var imgPart = slidePart.GetPartById(blip.Embed.Value!);
                using var stream = imgPart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                var base64 = Convert.ToBase64String(ms.ToArray());
                var contentType = SanitizeContentType(imgPart.ContentType ?? "image/png");

                // Crop
                var srcRect = blipFill?.GetFirstChild<Drawing.SourceRectangle>();
                var imgStyles = new List<string>();
                if (srcRect != null)
                {
                    var cl = (srcRect.Left?.Value ?? 0) / 1000.0;
                    var ct = (srcRect.Top?.Value ?? 0) / 1000.0;
                    var cr = (srcRect.Right?.Value ?? 0) / 1000.0;
                    var cb = (srcRect.Bottom?.Value ?? 0) / 1000.0;
                    if (cl != 0 || ct != 0 || cr != 0 || cb != 0)
                    {
                        // Use clip-path for cropping
                        imgStyles.Add($"clip-path:inset({ct:0.##}% {cr:0.##}% {cb:0.##}% {cl:0.##}%)");
                    }
                }

                var imgStyle = imgStyles.Count > 0 ? $" style=\"{string.Join(";", imgStyles)}\"" : "";
                sb.Append($"<img src=\"data:{contentType};base64,{base64}\"{imgStyle} loading=\"lazy\">");
            }
            catch
            {
                // Image extraction failed - show placeholder
                sb.Append("<div style=\"width:100%;height:100%;background:#e0e0e0;display:flex;align-items:center;justify-content:center;color:#999;font-size:12px\">Image</div>");
            }
        }

        sb.AppendLine("</div>");
    }

    // ==================== Table Rendering ====================

    private static void RenderTable(StringBuilder sb, GraphicFrame gf, Dictionary<string, string> themeColors)
    {
        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
        if (table == null) return;

        var offset = gf.Transform?.Offset;
        var extents = gf.Transform?.Extents;
        if (offset == null || extents == null) return;

        var x = offset.X?.Value ?? 0;
        var y = offset.Y?.Value ?? 0;
        var cx = extents.Cx?.Value ?? 0;
        var cy = extents.Cy?.Value ?? 0;

        sb.AppendLine($"    <div class=\"table-container\" style=\"left:{EmuToCm(x)}cm;top:{EmuToCm(y)}cm;width:{EmuToCm(cx)}cm;height:{EmuToCm(cy)}cm\">");
        sb.AppendLine("      <table class=\"slide-table\">");

        // Column widths
        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
        if (gridCols != null && gridCols.Count > 0)
        {
            sb.Append("        <colgroup>");
            long totalWidth = gridCols.Sum(gc => gc.Width?.Value ?? 0);
            foreach (var gc in gridCols)
            {
                var w = gc.Width?.Value ?? 0;
                var pct = totalWidth > 0 ? (w * 100.0 / totalWidth) : (100.0 / gridCols.Count);
                sb.Append($"<col style=\"width:{pct:0.##}%\">");
            }
            sb.AppendLine("</colgroup>");
        }

        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            sb.AppendLine("        <tr>");
            foreach (var cell in row.Elements<Drawing.TableCell>())
            {
                var cellStyles = new List<string>();

                // Cell fill
                var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                var cellSolid = tcPr?.GetFirstChild<Drawing.SolidFill>();
                var cellColor = ResolveFillColor(cellSolid, themeColors);
                if (cellColor != null)
                    cellStyles.Add($"background:{cellColor}");

                var cellGrad = tcPr?.GetFirstChild<Drawing.GradientFill>();
                if (cellGrad != null)
                    cellStyles.Add($"background:{GradientToCss(cellGrad, themeColors)}");

                // Vertical alignment
                if (tcPr?.Anchor?.HasValue == true)
                {
                    var va = tcPr.Anchor.InnerText switch
                    {
                        "ctr" => "middle",
                        "b" => "bottom",
                        _ => "top"
                    };
                    cellStyles.Add($"vertical-align:{va}");
                }

                // Cell text formatting
                var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
                if (firstRun?.RunProperties != null)
                {
                    var rp = firstRun.RunProperties;
                    if (rp.FontSize?.HasValue == true)
                        cellStyles.Add($"font-size:{rp.FontSize.Value / 100.0:0.##}pt");
                    if (rp.Bold?.Value == true)
                        cellStyles.Add("font-weight:bold");
                    var fontVal = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                        ?? rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                    if (fontVal != null && !fontVal.StartsWith("+", StringComparison.Ordinal))
                        cellStyles.Add($"font-family:'{CssSanitize(fontVal)}'");
                    var runColor = ResolveFillColor(rp.GetFirstChild<Drawing.SolidFill>(), themeColors);
                    if (runColor != null)
                        cellStyles.Add($"color:{runColor}");
                }

                // Paragraph alignment
                var firstPara = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
                {
                    var align = firstPara.ParagraphProperties.Alignment.InnerText switch
                    {
                        "ctr" => "center",
                        "r" => "right",
                        "just" => "justify",
                        _ => "left"
                    };
                    cellStyles.Add($"text-align:{align}");
                }

                var cellText = cell.TextBody?.InnerText ?? "";
                var styleStr = cellStyles.Count > 0 ? $" style=\"{string.Join(";", cellStyles)}\"" : "";

                // Column/row span (GridSpan and RowSpan are on the TableCell, not TableCellProperties)
                var gridSpan = cell.GridSpan?.Value;
                var rowSpan = cell.RowSpan?.Value;
                var spanAttrs = "";
                if (gridSpan > 1) spanAttrs += $" colspan=\"{gridSpan}\"";
                if (rowSpan > 1) spanAttrs += $" rowspan=\"{rowSpan}\"";

                // Skip merged continuation cells
                if (cell.HorizontalMerge?.Value == true || cell.VerticalMerge?.Value == true)
                    continue;

                sb.AppendLine($"          <td{spanAttrs}{styleStr}>{HtmlEncode(cellText)}</td>");
            }
            sb.AppendLine("        </tr>");
        }

        sb.AppendLine("      </table>");
        sb.AppendLine("    </div>");
    }

    // ==================== Chart Rendering ====================

    private static readonly string[] ChartColors = [
        "#E74C3C", "#3498DB", "#2ECC71", "#F39C12", "#9B59B6", "#1ABC9C",
        "#E67E22", "#34495E", "#E91E63", "#00BCD4", "#8BC34A", "#FF9800"
    ];

    private void RenderChart(StringBuilder sb, GraphicFrame gf, SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        // p:xfrm contains a:off and a:ext
        var pxfrm = gf.GetFirstChild<DocumentFormat.OpenXml.Presentation.Transform>();
        var off = pxfrm?.GetFirstChild<Drawing.Offset>();
        var ext = pxfrm?.GetFirstChild<Drawing.Extents>();
        if (off == null || ext == null) return;

        var x = EmuToCm(off.X?.Value ?? 0);
        var y = EmuToCm(off.Y?.Value ?? 0);
        var w = EmuToCm(ext.Cx?.Value ?? 0);
        var h = EmuToCm(ext.Cy?.Value ?? 0);

        // Read chart data — find c:chart element with r:id
        var chartEl = gf.Descendants().FirstOrDefault(e => e.LocalName == "chart" && e.NamespaceUri.Contains("chart"));
        var rId = chartEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri.Contains("relationships")).Value;
        if (rId == null) return;

        DocumentFormat.OpenXml.Drawing.Charts.Chart? chart;
        DocumentFormat.OpenXml.Drawing.Charts.PlotArea? plotArea;
        try
        {
            var chartPart = (ChartPart)slidePart.GetPartById(rId);
            chart = chartPart.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            plotArea = chart?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>();
            if (plotArea == null) return;
        }
        catch { return; }

        var chartType = ChartHelper.DetectChartType(plotArea) ?? "bar";
        var categories = ChartHelper.ReadCategories(plotArea) ?? [];
        var seriesList = ChartHelper.ReadAllSeries(plotArea);
        if (seriesList.Count == 0) return;

        // Read series colors
        var seriesColors = new List<string>();
        var serElements = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser").ToList();
        for (int i = 0; i < seriesList.Count; i++)
        {
            var serEl = i < serElements.Count ? serElements[i] : null;
            var spPr = serEl?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>();
            var fill = spPr?.GetFirstChild<Drawing.SolidFill>();
            var rgb = fill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            seriesColors.Add(rgb != null ? $"#{rgb}" : ChartColors[i % ChartColors.Length]);
        }

        // Title
        var titleText = chart?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>()
            ?.Descendants<Drawing.Text>().FirstOrDefault()?.Text ?? "";

        sb.AppendLine($"    <div class=\"shape\" style=\"left:{x}cm;top:{y}cm;width:{w}cm;height:{h}cm;background:rgba(255,255,255,0.05)\">");

        // Title
        if (!string.IsNullOrEmpty(titleText))
            sb.AppendLine($"      <div style=\"text-align:center;font-size:11px;font-weight:bold;padding:4px;color:#ccc\">{HtmlEncode(titleText)}</div>");

        // SVG chart area — proportional to actual shape dimensions
        var widthEmu = ext.Cx?.Value ?? 3600000;
        var heightEmu = ext.Cy?.Value ?? 2520000;
        var svgW = (int)(widthEmu / 10000.0); // scale down to reasonable SVG units
        var svgH = (int)(heightEmu / 10000.0);
        var titleH = string.IsNullOrEmpty(titleText) ? 0 : 20;
        var chartSvgH = svgH - titleH;
        var margin = new { top = 10, right = 15, bottom = 25, left = 40 };
        var plotW = svgW - margin.left - margin.right;
        var plotH = chartSvgH - margin.top - margin.bottom;

        sb.AppendLine($"      <svg viewBox=\"0 0 {svgW} {chartSvgH}\" style=\"width:100%;height:calc(100% - {titleH + 4}px)\" preserveAspectRatio=\"xMidYMin meet\">");

        if (chartType.Contains("pie") || chartType.Contains("doughnut"))
        {
            var isDoughnut = chartType.Contains("doughnut");
            var holeSize = 0.0;
            if (isDoughnut)
            {
                var holeSizeEl = plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.HoleSize>().FirstOrDefault();
                holeSize = (holeSizeEl?.Val?.Value ?? 50) / 100.0; // default 50%
            }
            RenderPieChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH, holeSize);
        }
        else if (chartType.Contains("area"))
        {
            RenderAreaChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType == "combo")
        {
            RenderComboChartSvg(sb, plotArea, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType.Contains("radar"))
        {
            RenderRadarChartSvg(sb, seriesList, categories, seriesColors, svgW, chartSvgH);
        }
        else if (chartType == "bubble")
        {
            RenderBubbleChartSvg(sb, plotArea, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType == "stock")
        {
            RenderStockChartSvg(sb, plotArea, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else if (chartType.Contains("line") || chartType == "scatter")
        {
            RenderLineChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH);
        }
        else
        {
            // Bar/column (default)
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            RenderBarChartSvg(sb, seriesList, categories, seriesColors, margin.left, margin.top, plotW, plotH, isHorizontal);
        }

        sb.AppendLine("      </svg>");

        // Legend
        if (seriesList.Count > 1)
        {
            sb.Append("      <div style=\"display:flex;justify-content:center;gap:8px;font-size:8px;color:#aaa;padding:2px\">");
            for (int i = 0; i < seriesList.Count; i++)
            {
                sb.Append($"<span><span style=\"display:inline-block;width:8px;height:8px;background:{seriesColors[i]};margin-right:2px;border-radius:1px\"></span>{HtmlEncode(seriesList[i].name)}</span>");
            }
            sb.AppendLine("</div>");
        }

        sb.AppendLine("    </div>");
    }

    private static void RenderBarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool horizontal)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;

        if (horizontal)
        {
            // Horizontal bar: categories along Y axis, values along X axis
            var hLabelMargin = 50; // extra left margin for category labels
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;
            var groupH = (double)ph / Math.Max(catCount, 1);
            var barH = groupH * 0.7 / serCount;
            var gap = groupH * 0.15;

            // Axis lines
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

            // Bars
            for (int s = 0; s < serCount; s++)
            {
                for (int c = 0; c < series[s].values.Length && c < catCount; c++)
                {
                    var val = series[s].values[c];
                    var barW = (val / maxVal) * plotPw;
                    var bx = plotOx;
                    var by = oy + c * groupH + gap + s * barH;
                    sb.AppendLine($"        <rect x=\"{bx}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                }
            }

            // Category labels (left side)
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var ly = oy + c * groupH + groupH / 2;
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"#999\" font-size=\"9\" text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }

            // Value axis labels (bottom)
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var tx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#777\" font-size=\"8\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            // Vertical column: categories along X axis, values along Y axis
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.7 / serCount;
            var gap = groupW * 0.15;

            // Axis lines
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

            // Bars
            for (int s = 0; s < serCount; s++)
            {
                for (int c = 0; c < series[s].values.Length && c < catCount; c++)
                {
                    var val = series[s].values[c];
                    var barH = (val / maxVal) * ph;
                    var bx = ox + c * groupW + gap + s * barW;
                    var by = oy + ph - barH;
                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                }
            }

            // Category labels (bottom)
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }

            // Value axis labels (left side)
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var ty = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    private static void RenderLineChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / maxVal) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
                // Dots
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s]}\"/>");
                }
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    private static void RenderPieChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, double holeRatio = 0.0)
    {
        // Use first series values
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.35;
        var innerR = r * holeRatio;
        var startAngle = -Math.PI / 2;

        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : ChartColors[i % ChartColors.Length];

            if (values.Length == 1 && holeRatio <= 0)
            {
                sb.AppendLine($"        <circle cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" r=\"{r:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            else if (holeRatio > 0)
            {
                // Doughnut: arc with inner hole
                var ox1 = cx + r * Math.Cos(startAngle);
                var oy1 = cy + r * Math.Sin(startAngle);
                var ox2 = cx + r * Math.Cos(endAngle);
                var oy2 = cy + r * Math.Sin(endAngle);
                var ix1 = cx + innerR * Math.Cos(endAngle);
                var iy1 = cy + innerR * Math.Sin(endAngle);
                var ix2 = cx + innerR * Math.Cos(startAngle);
                var iy2 = cy + innerR * Math.Sin(startAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                // Outer arc clockwise, line to inner, inner arc counter-clockwise, close
                sb.AppendLine($"        <path d=\"M {ox1:0.#},{oy1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {ox2:0.#},{oy2:0.#} L {ix1:0.#},{iy1:0.#} A {innerR:0.#},{innerR:0.#} 0 {largeArc},0 {ix2:0.#},{iy2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            else
            {
                var x1 = cx + r * Math.Cos(startAngle);
                var y1 = cy + r * Math.Sin(startAngle);
                var x2 = cx + r * Math.Cos(endAngle);
                var y2 = cy + r * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {cx:0.#},{cy:0.#} L {x1:0.#},{y1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }

            // Label — position between inner and outer radius for doughnut
            var midAngle = startAngle + sliceAngle / 2;
            var labelR = holeRatio > 0 ? (r + innerR) / 2 : r * 0.6;
            var lx = cx + labelR * Math.Cos(midAngle);
            var ly = cy + labelR * Math.Sin(midAngle);
            var label = i < categories.Length ? categories[i] : "";
            if (!string.IsNullOrEmpty(label))
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"white\" font-size=\"9\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");

            startAngle = endAngle;
        }
    }

    private static void RenderAreaChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        // Render in reverse order so first series is on top
        for (int s = series.Count - 1; s >= 0; s--)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / maxVal) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                // Close polygon to baseline
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastX = ox + (catCount > 1 ? (double)pw * (series[s].values.Length - 1) / (catCount - 1) : pw / 2.0);
                var polygonPoints = $"{firstX:0.#},{oy + ph} {string.Join(" ", points)} {lastX:0.#},{oy + ph}";
                sb.AppendLine($"        <polygon points=\"{polygonPoints}\" fill=\"{colors[s]}\" opacity=\"0.4\"/>");
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Value axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    private void RenderComboChartSvg(StringBuilder sb, DocumentFormat.OpenXml.Drawing.Charts.PlotArea plotArea,
        List<(string name, double[] values)> seriesList, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        // Combo: detect series type from parent chart element
        var barIndices = new HashSet<int>();
        var lineIndices = new HashSet<int>();
        var areaIndices = new HashSet<int>();
        var idx = 0;
        foreach (var chartEl in plotArea.ChildElements)
        {
            var serElements = chartEl.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser").ToList();
            if (serElements.Count == 0) continue;
            var localName = chartEl.LocalName.ToLowerInvariant();
            var isBar = localName.Contains("bar");
            var isArea = localName.Contains("area");
            foreach (var _ in serElements)
            {
                if (isBar) barIndices.Add(idx);
                else if (isArea) areaIndices.Add(idx);
                else lineIndices.Add(idx);
                idx++;
            }
        }

        var allValues = seriesList.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, seriesList.Max(s => s.values.Length));

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        // Bar series
        var barSeries = barIndices.Where(i => i < seriesList.Count).ToList();
        var barCount = barSeries.Count;
        if (barCount > 0)
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.7 / barCount;
            var gap = groupW * 0.15;
            for (int bi = 0; bi < barCount; bi++)
            {
                var s = barSeries[bi];
                for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
                {
                    var val = seriesList[s].values[c];
                    var barH = (val / maxVal) * ph;
                    var bx = ox + c * groupW + gap + bi * barW;
                    var by = oy + ph - barH;
                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                }
            }
        }

        // Area series (render before lines so lines appear on top)
        foreach (var s in areaIndices.Where(i => i < seriesList.Count))
        {
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (seriesList[s].values[c] / maxVal) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastX = ox + (catCount > 1 ? (double)pw * (seriesList[s].values.Length - 1) / (catCount - 1) : pw / 2.0);
                var polygonPoints = $"{firstX:0.#},{oy + ph} {string.Join(" ", points)} {lastX:0.#},{oy + ph}";
                sb.AppendLine($"        <polygon points=\"{polygonPoints}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.3\"/>");
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2\"/>");
            }
        }

        // Line series
        foreach (var s in lineIndices.Where(i => i < seriesList.Count))
        {
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (seriesList[s].values[c] / maxVal) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2.5\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s % colors.Count]}\"/>");
                }
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (double)pw * c / Math.Max(catCount, 1) + (double)pw / Math.Max(catCount, 1) / 2;
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Value axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    private static void RenderRadarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH)
    {
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        if (catCount < 3) return;
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.35;

        // Grid lines (3 rings)
        for (int ring = 1; ring <= 3; ring++)
        {
            var rr = r * ring / 3;
            var gridPoints = new List<string>();
            for (int c = 0; c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                gridPoints.Add($"{cx + rr * Math.Cos(angle):0.#},{cy + rr * Math.Sin(angle):0.#}");
            }
            sb.AppendLine($"        <polygon points=\"{string.Join(" ", gridPoints)}\" fill=\"none\" stroke=\"#444\" stroke-width=\"0.5\"/>");
        }

        // Axis lines
        for (int c = 0; c < catCount; c++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            var ax = cx + r * Math.Cos(angle);
            var ay = cy + r * Math.Sin(angle);
            sb.AppendLine($"        <line x1=\"{cx:0.#}\" y1=\"{cy:0.#}\" x2=\"{ax:0.#}\" y2=\"{ay:0.#}\" stroke=\"#444\" stroke-width=\"0.5\"/>");
        }

        // Data series
        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                var val = series[s].values[c] / maxVal * r;
                points.Add($"{cx + val * Math.Cos(angle):0.#},{cy + val * Math.Sin(angle):0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", points)}\" fill=\"{colors[s]}\" fill-opacity=\"0.2\" stroke=\"{colors[s]}\" stroke-width=\"2\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s]}\"/>");
                }
            }
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            var lx = cx + (r + 15) * Math.Cos(angle);
            var ly = cy + (r + 15) * Math.Sin(angle);
            var anchor = Math.Abs(Math.Cos(angle)) < 0.1 ? "middle" : (Math.Cos(angle) > 0 ? "start" : "end");
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#999\" font-size=\"9\" text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    private static void RenderBubbleChartSvg(StringBuilder sb,
        DocumentFormat.OpenXml.Drawing.Charts.PlotArea plotArea,
        List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        // Read X, Y, and bubble size from each series in the BubbleChart
        var bubbleSeries = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent?.LocalName == "bubbleChart").ToList();

        var allX = new List<double>();
        var allY = new List<double>();
        var allSize = new List<double>();
        var seriesData = new List<(double[] x, double[] y, double[] size)>();

        for (int s = 0; s < bubbleSeries.Count; s++)
        {
            var ser = bubbleSeries[s];
            var xVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "xVal")) ?? [];
            var yVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal")) ?? [];
            var sizeVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "bubbleSize")) ?? yVals;
            seriesData.Add((xVals, yVals, sizeVals));
            allX.AddRange(xVals);
            allY.AddRange(yVals);
            allSize.AddRange(sizeVals);
        }

        // Fallback if no bubble series found
        if (seriesData.Count == 0)
        {
            // Use regular series data as Y, index as X
            foreach (var s in series)
            {
                var xVals = Enumerable.Range(0, s.values.Length).Select(i => (double)i).ToArray();
                seriesData.Add((xVals, s.values, s.values));
                allX.AddRange(xVals);
                allY.AddRange(s.values);
                allSize.AddRange(s.values);
            }
        }

        if (allY.Count == 0) return;
        var minX = allX.Count > 0 ? allX.Min() : 0;
        var maxX = allX.Count > 0 ? allX.Max() : 1;
        if (maxX <= minX) maxX = minX + 1;
        var minY = allY.Min();
        var maxY = allY.Max();
        if (maxY <= minY) maxY = minY + 1;
        var maxSize = allSize.Count > 0 ? allSize.Max() : 1;
        if (maxSize <= 0) maxSize = 1;
        var maxRadius = Math.Min(pw, ph) * 0.08;

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        for (int s = 0; s < seriesData.Count; s++)
        {
            var (xVals, yVals, sizeVals) = seriesData[s];
            var count = Math.Min(xVals.Length, yVals.Length);
            for (int i = 0; i < count; i++)
            {
                var bx = ox + ((xVals[i] - minX) / (maxX - minX)) * pw;
                var by = oy + ph - ((yVals[i] - minY) / (maxY - minY)) * ph;
                var sz = i < sizeVals.Length ? sizeVals[i] : yVals[i];
                var r = (sz / maxSize) * maxRadius + 4;
                sb.AppendLine($"        <circle cx=\"{bx:0.#}\" cy=\"{by:0.#}\" r=\"{r:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.5\"/>");
            }
        }

        // X axis labels (5 ticks)
        for (int t = 0; t <= 4; t++)
        {
            var val = minX + (maxX - minX) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var tx = ox + (double)pw * t / 4;
            sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"8\" text-anchor=\"middle\">{label}</text>");
        }

        // Y axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = minY + (maxY - minY) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    private static void RenderStockChartSvg(StringBuilder sb,
        DocumentFormat.OpenXml.Drawing.Charts.PlotArea plotArea,
        List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        var minVal = allValues.Min();
        if (maxVal <= minVal) { maxVal = minVal + 1; }
        var range = maxVal - minVal;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        // Read up/down bar colors from StockChart
        var upColor = "#2ECC71";
        var downColor = "#E74C3C";
        var stockChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.StockChart>();
        if (stockChart != null)
        {
            var upBars = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "upBars");
            var upFill = upBars?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (upFill != null) upColor = $"#{upFill}";
            var downBars = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "downBars");
            var downFill = downBars?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (downFill != null) downColor = $"#{downFill}";
        }

        // Axis lines
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"#555\" stroke-width=\"1\"/>");

        var groupW = (double)pw / Math.Max(catCount, 1);

        if (series.Count >= 4)
        {
            // OHLC: Open, High, Low, Close
            for (int c = 0; c < catCount; c++)
            {
                var open = c < series[0].values.Length ? series[0].values[c] : 0;
                var high = c < series[1].values.Length ? series[1].values[c] : 0;
                var low = c < series[2].values.Length ? series[2].values[c] : 0;
                var close = c < series[3].values.Length ? series[3].values[c] : 0;
                var cx = ox + c * groupW + groupW / 2;
                var yHigh = oy + ph - ((high - minVal) / range) * ph;
                var yLow = oy + ph - ((low - minVal) / range) * ph;
                var yOpen = oy + ph - ((open - minVal) / range) * ph;
                var yClose = oy + ph - ((close - minVal) / range) * ph;
                var color = close >= open ? upColor : downColor;
                var barW = groupW * 0.5;

                // High-Low line
                sb.AppendLine($"        <line x1=\"{cx:0.#}\" y1=\"{yHigh:0.#}\" x2=\"{cx:0.#}\" y2=\"{yLow:0.#}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
                // Open-Close body
                var bodyTop = Math.Min(yOpen, yClose);
                var bodyH = Math.Abs(yOpen - yClose);
                if (bodyH < 1) bodyH = 1;
                sb.AppendLine($"        <rect x=\"{cx - barW / 2:0.#}\" y=\"{bodyTop:0.#}\" width=\"{barW:0.#}\" height=\"{bodyH:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
        }
        else
        {
            // Fallback: render as line chart
            RenderLineChartSvg(sb, series, categories, colors, ox, oy, pw, ph);
            return;
        }

        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + c * groupW + groupW / 2;
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 14}\" fill=\"#999\" font-size=\"9\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Value axis labels
        for (int t = 0; t <= 4; t++)
        {
            var val = minVal + range * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"#777\" font-size=\"8\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    // ==================== Connector Rendering ====================

    private static void RenderConnector(StringBuilder sb, ConnectionShape cxn, Dictionary<string, string> themeColors)
    {
        var xfrm = cxn.ShapeProperties?.Transform2D;
        if (xfrm?.Offset == null || xfrm?.Extents == null) return;

        var x = xfrm.Offset.X?.Value ?? 0;
        var y = xfrm.Offset.Y?.Value ?? 0;
        var cx = xfrm.Extents.Cx?.Value ?? 0;
        var cy = xfrm.Extents.Cy?.Value ?? 0;

        var flipH = xfrm.HorizontalFlip?.Value == true;
        var flipV = xfrm.VerticalFlip?.Value == true;

        // SVG line
        var outline = cxn.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        var lineColor = "#000000";
        var lineWidth = 1.0;
        if (outline != null)
        {
            var c = ResolveFillColor(outline.GetFirstChild<Drawing.SolidFill>(), themeColors);
            if (c != null) lineColor = c;
            if (outline.Width?.HasValue == true) lineWidth = outline.Width.Value / 12700.0;
        }

        // Ensure minimum dimensions so the line is visible
        // For horizontal lines (cy=0), the container needs height for stroke width
        // For vertical lines (cx=0), the container needs width for stroke width
        var minDimEmu = (long)(lineWidth * 12700 + 12700); // lineWidth + 1pt padding
        var renderCx = Math.Max(cx, cx == 0 ? minDimEmu : 1);
        var renderCy = Math.Max(cy, cy == 0 ? minDimEmu : 1);
        var widthCm = EmuToCm(renderCx);
        var heightCm = EmuToCm(renderCy);

        // Adjust y position upward by half the added height for zero-height lines
        var renderY = cy == 0 ? y - minDimEmu / 2 : y;
        var renderX = cx == 0 ? x - minDimEmu / 2 : x;

        var x1 = flipH ? "100%" : "0";
        var y1 = flipV ? "100%" : "0";
        var x2 = flipH ? "0" : "100%";
        var y2 = flipV ? "0" : "100%";

        // For straight lines (one dimension is 0), draw from center
        string svgY1, svgY2, svgX1, svgX2;
        if (cy == 0)
        {
            // Horizontal line: draw at vertical center
            svgX1 = flipH ? "100%" : "0";
            svgX2 = flipH ? "0" : "100%";
            svgY1 = svgY2 = "50%";
        }
        else if (cx == 0)
        {
            // Vertical line: draw at horizontal center
            svgX1 = svgX2 = "50%";
            svgY1 = flipV ? "100%" : "0";
            svgY2 = flipV ? "0" : "100%";
        }
        else
        {
            svgX1 = x1; svgY1 = y1; svgX2 = x2; svgY2 = y2;
        }

        // Dash pattern
        var dashAttr = "";
        var prstDash = outline?.GetFirstChild<Drawing.PresetDash>();
        if (prstDash?.Val?.HasValue == true)
        {
            var dashVal = prstDash.Val.InnerText;
            var dashArray = dashVal switch
            {
                "dash" or "lgDash" => $"{lineWidth * 4:0.##},{lineWidth * 3:0.##}",
                "sysDash" => $"{lineWidth * 3:0.##},{lineWidth * 1:0.##}",
                "dot" or "sysDot" => $"{lineWidth * 1:0.##},{lineWidth * 2:0.##}",
                "dashDot" => $"{lineWidth * 4:0.##},{lineWidth * 2:0.##},{lineWidth * 1:0.##},{lineWidth * 2:0.##}",
                "lgDashDot" => $"{lineWidth * 6:0.##},{lineWidth * 2:0.##},{lineWidth * 1:0.##},{lineWidth * 2:0.##}",
                "lgDashDotDot" => $"{lineWidth * 6:0.##},{lineWidth * 2:0.##},{lineWidth * 1:0.##},{lineWidth * 2:0.##},{lineWidth * 1:0.##},{lineWidth * 2:0.##}",
                _ => ""
            };
            if (!string.IsNullOrEmpty(dashArray))
                dashAttr = $" stroke-dasharray=\"{dashArray}\"";
        }

        // Arrow markers
        var headEnd = outline?.GetFirstChild<Drawing.HeadEnd>();
        var tailEnd = outline?.GetFirstChild<Drawing.TailEnd>();
        var hasHead = headEnd?.Type?.HasValue == true && headEnd.Type.InnerText != "none";
        var hasTail = tailEnd?.Type?.HasValue == true && tailEnd.Type.InnerText != "none";
        var markerDefs = "";
        var markerStartAttr = "";
        var markerEndAttr = "";
        var safeColor = CssSanitizeColor(lineColor);

        if (hasHead || hasTail)
        {
            var arrowSize = Math.Max(3, lineWidth * 3);
            var defs = new StringBuilder();
            defs.Append("<defs>");
            if (hasHead)
            {
                defs.Append($"<marker id=\"ah\" markerWidth=\"{arrowSize:0.#}\" markerHeight=\"{arrowSize:0.#}\" refX=\"{arrowSize:0.#}\" refY=\"{arrowSize / 2:0.#}\" orient=\"auto-start-reverse\"><polygon points=\"{arrowSize:0.#} 0,0 {arrowSize / 2:0.#},{arrowSize:0.#} {arrowSize:0.#}\" fill=\"{safeColor}\"/></marker>");
                markerStartAttr = " marker-start=\"url(#ah)\"";
            }
            if (hasTail)
            {
                defs.Append($"<marker id=\"at\" markerWidth=\"{arrowSize:0.#}\" markerHeight=\"{arrowSize:0.#}\" refX=\"0\" refY=\"{arrowSize / 2:0.#}\" orient=\"auto\"><polygon points=\"0 0,{arrowSize:0.#} {arrowSize / 2:0.#},0 {arrowSize:0.#}\" fill=\"{safeColor}\"/></marker>");
                markerEndAttr = " marker-end=\"url(#at)\"";
            }
            defs.Append("</defs>");
            markerDefs = defs.ToString();
        }

        sb.AppendLine($"    <div class=\"connector\" style=\"left:{EmuToCm(renderX)}cm;top:{EmuToCm(renderY)}cm;width:{widthCm}cm;height:{heightCm}cm\">");
        sb.AppendLine($"      <svg width=\"100%\" height=\"100%\" preserveAspectRatio=\"none\">");
        if (!string.IsNullOrEmpty(markerDefs))
            sb.AppendLine($"        {markerDefs}");
        sb.AppendLine($"        <line x1=\"{svgX1}\" y1=\"{svgY1}\" x2=\"{svgX2}\" y2=\"{svgY2}\" stroke=\"{safeColor}\" stroke-width=\"{lineWidth:0.##}\"{dashAttr}{markerStartAttr}{markerEndAttr}/>");
        sb.AppendLine("      </svg>");
        sb.AppendLine("    </div>");
    }

    // ==================== Group Rendering ====================

    private void RenderGroup(StringBuilder sb, GroupShape grp, SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        var grpXfrm = grp.GroupShapeProperties?.TransformGroup;
        if (grpXfrm?.Offset == null || grpXfrm?.Extents == null) return;

        var x = grpXfrm.Offset.X?.Value ?? 0;
        var y = grpXfrm.Offset.Y?.Value ?? 0;
        var cx = grpXfrm.Extents.Cx?.Value ?? 0;
        var cy = grpXfrm.Extents.Cy?.Value ?? 0;

        // Child offset/extents for coordinate transformation
        var childOff = grpXfrm.ChildOffset;
        var childExt = grpXfrm.ChildExtents;
        var scaleX = (childExt?.Cx?.Value ?? cx) != 0 ? (double)cx / (childExt?.Cx?.Value ?? cx) : 1.0;
        var scaleY = (childExt?.Cy?.Value ?? cy) != 0 ? (double)cy / (childExt?.Cy?.Value ?? cy) : 1.0;
        var offX = childOff?.X?.Value ?? 0;
        var offY = childOff?.Y?.Value ?? 0;

        sb.AppendLine($"    <div class=\"group\" style=\"left:{EmuToCm(x)}cm;top:{EmuToCm(y)}cm;width:{EmuToCm(cx)}cm;height:{EmuToCm(cy)}cm\">");

        foreach (var child in grp.ChildElements)
        {
            switch (child)
            {
                case Shape shape:
                {
                    var pos = CalcGroupChildPos(shape.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderShape(sb, shape, slidePart, themeColors, pos);
                    break;
                }
                case Picture pic:
                {
                    var pos = CalcGroupChildPos(pic.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderPicture(sb, pic, slidePart, themeColors, pos);
                    break;
                }
                case GroupShape nestedGrp:
                {
                    // Nested group: calculate the group's own position within parent group
                    var nestedXfrm = nestedGrp.GroupShapeProperties?.TransformGroup;
                    if (nestedXfrm?.Offset != null && nestedXfrm?.Extents != null)
                    {
                        var nx = (long)((( nestedXfrm.Offset.X?.Value ?? 0) - offX) * scaleX);
                        var ny = (long)(((nestedXfrm.Offset.Y?.Value ?? 0) - offY) * scaleY);
                        var ncx = (long)((nestedXfrm.Extents.Cx?.Value ?? 0) * scaleX);
                        var ncy = (long)((nestedXfrm.Extents.Cy?.Value ?? 0) * scaleY);
                        RenderNestedGroup(sb, nestedGrp, slidePart, themeColors, nx, ny, ncx, ncy);
                    }
                    break;
                }
                case ConnectionShape cxn:
                {
                    RenderConnector(sb, cxn, themeColors);
                    break;
                }
            }
        }

        sb.AppendLine("    </div>");
    }

    /// <summary>
    /// Pure calculation: compute adjusted coordinates for a group child element.
    /// Returns null if the element has no transform. NEVER modifies the original element.
    /// </summary>
    private static (long x, long y, long cx, long cy)? CalcGroupChildPos(
        Drawing.Transform2D? xfrm, long offX, long offY, double scaleX, double scaleY)
    {
        if (xfrm?.Offset == null || xfrm?.Extents == null) return null;

        var origX = xfrm.Offset.X?.Value ?? 0;
        var origY = xfrm.Offset.Y?.Value ?? 0;
        var origCx = xfrm.Extents.Cx?.Value ?? 0;
        var origCy = xfrm.Extents.Cy?.Value ?? 0;

        return (
            (long)((origX - offX) * scaleX),
            (long)((origY - offY) * scaleY),
            (long)(origCx * scaleX),
            (long)(origCy * scaleY)
        );
    }

    /// <summary>
    /// Render a nested group with pre-calculated position (from parent group transform).
    /// Recursively handles arbitrary nesting depth.
    /// </summary>
    private void RenderNestedGroup(StringBuilder sb, GroupShape grp, SlidePart slidePart,
        Dictionary<string, string> themeColors, long x, long y, long cx, long cy)
    {
        var grpXfrm = grp.GroupShapeProperties?.TransformGroup;

        // Child coordinate system of this nested group
        var childOff = grpXfrm?.ChildOffset;
        var childExt = grpXfrm?.ChildExtents;
        var scaleX = (childExt?.Cx?.Value ?? cx) != 0 ? (double)cx / (childExt?.Cx?.Value ?? cx) : 1.0;
        var scaleY = (childExt?.Cy?.Value ?? cy) != 0 ? (double)cy / (childExt?.Cy?.Value ?? cy) : 1.0;
        var offX = childOff?.X?.Value ?? 0;
        var offY = childOff?.Y?.Value ?? 0;

        sb.AppendLine($"    <div class=\"group\" style=\"left:{EmuToCm(x)}cm;top:{EmuToCm(y)}cm;width:{EmuToCm(cx)}cm;height:{EmuToCm(cy)}cm\">");

        foreach (var child in grp.ChildElements)
        {
            switch (child)
            {
                case Shape shape:
                {
                    var pos = CalcGroupChildPos(shape.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderShape(sb, shape, slidePart, themeColors, pos);
                    break;
                }
                case Picture pic:
                {
                    var pos = CalcGroupChildPos(pic.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderPicture(sb, pic, slidePart, themeColors, pos);
                    break;
                }
                case GroupShape nestedGrp:
                {
                    var nestedXfrm = nestedGrp.GroupShapeProperties?.TransformGroup;
                    if (nestedXfrm?.Offset != null && nestedXfrm?.Extents != null)
                    {
                        var nx = (long)(((nestedXfrm.Offset.X?.Value ?? 0) - offX) * scaleX);
                        var ny = (long)(((nestedXfrm.Offset.Y?.Value ?? 0) - offY) * scaleY);
                        var ncx = (long)((nestedXfrm.Extents.Cx?.Value ?? 0) * scaleX);
                        var ncy = (long)((nestedXfrm.Extents.Cy?.Value ?? 0) * scaleY);
                        RenderNestedGroup(sb, nestedGrp, slidePart, themeColors, nx, ny, ncx, ncy);
                    }
                    break;
                }
                case ConnectionShape cxn:
                    RenderConnector(sb, cxn, themeColors);
                    break;
            }
        }

        sb.AppendLine("    </div>");
    }

    // ==================== CSS Helper: Fill ====================

    private static string GetShapeFillCss(ShapeProperties? spPr, OpenXmlPart part, Dictionary<string, string> themeColors)
    {
        if (spPr == null) return "";

        // NoFill
        if (spPr.GetFirstChild<Drawing.NoFill>() != null)
            return "background:transparent";

        // Solid fill
        var solidFill = spPr.GetFirstChild<Drawing.SolidFill>();
        if (solidFill != null)
        {
            var color = ResolveFillColor(solidFill, themeColors);
            if (color != null) return $"background:{color}";
        }

        // Gradient fill
        var gradFill = spPr.GetFirstChild<Drawing.GradientFill>();
        if (gradFill != null)
            return $"background:{GradientToCss(gradFill, themeColors)}";

        // Image fill (blip)
        var blipFill = spPr.GetFirstChild<Drawing.BlipFill>();
        if (blipFill != null)
        {
            var dataUri = BlipToDataUri(blipFill, part);
            if (dataUri != null)
                return $"background:url('{dataUri}') center/cover no-repeat";
        }

        return "";
    }

    // ==================== CSS Helper: Custom Geometry ====================

    /// <summary>
    /// Convert OOXML CustomGeometry (a:custGeom) path data to CSS clip-path.
    /// Supports moveTo, lineTo, cubicBezTo, quadBezTo, close.
    /// Coordinates are in the path's own coordinate system (w/h),
    /// converted to percentages for clip-path.
    /// </summary>
    private static string CustomGeometryToClipPath(Drawing.CustomGeometry custGeom)
    {
        var pathList = custGeom.GetFirstChild<Drawing.PathList>();
        if (pathList == null) return "";

        var path = pathList.GetFirstChild<Drawing.Path>();
        if (path == null) return "";

        // Path coordinate system
        var pathW = path.Width?.HasValue == true ? path.Width.Value : 100000L;
        var pathH = path.Height?.HasValue == true ? path.Height.Value : 100000L;
        if (pathW == 0) pathW = 100000;
        if (pathH == 0) pathH = 100000;

        // Helper: parse Drawing.Point X/Y (StringValue) to double percentage
        static bool TryParsePoint(Drawing.Point? pt, double pw, double ph, out double px, out double py)
        {
            px = py = 0;
            if (pt?.X?.HasValue != true || pt?.Y?.HasValue != true) return false;
            if (!long.TryParse(pt.X.Value, out var xv) || !long.TryParse(pt.Y.Value, out var yv)) return false;
            px = xv * 100.0 / pw;
            py = yv * 100.0 / ph;
            return true;
        }

        // Try polygon first (only moveTo + lineTo + close = all straight lines)
        bool hasOnlyLines = true;
        foreach (var child in path.ChildElements)
        {
            if (child is Drawing.CubicBezierCurveTo or Drawing.QuadraticBezierCurveTo)
            {
                hasOnlyLines = false;
                break;
            }
        }

        if (hasOnlyLines)
        {
            // Use clip-path: polygon() — better browser support
            var points = new List<string>();
            foreach (var child in path.ChildElements)
            {
                switch (child)
                {
                    case Drawing.MoveTo moveTo:
                        if (TryParsePoint(moveTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var mx, out var my))
                            points.Add($"{mx:0.##}% {my:0.##}%");
                        break;
                    case Drawing.LineTo lineTo:
                        if (TryParsePoint(lineTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var lx, out var ly))
                            points.Add($"{lx:0.##}% {ly:0.##}%");
                        break;
                    case Drawing.CloseShapePath:
                        break; // polygon implicitly closes
                }
            }
            if (points.Count >= 3)
                return $"clip-path:polygon({string.Join(",", points)})";
        }
        else
        {
            // Has curves — use clip-path: path() with SVG path syntax
            var svgPath = new StringBuilder();
            foreach (var child in path.ChildElements)
            {
                switch (child)
                {
                    case Drawing.MoveTo moveTo:
                        if (TryParsePoint(moveTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var mx, out var my))
                            svgPath.Append($"M {mx:0.##} {my:0.##} ");
                        break;
                    case Drawing.LineTo lineTo:
                        if (TryParsePoint(lineTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var lx, out var ly))
                            svgPath.Append($"L {lx:0.##} {ly:0.##} ");
                        break;
                    case Drawing.CubicBezierCurveTo cubicBez:
                    {
                        var pts = cubicBez.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 3
                            && TryParsePoint(pts[0], pathW, pathH, out var c1x, out var c1y)
                            && TryParsePoint(pts[1], pathW, pathH, out var c2x, out var c2y)
                            && TryParsePoint(pts[2], pathW, pathH, out var c3x, out var c3y))
                            svgPath.Append($"C {c1x:0.##} {c1y:0.##},{c2x:0.##} {c2y:0.##},{c3x:0.##} {c3y:0.##} ");
                        break;
                    }
                    case Drawing.QuadraticBezierCurveTo quadBez:
                    {
                        var pts = quadBez.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 2
                            && TryParsePoint(pts[0], pathW, pathH, out var q1x, out var q1y)
                            && TryParsePoint(pts[1], pathW, pathH, out var q2x, out var q2y))
                            svgPath.Append($"Q {q1x:0.##} {q1y:0.##},{q2x:0.##} {q2y:0.##} ");
                        break;
                    }
                    case Drawing.CloseShapePath:
                        svgPath.Append("Z ");
                        break;
                }
            }
            var pathStr = svgPath.ToString().Trim();
            if (!string.IsNullOrEmpty(pathStr))
                return $"clip-path:path('{pathStr}')";
        }

        return "";
    }

    // ==================== CSS Helper: Gradient ====================

    private static string GradientToCss(Drawing.GradientFill gradFill, Dictionary<string, string> themeColors)
    {
        var stops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stops == null || stops.Count < 2) return "transparent";

        var cssStops = new List<string>();
        foreach (var gs in stops)
        {
            var color = ResolveFillColor(gs.GetFirstChild<Drawing.SolidFill>(), themeColors);
            if (color == null)
            {
                // Try direct color children
                var rgb = gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
                    color = $"#{rgb[..6]}";
                else
                {
                    var scheme = gs.GetFirstChild<Drawing.SchemeColor>()?.Val?.InnerText;
                    color = scheme != null && themeColors.TryGetValue(scheme, out var tc) ? $"#{tc}" : "#808080";
                }
            }
            var pos = gs.Position?.Value;
            if (pos.HasValue)
                cssStops.Add($"{color} {pos.Value / 1000.0:0.##}%");
            else
                cssStops.Add(color);
        }

        // Radial or linear?
        var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
        if (pathGrad != null)
            return $"radial-gradient(circle, {string.Join(", ", cssStops)})";

        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        var angleDeg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 90.0;
        // OOXML angle 0° = top→bottom (same as CSS 180deg), so CSS angle = OOXML + 90°
        // Actually OOXML: 0 = right, 90 = bottom; CSS: 0 = up, 90 = right
        var cssAngle = angleDeg + 90;

        return $"linear-gradient({cssAngle:0.##}deg, {string.Join(", ", cssStops)})";
    }

    // ==================== CSS Helper: Outline/Border ====================

    private static string OutlineToCss(Drawing.Outline outline, Dictionary<string, string> themeColors)
    {
        if (outline.GetFirstChild<Drawing.NoFill>() != null) return "";

        var color = ResolveFillColor(outline.GetFirstChild<Drawing.SolidFill>(), themeColors) ?? "#000000";
        var widthPt = outline.Width?.HasValue == true ? outline.Width.Value / 12700.0 : 1.0;
        if (widthPt < 0.5) widthPt = 0.5;

        var dash = outline.GetFirstChild<Drawing.PresetDash>();
        var borderStyle = "solid";
        if (dash?.Val?.HasValue == true)
        {
            borderStyle = dash.Val.InnerText switch
            {
                "dash" or "lgDash" or "sysDash" => "dashed",
                "dot" or "sysDot" => "dotted",
                "dashDot" or "lgDashDot" or "sysDashDot" or "sysDashDotDot" => "dashed",
                _ => "solid"
            };
        }

        return $"border:{widthPt:0.##}pt {borderStyle} {color}";
    }

    // ==================== CSS Helper: Shadow ====================

    private static string EffectListToShadowCss(Drawing.EffectList? effectList, Dictionary<string, string> themeColors)
    {
        if (effectList == null) return "";

        var shadow = effectList.GetFirstChild<Drawing.OuterShadow>();
        if (shadow == null) return "";

        var color = "rgba(0,0,0,0.3)";
        var rgb = shadow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        var alpha = shadow.Descendants<Drawing.Alpha>().FirstOrDefault()?.Val?.Value ?? 50000;
        var opacity = alpha / 100000.0;
        if (rgb != null)
        {
            var r = Convert.ToInt32(rgb[..2], 16);
            var g = Convert.ToInt32(rgb[2..4], 16);
            var b = Convert.ToInt32(rgb[4..6], 16);
            color = $"rgba({r},{g},{b},{opacity:0.##})";
        }

        var blurPt = shadow.BlurRadius?.HasValue == true ? shadow.BlurRadius.Value / 12700.0 : 4;
        var distPt = shadow.Distance?.HasValue == true ? shadow.Distance.Value / 12700.0 : 3;
        var angleDeg = shadow.Direction?.HasValue == true ? shadow.Direction.Value / 60000.0 : 45;
        var angleRad = angleDeg * Math.PI / 180;
        var offsetX = distPt * Math.Cos(angleRad);
        var offsetY = distPt * Math.Sin(angleRad);

        return $"box-shadow:{offsetX:0.##}pt {offsetY:0.##}pt {blurPt:0.##}pt {color}";
    }

    // ==================== CSS Helper: Preset Geometry ====================

    private static string PresetGeometryToCss(string preset)
    {
        return preset switch
        {
            // Rectangles
            "rect" => "",
            "roundRect" => "border-radius:8px",
            "snip1Rect" => "clip-path:polygon(0 0,92% 0,100% 8%,100% 100%,0 100%)",
            "snip2SameRect" => "clip-path:polygon(8% 0,92% 0,100% 8%,100% 100%,0 100%,0 8%)",
            "snip2DiagRect" => "clip-path:polygon(8% 0,100% 0,100% 92%,92% 100%,0 100%,0 8%)",
            "round1Rect" => "border-radius:8px 0 0 0",
            "round2SameRect" => "border-radius:8px 8px 0 0",
            "round2DiagRect" => "border-radius:8px 0 8px 0",

            // Ellipses
            "ellipse" => "border-radius:50%",

            // Triangles
            "triangle" or "isosTriangle" => "clip-path:polygon(50% 0,100% 100%,0 100%)",
            "rtTriangle" => "clip-path:polygon(0 0,100% 100%,0 100%)",

            // Diamonds and parallelograms
            "diamond" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            "parallelogram" => "clip-path:polygon(15% 0,100% 0,85% 100%,0 100%)",
            "trapezoid" => "clip-path:polygon(20% 0,80% 0,100% 100%,0 100%)",

            // Polygons
            "pentagon" => "clip-path:polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%)",
            "hexagon" => "clip-path:polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%)",
            "heptagon" => "clip-path:polygon(50% 0,90% 20%,100% 60%,75% 100%,25% 100%,0 60%,10% 20%)",
            "octagon" => "clip-path:polygon(29% 0,71% 0,100% 29%,100% 71%,71% 100%,29% 100%,0 71%,0 29%)",
            "decagon" => "clip-path:polygon(35% 0,65% 0,90% 12%,100% 38%,100% 62%,90% 88%,65% 100%,35% 100%,10% 88%,0 62%,0 38%,10% 12%)",
            "dodecagon" => "clip-path:polygon(37% 0,63% 0,87% 13%,100% 37%,100% 63%,87% 87%,63% 100%,37% 100%,13% 87%,0 63%,0 37%,13% 13%)",

            // Stars
            "star4" => "clip-path:polygon(50% 0,62% 38%,100% 50%,62% 62%,50% 100%,38% 62%,0 50%,38% 38%)",
            "star5" => "clip-path:polygon(50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)",
            "star6" => "clip-path:polygon(50% 0,63% 25%,100% 25%,75% 50%,100% 75%,63% 75%,50% 100%,37% 75%,0 75%,25% 50%,0 25%,37% 25%)",
            "star8" => "clip-path:polygon(50% 0,62% 19%,85% 15%,81% 38%,100% 50%,81% 62%,85% 85%,62% 81%,50% 100%,38% 81%,15% 85%,19% 62%,0 50%,19% 38%,15% 15%,38% 19%)",
            "star10" => "clip-path:polygon(50% 0,59% 19%,79% 5%,74% 27%,97% 25%,84% 43%,100% 50%,84% 57%,97% 75%,74% 73%,79% 95%,59% 81%,50% 100%,41% 81%,21% 95%,26% 73%,3% 75%,16% 57%,0 50%,16% 43%,3% 25%,26% 27%,21% 5%,41% 19%)",
            "star12" => "clip-path:polygon(50% 0,57% 15%,75% 7%,71% 25%,93% 25%,84% 42%,100% 50%,84% 58%,93% 75%,71% 75%,75% 93%,57% 85%,50% 100%,43% 85%,25% 93%,29% 75%,7% 75%,16% 58%,0 50%,16% 42%,7% 25%,29% 25%,25% 7%,43% 15%)",

            // Arrows
            "rightArrow" => "clip-path:polygon(0 20%,70% 20%,70% 0,100% 50%,70% 100%,70% 80%,0 80%)",
            "leftArrow" => "clip-path:polygon(30% 0,30% 20%,100% 20%,100% 80%,30% 80%,30% 100%,0 50%)",
            "upArrow" => "clip-path:polygon(20% 30%,50% 0,80% 30%,80% 100%,20% 100%)",
            "downArrow" => "clip-path:polygon(20% 0,80% 0,80% 70%,100% 70%,50% 100%,0 70%,20% 70%)",
            "leftRightArrow" => "clip-path:polygon(0 50%,15% 20%,15% 35%,85% 35%,85% 20%,100% 50%,85% 80%,85% 65%,15% 65%,15% 80%)",
            "upDownArrow" => "clip-path:polygon(50% 0,80% 15%,65% 15%,65% 85%,80% 85%,50% 100%,20% 85%,35% 85%,35% 15%,20% 15%)",
            "notchedRightArrow" => "clip-path:polygon(0 20%,70% 20%,70% 0,100% 50%,70% 100%,70% 80%,0 80%,10% 50%)",
            "bentArrow" => "clip-path:polygon(0 20%,60% 20%,60% 0,100% 35%,60% 70%,60% 50%,20% 50%,20% 100%,0 100%)",
            "chevron" => "clip-path:polygon(0 0,80% 0,100% 50%,80% 100%,0 100%,20% 50%)",
            "homePlate" => "clip-path:polygon(0 0,85% 0,100% 50%,85% 100%,0 100%)",
            "stripedRightArrow" => "clip-path:polygon(10% 20%,12% 20%,12% 80%,10% 80%,10% 20%,15% 20%,70% 20%,70% 0,100% 50%,70% 100%,70% 80%,15% 80%)",

            // Callouts
            "wedgeRoundRectCallout" => "border-radius:6px",
            "wedgeRectCallout" or "wedgeEllipseCallout" => "",
            "cloudCallout" => "border-radius:50%",

            // Crosses and plus
            "plus" or "cross" => "clip-path:polygon(33% 0,67% 0,67% 33%,100% 33%,100% 67%,67% 67%,67% 100%,33% 100%,33% 67%,0 67%,0 33%,33% 33%)",

            // Heart (polygon approximation)
            "heart" => "clip-path:polygon(50% 18%,65% 0,85% 0,100% 15%,100% 35%,50% 100%,0 35%,0 15%,15% 0,35% 0)",

            // Cloud (rounded blob)
            "cloud" or "cloudCallout" => "border-radius:50% 50% 45% 55% / 60% 40% 55% 45%",

            // Smiley (circle)
            "smileyFace" or "smiley" => "border-radius:50%",

            // Gear (polygon approximation of 6-tooth gear)
            "gear6" => "clip-path:polygon(50% 0,61% 10%,75% 3%,80% 18%,97% 25%,88% 38%,100% 50%,88% 62%,97% 75%,80% 82%,75% 97%,61% 90%,50% 100%,39% 90%,25% 97%,20% 82%,3% 75%,12% 62%,0 50%,12% 38%,3% 25%,20% 18%,25% 3%,39% 10%)",
            "gear9" => "clip-path:polygon(50% 0,56% 8%,65% 2%,68% 12%,78% 9%,78% 20%,88% 20%,85% 30%,95% 35%,90% 44%,100% 50%,90% 56%,95% 65%,85% 70%,88% 80%,78% 80%,78% 91%,68% 88%,65% 98%,56% 92%,50% 100%,44% 92%,35% 98%,32% 88%,22% 91%,22% 80%,12% 80%,15% 70%,5% 65%,10% 56%,0 50%,10% 44%,5% 35%,15% 30%,12% 20%,22% 20%,22% 9%,32% 12%,35% 2%,44% 8%)",

            // 3D-like shapes (rendered flat)
            "cube" => "",
            "can" or "cylinder" => "border-radius:50%/10%",
            "bevel" => "border:3px outset #888",
            "foldedCorner" => "clip-path:polygon(0 0,85% 0,100% 15%,100% 100%,0 100%)",
            "lightningBolt" => "clip-path:polygon(35% 0,55% 35%,100% 30%,45% 55%,80% 100%,25% 60%,0 80%,30% 45%)",

            // Misc shapes
            "frame" => "clip-path:polygon(0 0,100% 0,100% 100%,0 100%,0 12%,12% 12%,12% 88%,88% 88%,88% 12%,0 12%)",
            "donut" => "border-radius:50%", // approximate — real donut has inner hole
            "noSmoking" => "border-radius:50%",
            "halfFrame" => "clip-path:polygon(0 0,100% 0,100% 15%,15% 15%,15% 100%,0 100%)",
            "corner" => "clip-path:polygon(0 0,50% 0,50% 50%,100% 50%,100% 100%,0 100%)",
            "pie" or "arc" => "border-radius:50%",

            // Ribbons/banners
            "ribbon" or "ribbon2" or "wave" or "doubleWave" => "",
            "horizontalScroll" or "verticalScroll" => "border-radius:4px",

            // Flowchart
            "flowChartProcess" => "",
            "flowChartAlternateProcess" => "border-radius:8px",
            "flowChartDecision" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            "flowChartInputOutput" or "flowChartData" => "clip-path:polygon(15% 0,100% 0,85% 100%,0 100%)",
            "flowChartPredefinedProcess" => "border-left:3px double currentColor;border-right:3px double currentColor",
            "flowChartDocument" => "",
            "flowChartMultidocument" => "",
            "flowChartTerminator" => "border-radius:50%/100%",
            "flowChartPreparation" => "clip-path:polygon(17% 0,83% 0,100% 50%,83% 100%,17% 100%,0 50%)",
            "flowChartManualInput" => "clip-path:polygon(0 15%,100% 0,100% 100%,0 100%)",
            "flowChartManualOperation" => "clip-path:polygon(0 0,100% 0,85% 100%,15% 100%)",
            "flowChartMerge" => "clip-path:polygon(0 0,100% 0,50% 100%)",
            "flowChartExtract" => "clip-path:polygon(50% 0,100% 100%,0 100%)",
            "flowChartSort" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            "flowChartCollate" => "clip-path:polygon(0 0,100% 0,50% 50%,100% 100%,0 100%,50% 50%)",
            "flowChartDelay" => "border-radius:0 50% 50% 0",
            "flowChartDisplay" => "clip-path:polygon(0 50%,15% 0,85% 0,100% 50%,85% 100%,15% 100%)",
            "flowChartPunchedCard" => "clip-path:polygon(15% 0,100% 0,100% 100%,0 100%,0 15%)",
            "flowChartPunchedTape" => "",
            "flowChartOnlineStorage" => "border-radius:50% 0 0 50%",
            "flowChartOfflineStorage" => "clip-path:polygon(10% 0,90% 0,50% 100%)",
            "flowChartMagneticDisk" => "border-radius:50%/20%",
            "flowChartConnector" or "flowChartOffpageConnector" => "border-radius:50%",

            // Block arrows
            "curvedRightArrow" or "curvedLeftArrow" or "curvedUpArrow" or "curvedDownArrow" => "",
            "circularArrow" => "border-radius:50%",

            // Math
            "mathPlus" => "clip-path:polygon(33% 0,67% 0,67% 33%,100% 33%,100% 67%,67% 67%,67% 100%,33% 100%,33% 67%,0 67%,0 33%,33% 33%)",
            "mathMinus" => "clip-path:polygon(0 35%,100% 35%,100% 65%,0 65%)",
            "mathMultiply" => "clip-path:polygon(20% 0,50% 30%,80% 0,100% 20%,70% 50%,100% 80%,80% 100%,50% 70%,20% 100%,0 80%,30% 50%,0 20%)",
            "mathDivide" => "",
            "mathEqual" => "clip-path:polygon(0 25%,100% 25%,100% 40%,0 40%,0 60%,100% 60%,100% 75%,0 75%)",
            "mathNotEqual" => "",

            // Default: render as rectangle
            _ => ""
        };
    }

    // ==================== Color Resolution ====================

    private static string? ResolveFillColor(Drawing.SolidFill? solidFill, Dictionary<string, string> themeColors)
    {
        if (solidFill == null) return null;

        var rgb = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
        {
            var hexPart = rgb[..6]; // Only use first 6 hex chars, ignore any trailing data
            var alpha = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (alpha.HasValue && alpha.Value < 100000)
            {
                var r = Convert.ToInt32(hexPart[..2], 16);
                var g = Convert.ToInt32(hexPart[2..4], 16);
                var b = Convert.ToInt32(hexPart[4..6], 16);
                return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";
            }
            return $"#{hexPart}";
        }

        var schemeColor = solidFill.GetFirstChild<Drawing.SchemeColor>();
        if (schemeColor?.Val?.HasValue == true)
        {
            var schemeName = schemeColor.Val!.InnerText;
            if (schemeName != null && themeColors.TryGetValue(schemeName, out var themeHex))
            {
                // Check for lumMod/lumOff/tint/shade transforms
                var color = ApplyColorTransforms(themeHex, schemeColor);
                return color;
            }
            return null; // Unknown scheme color
        }

        return null;
    }

    private static string ApplyColorTransforms(string hex, Drawing.SchemeColor schemeColor)
    {
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        var lumMod = schemeColor.GetFirstChild<Drawing.LuminanceModulation>()?.Val?.Value;
        var lumOff = schemeColor.GetFirstChild<Drawing.LuminanceOffset>()?.Val?.Value;
        var tint = schemeColor.GetFirstChild<Drawing.Tint>()?.Val?.Value;
        var shade = schemeColor.GetFirstChild<Drawing.Shade>()?.Val?.Value;
        var alpha = schemeColor.GetFirstChild<Drawing.Alpha>()?.Val?.Value;

        // OOXML spec: tint blends toward white, shade blends toward black
        if (tint.HasValue)
        {
            var t = tint.Value / 100000.0;
            r = (int)(r + (255 - r) * (1 - t));
            g = (int)(g + (255 - g) * (1 - t));
            b = (int)(b + (255 - b) * (1 - t));
        }

        if (shade.HasValue)
        {
            var s = shade.Value / 100000.0;
            r = (int)(r * s);
            g = (int)(g * s);
            b = (int)(b * s);
        }

        // OOXML spec: lumMod/lumOff operate in HSL space
        if (lumMod.HasValue || lumOff.HasValue)
        {
            var mod = (lumMod ?? 100000) / 100000.0;
            var off = (lumOff ?? 0) / 100000.0;
            RgbToHsl(r, g, b, out var h, out var s, out var l);
            l = Math.Clamp(l * mod + off, 0, 1);
            HslToRgb(h, s, l, out r, out g, out b);
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);

        if (alpha.HasValue && alpha.Value < 100000)
            return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";

        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private static void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
    {
        var rf = r / 255.0;
        var gf = g / 255.0;
        var bf = b / 255.0;
        var max = Math.Max(rf, Math.Max(gf, bf));
        var min = Math.Min(rf, Math.Min(gf, bf));
        var delta = max - min;

        l = (max + min) / 2.0;

        if (delta < 1e-10)
        {
            h = 0;
            s = 0;
            return;
        }

        s = l < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);

        if (Math.Abs(max - rf) < 1e-10)
            h = ((gf - bf) / delta + (gf < bf ? 6 : 0)) / 6.0;
        else if (Math.Abs(max - gf) < 1e-10)
            h = ((bf - rf) / delta + 2) / 6.0;
        else
            h = ((rf - gf) / delta + 4) / 6.0;
    }

    private static void HslToRgb(double h, double s, double l, out int r, out int g, out int b)
    {
        if (s < 1e-10)
        {
            r = g = b = (int)Math.Round(l * 255);
            return;
        }

        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;

        r = (int)Math.Round(HueToRgb(p, q, h + 1.0 / 3) * 255);
        g = (int)Math.Round(HueToRgb(p, q, h) * 255);
        b = (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3) * 255);
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    /// <summary>
    /// Build a map of scheme color names to hex values from the presentation theme.
    /// </summary>
    private Dictionary<string, string> ResolveThemeColorMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var theme = _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault()?.ThemePart?.Theme;
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return map;

        void Add(string name, OpenXmlCompositeElement? color)
        {
            if (color == null) return;
            var rgb = color.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            var sys = color.GetFirstChild<Drawing.SystemColor>();
            var srgb = sys?.LastColor?.Value ?? sys?.Val?.InnerText;
            var hex = rgb ?? srgb;
            if (hex != null) map[name] = hex;
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
        if (map.TryGetValue("dk1", out var dk1)) { map["tx1"] = dk1; map["dark1"] = dk1; map["text1"] = dk1; }
        if (map.TryGetValue("dk2", out var dk2)) { map["dark2"] = dk2; map["text2"] = dk2; map["tx2"] = dk2; }
        if (map.TryGetValue("lt1", out var lt1)) { map["bg1"] = lt1; map["light1"] = lt1; map["background1"] = lt1; }
        if (map.TryGetValue("lt2", out var lt2)) { map["bg2"] = lt2; map["light2"] = lt2; map["background2"] = lt2; }

        return map;
    }

    // ==================== Image Helpers ====================

    private static string? BlipToDataUri(Drawing.BlipFill blipFill, OpenXmlPart part)
    {
        var blip = blipFill.GetFirstChild<Drawing.Blip>();
        if (blip?.Embed?.HasValue != true) return null;

        try
        {
            var imgPart = part.GetPartById(blip.Embed.Value!);
            using var stream = imgPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());
            return $"data:{imgPart.ContentType ?? "image/png"};base64,{base64}";
        }
        catch
        {
            return null;
        }
    }

    // ==================== Utility ====================

    private static double EmuToCm(long emu)
    {
        return Math.Round(emu / 360000.0, 3);
    }

    private static string HtmlEncode(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }

    /// <summary>
    /// Sanitize a value for use inside a CSS style attribute.
    /// Strips characters that could break out of the style context.
    /// </summary>
    private static string CssSanitize(string value)
    {
        // Remove characters that could escape the style attribute or inject HTML
        return value.Replace("\"", "").Replace("'", "").Replace("<", "").Replace(">", "")
            .Replace(";", "").Replace("{", "").Replace("}", "");
    }

    /// <summary>
    /// Sanitize a color value for safe embedding in CSS.
    /// Only allows hex colors (#RRGGBB), rgb/rgba() functions, and named CSS colors.
    /// </summary>
    private static string CssSanitizeColor(string color)
    {
        if (string.IsNullOrEmpty(color)) return "transparent";
        // Allow: #hex, rgb(), rgba(), named colors (alphanumeric only)
        var trimmed = color.Trim();
        if (trimmed.StartsWith('#') && trimmed.Length <= 9 && trimmed[1..].All(char.IsAsciiHexDigit))
            return trimmed;
        if (trimmed.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
            return CssSanitize(trimmed);
        if (trimmed.All(c => char.IsLetterOrDigit(c) || c == '.'))
            return trimmed;
        return "transparent";
    }

    /// <summary>
    /// Sanitize a MIME content type for safe embedding in a data URI.
    /// </summary>
    private static string SanitizeContentType(string contentType)
    {
        if (string.IsNullOrEmpty(contentType)) return "image/png";
        // Only allow alphanumeric, '/', '+', '-', '.'
        if (contentType.All(c => char.IsLetterOrDigit(c) || c is '/' or '+' or '-' or '.'))
            return contentType;
        return "image/png";
    }
}
