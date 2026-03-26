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
    // EMU to pixel conversion: 1 inch = 914400 EMU = 192 px (2x 96 DPI for retina)
    // So 1 px = 914400 / 192 = 4762.5 EMU
    // But to match officeshot's 1920×1080 from standard 10"×7.5" slides:
    //   10 inches * 914400 = 9144000 EMU → 1920 px → 1 px = 4762.5 EMU
    // Standard 13.333" × 7.5" (widescreen): 12192000 × 6858000 EMU → 1920 × 1080
    //   1 px = 12192000 / 1920 = 6350 EMU
    private const double EmuPerPx = 6350.0;

    private static double EmuToPx(long emu) => Math.Round(emu / EmuPerPx, 2);
    private static double EmuToPx(double emu) => Math.Round(emu / EmuPerPx, 2);

    /// <summary>
    /// Generate a self-contained native SVG for a single slide.
    /// ViewBox uses pixel coordinates (matching officeshot 1920×1080 output).
    /// </summary>
    public string ViewAsSvg(int slideNum)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideNum < 1 || slideNum > slideParts.Count)
            throw new CliException($"Slide {slideNum} does not exist. This presentation has {slideParts.Count} slide(s).")
            {
                Code = "out_of_range",
                Suggestion = $"Use a slide number between 1 and {slideParts.Count}."
            };

        var slidePart = slideParts[slideNum - 1];
        var (slideWidthEmu, slideHeightEmu) = GetSlideSize();
        var themeColors = ResolveThemeColorMap();

        double svgW = EmuToPx(slideWidthEmu);
        double svgH = EmuToPx(slideHeightEmu);

        var sb = new StringBuilder();
        var defsBuilder = new StringBuilder();
        int defIdCounter = 0;

        sb.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"");
        sb.AppendLine($"     width=\"{svgW:0.##}\" height=\"{svgH:0.##}\"");
        sb.AppendLine($"     viewBox=\"0 0 {svgW:0.##} {svgH:0.##}\">");

        const string defsPlaceholder = "<!--DEFS_PLACEHOLDER-->";
        sb.AppendLine(defsPlaceholder);

        // Slide background
        var bgColor = GetSlideBackgroundSvgColor(slidePart, themeColors);
        sb.AppendLine($"<rect width=\"{svgW:0.##}\" height=\"{svgH:0.##}\" fill=\"{bgColor}\"/>");

        // Render layout/master placeholders
        RenderLayoutPlaceholdersSvg(sb, defsBuilder, ref defIdCounter, slidePart, themeColors);

        // Render slide elements
        RenderSlideElementsSvg(sb, defsBuilder, ref defIdCounter, slidePart, slideNum, slideWidthEmu, slideHeightEmu, themeColors);

        sb.AppendLine("</svg>");

        // Insert accumulated defs
        var result = sb.ToString();
        var defsContent = defsBuilder.ToString();
        if (!string.IsNullOrEmpty(defsContent))
            result = result.Replace(defsPlaceholder, $"<defs>\n{defsContent}</defs>");
        else
            result = result.Replace(defsPlaceholder, "");

        return result;
    }

    private string GetSlideBackgroundSvgColor(SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        var bg = GetSlide(slidePart).CommonSlideData?.Background;
        if (bg != null)
        {
            var bgPr = bg.BackgroundProperties;
            if (bgPr != null)
            {
                var solidFill = bgPr.GetFirstChild<Drawing.SolidFill>();
                var color = ResolveFillColor(solidFill, themeColors);
                if (color != null) return color;
            }
        }
        return "white";
    }

    private void RenderSlideElementsSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        SlidePart slidePart, int slideNum, long slideWidthEmu, long slideHeightEmu,
        Dictionary<string, string> themeColors)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return;

        foreach (var element in shapeTree.ChildElements)
        {
            switch (element)
            {
                case Shape shape:
                    RenderShapeSvg(sb, defs, ref defId, shape, slidePart, themeColors);
                    break;
                case ConnectionShape cxn:
                    RenderConnectorSvg(sb, defs, ref defId, cxn, themeColors);
                    break;
                case Picture pic:
                    RenderPictureSvg(sb, defs, ref defId, pic, slidePart, themeColors);
                    break;
                case GraphicFrame gf:
                    if (gf.Descendants<Drawing.Table>().Any())
                        RenderTableSvg(sb, defs, ref defId, gf, themeColors);
                    break;
                case GroupShape grp:
                    RenderGroupSvg(sb, defs, ref defId, grp, slidePart, themeColors);
                    break;
                // TODO: Chart
            }
        }
    }

    private void RenderLayoutPlaceholdersSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        SlidePart slidePart, Dictionary<string, string> themeColors)
    {
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

        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart != null)
            RenderInheritedShapesSvg(sb, defs, ref defId, layoutPart.SlideLayout?.CommonSlideData?.ShapeTree,
                layoutPart, slidePlaceholders, themeColors);

        var masterPart = layoutPart?.SlideMasterPart;
        if (masterPart != null)
            RenderInheritedShapesSvg(sb, defs, ref defId, masterPart.SlideMaster?.CommonSlideData?.ShapeTree,
                masterPart, slidePlaceholders, themeColors);
    }

    private void RenderInheritedShapesSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        ShapeTree? shapeTree, OpenXmlPart part, HashSet<string> skipIndices,
        Dictionary<string, string> themeColors)
    {
        if (shapeTree == null) return;

        foreach (var element in shapeTree.ChildElements)
        {
            if (element is not Shape shape) continue;

            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph != null)
            {
                if (ph.Index?.HasValue == true && skipIndices.Contains($"idx:{ph.Index.Value}")) continue;
                if (ph.Type?.HasValue == true && skipIndices.Contains($"type:{ph.Type.InnerText}")) continue;
                if (string.IsNullOrWhiteSpace(GetShapeText(shape))) continue;
            }
            else
            {
                if (shape.ShapeProperties?.Transform2D == null) continue;
            }

            RenderShapeSvg(sb, defs, ref defId, shape, part, themeColors);
        }
    }

    // ==================== Shape Rendering (SVG) ====================

    private static void RenderShapeSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        Shape shape, OpenXmlPart part, Dictionary<string, string> themeColors,
        (long x, long y, long cx, long cy)? overridePos = null)
    {
        var xfrm = shape.ShapeProperties?.Transform2D;

        long xEmu, yEmu, cxEmu, cyEmu;
        if (overridePos != null)
        {
            (xEmu, yEmu, cxEmu, cyEmu) = overridePos.Value;
        }
        else if (xfrm?.Offset != null && xfrm?.Extents != null)
        {
            xEmu = xfrm.Offset.X?.Value ?? 0;
            yEmu = xfrm.Offset.Y?.Value ?? 0;
            cxEmu = xfrm.Extents.Cx?.Value ?? 0;
            cyEmu = xfrm.Extents.Cy?.Value ?? 0;
        }
        else
        {
            var resolved = ResolveInheritedPosition(shape, part);
            if (resolved == null)
            {
                if (string.IsNullOrWhiteSpace(GetShapeText(shape))) return;
                resolved = GetDefaultPlaceholderPosition(shape, part);
                if (resolved == null) return;
            }
            (xEmu, yEmu, cxEmu, cyEmu) = resolved.Value;
        }

        if (cxEmu <= 0 || cyEmu <= 0) return;

        // Convert to px
        double x = EmuToPx(xEmu), y = EmuToPx(yEmu);
        double w = EmuToPx(cxEmu), h = EmuToPx(cyEmu);

        // Resolve fill
        var spPr = shape.ShapeProperties;
        string fillColor = "none";
        double fillOpacity = 1.0;
        string? gradientRef = null;
        ResolveSvgFillWithOpacity(spPr, part, themeColors, out fillColor, out fillOpacity, defs, ref defId, out gradientRef);

        // Resolve outline
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
        string strokeColor = "none";
        double strokeWidth = 0;
        double strokeOpacity = 1.0;
        string strokeDasharray = "";
        if (outline != null && outline.GetFirstChild<Drawing.NoFill>() == null)
        {
            var outlineColor = ResolveFillColor(outline.GetFirstChild<Drawing.SolidFill>(), themeColors) ?? "#000000";
            ParseSvgColor(outlineColor, out strokeColor, out strokeOpacity);
            strokeWidth = EmuToPx(outline.Width?.HasValue == true ? outline.Width.Value : 12700);
            var dash = outline.GetFirstChild<Drawing.PresetDash>();
            if (dash?.Val?.HasValue == true)
            {
                var sw = strokeWidth;
                strokeDasharray = dash.Val.InnerText switch
                {
                    "dash" or "lgDash" or "sysDash" => $"{sw * 3:0.##} {sw * 2:0.##}",
                    "dot" or "sysDot" => $"{sw:0.##} {sw:0.##}",
                    "dashDot" or "lgDashDot" or "sysDashDot" => $"{sw * 3:0.##} {sw:0.##} {sw:0.##} {sw:0.##}",
                    _ => ""
                };
            }
        }

        // Build transform
        var transforms = new List<string>();
        transforms.Add($"translate({x:0.##},{y:0.##})");

        if (xfrm?.Rotation != null && xfrm.Rotation.Value != 0)
        {
            var deg = xfrm.Rotation.Value / 60000.0;
            transforms.Add($"rotate({deg:0.##},{w / 2:0.##},{h / 2:0.##})");
        }

        if (xfrm?.HorizontalFlip?.Value == true && xfrm.VerticalFlip?.Value == true)
            transforms.Add($"translate({w:0.##},{h:0.##}) scale(-1,-1)");
        else if (xfrm?.HorizontalFlip?.Value == true)
            transforms.Add($"translate({w:0.##},0) scale(-1,1)");
        else if (xfrm?.VerticalFlip?.Value == true)
            transforms.Add($"translate(0,{h:0.##}) scale(1,-1)");

        // Shadow effect → SVG filter
        var effectList = spPr?.GetFirstChild<Drawing.EffectList>();
        string? filterRef = null;
        if (effectList != null)
        {
            var shadowFilter = BuildSvgShadowFilter(effectList, themeColors, ref defId, defs);
            if (shadowFilter != null)
                filterRef = shadowFilter;
        }

        var gAttrs = $"transform=\"{string.Join(" ", transforms)}\"";
        if (filterRef != null)
            gAttrs += $" filter=\"url(#{filterRef})\"";
        sb.Append($"<g {gAttrs}>");

        // Resolve preset geometry for corner radius
        var presetGeom = spPr?.GetFirstChild<Drawing.PresetGeometry>();
        string presetName = presetGeom?.Preset?.InnerText ?? "rect";
        double rx = 0, ry = 0;
        if (presetName == "flowChartTerminator" || presetName == "flowChartAlternateProcess")
        {
            // Stadium/capsule shape — max border radius
            rx = ry = Math.Min(w, h) / 2;
        }
        else if (presetName is "roundRect" or "round1Rect" or "round2SameRect" or "round2DiagRect")
        {
            var minSide = Math.Min(cxEmu, cyEmu);
            long avVal = 16667; // default 16.667%
            var avList = presetGeom?.GetFirstChild<Drawing.AdjustValueList>();
            var gd = avList?.GetFirstChild<Drawing.ShapeGuide>();
            if (gd?.Formula?.Value != null && gd.Formula.Value.StartsWith("val "))
            {
                if (long.TryParse(gd.Formula.Value.AsSpan(4), out var parsed))
                    avVal = parsed;
            }
            var radiusEmu = minSide * avVal / 100000;
            rx = ry = EmuToPx(radiusEmu);
        }

        // Common fill/stroke attributes
        var fillValue = gradientRef != null ? $"url(#{gradientRef})" : fillColor;
        var fillStrokeAttrs = new List<string> { $"fill=\"{fillValue}\"" };
        if (fillOpacity < 1.0)
            fillStrokeAttrs.Add($"fill-opacity=\"{fillOpacity:0.##}\"");
        if (strokeColor != "none")
        {
            fillStrokeAttrs.Add($"stroke=\"{strokeColor}\"");
            fillStrokeAttrs.Add($"stroke-width=\"{strokeWidth:0.##}\"");
            if (strokeOpacity < 1.0)
                fillStrokeAttrs.Add($"stroke-opacity=\"{strokeOpacity:0.##}\"");
            if (!string.IsNullOrEmpty(strokeDasharray))
                fillStrokeAttrs.Add($"stroke-dasharray=\"{strokeDasharray}\"");
        }
        var fsStr = string.Join(" ", fillStrokeAttrs);

        // Draw shape based on geometry type
        var polygonPoints = GetPresetPolygonPoints(presetName, w, h);
        if (presetName == "ellipse")
        {
            sb.Append($"<ellipse cx=\"{w / 2:0.##}\" cy=\"{h / 2:0.##}\" rx=\"{w / 2:0.##}\" ry=\"{h / 2:0.##}\" {fsStr}/>");
        }
        else if (polygonPoints != null)
        {
            sb.Append($"<polygon points=\"{polygonPoints}\" {fsStr}/>");
        }
        else
        {
            // rect / roundRect / other rect variants
            var rectExtra = "";
            if (rx > 0)
                rectExtra = $" rx=\"{rx:0.##}\" ry=\"{ry:0.##}\"";
            sb.Append($"<rect width=\"{w:0.##}\" height=\"{h:0.##}\"{rectExtra} {fsStr}/>");
        }

        // Text content
        if (shape.TextBody != null)
        {
            var bodyPr = shape.TextBody.Elements<Drawing.BodyProperties>().FirstOrDefault();
            double lIns = EmuToPx(bodyPr?.LeftInset?.Value ?? 91440);
            double tIns = EmuToPx(bodyPr?.TopInset?.Value ?? 45720);
            double rIns = EmuToPx(bodyPr?.RightInset?.Value ?? 91440);
            double bIns = EmuToPx(bodyPr?.BottomInset?.Value ?? 45720);

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

            int? phDefaultFontSize = ResolvePlaceholderFontSize(shape, part);
            RenderTextBodySvg(sb, shape.TextBody, themeColors, w, h,
                lIns, tIns, rIns, bIns, valign, phDefaultFontSize);
        }

        sb.AppendLine("</g>");
    }

    // ==================== Fill Resolution (SVG) ====================

    /// <summary>
    /// Resolve fill color for SVG, separating color and opacity.
    /// Also handles gradient fills by creating SVG gradient definitions.
    /// </summary>
    private static void ResolveSvgFillWithOpacity(ShapeProperties? spPr, OpenXmlPart part,
        Dictionary<string, string> themeColors, out string color, out double opacity,
        StringBuilder defs, ref int defId, out string? gradientRef)
    {
        color = "none";
        opacity = 1.0;
        gradientRef = null;
        if (spPr == null) return;

        if (spPr.GetFirstChild<Drawing.NoFill>() != null)
            return;

        var solidFill = spPr.GetFirstChild<Drawing.SolidFill>();
        if (solidFill != null)
        {
            var resolved = ResolveFillColor(solidFill, themeColors);
            if (resolved != null)
            {
                ParseSvgColor(resolved, out color, out opacity);
                return;
            }
        }

        // Gradient fill
        var gradFill = spPr.GetFirstChild<Drawing.GradientFill>();
        if (gradFill != null)
        {
            gradientRef = BuildSvgGradient(gradFill, themeColors, ref defId, defs);
            if (gradientRef != null)
                return;
        }

        // TODO: blip fills (images)
    }

    /// <summary>
    /// Parse a CSS color (hex or rgba) into SVG-compatible color + opacity.
    /// </summary>
    private static void ParseSvgColor(string cssColor, out string svgColor, out double opacity)
    {
        opacity = 1.0;
        if (cssColor.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase))
        {
            // rgba(r,g,b,a)
            var inner = cssColor[5..^1];
            var parts = inner.Split(',');
            if (parts.Length >= 4)
            {
                var r = int.Parse(parts[0].Trim());
                var g = int.Parse(parts[1].Trim());
                var b = int.Parse(parts[2].Trim());
                opacity = double.Parse(parts[3].Trim(), System.Globalization.CultureInfo.InvariantCulture);
                svgColor = $"#{r:X2}{g:X2}{b:X2}";
                return;
            }
        }
        svgColor = cssColor;
    }

    // ==================== Text Rendering (SVG) ====================

    private static void RenderTextBodySvg(StringBuilder sb, OpenXmlElement textBody,
        Dictionary<string, string> themeColors,
        double shapeW, double shapeH,
        double lIns, double tIns, double rIns, double bIns,
        string valign, int? defaultFontSizeHundredths)
    {
        var paragraphs = textBody.Elements<Drawing.Paragraph>().ToList();
        if (paragraphs.Count == 0) return;

        double textW = shapeW - lIns - rIns;
        if (textW <= 0) return;

        // Gather paragraph info
        var paraInfos = new List<(Drawing.Paragraph para, double fontSizePt, string align)>();
        foreach (var para in paragraphs)
        {
            var firstRun = para.Elements<Drawing.Run>().FirstOrDefault();
            var rp = firstRun?.RunProperties;

            double fontSizePt = defaultFontSizeHundredths.HasValue ? defaultFontSizeHundredths.Value / 100.0 : 18;
            if (rp?.FontSize?.HasValue == true)
                fontSizePt = rp.FontSize.Value / 100.0;

            var align = "start";
            if (para.ParagraphProperties?.Alignment?.HasValue == true)
            {
                align = para.ParagraphProperties.Alignment.InnerText switch
                {
                    "ctr" => "middle",
                    "r" => "end",
                    _ => "start"
                };
            }

            paraInfos.Add((para, fontSizePt, align));
        }

        // Calculate total text height in px (pt → px: 1pt ≈ 1.333px at 96dpi)
        const double ptToPx = 96.0 / 72.0;
        double totalHeightPx = 0;
        foreach (var (_, fontSizePt, _) in paraInfos)
        {
            totalHeightPx += fontSizePt * ptToPx * 1.2; // 1.2 line height
        }

        // Vertical alignment
        double usableH = shapeH - tIns - bIns;
        double startY = valign switch
        {
            "center" => tIns + (usableH - totalHeightPx) / 2,
            "bottom" => tIns + usableH - totalHeightPx,
            _ => tIns
        };

        // Render each paragraph
        double currentY = startY;
        foreach (var (para, fontSizePt, align) in paraInfos)
        {
            double fontSizePx = fontSizePt * ptToPx;
            double lineHeightPx = fontSizePx * 1.2;
            double baselineY = currentY + fontSizePx * 0.85;

            double textAnchorX = align switch
            {
                "middle" => lIns + textW / 2.0,
                "end" => lIns + textW,
                _ => lIns
            };

            var runs = para.Elements<Drawing.Run>().ToList();
            if (runs.Count == 0)
            {
                currentY += lineHeightPx;
                continue;
            }

            sb.Append($"<text x=\"{textAnchorX:0.##}\" y=\"{baselineY:0.##}\" text-anchor=\"{align}\"");
            sb.Append($" font-size=\"{fontSizePx:0.##}\"");
            sb.Append($" font-family=\"Calibri, &apos;PingFang SC&apos;, &apos;Microsoft YaHei&apos;, sans-serif\"");
            sb.Append(">");

            foreach (var run in runs)
            {
                var text = run.Text?.Text ?? "";
                if (string.IsNullOrEmpty(text)) continue;

                var rp = run.RunProperties;
                var tspanAttrs = new List<string>();

                // Color
                var runFill = rp?.GetFirstChild<Drawing.SolidFill>();
                var runColorCss = ResolveFillColor(runFill, themeColors) ?? "#000000";
                ParseSvgColor(runColorCss, out var runColor, out var runOpacity);
                tspanAttrs.Add($"fill=\"{SvgEncode(runColor)}\"");
                if (runOpacity < 1.0)
                    tspanAttrs.Add($"fill-opacity=\"{runOpacity:0.##}\"");

                // Per-run font size
                if (rp?.FontSize?.HasValue == true)
                {
                    var runFontPx = rp.FontSize.Value / 100.0 * ptToPx;
                    tspanAttrs.Add($"font-size=\"{runFontPx:0.##}\"");
                }

                if (rp?.Bold?.Value == true)
                    tspanAttrs.Add("font-weight=\"bold\"");
                if (rp?.Italic?.Value == true)
                    tspanAttrs.Add("font-style=\"italic\"");

                var font = rp?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? rp?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                if (font != null && !font.StartsWith("+", StringComparison.Ordinal))
                    tspanAttrs.Add($"font-family=\"{SvgEncode(font)}\"");

                sb.Append($"<tspan {string.Join(" ", tspanAttrs)}>{SvgEncode(text)}</tspan>");
            }

            sb.Append("</text>");
            currentY += lineHeightPx;
        }
    }

    // ==================== Picture Rendering (SVG) ====================

    private static void RenderPictureSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        Picture pic, SlidePart slidePart, Dictionary<string, string> themeColors,
        (long x, long y, long cx, long cy)? overridePos = null)
    {
        var xfrm = pic.ShapeProperties?.Transform2D;
        if (xfrm?.Offset == null || xfrm?.Extents == null) return;

        double px = EmuToPx(overridePos?.x ?? xfrm.Offset.X?.Value ?? 0);
        double py = EmuToPx(overridePos?.y ?? xfrm.Offset.Y?.Value ?? 0);
        double pw = EmuToPx(overridePos?.cx ?? xfrm.Extents.Cx?.Value ?? 0);
        double ph = EmuToPx(overridePos?.cy ?? xfrm.Extents.Cy?.Value ?? 0);
        if (pw <= 0 || ph <= 0) return;

        // Extract image
        var blipFill = pic.BlipFill;
        var blip = blipFill?.GetFirstChild<Drawing.Blip>();
        if (blip?.Embed?.HasValue != true) return;

        string? dataUri = null;
        try
        {
            var imgPart = slidePart.GetPartById(blip.Embed.Value!);
            using var stream = imgPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());
            var contentType = SanitizeContentType(imgPart.ContentType ?? "image/png");
            dataUri = $"data:{contentType};base64,{base64}";
        }
        catch { return; }

        if (dataUri == null) return;

        // Transform
        var transforms = new List<string> { $"translate({px:0.##},{py:0.##})" };
        if (xfrm.Rotation != null && xfrm.Rotation.Value != 0)
            transforms.Add($"rotate({xfrm.Rotation.Value / 60000.0:0.##},{pw / 2:0.##},{ph / 2:0.##})");

        // Clip for crop
        string? clipId = null;
        var srcRect = blipFill?.GetFirstChild<Drawing.SourceRectangle>();
        if (srcRect != null)
        {
            var cl = (srcRect.Left?.Value ?? 0) / 100000.0;
            var ct = (srcRect.Top?.Value ?? 0) / 100000.0;
            var cr = (srcRect.Right?.Value ?? 0) / 100000.0;
            var cb = (srcRect.Bottom?.Value ?? 0) / 100000.0;
            if (cl != 0 || ct != 0 || cr != 0 || cb != 0)
            {
                clipId = $"clip{defId++}";
                defs.AppendLine($"<clipPath id=\"{clipId}\">");
                defs.AppendLine($"  <rect x=\"{pw * cl:0.##}\" y=\"{ph * ct:0.##}\" width=\"{pw * (1 - cl - cr):0.##}\" height=\"{ph * (1 - ct - cb):0.##}\"/>");
                defs.AppendLine("</clipPath>");
            }
        }

        sb.Append($"<g transform=\"{string.Join(" ", transforms)}\"");
        if (clipId != null) sb.Append($" clip-path=\"url(#{clipId})\"");
        sb.Append(">");
        sb.Append($"<image href=\"{dataUri}\" width=\"{pw:0.##}\" height=\"{ph:0.##}\" preserveAspectRatio=\"none\"/>");
        sb.AppendLine("</g>");
    }

    // ==================== Group Rendering (SVG) ====================

    private void RenderGroupSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        GroupShape grp, SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        var grpXfrm = grp.GroupShapeProperties?.TransformGroup;
        if (grpXfrm?.Offset == null || grpXfrm?.Extents == null) return;

        double gx = EmuToPx(grpXfrm.Offset.X?.Value ?? 0);
        double gy = EmuToPx(grpXfrm.Offset.Y?.Value ?? 0);
        long cx = grpXfrm.Extents.Cx?.Value ?? 0;
        long cy = grpXfrm.Extents.Cy?.Value ?? 0;

        var childOff = grpXfrm.ChildOffset;
        var childExt = grpXfrm.ChildExtents;
        var scaleX = (childExt?.Cx?.Value ?? cx) != 0 ? (double)cx / (childExt?.Cx?.Value ?? cx) : 1.0;
        var scaleY = (childExt?.Cy?.Value ?? cy) != 0 ? (double)cy / (childExt?.Cy?.Value ?? cy) : 1.0;
        var offX = childOff?.X?.Value ?? 0;
        var offY = childOff?.Y?.Value ?? 0;

        sb.Append($"<g transform=\"translate({gx:0.##},{gy:0.##})\">");

        foreach (var child in grp.ChildElements)
        {
            switch (child)
            {
                case Shape shape:
                {
                    var pos = CalcGroupChildPos(shape.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderShapeSvg(sb, defs, ref defId, shape, slidePart, themeColors, pos);
                    break;
                }
                case Picture pic:
                {
                    var pos = CalcGroupChildPos(pic.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderPictureSvg(sb, defs, ref defId, pic, slidePart, themeColors, pos);
                    break;
                }
                case ConnectionShape cxn:
                    RenderConnectorSvg(sb, defs, ref defId, cxn, themeColors);
                    break;
            }
        }

        sb.AppendLine("</g>");
    }

    // ==================== Connector Rendering (SVG) ====================

    private static void RenderConnectorSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        ConnectionShape cxn, Dictionary<string, string> themeColors)
    {
        var xfrm = cxn.ShapeProperties?.Transform2D;
        if (xfrm?.Offset == null || xfrm?.Extents == null) return;

        long xEmu = xfrm.Offset.X?.Value ?? 0;
        long yEmu = xfrm.Offset.Y?.Value ?? 0;
        long cxEmu = xfrm.Extents.Cx?.Value ?? 0;
        long cyEmu = xfrm.Extents.Cy?.Value ?? 0;
        var flipH = xfrm.HorizontalFlip?.Value == true;
        var flipV = xfrm.VerticalFlip?.Value == true;

        double px1 = EmuToPx(xEmu), py1 = EmuToPx(yEmu);
        double px2 = EmuToPx(xEmu + cxEmu), py2 = EmuToPx(yEmu + cyEmu);

        // Apply flips
        double lx1 = flipH ? px2 : px1, ly1 = flipV ? py2 : py1;
        double lx2 = flipH ? px1 : px2, ly2 = flipV ? py1 : py2;

        // Outline
        var outline = cxn.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        var defaultColor = themeColors.TryGetValue("tx1", out var txc) ? $"#{txc}"
            : themeColors.TryGetValue("dk1", out var dkc) ? $"#{dkc}" : "#000000";
        string strokeColor = defaultColor;
        double strokeOpacity = 1.0;
        double strokeWidth = 1.5; // px
        if (outline != null)
        {
            var c = ResolveFillColor(outline.GetFirstChild<Drawing.SolidFill>(), themeColors);
            if (c != null) ParseSvgColor(c, out strokeColor, out strokeOpacity);
            if (outline.Width?.HasValue == true) strokeWidth = EmuToPx(outline.Width.Value);
            if (strokeWidth < 0.5) strokeWidth = 0.5;
        }

        // Dash
        string dashAttr = "";
        var prstDash = outline?.GetFirstChild<Drawing.PresetDash>();
        if (prstDash?.Val?.HasValue == true)
        {
            var sw = strokeWidth;
            var dashArray = prstDash.Val.InnerText switch
            {
                "dash" or "lgDash" => $"{sw * 4:0.##},{sw * 3:0.##}",
                "dot" or "sysDot" => $"{sw:0.##},{sw * 2:0.##}",
                "dashDot" => $"{sw * 4:0.##},{sw * 2:0.##},{sw:0.##},{sw * 2:0.##}",
                _ => ""
            };
            if (!string.IsNullOrEmpty(dashArray))
                dashAttr = $" stroke-dasharray=\"{dashArray}\"";
        }

        // Arrow markers
        var headEnd = outline?.GetFirstChild<Drawing.HeadEnd>();
        var tailEnd = outline?.GetFirstChild<Drawing.TailEnd>();
        var hasTail = tailEnd?.Type?.HasValue == true && tailEnd.Type.InnerText != "none";
        var hasHead = headEnd?.Type?.HasValue == true && headEnd.Type.InnerText != "none";
        string markerStartAttr = "", markerEndAttr = "";

        if (hasTail)
        {
            var markerId = $"arrow{defId++}";
            var s = Math.Max(4, strokeWidth * 3);
            defs.AppendLine($"<marker id=\"{markerId}\" markerWidth=\"{s:0.#}\" markerHeight=\"{s:0.#}\" refX=\"0\" refY=\"{s / 2:0.#}\" orient=\"auto\">");
            defs.AppendLine($"  <polygon points=\"0,0 {s:0.#},{s / 2:0.#} 0,{s:0.#}\" fill=\"{strokeColor}\"/>");
            defs.AppendLine("</marker>");
            markerEndAttr = $" marker-end=\"url(#{markerId})\"";
        }
        if (hasHead)
        {
            var markerId = $"arrow{defId++}";
            var s = Math.Max(4, strokeWidth * 3);
            defs.AppendLine($"<marker id=\"{markerId}\" markerWidth=\"{s:0.#}\" markerHeight=\"{s:0.#}\" refX=\"{s:0.#}\" refY=\"{s / 2:0.#}\" orient=\"auto-start-reverse\">");
            defs.AppendLine($"  <polygon points=\"{s:0.#},0 0,{s / 2:0.#} {s:0.#},{s:0.#}\" fill=\"{strokeColor}\"/>");
            defs.AppendLine("</marker>");
            markerStartAttr = $" marker-start=\"url(#{markerId})\"";
        }

        var opacityAttr = strokeOpacity < 1.0 ? $" stroke-opacity=\"{strokeOpacity:0.##}\"" : "";
        sb.AppendLine($"<line x1=\"{lx1:0.##}\" y1=\"{ly1:0.##}\" x2=\"{lx2:0.##}\" y2=\"{ly2:0.##}\" stroke=\"{strokeColor}\" stroke-width=\"{strokeWidth:0.##}\"{opacityAttr}{dashAttr}{markerStartAttr}{markerEndAttr}/>");
    }

    // ==================== Table Rendering (SVG) ====================

    private static void RenderTableSvg(StringBuilder sb, StringBuilder defs, ref int defId,
        GraphicFrame gf, Dictionary<string, string> themeColors)
    {
        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
        if (table == null) return;

        var offset = gf.Transform?.Offset;
        var extents = gf.Transform?.Extents;
        if (offset == null || extents == null) return;

        double tx = EmuToPx(offset.X?.Value ?? 0);
        double ty = EmuToPx(offset.Y?.Value ?? 0);
        double tw = EmuToPx(extents.Cx?.Value ?? 0);
        double th = EmuToPx(extents.Cy?.Value ?? 0);

        // Column widths
        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
        long totalColWidth = gridCols?.Sum(gc => gc.Width?.Value ?? 0) ?? 0;
        var colWidths = new List<double>();
        if (gridCols != null && totalColWidth > 0)
        {
            foreach (var gc in gridCols)
                colWidths.Add(tw * (gc.Width?.Value ?? 0) / totalColWidth);
        }

        sb.Append($"<g transform=\"translate({tx:0.##},{ty:0.##})\">");

        // Table border
        sb.Append($"<rect width=\"{tw:0.##}\" height=\"{th:0.##}\" fill=\"none\" stroke=\"#D0D0D0\" stroke-width=\"0.5\"/>");

        double currentY = 0;
        int rowIndex = 0;
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            double rowH = EmuToPx(row.Height?.Value ?? 0);
            double currentX = 0;
            int colIndex = 0;

            foreach (var cell in row.Elements<Drawing.TableCell>())
            {
                double cellW = colIndex < colWidths.Count ? colWidths[colIndex] : tw / Math.Max(1, colWidths.Count);

                // Cell fill
                var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                var cellFill = ResolveFillColor(tcPr?.GetFirstChild<Drawing.SolidFill>(), themeColors);
                string cellFillColor = "none";
                double cellFillOpacity = 1.0;
                if (cellFill != null)
                    ParseSvgColor(cellFill, out cellFillColor, out cellFillOpacity);

                // Cell rect
                if (cellFillColor != "none")
                {
                    var opAttr = cellFillOpacity < 1.0 ? $" fill-opacity=\"{cellFillOpacity:0.##}\"" : "";
                    sb.Append($"<rect x=\"{currentX:0.##}\" y=\"{currentY:0.##}\" width=\"{cellW:0.##}\" height=\"{rowH:0.##}\" fill=\"{cellFillColor}\"{opAttr}/>");
                }

                // Cell border
                sb.Append($"<rect x=\"{currentX:0.##}\" y=\"{currentY:0.##}\" width=\"{cellW:0.##}\" height=\"{rowH:0.##}\" fill=\"none\" stroke=\"#BFBFBF\" stroke-width=\"0.5\"/>");

                // Cell text
                var textBody = cell.GetFirstChild<Drawing.TextBody>();
                if (textBody != null)
                {
                    double padding = 4;
                    RenderTextBodySvg(sb, textBody, themeColors, cellW, rowH,
                        padding, padding, padding, padding, "top", null);
                }

                currentX += cellW;
                colIndex++;
            }
            currentY += rowH;
            rowIndex++;
        }

        sb.AppendLine("</g>");
    }

    // ==================== SVG Preset Geometries ====================

    /// <summary>
    /// Returns SVG polygon points string for common preset shapes, or null if not a polygon shape.
    /// </summary>
    private static string? GetPresetPolygonPoints(string preset, double w, double h)
    {
        return preset switch
        {
            // Triangles
            "triangle" or "isosTriangle" => $"{w / 2:0.##},0 0,{h:0.##} {w:0.##},{h:0.##}",
            "rtTriangle" => $"0,0 0,{h:0.##} {w:0.##},{h:0.##}",

            // Diamond
            "diamond" => $"{w / 2:0.##},0 {w:0.##},{h / 2:0.##} {w / 2:0.##},{h:0.##} 0,{h / 2:0.##}",

            // Parallelogram
            "parallelogram" => $"{w * 0.25:0.##},0 {w:0.##},0 {w * 0.75:0.##},{h:0.##} 0,{h:0.##}",
            "trapezoid" => $"{w * 0.2:0.##},0 {w * 0.8:0.##},0 {w:0.##},{h:0.##} 0,{h:0.##}",

            // Pentagon, Hexagon, etc.
            "pentagon" => BuildRegularPolygon(5, w, h),
            "hexagon" => BuildRegularPolygon(6, w, h),
            "heptagon" => BuildRegularPolygon(7, w, h),
            "octagon" => BuildRegularPolygon(8, w, h),
            "decagon" => BuildRegularPolygon(10, w, h),
            "dodecagon" => BuildRegularPolygon(12, w, h),

            // Stars
            "star4" => BuildStar(4, w, h),
            "star5" => BuildStar(5, w, h),
            "star6" => BuildStar(6, w, h),
            "star8" => BuildStar(8, w, h),
            "star10" => BuildStar(10, w, h),
            "star12" => BuildStar(12, w, h),

            // Arrows
            "rightArrow" => $"0,{h * 0.25:0.##} {w * 0.7:0.##},{h * 0.25:0.##} {w * 0.7:0.##},0 {w:0.##},{h / 2:0.##} {w * 0.7:0.##},{h:0.##} {w * 0.7:0.##},{h * 0.75:0.##} 0,{h * 0.75:0.##}",
            "leftArrow" => $"{w:0.##},{h * 0.25:0.##} {w * 0.3:0.##},{h * 0.25:0.##} {w * 0.3:0.##},0 0,{h / 2:0.##} {w * 0.3:0.##},{h:0.##} {w * 0.3:0.##},{h * 0.75:0.##} {w:0.##},{h * 0.75:0.##}",
            "upArrow" => $"{w * 0.25:0.##},{h:0.##} {w * 0.25:0.##},{h * 0.3:0.##} 0,{h * 0.3:0.##} {w / 2:0.##},0 {w:0.##},{h * 0.3:0.##} {w * 0.75:0.##},{h * 0.3:0.##} {w * 0.75:0.##},{h:0.##}",
            "downArrow" => $"{w * 0.25:0.##},0 {w * 0.75:0.##},0 {w * 0.75:0.##},{h * 0.7:0.##} {w:0.##},{h * 0.7:0.##} {w / 2:0.##},{h:0.##} 0,{h * 0.7:0.##} {w * 0.25:0.##},{h * 0.7:0.##}",

            // Chevron
            "chevron" => $"0,0 {w * 0.8:0.##},0 {w:0.##},{h / 2:0.##} {w * 0.8:0.##},{h:0.##} 0,{h:0.##} {w * 0.2:0.##},{h / 2:0.##}",
            "homePlate" => $"0,0 {w * 0.85:0.##},0 {w:0.##},{h / 2:0.##} {w * 0.85:0.##},{h:0.##} 0,{h:0.##}",

            // Cross / Plus
            "plus" or "cross" => $"{w * 0.33:0.##},0 {w * 0.67:0.##},0 {w * 0.67:0.##},{h * 0.33:0.##} {w:0.##},{h * 0.33:0.##} {w:0.##},{h * 0.67:0.##} {w * 0.67:0.##},{h * 0.67:0.##} {w * 0.67:0.##},{h:0.##} {w * 0.33:0.##},{h:0.##} {w * 0.33:0.##},{h * 0.67:0.##} 0,{h * 0.67:0.##} 0,{h * 0.33:0.##} {w * 0.33:0.##},{h * 0.33:0.##}",

            // Heart (approximate with polygon)
            "heart" => BuildHeartPath(w, h),

            // Flowchart shapes
            "flowChartProcess" => null, // rect, handled by default
            "flowChartDecision" => $"{w / 2:0.##},0 {w:0.##},{h / 2:0.##} {w / 2:0.##},{h:0.##} 0,{h / 2:0.##}",
            "flowChartInputOutput" or "flowChartData" => $"{w * 0.2:0.##},0 {w:0.##},0 {w * 0.8:0.##},{h:0.##} 0,{h:0.##}",
            "flowChartManualInput" => $"0,{h * 0.15:0.##} {w:0.##},0 {w:0.##},{h:0.##} 0,{h:0.##}",
            "flowChartManualOperation" => $"0,0 {w:0.##},0 {w * 0.85:0.##},{h:0.##} {w * 0.15:0.##},{h:0.##}",
            "flowChartPreparation" => $"{w * 0.15:0.##},0 {w * 0.85:0.##},0 {w:0.##},{h / 2:0.##} {w * 0.85:0.##},{h:0.##} {w * 0.15:0.##},{h:0.##} 0,{h / 2:0.##}",
            "flowChartExtract" => $"{w / 2:0.##},0 {w:0.##},{h:0.##} 0,{h:0.##}",
            "flowChartMerge" => $"0,0 {w:0.##},0 {w / 2:0.##},{h:0.##}",
            "flowChartDocument" => $"0,0 {w:0.##},0 {w:0.##},{h * 0.85:0.##} {w * 0.75:0.##},{h * 0.75:0.##} {w * 0.5:0.##},{h * 0.9:0.##} {w * 0.25:0.##},{h:0.##} 0,{h * 0.85:0.##}",

            // Snip rectangles
            "snip1Rect" => $"0,0 {w * 0.92:0.##},0 {w:0.##},{h * 0.08:0.##} {w:0.##},{h:0.##} 0,{h:0.##}",
            "snip2SameRect" => $"{w * 0.08:0.##},0 {w * 0.92:0.##},0 {w:0.##},{h * 0.08:0.##} {w:0.##},{h:0.##} 0,{h:0.##} 0,{h * 0.08:0.##}",

            // Cloud / callout - approximate with polygon
            "cloud" or "cloudCallout" => BuildCloudPath(w, h),

            _ => null
        };
    }

    private static string BuildRegularPolygon(int sides, double w, double h)
    {
        var points = new List<string>();
        for (int i = 0; i < sides; i++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * i / sides;
            var px = w / 2 + w / 2 * Math.Cos(angle);
            var py = h / 2 + h / 2 * Math.Sin(angle);
            points.Add($"{px:0.##},{py:0.##}");
        }
        return string.Join(" ", points);
    }

    private static string BuildStar(int pointCount, double w, double h)
    {
        var points = new List<string>();
        var outerR = Math.Min(w, h) / 2;
        var innerR = outerR * 0.4;
        for (int i = 0; i < pointCount * 2; i++)
        {
            var angle = -Math.PI / 2 + Math.PI * i / pointCount;
            var r = i % 2 == 0 ? outerR : innerR;
            var px = w / 2 + r * Math.Cos(angle) * (w / Math.Min(w, h));
            var py = h / 2 + r * Math.Sin(angle) * (h / Math.Min(w, h));
            points.Add($"{px:0.##},{py:0.##}");
        }
        return string.Join(" ", points);
    }

    private static string BuildHeartPath(double w, double h)
    {
        // Approximate heart as polygon
        var points = new List<string>();
        int n = 32;
        for (int i = 0; i <= n; i++)
        {
            var t = 2 * Math.PI * i / n;
            var hx = 16 * Math.Pow(Math.Sin(t), 3);
            var hy = -(13 * Math.Cos(t) - 5 * Math.Cos(2 * t) - 2 * Math.Cos(3 * t) - Math.Cos(4 * t));
            var px = w / 2 + hx / 16 * w / 2;
            var py = h * 0.45 + hy / 17 * h / 2;
            points.Add($"{px:0.##},{py:0.##}");
        }
        return string.Join(" ", points);
    }

    private static string BuildCloudPath(double w, double h)
    {
        // Cloud shape approximated with overlapping circles as polygon
        var points = new List<string>();
        // Bottom arc
        AddArcPoints(points, w * 0.5, h * 0.85, w * 0.45, h * 0.2, Math.PI * 0.0, Math.PI * 1.0, 16);
        // Left arc
        AddArcPoints(points, w * 0.15, h * 0.55, w * 0.18, h * 0.35, Math.PI * 0.7, Math.PI * 1.5, 10);
        // Top-left arc
        AddArcPoints(points, w * 0.3, h * 0.25, w * 0.2, h * 0.22, Math.PI * 1.0, Math.PI * 1.8, 10);
        // Top arc
        AddArcPoints(points, w * 0.55, h * 0.18, w * 0.22, h * 0.2, Math.PI * 1.2, Math.PI * 2.0, 10);
        // Top-right arc
        AddArcPoints(points, w * 0.75, h * 0.28, w * 0.2, h * 0.25, Math.PI * 1.5, Math.PI * 2.2, 10);
        // Right arc
        AddArcPoints(points, w * 0.85, h * 0.55, w * 0.18, h * 0.35, Math.PI * 1.5, Math.PI * 2.3, 10);
        return string.Join(" ", points);
    }

    private static void AddArcPoints(List<string> points, double cx, double cy,
        double rx, double ry, double startAngle, double endAngle, int segments)
    {
        for (int i = 0; i <= segments; i++)
        {
            var angle = startAngle + (endAngle - startAngle) * i / segments;
            var px = cx + rx * Math.Cos(angle);
            var py = cy + ry * Math.Sin(angle);
            points.Add($"{px:0.##},{py:0.##}");
        }
    }

    // ==================== SVG Gradient ====================

    private static string? BuildSvgGradient(Drawing.GradientFill gradFill,
        Dictionary<string, string> themeColors, ref int defId, StringBuilder defs)
    {
        var stops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stops == null || stops.Count < 2) return null;

        // Build stop elements
        var stopElements = new List<string>();
        foreach (var gs in stops)
        {
            string stopColor = "#000000";
            double stopOpacity = 1.0;

            var color = ResolveFillColor(gs.GetFirstChild<Drawing.SolidFill>(), themeColors);
            if (color == null)
            {
                var rgb = gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (rgb != null && rgb.Length >= 6)
                {
                    color = $"#{rgb[..6]}";
                    var alpha = gs.GetFirstChild<Drawing.RgbColorModelHex>()?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
                    if (alpha.HasValue) stopOpacity = alpha.Value / 100000.0;
                }
                else
                {
                    var scheme = gs.GetFirstChild<Drawing.SchemeColor>();
                    if (scheme?.Val?.InnerText != null && themeColors.TryGetValue(scheme.Val.InnerText, out var tc))
                    {
                        color = $"#{ApplyColorTransforms(tc, scheme)}".Replace("rgba(", "").Replace(")", "");
                        // Re-resolve properly
                        var resolved = ApplyColorTransforms(tc, scheme);
                        ParseSvgColor(resolved, out color, out stopOpacity);
                    }
                }
            }
            else
            {
                ParseSvgColor(color, out stopColor, out stopOpacity);
                color = stopColor;
            }

            var pos = gs.Position?.Value;
            var offset = pos.HasValue ? $"{pos.Value / 1000.0:0.##}%" : "";
            var opacityAttr = stopOpacity < 1.0 ? $" stop-opacity=\"{stopOpacity:0.##}\"" : "";
            stopElements.Add($"  <stop offset=\"{offset}\" stop-color=\"{color}\"{opacityAttr}/>");
        }

        var gradId = $"grad{defId++}";

        // Radial or linear?
        var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
        if (pathGrad != null)
        {
            defs.AppendLine($"<radialGradient id=\"{gradId}\">");
            foreach (var s in stopElements) defs.AppendLine(s);
            defs.AppendLine("</radialGradient>");
        }
        else
        {
            var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
            var angleDeg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 90.0;
            // OOXML angle: 0=right, 90=bottom. Convert to SVG gradient coordinates.
            var angleRad = (angleDeg + 90) * Math.PI / 180;
            var x1 = 50 - 50 * Math.Cos(angleRad);
            var y1 = 50 - 50 * Math.Sin(angleRad);
            var x2 = 50 + 50 * Math.Cos(angleRad);
            var y2 = 50 + 50 * Math.Sin(angleRad);

            defs.AppendLine($"<linearGradient id=\"{gradId}\" x1=\"{x1:0.##}%\" y1=\"{y1:0.##}%\" x2=\"{x2:0.##}%\" y2=\"{y2:0.##}%\">");
            foreach (var s in stopElements) defs.AppendLine(s);
            defs.AppendLine("</linearGradient>");
        }

        return gradId;
    }

    // ==================== SVG Effects ====================

    private static string? BuildSvgShadowFilter(Drawing.EffectList effectList,
        Dictionary<string, string> themeColors, ref int defId, StringBuilder defs)
    {
        var shadow = effectList.GetFirstChild<Drawing.OuterShadow>();
        if (shadow == null) return null;

        var alpha = shadow.Descendants<Drawing.Alpha>().FirstOrDefault()?.Val?.Value ?? 50000;
        var opacity = alpha / 100000.0;

        var rgb = shadow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        int r = 0, g = 0, b = 0;
        if (rgb != null && rgb.Length >= 6)
        {
            r = Convert.ToInt32(rgb[..2], 16);
            g = Convert.ToInt32(rgb[2..4], 16);
            b = Convert.ToInt32(rgb[4..6], 16);
        }
        else
        {
            var schemeColor = shadow.GetFirstChild<Drawing.SchemeColor>()?.Val?.InnerText;
            if (schemeColor != null && themeColors.TryGetValue(schemeColor, out var sc) && sc.Length >= 6)
            {
                r = Convert.ToInt32(sc[..2], 16);
                g = Convert.ToInt32(sc[2..4], 16);
                b = Convert.ToInt32(sc[4..6], 16);
            }
        }

        var blurPx = EmuToPx(shadow.BlurRadius?.HasValue == true ? shadow.BlurRadius.Value : 50800);
        var distPx = EmuToPx(shadow.Distance?.HasValue == true ? shadow.Distance.Value : 38100);
        var angleDeg = shadow.Direction?.HasValue == true ? shadow.Direction.Value / 60000.0 : 45;
        var angleRad = angleDeg * Math.PI / 180;
        var dx = distPx * Math.Cos(angleRad);
        var dy = distPx * Math.Sin(angleRad);

        var filterId = $"shadow{defId++}";
        defs.AppendLine($"<filter id=\"{filterId}\" x=\"-20%\" y=\"-20%\" width=\"150%\" height=\"150%\">");
        defs.AppendLine($"  <feDropShadow dx=\"{dx:0.##}\" dy=\"{dy:0.##}\" stdDeviation=\"{blurPx / 2:0.##}\" flood-color=\"rgb({r},{g},{b})\" flood-opacity=\"{opacity:0.##}\"/>");
        defs.AppendLine("</filter>");

        return filterId;
    }

    // ==================== SVG Helpers ====================

    private static string SvgEncode(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }
}
