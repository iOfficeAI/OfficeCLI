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
    // ==================== Shape Rendering ====================

    /// <summary>
    /// Render a shape element to HTML. When called from a group, pass overridePos
    /// with the adjusted coordinates — the original element is NEVER modified.
    /// </summary>
    private static void RenderShape(StringBuilder sb, Shape shape, OpenXmlPart part,
        Dictionary<string, string> themeColors, (long x, long y, long cx, long cy)? overridePos = null)
    {
        var xfrm = shape.ShapeProperties?.Transform2D;

        long x, y, cx, cy;
        if (overridePos != null)
        {
            (x, y, cx, cy) = overridePos.Value;
        }
        else if (xfrm?.Offset != null && xfrm?.Extents != null)
        {
            x = xfrm.Offset.X?.Value ?? 0;
            y = xfrm.Offset.Y?.Value ?? 0;
            cx = xfrm.Extents.Cx?.Value ?? 0;
            cy = xfrm.Extents.Cy?.Value ?? 0;
        }
        else
        {
            // No xfrm — try to inherit position from matching layout/master placeholder
            var resolved = ResolveInheritedPosition(shape, part);
            if (resolved == null)
            {
                // No text content → skip silently
                if (string.IsNullOrWhiteSpace(GetShapeText(shape))) return;
                // Has text but no position can be resolved → use default placeholder position
                resolved = GetDefaultPlaceholderPosition(shape, part);
                if (resolved == null) return;
            }
            (x, y, cx, cy) = resolved.Value;
        }

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
        if (xfrm?.Rotation != null && xfrm.Rotation.Value != 0)
        {
            var deg = xfrm.Rotation.Value / 60000.0;
            transforms.Add($"rotate({deg:0.##}deg)");
        }

        // Flip
        if (xfrm?.HorizontalFlip?.Value == true && xfrm.VerticalFlip?.Value == true)
            transforms.Add("scale(-1,-1)");
        else if (xfrm?.HorizontalFlip?.Value == true)
            transforms.Add("scaleX(-1)");
        else if (xfrm?.VerticalFlip?.Value == true)
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
            var geomCss = PresetGeometryToCss(presetGeom.Preset!.InnerText!, cx, cy, presetGeom);
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

        // Reflection → CSS -webkit-box-reflect
        var reflectionCss = EffectListToReflectionCss(effectList);
        if (!string.IsNullOrEmpty(reflectionCss))
            styles.Add(reflectionCss);

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
        long lIns = bodyPr?.LeftInset?.Value ?? 91440;
        long tIns = bodyPr?.TopInset?.Value ?? 45720;
        long rIns = bodyPr?.RightInset?.Value ?? 91440;
        long bIns = bodyPr?.BottomInset?.Value ?? 45720;

        // For clip-path shapes (non-rectangular), add extra inner padding
        // so text doesn't appear outside the visible shape area.
        if (!string.IsNullOrEmpty(clipPathCss) && presetGeom?.Preset?.HasValue == true)
        {
            var (pctL, pctT, pctR, pctB) = GetShapeTextInsetPercent(presetGeom.Preset!.InnerText!);
            if (pctL > 0 || pctT > 0 || pctR > 0 || pctB > 0)
            {
                var extraL = (long)(cx * pctL);
                var extraT = (long)(cy * pctT);
                var extraR = (long)(cx * pctR);
                var extraB = (long)(cy * pctB);
                lIns = Math.Max(lIns, extraL);
                tIns = Math.Max(tIns, extraT);
                rIns = Math.Max(rIns, extraR);
                bIns = Math.Max(bIns, extraB);
            }
        }

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

            // Resolve placeholder-based default font size for inheritance
            int? phDefaultFontSize = ResolvePlaceholderFontSize(shape, part);

            RenderTextBody(sb, shape.TextBody, themeColors, phDefaultFontSize);
            sb.Append("</div>");
        }

        sb.AppendLine("</div>");
    }

    // ==================== Placeholder Position Inheritance ====================

    /// <summary>
    /// When a shape has no Transform2D, try to find position from matching placeholder
    /// on the slide layout or slide master (OOXML placeholder inheritance chain).
    /// </summary>
    private static (long x, long y, long cx, long cy)? ResolveInheritedPosition(Shape shape, OpenXmlPart part)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        // Only placeholder shapes can inherit position from layout/master
        if (ph == null) return null;

        var slidePart = part as SlidePart;
        if (slidePart == null) return null;

        // Search layout then master for a matching placeholder
        var layoutShapeTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
        var masterShapeTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;

        foreach (var tree in new[] { layoutShapeTree, masterShapeTree })
        {
            if (tree == null) continue;
            foreach (var candidate in tree.Elements<Shape>())
            {
                var candidatePh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                if (candidatePh == null) continue;

                if (!PlaceholderMatches(ph, candidatePh)) continue;

                var cxfrm = candidate.ShapeProperties?.Transform2D;
                if (cxfrm?.Offset != null && cxfrm?.Extents != null)
                {
                    return (
                        cxfrm.Offset.X?.Value ?? 0,
                        cxfrm.Offset.Y?.Value ?? 0,
                        cxfrm.Extents.Cx?.Value ?? 0,
                        cxfrm.Extents.Cy?.Value ?? 0
                    );
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Check if two placeholder shapes match by type and/or index.
    /// </summary>
    private static bool PlaceholderMatches(PlaceholderShape slidePh, PlaceholderShape layoutPh)
    {
        // Match by index first (most specific)
        if (slidePh.Index?.HasValue == true && layoutPh.Index?.HasValue == true)
            return slidePh.Index.Value == layoutPh.Index.Value;

        // Match by type
        if (slidePh.Type?.HasValue == true && layoutPh.Type?.HasValue == true)
            return slidePh.Type.Value == layoutPh.Type.Value;

        // If slide ph has no type/idx, match by name or consider it a body placeholder
        // Default placeholder type (when type is omitted) is "body" per OOXML spec
        if (slidePh.Type?.HasValue != true && slidePh.Index?.HasValue != true)
        {
            // A typeless/indexless placeholder matches title if the layout has title,
            // or body/subtitle by convention
            if (layoutPh.Type?.HasValue == true)
            {
                var lt = layoutPh.Type.Value;
                return lt == PlaceholderValues.Title || lt == PlaceholderValues.CenteredTitle
                    || lt == PlaceholderValues.SubTitle || lt == PlaceholderValues.Body;
            }
        }

        return false;
    }

    /// <summary>
    /// Last-resort fallback: provide default positions for placeholder shapes
    /// with text content when no layout/master placeholder can be matched.
    /// Uses standard PowerPoint default placeholder positions.
    /// </summary>
    private static (long x, long y, long cx, long cy)? GetDefaultPlaceholderPosition(Shape shape, OpenXmlPart part)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        // Get slide dimensions for proportional positioning
        long slideW = 12192000; // default 33.87cm
        long slideH = 6858000;  // default 19.05cm
        if (part is SlidePart sp)
        {
            var presDoc = sp.GetParentParts().OfType<PresentationPart>().FirstOrDefault();
            var slideSize = presDoc?.Presentation?.SlideSize;
            if (slideSize?.Cx?.HasValue == true) slideW = slideSize.Cx.Value;
            if (slideSize?.Cy?.HasValue == true) slideH = slideSize.Cy.Value;
        }

        // Standard PowerPoint default positions (in EMU)
        long margin = slideW / 16; // ~6.25% margin on each side
        long contentW = slideW - margin * 2;

        if (ph?.Type?.HasValue == true)
        {
            var t = ph.Type.Value;
            if (t == PlaceholderValues.Title || t == PlaceholderValues.CenteredTitle)
                return (margin, slideH / 8, contentW, slideH / 4);
            if (t == PlaceholderValues.SubTitle)
                return (margin, slideH * 3 / 8, contentW, slideH / 4);
            if (t == PlaceholderValues.Body || t == PlaceholderValues.Object)
                return (margin, slideH * 3 / 8, contentW, slideH / 2);
            return null;
        }

        // Placeholder with no type attribute — use a generous centered area
        if (ph != null)
        {
            // Determine position based on shape name as a hint
            // Check Subtitle before Title since "Subtitle" contains "Title"
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "";
            if (name.Contains("Subtitle", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("副标题", StringComparison.Ordinal))
                return (margin, slideH * 3 / 8, contentW, slideH / 4);
            if (name.Contains("Title", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("标题", StringComparison.Ordinal))
                return (margin, slideH / 8, contentW, slideH / 4);

            // Generic placeholder — use body area
            return (margin, slideH / 4, contentW, slideH / 2);
        }

        return null;
    }

    // ==================== Shape Text Inset for Clip-Path Shapes ====================

    /// <summary>
    /// Returns per-side inset percentages (left, top, right, bottom) for text inside a clip-path shape.
    /// Each value is 0-1, applied to the shape's width (left/right) or height (top/bottom).
    /// This keeps text within the visible shape interior.
    /// </summary>
    private static (double L, double T, double R, double B) GetShapeTextInsetPercent(string preset) => preset switch
    {
        "diamond" => (0.25, 0.25, 0.25, 0.25),
        "triangle" or "isosTriangle" => (0.20, 0.20, 0.20, 0),
        "rtTriangle" => (0, 0.15, 0.15, 0),
        "star4" => (0.28, 0.28, 0.28, 0.28),
        "star5" => (0.28, 0.28, 0.28, 0.28),
        "star6" => (0.25, 0.25, 0.25, 0.25),
        "star8" or "star10" or "star12" => (0.20, 0.20, 0.20, 0.20),
        "hexagon" => (0.25, 0.10, 0.25, 0.10),
        "pentagon" => (0.12, 0.12, 0.12, 0),
        "heptagon" or "octagon" or "decagon" or "dodecagon" => (0.08, 0.08, 0.08, 0.08),
        "parallelogram" => (0.15, 0, 0.15, 0),
        "trapezoid" => (0.12, 0, 0.12, 0),
        "rightArrow" or "notchedRightArrow" => (0, 0.20, 0.25, 0.20),
        "leftArrow" => (0.25, 0.20, 0, 0.20),
        "upArrow" => (0.20, 0.25, 0.20, 0),
        "downArrow" => (0.20, 0, 0.20, 0.25),
        "chevron" or "homePlate" => (0, 0, 0.15, 0),
        "heart" => (0.15, 0.15, 0.15, 0.15),
        "plus" or "cross" => (0.10, 0.10, 0.10, 0.10),
        "cloud" or "cloudCallout" => (0.12, 0.12, 0.12, 0.12),
        "sun" => (0.20, 0.20, 0.20, 0.20),
        "moon" => (0.15, 0, 0, 0),
        "cube" => (0, 0.08, 0.08, 0),
        "donut" => (0.25, 0.25, 0.25, 0.25),
        "wedgeRectCallout" or "wedgeRoundRectCallout" or "wedgeEllipseCallout" => (0.08, 0.08, 0.08, 0.08),
        "curvedRightArrow" or "curvedLeftArrow" or "curvedUpArrow" or "curvedDownArrow" => (0.12, 0.12, 0.12, 0.12),
        _ => (0, 0, 0, 0)
    };

    // ==================== Placeholder Font Size Inheritance ====================

    /// <summary>
    /// Resolve the default font size for a placeholder shape by walking the inheritance chain:
    /// shape listStyle → slide layout placeholder → slide master placeholder → master text styles → OOXML defaults.
    /// Returns font size in hundredths of a point (e.g. 4400 = 44pt), or null if no override.
    /// </summary>
    private static int? ResolvePlaceholderFontSize(Shape shape, OpenXmlPart part)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return null; // Not a placeholder

        // 1. Check shape's own list style for level 1 default run properties
        var lstStyle = shape.TextBody?.GetFirstChild<Drawing.ListStyle>();
        var lvl1 = lstStyle?.GetFirstChild<Drawing.Level1ParagraphProperties>();
        var defRp = lvl1?.GetFirstChild<Drawing.DefaultRunProperties>();
        if (defRp?.FontSize?.HasValue == true)
            return defRp.FontSize.Value;

        // Determine placeholder category
        var phType = ph.Type?.HasValue == true ? ph.Type.Value : PlaceholderValues.Body;
        bool isTitle = phType == PlaceholderValues.Title || phType == PlaceholderValues.CenteredTitle;
        bool isSubTitle = phType == PlaceholderValues.SubTitle;

        // 2. Check layout and master placeholder matching shapes for inherited font size
        if (part is SlidePart slidePart)
        {
            var layoutTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            var masterTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;

            foreach (var tree in new[] { layoutTree, masterTree })
            {
                if (tree == null) continue;
                foreach (var candidate in tree.Elements<Shape>())
                {
                    var cPh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    if (cPh == null) continue;
                    if (!PlaceholderMatches(ph, cPh)) continue;

                    // Check candidate's list style
                    var cLstStyle = candidate.TextBody?.GetFirstChild<Drawing.ListStyle>();
                    var cLvl1 = cLstStyle?.GetFirstChild<Drawing.Level1ParagraphProperties>();
                    var cDefRp = cLvl1?.GetFirstChild<Drawing.DefaultRunProperties>();
                    if (cDefRp?.FontSize?.HasValue == true)
                        return cDefRp.FontSize.Value;
                }
            }

            // 3. Check master text styles (titleStyle for titles, bodyStyle for body, otherStyle for others)
            var masterTxStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
            if (masterTxStyles != null)
            {
                OpenXmlCompositeElement? styleList = null;
                if (isTitle)
                    styleList = masterTxStyles.TitleStyle;
                else if (isSubTitle || phType == PlaceholderValues.Body || phType == PlaceholderValues.Object)
                    styleList = masterTxStyles.BodyStyle;
                else
                    styleList = masterTxStyles.OtherStyle;

                if (styleList != null)
                {
                    var sLvl1 = styleList.GetFirstChild<Drawing.Level1ParagraphProperties>();
                    var sDefRp = sLvl1?.GetFirstChild<Drawing.DefaultRunProperties>();
                    if (sDefRp?.FontSize?.HasValue == true)
                        return sDefRp.FontSize.Value;
                }
            }
        }

        // 4. OOXML spec defaults: Title=44pt, SubTitle=32pt, Body=24pt
        if (isTitle) return 4400;
        if (isSubTitle) return 3200;

        return null;
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

        // Reflection → CSS -webkit-box-reflect
        var reflectionCss = EffectListToReflectionCss(effectList);
        if (!string.IsNullOrEmpty(reflectionCss))
            styles.Add(reflectionCss);

        // Geometry (rounded corners)
        var presetGeom = pic.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
        {
            var geomCss = PresetGeometryToCss(presetGeom.Preset!.InnerText!, cx, cy, presetGeom);
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
                sb.Append("<div style=\"width:100%;height:100%;background:rgba(128,128,128,0.15);display:flex;align-items:center;justify-content:center;color:rgba(128,128,128,0.5);font-size:12px\">Image</div>");
            }
        }

        sb.AppendLine("</div>");
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
        var defaultLineColor = themeColors.TryGetValue("tx1", out var txc) ? $"#{txc}"
            : themeColors.TryGetValue("dk1", out var dkc) ? $"#{dkc}" : "#000000";
        var lineColor = defaultLineColor;
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

    // ==================== AlternateContent (3D Model, Zoom) Rendering ====================

    /// <summary>
    /// Render mc:AlternateContent elements. For 3D models, embeds the GLB as base64
    /// and uses Three.js to render it interactively in the browser.
    /// </summary>
    private static void RenderAlternateContent(StringBuilder sb, OpenXmlElement acElement,
        SlidePart slidePart, Dictionary<string, string> themeColors)
    {
        var isModel3D = acElement.Descendants().Any(d => d.LocalName == "model3d");
        var isZoom = acElement.Descendants().Any(d => d.LocalName == "sldZm");
        if (!isModel3D && !isZoom) return;

        // Extract position from mc:Choice > graphicFrame/sp > xfrm
        var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
        var frame = choice?.ChildElements.FirstOrDefault(e =>
            e.LocalName == "graphicFrame" || e.LocalName == "sp");
        var xfrm = frame?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
        xfrm ??= frame?.Descendants().FirstOrDefault(e =>
            e.LocalName == "xfrm" && e.Parent?.LocalName == (frame?.LocalName == "sp" ? "spPr" : frame?.LocalName));
        if (xfrm == null) return;

        var off = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
        var ext = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
        if (off == null || ext == null) return;

        long.TryParse(off.GetAttribute("x", "").Value, out var x);
        long.TryParse(off.GetAttribute("y", "").Value, out var y);
        long.TryParse(ext.GetAttribute("cx", "").Value, out var cx);
        long.TryParse(ext.GetAttribute("cy", "").Value, out var cy);
        if (cx == 0 || cy == 0) return;

        var leftCm = x / 360000.0;
        var topCm = y / 360000.0;
        var widthCm = cx / 360000.0;
        var heightCm = cy / 360000.0;

        if (isModel3D)
        {
            RenderModel3D(sb, acElement, slidePart, leftCm, topCm, widthCm, heightCm);
        }
        else
        {
            // Zoom: render fallback image
            RenderZoomFallback(sb, acElement, slidePart, leftCm, topCm, widthCm, heightCm);
        }
    }

    private static int _model3dCounter;
    // Cache: part URI → JS variable name, to avoid embedding the same GLB multiple times
    private static readonly Dictionary<string, string> _glbDataCache = new();

    /// <summary>
    /// Render a 3D model using Three.js with the embedded GLB data.
    /// Same GLB files across slides are deduplicated — embedded once, referenced by variable.
    /// </summary>
    private static void RenderModel3D(StringBuilder sb, OpenXmlElement acElement,
        SlidePart slidePart, double leftCm, double topCm, double widthCm, double heightCm)
    {
        // Find the model3d element and get the GLB relationship
        var model3d = acElement.Descendants().FirstOrDefault(d => d.LocalName == "model3d");
        if (model3d == null) return;

        var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var embedId = model3d.GetAttribute("embed", rNs).Value;
        if (string.IsNullOrEmpty(embedId)) return;

        // Deduplicate: use content hash so identical GLBs across slides share one copy
        string glbVarName;
        try
        {
            var part = slidePart.GetPartById(embedId);
            using var stream = part.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var bytes = ms.ToArray();
            var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes))[..16];
            if (!_glbDataCache.TryGetValue(hash, out glbVarName!))
            {
                glbVarName = $"_glb{_glbDataCache.Count}";
                sb.AppendLine($"<script>window.{glbVarName}='{Convert.ToBase64String(bytes)}';</script>");
                _glbDataCache[hash] = glbVarName;
            }
        }
        catch { return; }

        var canvasId = $"model3d_{_model3dCounter++}";

        // Extract rotation from am3d:rot
        var rot = model3d.Descendants().FirstOrDefault(d => d.LocalName == "rot");
        double rotX = 0, rotY = 0, rotZ = 0;
        if (rot != null)
        {
            var ax = rot.GetAttribute("ax", "").Value;
            var ay = rot.GetAttribute("ay", "").Value;
            var az = rot.GetAttribute("az", "").Value;
            if (!string.IsNullOrEmpty(ax) && int.TryParse(ax, out var axv)) rotX = axv / 60000.0 * Math.PI / 180.0;
            if (!string.IsNullOrEmpty(ay) && int.TryParse(ay, out var ayv)) rotY = ayv / 60000.0 * Math.PI / 180.0;
            if (!string.IsNullOrEmpty(az) && int.TryParse(az, out var azv)) rotZ = azv / 60000.0 * Math.PI / 180.0;
        }

        // Extract fallback image from mc:Fallback for WebGL-unavailable environments
        string? fallbackImgSrc = null;
        var fallback = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
        var fbBlip = fallback?.Descendants().FirstOrDefault(d => d.LocalName == "blip");
        if (fbBlip != null)
        {
            var fbRNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var fbEmbedId = fbBlip.GetAttribute("embed", fbRNs).Value;
            if (!string.IsNullOrEmpty(fbEmbedId))
            {
                try
                {
                    var fbPart = slidePart.GetPartById(fbEmbedId);
                    using var fbStream = fbPart.GetStream();
                    using var fbMs = new MemoryStream();
                    fbStream.CopyTo(fbMs);
                    var fbBytes = fbMs.ToArray();
                    if (fbBytes.Length > 200)
                        fallbackImgSrc = $"data:{fbPart.ContentType ?? "image/png"};base64,{Convert.ToBase64String(fbBytes)}";
                }
                catch { }
            }
        }

        var containerId = $"m3d_wrap_{canvasId}";
        sb.AppendLine($"    <div id=\"{containerId}\" style=\"position:absolute;" +
            $"left:{leftCm:F3}cm;top:{topCm:F3}cm;" +
            $"width:{widthCm:F3}cm;height:{heightCm:F3}cm;" +
            $"overflow:hidden;\">");
        sb.AppendLine($"      <canvas id=\"{canvasId}\" style=\"width:100%;height:100%;\"></canvas>");
        if (fallbackImgSrc != null)
            sb.AppendLine($"      <img class=\"m3d-fallback\" src=\"{fallbackImgSrc}\" style=\"width:100%;height:100%;object-fit:contain;display:none;\" />");
        sb.AppendLine("    </div>");

        sb.AppendLine($@"    <script type=""module"">
    import * as THREE from 'three';
    import {{ GLTFLoader }} from 'three/addons/loaders/GLTFLoader.js';
    (function() {{
      const canvas = document.getElementById('{canvasId}');
      if (!canvas) return;
      const container = canvas.parentElement;
      try {{
        const designW = {widthCm:F3} * 96 / 2.54;
        const designH = {heightCm:F3} * 96 / 2.54;
        canvas.width = designW * 2; canvas.height = designH * 2;
        canvas.style.width = '100%'; canvas.style.height = '100%';

        const w = designW, h = designH;
        const dpr = window.devicePixelRatio || 1;
        const renderer = new THREE.WebGLRenderer({{ canvas, alpha: true, antialias: true }});
        renderer.setSize(canvas.width / dpr, canvas.height / dpr);
        renderer.setPixelRatio(dpr);
        renderer.outputColorSpace = THREE.SRGBColorSpace;

        const scene = new THREE.Scene();
        const camera = new THREE.PerspectiveCamera(45, w / h, 0.01, 1000);

        // Lighting (matches PowerPoint 3-point setup)
        scene.add(new THREE.AmbientLight(0x808080, 0.8));
        const key = new THREE.DirectionalLight(0xfff0e0, 1.2);
        key.position.set(2, 3, 4);
        scene.add(key);
        const fill = new THREE.DirectionalLight(0x6090e0, 0.6);
        fill.position.set(-3, 2, -1);
        scene.add(fill);
        const rim = new THREE.DirectionalLight(0xd0b0ff, 0.4);
        rim.position.set(-1, 1, -3);
        scene.add(rim);

        // Load GLB from base64
        const b64 = window.{glbVarName};
        const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        const loader = new GLTFLoader();
        loader.parse(bin.buffer, '', (gltf) => {{
          const model = gltf.scene;
          // Center and fit model
          const box = new THREE.Box3().setFromObject(model);
          const center = box.getCenter(new THREE.Vector3());
          const size = box.getSize(new THREE.Vector3());
          model.position.sub(center);
          const maxDim = Math.max(size.x, size.y, size.z);
          const scale = 2.0 / maxDim;
          model.scale.setScalar(scale);
          // Apply rotation from PowerPoint
          model.rotation.x = {rotX:F6};
          model.rotation.y = {rotY:F6};
          model.rotation.z = {rotZ:F6};
          scene.add(model);
          // Position camera
          camera.position.set(0, 0, 3.2);
          camera.lookAt(0, 0, 0);
          // Auto-rotate animation
          let baseRotY = {rotY:F6};
          function animate() {{
            requestAnimationFrame(animate);
            baseRotY += 0.003;
            model.rotation.y = baseRotY;
            renderer.render(scene, camera);
          }}
          animate();
        }});
      }} catch(e) {{
        // WebGL unavailable — show fallback image
        canvas.style.display = 'none';
        const fb = container?.querySelector('.m3d-fallback');
        if (fb) fb.style.display = 'block';
      }}
    }})();
    </script>");
    }

    /// <summary>
    /// Render a zoom element using its fallback image.
    /// </summary>
    private static void RenderZoomFallback(StringBuilder sb, OpenXmlElement acElement,
        SlidePart slidePart, double leftCm, double topCm, double widthCm, double heightCm)
    {
        var fallback = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
        var fbBlip = fallback?.Descendants().FirstOrDefault(d => d.LocalName == "blip");
        string? imgSrc = null;
        if (fbBlip != null)
        {
            var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var embedId = fbBlip.GetAttribute("embed", rNs).Value;
            if (!string.IsNullOrEmpty(embedId))
            {
                try
                {
                    var part = slidePart.GetPartById(embedId);
                    using var stream = part.GetStream();
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    var bytes = ms.ToArray();
                    if (bytes.Length > 200)
                        imgSrc = $"data:{part.ContentType ?? "image/png"};base64,{Convert.ToBase64String(bytes)}";
                }
                catch { }
            }
        }

        sb.AppendLine($"    <div style=\"position:absolute;" +
            $"left:{leftCm:F3}cm;top:{topCm:F3}cm;" +
            $"width:{widthCm:F3}cm;height:{heightCm:F3}cm;" +
            $"border:2px dashed rgba(255,193,7,0.6);border-radius:8px;" +
            $"overflow:hidden;\">");
        if (imgSrc != null)
            sb.AppendLine($"      <img src=\"{imgSrc}\" style=\"width:100%;height:100%;object-fit:contain;\" />");
        sb.AppendLine("    </div>");
    }
}
