// Copyright 2025 OfficeCLI (officecli.ai)
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
        Dictionary<string, string> themeColors, (long x, long y, long cx, long cy)? overridePos = null,
        string? dataPath = null, bool suppressText = false)
    {
        // prst="line" auto-shapes are line-segment geometry; render as SVG
        // through the connector pipeline so they don't degrade to a div with
        // border (which fakes a thin filled rect and loses zero-width/height
        // line semantics — observed on slide 2 of test-samples/07.pptx).
        var prstGeomEarly = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (overridePos == null && prstGeomEarly?.Preset?.HasValue == true
            && prstGeomEarly.Preset.InnerText == "line")
        {
            RenderConnector(sb, shape.ShapeProperties, themeColors, dataPath);
            return;
        }

        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        var xfrm = shape.ShapeProperties?.Transform2D;

        // Shape-level hyperlink → wrap rendered shape <div> in <a> for clickability in HTML preview.
        // Only external URLs are wrapped; internal slide-jump links (ppaction://hlinksldjump) are
        // skipped because there is no corresponding external href in this static HTML context.
        string? shapeHrefUrl = null;
        string? shapeHrefTooltip = null;
        {
            var nvHlink = shape.NonVisualShapeProperties?.NonVisualDrawingProperties
                ?.GetFirstChild<Drawing.HyperlinkOnClick>();
            if (nvHlink != null)
            {
                shapeHrefTooltip = nvHlink.Tooltip?.Value;
                var action = nvHlink.Action?.Value;
                var hlId = nvHlink.Id?.Value;
                // Skip if this is a slide-jump action (no external URL target)
                if (string.IsNullOrEmpty(action) || !action.Contains("hlink"))
                {
                    // Plain external: no action + r:id → look up external relationship
                    if (!string.IsNullOrEmpty(hlId))
                    {
                        try
                        {
                            var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == hlId);
                            if (rel?.Uri != null) shapeHrefUrl = rel.Uri.ToString();
                        }
                        catch { }
                    }
                }
                else if (action.Contains("hlinksldjump"))
                {
                    // Internal slide-jump — deliberately not wrapped (no navigable href in static HTML)
                    shapeHrefUrl = null;
                }
            }
        }

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
            $"left:{Units.EmuToPt(x)}pt",
            $"top:{Units.EmuToPt(y)}pt",
            $"width:{Units.EmuToPt(cx)}pt",
            $"height:{Units.EmuToPt(cy)}pt"
        };

        // Fill
        var fillCss = GetShapeFillCss(shape.ShapeProperties, part, themeColors);
        if (!string.IsNullOrEmpty(fillCss))
            styles.Add(fillCss);

        // Border/outline — parse for later; solid goes to CSS, non-solid to SVG
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        var parsedOutline = outline != null ? ParseOutline(outline, themeColors) : null;
        if (parsedOutline != null && parsedOutline.Value.dashType == "solid")
        {
            styles.Add($"border:{parsedOutline.Value.widthPt:0.##}pt solid {parsedOutline.Value.color}");
        }
        // Non-solid outlines rendered as SVG after the shape div

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

        // Shadow + Glow → combine into single filter property
        var effectList = shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        var shadowCss = EffectListToShadowCss(effectList, themeColors);
        var glowCss = EffectListToGlowCss(effectList, themeColors);
        // Merge multiple filter:drop-shadow into one filter property
        var filterParts = new List<string>();
        if (!string.IsNullOrEmpty(shadowCss))
            filterParts.Add(shadowCss.Replace("filter:", ""));
        if (!string.IsNullOrEmpty(glowCss))
            filterParts.Add(glowCss.Replace("filter:", ""));
        if (filterParts.Count > 0)
            styles.Add($"filter:{string.Join(" ", filterParts)}");

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
            var bevelW = sp3d.BevelTop.Width?.HasValue == true ? sp3d.BevelTop.Width.Value / 12700.0 : 6; // OOXML default 76200 EMU = 6pt
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

        // For non-rectangular shapes (clip-path or border-radius), add extra inner padding
        // so text doesn't appear outside the visible shape area.
        if ((!string.IsNullOrEmpty(clipPathCss) || !string.IsNullOrEmpty(borderRadiusCss)) && presetGeom?.Preset?.HasValue == true)
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

        // Skip text-frame padding for shapes with no real text content. With
        // box-sizing:border-box, when default padding (~7.2pt L/R) exceeds the
        // shape's outer width, Chromium expands the rendered box to fit the
        // padding instead of clamping content to 0 — turning small decorative
        // shapes (e.g. 5.76pt vertex-marker ellipses) into wide pills.
        if (!string.IsNullOrWhiteSpace(GetShapeText(shape)))
            styles.Add($"padding:{Units.EmuToPt(tIns)}pt {Units.EmuToPt(rIns)}pt {Units.EmuToPt(bIns)}pt {Units.EmuToPt(lIns)}pt");

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

        // bodyPr/@wrap="none": text does not wrap inside the shape. Combined
        // with noAutofit (or by itself, since spAutoFit only adjusts the
        // shape's own bounds and normAutofit only scales), real PowerPoint
        // lets the rendered line extend beyond the shape's right edge.
        // Detect explicitly so we can suppress wrap and unclip overflow.
        var wrapNone = bodyPr?.Wrap?.Value == Drawing.TextWrappingValues.None;
        // noAutofit / spAutoFit (and absence of any autofit child = default
        // is shape-dependent, but textbox defaults to noAutofit) mean we must
        // not clip vertical overflow either.
        var noAutofit = bodyPr?.GetFirstChild<Drawing.NoAutoFit>() != null
            || (bodyPr != null
                && bodyPr.GetFirstChild<Drawing.NormalAutoFit>() == null
                && bodyPr.GetFirstChild<Drawing.ShapeAutoFit>() == null);

        // Add has-fill class to clip overflow when shape has a visible background.
        // wrap=none AND explicit noAutofit (<a:noAutofit/> child) both take
        // priority: real PowerPoint lets text overflow past the shape edges
        // in either mode (wrap=none = horizontal overflow past the right edge;
        // explicit noAutofit = vertical overflow past the bottom edge for
        // over-long body text). The earlier `noAutofit` local also fires when
        // no autofit child is present at all (textbox default); we deliberately
        // do NOT use that for the overflow decision, because doing so would
        // turn off clipping on every plain filled shape and let stray text
        // bleed across decorative buttons.
        var hasFillBg = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>() != null
            || shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null
            || shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>() != null;
        var explicitNoAutofit = bodyPr?.GetFirstChild<Drawing.NoAutoFit>() != null;
        var allowOverflow = wrapNone || explicitNoAutofit;
        var shapeClass = hasFillBg && !allowOverflow ? "shape has-fill" : "shape";
        if (allowOverflow) styles.Add("overflow:visible");

        // Open <a> wrapper for shape-level hyperlink (before the shape <div>)
        if (!string.IsNullOrEmpty(shapeHrefUrl))
        {
            var tooltipAttr = !string.IsNullOrEmpty(shapeHrefTooltip)
                ? $" title=\"{HtmlEncode(shapeHrefTooltip!)}\"" : "";
            sb.Append($"    <a class=\"shape-link\" href=\"{HtmlEncode(shapeHrefUrl!)}\" rel=\"noopener\" target=\"_blank\"{tooltipAttr} style=\"display:contents;cursor:pointer\">");
        }

        if (!string.IsNullOrEmpty(clipPathCss))
        {
            // For clip-path shapes: move fill to a clipped background layer, keep text unclipped
            // Extract fill-related styles for the clipped background layer
            var fillStyles = new List<string>();
            var borderStyles = new List<string>();
            var outerStyles = new List<string>();
            foreach (var s in styles)
            {
                if (s.StartsWith("background:") || s.StartsWith("background-image:"))
                    fillStyles.Add(s);
                else if (s.StartsWith("border"))
                    borderStyles.Add(s);
                else
                    outerStyles.Add(s);
            }
            // When wrapped in a link, add cursor:pointer to the shape <div> itself
            if (!string.IsNullOrEmpty(shapeHrefUrl)) outerStyles.Add("cursor:pointer");
            sb.Append($"    <div class=\"{shapeClass}\"{dataPathAttr} style=\"{string.Join(";", outerStyles)}\">");
            // Fill layer (clipped)
            if (fillStyles.Count > 0)
                sb.Append($"<div style=\"position:absolute;inset:0;{clipPathCss};{string.Join(";", fillStyles)}\"></div>");
            // Border layer for clip-path shapes: always use SVG polygon stroke
            if (parsedOutline != null && clipPathCss.StartsWith("clip-path:polygon("))
            {
                var (bw, dt, bc) = parsedOutline.Value;
                var polyStr = clipPathCss["clip-path:polygon(".Length..^1];
                var svgPoints = polyStr.Replace("%", "");
                var dashArr = DashTypeToSvgDasharray(dt, bw);
                var dashAttr = !string.IsNullOrEmpty(dashArr) ? $" stroke-dasharray=\"{dashArr}\"" : "";
                var safeColor = CssSanitizeColor(bc);
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\">");
                sb.Append($"<polygon points=\"{svgPoints}\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" vector-effect=\"non-scaling-stroke\" stroke-linecap=\"butt\"{dashAttr}/>");
                sb.Append("</svg>");
            }
        }
        else
        {
            if (!string.IsNullOrEmpty(shapeHrefUrl)) styles.Add("cursor:pointer");
            sb.Append($"    <div class=\"{shapeClass}\"{dataPathAttr} style=\"{string.Join(";", styles)}\">");
        }

        // Text content. `suppressText` is set by RenderInheritedShapes for layout/master
        // content placeholders: their <p:txBody> holds edit-prompt text ("Click to add
        // title") that belongs to the slide, not the layout. We still render the shape
        // chrome (fill/outline/geometry) so themed placeholder backgrounds survive.
        if (shape.TextBody != null && !suppressText)
        {
            // PowerPoint mirrors text along with the shape on flipH/flipV (e.g. flipH
            // renders "AI" as "IA"; flipV renders text upside-down). We deliberately
            // do NOT counter-flip — the parent shape transform applies to the inner
            // text, matching PowerPoint's rendering. An earlier implementation
            // counter-flipped here, but that diverged from real PowerPoint output.
            var flipStyle = "";

            // Shape-level RTL column flow: <a:bodyPr rtlCol="1"/> reverses
            // the column flow for the whole text body. Mirror with CSS so
            // Arabic / Hebrew shapes lay out the same way in HTML preview
            // as in PowerPoint.
            string rtlColStyle = "";
            if (bodyPr != null)
            {
                foreach (var attr in bodyPr.GetAttributes())
                {
                    if (attr.LocalName == "rtlCol" && (attr.Value == "1" || string.Equals(attr.Value, "true", StringComparison.OrdinalIgnoreCase)))
                    {
                        rtlColStyle = "direction:rtl;";
                        break;
                    }
                }
            }

            // wrap=none: suppress the .shape inherited `white-space:pre-wrap`
            // on the inner text container so the line extends horizontally
            // rather than wrapping inside the shape's width box.
            var wrapNoneStyle = wrapNone ? "white-space:nowrap;overflow:visible;" : "";
            var textStyle = !string.IsNullOrEmpty(flipStyle) || !string.IsNullOrEmpty(clipPathCss) || !string.IsNullOrEmpty(rtlColStyle) || !string.IsNullOrEmpty(wrapNoneStyle)
                ? $" style=\"{flipStyle}{rtlColStyle}{wrapNoneStyle}{(string.IsNullOrEmpty(clipPathCss) ? "" : "position:relative;")}\""
                : "";
            sb.Append($"<div class=\"shape-text valign-{valign}\"{textStyle}>");

            RenderTextBody(sb, shape.TextBody, themeColors, shape, part);
            sb.Append("</div>");
        }

        // SVG border overlay for non-solid outlines (dashed, dotted, dashDot etc.)
        if (parsedOutline != null && parsedOutline.Value.dashType != "solid")
        {
            var (bw, dt, bc) = parsedOutline.Value;
            var dashArr = DashTypeToSvgDasharray(dt, bw);
            var dashAttr = !string.IsNullOrEmpty(dashArr) ? $" stroke-dasharray=\"{dashArr}\"" : "";
            var safeColor = CssSanitizeColor(bc);

            if (!string.IsNullOrEmpty(clipPathCss) && clipPathCss.StartsWith("clip-path:polygon("))
            {
                // Polygon shapes — reuse existing polygon SVG approach
                var polyStr = clipPathCss["clip-path:polygon(".Length..^1];
                var svgPoints = polyStr.Replace("%", "");
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\">");
                sb.Append($"<polygon points=\"{svgPoints}\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" vector-effect=\"non-scaling-stroke\" stroke-linecap=\"butt\"{dashAttr}/>");
                sb.Append("</svg>");
            }
            else if (!string.IsNullOrEmpty(borderRadiusCss))
            {
                // Rounded rect — use SVG rect with rx/ry
                var rxMatch = System.Text.RegularExpressions.Regex.Match(borderRadiusCss, @"border-radius:([\d.]+)");
                var rx = rxMatch.Success ? rxMatch.Groups[1].Value : "0";
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
                sb.Append($"<rect x=\"{bw / 2:0.##}pt\" y=\"{bw / 2:0.##}pt\" width=\"calc(100% - {bw:0.##}pt)\" height=\"calc(100% - {bw:0.##}pt)\" rx=\"{rx}\" ry=\"{rx}\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" stroke-linecap=\"butt\"{dashAttr}/>");
                sb.Append("</svg>");
            }
            else if (presetGeom?.Preset?.InnerText == "ellipse")
            {
                // Ellipse — size in pt so stroke-width matches CSS border path.
                // CONSISTENCY(shape-stroke-unit): keep stroke-width in pt across solid/non-solid paths.
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
                sb.Append($"<ellipse cx=\"50%\" cy=\"50%\" rx=\"calc(50% - {bw / 2:0.##}pt)\" ry=\"calc(50% - {bw / 2:0.##}pt)\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" stroke-linecap=\"butt\"{dashAttr}/>");
                sb.Append("</svg>");
            }
            else
            {
                // Plain rect — use SVG rect sized in pt so stroke-width matches the CSS
                // `border:Npt solid` path (same visual weight). Inset by bw/2 so the stroke
                // sits entirely inside the content box (box-sizing:border-box equivalent).
                // CONSISTENCY(shape-stroke-unit): keep stroke-width in pt across solid/non-solid paths.
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
                sb.Append($"<rect x=\"{bw / 2:0.##}pt\" y=\"{bw / 2:0.##}pt\" width=\"calc(100% - {bw:0.##}pt)\" height=\"calc(100% - {bw:0.##}pt)\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" stroke-linecap=\"butt\"{dashAttr}/>");
                sb.Append("</svg>");
            }
        }

        sb.Append("</div>");
        if (!string.IsNullOrEmpty(shapeHrefUrl))
            sb.Append("</a>");
        sb.AppendLine();
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
        long slideW = SlideSizeDefaults.Widescreen16x9Cx;
        long slideH = SlideSizeDefaults.Widescreen16x9Cy;
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
        "roundRect" => (0.07, 0.07, 0.07, 0.07),
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
    private static int? ResolvePlaceholderFontSize(Shape shape, OpenXmlPart part, int level = 0)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return null; // Not a placeholder

        // 1. Check shape's own list style for the paragraph's level
        var lstStyle = shape.TextBody?.GetFirstChild<Drawing.ListStyle>();
        var defRp = GetLevelDefRp(lstStyle, level);
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

                    // Check candidate's list style at the correct level
                    var cLstStyle = candidate.TextBody?.GetFirstChild<Drawing.ListStyle>();
                    var cDefRp = GetLevelDefRp(cLstStyle, level);
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
                    var sDefRp = GetLevelDefRp(styleList, level);
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

    /// <summary>
    /// Get the DefaultRunProperties for a given paragraph level (0-8) from a list style or text style element.
    /// Maps level 0 → Level1ParagraphProperties, level 1 → Level2ParagraphProperties, etc.
    /// </summary>
    private static Drawing.DefaultRunProperties? GetLevelDefRp(OpenXmlCompositeElement? styleList, int level)
    {
        if (styleList == null) return null;
        OpenXmlElement? lvlPpr = level switch
        {
            0 => styleList.GetFirstChild<Drawing.Level1ParagraphProperties>(),
            1 => styleList.GetFirstChild<Drawing.Level2ParagraphProperties>(),
            2 => styleList.GetFirstChild<Drawing.Level3ParagraphProperties>(),
            3 => styleList.GetFirstChild<Drawing.Level4ParagraphProperties>(),
            4 => styleList.GetFirstChild<Drawing.Level5ParagraphProperties>(),
            5 => styleList.GetFirstChild<Drawing.Level6ParagraphProperties>(),
            6 => styleList.GetFirstChild<Drawing.Level7ParagraphProperties>(),
            7 => styleList.GetFirstChild<Drawing.Level8ParagraphProperties>(),
            8 => styleList.GetFirstChild<Drawing.Level9ParagraphProperties>(),
            _ => styleList.GetFirstChild<Drawing.Level1ParagraphProperties>(),
        };
        return lvlPpr?.GetFirstChild<Drawing.DefaultRunProperties>();
    }

    // ==================== Picture Rendering ====================

    /// <summary>
    /// Render a picture element to HTML. When called from a group, pass overridePos
    /// with the adjusted coordinates — the original element is NEVER modified.
    /// </summary>
    private static void RenderPicture(StringBuilder sb, Picture pic, OpenXmlPart slidePart,
        Dictionary<string, string> themeColors, (long x, long y, long cx, long cy)? overridePos = null,
        string? dataPath = null)
    {
        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        var xfrm = pic.ShapeProperties?.Transform2D;
        if (xfrm?.Offset == null || xfrm?.Extents == null) return;

        var x = overridePos?.x ?? xfrm.Offset.X?.Value ?? 0;
        var y = overridePos?.y ?? xfrm.Offset.Y?.Value ?? 0;
        var cx = overridePos?.cx ?? xfrm.Extents.Cx?.Value ?? 0;
        var cy = overridePos?.cy ?? xfrm.Extents.Cy?.Value ?? 0;

        var styles = new List<string>
        {
            $"left:{Units.EmuToPt(x)}pt",
            $"top:{Units.EmuToPt(y)}pt",
            $"width:{Units.EmuToPt(cx)}pt",
            $"height:{Units.EmuToPt(cy)}pt"
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

        // Effects: brightness, contrast, glow, shadow, opacity all roll
        // into one CSS `filter` property (drop-shadow / brightness /
        // contrast) so they compose. Mirror the shape renderer above:
        // shadowCss + glowCss merged into filter:..., reflection separate.
        var effectList = pic.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        var shadowCss = EffectListToShadowCss(effectList, themeColors);
        var glowCss = EffectListToGlowCss(effectList, themeColors);

        // brightness / contrast — Set.Media writes lumOff (brightness) and
        // lumMod (contrast) under a:blip; the SDK stores them as
        // OpenXmlUnknownElement on re-parse, so walk children by LocalName.
        var picBlipForFx = pic.BlipFill?.GetFirstChild<Drawing.Blip>();
        double? brightnessPct = null, contrastPct = null;
        if (picBlipForFx != null)
        {
            foreach (var kid in picBlipForFx.ChildElements)
            {
                if (kid.NamespaceUri != "http://schemas.openxmlformats.org/drawingml/2006/main") continue;
                var attr = kid.GetAttribute("val", "").Value;
                if (string.IsNullOrEmpty(attr) || !int.TryParse(attr, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var iv)) continue;
                if (kid.LocalName == "lumOff") brightnessPct = iv / 1000.0;
                else if (kid.LocalName == "lumMod") contrastPct = (iv - 100000) / 1000.0;
            }
        }

        var filterParts = new List<string>();
        // CSS brightness(1) = no change; +N% brightness → brightness(1 + N/100).
        if (brightnessPct.HasValue && Math.Abs(brightnessPct.Value) > 0.01)
            filterParts.Add($"brightness({1 + brightnessPct.Value / 100.0:0.###})");
        // CSS contrast(1) = no change; +N% contrast → contrast(1 + N/100).
        if (contrastPct.HasValue && Math.Abs(contrastPct.Value) > 0.01)
            filterParts.Add($"contrast({1 + contrastPct.Value / 100.0:0.###})");
        if (!string.IsNullOrEmpty(shadowCss))
            filterParts.Add(shadowCss.Replace("filter:", ""));
        if (!string.IsNullOrEmpty(glowCss))
            filterParts.Add(glowCss.Replace("filter:", ""));
        if (filterParts.Count > 0)
            styles.Add($"filter:{string.Join(" ", filterParts)}");

        // Opacity — a:blip/a:alphaModFix amount is 0-100000 (1000 = 1%).
        var picAlphaMod = picBlipForFx?.GetFirstChild<Drawing.AlphaModulationFixed>();
        if (picAlphaMod?.Amount?.HasValue == true && picAlphaMod.Amount.Value < 100000)
            styles.Add($"opacity:{picAlphaMod.Amount.Value / 100000.0:0.###}");

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

        sb.Append($"    <div class=\"picture\"{dataPathAttr} style=\"{string.Join(";", styles)}\">");

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

                // Crop — PowerPoint srcRect semantics: select a rectangular region of the
                // source image, then scale that region to fill the container.
                // CSS equivalent: render as a <div> with background-image, setting
                // background-size = container / visibleFraction and background-position
                // so the srcRect region aligns to the container edge.
                var srcRect = blipFill?.GetFirstChild<Drawing.SourceRectangle>();
                double srcL = 0, srcT = 0, srcR = 0, srcB = 0;
                if (srcRect != null)
                {
                    srcL = (srcRect.Left?.Value ?? 0) / 100000.0;
                    srcT = (srcRect.Top?.Value ?? 0) / 100000.0;
                    srcR = (srcRect.Right?.Value ?? 0) / 100000.0;
                    srcB = (srcRect.Bottom?.Value ?? 0) / 100000.0;
                }
                var hasCrop = srcL != 0 || srcT != 0 || srcR != 0 || srcB != 0;
                // Degenerate crop: L+R >= 100% or T+B >= 100% means zero/negative
                // visible area. PowerPoint renders nothing in this case; HTML
                // preview previously averaged the background-image into a muddy
                // block. Skip the picture draw entirely to match real PPT.
                var degenerateCrop = hasCrop && (srcL + srcR >= 1.0 || srcT + srcB >= 1.0);
                if (degenerateCrop)
                {
                    // Render nothing — matches PowerPoint's zero-area behavior.
                }
                else if (hasCrop)
                {
                    var visibleW = Math.Max(1 - srcL - srcR, 0.0001);
                    var visibleH = Math.Max(1 - srcT - srcB, 0.0001);
                    var bgSizeW = 100.0 / visibleW;
                    var bgSizeH = 100.0 / visibleH;
                    // background-position percentage semantics: pos% aligns pos%-of-image with pos%-of-container.
                    // To align srcRect (image region starting at fraction L) with container's left edge:
                    //   pos_x% = L / (srcL + srcR) * 100   (denominator = 1 - visibleW)
                    // Fallback to 0 when there's no crop on that axis (denominator == 0).
                    var denomX = srcL + srcR;
                    var denomY = srcT + srcB;
                    var bgPosX = denomX > 0 ? (srcL / denomX) * 100.0 : 0.0;
                    var bgPosY = denomY > 0 ? (srcT / denomY) * 100.0 : 0.0;
                    var bgStyle = $"width:100%;height:100%;background-image:url(data:{contentType};base64,{base64});background-repeat:no-repeat;background-size:{bgSizeW:0.##}% {bgSizeH:0.##}%;background-position:{bgPosX:0.##}% {bgPosY:0.##}%";
                    sb.Append($"<div style=\"{bgStyle}\"></div>");
                }
                else
                {
                    sb.Append($"<img src=\"data:{contentType};base64,{base64}\" loading=\"lazy\">");
                }
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

    private static void RenderConnector(StringBuilder sb, ConnectionShape cxn, Dictionary<string, string> themeColors, string? dataPath = null,
        (long x, long y, long cx, long cy)? overridePos = null)
        => RenderConnector(sb, cxn.ShapeProperties, themeColors, dataPath, overridePos);

    // Shared SVG line/polyline/path renderer for both <p:cxnSp> connectors and
    // <p:sp> shapes with prst="line". Reads geometry + outline from a
    // ShapeProperties and emits a connector-style div.
    // overridePos: when rendering inside a group, the caller supplies coordinates
    // already transformed into the group's child coordinate system (see RenderShape /
    // RenderPicture's parallel parameter). Without it the connector's raw slide-absolute
    // EMU coords are emitted as offsets from the group container — placing the line
    // far outside the group div, where it disappears.
    private static void RenderConnector(StringBuilder sb, ShapeProperties? spPr, Dictionary<string, string> themeColors, string? dataPath = null,
        (long x, long y, long cx, long cy)? overridePos = null)
    {
        var xfrm = spPr?.Transform2D;
        if (overridePos == null && (xfrm?.Offset == null || xfrm?.Extents == null)) return;

        long x, y, cx, cy;
        if (overridePos != null)
        {
            (x, y, cx, cy) = overridePos.Value;
        }
        else
        {
            x = xfrm!.Offset!.X?.Value ?? 0;
            y = xfrm.Offset.Y?.Value ?? 0;
            cx = xfrm.Extents!.Cx?.Value ?? 0;
            cy = xfrm.Extents.Cy?.Value ?? 0;
        }

        var flipH = xfrm?.HorizontalFlip?.Value == true;
        var flipV = xfrm?.VerticalFlip?.Value == true;

        // SVG line
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
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
        var widthPt = Units.EmuToPt(renderCx);
        var heightPt = Units.EmuToPt(renderCy);

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
            // Both markers use a right-pointing triangle with tip at (arrowSize, arrowSize/2).
            // For marker-start we use orient="auto-start-reverse" so SVG flips the right-pointing
            // triangle to point outward (leftward) at the line's start. Authoring both markers
            // with the same geometry avoids a past bug where the head marker was authored
            // leftward-pointing and the reverse flipped it inward on straight connectors.
            if (hasHead)
            {
                defs.Append($"<marker id=\"ah\" markerWidth=\"{arrowSize:0.#}\" markerHeight=\"{arrowSize:0.#}\" refX=\"{arrowSize:0.#}\" refY=\"{arrowSize / 2:0.#}\" orient=\"auto-start-reverse\"><polygon points=\"0 0,{arrowSize:0.#} {arrowSize / 2:0.#},0 {arrowSize:0.#}\" fill=\"{safeColor}\"/></marker>");
                markerStartAttr = " marker-start=\"url(#ah)\"";
            }
            if (hasTail)
            {
                defs.Append($"<marker id=\"at\" markerWidth=\"{arrowSize:0.#}\" markerHeight=\"{arrowSize:0.#}\" refX=\"{arrowSize:0.#}\" refY=\"{arrowSize / 2:0.#}\" orient=\"auto\"><polygon points=\"0 0,{arrowSize:0.#} {arrowSize / 2:0.#},0 {arrowSize:0.#}\" fill=\"{safeColor}\"/></marker>");
                markerEndAttr = " marker-end=\"url(#at)\"";
            }
            defs.Append("</defs>");
            markerDefs = defs.ToString();
        }

        // Branch on preset geometry: straightConnectorN -> line; bentConnectorN -> polyline;
        // curvedConnectorN -> cubic bezier path. Falls back to straight line for unknown presets.
        var prstGeom = spPr?.GetFirstChild<Drawing.PresetGeometry>();
        var preset = prstGeom?.Preset?.HasValue == true ? (prstGeom.Preset.InnerText ?? "straightConnector1") : "straightConnector1";

        // Bent/curved connectors need both axes to draw their perpendicular segment.
        // When one axis is 0 (degenerate — typical when from=/to= shapes are aligned
        // horizontally or vertically), the polyline/bezier collapses into a 1-2pt strip
        // and any arrow marker covers the whole thing, producing a "dot". Degrade to a
        // straight line in that case so the rendered output stays meaningful.
        // PowerPoint would route the elbow above/below using connection points, but we
        // don't compute those — straight is the honest fallback.
        if ((cx == 0 || cy == 0)
            && (preset.StartsWith("bentConnector", StringComparison.Ordinal)
                || preset.StartsWith("curvedConnector", StringComparison.Ordinal)))
        {
            preset = "straightConnector1";
        }

        // CONSISTENCY(shape-stroke-unit): stroke-width in pt matches CSS border path (see R3 fix).
        var strokeAttrs = $"stroke=\"{safeColor}\" stroke-width=\"{lineWidth:0.##}pt\" fill=\"none\"{dashAttr}{markerStartAttr}{markerEndAttr}";

        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        // CONSISTENCY(shape-rotation): connectors use the same Transform2D.Rotation
        // slot as shapes/pictures/groups; apply the same CSS transform so the rendered
        // line matches PowerPoint. Default transform-origin (50% 50%) matches OOXML
        // rotation pivot (bounding-box center).
        var cxnRotTransform = "";
        if (xfrm?.Rotation != null && xfrm.Rotation.Value != 0)
            cxnRotTransform = $";transform:rotate({xfrm.Rotation.Value / 60000.0:0.##}deg)";
        sb.AppendLine($"    <div class=\"connector\"{dataPathAttr} style=\"left:{Units.EmuToPt(renderX)}pt;top:{Units.EmuToPt(renderY)}pt;width:{widthPt}pt;height:{heightPt}pt{cxnRotTransform}\">");

        if (preset.StartsWith("bentConnector", StringComparison.Ordinal))
        {
            // Bent connectors: right-angle polyline. Use viewBox=0..100 so stretched
            // preserveAspectRatio=none fills the container.
            // bentConnector2: single 90-degree bend (2 segments, 3 points).
            // bentConnector3 (default): 3 segments with mid bend — (0,0) -> (50,0) -> (50,100) -> (100,100).
            // bentConnector4/5: approximate with 25/75 splits when no adjustments set.
            string points = preset switch
            {
                "bentConnector2" => "0,0 100,0 100,100",
                "bentConnector4" or "bentConnector5" => "0,0 25,0 25,50 75,50 75,100 100,100",
                _ => "0,0 50,0 50,100 100,100", // bentConnector3
            };
            sb.AppendLine("      <svg width=\"100%\" height=\"100%\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\" style=\"overflow:visible;display:block\">");
            if (!string.IsNullOrEmpty(markerDefs))
                sb.AppendLine($"        {markerDefs}");
            sb.AppendLine($"        <polyline points=\"{points}\" {strokeAttrs}/>");
            sb.AppendLine("      </svg>");
        }
        else if (preset.StartsWith("curvedConnector", StringComparison.Ordinal))
        {
            // Curved connectors: cubic bezier S-curve. Author in 0..100 viewBox.
            // curvedConnector3 default: M 0,0 C 50,0 50,100 100,100 (horizontal-entry S).
            string d = preset switch
            {
                "curvedConnector2" => "M 0,0 Q 100,0 100,100",
                "curvedConnector4" or "curvedConnector5" => "M 0,0 C 25,0 25,50 50,50 C 75,50 75,100 100,100",
                _ => "M 0,0 C 50,0 50,100 100,100", // curvedConnector3
            };
            sb.AppendLine("      <svg width=\"100%\" height=\"100%\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\" style=\"overflow:visible;display:block\">");
            if (!string.IsNullOrEmpty(markerDefs))
                sb.AppendLine($"        {markerDefs}");
            sb.AppendLine($"        <path d=\"{d}\" {strokeAttrs}/>");
            sb.AppendLine("      </svg>");
        }
        else
        {
            sb.AppendLine("      <svg width=\"100%\" height=\"100%\" preserveAspectRatio=\"none\" style=\"overflow:visible;display:block\">");
            if (!string.IsNullOrEmpty(markerDefs))
                sb.AppendLine($"        {markerDefs}");
            sb.AppendLine($"        <line x1=\"{svgX1}\" y1=\"{svgY1}\" x2=\"{svgX2}\" y2=\"{svgY2}\" {strokeAttrs}/>");
            sb.AppendLine("      </svg>");
        }
        sb.AppendLine("    </div>");
    }

    // ==================== Group Rendering ====================

    private void RenderGroup(StringBuilder sb, GroupShape grp, SlidePart slidePart, Dictionary<string, string> themeColors, string? dataPath = null)
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

        // Group is selected as a whole. Children inside the group don't get their own
        // data-path because nested @id= addressing isn't currently supported by
        // ResolveIdPath — clicks inside walk up via closest('[data-path]') and select
        // the group container.
        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        // CONSISTENCY(group-rotation): match single-shape rotation idiom from RenderShape
        // (transform:rotate(Ndeg)). OOXML group rotation rotates children as a composite
        // around the group's bounding-box center; CSS default transform-origin (50% 50%)
        // matches this.
        var grpTransform = "";
        if (grpXfrm?.Rotation != null && grpXfrm.Rotation.Value != 0)
            grpTransform = $";transform:rotate({grpXfrm.Rotation.Value / 60000.0:0.##}deg)";
        sb.AppendLine($"    <div class=\"group\"{dataPathAttr} style=\"left:{Units.EmuToPt(x)}pt;top:{Units.EmuToPt(y)}pt;width:{Units.EmuToPt(cx)}pt;height:{Units.EmuToPt(cy)}pt{grpTransform}\">");

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
                    // CONSISTENCY(group-child-pos): mirror Shape/Picture branches above —
                    // a connector inside a group must have its slide-absolute EMU coords
                    // re-projected into the group's child coordinate system. Previously
                    // the raw coords were emitted as offsets inside the group div, which
                    // placed the connector far outside the group (invisible).
                    var pos = CalcGroupChildPos(cxn.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderConnector(sb, cxn, themeColors, dataPath: null, overridePos: pos);
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

        // CONSISTENCY(group-rotation): same idiom as RenderGroup
        var grpTransform = "";
        if (grpXfrm?.Rotation != null && grpXfrm.Rotation.Value != 0)
            grpTransform = $";transform:rotate({grpXfrm.Rotation.Value / 60000.0:0.##}deg)";
        sb.AppendLine($"    <div class=\"group\" style=\"left:{Units.EmuToPt(x)}pt;top:{Units.EmuToPt(y)}pt;width:{Units.EmuToPt(cx)}pt;height:{Units.EmuToPt(cy)}pt{grpTransform}\">");

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
                {
                    // CONSISTENCY(group-child-pos): see RenderGroup ConnectionShape branch.
                    var pos = CalcGroupChildPos(cxn.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderConnector(sb, cxn, themeColors, dataPath: null, overridePos: pos);
                    break;
                }
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
        SlidePart slidePart, Dictionary<string, string> themeColors, string? dataPath = null)
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

        var leftPt = Units.EmuToPt(x);
        var topPt = Units.EmuToPt(y);
        var widthPt2 = Units.EmuToPt(cx);
        var heightPt2 = Units.EmuToPt(cy);

        if (isModel3D)
        {
            RenderModel3D(sb, acElement, slidePart, leftPt, topPt, widthPt2, heightPt2, dataPath);
        }
        else
        {
            // Zoom: render fallback image
            RenderZoomFallback(sb, acElement, slidePart, leftPt, topPt, widthPt2, heightPt2);
        }
    }

    private static int _model3dCounter;
    // Cache: GLB content hash → JS variable name, to avoid embedding the same
    // GLB multiple times within a single render. MUST be reset between renders
    // (see ResetModel3DRenderState) — otherwise call N+1 hits the cache and
    // skips emitting the data script that the new HTML's module script needs.
    private static readonly Dictionary<string, string> _glbDataCache = new();

    internal static void ResetModel3DRenderState()
    {
        _model3dCounter = 0;
        _glbDataCache.Clear();
    }

    /// <summary>
    /// Render a 3D model using Three.js with the embedded GLB data.
    /// Same GLB files across slides are deduplicated — embedded once, referenced by variable.
    /// </summary>
    private static void RenderModel3D(StringBuilder sb, OpenXmlElement acElement,
        SlidePart slidePart, double leftPt, double topPt, double widthPt, double heightPt,
        string? dataPath = null)
    {
        // Find the model3d element and get the GLB relationship
        var model3d = acElement.Descendants().FirstOrDefault(d => d.LocalName == "model3d");
        if (model3d == null) return;

        var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var embedId = model3d.GetAttribute("embed", rNs).Value;
        if (string.IsNullOrEmpty(embedId)) return;

        // Deduplicate: use content hash so identical GLBs across slides share one copy.
        // Also surface the GLB filename (relationship target) for the placeholder label.
        string glbVarName;
        string? glbFileName = null;
        try
        {
            var part = slidePart.GetPartById(embedId);
            try
            {
                var rel = slidePart.GetReferenceRelationship(embedId);
                glbFileName = System.IO.Path.GetFileName(rel.Uri?.ToString() ?? "");
            }
            catch { }
            if (string.IsNullOrEmpty(glbFileName))
                glbFileName = System.IO.Path.GetFileName(part.Uri?.ToString() ?? "");
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

        // Bordered placeholder underlay: visible until Three.js paints the canvas
        // over it, and remains the only visible surface when Three.js / WebGL are
        // unavailable and no mc:Fallback image was authored. Mirrors the OLE
        // placeholder pattern — surface presence + a data-path for selection.
        var containerId = $"m3d_wrap_{canvasId}";
        var label = string.IsNullOrEmpty(glbFileName)
            ? "3D Model"
            : HtmlEncode($"3D Model: {glbFileName}");
        var dpAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath!)}\"";
        sb.AppendLine($"    <div id=\"{containerId}\"{dpAttr} style=\"position:absolute;" +
            $"left:{leftPt:0.##}pt;top:{topPt:0.##}pt;" +
            $"width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;" +
            $"border:2px dashed rgba(108,117,125,0.6);border-radius:4px;" +
            $"background:rgba(248,249,250,0.7);" +
            $"overflow:hidden;box-sizing:border-box;\">");
        sb.AppendLine($"      <div class=\"m3d-label\" style=\"position:absolute;inset:0;" +
            $"display:flex;align-items:center;justify-content:center;" +
            $"font:11pt sans-serif;color:#495057;text-align:center;padding:4px;" +
            $"pointer-events:none;\">{label}</div>");
        sb.AppendLine($"      <canvas id=\"{canvasId}\" style=\"position:relative;width:100%;height:100%;\"></canvas>");
        if (fallbackImgSrc != null)
            sb.AppendLine($"      <img class=\"m3d-fallback\" src=\"{fallbackImgSrc}\" style=\"width:100%;height:100%;object-fit:contain;display:none;\" />");
        sb.AppendLine("    </div>");

        sb.AppendLine($@"    <script type=""module"">
    let THREE, GLTFLoader;
    try {{
      THREE = await import('three');
      ({{ GLTFLoader }} = await import('three/addons/loaders/GLTFLoader.js'));
    }} catch(e) {{
      // Three.js unavailable (offline) — show fallback image
      const c = document.getElementById('{canvasId}');
      if (c) {{ c.style.display='none'; const fb=c.parentElement?.querySelector('.m3d-fallback'); if(fb) fb.style.display='block'; }}
      throw e; // stop execution of this module
    }}
    (function() {{
      const canvas = document.getElementById('{canvasId}');
      if (!canvas) return;
      const container = canvas.parentElement;
      try {{
        const designW = {widthPt:0.##} * 96 / 72;
        const designH = {heightPt:0.##} * 96 / 72;
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
        SlidePart slidePart, double leftPt, double topPt, double widthPt, double heightPt)
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
            $"left:{leftPt:0.##}pt;top:{topPt:0.##}pt;" +
            $"width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;" +
            $"border:2px dashed rgba(255,193,7,0.6);border-radius:8px;" +
            $"overflow:hidden;\">");
        if (imgSrc != null)
            sb.AppendLine($"      <img src=\"{imgSrc}\" style=\"width:100%;height:100%;object-fit:contain;\" />");
        sb.AppendLine("    </div>");
    }

    /// <summary>
    /// Render an OLE GraphicFrame as a bordered placeholder carrying the
    /// ProgID label. Mirrors the model3d / zoom fallback pattern. Real
    /// rendering of the embedded payload is intentionally not attempted —
    /// PowerPoint itself stamps a pre-baked thumbnail (the inner p:pic blip)
    /// at author time; for the HTML preview we surface the OLE's presence so
    /// the slide canvas is not silently empty and selection has a data-path
    /// to bind to.
    /// </summary>
    private static void RenderOlePlaceholder(StringBuilder sb, GraphicFrame gf, string? dataPath = null)
    {
        var xfrm = gf.Transform;
        var x = xfrm?.Offset?.X?.Value ?? 0;
        var y = xfrm?.Offset?.Y?.Value ?? 0;
        var cx = xfrm?.Extents?.Cx?.Value ?? 0;
        var cy = xfrm?.Extents?.Cy?.Value ?? 0;
        var leftPt = Units.EmuToPt(x);
        var topPt = Units.EmuToPt(y);
        var widthPt = Units.EmuToPt(cx);
        var heightPt = Units.EmuToPt(cy);

        var oleEl = gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().First();
        var progId = oleEl.ProgId?.Value ?? "Embedded Object";
        var label = HtmlEncode($"OLE: {progId}");
        var dpAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";

        sb.AppendLine($"    <div class=\"ole-placeholder\"{dpAttr} style=\"position:absolute;" +
            $"left:{leftPt:0.##}pt;top:{topPt:0.##}pt;" +
            $"width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;" +
            $"border:2px dashed rgba(108,117,125,0.6);border-radius:4px;" +
            $"display:flex;align-items:center;justify-content:center;" +
            $"font:11pt sans-serif;color:#495057;background:rgba(248,249,250,0.7);" +
            $"overflow:hidden;text-align:center;padding:4px;box-sizing:border-box;\">" +
            $"{label}</div>");
    }
}
