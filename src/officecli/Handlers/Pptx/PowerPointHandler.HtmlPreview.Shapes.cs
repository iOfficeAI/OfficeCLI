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
    // ==================== Shape Rendering ====================

    /// <summary>
    /// Render a shape element to HTML. When called from a group, pass overridePos
    /// with the adjusted coordinates — the original element is NEVER modified.
    /// </summary>
    private static void RenderShape(StringBuilder sb, Shape shape, OpenXmlPart part,
        Dictionary<string, string> themeColors, (long x, long y, long cx, long cy)? overridePos = null,
        string? dataPath = null, bool suppressText = false, int? slideNumber = null)
    {
        // prst="line" auto-shapes are line-segment geometry; render as SVG
        // through the connector pipeline so they don't degrade to a div with
        // border (which fakes a thin filled rect and loses zero-width/height
        // line semantics — observed on slide 2 of test-samples/07.pptx).
        var prstGeomEarly = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (prstGeomEarly?.Preset?.HasValue == true
            && prstGeomEarly.Preset.InnerText == "line")
        {
            // Forward the shape's text body: a prstGeom="line" sp can carry a <p:txBody>
            // label (PowerPoint renders it over the line); the connector overlay draws it.
            // Without this the label was silently dropped.
            RenderConnector(sb, shape.ShapeProperties, themeColors, dataPath, overridePos,
                cxnTextBody: shape.TextBody, style: shape.ShapeStyle, part: part);
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
                            // Reject javascript:/vbscript:/data: etc. — OOXML hyperlink
                            // relationships are attacker-controlled and HtmlEncode does not
                            // neutralize a dangerous scheme. Mirrors the Word/Excel previews.
                            if (rel?.Uri != null && Core.HyperlinkUriValidator.IsSafeScheme(rel.Uri.ToString()))
                                shapeHrefUrl = rel.Uri.ToString();
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

        // Zero-thickness rule: a shape whose box collapses to one dimension
        // (cy==0 horizontal, cx==0 vertical) with an outline is a divider line.
        // A border-box <div> can't draw a 1px line in the collapsed axis (the
        // solid path renders a 2*width-tall strip; the SVG-dash path computes a
        // negative rect height and vanishes). Route through the connector
        // pipeline, which already draws zero-dimension lines with the correct
        // color/width/dash (DashTypeToSvgDasharray). Only when the shape has no
        // text — a text-bearing collapsed shape is unusual; keep the div path.
        if ((cx == 0 || cy == 0)
            && shape.ShapeProperties?.GetFirstChild<Drawing.Outline>() != null
            && string.IsNullOrWhiteSpace(GetShapeText(shape)))
        {
            RenderConnector(sb, shape.ShapeProperties, themeColors, dataPath, overridePos,
                style: shape.ShapeStyle, part: part);
            return;
        }

        // Bug #8(A): a shape with <a:spAutoFit/> grows to fit its text in real
        // PowerPoint. A fixed pt height clips overflowing content, so emit
        // min-height + height:auto for spAutoFit. All other autofit modes
        // (normAutofit / noAutofit / none) keep the fixed OOXML height.
        var autoFitBodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        var isSpAutoFit = autoFitBodyPr?.GetFirstChild<Drawing.ShapeAutoFit>() != null;
        var heightStyle = isSpAutoFit
            ? $"min-height:{Units.EmuToPt(cy)}pt;height:auto"
            : $"height:{Units.EmuToPt(cy)}pt";

        var styles = new List<string>
        {
            $"left:{Units.EmuToPt(x)}pt",
            $"top:{Units.EmuToPt(y)}pt",
            $"width:{Units.EmuToPt(cx)}pt",
            heightStyle
        };

        // Fill
        var fillCss = GetShapeFillCss(shape.ShapeProperties, part, themeColors);
        // R12-5: a slide PLACEHOLDER with no own fill inherits the matching
        // layout (then master) placeholder's fill. Parallels the R7
        // ResolveInheritedPosition walk. Explicit slide fill still wins.
        if (string.IsNullOrEmpty(fillCss))
        {
            var inheritedPh = ResolveInheritedPlaceholderShape(shape, part);
            if (inheritedPh != null)
                fillCss = GetShapeFillCss(inheritedPh.ShapeProperties, part, themeColors);
        }
        // Style-matrix fallback (R11-1): when spPr carries no fill element, resolve the
        // shape's <p:style>/<a:fillRef> against the theme FormatScheme. Explicit spPr
        // fill always wins — only consult fillRef when GetShapeFillCss returned "".
        if (string.IsNullOrEmpty(fillCss))
            fillCss = GetStyleFillRefCss(shape.ShapeStyle, part, themeColors);
        if (!string.IsNullOrEmpty(fillCss))
            styles.Add(fillCss);

        // Border/outline — parse for later; solid goes to CSS, non-solid to SVG
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        // Gradient outline (<a:ln>/<a:gradFill>): a solid border can't represent
        // the gradient, so render the stops + direction via CSS border-image with
        // a linear-gradient — reusing the same stop/angle conversion the shape
        // FILL gradient uses (GradientToCss). Takes precedence over the solid /
        // SVG-dashed paths. Without this the gradient line color resolves to null
        // in ParseOutline and degrades to solid #000000.
        var outlineGradFill = outline?.GetFirstChild<Drawing.GradientFill>();
        var parsedOutline = outline != null ? ParseOutline(outline, themeColors) : null;
        // R23-B: a gradient outline used to always emit `border:Npt solid transparent`
        // + `border-image:<gradient> 1`. But CSS `border-image` is IGNORED whenever
        // `border-radius` is also set (spec) — so a roundRect (or any preset whose
        // PresetGeometryToCss yields border-radius / a clip-path) rendered the
        // gradient outline as a SQUARE solid-ish border. Defer rounded / non-rect
        // gradient outlines to the SVG stroke overlay below (stroke=url(#grad)),
        // mirroring the connector gradient-stroke path; plain rects keep border-image
        // (which works there).
        double gradOutlineW = 0;
        if (outlineGradFill != null)
        {
            gradOutlineW = outline!.Width?.HasValue == true
                ? outline.Width.Value / EmuConverter.EmuPerPointF : 1.0;
            if (gradOutlineW < 0.5) gradOutlineW = 0.5;
            // Suppress the solid-CSS and SVG-dashed outline paths below; the gradient
            // is rendered either as border-image (plain rect) or SVG stroke (rounded).
            parsedOutline = null;
        }
        else if (parsedOutline != null && parsedOutline.Value.dashType == "solid")
        {
            // Compound line (cmpd=dbl/thickThin/thinThick/tri) draws multiple parallel
            // lines; CSS `double` renders two parallel lines when border-width >= ~3px.
            var solidStyle = parsedOutline.Value.cmpd != "sng" ? "double" : "solid";
            styles.Add($"border:{parsedOutline.Value.widthPt:0.##}pt {solidStyle} {parsedOutline.Value.color}");
        }
        // Non-solid outlines rendered as SVG after the shape div

        // Style-matrix fallback (R11-2): when spPr carries no <a:ln>, resolve the
        // shape's <p:style>/<a:lnRef> against the theme FormatScheme.LineStyleList.
        // Explicit spPr outline always wins — only consult lnRef when outline == null.
        if (outline == null)
        {
            var lnRefCss = GetStyleLineRefCss(shape.ShapeStyle, part, themeColors);
            if (!string.IsNullOrEmpty(lnRefCss))
                styles.Add(lnRefCss);
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
        // R48: multi-subpath / fill="none" custGeom rendered as an inline SVG <path>
        // overlay instead of clip-path:polygon (which can hold only one polygon).
        string custGeomSvgFillD = "";
        string custGeomSvgStrokeD = "";
        long custGeomSvgW = 100000L;
        long custGeomSvgH = 100000L;
        // R48: true once a custGeom routes its fill/stroke through the SVG overlay, so the
        // shape <div> must NOT also paint the solidFill background (or for an all-fill="none"
        // path, must paint nothing at all).
        bool custGeomSvgRouted = false;
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
                // R48: a single filled subpath keeps the clip-path:polygon path (broadest
                // browser support, unchanged common case). Multiple subpaths OR any
                // fill="none" subpath cannot be expressed by one clip-path polygon, so
                // route those through an inline SVG <path> overlay (set below).
                var pathLst = custGeom.GetFirstChild<Drawing.PathList>();
                var subPaths = pathLst?.Elements<Drawing.Path>().ToList() ?? new List<Drawing.Path>();
                var needsSvg = subPaths.Count > 1
                    || subPaths.Any(p => p.Fill?.Value == Drawing.PathFillModeValues.None);

                if (needsSvg)
                {
                    // Multi-subpath or fill="none": never use clip-path (it can hold only
                    // one polygon and cannot express a fill-less stroke-only path). The SVG
                    // overlay carries fill+stroke; if every subpath is both fill="none" and
                    // stroke=off the shape is correctly invisible (no fill, no clip-path).
                    custGeomSvgRouted = true;
                    CustomGeometryToSvgPaths(custGeom, out var fd, out var sd, out var pw, out var ph);
                    custGeomSvgFillD = fd;
                    custGeomSvgStrokeD = sd;
                    custGeomSvgW = pw;
                    custGeomSvgH = ph;
                }
                else
                {
                    var clipPath = CustomGeometryToClipPath(custGeom);
                    if (!string.IsNullOrEmpty(clipPath))
                        clipPathCss = clipPath;
                }
            }
        }

        // R23-B: gradient outline — pick the render strategy now that geometry is
        // known. Plain rect (no border-radius, no clip-path) keeps border-image
        // (it renders correctly). Rounded / clipped geometry uses the SVG stroke
        // overlay emitted in the post-div block (border-image is silently ignored
        // when border-radius is present, which produced a square solid border).
        string? gradOutlineSvg = null;
        bool gradOutlineRounded = false;
        if (outlineGradFill != null)
        {
            gradOutlineRounded = !string.IsNullOrEmpty(borderRadiusCss)
                || !string.IsNullOrEmpty(clipPathCss);
            if (!gradOutlineRounded)
            {
                var gradCss = GradientToCss(outlineGradFill, themeColors);
                styles.Add($"border:{gradOutlineW:0.##}pt solid transparent");
                styles.Add($"border-image:{gradCss} 1");
            }
            else
            {
                var gradId = $"shg{_markerCounter++}";
                gradOutlineSvg = BuildSvgLinearGradient(outlineGradFill, gradId, themeColors, out _);
                if (string.IsNullOrEmpty(gradOutlineSvg))
                    gradOutlineSvg = null;
                else
                    gradOutlineSvg = "url(#" + gradId + ")|" + gradOutlineSvg;
            }
        }

        // Shadow + Glow → combine into single filter property
        var effectList = shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        // Style-matrix fallback (R11-4 / R14-1): when spPr carries no <a:effectLst>,
        // resolve the shape's <p:style>/<a:effectRef> against the theme
        // FormatScheme.EffectStyleList and emit shadow + glow + reflection from it.
        // Explicit spPr effects always win — only consult effectRef when none present.
        var effectFor = effectList;
        if (effectList == null)
            effectFor = ResolveStyleEffectRefList(shape.ShapeStyle, part);
        var shadowCss = EffectListToShadowCss(effectFor, themeColors);
        var glowCss = EffectListToGlowCss(effectFor, themeColors);
        // Merge multiple filter:drop-shadow into one filter property.
        // EffectListToShadowCss returns either a "filter:..." value (outer/preset
        // shadow) or a "box-shadow:inset ..." declaration (innerShdw, which has no
        // CSS filter equivalent). Route each to the right place: filter values into
        // filterParts, box-shadow segments into boxShadowParts so a single combined
        // box-shadow declaration is emitted (CSS only honors the last box-shadow).
        var filterParts = new List<string>();
        var boxShadowParts = new List<string>();
        if (!string.IsNullOrEmpty(shadowCss))
        {
            if (shadowCss.StartsWith("box-shadow:"))
                boxShadowParts.Add(shadowCss["box-shadow:".Length..]);
            else
                filterParts.Add(shadowCss.Replace("filter:", ""));
        }
        if (!string.IsNullOrEmpty(glowCss))
            filterParts.Add(glowCss.Replace("filter:", ""));
        var blurCss = EffectListToBlurCss(effectFor);
        if (!string.IsNullOrEmpty(blurCss))
            filterParts.Add(blurCss);
        if (filterParts.Count > 0)
            styles.Add($"filter:{string.Join(" ", filterParts)}");

        // Reflection → CSS -webkit-box-reflect
        var reflectionCss = EffectListToReflectionCss(effectFor);
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
            var edgePx = Math.Max(2, softEdge.Radius.Value / EmuConverter.EmuPerPointF * 0.8);
            // Use linear-gradient masks on all 4 edges to create edge fade-out
            styles.Add($"-webkit-mask-image:linear-gradient(to right,transparent 0,black {edgePx:0.#}px,black calc(100% - {edgePx:0.#}px),transparent 100%)," +
                       $"linear-gradient(to bottom,transparent 0,black {edgePx:0.#}px,black calc(100% - {edgePx:0.#}px),transparent 100%)");
            styles.Add("-webkit-mask-composite:source-in;mask-composite:intersect");
        }

        // Bevel → approximate with inset box-shadow for a subtle 3D appearance
        var sp3d = shape.ShapeProperties?.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d?.BevelTop != null)
        {
            var bevelW = sp3d.BevelTop.Width?.HasValue == true ? sp3d.BevelTop.Width.Value / EmuConverter.EmuPerPointF : 6; // OOXML default 76200 EMU = 6pt
            var bW = Math.Max(1, bevelW * 0.5);
            boxShadowParts.Add($"inset {bW:0.#}px {bW:0.#}px {bW * 1.5:0.#}px rgba(255,255,255,0.25),inset -{bW:0.#}px -{bW:0.#}px {bW * 1.5:0.#}px rgba(0,0,0,0.15)");
        }

        // Emit one combined box-shadow (inner shadow + bevel). CSS only honors the
        // last box-shadow in an inline style, so they must share a single declaration.
        if (boxShadowParts.Count > 0)
            styles.Add($"box-shadow:{string.Join(",", boxShadowParts)}");

        // Note: fill opacity (alpha) is already baked into rgba() by ResolveFillColor.
        // Do NOT add a separate CSS opacity here — it would double-apply.

        // Text margins
        var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        // R4-2: surface <a:prstTxWarp prst="…"/> (WordArt text warp) as a
        // data-textwarp attribute + a "text-warp" marker class on the shape div.
        // CSS cannot faithfully reproduce per-glyph warp paths, but emitting the
        // preset name makes a warped shape visibly distinct from an unwarped one
        // in the HTML (and lets downstream tooling/JS apply an SVG textPath).
        var prstTxWarp = bodyPr?.GetFirstChild<Drawing.PresetTextWarp>();
        var textWarpAttr = prstTxWarp?.Preset?.HasValue == true
            ? $" data-textwarp=\"{HtmlEncode(prstTxWarp.Preset.InnerText!)}\""
            : "";
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

        // Vertical alignment class.
        //
        // Default (no explicit anchor, nothing inherited) is shape-kind
        // dependent in real PowerPoint:
        //   - text box (<p:cNvSpPr txBox="1">)  → top
        //   - placeholder (<p:ph>)              → inherits from layout/master
        //                                          (resolved below via
        //                                          ResolveInheritedAnchor)
        //   - autoshape (prstGeom/custGeom, no txBox) → center
        // An explicit <a:bodyPr anchor="t|ctr|b"> always wins.
        var isPlaceholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>() != null;
        var isTextBox = shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties
            ?.TextBox?.Value == true;
        var valign = (!isPlaceholder && !isTextBox) ? "center" : "top";
        if (bodyPr?.Anchor?.HasValue == true)
        {
            valign = bodyPr.Anchor.InnerText switch
            {
                "ctr" => "center",
                "b" => "bottom",
                _ => "top"
            };
        }
        else if (isPlaceholder)
        {
            // Slide bodyPr has no explicit anchor: resolve via the standard
            // slide→layout→master placeholder inheritance (same ph type/idx
            // matching used for position/font/color). The first ancestor
            // bodyPr that declares an anchor wins. Absent everywhere → "top".
            var inheritedAnchor = ResolveInheritedAnchor(shape, part);
            if (inheritedAnchor != null)
            {
                valign = inheritedAnchor switch
                {
                    "ctr" => "center",
                    "b" => "bottom",
                    _ => "top"
                };
            }
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
            || shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>() != null
            || shape.ShapeProperties?.GetFirstChild<Drawing.PatternFill>() != null;
        var explicitNoAutofit = bodyPr?.GetFirstChild<Drawing.NoAutoFit>() != null;
        var allowOverflow = wrapNone || explicitNoAutofit;
        var shapeClass = hasFillBg && !allowOverflow ? "shape has-fill" : "shape";
        if (!string.IsNullOrEmpty(textWarpAttr)) shapeClass += " text-warp";
        if (allowOverflow) styles.Add("overflow:visible");

        // Open <a> wrapper for shape-level hyperlink (before the shape <div>)
        if (!string.IsNullOrEmpty(shapeHrefUrl))
        {
            var tooltipAttr = !string.IsNullOrEmpty(shapeHrefTooltip)
                ? $" title=\"{HtmlEncode(shapeHrefTooltip!)}\"" : "";
            sb.Append($"    <a class=\"shape-link\" href=\"{HtmlEncode(shapeHrefUrl!)}\" rel=\"noopener\" target=\"_blank\"{tooltipAttr} style=\"display:contents;cursor:pointer\">");
        }

        if (custGeomSvgRouted)
        {
            // R48: multi-subpath / fill="none" custGeom. The shape fill is drawn by the
            // SVG <path> (so background:/border: must NOT also be applied to the div).
            // Pull the fill color out of the CSS background declaration; gradients
            // degrade gracefully to the first parsed color (or no fill).
            var outerStyles = new List<string>();
            string svgFill = "none";
            foreach (var s in styles)
            {
                if (s.StartsWith("background:"))
                    svgFill = s["background:".Length..].Trim();
                else if (s.StartsWith("background-image:") || s.StartsWith("border"))
                    continue; // gradient/image fill or border handled by SVG paths
                else
                    outerStyles.Add(s);
            }
            if (!string.IsNullOrEmpty(shapeHrefUrl)) outerStyles.Add("cursor:pointer");
            sb.Append($"    <div class=\"{shapeClass}\"{dataPathAttr}{textWarpAttr} style=\"{string.Join(";", outerStyles)}\">");

            // A gradient fill can't be expressed as a single SVG `fill` color
            // (CssSanitizeColor would reject "linear-gradient(...)" → transparent →
            // invisible shape). Build a real SVG <linearGradient> def and reference it
            // via url(#id); degrade to the first stop color if the gradient has <2 stops.
            var fillGrad = shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
            string gradFillDef = "";
            string fillAttr = CssSanitizeColor(svgFill);
            if (fillGrad != null)
            {
                var gid = $"cgg{_markerCounter++}";
                gradFillDef = BuildSvgLinearGradient(fillGrad, gid, themeColors, out var fs);
                if (!string.IsNullOrEmpty(gradFillDef)) fillAttr = $"url(#{gid})";
                else if (!string.IsNullOrEmpty(fs)) fillAttr = CssSanitizeColor(fs);
            }
            var drawFill = !string.IsNullOrEmpty(custGeomSvgFillD) && (svgFill != "none" || fillGrad != null);
            sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\" viewBox=\"0 0 {custGeomSvgW} {custGeomSvgH}\" preserveAspectRatio=\"none\">");
            if (!string.IsNullOrEmpty(gradFillDef))
                sb.Append($"<defs>{gradFillDef}</defs>");
            if (drawFill)
                sb.Append($"<path d=\"{custGeomSvgFillD}\" fill=\"{fillAttr}\" fill-rule=\"evenodd\"/>");
            if (!string.IsNullOrEmpty(custGeomSvgStrokeD) && parsedOutline != null)
            {
                var (bw, dt, bc, cap, _, join) = parsedOutline.Value;
                var dashArr = DashTypeToSvgDasharray(dt, bw);
                var dashAttr = !string.IsNullOrEmpty(dashArr) ? $" stroke-dasharray=\"{dashArr}\"" : "";
                var linecap = CapToSvgLinecap(cap);
                sb.Append($"<path d=\"{custGeomSvgStrokeD}\" fill=\"none\" stroke=\"{CssSanitizeColor(bc)}\" stroke-width=\"{bw:0.##}pt\" vector-effect=\"non-scaling-stroke\" stroke-linecap=\"{linecap}\" stroke-linejoin=\"{join}\"{dashAttr}/>");
            }
            sb.Append("</svg>");
        }
        else if (!string.IsNullOrEmpty(clipPathCss))
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
            sb.Append($"    <div class=\"{shapeClass}\"{dataPathAttr}{textWarpAttr} style=\"{string.Join(";", outerStyles)}\">");
            // Fill layer (clipped). Emit even when the shape has no explicit fill
            // background so the clip-path silhouette is still present in the HTML
            // (a snip/clip shape with no fill must still show its clipped outline;
            // previously the clip-path was dropped entirely when fillStyles was empty).
            if (fillStyles.Count > 0)
                sb.Append($"<div style=\"position:absolute;inset:0;{clipPathCss};{string.Join(";", fillStyles)}\"></div>");
            else
                sb.Append($"<div style=\"position:absolute;inset:0;{clipPathCss}\"></div>");
            // Border layer for clip-path shapes: always use SVG polygon stroke
            if (parsedOutline != null && clipPathCss.StartsWith("clip-path:polygon("))
            {
                var (bw, dt, bc, cap, _, join) = parsedOutline.Value;
                var polyStr = clipPathCss["clip-path:polygon(".Length..^1];
                var svgPoints = polyStr.Replace("%", "");
                var dashArr = DashTypeToSvgDasharray(dt, bw);
                var dashAttr = !string.IsNullOrEmpty(dashArr) ? $" stroke-dasharray=\"{dashArr}\"" : "";
                var linecap = CapToSvgLinecap(cap);
                var safeColor = CssSanitizeColor(bc);
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\">");
                sb.Append($"<polygon points=\"{svgPoints}\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" vector-effect=\"non-scaling-stroke\" stroke-linecap=\"{linecap}\" stroke-linejoin=\"{join}\"{dashAttr}/>");
                sb.Append("</svg>");
            }
        }
        else
        {
            if (!string.IsNullOrEmpty(shapeHrefUrl)) styles.Add("cursor:pointer");
            sb.Append($"    <div class=\"{shapeClass}\"{dataPathAttr}{textWarpAttr} style=\"{string.Join(";", styles)}\">");
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

            // Vertical text: <a:bodyPr vert="vert|vert270|eaVert|wordArtVert|mongolianVert"/>
            // rotates the text flow into a column. Map to CSS writing-mode so HTML
            // preview lays text out vertically like PowerPoint instead of as a
            // normal horizontal line.
            //   vert       — top-to-bottom, glyphs rotated 90° CW   → vertical-rl
            //   eaVert     — East-Asian top-to-bottom (upright)     → vertical-rl + text-orientation:upright
            //   vert270    — bottom-to-top, glyphs rotated 90° CCW  → vertical-rl + rotate(180deg)
            //   mongolianVert — left-to-right columns               → vertical-lr (best-effort)
            //   wordArtVert   — stacked upright                     → vertical-rl + upright (best-effort)
            string vertStyle = "";
            var vertVal = bodyPr?.Vertical?.HasValue == true ? bodyPr.Vertical.InnerText : null;
            switch (vertVal)
            {
                case "vert":
                    vertStyle = "writing-mode:vertical-rl;";
                    break;
                case "eaVert":
                    vertStyle = "writing-mode:vertical-rl;text-orientation:upright;";
                    break;
                case "vert270":
                    vertStyle = "writing-mode:vertical-rl;transform:rotate(180deg);";
                    break;
                case "mongolianVert":
                    vertStyle = "writing-mode:vertical-lr;text-orientation:upright;";
                    break;
                case "wordArtVert":
                case "wordArtVertRtl":
                    vertStyle = "writing-mode:vertical-rl;text-orientation:upright;";
                    break;
            }

            // Text-body rotation: <a:bodyPr rot="..."/> rotates the whole text body
            // about the text-box center (units = 60000ths of a degree). This is an
            // ADDITIONAL rotation on top of any shape-container rotation (xfrm rot),
            // and does NOT change the shape box rectangle — only the text rotates,
            // overflowing the box like PowerPoint. `vert` (writing-mode) and `rot`
            // are normally mutually exclusive; if `vert270` already set a transform
            // we append the bodyPr rotation rather than overwrite it.
            string rotStyle = "";
            if (bodyPr?.Rotation?.HasValue == true && bodyPr.Rotation.Value != 0)
            {
                var rotDeg = bodyPr.Rotation.Value / 60000.0;
                if (vertStyle.Contains("transform:rotate("))
                {
                    // vert270 path: compose with its existing rotate(180deg)
                    vertStyle = vertStyle.Replace("transform:rotate(180deg);",
                        $"transform:rotate({(180 + rotDeg):0.##}deg);transform-origin:center center;");
                }
                else
                {
                    rotStyle = $"transform:rotate({rotDeg:0.##}deg);transform-origin:center center;";
                }
            }

            // Horizontal block centering: <a:bodyPr anchorCtr="1"/> centers the
            // text BLOCK horizontally within the text frame, independent of each
            // paragraph's algn. PowerPoint centers short left-aligned text in a wide
            // shape. The `.shape-text` flex column stretches children full-width by
            // default; add `align-items:center` AND the `anchor-ctr` class (whose CSS
            // rule shrink-wraps the otherwise width:100% .para divs) so the block
            // visibly centers like PowerPoint instead of staying left-flush.
            bool anchorCtr = bodyPr?.AnchorCenter?.Value == true;
            string anchorCtrStyle = anchorCtr ? "align-items:center;" : "";

            // wrap=none: suppress the .shape inherited `white-space:pre-wrap`
            // on the inner text container so the line extends horizontally
            // rather than wrapping inside the shape's width box.
            var wrapNoneStyle = wrapNone ? "white-space:nowrap;overflow:visible;" : "";

            // Multi-column text: <a:bodyPr numCol="N" spcCol="EMU"/> lays the
            // text body out in N columns. CSS column-count is INERT on a flex
            // container, and `.shape-text` is display:flex (for valign). So the
            // columns CSS must live on a BLOCK-level wrapper INSIDE the flex div
            // — the flex parent still handles vertical anchoring while the inner
            // block establishes the multi-column formatting context.
            // column-fill:auto + a BOUNDED height make PowerPoint's newspaper-style
            // sequential fill (column 1 filled top-to-bottom before column 2) instead
            // of the CSS default `balance` (even split). The wrapper is a block child
            // of the display:flex .shape-text; a flex child's height:100% does not
            // reliably resolve into a definite height for column-fill, so we set an
            // explicit height equal to the shape's text content box (shape ext height
            // minus the top+bottom body insets), in pt — the same insets used above
            // for the text-frame padding.
            string columnStyle = "";
            if (bodyPr?.ColumnCount?.HasValue == true && bodyPr.ColumnCount.Value > 1)
            {
                columnStyle = $"column-count:{bodyPr.ColumnCount.Value};column-fill:auto;";
                if (bodyPr.ColumnSpacing?.HasValue == true)
                    columnStyle += $"column-gap:{Units.EmuToPt(bodyPr.ColumnSpacing.Value):0.##}pt;";
                var contentHeightEmu = cy - tIns - bIns;
                if (contentHeightEmu > 0)
                    columnStyle += $"height:{Units.EmuToPt(contentHeightEmu):0.##}pt;";
            }

            var textStyle = !string.IsNullOrEmpty(flipStyle) || !string.IsNullOrEmpty(clipPathCss) || !string.IsNullOrEmpty(rtlColStyle) || !string.IsNullOrEmpty(wrapNoneStyle) || !string.IsNullOrEmpty(vertStyle) || !string.IsNullOrEmpty(rotStyle) || !string.IsNullOrEmpty(anchorCtrStyle)
                ? $" style=\"{flipStyle}{rtlColStyle}{vertStyle}{rotStyle}{wrapNoneStyle}{anchorCtrStyle}{(string.IsNullOrEmpty(clipPathCss) ? "" : "position:relative;")}\""
                : "";
            var anchorCtrClass = anchorCtr ? " anchor-ctr" : "";
            // Vertical text (writing-mode) rotates the flex main axis to horizontal,
            // so valign-* (justify-content) now anchors the text column HORIZONTALLY
            // — which matches PowerPoint (anchor="ctr" centers a vert column across
            // the frame width). But a width:100% .para fills that axis and blocks the
            // centering, so flag the shape to shrink-wrap its paragraphs (same trick
            // the anchor-ctr CSS uses).
            var vertTextClass = !string.IsNullOrEmpty(vertStyle) ? " has-vert-text" : "";
            sb.Append($"<div class=\"shape-text valign-{valign}{anchorCtrClass}{vertTextClass}\"{textStyle}>");

            // Block-level column wrapper: column-count works here (normal block
            // formatting context), unlike on the flex .shape-text parent.
            if (!string.IsNullOrEmpty(columnStyle))
                sb.Append($"<div class=\"text-columns\" style=\"display:block;width:100%;{columnStyle}\">");

            // R11-3: pass the shape's <p:style>/<a:fontRef> schemeClr down as the
            // final fallback run color (used only when no explicit/inherited color).
            var fontRefDefaultColor = ResolveStyleRefSchemeColor(shape.ShapeStyle?.FontReference, themeColors);
            RenderTextBody(sb, shape.TextBody, themeColors, shape, part, fontRefDefaultColor, slideNumber);

            if (!string.IsNullOrEmpty(columnStyle))
                sb.Append("</div>");
            sb.Append("</div>");
        }

        // R23-B: SVG gradient-stroke overlay for rounded / clipped gradient outlines.
        // border-image can't follow border-radius (CSS ignores it), so stroke a
        // geometry-matching shape with stroke=url(#grad) — same approach the
        // connector gradient stroke uses.
        if (gradOutlineSvg != null)
        {
            var sep = gradOutlineSvg.IndexOf('|');
            var strokeRef = gradOutlineSvg[..sep];
            var gradDef = gradOutlineSvg[(sep + 1)..];
            var bw = gradOutlineW;
            sb.Append("<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\"><defs>");
            sb.Append(gradDef);
            sb.Append("</defs>");
            if (!string.IsNullOrEmpty(clipPathCss) && clipPathCss.StartsWith("clip-path:polygon("))
            {
                var polyStr = clipPathCss["clip-path:polygon(".Length..^1];
                var svgPoints = polyStr.Replace("%", "");
                // viewBox 0..100 + non-scaling-stroke keeps the polygon in % space.
                sb.Append("</svg>");
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\"><defs>{gradDef}</defs>");
                sb.Append($"<polygon points=\"{svgPoints}\" fill=\"none\" stroke=\"{strokeRef}\" stroke-width=\"{bw:0.##}pt\" vector-effect=\"non-scaling-stroke\"/>");
            }
            else
            {
                var rxMatch = System.Text.RegularExpressions.Regex.Match(borderRadiusCss, @"border-radius:([\d.]+)");
                var rx = rxMatch.Success ? rxMatch.Groups[1].Value : "0";
                sb.Append($"<rect x=\"{bw / 2:0.##}pt\" y=\"{bw / 2:0.##}pt\" width=\"calc(100% - {bw:0.##}pt)\" height=\"calc(100% - {bw:0.##}pt)\" rx=\"{rx}pt\" ry=\"{rx}pt\" fill=\"none\" stroke=\"{strokeRef}\" stroke-width=\"{bw:0.##}pt\"/>");
            }
            sb.Append("</svg>");
        }

        // SVG border overlay for non-solid outlines (dashed, dotted, dashDot etc.)
        if (parsedOutline != null && parsedOutline.Value.dashType != "solid")
        {
            var (bw, dt, bc, cap, _, join) = parsedOutline.Value;
            var dashArr = DashTypeToSvgDasharray(dt, bw);
            var dashAttr = !string.IsNullOrEmpty(dashArr) ? $" stroke-dasharray=\"{dashArr}\"" : "";
            var linecap = CapToSvgLinecap(cap);
            var linejoinAttr = $" stroke-linejoin=\"{join}\"";
            var safeColor = CssSanitizeColor(bc);

            if (!string.IsNullOrEmpty(clipPathCss) && clipPathCss.StartsWith("clip-path:polygon("))
            {
                // Polygon shapes — reuse existing polygon SVG approach
                var polyStr = clipPathCss["clip-path:polygon(".Length..^1];
                var svgPoints = polyStr.Replace("%", "");
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\">");
                sb.Append($"<polygon points=\"{svgPoints}\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" vector-effect=\"non-scaling-stroke\" stroke-linecap=\"{linecap}\"{linejoinAttr}{dashAttr}/>");
                sb.Append("</svg>");
            }
            else if (!string.IsNullOrEmpty(borderRadiusCss))
            {
                // Rounded rect — use SVG rect with rx/ry
                var rxMatch = System.Text.RegularExpressions.Regex.Match(borderRadiusCss, @"border-radius:([\d.]+)");
                var rx = rxMatch.Success ? rxMatch.Groups[1].Value : "0";
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
                sb.Append($"<rect x=\"{bw / 2:0.##}pt\" y=\"{bw / 2:0.##}pt\" width=\"calc(100% - {bw:0.##}pt)\" height=\"calc(100% - {bw:0.##}pt)\" rx=\"{rx}pt\" ry=\"{rx}pt\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" stroke-linecap=\"{linecap}\"{linejoinAttr}{dashAttr}/>");
                sb.Append("</svg>");
            }
            else if (presetGeom?.Preset?.InnerText == "ellipse")
            {
                // Ellipse — size in pt so stroke-width matches CSS border path.
                // CONSISTENCY(shape-stroke-unit): keep stroke-width in pt across solid/non-solid paths.
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
                sb.Append($"<ellipse cx=\"50%\" cy=\"50%\" rx=\"calc(50% - {bw / 2:0.##}pt)\" ry=\"calc(50% - {bw / 2:0.##}pt)\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" stroke-linecap=\"{linecap}\"{dashAttr}/>");
                sb.Append("</svg>");
            }
            else
            {
                // Plain rect — use SVG rect sized in pt so stroke-width matches the CSS
                // `border:Npt solid` path (same visual weight). Inset by bw/2 so the stroke
                // sits entirely inside the content box (box-sizing:border-box equivalent).
                // CONSISTENCY(shape-stroke-unit): keep stroke-width in pt across solid/non-solid paths.
                sb.Append($"<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
                sb.Append($"<rect x=\"{bw / 2:0.##}pt\" y=\"{bw / 2:0.##}pt\" width=\"calc(100% - {bw:0.##}pt)\" height=\"calc(100% - {bw:0.##}pt)\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" stroke-linecap=\"{linecap}\"{linejoinAttr}{dashAttr}/>");
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
    /// R12-5: find the layout (then master) placeholder shape that the given
    /// slide placeholder inherits from. Same ph type/idx matching as
    /// ResolveInheritedPosition, but returns the whole shape so callers can
    /// read inherited spPr fill/etc. Returns null for non-placeholders.
    /// </summary>
    private static Shape? ResolveInheritedPlaceholderShape(Shape shape, OpenXmlPart part)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return null;

        var slidePart = part as SlidePart;
        if (slidePart == null) return null;

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
                if (PlaceholderMatches(ph, candidatePh)) return candidate;
            }
        }

        return null;
    }

    /// <summary>
    /// Resolve the text vertical anchor (<a:bodyPr anchor="…">) for a placeholder
    /// shape whose own bodyPr declares no anchor, walking the layout then master
    /// matching placeholder (same ph type/idx matching as ResolveInheritedPosition).
    /// Returns the first ancestor bodyPr that carries an anchor attribute
    /// ("t" | "ctr" | "b"), or null if none specify one.
    /// </summary>
    private static string? ResolveInheritedAnchor(Shape shape, OpenXmlPart part)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return null;

        var slidePart = part as SlidePart;
        if (slidePart == null) return null;

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

                var candBodyPr = candidate.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                if (candBodyPr?.Anchor?.HasValue == true)
                    return candBodyPr.Anchor.InnerText;
                // Matched this ancestor but it has no anchor → fall through to
                // the next ancestor (layout matched but silent → consult master).
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

        // R26-5: slide ph has idx but NO type, layout ph has type but NO idx.
        // OOXML: a <p:ph idx=N/> with no type defaults to type=body, so it
        // should inherit from a type=body (or object) layout/master placeholder.
        // Without this branch all inheritance silently drops for idx-only slide
        // placeholders bound to a typed layout placeholder.
        if (slidePh.Index?.HasValue == true && slidePh.Type?.HasValue != true
            && layoutPh.Type?.HasValue == true && layoutPh.Index?.HasValue != true)
        {
            var lt = layoutPh.Type.Value;
            return lt == PlaceholderValues.Body || lt == PlaceholderValues.Object;
        }

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
        "ellipse" => (0.15, 0.15, 0.15, 0.15),
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
    /// R7-2: Resolve the inherited default run COLOR for a placeholder shape, walking the
    /// same inheritance chain as ResolvePlaceholderFontSize (shape lstStyle → layout/master
    /// placeholder → master text styles). Returns the defRPr solidFill resolved to a hex/theme
    /// CSS color, or null if no layer carries a color. Mirrors the size resolver so a master
    /// bodyStyle defRPr solidFill propagates to body placeholders in HTML preview.
    /// </summary>
    private static string? ResolvePlaceholderDefaultColor(Shape shape, OpenXmlPart part,
        Dictionary<string, string> themeColors, int level = 0)
    {
        var defRp = ResolvePlaceholderDefRp(shape, part, level,
            dr => dr.GetFirstChild<Drawing.SolidFill>() != null);
        var fill = defRp?.GetFirstChild<Drawing.SolidFill>();
        return ResolveFillColor(fill, themeColors);
    }

    /// <summary>
    /// R7-3: Resolve the inherited default line-spacing for a placeholder shape from the same
    /// inheritance chain. Returns the master/layout defRPr-level lnSpc as a CSS line-height
    /// fragment (e.g. "line-height:2") or null when no layer specifies lnSpc.
    /// </summary>
    private static string? ResolvePlaceholderLineSpacing(Shape shape, OpenXmlPart part, int level = 0)
    {
        var lvlPpr = ResolvePlaceholderLevelPpr(shape, part, level,
            p => p.GetFirstChild<Drawing.LineSpacing>() != null);
        var lnSpc = lvlPpr?.GetFirstChild<Drawing.LineSpacing>();
        if (lnSpc == null) return null;
        var pct = lnSpc.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
        if (pct.HasValue) return $"line-height:{pct.Value / 100000.0:0.##}";
        var pts = lnSpc.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
        if (pts.HasValue) return $"line-height:{pts.Value / 100.0:0.##}pt";
        return null;
    }

    /// <summary>
    /// Shared inheritance walk for placeholder default-run-property resolution. Walks
    /// shape lstStyle → layout/master matching placeholder lstStyle → master text styles
    /// and returns the first level paragraph-properties element matching <paramref name="match"/>.
    /// </summary>
    private static Drawing.DefaultRunProperties? ResolvePlaceholderDefRp(Shape shape, OpenXmlPart part,
        int level, Func<Drawing.DefaultRunProperties, bool> match)
    {
        var lvlPpr = ResolvePlaceholderLevelPpr(shape, part, level,
            p => p.GetFirstChild<Drawing.DefaultRunProperties>() is { } dr && match(dr));
        return lvlPpr?.GetFirstChild<Drawing.DefaultRunProperties>();
    }

    /// <summary>
    /// Shared inheritance walk that returns the first level-paragraph-properties element
    /// (Level1ParagraphProperties etc.) in the placeholder chain matching <paramref name="match"/>.
    /// </summary>
    private static OpenXmlElement? ResolvePlaceholderLevelPpr(Shape shape, OpenXmlPart part,
        int level, Func<OpenXmlElement, bool> match)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return null;

        // 1. Shape's own list style
        var lstStyle = shape.TextBody?.GetFirstChild<Drawing.ListStyle>();
        if (GetLevelPpr(lstStyle, level) is { } own && match(own)) return own;

        var phType = ph.Type?.HasValue == true ? ph.Type.Value : PlaceholderValues.Body;
        bool isTitle = phType == PlaceholderValues.Title || phType == PlaceholderValues.CenteredTitle;
        bool isSubTitle = phType == PlaceholderValues.SubTitle;

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
                    var cLstStyle = candidate.TextBody?.GetFirstChild<Drawing.ListStyle>();
                    if (GetLevelPpr(cLstStyle, level) is { } cppr && match(cppr)) return cppr;
                }
            }

            var masterTxStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
            if (masterTxStyles != null)
            {
                OpenXmlCompositeElement? styleList = isTitle ? masterTxStyles.TitleStyle
                    : (isSubTitle || phType == PlaceholderValues.Body || phType == PlaceholderValues.Object)
                        ? masterTxStyles.BodyStyle
                        : masterTxStyles.OtherStyle;
                if (GetLevelPpr(styleList, level) is { } sppr && match(sppr)) return sppr;
            }
        }
        return null;
    }

    /// <summary>
    /// Get the level paragraph-properties element (Level1ParagraphProperties etc.) for a
    /// given level from a list style or text style element.
    /// </summary>
    private static OpenXmlElement? GetLevelPpr(OpenXmlCompositeElement? styleList, int level)
    {
        if (styleList == null) return null;
        return level switch
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
    /// Compute the hue angle (degrees, 0-360) of an RRGGBB hex color. Used to
    /// approximate a duotone recolor in CSS via sepia + hue-rotate toward the
    /// target color's hue. sepia produces a brownish base (~hue 35°), so the
    /// emitted rotation is relative to that.
    /// </summary>
    private static double RgbHexToHueDeg(string hex)
    {
        var clean = hex.TrimStart('#');
        if (clean.Length < 6) return 0.0;
        var (r, g, b) = ColorMath.HexToRgb(clean[..6]);
        double rf = r / 255.0, gf = g / 255.0, bf = b / 255.0;
        double max = Math.Max(rf, Math.Max(gf, bf)), min = Math.Min(rf, Math.Min(gf, bf));
        double delta = max - min;
        if (delta < 1e-6) return 0.0;
        double hue;
        if (max == rf) hue = ((gf - bf) / delta) % 6;
        else if (max == gf) hue = (bf - rf) / delta + 2;
        else hue = (rf - gf) / delta + 4;
        hue *= 60;
        if (hue < 0) hue += 360;
        // sepia base hue is ~35°; rotate from there toward the target hue.
        var rotate = hue - 35;
        if (rotate < 0) rotate += 360;
        return rotate;
    }

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

        // Picture-level hyperlink → wrap the picture <div> in <a> for clickability in
        // HTML preview. CONSISTENCY(shape-picture-parity): RenderShape already does
        // this for shapes (and NodeBuilder surfaces Format["link"] for pictures), but
        // RenderPicture dropped it — a hyperlinked image rendered un-clickable. Same
        // rules: external URLs only; internal slide-jump (ppaction://hlinksldjump) and
        // unsafe schemes (javascript:/data: …) are skipped.
        string? picHrefUrl = null;
        string? picHrefTooltip = null;
        {
            var nvHlink = pic.NonVisualPictureProperties?.NonVisualDrawingProperties
                ?.GetFirstChild<Drawing.HyperlinkOnClick>();
            if (nvHlink != null)
            {
                picHrefTooltip = nvHlink.Tooltip?.Value;
                var action = nvHlink.Action?.Value;
                var hlId = nvHlink.Id?.Value;
                if (string.IsNullOrEmpty(action) || !action.Contains("hlink"))
                {
                    if (!string.IsNullOrEmpty(hlId))
                    {
                        try
                        {
                            var rel = slidePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == hlId);
                            if (rel?.Uri != null && Core.HyperlinkUriValidator.IsSafeScheme(rel.Uri.ToString()))
                                picHrefUrl = rel.Uri.ToString();
                        }
                        catch { }
                    }
                }
            }
        }

        var styles = new List<string>
        {
            $"left:{Units.EmuToPt(x)}pt",
            $"top:{Units.EmuToPt(y)}pt",
            $"width:{Units.EmuToPt(cx)}pt",
            $"height:{Units.EmuToPt(cy)}pt"
        };

        // Transform chain (rotation + 3D) — combined into one transform property.
        var picTransforms = new List<string>();

        // Rotation
        if (xfrm.Rotation != null && xfrm.Rotation.Value != 0)
            picTransforms.Add($"rotate({xfrm.Rotation.Value / 60000.0:0.##}deg)");

        // Flip — CONSISTENCY(shape-picture-parity): mirror RenderShape's flip
        // block (rotate before scale). A flipH/flipV picture mirrors in real
        // PowerPoint, so view html must do the same.
        if (xfrm.HorizontalFlip?.Value == true && xfrm.VerticalFlip?.Value == true)
            picTransforms.Add("scale(-1,-1)");
        else if (xfrm.HorizontalFlip?.Value == true)
            picTransforms.Add("scaleX(-1)");
        else if (xfrm.VerticalFlip?.Value == true)
            picTransforms.Add("scaleY(-1)");

        // 3D rotation (scene3d camera rotation) → CSS perspective transform.
        // CONSISTENCY(shape-picture-parity): mirror RenderShape's scene3d block.
        var picScene3d = pic.ShapeProperties?.GetFirstChild<Drawing.Scene3DType>();
        var picCam = picScene3d?.Camera;
        var picRot3d = picCam?.Rotation;
        if (picRot3d != null)
        {
            var rx = (picRot3d.Latitude?.Value ?? 0) / 60000.0;
            var ry = (picRot3d.Longitude?.Value ?? 0) / 60000.0;
            var rz = (picRot3d.Revolution?.Value ?? 0) / 60000.0;
            if (rx != 0 || ry != 0 || rz != 0)
            {
                styles.Add("perspective:800px");
                if (rx != 0) picTransforms.Add($"rotateX({rx:0.##}deg)");
                if (ry != 0) picTransforms.Add($"rotateY({ry:0.##}deg)");
                if (rz != 0) picTransforms.Add($"rotateZ({rz:0.##}deg)");
            }
        }

        if (picTransforms.Count > 0)
            styles.Add($"transform:{string.Join(" ", picTransforms)}");

        // Border
        var outline = pic.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        // Parse once so a non-solid dash (dot/dashDot/lgDash...) can be rendered as an
        // accurate SVG stroke-dasharray overlay below — exactly as RenderShape does.
        // OutlineToCss collapses every non-solid dash to a generic CSS border-style
        // (dashed/dotted) that cannot express e.g. dashDot, so a picture border looked
        // wrong vs PowerPoint (and vs a sibling shape with the same a:ln).
        var parsedPicOutline = outline != null ? ParseOutline(outline, themeColors) : null;
        if (outline != null)
        {
            // Solid (or unparseable) → CSS border as before. Non-solid is deferred to
            // the SVG overlay appended just before the picture's closing </div>.
            if (parsedPicOutline == null || parsedPicOutline.Value.dashType == "solid")
            {
                var borderCss = OutlineToCss(outline, themeColors);
                if (!string.IsNullOrEmpty(borderCss))
                    styles.Add(borderCss);
            }
        }
        else
        {
            // Style-matrix lnRef fallback (parity with RenderShape): a "Picture Style"
            // preset (Picture Format → Picture Styles) encodes its border as
            // p:style/a:lnRef into the theme line-style matrix with no explicit a:ln.
            // RenderPicture previously ignored pic.ShapeStyle entirely, so a styled
            // picture rendered with no border.
            var lnRefCss = GetStyleLineRefCss(pic.ShapeStyle, slidePart, themeColors);
            if (!string.IsNullOrEmpty(lnRefCss))
                styles.Add(lnRefCss);
        }

        // Effects: brightness, contrast, glow, shadow, opacity all roll
        // into one CSS `filter` property (drop-shadow / brightness /
        // contrast) so they compose. Mirror the shape renderer above:
        // shadowCss + glowCss merged into filter:..., reflection separate.
        // Style-matrix effectRef fallback (parity with RenderShape): a Picture Style
        // preset encodes its shadow/glow as p:style/a:effectRef with no explicit
        // effectLst — resolve it the same way shapes do.
        var effectList = pic.ShapeProperties?.GetFirstChild<Drawing.EffectList>()
            ?? ResolveStyleEffectRefList(pic.ShapeStyle, slidePart);
        var shadowCss = EffectListToShadowCss(effectList, themeColors);
        var glowCss = EffectListToGlowCss(effectList, themeColors);

        // brightness / contrast — Set.Media writes <a:lum bright="N"
        // contrast="M"/> under a:blip. Tolerate legacy <a:lumMod>/<a:lumOff>
        // children written by older builds (invalid per CT_Blip but found
        // in the wild) so existing decks still preview correctly.
        var picBlipForFx = pic.BlipFill?.GetFirstChild<Drawing.Blip>();
        double? brightnessPct = null, contrastPct = null;
        if (picBlipForFx != null)
        {
            foreach (var kid in picBlipForFx.ChildElements)
            {
                if (kid.NamespaceUri != "http://schemas.openxmlformats.org/drawingml/2006/main") continue;
                if (kid is Drawing.LuminanceEffect lumElem)
                {
                    if (lumElem.Brightness?.HasValue == true) brightnessPct = lumElem.Brightness.Value / 1000.0;
                    if (lumElem.Contrast?.HasValue == true) contrastPct = lumElem.Contrast.Value / 1000.0;
                }
                else if (kid.LocalName == "lumOff" || kid.LocalName == "lumMod")
                {
                    var attr = kid.GetAttribute("val", "").Value;
                    if (string.IsNullOrEmpty(attr) || !int.TryParse(attr, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var iv)) continue;
                    if (kid.LocalName == "lumOff") brightnessPct ??= iv / 1000.0;
                    else if (kid.LocalName == "lumMod") contrastPct ??= (iv - 100000) / 1000.0;
                }
            }
        }

        // biLevel — Set.Media writes <a:biLevel thresh="N"/> under a:blip,
        // converting the picture to 1-bit black/white at threshold N. CSS has
        // no true threshold, so approximate with grayscale(1) plus a heavy
        // contrast boost to push midtones toward pure black/white.
        var picBiLevel = picBlipForFx?.GetFirstChild<Drawing.BiLevel>();

        // Duotone — Set.Media writes <a:duotone> with two color children under
        // a:blip, remapping the image's luminance gradient between the two
        // stops. CSS has no true duotone, so approximate (same philosophy as
        // the biLevel approximation): grayscale + sepia + hue-rotate toward the
        // highlight color's hue so a duotone picture is visibly tinted and
        // distinct from the untinted original.
        var picDuotone = picBlipForFx?.GetFirstChild<Drawing.Duotone>();
        string? duotoneFilter = null;
        if (picDuotone != null)
        {
            // The highlight (lighter) color is the LAST color child. It may be a
            // <a:srgbClr> OR a <a:schemeClr> (PowerPoint's built-in duotone presets and
            // the CLI both write schemeClr). Resolving only srgbClr left schemeClr
            // duotones at hue 0 (sepia-orange) regardless of the actual accent — a blue
            // accent duotone previewed orange. Resolve the scheme color via themeColors.
            var lastColorEl = picDuotone.ChildElements
                .LastOrDefault(e => e is Drawing.RgbColorModelHex or Drawing.SchemeColor);
            string? highlight = lastColorEl switch
            {
                Drawing.RgbColorModelHex rgbH => rgbH.Val?.Value,
                Drawing.SchemeColor schH when schH.Val?.InnerText is string sn
                    && themeColors.TryGetValue(sn, out var shx) => shx,
                _ => null,
            };
            var hue = !string.IsNullOrEmpty(highlight) ? RgbHexToHueDeg(highlight!) : 0.0;
            duotoneFilter = $"grayscale(1) sepia(1) saturate(3) hue-rotate({hue:0.#}deg)";
        }

        // Grayscale — Set.Media writes a bare <a:grayscl/> under a:blip,
        // converting the picture to luminance grayscale (Picture Format →
        // Color → Recolor → Grayscale). Maps directly to CSS grayscale(1).
        // grayscl is mutually exclusive with duotone/biLevel in practice;
        // skip if duotone already supplies a richer (grayscale-inclusive)
        // filter so we don't emit two conflicting filters.
        var picGrayscl = picBlipForFx?.GetFirstChild<Drawing.Grayscale>();

        var filterParts = new List<string>();
        var boxShadowParts = new List<string>();
        if (picBiLevel != null)
            filterParts.Add("grayscale(1) contrast(1000%)");
        if (duotoneFilter != null)
            filterParts.Add(duotoneFilter);
        if (picGrayscl != null && duotoneFilter == null && picBiLevel == null)
            filterParts.Add("grayscale(1)");
        // CSS brightness(1) = no change; +N% brightness → brightness(1 + N/100).
        if (brightnessPct.HasValue && Math.Abs(brightnessPct.Value) > 0.01)
            filterParts.Add($"brightness({1 + brightnessPct.Value / 100.0:0.###})");
        // CSS contrast(1) = no change; +N% contrast → contrast(1 + N/100).
        if (contrastPct.HasValue && Math.Abs(contrastPct.Value) > 0.01)
            filterParts.Add($"contrast({1 + contrastPct.Value / 100.0:0.###})");
        // innerShdw returns "box-shadow:inset ..." (no CSS filter equivalent); route
        // it to boxShadowParts. Outer/preset shadow returns a "filter:..." value.
        if (!string.IsNullOrEmpty(shadowCss))
        {
            if (shadowCss.StartsWith("box-shadow:"))
                boxShadowParts.Add(shadowCss["box-shadow:".Length..]);
            else
                filterParts.Add(shadowCss.Replace("filter:", ""));
        }
        if (!string.IsNullOrEmpty(glowCss))
            filterParts.Add(glowCss.Replace("filter:", ""));
        var picBlurCss = EffectListToBlurCss(effectList);
        if (!string.IsNullOrEmpty(picBlurCss))
            filterParts.Add(picBlurCss);
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

        // Soft edge → fade out at edges using CSS mask-image.
        // CONSISTENCY(shape-picture-parity): mirror RenderShape's softEdge block.
        var picSoftEdge = effectList?.GetFirstChild<Drawing.SoftEdge>();
        if (picSoftEdge?.Radius?.HasValue == true)
        {
            var edgePx = Math.Max(2, picSoftEdge.Radius.Value / EmuConverter.EmuPerPointF * 0.8);
            styles.Add($"-webkit-mask-image:linear-gradient(to right,transparent 0,black {edgePx:0.#}px,black calc(100% - {edgePx:0.#}px),transparent 100%)," +
                       $"linear-gradient(to bottom,transparent 0,black {edgePx:0.#}px,black calc(100% - {edgePx:0.#}px),transparent 100%)");
            styles.Add("-webkit-mask-composite:source-in;mask-composite:intersect");
        }

        // Bevel → approximate with inset box-shadow for a subtle 3D appearance.
        // CONSISTENCY(shape-picture-parity): mirror RenderShape's sp3d block.
        var picSp3d = pic.ShapeProperties?.GetFirstChild<Drawing.Shape3DType>();
        if (picSp3d?.BevelTop != null)
        {
            var bevelW = picSp3d.BevelTop.Width?.HasValue == true ? picSp3d.BevelTop.Width.Value / EmuConverter.EmuPerPointF : 6;
            var bW = Math.Max(1, bevelW * 0.5);
            boxShadowParts.Add($"inset {bW:0.#}px {bW:0.#}px {bW * 1.5:0.#}px rgba(255,255,255,0.25),inset -{bW:0.#}px -{bW:0.#}px {bW * 1.5:0.#}px rgba(0,0,0,0.15)");
        }

        // Emit one combined box-shadow (inner shadow + bevel). CSS only honors the
        // last box-shadow in an inline style, so they must share a single declaration.
        if (boxShadowParts.Count > 0)
            styles.Add($"box-shadow:{string.Join(",", boxShadowParts)}");

        // Geometry (rounded corners)
        var presetGeom = pic.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
        {
            var geomCss = PresetGeometryToCss(presetGeom.Preset!.InnerText!, cx, cy, presetGeom);
            if (!string.IsNullOrEmpty(geomCss))
                styles.Add(geomCss);
        }

        // Open <a> wrapper for picture-level hyperlink (before the picture <div>).
        if (!string.IsNullOrEmpty(picHrefUrl))
        {
            var tooltipAttr = !string.IsNullOrEmpty(picHrefTooltip)
                ? $" title=\"{HtmlEncode(picHrefTooltip!)}\"" : "";
            sb.Append($"    <a class=\"shape-link\" href=\"{HtmlEncode(picHrefUrl!)}\" rel=\"noopener\" target=\"_blank\"{tooltipAttr} style=\"display:contents;cursor:pointer\">");
        }

        sb.Append($"    <div class=\"picture\"{dataPathAttr} style=\"{string.Join(";", styles)}\">");

        // Extract image data
        var blipFill = pic.BlipFill;
        var blip = blipFill?.GetFirstChild<Drawing.Blip>();
        // R4-1: prefer the asvg:svgBlip extension's rel-id over the raster
        // fallback (blip.Embed → 1x1 PNG) so an SVG picture renders its real
        // vector artwork instead of a blank 1x1 placeholder.
        var picSvgRelId = blip != null ? OfficeCli.Core.SvgImageHelper.GetSvgRelId(blip) : null;
        var picEmbedId = !string.IsNullOrEmpty(picSvgRelId) ? picSvgRelId : blip?.Embed?.Value;
        if (!string.IsNullOrEmpty(picEmbedId))
        {
            try
            {
                var imgPart = slidePart.GetPartById(picEmbedId!);
                using var stream = imgPart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                var base64 = Convert.ToBase64String(ms.ToArray());
                var contentType = !string.IsNullOrEmpty(picSvgRelId)
                    ? "image/svg+xml"
                    : SanitizeContentType(imgPart.ContentType ?? "image/png");

                // EMF/WMF are vector metafiles browsers cannot display via <img> or
                // background-image — emitting a data:image/x-emf URI shows a broken-image
                // icon. Emit the graceful placeholder used for other unrenderable images.
                if (contentType is "image/x-emf" or "image/x-wmf" or "image/emf" or "image/wmf")
                {
                    sb.Append("<div style=\"width:100%;height:100%;background:rgba(128,128,128,0.15);display:flex;align-items:center;justify-content:center;color:rgba(128,128,128,0.5);font-size:12px\">Image</div>");
                }
                else
                {

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
                // R4-4: a <a:tile> blip fill repeats the image at its native size
                // rather than stretching it to cover. Emit a repeating background
                // (background-repeat:repeat; background-size:auto) instead of the
                // default stretched <img>; previously tile was ignored and the
                // image rendered stretched-to-fit.
                var tile = blipFill?.GetFirstChild<Drawing.Tile>();
                // R49-02: <a:stretch><a:fillRect l/t/r/b> insets the stretched
                // image from each edge (1/1000 percent; can be negative=outset).
                // PowerPoint scales the image into the inner rect, leaving the
                // border transparent. CSS: background-size = remaining %, and
                // background-position derived from the l/t vs r/b split.
                var stretch = blipFill?.GetFirstChild<Drawing.Stretch>();
                var fillRect = stretch?.FillRectangle;
                double frL = (fillRect?.Left?.Value ?? 0) / 1000.0;
                double frT = (fillRect?.Top?.Value ?? 0) / 1000.0;
                double frR = (fillRect?.Right?.Value ?? 0) / 1000.0;
                double frB = (fillRect?.Bottom?.Value ?? 0) / 1000.0;
                var hasFillRectInset = fillRect != null && (frL != 0 || frT != 0 || frR != 0 || frB != 0);
                // Degenerate crop: L+R >= 100% or T+B >= 100% means zero/negative
                // visible area. PowerPoint renders nothing in this case; HTML
                // preview previously averaged the background-image into a muddy
                // block. Skip the picture draw entirely to match real PPT.
                var degenerateCrop = hasCrop && (srcL + srcR >= 1.0 || srcT + srcB >= 1.0);
                if (degenerateCrop)
                {
                    // Render nothing — matches PowerPoint's zero-area behavior.
                }
                else if (tile != null)
                {
                    // R49-01: honor <a:tile sx/sy> scale + algn. sx/sy are
                    // ×1000 percent (30000 → 30%) of the image's NATIVE pixel
                    // size. CSS background-size:30% is 30% of the container, not
                    // the image, so a percentage cannot reproduce PowerPoint's
                    // semantics — instead compute pixel dimensions from the
                    // decoded image's natural size × ratio. When sx/sy are absent
                    // keep background-size:auto (native size = the 100% default).
                    var sx = tile.HorizontalRatio?.Value;
                    var sy = tile.VerticalRatio?.Value;
                    string bgSize = "auto";
                    // 100% (or absent) == native size == auto; only a non-100%
                    // ratio needs an explicit scaled background-size.
                    var nonDefaultScale = (sx.HasValue && sx.Value != 100000)
                        || (sy.HasValue && sy.Value != 100000);
                    if (nonDefaultScale)
                    {
                        var nat = OfficeCli.Core.ImageSource.TryGetDimensions(ms);
                        var sxPct = (sx ?? 100000) / 100000.0;
                        var syPct = (sy ?? 100000) / 100000.0;
                        if (nat != null)
                            bgSize = $"{nat.Value.Width * sxPct:0.##}px {nat.Value.Height * syPct:0.##}px";
                        else
                            // No natural size: fall back to percent-of-container
                            // (approximate, but still distinct from native auto).
                            bgSize = $"{sxPct * 100:0.##}% {syPct * 100:0.##}%";
                    }
                    var bgPos = TileAlignToBackgroundPosition(tile.Alignment);
                    sb.Append($"<div style=\"width:100%;height:100%;background-image:url(data:{contentType};base64,{base64});background-repeat:repeat;background-position:{bgPos};background-size:{bgSize}\"></div>");
                }
                else if (hasFillRectInset)
                {
                    // Image occupies the inner rect (100-l-r) × (100-t-b);
                    // negative insets bleed outside and are clipped by the
                    // .picture div's overflow. Position by the l vs r ratio.
                    var sizeW = Math.Max(100.0 - frL - frR, 0.01);
                    var sizeH = Math.Max(100.0 - frT - frB, 0.01);
                    var posDenomX = frL + frR;
                    var posDenomY = frT + frB;
                    var posX = posDenomX != 0 ? frL / posDenomX * 100.0 : 0.0;
                    var posY = posDenomY != 0 ? frT / posDenomY * 100.0 : 0.0;
                    var bgStyle = $"width:100%;height:100%;overflow:hidden;background-image:url(data:{contentType};base64,{base64});background-repeat:no-repeat;background-size:{sizeW:0.##}% {sizeH:0.##}%;background-position:{posX:0.##}% {posY:0.##}%";
                    sb.Append($"<div style=\"{bgStyle}\"></div>");
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
                } // end else (renderable content type)
            }
            catch
            {
                // Image extraction failed - show placeholder
                sb.Append("<div style=\"width:100%;height:100%;background:rgba(128,128,128,0.15);display:flex;align-items:center;justify-content:center;color:rgba(128,128,128,0.5);font-size:12px\">Image</div>");
            }
        }
        else if (blip?.Link?.Value is { Length: > 0 })
        {
            // R15-3: a linked picture (<a:blip r:link="...">, no r:embed) references an
            // external image we do not fetch. Mirror the OLE/3D placeholder pattern and
            // emit a grey "Linked image" surface so the shape is visible rather than an
            // empty div. We deliberately do NOT resolve/download the external target.
            sb.Append("<div style=\"width:100%;height:100%;background:rgba(128,128,128,0.15);display:flex;align-items:center;justify-content:center;color:rgba(128,128,128,0.5);font-size:12px\">Linked image</div>");
        }

        // Non-solid outline (dash/dot/dashDot/lgDash...) → accurate SVG stroke-dasharray
        // overlay, mirroring RenderShape's plain-rect branch. A picture is always a plain
        // rectangle, so only that branch is needed. The stroke is inset by bw/2 so it sits
        // inside the content box (matching the solid CSS-border path's visual weight).
        if (parsedPicOutline != null && parsedPicOutline.Value.dashType != "solid")
        {
            var (bw, dt, bc, cap, _, join) = parsedPicOutline.Value;
            var dashArr = DashTypeToSvgDasharray(dt, bw);
            var dashAttr = !string.IsNullOrEmpty(dashArr) ? $" stroke-dasharray=\"{dashArr}\"" : "";
            var linecap = CapToSvgLinecap(cap);
            var safeColor = CssSanitizeColor(bc);
            sb.Append("<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
            sb.Append($"<rect x=\"{bw / 2:0.##}pt\" y=\"{bw / 2:0.##}pt\" width=\"calc(100% - {bw:0.##}pt)\" height=\"calc(100% - {bw:0.##}pt)\" fill=\"none\" stroke=\"{safeColor}\" stroke-width=\"{bw:0.##}pt\" stroke-linecap=\"{linecap}\" stroke-linejoin=\"{join}\"{dashAttr}/>");
            sb.Append("</svg>");
        }

        sb.AppendLine("</div>");

        // Close <a> wrapper for picture-level hyperlink.
        if (!string.IsNullOrEmpty(picHrefUrl))
            sb.Append("</a>");
    }

    // ==================== Connector Rendering ====================

    private static void RenderConnector(StringBuilder sb, ConnectionShape cxn, Dictionary<string, string> themeColors, string? dataPath = null,
        (long x, long y, long cx, long cy)? overridePos = null, OpenXmlPart? part = null)
        // R15-1: resolve the connector's txBody (handles the SDK OpenXmlUnknownElement
        // parse) and forward it so the label renders as a centered overlay on the line.
        // R234: forward ShapeStyle + part so a style-only connector (no <a:ln>) resolves
        // its lnRef stroke color/width from the theme line-style matrix.
        => RenderConnector(sb, cxn.ShapeProperties, themeColors, dataPath, overridePos,
            ResolveConnectorTextBody(cxn), cxn.ShapeStyle, part);

    // Shared SVG line/polyline/path renderer for both <p:cxnSp> connectors and
    // <p:sp> shapes with prst="line". Reads geometry + outline from a
    // ShapeProperties and emits a connector-style div.
    // overridePos: when rendering inside a group, the caller supplies coordinates
    // already transformed into the group's child coordinate system (see RenderShape /
    // RenderPicture's parallel parameter). Without it the connector's raw slide-absolute
    // EMU coords are emitted as offsets from the group container — placing the line
    // far outside the group div, where it disappears.
    private static void RenderConnector(StringBuilder sb, ShapeProperties? spPr, Dictionary<string, string> themeColors, string? dataPath = null,
        (long x, long y, long cx, long cy)? overridePos = null,
        DocumentFormat.OpenXml.Presentation.TextBody? cxnTextBody = null,
        ShapeStyle? style = null, OpenXmlPart? part = null)
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
        var lineCap = "flat";
        var lineCmpd = "sng";
        // R15-2: a connector outline <a:ln>/<a:gradFill> can't be represented by a
        // single solid stroke; mirror the shape outline path (line ~167) by building
        // an SVG <linearGradient> def from the stops and stroking with url(#id). When
        // the gradient is present, gradStrokeId/gradStrokeDef are populated and the
        // solid stroke color is set to the first stop (fallback for renderers that
        // can't resolve the def). Without this the stroke fell back to theme tx1/dk1.
        string? gradStrokeId = null;
        string? gradStrokeDef = null;
        if (outline != null)
        {
            var outlineGradFill = outline.GetFirstChild<Drawing.GradientFill>();
            if (outlineGradFill != null)
            {
                gradStrokeId = $"cxg{_markerCounter++}";
                gradStrokeDef = BuildSvgLinearGradient(outlineGradFill, gradStrokeId, themeColors, out var firstStop);
                if (firstStop != null) lineColor = firstStop;
            }
            else
            {
                var c = ResolveFillColor(outline.GetFirstChild<Drawing.SolidFill>(), themeColors);
                if (c != null) lineColor = c;
            }
            if (outline.Width?.HasValue == true) lineWidth = outline.Width.Value / EmuConverter.EmuPerPointF;
            if (outline.CapType?.HasValue == true) lineCap = outline.CapType.InnerText ?? "flat";
            if (outline.CompoundLineType?.HasValue == true) lineCmpd = outline.CompoundLineType.InnerText ?? "sng";
        }
        else
        {
            // Style-matrix lnRef fallback (parity with RenderShape / RenderPicture):
            // a connector with NO explicit <a:ln> takes its color and weight from
            // <p:style>/<a:lnRef idx=N> against the theme line-style matrix. Every
            // default PowerPoint connector is styled this way (typically an accent
            // color), so without this the line wrongly rendered theme-tx1 black.
            var stroke = ResolveStyleLineRefStroke(style, part, themeColors);
            if (stroke != null)
            {
                lineColor = stroke.Value.color;
                lineWidth = stroke.Value.widthPt;
            }
        }

        // Ensure minimum dimensions so the line is visible
        // For horizontal lines (cy=0), the container needs height for stroke width
        // For vertical lines (cx=0), the container needs width for stroke width
        var minDimEmu = (long)(lineWidth * EmuConverter.EmuPerPoint + 12700); // lineWidth + 1pt padding
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
        if (prstDash?.Val?.HasValue == true && prstDash.Val.InnerText is { } dashVal)
        {
            // CONSISTENCY(dash-presets): reuse the canonical CSS-tables dash converter
            // so the shapes/connector path emits the SAME dasharray as table borders.
            // Previously this path had its own divergent switch where lgDash == dash
            // ({w*4},{w*3}); the canonical helper draws lgDash with the longer {w*8} on-length.
            var dashArray = DashTypeToSvgDasharray(dashVal, lineWidth);
            if (!string.IsNullOrEmpty(dashArray))
                dashAttr = $" stroke-dasharray=\"{dashArray}\"";
        }
        else
        {
            // CONSISTENCY(dash-presets): <a:custDash> (mutually exclusive with prstDash)
            // is a list of <a:ds d sp/> stops (ST_PositivePercentage of line width).
            // Mirror ParseOutline's encoding into "custom:<onMult>,<offMult>,..." and
            // run it through the same DashTypeToSvgDasharray converter. Without this,
            // a connector with a custom dash rendered SOLID.
            var custDash = outline?.GetFirstChild<Drawing.CustomDash>();
            if (custDash != null)
            {
                var ci = System.Globalization.CultureInfo.InvariantCulture;
                var segs = new List<string>();
                foreach (var ds in custDash.Elements<Drawing.DashStop>())
                {
                    double d = (ds.DashLength?.Value ?? 0) / 100000.0;
                    double sp = (ds.SpaceLength?.Value ?? 0) / 100000.0;
                    if (d <= 0 && sp <= 0) continue;
                    segs.Add(d.ToString("0.##", ci));
                    segs.Add(sp.ToString("0.##", ci));
                }
                if (segs.Count > 0)
                {
                    var dashArray = DashTypeToSvgDasharray("custom:" + string.Join(",", segs), lineWidth);
                    if (!string.IsNullOrEmpty(dashArray))
                        dashAttr = $" stroke-dasharray=\"{dashArray}\"";
                }
            }
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
            // R37-A: marker dimensions are in strokeWidth units (SVG default
            // markerUnits="strokeWidth"), so the rendered arrowhead size is
            // markerWidth × strokeWidth. The base must therefore be a SMALL CONSTANT
            // (NOT multiplied by lineWidth) — otherwise the effective size grows as
            // O(lineWidth²). A "med" triangle in PowerPoint is ~3-4.5× the stroke width,
            // so base ≈ 4 gives the right proportion and scales LINEARLY with the line.
            var baseArrowSize = 4.0;
            // R4-5: scale the marker by the head/tail @w / @len size enum
            // (sm/med/lg) so a large arrowhead renders visibly bigger than the
            // default; previously arrowSize was uniform and ignored @w/@len.
            // ST_LineEndWidth/Length default is "med" → ×1.0.
            static double TokenScale(string? s) => s switch { "sm" => 0.6, "lg" => 1.6, _ => 1.0 };
            // @w (ST_LineEndWidth) = arrowhead width PERPENDICULAR to the line;
            // @len (ST_LineEndLength) = arrowhead length ALONG the line. They are
            // independent — a "w=lg len=sm" head is wide but short. Return both
            // scales so the marker can be a rectangle, not a Math.Max square.
            static (double lenScale, double wScale) ArrowSizeScales(OpenXmlElement? end)
            {
                if (end == null) return (1.0, 1.0);
                var attrs = end.GetAttributes();
                var w = attrs.FirstOrDefault(a => a.LocalName == "w").Value;
                var l = attrs.FirstOrDefault(a => a.LocalName == "len").Value;
                return (TokenScale(l), TokenScale(w));
            }
            var (headLenScale, headWScale) = ArrowSizeScales(headEnd);
            var (tailLenScale, tailWScale) = ArrowSizeScales(tailEnd);
            var headLen = baseArrowSize * headLenScale;
            var headW = baseArrowSize * headWScale;
            var tailLen = baseArrowSize * tailLenScale;
            var tailW = baseArrowSize * tailWScale;
            var defs = new StringBuilder();
            defs.Append("<defs>");
            // BUG1(marker-id-collision): marker ids are document-global in HTML even when each
            // marker sits in its own <svg>. A fixed id ("ah"/"at") makes every line's
            // marker-end="url(#at)" resolve to the FIRST line's marker, so all arrowheads
            // inherit the first line's color. Use a per-render counter to make each marker id
            // unique so each line references its own correctly-colored marker.
            var markerSeq = _markerCounter++;
            // BUG2(marker-shape): emit the correct geometry per head/tail type (triangle,
            // diamond/rhombus, oval/circle, stealth/notched) instead of always a triangle.
            // For marker-start we use orient="auto-start-reverse" so SVG flips the right-pointing
            // geometry to point outward (leftward) at the line's start.
            if (hasHead)
            {
                var hid = $"ah{markerSeq}";
                defs.Append($"<marker id=\"{hid}\" markerWidth=\"{headLen:0.#}\" markerHeight=\"{headW:0.#}\" refX=\"{headLen:0.#}\" refY=\"{headW / 2:0.#}\" orient=\"auto-start-reverse\">{ArrowMarkerGeometry(headEnd!.Type!.InnerText ?? "triangle", headLen, headW, safeColor)}</marker>");
                markerStartAttr = $" marker-start=\"url(#{hid})\"";
            }
            if (hasTail)
            {
                var tid = $"at{markerSeq}";
                defs.Append($"<marker id=\"{tid}\" markerWidth=\"{tailLen:0.#}\" markerHeight=\"{tailW:0.#}\" refX=\"{tailLen:0.#}\" refY=\"{tailW / 2:0.#}\" orient=\"auto\">{ArrowMarkerGeometry(tailEnd!.Type!.InnerText ?? "triangle", tailLen, tailW, safeColor)}</marker>");
                markerEndAttr = $" marker-end=\"url(#{tid})\"";
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

        // Line cap (rnd→round pill dash ends, sq→square, flat→butt) applies to the stroke.
        var linecapAttr = $" stroke-linecap=\"{CapToSvgLinecap(lineCap)}\"";
        // CONSISTENCY(shape-stroke-unit): stroke-width in pt matches CSS border path (see R3 fix).
        // R15-2: stroke with the gradient def when present, else the (first-stop/solid) color.
        var strokePaint = gradStrokeId != null ? $"url(#{gradStrokeId})" : safeColor;
        var strokeAttrs = $"stroke=\"{strokePaint}\" stroke-width=\"{lineWidth:0.##}pt\" fill=\"none\"{linecapAttr}{dashAttr}{markerStartAttr}{markerEndAttr}";
        // Merge the gradient def into the SVG <defs> alongside any marker defs.
        if (gradStrokeDef != null)
        {
            markerDefs = string.IsNullOrEmpty(markerDefs)
                ? $"<defs>{gradStrokeDef}</defs>"
                : markerDefs.Replace("</defs>", $"{gradStrokeDef}</defs>");
        }
        // Compound line (cmpd=dbl/thickThin/thinThick/tri): SVG strokes are a single path,
        // so we approximate "two parallel lines" by overlaying a thinner transparent-gap
        // stroke down the center of the full-width stroke. This splits the visible stroke
        // into two parallel runs without computing offset geometry.
        var isCompound = lineCmpd != "sng";
        var compoundGapAttrs = isCompound
            ? $"stroke=\"transparent\" stroke-width=\"{lineWidth / 3.0:0.##}pt\" fill=\"none\"{linecapAttr}"
            : "";

        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        // CONSISTENCY(shape-rotation): connectors use the same Transform2D.Rotation
        // slot as shapes/pictures/groups; apply the same CSS transform so the rendered
        // line matches PowerPoint. Default transform-origin (50% 50%) matches OOXML
        // rotation pivot (bounding-box center).
        var cxnRotTransform = "";
        if (xfrm?.Rotation != null && xfrm.Rotation.Value != 0)
            cxnRotTransform = $";transform:rotate({xfrm.Rotation.Value / 60000.0:0.##}deg)";
        // Effects (outer shadow / glow / blur) — connectors carry a:effectLst on their
        // spPr just like shapes/pictures and PowerPoint renders the shadow beneath the
        // line. RenderConnector previously dropped it entirely; mirror the shape path
        // (the SVG sits in an overflow:visible div, so filter:drop-shadow applies to the
        // stroke). Inner shadow has no CSS-filter equivalent on a line, so skip it.
        var cxnEffectList = spPr?.GetFirstChild<Drawing.EffectList>();
        // Style-matrix fallback: when the connector's spPr carries no explicit
        // <a:effectLst>, resolve its <p:style>/<a:effectRef> against the theme
        // EffectStyleList — exactly as the shape (RenderShape ~line 347) and
        // picture (~line 1508) paths already do. A connector styled via the
        // "Shadow"/effect connector-style gallery stores its shadow in effectRef,
        // not in spPr; without this fallback PowerPoint draws the shadow but the
        // preview dropped it. Explicit spPr effects still win.
        if (cxnEffectList == null && part != null)
            cxnEffectList = ResolveStyleEffectRefList(style, part);
        var cxnFilterParts = new List<string>();
        var cxnShadowCss = EffectListToShadowCss(cxnEffectList, themeColors);
        if (!string.IsNullOrEmpty(cxnShadowCss) && !cxnShadowCss.StartsWith("box-shadow:"))
            cxnFilterParts.Add(cxnShadowCss.Replace("filter:", ""));
        var cxnGlowCss = EffectListToGlowCss(cxnEffectList, themeColors);
        if (!string.IsNullOrEmpty(cxnGlowCss))
            cxnFilterParts.Add(cxnGlowCss.Replace("filter:", ""));
        var cxnBlurCss = EffectListToBlurCss(cxnEffectList);
        if (!string.IsNullOrEmpty(cxnBlurCss))
            cxnFilterParts.Add(cxnBlurCss);
        var cxnFilter = cxnFilterParts.Count > 0 ? $";filter:{string.Join(" ", cxnFilterParts)}" : "";
        sb.AppendLine($"    <div class=\"connector\"{dataPathAttr} style=\"left:{Units.EmuToPt(renderX)}pt;top:{Units.EmuToPt(renderY)}pt;width:{widthPt}pt;height:{heightPt}pt{cxnRotTransform}{cxnFilter}\">");

        if (preset.StartsWith("bentConnector", StringComparison.Ordinal))
        {
            // Bent connectors: right-angle polyline. Use viewBox=0..100 so stretched
            // preserveAspectRatio=none fills the container.
            // bentConnector2: single 90-degree bend (2 segments, 3 points).
            // bentConnector3 (default): 3 segments with mid bend — (0,0) -> (50,0) -> (50,100) -> (100,100).
            // bentConnector4/5: approximate with 25/75 splits when no adjustments set.
            // The elbow position(s) are driven by avLst adjusts (a user dragging the
            // elbow handle writes <a:gd name="adjN" fmla="val P"/>, P in 1/100000).
            // Previously the elbows were hardcoded (50% for bentConnector3; 25/50/75 for
            // 4/5), so any non-default elbow rendered in the wrong place.
            var ci = System.Globalization.CultureInfo.InvariantCulture;
            double ConnAdj(string name, double def)
            {
                var f = prstGeom?.GetFirstChild<Drawing.AdjustValueList>()?.Elements<Drawing.ShapeGuide>()
                    .FirstOrDefault(g => g.Name?.Value == name)?.Formula?.Value;
                if (f != null && f.StartsWith("val ", StringComparison.Ordinal)
                    && double.TryParse(f.AsSpan(4), System.Globalization.NumberStyles.Float, ci, out var v))
                    return v / 100000.0 * 100.0;
                return def;
            }
            string N(double v) => v.ToString("0.##", ci);
            string points;
            if (preset == "bentConnector2")
                points = "0,0 100,0 100,100";
            else if (preset is "bentConnector4" or "bentConnector5")
            {
                var a1 = ConnAdj("adj1", 25); var a2 = ConnAdj("adj2", 50); var a3 = ConnAdj("adj3", 75);
                points = $"0,0 {N(a1)},0 {N(a1)},{N(a2)} {N(a3)},{N(a2)} {N(a3)},100 100,100";
            }
            else // bentConnector3
            {
                var ex = ConnAdj("adj1", 50);
                points = $"0,0 {N(ex)},0 {N(ex)},100 100,100";
            }
            // R27: mirror the polyline in the 0..100 viewBox when flipH/flipV is
            // set so a flipped elbow lands on the shape edges (the straight branch
            // already flips via svgX1/Y1/X2/Y2). flipH → x'=100-x, flipV → y'=100-y.
            points = MirrorConnectorPoints(points, flipH, flipV);
            sb.AppendLine("      <svg width=\"100%\" height=\"100%\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\" style=\"overflow:visible;display:block\">");
            if (!string.IsNullOrEmpty(markerDefs))
                sb.AppendLine($"        {markerDefs}");
            sb.AppendLine($"        <polyline points=\"{points}\" {strokeAttrs}/>");
            if (isCompound)
                sb.AppendLine($"        <polyline points=\"{points}\" {compoundGapAttrs}/>");
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
            // R27: mirror the bezier control points in the 0..100 viewBox when
            // flipH/flipV is set (parity with the bent + straight branches).
            d = MirrorConnectorPath(d, flipH, flipV);
            sb.AppendLine("      <svg width=\"100%\" height=\"100%\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\" style=\"overflow:visible;display:block\">");
            if (!string.IsNullOrEmpty(markerDefs))
                sb.AppendLine($"        {markerDefs}");
            sb.AppendLine($"        <path d=\"{d}\" {strokeAttrs}/>");
            if (isCompound)
                sb.AppendLine($"        <path d=\"{d}\" {compoundGapAttrs}/>");
            sb.AppendLine("      </svg>");
        }
        else
        {
            sb.AppendLine("      <svg width=\"100%\" height=\"100%\" preserveAspectRatio=\"none\" style=\"overflow:visible;display:block\">");
            if (!string.IsNullOrEmpty(markerDefs))
                sb.AppendLine($"        {markerDefs}");
            sb.AppendLine($"        <line x1=\"{svgX1}\" y1=\"{svgY1}\" x2=\"{svgX2}\" y2=\"{svgY2}\" {strokeAttrs}/>");
            if (isCompound)
                sb.AppendLine($"        <line x1=\"{svgX1}\" y1=\"{svgY1}\" x2=\"{svgX2}\" y2=\"{svgY2}\" {compoundGapAttrs}/>");
            sb.AppendLine("      </svg>");
        }
        // R15-1: overlay the connector's text label, centered over the line's bounding
        // box. The connector div is position:absolute with no intrinsic text host, so we
        // emit an absolutely-positioned, centered inner div and render the txBody into it.
        if (cxnTextBody != null && !string.IsNullOrWhiteSpace(cxnTextBody.InnerText))
        {
            sb.Append("      <div style=\"position:absolute;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none\">");
            RenderTextBody(sb, cxnTextBody, themeColors);
            sb.AppendLine("</div>");
        }
        sb.AppendLine("    </div>");
    }

    // R27: mirror a "x,y x,y ..." polyline-points string in the 0..100 viewBox.
    // flipH → x'=100-x; flipV → y'=100-y. Both axes are applied independently so
    // a bent/curved connector flips to the same orientation real PowerPoint routes.
    private static string MirrorConnectorPoints(string points, bool flipH, bool flipV)
    {
        if (!flipH && !flipV) return points;
        var pairs = points.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < pairs.Length; i++)
        {
            var xy = pairs[i].Split(',');
            if (xy.Length != 2) continue;
            if (flipH && double.TryParse(xy[0], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var px))
                xy[0] = (100 - px).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
            if (flipV && double.TryParse(xy[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var py))
                xy[1] = (100 - py).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
            pairs[i] = $"{xy[0]},{xy[1]}";
        }
        return string.Join(' ', pairs);
    }

    // R27: mirror an SVG path "d" of the form "M x,y C/Q x,y x,y x,y" — the
    // command letters (M/C/Q) pass through; every "x,y" coordinate token is
    // mirrored via MirrorConnectorPoints' per-pair logic.
    private static string MirrorConnectorPath(string d, bool flipH, bool flipV)
    {
        if (!flipH && !flipV) return d;
        var tokens = d.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < tokens.Length; i++)
        {
            if (!tokens[i].Contains(',')) continue; // command letter (M/C/Q)
            tokens[i] = MirrorConnectorPoints(tokens[i], flipH, flipV);
        }
        return string.Join(' ', tokens);
    }

    // ==================== Group Rendering ====================

    // CONSISTENCY(shape-group-parity): build the group container transform from
    // rotation + flip, mirroring RenderShape (rotate before scale). A flipped
    // group mirrors all its children in real PowerPoint, so the container div
    // must carry the flip too. Returns "" or ";transform:...".
    private static string BuildGroupTransform(Drawing.TransformGroup? grpXfrm)
    {
        var parts = new List<string>();
        if (grpXfrm?.Rotation != null && grpXfrm.Rotation.Value != 0)
            parts.Add($"rotate({grpXfrm.Rotation.Value / 60000.0:0.##}deg)");
        if (grpXfrm?.HorizontalFlip?.Value == true && grpXfrm.VerticalFlip?.Value == true)
            parts.Add("scale(-1,-1)");
        else if (grpXfrm?.HorizontalFlip?.Value == true)
            parts.Add("scaleX(-1)");
        else if (grpXfrm?.VerticalFlip?.Value == true)
            parts.Add("scaleY(-1)");
        return parts.Count > 0 ? $";transform:{string.Join(" ", parts)}" : "";
    }

    private void RenderGroup(StringBuilder sb, GroupShape grp, OpenXmlPart slidePart, Dictionary<string, string> themeColors, string? dataPath = null)
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
        var grpTransform = BuildGroupTransform(grpXfrm);
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
                        RenderNestedGroup(sb, nestedGrp, slidePart, themeColors, nx, ny, ncx, ncy, depth: 1);
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
                        RenderConnector(sb, cxn, themeColors, dataPath: null, overridePos: pos, part: slidePart);
                    break;
                }
                case GraphicFrame gf:
                {
                    // R14-2: chart/table inside a group. Re-project the graphicFrame's
                    // <p:xfrm> into the group's child coordinate system, then route to
                    // RenderTable / RenderChart like the slide-level dispatch does.
                    var pos = CalcGraphicFramePos(gf, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                    {
                        if (gf.Descendants<Drawing.Table>().Any())
                            RenderTable(sb, gf, themeColors, dataPath: null, overridePos: pos, part: slidePart);
                        else if (slidePart is SlidePart sp)
                            RenderChart(sb, gf, sp, themeColors, dataPath: null, overridePos: pos);
                    }
                    break;
                }
                default:
                {
                    // mc:AlternateContent inside a group (e.g. an a14 math text box
                    // grouped with other shapes) — see RenderAlternateContent (#228).
                    if (child.LocalName == "AlternateContent")
                        RenderGroupAltContent(sb, child, slidePart, themeColors, offX, offY, scaleX, scaleY);
                    break;
                }
            }
        }

        sb.AppendLine("    </div>");
    }

    // Shapes inside an mc:AlternateContent child of a group. Same drop-fix as
    // RenderAlternateContent's general branch (#228), but mirroring the group
    // child branches: each inner shape's slide-space xfrm must be re-projected
    // into the group's child coordinate system (CONSISTENCY(group-child-pos)).
    private static void RenderGroupAltContent(StringBuilder sb, OpenXmlElement acElement,
        OpenXmlPart slidePart, Dictionary<string, string> themeColors,
        long offX, long offY, double scaleX, double scaleY)
    {
        var altChild = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice")
                    ?? acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
        if (altChild == null) return;
        foreach (var inner in altChild.ChildElements)
        {
            // Inside a group the SDK leaves the mc:AlternateContent subtree
            // untyped (OpenXmlUnknownElement), so `is Shape` never matches —
            // re-hydrate the strongly-typed element from its XML.
            switch (CoerceAltContentChild(inner))
            {
                case Shape sp:
                {
                    var pos = CalcGroupChildPos(sp.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderShape(sb, sp, slidePart, themeColors, pos);
                    break;
                }
                case Picture pic:
                {
                    var pos = CalcGroupChildPos(pic.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderPicture(sb, pic, slidePart, themeColors, pos);
                    break;
                }
                case ConnectionShape cxn:
                {
                    var pos = CalcGroupChildPos(cxn.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderConnector(sb, cxn, themeColors, dataPath: null, overridePos: pos, part: slidePart);
                    break;
                }
            }
        }
    }

    // Re-hydrate an mc:Choice/mc:Fallback child into its strongly-typed SDK
    // element. Where the schema doesn't admit AlternateContent (e.g. inside
    // p:grpSp) the SDK parses the subtree as OpenXmlUnknownElement, so typed
    // pattern matches and typed property accessors (ShapeProperties, TextBody)
    // silently yield null. Already-typed elements pass through unchanged.
    private static OpenXmlElement CoerceAltContentChild(OpenXmlElement inner)
    {
        if (inner is not OpenXmlUnknownElement) return inner;
        try
        {
            return inner.LocalName switch
            {
                "sp" => new Shape(inner.OuterXml),
                "pic" => new Picture(inner.OuterXml),
                "cxnSp" => new ConnectionShape(inner.OuterXml),
                "graphicFrame" => new GraphicFrame(inner.OuterXml),
                _ => inner,
            };
        }
        catch { return inner; }
    }

    /// <summary>
    /// Pure calculation: re-project a grouped GraphicFrame's <p:xfrm> into the
    /// group's child coordinate system. Returns null if the frame has no xfrm.
    /// </summary>
    private static (long x, long y, long cx, long cy)? CalcGraphicFramePos(
        GraphicFrame gf, long offX, long offY, double scaleX, double scaleY)
    {
        var off = gf.Transform?.Offset;
        var ext = gf.Transform?.Extents;
        if (off == null || ext == null) return null;
        return (
            (long)(((off.X?.Value ?? 0) - offX) * scaleX),
            (long)(((off.Y?.Value ?? 0) - offY) * scaleY),
            (long)((ext.Cx?.Value ?? 0) * scaleX),
            (long)((ext.Cy?.Value ?? 0) * scaleY)
        );
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
    private void RenderNestedGroup(StringBuilder sb, GroupShape grp, OpenXmlPart slidePart,
        Dictionary<string, string> themeColors, long x, long y, long cx, long cy, int depth = 0)
    {
        // CONSISTENCY(dos-hardening): nested-group recursion is unbounded; a
        // crafted deeply-nested grpSp would overflow the stack during
        // `view html`. See DocumentLimits.
        DocumentLimits.EnsureDepth(depth);

        var grpXfrm = grp.GroupShapeProperties?.TransformGroup;

        // Child coordinate system of this nested group
        var childOff = grpXfrm?.ChildOffset;
        var childExt = grpXfrm?.ChildExtents;
        var scaleX = (childExt?.Cx?.Value ?? cx) != 0 ? (double)cx / (childExt?.Cx?.Value ?? cx) : 1.0;
        var scaleY = (childExt?.Cy?.Value ?? cy) != 0 ? (double)cy / (childExt?.Cy?.Value ?? cy) : 1.0;
        var offX = childOff?.X?.Value ?? 0;
        var offY = childOff?.Y?.Value ?? 0;

        // CONSISTENCY(group-rotation): same idiom as RenderGroup
        var grpTransform = BuildGroupTransform(grpXfrm);
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
                        RenderNestedGroup(sb, nestedGrp, slidePart, themeColors, nx, ny, ncx, ncy, depth: depth + 1);
                    }
                    break;
                }
                case ConnectionShape cxn:
                {
                    // CONSISTENCY(group-child-pos): see RenderGroup ConnectionShape branch.
                    var pos = CalcGroupChildPos(cxn.ShapeProperties?.Transform2D, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                        RenderConnector(sb, cxn, themeColors, dataPath: null, overridePos: pos, part: slidePart);
                    break;
                }
                case GraphicFrame gf:
                {
                    // R14-2: see RenderGroup GraphicFrame branch.
                    var pos = CalcGraphicFramePos(gf, offX, offY, scaleX, scaleY);
                    if (pos.HasValue)
                    {
                        if (gf.Descendants<Drawing.Table>().Any())
                            RenderTable(sb, gf, themeColors, dataPath: null, overridePos: pos, part: slidePart);
                        else if (slidePart is SlidePart sp)
                            RenderChart(sb, gf, sp, themeColors, dataPath: null, overridePos: pos);
                    }
                    break;
                }
                default:
                {
                    // mc:AlternateContent — see RenderGroup default branch (#228).
                    if (child.LocalName == "AlternateContent")
                        RenderGroupAltContent(sb, child, slidePart, themeColors, offX, offY, scaleX, scaleY);
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
        if (!isModel3D && !isZoom)
        {
            // General mc:AlternateContent — e.g. a text box whose body carries
            // Office 2010 math (mc:Choice Requires="a14" around a p:sp). Returning
            // early here silently dropped the whole shape, equations and text alike
            // (#228). Route the shapes inside mc:Choice (preferred: it carries the
            // full-fidelity content our renderer understands, e.g. a14:m equations)
            // or mc:Fallback through the normal shape/picture pipeline instead.
            var altChild = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice")
                        ?? acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
            if (altChild == null) return;
            foreach (var acInner in altChild.ChildElements)
            {
                switch (CoerceAltContentChild(acInner))
                {
                    case Shape acSp:
                        RenderShape(sb, acSp, slidePart, themeColors, dataPath: dataPath);
                        break;
                    case Picture acPic:
                        RenderPicture(sb, acPic, slidePart, themeColors, dataPath: dataPath);
                        break;
                    case ConnectionShape acCxn:
                        RenderConnector(sb, acCxn, themeColors, dataPath: dataPath, part: slidePart);
                        break;
                    case GraphicFrame acGf when acGf.Descendants<Drawing.Table>().Any():
                        RenderTable(sb, acGf, themeColors, dataPath: dataPath, part: slidePart);
                        break;
                }
            }
            return;
        }

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

    // Per-render counter ensuring each line's arrowhead marker gets a globally-unique
    // HTML id (markers share the document id-space across separate <svg> elements).
    private static int _markerCounter;

    // Build the SVG geometry for an arrowhead marker of the given OOXML end type.
    // Coordinate space: 0..len along the line (x), 0..width perpendicular (y); line
    // enters from the left, tip at the right (refX=len, refY=width/2). The two
    // dimensions are independent so @w (width) and @len (length) render with the
    // correct aspect ratio — a wide-but-short arrowhead is NOT an equilateral one.
    // marker-start flips this via orient.
    private static string ArrowMarkerGeometry(string endType, double len, double width, string color)
    {
        var s = len;            // extent along the line (x)
        var w = width;          // extent perpendicular (y)
        var h = width / 2;      // perpendicular centre
        var m = len / 2;        // mid along the line
        return endType switch
        {
            // Rhombus centered on the tip: left, top, right, bottom vertices.
            "diamond" => $"<polygon points=\"0 {h:0.#},{m:0.#} 0,{s:0.#} {h:0.#},{m:0.#} {w:0.#}\" fill=\"{color}\"/>",
            // Filled ellipse sitting at the line end (circle when len==width).
            "oval" => $"<ellipse cx=\"{m:0.#}\" cy=\"{h:0.#}\" rx=\"{m:0.#}\" ry=\"{h:0.#}\" fill=\"{color}\"/>",
            // Concave/notched arrow: triangle with a notch cut into the back edge.
            "stealth" => $"<polygon points=\"0 0,{s:0.#} {h:0.#},0 {w:0.#},{m:0.#} {h:0.#}\" fill=\"{color}\"/>",
            // R37-B: OOXML type="arrow" is an OPEN arrowhead — two strokes meeting at the
            // tip (like ">"), NOT a filled area. Emit an open chevron via <polyline> with
            // fill="none" and an explicit stroke (the marker's internal coordinate space does
            // not inherit the outer line's stroke-width, so set it explicitly ≈ width/6).
            // Back corners at (0,0) and (0,w), tip at the right (s,h).
            "arrow" => $"<polyline points=\"0 0,{s:0.#} {h:0.#},0 {w:0.#}\" fill=\"none\" stroke=\"{color}\" stroke-width=\"{w / 6:0.##}\"/>",
            // triangle / default: right-pointing solid triangle (▶).
            _ => $"<polygon points=\"0 0,{s:0.#} {h:0.#},0 {w:0.#}\" fill=\"{color}\"/>",
        };
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
        _markerCounter = 0;
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
      // Mirror-first via the importmap (CONSISTENCY(katex-mirror), see Core/ThreeAssets)
      THREE = await import('three');
      ({{ GLTFLoader }} = await import('three/addons/loaders/GLTFLoader.js'));
    }} catch(e) {{
      try {{
        // Mirror unreachable — CDN /+esm bypasses the importmap (absolute rewritten imports)
        THREE = await import('{Core.ThreeAssets.CdnEsmThreeUrl}');
        ({{ GLTFLoader }} = await import('{Core.ThreeAssets.CdnEsmGltfLoaderUrl}'));
      }} catch(e2) {{
        // Three.js unavailable (offline) — show fallback image
        const c = document.getElementById('{canvasId}');
        if (c) {{ c.style.display='none'; const fb=c.parentElement?.querySelector('.m3d-fallback'); if(fb) fb.style.display='block'; }}
        throw e2; // stop execution of this module
      }}
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
    // SmartArt (diagram) GraphicFrame. Prefer rendering the cached dsp drawing
    // part (the fully laid-out <dsp:sp> shapes). Resolve chain:
    //   gf → dgm:relIds/@r:dm → DiagramDataPart
    //   data part → <dsp:dataModelExt @relId> → DiagramPersistLayoutPart (slide rel)
    //   drawing part → <dsp:spTree> with <dsp:sp> shapes
    // Each <dsp:sp> is structurally identical to <p:sp>; we namespace-swap the
    // drawing XML to the presentation namespace, reparse into a ShapeTree, and
    // reuse RenderShape. dsp shape xfrm offsets are in the diagram's own space,
    // which (per the dsp drawing contract) shares the graphicFrame's EMU origin,
    // so we offset each shape by the frame's top-left. When the drawing part is
    // absent/unresolvable, fall back to a labeled placeholder (mirrors
    // RenderOlePlaceholder) so the SmartArt is never silently invisible.
    private void RenderSmartArt(StringBuilder sb, GraphicFrame gf, OpenXmlPart slidePart,
        Dictionary<string, string> themeColors, string? dataPath = null)
    {
        var xfrm = gf.Transform;
        var fx = xfrm?.Offset?.X?.Value ?? 0;
        var fy = xfrm?.Offset?.Y?.Value ?? 0;
        var fcx = xfrm?.Extents?.Cx?.Value ?? 0;
        var fcy = xfrm?.Extents?.Cy?.Value ?? 0;

        var drawingPart = TryResolveDiagramDrawingPart(gf, slidePart);
        if (drawingPart != null)
        {
            try
            {
                string raw;
                using (var s = drawingPart.GetStream(FileMode.Open, FileAccess.Read))
                using (var r = new StreamReader(s))
                    raw = r.ReadToEnd();

                // Swap the dsp namespace to the presentation namespace so <dsp:sp>
                // (and its nvSpPr/spPr/txBody) reparse as p:sp / Shape, and the
                // drawing/spTree roots as p:cSld/p:spTree. dsp shares the same
                // child element local-names + Drawing (a:) children as p:.
                const string dspNs = "http://schemas.microsoft.com/office/drawing/2008/diagram";
                const string pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
                // Swap both the element prefix (<dsp:spTree> → <p:spTree>, and the
                // xmlns:dsp= declaration → xmlns:p=) AND the namespace URI. Replacing
                // only the URI leaves the elements as <dsp:...>, so the spTree
                // extraction below would never match (dead drawing-render path).
                var swapped = raw.Replace("dsp:", "p:").Replace(dspNs, pNs);

                // Wrap the spTree in a throwaway slide so reparse yields a Slide we
                // can walk for Shape children. The dsp drawing root is <dsp:drawing>
                // with a <dsp:spTree>; after the swap it's <p:drawing>/<p:spTree>.
                // Extract just the spTree subtree and host it in a minimal p:sld.
                int treeStart = swapped.IndexOf("<p:spTree", StringComparison.Ordinal);
                int treeEnd = swapped.LastIndexOf("</p:spTree>", StringComparison.Ordinal);
                if (treeStart >= 0 && treeEnd > treeStart)
                {
                    var treeXml = swapped.Substring(treeStart, treeEnd - treeStart + "</p:spTree>".Length);
                    var sldXml =
                        $"<p:sld xmlns:p=\"{pNs}\" " +
                        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<p:cSld>" + treeXml + "</p:cSld></p:sld>";
                    var slide = new Slide(sldXml);
                    var spTree = slide.CommonSlideData?.ShapeTree;
                    if (spTree != null)
                    {
                        bool emittedAny = false;
                        bool dpUsed = false;
                        foreach (var shape in spTree.Elements<Shape>())
                        {
                            var sx = shape.ShapeProperties?.Transform2D?.Offset?.X?.Value ?? 0;
                            var sy = shape.ShapeProperties?.Transform2D?.Offset?.Y?.Value ?? 0;
                            var scx = shape.ShapeProperties?.Transform2D?.Extents?.Cx?.Value ?? 0;
                            var scy = shape.ShapeProperties?.Transform2D?.Extents?.Cy?.Value ?? 0;
                            // The dsp spTree origin maps to the frame origin; offset
                            // each shape's local position by the frame's top-left.
                            var pos = (x: fx + sx, y: fy + sy, cx: scx, cy: scy);
                            RenderShape(sb, shape, slidePart, themeColors, overridePos: pos,
                                dataPath: dpUsed ? null : dataPath);
                            dpUsed = true;
                            emittedAny = true;
                        }
                        if (emittedAny) return;
                    }
                }
            }
            catch { /* fall through to placeholder */ }
        }

        // Fallback: labeled placeholder so the element is never silently invisible.
        var leftPt = Units.EmuToPt(fx);
        var topPt = Units.EmuToPt(fy);
        var widthPt = Units.EmuToPt(fcx);
        var heightPt = Units.EmuToPt(fcy);
        var dpAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        sb.AppendLine($"    <div class=\"smartart-placeholder\"{dpAttr} style=\"position:absolute;" +
            $"left:{leftPt:0.##}pt;top:{topPt:0.##}pt;" +
            $"width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;" +
            $"border:2px dashed rgba(108,117,125,0.6);border-radius:4px;" +
            $"display:flex;align-items:center;justify-content:center;" +
            $"font:11pt sans-serif;color:#495057;background:rgba(248,249,250,0.7);" +
            $"overflow:hidden;text-align:center;padding:4px;box-sizing:border-box;\">" +
            $"SmartArt</div>");
    }

    // Resolve the dsp cached-drawing part for a SmartArt graphicFrame, or null.
    private static DiagramPersistLayoutPart? TryResolveDiagramDrawingPart(GraphicFrame gf, OpenXmlPart slidePart)
    {
        try
        {
            const string dgmNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            const string dspNs = "http://schemas.microsoft.com/office/drawing/2008/diagram";
            var relIds = gf.Descendants().FirstOrDefault(e =>
                e.LocalName == "relIds" && e.NamespaceUri == dgmNs);
            if (relIds == null) return null;
            string? dataRid = null;
            foreach (var a in relIds.GetAttributes())
                if (a.LocalName == "dm") { dataRid = a.Value; break; }
            if (string.IsNullOrEmpty(dataRid)) return null;

            if (slidePart.GetPartById(dataRid!) is not DiagramDataPart dataPart) return null;
            var ext = dataPart.DataModelRoot?.Descendants().FirstOrDefault(e =>
                e.LocalName == "dataModelExt" && e.NamespaceUri == dspNs);
            string? drawingRelId = null;
            if (ext != null)
                foreach (var a in ext.GetAttributes())
                    if (a.LocalName == "relId") { drawingRelId = a.Value; break; }
            if (string.IsNullOrEmpty(drawingRelId)) return null;

            return slidePart.GetPartById(drawingRelId!) as DiagramPersistLayoutPart;
        }
        catch { return null; }
    }

    private static void RenderOlePlaceholder(StringBuilder sb, GraphicFrame gf, OpenXmlPart slidePart, string? dataPath = null)
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

        // RR5: PowerPoint stores the OLE's rendered preview as an embedded p:pic
        // blip (inside p:oleObj, or in an mc:AlternateContent/mc:Fallback wrapper
        // around the graphicFrame). Surface that thumbnail as an inline <img>
        // data-uri — mirroring RenderPicture's blip → base64 emission — instead
        // of only showing the dashed-box "OLE: {progId}" text placeholder.
        var oleImgSrc = TryExtractOlePreviewDataUri(gf, slidePart);
        if (oleImgSrc != null)
        {
            sb.AppendLine($"    <img class=\"ole-preview\"{dpAttr} src=\"{oleImgSrc}\" alt=\"{label}\" style=\"position:absolute;" +
                $"left:{leftPt:0.##}pt;top:{topPt:0.##}pt;" +
                $"width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;" +
                $"object-fit:contain;box-sizing:border-box;\"/>");
            return;
        }

        sb.AppendLine($"    <div class=\"ole-placeholder\"{dpAttr} style=\"position:absolute;" +
            $"left:{leftPt:0.##}pt;top:{topPt:0.##}pt;" +
            $"width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;" +
            $"border:2px dashed rgba(108,117,125,0.6);border-radius:4px;" +
            $"display:flex;align-items:center;justify-content:center;" +
            $"font:11pt sans-serif;color:#495057;background:rgba(248,249,250,0.7);" +
            $"overflow:hidden;text-align:center;padding:4px;box-sizing:border-box;\">" +
            $"{label}</div>");
    }

    // RR5: locate the OLE preview blip (any a:blip with an r:embed under the
    // graphicFrame — covers p:oleObj/p:pic and the mc:Fallback variant) and
    // return its image bytes as a data-uri, or null when none is present.
    private static string? TryExtractOlePreviewDataUri(GraphicFrame gf, OpenXmlPart slidePart)
    {
        const string rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        foreach (var blip in gf.Descendants().Where(d => d.LocalName == "blip"))
        {
            var embedId = blip.GetAttribute("embed", rNs).Value;
            if (string.IsNullOrEmpty(embedId)) continue;
            try
            {
                var imgPart = slidePart.GetPartById(embedId!);
                using var stream = imgPart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                var base64 = Convert.ToBase64String(ms.ToArray());
                var contentType = SanitizeContentType(imgPart.ContentType ?? "image/png");
                return $"data:{contentType};base64,{base64}";
            }
            catch { /* unresolved rel — try the next blip */ }
        }
        return null;
    }
}
