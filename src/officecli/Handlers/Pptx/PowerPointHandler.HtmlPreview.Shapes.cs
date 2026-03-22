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
                sb.Append("<div style=\"width:100%;height:100%;background:#e0e0e0;display:flex;align-items:center;justify-content:center;color:#999;font-size:12px\">Image</div>");
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
}
