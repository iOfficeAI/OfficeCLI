// Copyright 2025 OfficeCli (officecli.ai)
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
    // ==================== Drawing with Overlaid Images ====================

    private void RenderDrawingWithOverlaidImages(StringBuilder sb, Drawing groupDrawing, List<Drawing> overlaidImages)
    {
        if (overlaidImages.Count == 0)
        {
            RenderDrawingHtml(sb, groupDrawing, null);
            return;
        }

        RenderDrawingHtml(sb, groupDrawing, overlaidImages);
    }

    // ==================== Drawing Rendering (images, groups, shapes) ====================

    /// <summary>Check if a paragraph contains drawings with actual text box content (txbxContent).</summary>
    private static bool HasTextBoxContent(Paragraph para)
    {
        foreach (var run in para.Elements<Run>())
        {
            var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();
            if (drawing != null && HasTextBox(drawing))
                return true;
        }
        return false;
    }

    /// <summary>Check if paragraph contains any drawing that renders as block-level HTML (text box, chart, shape).</summary>
    private static bool HasBlockLevelDrawing(Paragraph para)
    {
        // Check all descendants (including inside mc:AlternateContent)
        foreach (var drawing in para.Descendants<Drawing>())
        {
            if (HasGroupOrShape(drawing)) return true;
            if (drawing.Descendants().Any(e => e.LocalName == "chart")) return true;
        }
        // Also check for text box content via localName (catches mc:AlternateContent cases)
        if (para.Descendants().Any(e => e.LocalName == "txbxContent"))
            return true;
        return false;
    }

    /// <summary>Check if a drawing contains groups or shapes (for rendering).</summary>
    private static bool HasGroupOrShape(Drawing drawing)
    {
        return drawing.Descendants().Any(e => e.LocalName == "wgp" || e.LocalName == "wsp");
    }

    /// <summary>Check if a drawing contains actual text box content with text (not empty decorative shapes).</summary>
    private static bool HasTextBox(Drawing drawing)
    {
        foreach (var txbx in drawing.Descendants().Where(e => e.LocalName == "txbxContent"))
        {
            // Check if any paragraph inside has actual text
            if (txbx.Descendants<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text)))
                return true;
        }
        return false;
    }

    private void RenderDrawingHtml(StringBuilder sb, Drawing drawing, List<Drawing>? floatImages = null)
    {
        // Check for chart (c:chart inside a:graphicData)
        var chartRef = drawing.Descendants().FirstOrDefault(e => e.LocalName == "chart" &&
            e.GetAttributes().Any(a => a.LocalName == "id"));
        if (chartRef != null)
        {
            RenderChartHtml(sb, drawing, chartRef);
            return;
        }

        // Check for groups/shapes first (text boxes, decorated shapes)
        var group = drawing.Descendants().FirstOrDefault(e => e.LocalName == "wgp");
        if (group != null)
        {
            // Get overall extent from wp:inline or wp:anchor
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            long groupWidthEmu = extent?.Cx?.Value ?? 0;
            long groupHeightEmu = extent?.Cy?.Value ?? 0;

            if (groupWidthEmu > 0 && groupHeightEmu > 0)
            {
                RenderGroupHtml(sb, group, groupWidthEmu, groupHeightEmu, floatImages);
                return;
            }
        }

        // Check for standalone shape (wsp without group)
        var shape = drawing.Descendants().FirstOrDefault(e => e.LocalName == "wsp");
        if (shape != null)
        {
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            long shapeWidth = extent?.Cx?.Value ?? 0;
            long shapeHeight = extent?.Cy?.Value ?? 0;
            if (shapeWidth > 0 && shapeHeight > 0)
            {
                // Full-page shapes → render as background layer
                if (IsFullPageSize(shapeWidth, shapeHeight))
                {
                    var fillCss = ResolveShapeFillCss(shape.Elements().FirstOrDefault(e => e.LocalName == "spPr"));
                    if (!string.IsNullOrEmpty(fillCss))
                        sb.Append($"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;{fillCss}\"></div>");
                    return;
                }
                // Standalone shape — render as inline block, not absolute positioned
                RenderStandaloneShapeHtml(sb, shape, shapeWidth, shapeHeight, floatImages);
                return;
            }
        }

        // Fall back to image rendering
        RenderImageHtml(sb, drawing);
    }

    private void RenderImageHtml(StringBuilder sb, Drawing drawing)
    {
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value == null) return;

        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return;

        try
        {
            var imagePart = mainPart.GetPartById(blip.Embed.Value) as ImagePart;
            if (imagePart == null) return;

            var contentType = imagePart.ContentType;
            using var stream = imagePart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());

            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault()
                ?? drawing.Descendants<A.Extents>().FirstOrDefault() as OpenXmlElement;
            long imgCxEmu = 0, imgCyEmu = 0;
            if (extent is DW.Extent dwExt) { imgCxEmu = dwExt.Cx?.Value ?? 0; imgCyEmu = dwExt.Cy?.Value ?? 0; }
            else if (extent is A.Extents aExt) { imgCxEmu = aExt.Cx?.Value ?? 0; imgCyEmu = aExt.Cy?.Value ?? 0; }

            var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            var alt = docProps?.Description?.Value ?? docProps?.Name?.Value ?? "image";
            var dataUri = $"data:{contentType};base64,{base64}";

            // Detect full-page background images → render as absolute background
            if (IsFullPageSize(imgCxEmu, imgCyEmu))
            {
                sb.Append($"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;overflow:hidden\">");
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncode(alt)}\" style=\"width:100%;height:100%;object-fit:cover\">");
                sb.Append("</div>");
                return;
            }

            var widthPx = imgCxEmu / 9525;
            var heightPx = imgCyEmu / 9525;
            string widthAttr = widthPx > 0 ? $" width=\"{widthPx}\"" : "";
            string heightAttr = heightPx > 0 ? $" height=\"{heightPx}\"" : "";

            // Detect anchored/floating positioning
            var anchor = drawing.Descendants<DW.Anchor>().FirstOrDefault();
            var floatCss = "";
            if (anchor != null)
            {
                // Check wrap type for float direction
                var wrapLeft = anchor.Elements().Any(e => e.LocalName == "wrapSquare" || e.LocalName == "wrapTight");
                if (wrapLeft)
                {
                    var hPosFrom = anchor.GetFirstChild<DW.HorizontalPosition>()?.RelativeFrom?.Value;
                    floatCss = hPosFrom == DW.HorizontalRelativePositionValues.RightMargin
                        ? "float:right;margin:0 0 8px 8px"
                        : "float:left;margin:0 8px 8px 0";
                }
            }

            // Crop support: container-based cropping
            var crop = GetCropPercents(drawing);
            var styleParts = new List<string> { "max-width:100%", "height:auto" };
            if (!string.IsNullOrEmpty(floatCss)) styleParts.Add(floatCss);

            if (crop.HasValue)
            {
                RenderCroppedImage(sb, dataUri, widthPx, heightPx, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, HtmlEncode(alt), floatCss);
            }
            else
            {
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncode(alt)}\"{widthAttr}{heightAttr} style=\"{string.Join(";", styleParts)}\">");
            }
        }
        catch
        {
            sb.Append("<span class=\"img-error\">[Image]</span>");
        }
    }

    /// <summary>
    /// Get crop percentages from a:srcRect.
    /// Values are in 1/1000 of a percent (e.g., 25000 = 25%).
    /// Negative values mean extend (treated as 0).
    /// Returns (left, top, right, bottom) as CSS percentages, or null if no crop.
    /// </summary>
    private static (double l, double t, double r, double b)? GetCropPercents(OpenXmlElement container)
    {
        var srcRect = container.Descendants().FirstOrDefault(e => e.LocalName == "srcRect");
        if (srcRect == null) return null;

        var l = Math.Max(0, GetIntAttr(srcRect, "l") / 1000.0);
        var t = Math.Max(0, GetIntAttr(srcRect, "t") / 1000.0);
        var r = Math.Max(0, GetIntAttr(srcRect, "r") / 1000.0);
        var b = Math.Max(0, GetIntAttr(srcRect, "b") / 1000.0);

        if (l == 0 && t == 0 && r == 0 && b == 0) return null;
        return (l, t, r, b);
    }

    /// <summary>
    /// Render a cropped image using a container div with overflow:hidden.
    /// The image is scaled to its original size and positioned to show only the cropped region.
    /// </summary>
    private static void RenderCroppedImage(StringBuilder sb, string dataUri, long displayWidthPx, long displayHeightPx,
        double cropL, double cropT, double cropR, double cropB, string alt, string extraStyle = "")
    {
        // The display size is the cropped result size.
        // Original image visible fraction: (1 - cropL/100 - cropR/100) horizontally, (1 - cropT/100 - cropB/100) vertically.
        var fracW = 1.0 - cropL / 100.0 - cropR / 100.0;
        var fracH = 1.0 - cropT / 100.0 - cropB / 100.0;
        if (fracW <= 0) fracW = 1; if (fracH <= 0) fracH = 1;

        // Original image size in CSS
        var imgW = displayWidthPx / fracW;
        var imgH = displayHeightPx / fracH;
        // Offset to show the cropped region
        var offsetX = -imgW * (cropL / 100.0);
        var offsetY = -imgH * (cropT / 100.0);

        var containerStyle = $"display:inline-block;width:{displayWidthPx}px;height:{displayHeightPx}px;overflow:hidden";
        if (!string.IsNullOrEmpty(extraStyle)) containerStyle += $";{extraStyle}";
        sb.Append($"<div style=\"{containerStyle}\">");
        sb.Append($"<img src=\"{dataUri}\" alt=\"{alt}\" style=\"width:{imgW:0}px;height:{imgH:0}px;margin-left:{offsetX:0}px;margin-top:{offsetY:0}px\">");
        sb.Append("</div>");
    }

    private static int GetIntAttr(OpenXmlElement el, string attrName)
    {
        var val = el.GetAttributes().FirstOrDefault(a => a.LocalName == attrName).Value;
        return val != null && int.TryParse(val, out var v) ? v : 0;
    }

    // ==================== Group / Shape Rendering ====================

    private void RenderGroupHtml(StringBuilder sb, OpenXmlElement group, long groupWidthEmu, long groupHeightEmu,
        List<Drawing>? floatImages = null)
    {
        var widthPx = groupWidthEmu / 9525;
        var heightPx = groupHeightEmu / 9525;

        // Get the group's child coordinate space from grpSpPr > xfrm
        long chOffX = 0, chOffY = 0, chExtCx = groupWidthEmu, chExtCy = groupHeightEmu;
        var grpSpPr = group.Elements().FirstOrDefault(e => e.LocalName == "grpSpPr");
        var grpXfrm = grpSpPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
        if (grpXfrm != null)
        {
            var chOff = grpXfrm.Elements().FirstOrDefault(e => e.LocalName == "chOff");
            var chExt = grpXfrm.Elements().FirstOrDefault(e => e.LocalName == "chExt");
            if (chOff != null)
            {
                chOffX = GetLongAttr(chOff, "x");
                chOffY = GetLongAttr(chOff, "y");
            }
            if (chExt != null)
            {
                chExtCx = GetLongAttr(chExt, "cx");
                chExtCy = GetLongAttr(chExt, "cy");
            }
        }

        sb.Append($"<div class=\"wg\" style=\"position:relative;width:{widthPx}px;height:{heightPx}px;display:inline-block;overflow:hidden\">");

        // Render each child element (shapes, pictures, nested groups)
        foreach (var child in group.Elements())
        {
            if (child.LocalName is "wsp" or "pic" or "grpSp")
            {
                // Get transform from xfrm (may be in spPr or grpSpPr)
                var xfrm = child.Descendants().FirstOrDefault(e => e.LocalName == "xfrm");
                long offX = 0, offY = 0, extCx = 0, extCy = 0;
                if (xfrm != null)
                {
                    var off = xfrm.Elements().FirstOrDefault(e => e.LocalName == "off");
                    var ext = xfrm.Elements().FirstOrDefault(e => e.LocalName == "ext");
                    if (off != null) { offX = GetLongAttr(off, "x"); offY = GetLongAttr(off, "y"); }
                    if (ext != null) { extCx = GetLongAttr(ext, "cx"); extCy = GetLongAttr(ext, "cy"); }
                }

                // Pass floatImages to first text box shape, then clear
                RenderShapeHtml(sb, child, offX - chOffX, offY - chOffY, extCx, extCy, chExtCx, chExtCy, floatImages);
                floatImages = null; // only inject into first shape
            }
        }

        sb.Append("</div>");
    }

    private void RenderStandaloneShapeHtml(StringBuilder sb, OpenXmlElement shape, long widthEmu, long heightEmu,
        List<Drawing>? floatImages)
    {
        var widthPx = widthEmu / 9525;
        var heightPx = heightEmu / 9525;
        var spPr = shape.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        var fillCss = ResolveShapeFillCss(spPr);
        var borderCss = ResolveShapeBorderCss(spPr);
        var txbx = shape.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");

        var style = $"display:inline-block;width:{widthPx}px;min-height:{heightPx}px;vertical-align:top";
        if (!string.IsNullOrEmpty(fillCss)) style += $";{fillCss}";
        if (!string.IsNullOrEmpty(borderCss)) style += $";{borderCss}";

        var bodyPr = shape.Elements().FirstOrDefault(e => e.LocalName == "bodyPr");
        var lIns = GetLongAttr(bodyPr, "lIns", 91440);
        var tIns = GetLongAttr(bodyPr, "tIns", 45720);
        var rIns = GetLongAttr(bodyPr, "rIns", 91440);
        var bIns = GetLongAttr(bodyPr, "bIns", 45720);
        style += $";padding:{tIns / 9525}px {rIns / 9525}px {bIns / 9525}px {lIns / 9525}px";

        sb.Append($"<div style=\"{style}\">");
        if (txbx != null)
        {
            foreach (var para in txbx.Descendants<Paragraph>())
                RenderParagraphHtml(sb, para);
        }
        else
        {
            var embedAttr = FindEmbedInDescendants(shape);
            if (embedAttr != null)
            {
                try
                {
                    var imagePart = _doc.MainDocumentPart?.GetPartById(embedAttr) as ImagePart;
                    if (imagePart != null)
                    {
                        using var stream = imagePart.GetStream();
                        using var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        sb.Append($"<img src=\"data:{imagePart.ContentType};base64,{Convert.ToBase64String(ms.ToArray())}\" style=\"width:100%;height:auto\">");
                    }
                }
                catch { }
            }
        }
        sb.Append("</div>");
    }

    private void RenderShapeHtml(StringBuilder sb, OpenXmlElement shape, long offX, long offY,
        long extCx, long extCy, long coordSpaceCx, long coordSpaceCy,
        List<Drawing>? floatImages = null)
    {
        // Convert child coordinates to percentage of group
        double leftPct = coordSpaceCx > 0 ? (double)offX / coordSpaceCx * 100 : 0;
        double topPct = coordSpaceCy > 0 ? (double)offY / coordSpaceCy * 100 : 0;
        double widthPct = coordSpaceCx > 0 ? (double)extCx / coordSpaceCx * 100 : 100;
        double heightPct = coordSpaceCy > 0 ? (double)extCy / coordSpaceCy * 100 : 100;

        // Get fill color
        var spPr = shape.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        var fillCss = ResolveShapeFillCss(spPr);

        // Get border
        var borderCss = ResolveShapeBorderCss(spPr);

        // pic elements are always images, not text boxes
        var txbx = shape.LocalName == "pic" ? null
            : shape.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");

        // Rotation from xfrm rot attribute (60000ths of a degree)
        var xfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
        var rot = GetLongAttr(xfrm, "rot");
        var rotCss = rot != 0 ? $";transform:rotate({rot / 60000.0:0.##}deg)" : "";

        // Build style
        var style = $"position:absolute;left:{leftPct:0.##}%;top:{topPct:0.##}%;width:{widthPct:0.##}%;height:{heightPct:0.##}%";
        if (!string.IsNullOrEmpty(fillCss)) style += $";{fillCss}";
        if (!string.IsNullOrEmpty(borderCss)) style += $";{borderCss}";
        style += rotCss;

        // Get body properties for text layout
        var bodyPr = shape.Elements().FirstOrDefault(e => e.LocalName == "bodyPr");
        var vAnchor = bodyPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "anchor").Value;
        if (vAnchor == "ctr") style += ";display:flex;align-items:center";
        else if (vAnchor == "b") style += ";display:flex;align-items:flex-end";

        // Padding from bodyPr insets (EMU → px)
        var lIns = GetLongAttr(bodyPr, "lIns", 91440);
        var tIns = GetLongAttr(bodyPr, "tIns", 45720);
        var rIns = GetLongAttr(bodyPr, "rIns", 91440);
        var bIns = GetLongAttr(bodyPr, "bIns", 45720);
        style += $";padding:{tIns / 9525}px {rIns / 9525}px {bIns / 9525}px {lIns / 9525}px";

        sb.Append($"<div style=\"{style}\">");

        if (txbx != null)
        {
            // Render text box content (standard Word paragraphs)
            sb.Append("<div style=\"width:100%\">");

            // Inject pending float images into this text box
            if (floatImages != null && floatImages.Count > 0)
            {
                foreach (var imgDrawing in floatImages)
                {
                    var imgBlip = imgDrawing.Descendants<A.Blip>().FirstOrDefault();
                    if (imgBlip?.Embed?.Value == null) continue;
                    try
                    {
                        var imgPart = _doc.MainDocumentPart?.GetPartById(imgBlip.Embed.Value) as ImagePart;
                        if (imgPart == null) continue;
                        using var imgStream = imgPart.GetStream();
                        using var imgMs = new MemoryStream();
                        imgStream.CopyTo(imgMs);
                        var imgBase64 = Convert.ToBase64String(imgMs.ToArray());
                        var imgExtent = imgDrawing.Descendants<DW.Extent>().FirstOrDefault();
                        var imgW = imgExtent?.Cx?.Value > 0 ? imgExtent.Cx.Value / 9525 : 100;
                        var imgH = imgExtent?.Cy?.Value > 0 ? imgExtent.Cy.Value / 9525 : 100;
                        var imgDataUri = $"data:{imgPart.ContentType};base64,{imgBase64}";
                        // Read distT/distB/distL/distR for image margins (EMU)
                        var inline = imgDrawing.Descendants<DW.Inline>().FirstOrDefault();
                        var anchor = imgDrawing.Descendants<DW.Anchor>().FirstOrDefault();
                        long distT = 0, distB = 0, distL = 0, distR = 0;
                        if (inline != null)
                        {
                            distT = (long)(inline.DistanceFromTop?.Value ?? 0);
                            distB = (long)(inline.DistanceFromBottom?.Value ?? 0);
                            distL = (long)(inline.DistanceFromLeft?.Value ?? 0);
                            distR = (long)(inline.DistanceFromRight?.Value ?? 0);
                        }
                        else if (anchor != null)
                        {
                            distT = (long)(anchor.DistanceFromTop?.Value ?? 0);
                            distB = (long)(anchor.DistanceFromBottom?.Value ?? 0);
                            distL = (long)(anchor.DistanceFromLeft?.Value ?? 0);
                            distR = (long)(anchor.DistanceFromRight?.Value ?? 0);
                        }
                        var marginCss = $"margin:{distT/9525}px {distR/9525}px {distB/9525}px {distL/9525}px";
                        var crop = GetCropPercents(imgDrawing);
                        if (crop.HasValue)
                        {
                            sb.Append($"<div style=\"float:left;{marginCss}\">");
                            RenderCroppedImage(sb, imgDataUri, imgW, imgH, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, "");
                            sb.Append("</div>");
                        }
                        else
                        {
                            sb.Append($"<img src=\"{imgDataUri}\" style=\"float:left;width:{imgW}px;height:{imgH}px;object-fit:cover;{marginCss}\">");
                        }
                    }
                    catch { }
                }
                floatImages = null;
            }

            foreach (var para in txbx.Descendants<Paragraph>())
            {
                RenderParagraphHtml(sb, para);
            }
            sb.Append("</div>");
        }
        else
        {
            // Check for image inside shape (spPr blip or pic:blipFill blip)
            var embedAttr = FindEmbedInDescendants(shape);
            if (embedAttr != null)
            {
                try
                {
                    var mainPart = _doc.MainDocumentPart;
                    var imagePart = mainPart?.GetPartById(embedAttr) as ImagePart;
                    if (imagePart != null)
                    {
                        using var stream = imagePart.GetStream();
                        using var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        var base64 = Convert.ToBase64String(ms.ToArray());
                        sb.Append($"<img src=\"data:{imagePart.ContentType};base64,{base64}\" style=\"width:100%;height:100%;object-fit:contain\">");
                    }
                }
                catch { }
            }
        }

        sb.Append("</div>");
    }

}
