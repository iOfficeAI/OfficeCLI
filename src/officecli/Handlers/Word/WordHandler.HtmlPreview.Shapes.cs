// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        // LocalName-based sweep over all descendants. Typed Descendants<Drawing>()
        // misses w:drawing nodes whose AddTextbox emit re-declared xmlns:w inline
        // (parser treats them as untyped). LocalName matches across both shapes.
        foreach (var e in para.Descendants())
        {
            var ln = e.LocalName;
            if (ln == "wsp" || ln == "wgp" || ln == "chart" || ln == "txbxContent")
                return true;
        }
        // Some drawings ship with inline xmlns:w re-declaration on <w:drawing>,
        // which causes the SDK to materialize them as untyped OpenXmlUnknownElement
        // with no child Descendants() walk past the drawing root. Fall back to
        // raw OuterXml substring scan so inline-redeclared drawings still register.
        var rawOuter = para.OuterXml;
        if (rawOuter.IndexOf("<wps:wsp", StringComparison.Ordinal) >= 0
            || rawOuter.IndexOf(":txbxContent", StringComparison.Ordinal) >= 0
            || rawOuter.IndexOf("<wpg:", StringComparison.Ordinal) >= 0)
            return true;
        return false;
    }

    /// <summary>Find VML horizontal rule shape in a paragraph (w:pict > v:rect/v:line with o:hr="t").</summary>
    private static OpenXmlElement? FindVmlHorizontalRule(Paragraph para)
    {
        // Search all descendants to handle both direct w:pict and mc:AlternateContent wrapping
        foreach (var pict in para.Descendants().Where(e => e.LocalName == "pict"))
        {
            var hrShape = pict.ChildElements.FirstOrDefault(c =>
                (c.LocalName == "rect" || c.LocalName == "line") &&
                c.GetAttributes().Any(a => a.LocalName == "hr" && a.Value == "t"));
            if (hrShape != null) return hrShape;
        }
        return null;
    }

    /// <summary>Check if a paragraph contains a VML horizontal rule.</summary>
    private static bool IsVmlHorizontalRule(Paragraph para) => FindVmlHorizontalRule(para) != null;

    /// <summary>Render a VML horizontal rule as an HTML hr element.</summary>
    private static void RenderVmlHorizontalRule(StringBuilder sb, Paragraph para)
    {
        var shape = FindVmlHorizontalRule(para)!;

        // Color from fillcolor attribute
        var fillColor = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "fillcolor").Value ?? "#a0a0a0";
        if (!fillColor.StartsWith("#")) fillColor = "#" + fillColor;

        // Height from VML style (e.g. style="width:0;height:1.5pt")
        var heightPx = 1.5;
        var vmlStyle = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
        if (vmlStyle != null)
        {
            var hMatch = System.Text.RegularExpressions.Regex.Match(vmlStyle, @"height:\s*([\d.]+)pt");
            if (hMatch.Success && double.TryParse(hMatch.Groups[1].Value, out var hPt))
                heightPx = hPt;
        }

        // Width percentage from o:hrpct (value in tenths of a percent, e.g. 1000 = 100%)
        var widthCss = "100%";
        var hrpct = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "hrpct").Value;
        if (hrpct != null && int.TryParse(hrpct, out var pctVal) && pctVal > 0 && pctVal < 1000)
            widthCss = $"{pctVal / 10.0:0.#}%";

        // Alignment from o:hralign
        var align = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "hralign").Value ?? "center";
        var marginCss = align switch
        {
            "left" => "margin:0.5em auto 0.5em 0",
            "right" => "margin:0.5em 0 0.5em auto",
            _ => "margin:0.5em auto"
        };

        sb.AppendLine($"<hr style=\"border:none;border-top:{heightPx:0.#}px solid {fillColor};width:{widthCss};{marginCss}\">");
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
                // Full-page shapes → render as a page-level background layer.
                // The fill div is position:absolute;height:100%, so its
                // height resolves against the nearest POSITIONED ancestor.
                // Emitting it inline keeps it inside the host paragraph; when
                // that paragraph ALSO anchors a non-full-page sub-paragraph
                // shape (e.g. a checkbox/connector/logo with its own
                // posOffset), BuildParagraphOpenTag makes the <p>
                // position:relative for those siblings, and the fill's
                // height:100% then collapses onto the ~19px paragraph box
                // instead of the page. Hoist the fill div to page start via
                // the TopAnchoredImages marker (same mechanism the top-anchored
                // background IMAGE path uses) so it becomes a direct child of
                // .page (position:relative, definite min-height) and covers the
                // full page regardless of host-paragraph siblings.
                if (IsFullPageSize(shapeWidth, shapeHeight))
                {
                    var fillCss = ResolveShapeFillCss(shape.Elements().FirstOrDefault(e => e.LocalName == "spPr"));
                    if (!string.IsNullOrEmpty(fillCss))
                    {
                        var fillDiv = $"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;{fillCss}\"></div>";
                        var markerId = $"TOP_ANCHOR_{_ctx.TopAnchoredImages.Count}";
                        _ctx.TopAnchoredImages.Add((markerId, fillDiv));
                        sb.Append($"<!--{markerId}-->");
                    }
                    return;
                }
                // wrapNone shape anchored relative to the column/paragraph with
                // explicit posOffsets (e.g. checkbox rectangles floated over a
                // label list) → absolutely position it from the host paragraph's
                // top-left so each shape lands at its posOffset instead of
                // stacking inline at the cell's left edge. The host paragraph is
                // made position:relative in BuildParagraphOpenTag.
                var paraAbsCss = ComputeParagraphAnchorAbsoluteCss(drawing);
                if (paraAbsCss != null)
                {
                    RenderStandaloneShapeHtml(sb, shape, shapeWidth, shapeHeight, floatImages, paraAbsCss);
                    return;
                }
                // Anchored (floating) shape/textbox with wrapSquare/wrapTight
                // must float so following text wraps beside it — mirror the
                // anchored-image float logic. Inline shapes and
                // wrapNone/behind/in-front keep inline-block positioning.
                var floatCss = ComputeAnchorWrapFloatCss(drawing, shapeWidth);
                RenderStandaloneShapeHtml(sb, shape, shapeWidth, shapeHeight, floatImages, floatCss);
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

        // Prefer the SVG extension rel if present (Office 2019+ keeps a PNG
        // raster in Embed plus an SVG via a:extLst/asvg:svgBlip). PNG fallback
        // is often a 1×1 transparent pixel that renders as a blank, so SVG
        // wins for modern documents that embed vector art.
        string blipRelId = blip.Embed.Value;
        var svgBlip = blip.Descendants().FirstOrDefault(e => e.LocalName == "svgBlip");
        if (svgBlip != null)
        {
            var svgRel = svgBlip.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "embed" || a.LocalName == "link").Value;
            if (!string.IsNullOrEmpty(svgRel))
                blipRelId = svgRel;
        }
        var dataUri = LoadImageAsDataUri(blipRelId);
        if (dataUri == null) return;

        try
        {

            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault()
                ?? drawing.Descendants<A.Extents>().FirstOrDefault() as OpenXmlElement;
            long imgCxEmu = 0, imgCyEmu = 0;
            if (extent is DW.Extent dwExt) { imgCxEmu = dwExt.Cx?.Value ?? 0; imgCyEmu = dwExt.Cy?.Value ?? 0; }
            else if (extent is A.Extents aExt) { imgCxEmu = aExt.Cx?.Value ?? 0; imgCyEmu = aExt.Cy?.Value ?? 0; }

            var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            var alt = docProps?.Description?.Value ?? docProps?.Name?.Value ?? "image";

            // Detect full-page background images → render as absolute background
            if (IsFullPageSize(imgCxEmu, imgCyEmu))
            {
                sb.Append($"<div style=\"position:absolute;top:0;left:0;width:100%;height:100%;z-index:-1;overflow:hidden\">");
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncodeAttr(alt)}\" style=\"width:100%;height:100%;object-fit:cover\">");
                sb.Append("</div>");
                return;
            }

            var widthPx = imgCxEmu / EmuConverter.EmuPerPx;
            var heightPx = imgCyEmu / EmuConverter.EmuPerPx;
            string widthAttr = widthPx > 0 ? $" width=\"{widthPx}\"" : "";
            string heightAttr = heightPx > 0 ? $" height=\"{heightPx}\"" : "";

            // Detect anchored/floating positioning
            var anchor = drawing.Descendants<DW.Anchor>().FirstOrDefault();
            var floatCss = "";
            if (anchor != null)
            {
                var hPos = anchor.GetFirstChild<DW.HorizontalPosition>();
                var hAlign = hPos?.Descendants().FirstOrDefault(e => e.LocalName == "align")?.InnerText;
                var hPosFrom = hPos?.RelativeFrom?.Value;

                // wrapNone → image floats over (or under) the text rather than
                // wrapping it. behindDoc="1" paints behind the body text like a
                // watermark (negative z-index); behindDoc="0" paints on top
                // (positive z-index). Either way the image is absolutely
                // positioned relative to the .page box (which is position:relative)
                // at the anchored hPosition/vPosition, so the text column flows
                // independently and visually overlaps the image — matching Word.
                if (anchor.Elements().Any(e => e.LocalName == "wrapNone"))
                {
                    RenderWrapNoneOverlayImage(sb, drawing, anchor, dataUri, alt, widthPx, heightPx);
                    return;
                }

                // wrapTopAndBottom → centered block image (no text beside it)
                var wrapTopBottom = anchor.Elements().Any(e => e.LocalName == "wrapTopAndBottom");
                if (wrapTopBottom)
                {
                    floatCss = "display:block;margin:8px auto";
                }
                // wrapSquare / wrapTight / wrapThrough → float left or right
                else if (anchor.Elements().Any(e => e.LocalName == "wrapSquare" || e.LocalName == "wrapTight" || e.LocalName == "wrapThrough"))
                {
                    var isRight = hAlign == "right"
                        || hPosFrom == DW.HorizontalRelativePositionValues.RightMargin;
                    // Also check posOffset — float side follows where the image's
                    // horizontal CENTER lands within the text column. The offset is
                    // interpreted relative to hRelative: margin/column offsets start
                    // at the left text edge (column origin), page offsets start at the
                    // physical page left. Comparing the center against the column
                    // midpoint (not the full-page midpoint) is what makes a
                    // right-half image float:right with text wrapping on its left,
                    // matching Word.
                    if (!isRight && hAlign != "left" && hAlign != "center" && hPos != null)
                    {
                        var offsetEl = hPos.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
                        if (offsetEl != null && long.TryParse(offsetEl.InnerText, out var offsetEmu))
                        {
                            var pg = GetPageLayout();
                            var marginLeftEmu = pg.MarginLeftPt * EmuConverter.EmuPerPoint;
                            var colWidthEmu = (pg.WidthPt - pg.MarginLeftPt - pg.MarginRightPt) * EmuConverter.EmuPerPoint;
                            // Convert the offset to a left-edge coordinate measured
                            // from the column origin.
                            double leftInColEmu = hPosFrom == DW.HorizontalRelativePositionValues.Page
                                ? offsetEmu - marginLeftEmu
                                : offsetEmu; // margin/column/character → already column-relative
                            var imgCenterEmu = leftInColEmu + imgCxEmu / 2.0;
                            isRight = imgCenterEmu > colWidthEmu / 2.0;
                        }
                    }
                    else if (hAlign == "center")
                    {
                        // centered alignment — keep the existing float:left default
                        // (block-centering is handled by wrapTopAndBottom branch).
                    }
                    // #7b: use the anchor's distT/distB/distL/distR for the
                    // float margin instead of a hardcoded 8px. The emu→pt
                    // conversion keeps spacing in line with what Word paints.
                    var distT = (long)(anchor.DistanceFromTop?.Value ?? 0) / EmuConverter.EmuPerPointF;
                    var distB = (long)(anchor.DistanceFromBottom?.Value ?? 0) / EmuConverter.EmuPerPointF;
                    var distL = (long)(anchor.DistanceFromLeft?.Value ?? 0) / EmuConverter.EmuPerPointF;
                    var distR = (long)(anchor.DistanceFromRight?.Value ?? 0) / EmuConverter.EmuPerPointF;
                    // Floor the "inside" margin (right for float:left, left for
                    // float:right) so text always has breathing room.
                    if (isRight)
                    {
                        if (distL < 6) distL = 6;
                    }
                    else
                    {
                        if (distR < 6) distR = 6;
                    }

                    // Anchored at top of margin — emit marker for relocation to page start
                    var vPos = anchor.GetFirstChild<DW.VerticalPosition>();
                    var vAlign = vPos?.Descendants().FirstOrDefault(e => e.LocalName == "align")?.InnerText;
                    var vFrom = vPos?.RelativeFrom?.Value;

                    // Approximate vPosition: an explicit vertical offset relative to
                    // the margin or page pushes the image down. Fold it into the top
                    // float margin (best-effort, not pixel-perfect) — previously the
                    // vertical offset was ignored and the image hugged the top of the
                    // wrapping paragraph.
                    if (vAlign == null && vPos != null &&
                        (vFrom == DW.VerticalRelativePositionValues.Margin
                         || vFrom == DW.VerticalRelativePositionValues.Page
                         || vFrom == DW.VerticalRelativePositionValues.Paragraph
                         || vFrom == DW.VerticalRelativePositionValues.Line))
                    {
                        var vOffEl = vPos.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
                        if (vOffEl != null && long.TryParse(vOffEl.InnerText, out var vOffEmu) && vOffEmu > 0)
                            distT += vOffEmu / EmuConverter.EmuPerPointF;
                    }

                    floatCss = isRight
                        ? $"float:right;margin:{distT:0.#}pt {distR:0.#}pt {distB:0.#}pt {distL:0.#}pt"
                        : $"float:left;margin:{distT:0.#}pt {distR:0.#}pt {distB:0.#}pt {distL:0.#}pt";

                    if (vAlign == "top" && vFrom == DW.VerticalRelativePositionValues.Margin)
                    {
                        var fc = isRight ? "float:right;margin:0 0 8px 8px" : "float:left;margin:0 8px 8px 0";
                        var cropVal = GetCropPercents(drawing);
                        var imgHtml = new StringBuilder();
                        if (cropVal.HasValue)
                            RenderCroppedImage(imgHtml, dataUri, widthPx, heightPx, cropVal.Value.l, cropVal.Value.t, cropVal.Value.r, cropVal.Value.b, HtmlEncodeAttr(alt), fc);
                        else
                            imgHtml.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncodeAttr(alt)}\" width=\"{widthPx}\" height=\"{heightPx}\" style=\"max-width:100%;height:{heightPx}px;{fc}\">");
                        var markerId = $"TOP_ANCHOR_{_ctx.TopAnchoredImages.Count}";
                        _ctx.TopAnchoredImages.Add((markerId, imgHtml.ToString()));
                        sb.Append($"<!--{markerId}-->");
                        return;
                    }
                }
            }

            // Crop support: container-based cropping
            var crop = GetCropPercents(drawing);
            // #7a001: when the image's native width exceeds the page body's
            // content width, drop `max-width:100%` so the image paints at
            // native size and overflows the margin the way Word does.
            // Otherwise `max-width:100%` + explicit width + flex-column parent
            // can collapse the layout slot to zero.
            var pgLayout = GetPageLayout();
            var contentWidthPt = pgLayout.WidthPt - pgLayout.MarginLeftPt - pgLayout.MarginRightPt;
            var imgWidthPt = widthPx * 72.0 / 96.0; // 96 DPI → pt
            var overflows = widthPx > 0 && imgWidthPt > contentWidthPt;
            // When the drawing carries an explicit extent (cx/cy), Word renders
            // at exactly that size even if it distorts the source aspect ratio.
            // Pin the height in px so the browser honors the declared dimensions;
            // `height:auto` would let it recompute height from the source ratio.
            var hasExplicitSize = widthPx > 0 && heightPx > 0;
            var heightStyle = hasExplicitSize ? $"height:{heightPx}px" : "height:auto";
            var styleParts = overflows
                ? new List<string> { $"width:{imgWidthPt:0.#}pt", heightStyle }
                : new List<string> { "max-width:100%", heightStyle };
            if (!string.IsNullOrEmpty(floatCss)) styleParts.Add(floatCss);

            // Picture effects from pic:spPr — rotation, flip, border, shadow
            var spPr = drawing.Descendants().FirstOrDefault(e => e.LocalName == "spPr");
            var effectCss = spPr != null ? GetPictureEffectsCss(spPr) : "";
            if (!string.IsNullOrEmpty(effectCss)) styleParts.Add(effectCss);

            if (crop.HasValue)
            {
                RenderCroppedImage(sb, dataUri, widthPx, heightPx, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, HtmlEncodeAttr(alt), floatCss + (string.IsNullOrEmpty(effectCss) ? "" : ";" + effectCss));
            }
            else
            {
                sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncodeAttr(alt)}\"{widthAttr}{heightAttr} style=\"{string.Join(";", styleParts)}\">");
            }
        }
        catch
        {
            sb.Append("<span class=\"img-error\">[Image]</span>");
        }
    }

    /// <summary>
    /// Render a wrapNone anchored image as an absolutely-positioned overlay so
    /// the body text flows independently of it. behindDoc="1" paints the image
    /// behind the text (negative z-index, watermark-style); behindDoc="0" paints
    /// it on top (positive z-index). Position is computed from the anchor's
    /// hPosition/vPosition offsets relative to the .page box (position:relative).
    /// margin/column/paragraph offsets are measured from the page content edge,
    /// so the page margins are added; page-relative offsets are absolute from
    /// the physical page edge (= the .page padding-box origin).
    /// </summary>
    private void RenderWrapNoneOverlayImage(StringBuilder sb, Drawing drawing, DW.Anchor anchor,
        string dataUri, string alt, long widthPx, long heightPx)
    {
        var pg = GetPageLayout();

        // A header/footer image's wrapNone overlay is emitted INSIDE the
        // .doc-header / .doc-footer band (RenderHeaderFooterHtml sets
        // _ctx.ImageHostPart to the HeaderPart/FooterPart). That band is itself
        // position:absolute (its own containing block for absolute children) and
        // is already inset to the page's left margin and offset to
        // HeaderDistancePt — so the overlay's origin is the band's top-left, NOT
        // the physical page. Adding the body's top/left margins here (the body
        // path below) would push a paragraph/column-relative header logo down
        // into the body area. When hosted in a header/footer, measure
        // column/paragraph offsets from the band origin (0,0) instead.
        bool inHeaderFooter = _ctx.ImageHostPart is HeaderPart or FooterPart;

        var hPos = anchor.GetFirstChild<DW.HorizontalPosition>();
        var vPos = anchor.GetFirstChild<DW.VerticalPosition>();
        var hFrom = hPos?.RelativeFrom?.Value;
        var vFrom = vPos?.RelativeFrom?.Value;

        // Body baselines add the page margins; header/footer baselines are zero
        // because the band is already inset to those margins.
        double leftBase = inHeaderFooter ? 0 : pg.MarginLeftPt;
        double topBase = inHeaderFooter ? 0 : pg.MarginTopPt;

        double leftPt = leftBase;
        // <wp:align> (center/left/right) is an alternative to <wp:posOffset>:
        // Word resolves it against the relativeFrom frame's extent. The
        // wrapSquare/Tight/Through paths above already read this element
        // (e.LocalName == "align"); the overlay path previously ignored it and
        // fell back to leftBase (= 0 in header/footer), pinning a page-centered
        // full-width banner to the left margin and clipping its left edge.
        // Resolve align → an absolute left coordinate the same way: center →
        // (frameW - imgW)/2, right → frameW - imgW, left → 0, all measured from
        // the frame origin, then shifted by the frame's left edge.
        var hAlign = hPos?.Descendants().FirstOrDefault(e => e.LocalName == "align")?.InnerText;
        var hOffEl = hPos?.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
        if (hAlign is "center" or "right" or "left")
        {
            // Frame width + origin per relativeFrom. page → full page width from
            // the physical edge (0); margin/column/character → the content
            // column from the left margin. In a header/footer band the origin is
            // already inset to the left margin, so frameOrigin is 0 there.
            double imgWidthPt = widthPx * 72.0 / 96.0; // 96 DPI px → pt
            double frameWidthPt, frameOriginPt;
            if (hFrom == DW.HorizontalRelativePositionValues.Page)
            {
                frameWidthPt = pg.WidthPt;
                frameOriginPt = inHeaderFooter ? -pg.MarginLeftPt : 0;
            }
            else // margin / column / character / leftMargin / insideMargin / …
            {
                frameWidthPt = pg.WidthPt - pg.MarginLeftPt - pg.MarginRightPt;
                frameOriginPt = leftBase;
            }
            double alignedPt = hAlign switch
            {
                "center" => (frameWidthPt - imgWidthPt) / 2.0,
                "right" => frameWidthPt - imgWidthPt,
                _ => 0, // left
            };
            leftPt = frameOriginPt + alignedPt;
        }
        else if (hOffEl != null && long.TryParse(hOffEl.InnerText, out var hOffEmu))
        {
            leftPt = hFrom == DW.HorizontalRelativePositionValues.Page
                ? hOffEmu / EmuConverter.EmuPerPointF
                : leftBase + hOffEmu / EmuConverter.EmuPerPointF;
        }

        double topPt = topBase;
        var vOffEl = vPos?.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
        if (vOffEl != null && long.TryParse(vOffEl.InnerText, out var vOffEmu))
        {
            topPt = vFrom == DW.VerticalRelativePositionValues.Page
                ? vOffEmu / EmuConverter.EmuPerPointF
                : topBase + vOffEmu / EmuConverter.EmuPerPointF;
        }

        // behindDoc="1" → behind text (watermark); else in front.
        bool behind = anchor.BehindDoc?.Value == true;
        var zIndex = behind ? "-1" : "10";

        var widthAttr = widthPx > 0 ? $" width=\"{widthPx}\"" : "";
        var heightAttr = heightPx > 0 ? $" height=\"{heightPx}\"" : "";
        // Absolutely-positioned overlay: write the declared px dimensions (from
        // the EMU extent) into inline style + max-width:none so the global
        // img{max-width:100%} rule can't clamp the image to the .page width.
        // A behindDoc full-page cover (declared wider than the page) must bleed
        // past the page edges to cover everything; clamping shrank it to ~61%.
        // Both dims come from the same extent, so aspect ratio is preserved.
        var sizeCss = "";
        if (widthPx > 0) sizeCss += $";width:{widthPx}px;max-width:none";
        if (heightPx > 0) sizeCss += $";height:{heightPx}px";
        var style = $"position:absolute;left:{leftPt:0.#}pt;top:{topPt:0.#}pt;z-index:{zIndex}{sizeCss}";

        var crop = GetCropPercents(drawing);
        if (crop.HasValue)
            RenderCroppedImage(sb, dataUri, widthPx, heightPx, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, HtmlEncodeAttr(alt), style);
        else
            sb.Append($"<img src=\"{dataUri}\" alt=\"{HtmlEncodeAttr(alt)}\"{widthAttr}{heightAttr} style=\"{style}\">");
    }

    /// <summary>
    /// Extract CSS for picture visual effects from a:xfrm (rotation, flip),
    /// a:ln (border), and a:effectLst (shadow/glow). All live under pic:spPr.
    /// </summary>
    private static string GetPictureEffectsCss(OpenXmlElement spPr)
    {
        var parts = new List<string>();

        // Rotation + flip from a:xfrm
        var xfrm = spPr.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
        if (xfrm != null)
        {
            var rot = xfrm.GetAttributes().FirstOrDefault(a => a.LocalName == "rot").Value;
            var flipH = xfrm.GetAttributes().FirstOrDefault(a => a.LocalName == "flipH").Value;
            var flipV = xfrm.GetAttributes().FirstOrDefault(a => a.LocalName == "flipV").Value;

            var transforms = new List<string>();
            if (long.TryParse(rot, out var rotVal) && rotVal != 0)
            {
                // OOXML rotation is in 60000ths of a degree
                var deg = rotVal / 60000.0;
                transforms.Add($"rotate({deg:0.##}deg)");
            }
            if (flipH == "1" || flipH == "true") transforms.Add("scaleX(-1)");
            if (flipV == "1" || flipV == "true") transforms.Add("scaleY(-1)");
            if (transforms.Count > 0)
                parts.Add($"transform:{string.Join(" ", transforms)}");
        }

        // Border from a:ln. An <a:ln> with <a:noFill/> (or w="0") is an EXPLICIT
        // declaration of "no outline" — Word renders no border, so we must NOT
        // emit a default one. A bare self-closing <a:ln/> (no w, no fill child)
        // is likewise NOT a paintable outline: with neither a width nor a fill
        // it inherits nothing meaningful for a style-less picture and Word draws
        // nothing, so we must require an explicit width OR fill before emitting.
        var ln = spPr.Elements().FirstOrDefault(e => e.LocalName == "ln");
        if (ln != null)
        {
            var noFill = ln.Elements().Any(e => e.LocalName == "noFill");
            var wAttr = ln.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
            var hasZeroWidth = long.TryParse(wAttr, out var wEmu0) && wEmu0 == 0;
            var hasWidth = long.TryParse(wAttr, out var wEmuPos) && wEmuPos > 0;
            var hasFill = ln.Elements().Any(e => e.LocalName is "solidFill" or "gradFill" or "pattFill");
            if (!noFill && !hasZeroWidth && (hasWidth || hasFill))
            {
                double borderPx = 1;
                if (long.TryParse(wAttr, out var wEmu) && wEmu > 0)
                    borderPx = Math.Max(1, wEmu / EmuConverter.EmuPerPxF); // EMU → px
                var solidFill = ln.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
                var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
                var colorHex = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                // The raw a:srgbClr/@val is document-controlled; it is not a typed
                // SDK value, so nothing has validated it. Only a hex colour may
                // reach the style attribute — anything else falls back to black.
                var borderColor = !string.IsNullOrEmpty(colorHex) && IsHexColor(colorHex) ? $"#{colorHex}" : "#000";
                parts.Add($"border:{borderPx:0.##}px solid {borderColor}");
            }
        }

        // Outer shadow from a:effectLst/a:outerShdw — map to box-shadow
        var shadowCss = ResolveOuterShadowCss(spPr);
        if (!string.IsNullOrEmpty(shadowCss)) parts.Add(shadowCss);

        return string.Join(";", parts);
    }

    /// <summary>
    /// Map a shape/picture spPr's a:effectLst/a:outerShdw to a CSS box-shadow.
    /// Returns "" when no outer shadow is present. Shared by the picture path
    /// and the wps shape style builder so both render drop shadows identically.
    /// </summary>
    private static string ResolveOuterShadowCss(OpenXmlElement? spPr)
    {
        var effectLst = spPr?.Elements().FirstOrDefault(e => e.LocalName == "effectLst");
        var outerShdw = effectLst?.Elements().FirstOrDefault(e => e.LocalName == "outerShdw");
        if (outerShdw == null) return "";

        // blurRad, dist, dir (60000ths of a degree) — simplified offset projection
        var blurAttr = outerShdw.GetAttributes().FirstOrDefault(a => a.LocalName == "blurRad").Value;
        var distAttr = outerShdw.GetAttributes().FirstOrDefault(a => a.LocalName == "dist").Value;
        var dirAttr = outerShdw.GetAttributes().FirstOrDefault(a => a.LocalName == "dir").Value;
        double blurPx = long.TryParse(blurAttr, out var blurEmu) ? blurEmu / EmuConverter.EmuPerPxF : 4;
        double distPx = long.TryParse(distAttr, out var distEmu) ? distEmu / EmuConverter.EmuPerPxF : 4;
        double dirDeg = long.TryParse(dirAttr, out var dirVal) ? dirVal / 60000.0 : 45;
        var offX = distPx * Math.Cos(dirDeg * Math.PI / 180);
        var offY = distPx * Math.Sin(dirDeg * Math.PI / 180);
        var shdwFill = outerShdw.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var shdwHex = shdwFill?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value ?? "000000";
        // Document-controlled raw attribute: only a hex colour may reach CSS.
        if (!IsHexColor(shdwHex)) shdwHex = "000000";
        return $"box-shadow:{offX:0.#}px {offY:0.#}px {blurPx:0.#}px #{shdwHex}";
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
    /// The image is absolutely positioned inside the container (NOT a baseline-
    /// dependent inline element with negative margins): a 128px-tall inline img in
    /// a 128px container sits on the text baseline, so without a vertical-crop
    /// margin to pull it back (e.g. cropLeft-only crops where margin-top is 0) it
    /// is pushed out of the overflow:hidden window and vanishes entirely. Absolute
    /// positioning ties the offset to the container box, not the line box, so every
    /// crop combination (symmetric / single-side / mixed) clips correctly.
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

        var containerStyle = $"position:relative;display:inline-block;width:{displayWidthPx}px;height:{displayHeightPx}px;overflow:hidden";
        if (!string.IsNullOrEmpty(extraStyle)) containerStyle += $";{extraStyle}";
        sb.Append($"<div style=\"{containerStyle}\">");
        sb.Append($"<img src=\"{dataUri}\" alt=\"{alt}\" style=\"position:absolute;left:{offsetX:0}px;top:{offsetY:0}px;width:{imgW:0}px;height:{imgH:0}px;max-width:none\">");
        sb.Append("</div>");
    }

    private static int GetIntAttr(OpenXmlElement el, string attrName)
    {
        var val = el.GetAttributes().FirstOrDefault(a => a.LocalName == attrName).Value;
        return val != null && int.TryParse(val, out var v) ? v : 0;
    }

    /// <summary>Load an image part by relationship ID and return as a base64 data URI.</summary>
    private string? LoadImageAsDataUri(string relId)
    {
        // Header/footer images store their ImagePart + relationship on the
        // HeaderPart/FooterPart, not MainDocumentPart. Use the host part for
        // the element currently being rendered when set; else fall back to
        // the document part (body path).
        var hostPart = _ctx.ImageHostPart ?? (DocumentFormat.OpenXml.Packaging.OpenXmlPart?)_doc.MainDocumentPart;
        if (hostPart == null) return null;
        return HtmlPreviewHelper.PartToDataUri(hostPart, relId);
    }

    /// <summary>
    /// Resolve the raw content type of an image part by relationship ID, using the
    /// same host-part fallback as LoadImageAsDataUri. Returns null if the part
    /// cannot be found. Callers that must distinguish a genuinely browser-renderable
    /// image from a degraded placeholder (PartToDataUri rewrites WMF/EMF to an SVG
    /// placeholder) should branch on this, not on the returned data URI string.
    /// </summary>
    private string? LoadImageContentType(string relId)
    {
        var hostPart = _ctx.ImageHostPart ?? (DocumentFormat.OpenXml.Packaging.OpenXmlPart?)_doc.MainDocumentPart;
        if (hostPart == null) return null;
        try
        {
            return hostPart.GetPartById(relId)?.ContentType;
        }
        catch
        {
            return null;
        }
    }

    // ==================== Group / Shape Rendering ====================

    private void RenderGroupHtml(StringBuilder sb, OpenXmlElement group, long groupWidthEmu, long groupHeightEmu,
        List<Drawing>? floatImages = null)
    {
        var widthPx = groupWidthEmu / EmuConverter.EmuPerPx;
        var heightPx = groupHeightEmu / EmuConverter.EmuPerPx;

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
        List<Drawing>? floatImages, string? floatCss = null)
    {
        // Standalone shapes use inline positioning with pixel dimensions
        RenderShapeHtml(sb, shape, 0, 0, widthEmu, heightEmu, widthEmu, heightEmu, floatImages, standalone: true, floatCss: floatCss);
    }

    /// <summary>
    /// For an anchored (wp:anchor) drawing with a wrapSquare/wrapTight wrap
    /// type, compute the float CSS (float:left / float:right + margin) so the
    /// following text wraps beside it — mirroring the anchored-image float
    /// logic in RenderImageHtml. Returns null for inline drawings and for
    /// wrapNone / behind-text / in-front-of-text (those keep inline/absolute
    /// positioning). Float side follows the anchor's horizontal position.
    /// </summary>
    private string? ComputeAnchorWrapFloatCss(Drawing drawing, long widthEmu)
    {
        var anchor = drawing.Descendants<DW.Anchor>().FirstOrDefault();
        if (anchor == null) return null;

        // Only square/tight/through wrap floats text beside the shape. wrapNone /
        // in-front-of-text legitimately overlap.
        if (!anchor.Elements().Any(e => e.LocalName == "wrapSquare" || e.LocalName == "wrapTight" || e.LocalName == "wrapThrough"))
            return null;

        // Page-anchored shapes with an explicit posOffset (both H and V relative
        // to the page) live at a FIXED page location, not "beside this paragraph".
        // Floating them into the anchoring paragraph's inline flow both
        // mispositions them (they belong at their page coords, often page bottom)
        // and steals horizontal width from that paragraph — e.g. a 36pt cover
        // title forced to wrap in the narrow gap left of a page-bottom address
        // box, which then mid-word-breaks ("produc/t"). Position such shapes
        // absolutely against the .page (position:relative) so the wrapping text
        // keeps its full column width. Word's own square-wrap of body text around
        // a page-bottom box is negligible here (text ends far above it).
        var vPosPage = anchor.GetFirstChild<DW.VerticalPosition>();
        var hPosPage = anchor.GetFirstChild<DW.HorizontalPosition>();
        if (vPosPage?.RelativeFrom?.Value == DW.VerticalRelativePositionValues.Page
            && hPosPage?.RelativeFrom?.Value == DW.HorizontalRelativePositionValues.Page)
        {
            var vOff = vPosPage.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
            var hOff = hPosPage.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
            if (vOff != null && hOff != null
                && long.TryParse(vOff.InnerText, out var vEmu)
                && long.TryParse(hOff.InnerText, out var hEmu))
            {
                // posOffset is from the physical page edge (0,0); an absolute
                // child resolves against .page's padding box, so subtract the
                // page margin (== .page padding) to convert.
                var pg = GetPageLayout();
                var topPt = vEmu / EmuConverter.EmuPerPointF - pg.MarginTopPt;
                var leftPt = hEmu / EmuConverter.EmuPerPointF - pg.MarginLeftPt;
                return $"position:absolute;top:{topPt:0.#}pt;left:{leftPt:0.#}pt;z-index:1";
            }
        }

        var hPos = anchor.GetFirstChild<DW.HorizontalPosition>();
        var hAlign = hPos?.Descendants().FirstOrDefault(e => e.LocalName == "align")?.InnerText;
        var hPosFrom = hPos?.RelativeFrom?.Value;

        var isRight = hAlign == "right"
            || hPosFrom == DW.HorizontalRelativePositionValues.RightMargin;
        // Mirror the image path: when there's an explicit posOffset, float to
        // the side the shape's horizontal center lands within the text column.
        if (!isRight && hAlign != "left" && hAlign != "center" && hPos != null)
        {
            var offsetEl = hPos.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
            if (offsetEl != null && long.TryParse(offsetEl.InnerText, out var offsetEmu))
            {
                var pg = GetPageLayout();
                var marginLeftEmu = pg.MarginLeftPt * EmuConverter.EmuPerPoint;
                var colWidthEmu = (pg.WidthPt - pg.MarginLeftPt - pg.MarginRightPt) * EmuConverter.EmuPerPoint;
                double leftInColEmu = hPosFrom == DW.HorizontalRelativePositionValues.Page
                    ? offsetEmu - marginLeftEmu
                    : offsetEmu;
                var centerEmu = leftInColEmu + widthEmu / 2.0;
                isRight = centerEmu > colWidthEmu / 2.0;
            }
        }

        var distT = (long)(anchor.DistanceFromTop?.Value ?? 0) / EmuConverter.EmuPerPointF;
        var distB = (long)(anchor.DistanceFromBottom?.Value ?? 0) / EmuConverter.EmuPerPointF;
        var distL = (long)(anchor.DistanceFromLeft?.Value ?? 0) / EmuConverter.EmuPerPointF;
        var distR = (long)(anchor.DistanceFromRight?.Value ?? 0) / EmuConverter.EmuPerPointF;
        // Floor the "inside" margin so text always has breathing room.
        if (isRight) { if (distL < 6) distL = 6; }
        else { if (distR < 6) distR = 6; }

        return isRight
            ? $"float:right;margin:{distT:0.#}pt {distR:0.#}pt {distB:0.#}pt {distL:0.#}pt"
            : $"float:left;margin:{distT:0.#}pt {distR:0.#}pt {distB:0.#}pt {distL:0.#}pt";
    }

    // Horizontal anchor origins that coincide with the text-column left edge
    // (i.e. the start of the cell/paragraph content box). A posOffset relative
    // to any of these is the distance from the paragraph's own left edge, so it
    // can be emitted directly as `left:` inside the position:relative paragraph.
    private static bool IsColumnLeftRelative(DW.HorizontalRelativePositionValues? from)
        => from == DW.HorizontalRelativePositionValues.Column
        || from == DW.HorizontalRelativePositionValues.Character
        || from == DW.HorizontalRelativePositionValues.LeftMargin
        || from == DW.HorizontalRelativePositionValues.InsideMargin;

    // Vertical anchor origins measured from the paragraph/line top — the
    // posOffset is the distance below the paragraph's own top edge, emitted
    // directly as `top:` inside the position:relative paragraph.
    private static bool IsParagraphTopRelative(DW.VerticalRelativePositionValues? from)
        => from == DW.VerticalRelativePositionValues.Paragraph
        || from == DW.VerticalRelativePositionValues.Line;

    /// <summary>
    /// True when the paragraph anchors at least one wrapNone shape positioned
    /// relative to the column/paragraph with explicit H+V posOffsets — the case
    /// ComputeParagraphAnchorAbsoluteCss positions absolutely. Drives the
    /// position:relative on the paragraph's host div so those absolute children
    /// resolve against the paragraph instead of the .page box.
    /// </summary>
    private bool ParagraphAnchorsSubParagraphShape(Paragraph para)
    {
        foreach (var drawing in para.Descendants<Drawing>())
        {
            if (!drawing.Descendants().Any(e => e.LocalName == "wsp"))
                continue;
            // A full-page-size shape is rendered as a page-background fill
            // (RenderDrawingHtml line ~178) whose width/height:100% must resolve
            // against the .page box. If we made the host paragraph
            // position:relative for it, that 100% would collapse onto the single
            // paragraph box (a ~478×448px sliver) instead of covering the page.
            // Such shapes get NO per-paragraph relative containing block.
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            if (extent != null && IsFullPageSize(extent.Cx?.Value ?? 0, extent.Cy?.Value ?? 0))
                continue;
            if (ComputeParagraphAnchorAbsoluteCss(drawing) != null)
                return true;
        }
        return false;
    }

    /// <summary>
    /// True when any paragraph in the table cell anchors a column/paragraph
    /// wrapNone shape positioned absolutely (see ComputeParagraphAnchorAbsoluteCss).
    /// Drives position:relative on the host &lt;td&gt; so those absolute shapes
    /// resolve against the cell content box. Applied on the cell — not the inner
    /// paragraph div — because a relative div whose only in-flow content is
    /// wrapped text inside a table cell collapses the row height to zero.
    /// </summary>
    private bool CellAnchorsSubParagraphShape(TableCell cell)
    {
        foreach (var para in cell.Descendants<Paragraph>())
        {
            if (ParagraphAnchorsSubParagraphShape(para))
                return true;
        }
        return false;
    }

    // Horizontal anchor origins that span the full text column (margin/page) and
    // therefore carry an <wp:align> (left/center/right) we can map to a CSS
    // alignment inside the position:relative paragraph box — as opposed to a
    // posOffset coordinate. The paragraph host box's content width equals the
    // text column, so align=center → centered, left/right → edge-pinned.
    private static bool IsColumnSpanRelative(DW.HorizontalRelativePositionValues? from)
        => from == DW.HorizontalRelativePositionValues.Margin
        || from == DW.HorizontalRelativePositionValues.Page
        || from == DW.HorizontalRelativePositionValues.Column;

    /// <summary>
    /// For a wrapNone shape anchored relative to the column/paragraph, compute
    /// absolute positioning CSS measured from the host paragraph's top-left.
    ///
    /// Horizontal placement comes from EITHER an explicit posOffset (H relative
    /// to column/character/left-margin/inside-margin — the distance from the
    /// paragraph's own left edge) OR an &lt;wp:align&gt; (H relative to
    /// margin/page/column, which spans the text column so the paragraph box's
    /// content width matches): center → left:50%;translateX(-50%); left → 0;
    /// right → pinned to the right edge.
    ///
    /// Vertical placement comes from a posOffset relative to the paragraph/line
    /// (distance below the paragraph's own top edge).
    ///
    /// Returns null for any other anchor shape — page/page (handled against
    /// .page in ComputeAnchorWrapFloatCss), wrapped, or inline — so those paths
    /// keep their existing behaviour.
    ///
    /// This recovers per-shape placement for forms whose checkbox/marker
    /// rectangles float over a label list via column/paragraph posOffsets (in
    /// flow HTML they otherwise collapse to a left-edge ladder), and for
    /// margin-centered floating text boxes positioned a fixed distance below
    /// their anchoring paragraph.
    /// </summary>
    private static string? ComputeParagraphAnchorAbsoluteCss(Drawing drawing)
    {
        var anchor = drawing.Descendants<DW.Anchor>().FirstOrDefault();
        if (anchor == null) return null;

        // Only wrapNone (overlap) shapes — wrapped shapes float and own their
        // own square/tight path; this is purely the over-text overlay case.
        if (!anchor.Elements().Any(e => e.LocalName == "wrapNone")) return null;

        var hPos = anchor.GetFirstChild<DW.HorizontalPosition>();
        var vPos = anchor.GetFirstChild<DW.VerticalPosition>();
        if (hPos == null || vPos == null) return null;

        // Vertical: distance below the paragraph's own top edge (posOffset).
        if (!IsParagraphTopRelative(vPos.RelativeFrom?.Value)) return null;
        var vOff = vPos.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
        if (vOff == null || !long.TryParse(vOff.InnerText, out var vEmu)) return null;
        var topPt = vEmu / EmuConverter.EmuPerPointF;

        // Horizontal: posOffset from the column/paragraph left edge, OR an
        // <wp:align> against a column-spanning origin (margin/page/column).
        string horizCss;
        var hFrom = hPos.RelativeFrom?.Value;
        var hOff = hPos.Descendants().FirstOrDefault(e => e.LocalName == "posOffset");
        if (IsColumnLeftRelative(hFrom) && hOff != null
            && long.TryParse(hOff.InnerText, out var hEmu))
        {
            var leftPt = hEmu / EmuConverter.EmuPerPointF;
            horizCss = $"left:{leftPt:0.#}pt";
        }
        else if (IsColumnSpanRelative(hFrom))
        {
            var hAlign = hPos.Descendants().FirstOrDefault(e => e.LocalName == "align")?.InnerText;
            horizCss = hAlign switch
            {
                "center" => "left:50%;transform:translateX(-50%)",
                "left" => "left:0",
                "right" => "right:0",
                _ => "",  // inside/outside or no align → not resolvable here
            };
            if (horizCss.Length == 0) return null;
        }
        else
        {
            return null;
        }

        // behindDoc="1" → under the text (e.g. shaded marker); else over it.
        var z = anchor.BehindDoc?.Value == true ? "-1" : "5";
        return $"position:absolute;{horizCss};top:{topPt:0.#}pt;z-index:{z}";
    }

    /// <summary>
    /// Render a shape element (wsp, pic, grpSp) with either absolute (inside group) or inline (standalone) positioning.
    /// </summary>
    private void RenderShapeHtml(StringBuilder sb, OpenXmlElement shape, long offX, long offY,
        long extCx, long extCy, long coordSpaceCx, long coordSpaceCy,
        List<Drawing>? floatImages = null, bool standalone = false, string? floatCss = null)
    {
        // Common shape properties
        var spPr = shape.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        var fillCss = ResolveShapeFillCss(spPr);
        var borderCss = ResolveShapeBorderCss(spPr);
        var txbx = shape.LocalName == "pic" ? null
            : shape.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");

        // Build positioning style
        string style;
        if (standalone)
        {
            var widthPx = extCx / EmuConverter.EmuPerPx;
            var heightPx = extCy / EmuConverter.EmuPerPx;
            // A shape that receives overlaid header images (floatImages, e.g. a
            // cover banner + logo floated into a header text box) acts as a
            // full-width header container, not a sized box. Shrink-wrapping it
            // to its own (often tiny) text extent makes the global
            // `img{max-width:100%}` rule clamp a 940px banner to the box width —
            // collapsing it to a thin strip. Render it as a non-shrink-wrapping
            // full-width block so the overlay images resolve against the header
            // content width instead. The overlay imgs get `max-width:none`
            // (see the floatImages inject loop) so their declared px width wins.
            bool isOverlayContainer = floatImages is { Count: > 0 };

            // Box sizing model: autofit vs fixed.
            //
            // A fixed-size text box (bodyPr/a:noAutofit, Word "Do not autofit")
            // with a solid fill paints the fill ONLY over its declared height.
            // When its content (e.g. an inner table whose rows exceed the box)
            // overflows — vertOverflow="overflow" — Word draws the overflowing
            // content beyond the box edge WITHOUT extending the fill. Emitting
            // `min-height` here lets the host div grow to the content and paints
            // `background-color` across the whole grown height, so any
            // transparent lower region (e.g. an unshaded table row) exposes the
            // box fill below the real box — a phantom colored band that Word
            // never shows. Pin the declared `height` and clip to the box
            // (overflow:hidden) so the fill — and any content taller than the
            // box — stays confined to the declared height, matching the box
            // Word paints. Only fixed boxes WITH a fill need this; autofit boxes
            // (spAutoFit / normAutofit) and fill-less fixed boxes keep min-height
            // (grow-to-content) so short content doesn't leave a gap.
            // Autofit detection: a box is autofit only when it carries an
            // explicit a:spAutoFit (resize box to text) or a:normAutofit (shrink
            // text to box). Everything else — explicit a:noAutofit OR no autofit
            // child at all (OOXML default == noAutofit) — is a fixed box.
            var bodyPrAf = shape.Elements().FirstOrDefault(e => e.LocalName == "bodyPr");
            bool isAutofitBox = bodyPrAf?.Elements().Any(e =>
                e.LocalName == "spAutoFit" || e.LocalName == "normAutofit") == true;
            bool isFixedBox = !isAutofitBox;
            bool hasFillBg = fillCss.Contains("background", StringComparison.Ordinal);

            // Horizontal overflow: a fixed (noAutofit) box can declare a narrow
            // ext cx yet hold a wider inner table (cover layouts commonly pin a
            // tall/thin box but lay the title + multi-column body in a table
            // whose grid far exceeds the box). bodyPr/@horzOverflow defaults to
            // "overflow"; Word then paints the content at its NATURAL width,
            // spilling past the declared box edge — it does NOT crush the table
            // into the narrow box. Pinning width:{widthPx}px here instead
            // collapses the inner width:100% table into the box and detonates
            // wrapping (one glyph/word per line, hyperlinks sliced). When the
            // box is fixed, horzOverflow is not clipped, and it carries an inner
            // table whose grid width exceeds the declared box width, widen the
            // host to the table's natural width so the content renders un-crushed
            // the way Word draws it. Mirrors the image-overflow precedent above
            // (drop the clamp, paint at native width). Excludes the DRAFT-stamp /
            // classification-banner / overlay-container boxes — none carry an
            // inner table, so this never perturbs R88/R109.
            var horzClip = bodyPrAf?.GetAttributes()
                .Any(a => a.LocalName == "horzOverflow" && a.Value == "clip") == true;
            long boxWidthPx = widthPx;
            if (isFixedBox && !horzClip)
            {
                var innerTbl = txbx?.Elements().FirstOrDefault(e => e.LocalName == "tbl");
                var tblGrid = innerTbl?.Elements().FirstOrDefault(e => e.LocalName == "tblGrid");
                if (tblGrid != null)
                {
                    long gridTwips = 0;
                    foreach (var gc in tblGrid.Elements().Where(e => e.LocalName == "gridCol"))
                        gridTwips += GetLongAttr(gc, "w");
                    // twips → px (1 twip = 1/15 px at 96dpi)
                    int gridPx = (int)(gridTwips / 15);
                    if (gridPx > boxWidthPx) boxWidthPx = gridPx;
                }
            }

            // When the box widens to the table, never clip horizontally (the
            // fill-confining overflow:hidden below would otherwise re-crush it);
            // keep vertical clipping intent via the height path only.
            bool widened = boxWidthPx > widthPx;
            var heightProp = isFixedBox && hasFillBg && !widened
                ? $"height:{heightPx}px;overflow:hidden"
                : $"min-height:{heightPx}px";

            // Anchored wrapSquare/wrapTight shape → float so following text
            // wraps beside it; otherwise inline-block (inline / wrapNone /
            // behind / in-front-of-text).
            style = floatCss != null
                ? $"{floatCss};width:{boxWidthPx}px;{heightProp};box-sizing:border-box"
                : isOverlayContainer
                    ? $"display:block;width:100%;{heightProp}"
                    : $"display:inline-block;width:{boxWidthPx}px;{heightProp};vertical-align:top";

            // Rotation on standalone shapes too (was only applied inside groups)
            var sXfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
            var sRot = GetLongAttr(sXfrm, "rot");
            if (sRot != 0) style += $";transform:rotate({sRot / 60000.0:0.##}deg)";
        }
        else
        {
            double leftPct = coordSpaceCx > 0 ? (double)offX / coordSpaceCx * 100 : 0;
            double topPct = coordSpaceCy > 0 ? (double)offY / coordSpaceCy * 100 : 0;
            double widthPct = coordSpaceCx > 0 ? (double)extCx / coordSpaceCx * 100 : 100;
            double heightPct = coordSpaceCy > 0 ? (double)extCy / coordSpaceCy * 100 : 100;
            style = $"position:absolute;left:{leftPct:0.##}%;top:{topPct:0.##}%;width:{widthPct:0.##}%;height:{heightPct:0.##}%";

            // Rotation (only for positioned shapes inside groups)
            var xfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
            var rot = GetLongAttr(xfrm, "rot");
            if (rot != 0) style += $";transform:rotate({rot / 60000.0:0.##}deg)";
        }

        // prstGeom → border-radius for ellipse, round rect, etc.
        var prstGeom = spPr?.Elements().FirstOrDefault(e => e.LocalName == "prstGeom");
        var prst = prstGeom?.GetAttributes().FirstOrDefault(a => a.LocalName == "prst").Value;
        if (prst == "ellipse" || prst == "oval")
            style += ";border-radius:50%";
        else if (prst == "roundRect")
            style += ";border-radius:12px";

        // #7a: for complex preset geometries (line, arrows, callouts) the
        // background/border approach collapses to a plain rect. Render
        // those as inline SVG overlays using the shape's fill/border colors.
        var svgPrst = prst is "line" or "straightConnector1"
            or "rightArrow" or "leftArrow" or "upArrow" or "downArrow"
            or "wedgeRoundRectCallout"
            or "diamond" or "flowChartDecision";
        if (svgPrst)
        {
            // Defer fill/border to the SVG so the host div stays transparent.
            style += ";overflow:visible";

            // The overlay SVG uses height:100%, which only resolves when the
            // host div has a *definite* height. The standalone path emits
            // `min-height:{h}px` (grow-to-content) — not a definite height —
            // so an SVG with viewBox 0 0 100 100 and width:100% falls back to
            // its 1:1 intrinsic aspect ratio and renders as a tall square. For
            // an extremely wide/short connector (e.g. a signature line:
            // cx=4524375 cy=9525 EMU → 475px × 1px), that square turns the
            // box-diagonal line endpoint (0,0→100,100) into a long page-spanning
            // diagonal instead of a near-horizontal stroke. Pin a definite
            // height equal to the shape's ext cy so the SVG squashes to the real
            // box, collapsing the diagonal to the connector's true orientation.
            // (The positioned/group path already emits a definite `height:%`.)
            if (standalone)
            {
                // Clamp to >=1px: a perfectly horizontal connector (cy≈0) would
                // otherwise collapse the box to 0px and hide the stroke.
                var svgHeightPx = Math.Max(1, extCy / EmuConverter.EmuPerPx);
                style = System.Text.RegularExpressions.Regex.Replace(
                    style, @"min-height:\d+px", $"height:{svgHeightPx}px");
            }
        }
        else
        {
            if (!string.IsNullOrEmpty(fillCss)) style += $";{fillCss}";
            if (!string.IsNullOrEmpty(borderCss)) style += $";{borderCss}";
        }

        // Outer shadow (a:effectLst/a:outerShdw) → box-shadow. Shares the
        // picture path's projection so wps shapes and pictures drop shadows
        // identically. Applies to the host div even for svg-overlay presets.
        var shadowCss = ResolveOuterShadowCss(spPr);
        if (!string.IsNullOrEmpty(shadowCss)) style += $";{shadowCss}";

        // Body properties: text layout + padding
        var bodyPr = shape.Elements().FirstOrDefault(e => e.LocalName == "bodyPr");
        // Vertical text anchor applies to both standalone and positioned shapes
        var vAnchor = bodyPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "anchor").Value;
        if (vAnchor == "ctr") style += ";display:flex;align-items:center";
        else if (vAnchor == "b") style += ";display:flex;align-items:flex-end";

        var lIns = GetLongAttr(bodyPr, "lIns", 91440);
        var tIns = GetLongAttr(bodyPr, "tIns", 45720);
        var rIns = GetLongAttr(bodyPr, "rIns", 91440);
        var bIns = GetLongAttr(bodyPr, "bIns", 45720);
        style += $";padding:{tIns / EmuConverter.EmuPerPx}px {rIns / EmuConverter.EmuPerPx}px {bIns / EmuConverter.EmuPerPx}px {lIns / EmuConverter.EmuPerPx}px";

        // Vertical text direction (bodyPr/@vert): rotate text via CSS writing-mode.
        // OOXML vert values map to writing-mode the same way table-cell tcDir
        // (Css.cs) and Excel textRotation (ExcelHandler.HtmlPreview.cs) do.
        // CONSISTENCY(vertical-text): vertical-rl + text-orientation, see sibling renderers.
        var vert = bodyPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "vert").Value;
        switch (vert)
        {
            case "eaVert":          // East Asian vertical: glyphs upright, columns right→left
            case "mongolianVert":   // rare; degrade to upright vertical
                style += ";writing-mode:vertical-rl;text-orientation:upright";
                break;
            case "vert":            // Latin rotated 90° CW (glyphs lie on their side)
                style += ";writing-mode:vertical-rl";
                break;
            case "vert270":         // Latin rotated 90° CCW
                style += ";writing-mode:vertical-rl;transform:rotate(180deg)";
                break;
            // "horz", null, or unknown → no writing-mode (stay horizontal)
        }

        sb.Append($"<div style=\"{style}\">");

        // #7a: paint the geometry via inline SVG overlay when the preset
        // needs real polygon/path geometry (line, arrows, callouts).
        if (svgPrst)
        {
            var svgFill = ExtractCssColor(fillCss, "background-color")
                ?? ExtractFirstGradientColor(fillCss)
                ?? "transparent";
            var (borderColor, borderWidth) = ExtractBorderParts(borderCss);
            // Connector orientation: flipH/flipV on the shape's a:xfrm decide
            // which box diagonal the stroke runs along. No flip → TL→BR;
            // flipV → BL→TR; flipH → TR→BL; both → BR→TL.
            var geomXfrm = spPr?.Elements().FirstOrDefault(e => e.LocalName == "xfrm");
            bool flipH = IsFlipSet(geomXfrm, "flipH");
            bool flipV = IsFlipSet(geomXfrm, "flipV");
            RenderPrstGeomSvg(sb, prst!, svgFill, borderColor ?? "#000", borderWidth ?? 1, flipH, flipV);
        }

        if (txbx != null)
        {
            // Render text box content (standard Word paragraphs).
            //
            // A shape's <wps:style><a:fontRef> supplies the DEFAULT text color for
            // the text box: a cover-title box filled dark teal commonly carries
            // <a:fontRef><a:schemeClr val="lt1"/> (white) with the title runs
            // themselves carrying NO explicit w:color. Word paints the runs white
            // via the fontRef; emit that color on the content wrapper so runs
            // without an explicit color inherit it (runs WITH a w:color override
            // it via their own inline color:). Without this the title reads black
            // on the dark fill. lt1→white / dk1→black / accentN resolve through
            // the theme via ResolveSchemeColor.
            var fontRefColor = ResolveShapeFontRefColor(shape);
            // overflow-wrap:normal + word-break:normal on the text-box content
            // wrapper to defeat the inherited .page-body{overflow-wrap:break-word}.
            // A fixed-width text box (e.g. a 161px "DRAFT" stamp with leading
            // nbsp for centering) would otherwise let break-word split a single
            // Latin word mid-token ("DRA"/"FT") when nbsp+word exceeds the inner
            // width. Word autofits / keeps the word whole, overflowing slightly.
            // Text-box only — body/table cells keep break-word/anywhere so long
            // content still wraps inside fixed columns.
            // bodyPr/@wrap="none" (e.g. a deliberately narrow classification
            // banner like "IN-CONFIDENCE") tells Word NOT to wrap — it renders a
            // single line that overflows the small box. Honor it with nowrap +
            // overflow:visible so the centered line stays whole instead of
            // breaking at a hyphen ("IN-"/"CONFIDENCE") inside the tiny box.
            var bodyPrEl = shape.Descendants().FirstOrDefault(e => e.LocalName == "bodyPr");
            var noWrap = bodyPrEl != null
                && bodyPrEl.GetAttributes().Any(a => a.LocalName == "wrap" && a.Value == "none");
            var txbxWrap = noWrap
                ? "overflow-wrap:normal;word-break:normal;white-space:nowrap;overflow:visible"
                : "overflow-wrap:normal;word-break:normal";
            var txbxWrapStyle = fontRefColor != null
                ? $"width:100%;color:{fontRefColor};{txbxWrap}"
                : $"width:100%;{txbxWrap}";
            sb.Append($"<div style=\"{txbxWrapStyle}\">");

            // Inject pending float images into this text box
            if (floatImages != null && floatImages.Count > 0)
            {
                foreach (var imgDrawing in floatImages)
                {
                    var imgBlip = imgDrawing.Descendants<A.Blip>().FirstOrDefault();
                    if (imgBlip?.Embed?.Value == null) continue;
                    var imgDataUri = LoadImageAsDataUri(imgBlip.Embed.Value);
                    if (imgDataUri == null) continue;
                    try
                    {
                        var imgExtent = imgDrawing.Descendants<DW.Extent>().FirstOrDefault();
                        var imgW = imgExtent?.Cx?.Value > 0 ? imgExtent.Cx.Value / EmuConverter.EmuPerPx : 100;
                        var imgH = imgExtent?.Cy?.Value > 0 ? imgExtent.Cy.Value / EmuConverter.EmuPerPx : 100;
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
                        var marginCss = $"margin:{distT/EmuConverter.EmuPerPx}px {distR/EmuConverter.EmuPerPx}px {distB/EmuConverter.EmuPerPx}px {distL/EmuConverter.EmuPerPx}px";
                        var crop = GetCropPercents(imgDrawing);
                        if (crop.HasValue)
                        {
                            sb.Append($"<div style=\"float:left;{marginCss}\">");
                            RenderCroppedImage(sb, imgDataUri, imgW, imgH, crop.Value.l, crop.Value.t, crop.Value.r, crop.Value.b, "");
                            sb.Append("</div>");
                        }
                        else
                        {
                            // max-width:none so the overlay's declared px width
                            // wins over the global img{max-width:100%}: a
                            // full-width banner (e.g. 940px) must not be clamped
                            // to the container width and collapse to a strip.
                            sb.Append($"<img src=\"{imgDataUri}\" style=\"float:left;width:{imgW}px;height:{imgH}px;max-width:none;object-fit:cover;{marginCss}\">");
                        }
                    }
                    catch { }
                }
                floatImages = null;
            }

            // Walk txbxContent's direct children — Descendants<Paragraph>()
            // alone would skip <w:tbl> entirely (its row cell paragraphs would
            // surface as bare <p>s, losing the table structure). Mirror the
            // body-render pattern: Paragraph → RenderParagraphHtml,
            // Table → RenderTableHtml, SdtBlock → recurse into content.
            // List grouping inside a text box mirrors the body/cell paths:
            // a run of ListBullet/numbered paragraphs becomes <ul>/<ol> with
            // <li> children instead of bare <p>s. Without this, a bullet list
            // authored inside a DrawingML text box (e.g. a sidebar/cover layout
            // box) lost every marker and collapsed to indented plain paragraphs.
            var txbxOl = new OrderedListNumberingState();
            string? txbxListTag = null;
            RenderTextBoxContentChildren(sb, txbx, ref txbxListTag, txbxOl);
            if (txbxListTag != null) sb.Append($"</{txbxListTag}>");
            sb.Append("</div>");
        }
        else
        {
            // Check for image inside shape
            var embedAttr = FindEmbedInDescendants(shape);
            if (embedAttr != null)
            {
                var dataUri = LoadImageAsDataUri(embedAttr);
                if (dataUri != null)
                    sb.Append($"<img src=\"{dataUri}\" style=\"width:100%;height:100%;object-fit:contain\">");
            }
        }

        sb.Append("</div>");
    }

    /// <summary>
    /// Render the block-level children of a text-box <c>w:txbxContent</c>
    /// (DrawingML <c>wps:txbx</c> or VML <c>v:textbox</c>). Mirrors the
    /// body/header-footer child dispatch: Paragraph → RenderParagraphHtml,
    /// Table → RenderTableHtml, SdtBlock → recurse into the SDT content so
    /// content controls (e.g. placeholder contact-info text inside a sidebar
    /// text box) aren't silently dropped. Block-level SDTs wrap real
    /// paragraphs/tables; iterating only Paragraph/Table here lost every run
    /// nested under a <c>w:sdt</c>.
    /// </summary>
    private void RenderTextBoxContentChildren(StringBuilder sb, OpenXmlElement container, ref string? txbxListTag, OrderedListNumberingState olState)
    {
        foreach (var child in container.ChildElements)
        {
            if (TryEmitContainerBookmarkAnchor(sb, child)) continue;
            if (child is Paragraph para)
            {
                // List item → reuse the cell list renderer (opens <ul>/<ol>,
                // renders the bullet glyph / ordered marker, carries indent).
                var listStyle = GetParagraphListStyle(para);
                if (listStyle != null)
                {
                    RenderCellListItem(sb, para, listStyle, ref txbxListTag, olState);
                    continue;
                }
                // Non-list paragraph closes any open list, then renders flat.
                if (txbxListTag != null) { sb.Append($"</{txbxListTag}>"); txbxListTag = null; }
                RenderParagraphHtml(sb, para);
            }
            else if (child is Table tbl)
            {
                if (txbxListTag != null) { sb.Append($"</{txbxListTag}>"); txbxListTag = null; }
                RenderTableHtml(sb, tbl);
            }
            else if (child is SdtBlock sdt && sdt.SdtContentBlock is { } content)
                RenderTextBoxContentChildren(sb, content, ref txbxListTag, olState);
        }
    }

    /// <summary>
    /// Resolve the default text color a DrawingML shape's
    /// <c>&lt;wps:style&gt;&lt;a:fontRef&gt;</c> contributes to its text box.
    /// fontRef's color is either an <c>&lt;a:schemeClr&gt;</c> (lt1/dk1/accentN —
    /// resolved through the theme, lt1 → white, dk1 → black) or a literal
    /// <c>&lt;a:srgbClr&gt;</c>. Returns a CSS color string or null when the shape
    /// has no style/fontRef or the color can't be resolved.
    /// </summary>
    private string? ResolveShapeFontRefColor(OpenXmlElement shape)
    {
        var style = shape.Elements().FirstOrDefault(e => e.LocalName == "style");
        var fontRef = style?.Elements().FirstOrDefault(e => e.LocalName == "fontRef");
        if (fontRef == null) return null;

        var scheme = fontRef.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
        if (scheme != null)
            return ResolveSchemeColor(scheme);

        var rgb = fontRef.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var val = rgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (val != null && IsHexColor(val))
            return $"#{val}";

        return null;
    }

    // ==================== #7a prstGeom SVG helpers ====================

    /// <summary>
    /// Pull a CSS property's color value out of strings like
    /// <c>background-color:#FF0000</c> or
    /// <c>background:linear-gradient(...)</c>. Returns null if not present.
    /// </summary>
    private static string? ExtractCssColor(string css, string prop)
    {
        if (string.IsNullOrEmpty(css)) return null;
        var m = System.Text.RegularExpressions.Regex.Match(
            css, $@"{prop}\s*:\s*(#[0-9A-Fa-f]{{3,8}}|[a-zA-Z]+)");
        return m.Success ? m.Groups[1].Value : null;
    }

    // Pull the first hex color out of a `background:linear-gradient(...)`
    // / `background-image:linear-gradient(...)` rule so SVG prstGeom shapes
    // don't degrade to transparent when only a gradient fill is available.
    private static string? ExtractFirstGradientColor(string css)
    {
        if (string.IsNullOrEmpty(css)) return null;
        if (css.IndexOf("gradient", StringComparison.OrdinalIgnoreCase) < 0) return null;
        var m = System.Text.RegularExpressions.Regex.Match(
            css, @"#[0-9A-Fa-f]{3,8}");
        return m.Success ? m.Value : null;
    }

    private static (string? color, double? width) ExtractBorderParts(string css)
    {
        if (string.IsNullOrEmpty(css)) return (null, null);
        // e.g. "border:1.5px solid #336699"
        var m = System.Text.RegularExpressions.Regex.Match(
            css, @"border\s*:\s*([\d.]+)px\s+\w+\s+(#[0-9A-Fa-f]{3,8}|[a-zA-Z]+)");
        if (!m.Success) return (null, null);
        return (m.Groups[2].Value,
            double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var w) ? w : 1);
    }

    /// <summary>
    /// Emit an inline SVG overlay rendering the given preset geometry.
    /// The SVG uses viewBox="0 0 100 100" and preserveAspectRatio="none"
    /// so it stretches to the host div's full size.
    /// </summary>
    /// <summary>Read a flipH/flipV boolean off an a:xfrm element.</summary>
    private static bool IsFlipSet(OpenXmlElement? xfrm, string name)
    {
        var v = xfrm?.GetAttributes().FirstOrDefault(a => a.LocalName == name).Value;
        return v == "1" || v == "true";
    }

    private static void RenderPrstGeomSvg(
        StringBuilder sb, string prst, string fill, string stroke, double strokeW,
        bool flipH = false, bool flipV = false)
    {
        // Normalize stroke width to viewBox coordinates: at 100-unit viewBox
        // and typical host size ~150px, 1px ≈ 0.67 units. Keep as-is since
        // preserveAspectRatio=none scales X/Y differently anyway; ok for
        // approximation.
        // Display:block + width/height:100% makes the SVG fill the host
        // <div> without needing position:absolute (which would anchor to
        // the nearest positioned ancestor and cause all shapes on a page
        // to stack on top of each other).
        sb.Append(
            "<svg style=\"display:block;width:100%;height:100%;overflow:visible\" " +
            "viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\" xmlns=\"http://www.w3.org/2000/svg\">");
        var sw = strokeW.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        switch (prst)
        {
            case "line":
            case "straightConnector1":
                // The stroke runs along a box diagonal; flipH/flipV pick which
                // one. Within the connector's wide/short bounding box this
                // diagonal renders as the true near-horizontal (or near-vertical)
                // line. No flip → TL→BR; flipV only → BL→TR; flipH only → TR→BL;
                // both → BR→TL.
                int x1 = flipH ? 100 : 0;
                int y1 = flipV ? 100 : 0;
                int x2 = flipH ? 0 : 100;
                int y2 = flipV ? 0 : 100;
                sb.Append($"<line x1=\"{x1}\" y1=\"{y1}\" x2=\"{x2}\" y2=\"{y2}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "rightArrow":
                // Classic block arrow pointing right: body 0..70, head 70..100.
                sb.Append($"<polygon points=\"0,30 70,30 70,10 100,50 70,90 70,70 0,70\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "leftArrow":
                sb.Append($"<polygon points=\"100,30 30,30 30,10 0,50 30,90 30,70 100,70\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "downArrow":
                sb.Append($"<polygon points=\"30,0 70,0 70,70 90,70 50,100 10,70 30,70\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "upArrow":
                sb.Append($"<polygon points=\"30,100 70,100 70,30 90,30 50,0 10,30 30,30\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "diamond":
            case "flowChartDecision":
                // Flowchart decision node / rhombus: vertices at the midpoint of
                // each box edge (top, right, bottom, left). flowChartDecision is
                // the same rhombus geometry as diamond.
                sb.Append($"<polygon points=\"50,0 100,50 50,100 0,50\" fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
            case "wedgeRoundRectCallout":
                // Rounded rect (80% height) + triangular pointer down-left.
                // Rect corners rounded at 10 units; pointer tip at (15, 95).
                sb.Append($"<path d=\"M 10,0 L 90,0 Q 100,0 100,10 L 100,70 Q 100,80 90,80 L 45,80 L 15,95 L 30,80 L 10,80 Q 0,80 0,70 L 0,10 Q 0,0 10,0 Z\" " +
                          $"fill=\"{fill}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"/>");
                break;
        }
        sb.Append("</svg>");
    }

}
