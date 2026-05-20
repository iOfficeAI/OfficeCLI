// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string AddPicture(string parentPath, int? index, Dictionary<string, string> properties)
    {
                if (!properties.TryGetValue("path", out var imgPath)
                    && !properties.TryGetValue("src", out imgPath))
                    throw new ArgumentException("'src' property is required for picture type");

                var imgSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!imgSlideMatch.Success)
                    throw new ArgumentException($"Pictures must be added to a slide: /slide[N]");

                var imgSlideIdx = int.Parse(imgSlideMatch.Groups[1].Value);
                var imgSlideParts = GetSlideParts().ToList();
                if (imgSlideIdx < 1 || imgSlideIdx > imgSlideParts.Count)
                    throw new ArgumentException($"Slide {imgSlideIdx} not found (total: {imgSlideParts.Count})");

                var imgSlidePart = imgSlideParts[imgSlideIdx - 1];
                var imgShapeTree = GetSlide(imgSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                // Resolve image from file/base64/URL and buffer for
                // both embedding and dimension sniffing (aspect ratio).
                var (rawImgStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
                using var rawImgDispose = rawImgStream;
                using var imgStream = new MemoryStream();
                rawImgStream.CopyTo(imgStream);
                imgStream.Position = 0;

                // Embed image into slide part. For SVG, emit the dual
                // representation Office requires: PNG fallback at r:embed,
                // SVG referenced via a:blip/a:extLst asvg:svgBlip.
                string imgRelId;
                string? picSvgRelId = null;
                if (imgPartType == ImagePartType.Svg)
                {
                    var svgPart = imgSlidePart.AddImagePart(ImagePartType.Svg);
                    svgPart.FeedData(imgStream);
                    imgStream.Position = 0;
                    picSvgRelId = imgSlidePart.GetIdOfPart(svgPart);

                    if (properties.TryGetValue("fallback", out var picFallback) && !string.IsNullOrWhiteSpace(picFallback))
                    {
                        var (fbRaw, fbType) = OfficeCli.Core.ImageSource.Resolve(picFallback);
                        using var fbDispose = fbRaw;
                        var fbPart = imgSlidePart.AddImagePart(fbType);
                        fbPart.FeedData(fbRaw);
                        imgRelId = imgSlidePart.GetIdOfPart(fbPart);
                    }
                    else
                    {
                        var pngPart = imgSlidePart.AddImagePart(ImagePartType.Png);
                        pngPart.FeedData(new MemoryStream(
                            OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                        imgRelId = imgSlidePart.GetIdOfPart(pngPart);
                    }
                }
                else
                {
                    var imagePart = imgSlidePart.AddImagePart(imgPartType);
                    imagePart.FeedData(imgStream);
                    imgStream.Position = 0;
                    imgRelId = imgSlidePart.GetIdOfPart(imagePart);
                }

                // Dimensions (default: 6in x 4in, with auto aspect-ratio)
                // CONSISTENCY(picture-aspect): when only one dimension is
                // supplied, compute the other from native pixel ratio — same
                // behavior as WordHandler.AddPicture.
                bool hasWidth = properties.TryGetValue("width", out var widthStr);
                bool hasHeight = properties.TryGetValue("height", out var heightStr);
                long cxEmu = hasWidth ? ParseEmu(widthStr!) : 5486400;  // 6 inches fallback
                long cyEmu = hasHeight ? ParseEmu(heightStr!) : 3657600; // 4 inches fallback
                // CONSISTENCY(positive-size): symmetric with Add.Shape negative-size guard
                // so picture / chart / connector / media all reject inverted dimensions.
                if (cxEmu < 0) throw new ArgumentException($"Negative width is not allowed: '{widthStr}'.");
                if (cyEmu < 0) throw new ArgumentException($"Negative height is not allowed: '{heightStr}'.");

                if (!hasWidth || !hasHeight)
                {
                    var dims = OfficeCli.Core.ImageSource.TryGetDimensions(imgStream);
                    if (dims is { Width: > 0, Height: > 0 } d)
                    {
                        double ratio = (double)d.Height / d.Width;
                        if (hasWidth && !hasHeight)
                            cyEmu = (long)(cxEmu * ratio);
                        else if (!hasWidth && hasHeight)
                            cxEmu = (long)(cyEmu / ratio);
                        else // neither supplied — default width, compute height
                            cyEmu = (long)(cxEmu * ratio);
                    }
                }

                // Position (default: centered on slide)
                var (slideW, slideH) = GetSlideSize();
                long xEmu = (slideW - cxEmu) / 2;
                long yEmu = (slideH - cyEmu) / 2;
                if (properties.TryGetValue("x", out var xStr) || properties.TryGetValue("left", out xStr))
                    xEmu = ParseEmu(xStr);
                if (properties.TryGetValue("y", out var yStr) || properties.TryGetValue("top", out yStr))
                    yEmu = ParseEmu(yStr);

                var imgShapeId = AcquireShapeId(imgShapeTree, properties);
                var imgName = properties.GetValueOrDefault("name", $"Picture {imgShapeTree.Elements<Picture>().Count() + 1}");
                // BUG-R5-02: data URIs / raw base64 blobs make Path.GetFileName
                // return a meaningless tail (e.g. "png;base64,iVBOR..."). Use a
                // placeholder unless the caller supplied an explicit alt=.
                string DefaultPictureAlt()
                {
                    if (string.IsNullOrEmpty(imgPath)) return imgName;
                    if (imgPath.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return imgName;
                    if (imgPath.Length > 256 && imgPath.IndexOf('/') < 0 && imgPath.IndexOf('\\') < 0) return imgName;
                    try { return Path.GetFileName(imgPath); } catch { return imgName; }
                }
                var altText = properties.TryGetValue("alt", out var altOverride) && !string.IsNullOrEmpty(altOverride)
                    ? altOverride
                    : DefaultPictureAlt();

                // Build Picture element following Open-XML-SDK conventions
                var picture = new Picture();

                picture.NonVisualPictureProperties = new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = imgShapeId, Name = imgName, Description = altText },
                    new NonVisualPictureDrawingProperties(
                        new Drawing.PictureLocks { NoChangeAspect = true }
                    ),
                    new ApplicationNonVisualDrawingProperties()
                );

                picture.BlipFill = new BlipFill();
                picture.BlipFill.Blip = new Drawing.Blip { Embed = imgRelId };
                if (picSvgRelId != null)
                    OfficeCli.Core.SvgImageHelper.AppendSvgExtension(picture.BlipFill.Blip, picSvgRelId);

                // Crop support (mirrors Set's crop emitter — keep keys/semantics
                // identical per CLAUDE.md Feature Implementation Checklist).
                // CONSISTENCY(ooxml-element-order): in CT_BlipFillProperties
                // srcRect must precede the fill-mode element (stretch/tile);
                // PowerPoint silently ignores an out-of-order srcRect.
                int? cropL = null, cropT = null, cropR = null, cropB = null;
                if (properties.TryGetValue("crop", out var cropAll))
                {
                    var parts = cropAll.Split(',');
                    double Parse1(string s)
                    {
                        // R10: accept trailing '%' suffix on each comma-separated value.
                        var stripped = s.Trim();
                        if (stripped.EndsWith("%", StringComparison.Ordinal)) stripped = stripped[..^1].Trim();
                        var v = ParseHelpers.SafeParseDouble(stripped, "crop");
                        if (v < 0 || v > 100)
                            throw new ArgumentException($"Invalid 'crop' value: '{s.Trim()}'. Crop percentage must be between 0 and 100.");
                        return v;
                    }
                    if (parts.Length == 4)
                    {
                        cropL = (int)(Parse1(parts[0]) * 1000);
                        cropT = (int)(Parse1(parts[1]) * 1000);
                        cropR = (int)(Parse1(parts[2]) * 1000);
                        cropB = (int)(Parse1(parts[3]) * 1000);
                    }
                    else if (parts.Length == 2)
                    {
                        var v = (int)(Parse1(parts[0]) * 1000);
                        var h = (int)(Parse1(parts[1]) * 1000);
                        cropT = v; cropB = v; cropL = h; cropR = h;
                    }
                    else if (parts.Length == 1)
                    {
                        var p = (int)(Parse1(parts[0]) * 1000);
                        cropL = p; cropT = p; cropR = p; cropB = p;
                    }
                    else
                    {
                        throw new ArgumentException($"Invalid 'crop' value: '{cropAll}'. Expected 1, 2, or 4 comma-separated percentages.");
                    }
                }
                int? SidePct(string k)
                {
                    if (!properties.TryGetValue(k, out var v)) return null;
                    // R10: accept trailing '%' suffix — error message already says
                    // "Expected a percentage (0-100)", so the % literal is the
                    // natural input form and rejecting it was self-contradictory.
                    var stripped = v.Trim();
                    if (stripped.EndsWith("%", StringComparison.Ordinal)) stripped = stripped[..^1].Trim();
                    if (!double.TryParse(stripped, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var d))
                        throw new ArgumentException($"Invalid '{k}' value: '{v}'. Expected a percentage (0-100).");
                    if (d < 0 || d > 100)
                        throw new ArgumentException($"Invalid '{k}' value: '{v}'. Crop percentage must be between 0 and 100.");
                    return (int)(d * 1000);
                }
                cropL = SidePct("cropleft") ?? cropL;
                cropT = SidePct("croptop") ?? cropT;
                cropR = SidePct("cropright") ?? cropR;
                cropB = SidePct("cropbottom") ?? cropB;
                var hasCrop = cropL is not null || cropT is not null || cropR is not null || cropB is not null;
                var anyNonZero = (cropL ?? 0) != 0 || (cropT ?? 0) != 0 || (cropR ?? 0) != 0 || (cropB ?? 0) != 0;
                if (hasCrop && anyNonZero)
                {
                    var srcRect = new Drawing.SourceRectangle();
                    if (cropL is not null) srcRect.Left = cropL;
                    if (cropT is not null) srcRect.Top = cropT;
                    if (cropR is not null) srcRect.Right = cropR;
                    if (cropB is not null) srcRect.Bottom = cropB;
                    picture.BlipFill.AppendChild(srcRect); // stretch not yet appended
                }
                // Fill mode: stretch (default) | contain (letterbox) |
                // cover (crop) | tile. stretch preserves the historical
                // <a:stretch><a:fillRect/></a:stretch> emission so existing
                // docs stay byte-identical. contain/cover require image and
                // container dimensions; if either is unknown, we fall back
                // to a bare stretch.
                var fillMode = (properties.GetValueOrDefault("fill", "stretch") ?? "stretch")
                    .Trim().ToLowerInvariant();
                if (fillMode == "tile")
                {
                    double tileScale = 1.0;
                    if (properties.TryGetValue("tilescale", out var tsStr)
                        && double.TryParse(tsStr, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var ts) && ts > 0)
                        tileScale = ts;
                    var tile = new Drawing.Tile
                    {
                        HorizontalRatio = (int)(tileScale * 100000),
                        VerticalRatio = (int)(tileScale * 100000),
                        Flip = Drawing.TileFlipValues.None,
                        Alignment = Drawing.RectangleAlignmentValues.TopLeft,
                    };
                    if (properties.TryGetValue("tilealign", out var taStr))
                    {
                        tile.Alignment = taStr.Trim().ToLowerInvariant() switch
                        {
                            "tl" or "topleft" => Drawing.RectangleAlignmentValues.TopLeft,
                            "t" or "top" => Drawing.RectangleAlignmentValues.Top,
                            "tr" or "topright" => Drawing.RectangleAlignmentValues.TopRight,
                            "l" or "left" => Drawing.RectangleAlignmentValues.Left,
                            "ctr" or "center" or "centre" => Drawing.RectangleAlignmentValues.Center,
                            "r" or "right" => Drawing.RectangleAlignmentValues.Right,
                            "bl" or "bottomleft" => Drawing.RectangleAlignmentValues.BottomLeft,
                            "b" or "bottom" => Drawing.RectangleAlignmentValues.Bottom,
                            "br" or "bottomright" => Drawing.RectangleAlignmentValues.BottomRight,
                            _ => Drawing.RectangleAlignmentValues.TopLeft,
                        };
                    }
                    if (properties.TryGetValue("tileflip", out var tfStr))
                    {
                        tile.Flip = tfStr.Trim().ToLowerInvariant() switch
                        {
                            "none" => Drawing.TileFlipValues.None,
                            "x" => Drawing.TileFlipValues.Horizontal,
                            "y" => Drawing.TileFlipValues.Vertical,
                            "xy" or "both" => Drawing.TileFlipValues.HorizontalAndVertical,
                            _ => Drawing.TileFlipValues.None,
                        };
                    }
                    picture.BlipFill.AppendChild(tile);
                }
                else if (fillMode == "contain" || fillMode == "cover")
                {
                    // Compute native-vs-container aspect to derive fillRect
                    // offsets. a:fillRect insets are in thousandths of a
                    // percent (100000 = 100%). Positive insets shrink the
                    // stretched area (letterbox for contain), negatives
                    // enlarge it (crop for cover).
                    imgStream.Position = 0;
                    var dims = OfficeCli.Core.ImageSource.TryGetDimensions(imgStream);
                    if (dims is { Width: > 0, Height: > 0 } d2 && cxEmu > 0 && cyEmu > 0)
                    {
                        double imgAspect = (double)d2.Width / d2.Height;
                        double boxAspect = (double)cxEmu / cyEmu;
                        var fr = new Drawing.FillRectangle();
                        if (fillMode == "contain")
                        {
                            if (imgAspect > boxAspect)
                            {
                                // Image wider than box — pad top/bottom
                                var pad = (int)Math.Round(((1.0 - boxAspect / imgAspect) / 2.0) * 100000);
                                fr.Top = pad; fr.Bottom = pad;
                            }
                            else
                            {
                                var pad = (int)Math.Round(((1.0 - imgAspect / boxAspect) / 2.0) * 100000);
                                fr.Left = pad; fr.Right = pad;
                            }
                        }
                        else // cover
                        {
                            if (imgAspect > boxAspect)
                            {
                                // Image wider than box — crop left/right (negative inset)
                                var crop = (int)Math.Round(((imgAspect / boxAspect - 1.0) / 2.0) * 100000);
                                fr.Left = -crop; fr.Right = -crop;
                            }
                            else
                            {
                                var crop = (int)Math.Round(((boxAspect / imgAspect - 1.0) / 2.0) * 100000);
                                fr.Top = -crop; fr.Bottom = -crop;
                            }
                        }
                        picture.BlipFill.AppendChild(new Drawing.Stretch(fr));
                    }
                    else
                    {
                        picture.BlipFill.AppendChild(new Drawing.Stretch(new Drawing.FillRectangle()));
                    }
                }
                else
                {
                    picture.BlipFill.AppendChild(new Drawing.Stretch(new Drawing.FillRectangle()));
                }

                picture.ShapeProperties = new ShapeProperties();
                picture.ShapeProperties.Transform2D = new Drawing.Transform2D();
                picture.ShapeProperties.Transform2D.Offset = new Drawing.Offset { X = xEmu, Y = yEmu };
                picture.ShapeProperties.Transform2D.Extents = new Drawing.Extents { Cx = cxEmu, Cy = cyEmu };
                var picGeomName = "rect";
                if (properties.TryGetValue("geometry", out var picGeom) || properties.TryGetValue("shape", out picGeom))
                    picGeomName = picGeom;
                picture.ShapeProperties.AppendChild(
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = ParsePresetShape(picGeomName) }
                );

                // CONSISTENCY(shape-picture-parity): rotation lives on the
                // same Transform2D as shape/connector/group; PowerPoint
                // applies it identically to pictures. Set.Media already
                // accepts rotation on picture; Add must mirror so Add and
                // Set agree on the property surface.
                if (properties.TryGetValue("rotation", out var picRotStr)
                    || properties.TryGetValue("rotate", out picRotStr))
                {
                    picture.ShapeProperties.Transform2D!.Rotation =
                        (int)(ParseHelpers.SafeParseDouble(picRotStr, "rotation") * 60000);
                }

                // CONSISTENCY(shape-picture-parity): pictures are routinely
                // click-targets — wire link= the same way shape does.
                // Tooltip is the same secondary key as on shape.
                if (properties.TryGetValue("link", out var picLink))
                {
                    var picTip = properties.GetValueOrDefault("tooltip");
                    ApplyPictureHyperlink(imgSlidePart, picture, picLink, picTip);
                }

                InsertAtPosition(imgShapeTree, picture, index);

                // CONSISTENCY(zorder-on-add): dump-emit carries zorder=N; without
                // consuming it here every picture appended at the end of the tree.
                if (properties.TryGetValue("zorder", out var picZ)
                    || properties.TryGetValue("z-order", out picZ)
                    || properties.TryGetValue("order", out picZ))
                {
                    ApplyZOrder(imgSlidePart, picture, picZ);
                }

                GetSlide(imgSlidePart).Save();

                return $"/slide[{imgSlideIdx}]/{BuildElementPathSegment("picture", picture, imgShapeTree.Elements<Picture>().Count())}";
    }


    private string AddChart(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var chartSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!chartSlideMatch.Success)
                    throw new ArgumentException("Charts must be added to a slide: /slide[N]");

                var chartSlideIdx = int.Parse(chartSlideMatch.Groups[1].Value);
                var chartSlideParts = GetSlideParts().ToList();
                if (chartSlideIdx < 1 || chartSlideIdx > chartSlideParts.Count)
                    throw new ArgumentException($"Slide {chartSlideIdx} not found (total: {chartSlideParts.Count})");

                var chartSlidePart = chartSlideParts[chartSlideIdx - 1];
                var chartShapeTree = GetSlide(chartSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                // Parse chart data. Use TryGetValue(case-insensitive) instead
                // of LINQ FirstOrDefault to play well with TrackingPropertyDictionary.
                string chartType = "column";
                if (properties.TryGetValue("charttype", out var ct) || properties.TryGetValue("type", out ct))
                    chartType = ct;
                var chartTitle = properties.GetValueOrDefault("title");
                var categories = ChartHelper.ParseCategories(properties);
                var seriesData = ChartHelper.ParseSeriesData(properties);

                if (seriesData.Count == 0)
                    throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                        "or series1=\"Revenue:100,200,300\"");

                // Position. Accept `anchor=x,y` (2-arg, position only — size
                // falls back to defaults) or `anchor=x,y,w,h` (full rect).
                // Schema documents both 2- and 4-arg forms; explicit
                // `x=`/`y=`/`width=`/`height=` props still override anchor
                // for that axis (matches OLE/picture/shape add convention).
                long chartX = properties.TryGetValue("x", out var xv) ? ParseEmu(xv) : 838200;     // ~2.3cm
                long chartY = properties.TryGetValue("y", out var yv) ? ParseEmu(yv) : 1825625;     // ~5cm
                long chartCx = properties.TryGetValue("width", out var wv) ? ParseEmu(wv) : 8229600; // ~22.9cm
                long chartCy = properties.TryGetValue("height", out var hv) ? ParseEmu(hv) : 4572000; // ~12.7cm
                if (properties.TryGetValue("anchor", out var chartAnchorRaw) && !string.IsNullOrWhiteSpace(chartAnchorRaw))
                {
                    var anchorParts = chartAnchorRaw.Split(',', StringSplitOptions.TrimEntries);
                    if (anchorParts.Length != 2 && anchorParts.Length != 4)
                        throw new ArgumentException(
                            $"Invalid anchor: '{chartAnchorRaw}'. Expected 'x,y' (e.g. '2cm,3cm') or 'x,y,w,h' (e.g. '2cm,3cm,18cm,10cm'). " +
                            "Cell-range form (e.g. 'D2:J18') is xlsx-only.");
                    if (!properties.ContainsKey("x")) chartX = ParseEmu(anchorParts[0]);
                    if (!properties.ContainsKey("y")) chartY = ParseEmu(anchorParts[1]);
                    if (anchorParts.Length == 4)
                    {
                        if (!properties.ContainsKey("width")) chartCx = ParseEmu(anchorParts[2]);
                        if (!properties.ContainsKey("height")) chartCy = ParseEmu(anchorParts[3]);
                    }
                }
                // CONSISTENCY(positive-size): symmetric with Add.Shape negative-size guard.
                if (chartCx < 0) throw new ArgumentException($"Negative width is not allowed: '{wv}'.");
                if (chartCy < 0) throw new ArgumentException($"Negative height is not allowed: '{hv}'.");
                var chartId = AcquireShapeId(chartShapeTree, properties);
                var chartName = properties.GetValueOrDefault("name", chartTitle ?? $"Chart {chartShapeTree.Elements<GraphicFrame>().Count(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any() || IsExtendedChartFrame(gf)) + 1}");

                // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
                if (ChartExBuilder.IsExtendedChartType(chartType))
                {
                    var cxChartSpace = ChartExBuilder.BuildExtendedChartSpace(
                        chartType, chartTitle, categories, seriesData, properties);
                    var extChartPart = chartSlidePart.AddNewPart<ExtendedChartPart>();
                    extChartPart.ChartSpace = cxChartSpace;
                    extChartPart.ChartSpace.Save();

                    // CONSISTENCY(chartex-sidecars): every chartEx part needs
                    // three sibling parts wired via specific relationship IDs:
                    //   rId1 → embedded .xlsx (cx:externalData target)
                    //   rId2 → chartStyle.xml
                    //   rId3 → colors.xml
                    // PowerPoint silently repairs (drops the chart, sometimes
                    // the entire shape group) if any of these are missing.
                    var embPart = extChartPart.AddNewPart<EmbeddedPackagePart>(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId1");
                    var xlsxBytes = ChartExResources.BuildMinimalEmbeddedXlsx(categories, seriesData);
                    using (var emsr = new MemoryStream(xlsxBytes))
                        embPart.FeedData(emsr);

                    var stylePart = extChartPart.AddNewPart<ChartStylePart>("rId2");
                    using (var styleStream = ChartExResources.OpenChartStyleXml())
                        stylePart.FeedData(styleStream);

                    var colorPart = extChartPart.AddNewPart<ChartColorStylePart>("rId3");
                    using (var colorStream = ChartExResources.OpenChartColorStyleXml())
                        colorPart.FeedData(colorStream);

                    var chartGfEx = BuildExtendedChartGraphicFrame(chartSlidePart, extChartPart,
                        chartId, chartName, chartX, chartY, chartCx, chartCy);
                    InsertAtPosition(chartShapeTree, chartGfEx, index);
                    if (properties.TryGetValue("zorder", out var cxZ)
                        || properties.TryGetValue("z-order", out cxZ)
                        || properties.TryGetValue("order", out cxZ))
                        ApplyZOrder(chartSlidePart, chartGfEx, cxZ);
                    GetSlide(chartSlidePart).Save();

                    // Count all charts (both regular and extended)
                    var totalCharts = chartShapeTree.Elements<GraphicFrame>()
                        .Count(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf));
                    return $"/slide[{chartSlideIdx}]/{BuildElementPathSegment("chart", chartGfEx, totalCharts)}";
                }

                // Build chart content BEFORE adding part (invalid type throws, must not leave empty part)
                var chartSpace = ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
                var chartPart = chartSlidePart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = chartSpace;
                chartPart.ChartSpace.Save();

                // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
                var deferredProps = properties
                    .Where(kv => ChartHelper.IsDeferredKey(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (deferredProps.Count > 0)
                    ChartHelper.SetChartProperties(chartPart, deferredProps);

                var chartGf = BuildChartGraphicFrame(chartSlidePart, chartPart, chartId, chartName,
                    chartX, chartY, chartCx, chartCy);
                InsertAtPosition(chartShapeTree, chartGf, index);
                if (properties.TryGetValue("zorder", out var stdZ)
                    || properties.TryGetValue("z-order", out stdZ)
                    || properties.TryGetValue("order", out stdZ))
                    ApplyZOrder(chartSlidePart, chartGf, stdZ);
                GetSlide(chartSlidePart).Save();

                var chartCount = chartShapeTree.Elements<GraphicFrame>()
                    .Count(gf => gf.Descendants<C.ChartReference>().Any());
                return $"/slide[{chartSlideIdx}]/{BuildElementPathSegment("chart", chartGf, chartCount)}";
    }


    private string AddMedia(string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
                var mediaSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!mediaSlideMatch.Success)
                    throw new ArgumentException("Media must be added to a slide: /slide[N]");

                if (!properties.TryGetValue("path", out var mediaPath)
                    && !properties.TryGetValue("src", out mediaPath))
                    throw new ArgumentException("'src' property is required for media type");

                var (mediaStream, ext) = OfficeCli.Core.FileSource.Resolve(mediaPath);
                using var mediaStreamDispose = mediaStream;

                var mediaSlideIdx = int.Parse(mediaSlideMatch.Groups[1].Value);
                var mediaSlideParts = GetSlideParts().ToList();
                if (mediaSlideIdx < 1 || mediaSlideIdx > mediaSlideParts.Count)
                    throw new ArgumentException($"Slide {mediaSlideIdx} not found (total: {mediaSlideParts.Count})");

                var mediaSlidePart = mediaSlideParts[mediaSlideIdx - 1];
                var mediaShapeTree = GetSlide(mediaSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var isVideo = type.ToLowerInvariant() == "video" ||
                    (type.ToLowerInvariant() == "media" && ext is ".mp4" or ".avi" or ".wmv" or ".mpg" or ".mov");

                var contentType = ext switch
                {
                    ".mp4" => "video/mp4", ".avi" => "video/x-msvideo", ".wmv" => "video/x-ms-wmv",
                    ".mpg" or ".mpeg" => "video/mpeg", ".mov" => "video/quicktime",
                    ".mp3" => "audio/mpeg", ".wav" => "audio/wav", ".wma" => "audio/x-ms-wma",
                    ".m4a" => "audio/mp4", _ => isVideo ? "video/mp4" : "audio/mpeg"
                };

                // 1. Create MediaDataPart and feed binary data
                var mediaDataPart = _doc.CreateMediaDataPart(contentType, ext);
                mediaDataPart.FeedData(mediaStream);

                // 2. Create relationships: Video/Audio + Media
                string videoRelId, mediaRelId;
                if (isVideo)
                {
                    videoRelId = mediaSlidePart.AddVideoReferenceRelationship(mediaDataPart).Id;
                    mediaRelId = mediaSlidePart.AddMediaReferenceRelationship(mediaDataPart).Id;
                }
                else
                {
                    videoRelId = mediaSlidePart.AddAudioReferenceRelationship(mediaDataPart).Id;
                    mediaRelId = mediaSlidePart.AddMediaReferenceRelationship(mediaDataPart).Id;
                }

                // 3. Add poster/thumbnail image
                ImagePart posterPart;
                if (properties.TryGetValue("poster", out var posterPath))
                {
                    var (posterStream, posterType) = OfficeCli.Core.ImageSource.Resolve(posterPath);
                    using var posterDispose = posterStream;
                    posterPart = mediaSlidePart.AddImagePart(posterType);
                    posterPart.FeedData(posterStream);
                }
                else
                {
                    // Minimal 1x1 transparent PNG placeholder
                    posterPart = mediaSlidePart.AddImagePart(ImagePartType.Png);
                    var posterPng = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
                        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
                        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
                        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
                    };
                    using var ms = new MemoryStream(posterPng);
                    posterPart.FeedData(ms);
                }
                var posterRelId = mediaSlidePart.GetIdOfPart(posterPart);

                // Position
                var (mediaSlideW, mediaSlideH) = GetSlideSize();
                long mCx = properties.TryGetValue("width", out var mwv) ? ParseEmu(mwv) : (long)(mediaSlideW * 0.75);
                long mCy = properties.TryGetValue("height", out var mhv) ? ParseEmu(mhv) : (long)(mediaSlideH * 0.75);
                // CONSISTENCY(positive-size): symmetric with Add.Shape negative-size guard.
                if (mCx < 0) throw new ArgumentException($"Negative width is not allowed: '{mwv}'.");
                if (mCy < 0) throw new ArgumentException($"Negative height is not allowed: '{mhv}'.");
                long mX = properties.TryGetValue("x", out var mxv) ? ParseEmu(mxv) : (mediaSlideW - mCx) / 2;
                long mY = properties.TryGetValue("y", out var myv) ? ParseEmu(myv) : (mediaSlideH - mCy) / 2;

                var mediaId = AcquireShapeId(mediaShapeTree, properties);
                var mediaName = properties.GetValueOrDefault("name", isVideo ? "video" : "audio");

                // 4. Build Picture element with proper video/audio structure
                // cNvPr with hlinkClick action="ppaction://media"
                var cNvPr = new NonVisualDrawingProperties { Id = mediaId, Name = mediaName };
                cNvPr.AppendChild(new Drawing.HyperlinkOnClick { Id = "", Action = "ppaction://media" });

                // nvPr with VideoFromFile/AudioFromFile + p14:media extension
                var appNvPr = new ApplicationNonVisualDrawingProperties();
                if (isVideo)
                    appNvPr.AppendChild(new Drawing.VideoFromFile { Link = videoRelId });
                else
                    appNvPr.AppendChild(new Drawing.AudioFromFile { Link = videoRelId });

                // p14:media extension (PowerPoint 2010+)
                var p14Media = new DocumentFormat.OpenXml.Office2010.PowerPoint.Media { Embed = mediaRelId };
                p14Media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

                var extList = new ApplicationNonVisualDrawingPropertiesExtensionList();
                var appExt = new ApplicationNonVisualDrawingPropertiesExtension
                    { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };
                appExt.AppendChild(p14Media);
                extList.AppendChild(appExt);
                appNvPr.AppendChild(extList);

                var mediaPic = new Picture();
                mediaPic.NonVisualPictureProperties = new NonVisualPictureProperties(
                    cNvPr,
                    new NonVisualPictureDrawingProperties(new Drawing.PictureLocks { NoChangeAspect = true }),
                    appNvPr
                );
                mediaPic.BlipFill = new BlipFill(
                    new Drawing.Blip { Embed = posterRelId },
                    new Drawing.Stretch(new Drawing.FillRectangle())
                );
                mediaPic.ShapeProperties = new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = mX, Y = mY },
                        new Drawing.Extents { Cx = mCx, Cy = mCy }
                    ),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                );

                // p14:trim (optional start/end trim). Schema accepts both
                // "hh:mm:ss.fff" timestamps and bare millisecond counts on
                // input; normalize to ms-int on write — PowerPoint refuses
                // to open a file whose p14:trim/@st is a timestamp literal
                // (0x80070570 "file corrupted").
                properties.TryGetValue("trimstart", out var trimStart);
                properties.TryGetValue("trimend", out var trimEnd);
                if (trimStart != null || trimEnd != null)
                {
                    var trim = new DocumentFormat.OpenXml.Office2010.PowerPoint.MediaTrim();
                    if (trimStart != null) trim.Start = NormalizeMediaTimeMs(trimStart, "trimStart");
                    if (trimEnd != null) trim.End = NormalizeMediaTimeMs(trimEnd, "trimEnd");
                    p14Media.MediaTrim = trim;
                }

                InsertAtPosition(mediaShapeTree, mediaPic, index);

                // 5. Add media timing node (controls playback behavior)
                var mediaSlide = GetSlide(mediaSlidePart);
                var vol = 80000; // default 80%
                if (properties.TryGetValue("volume", out var volStr))
                {
                    if (!double.TryParse(volStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var volDbl))
                        throw new ArgumentException($"Invalid 'volume' value: '{volStr}'. Expected a number 0-100 (e.g. 80 = 80%).");
                    // Detect 0-1 range input (e.g. 0.8 meaning 80%) and normalize to 0-100
                    if (volDbl > 0 && volDbl <= 1.0) volDbl *= 100;
                    vol = (int)(volDbl * 1000); // 0-100 → 0-100000
                }
                // CONSISTENCY(media-autoplay-aliases): autoStart is the
                // PowerPoint UI label and matches the OOXML verb naming;
                // autoplay is the existing canonical. Accept both on add
                // — Set already needs the same alias surface.
                string? apRaw = null;
                if (properties.TryGetValue("autoplay", out var apV)) apRaw = apV;
                else if (properties.TryGetValue("autostart", out apV)) apRaw = apV;
                var autoPlay = apRaw != null
                    && apRaw.Equals("true", StringComparison.OrdinalIgnoreCase);

                // loop: PowerPoint stores "Loop until Stopped" as
                // repeatCount="indefinite" on the player's cTn (inside the
                // CommonMediaNode). Default false.
                var loop = properties.TryGetValue("loop", out var loopV) && IsTruthy(loopV);

                AddMediaTimingNode(mediaSlide, mediaId, isVideo, vol, autoPlay, loop);

                mediaSlide.Save();

                // Count how many audio/video items of the same type are on the slide
                var sameTypeCount = mediaShapeTree.Elements<Picture>().Count(p =>
                {
                    var nvPr = p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                    return isVideo
                        ? nvPr?.GetFirstChild<Drawing.VideoFromFile>() != null
                        : nvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;
                });
                return $"/slide[{mediaSlideIdx}]/{(isVideo ? "video" : "audio")}[{sameTypeCount}]";
    }

    // ==================== OLE Object Insertion ====================
    //
    // Inserts an embedded OLE object into a slide. The structure follows
    // the PresentationML spec: a GraphicFrame hosting
    //   <a:graphicData uri="…/ole"><p:oleObj ... /></a:graphicData>
    // where p:oleObj carries progId + r:id (the payload relationship) and
    // an inner p:pic element rendering the icon preview.
    //
    // Caller props:
    //   src (required)  path to the file to embed
    //   progId          defaults to OleHelper.DetectProgId(src)
    //   width / height  EMU-parsed; defaults to 2in × 0.75in
    //   x / y           position in EMU; defaults to top-left (457200,457200)
    //   icon            path to a custom icon (png/jpg/emf); defaults to tiny PNG
    //   display         "icon" (default, sets showAsIcon) or "content"
    private string AddOle(string parentPath, int? index, Dictionary<string, string> properties)
    {
        properties ??= new Dictionary<string, string>();
        var srcPath = OfficeCli.Core.OleHelper.RequireSource(properties);
        OfficeCli.Core.OleHelper.WarnOnUnknownOleProps(properties);

        var oleSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
        if (!oleSlideMatch.Success)
            throw new ArgumentException("OLE objects must be added to a slide: /slide[N]");

        var oleSlideIdx = int.Parse(oleSlideMatch.Groups[1].Value);
        var oleSlideParts = GetSlideParts().ToList();
        if (oleSlideIdx < 1 || oleSlideIdx > oleSlideParts.Count)
            throw new ArgumentException($"Slide {oleSlideIdx} not found (total: {oleSlideParts.Count})");

        var oleSlidePart = oleSlideParts[oleSlideIdx - 1];
        var oleShapeTree = GetSlide(oleSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // 1. Create the embedded payload part.
        var (embedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(oleSlidePart, srcPath, _filePath);

        // 2. ProgID (explicit or auto-detected).
        var progId = OfficeCli.Core.OleHelper.ResolveProgId(properties, srcPath);

        // 3. Icon image part (placeholder PNG or user-supplied).
        var (_, oleIconRelId) = OfficeCli.Core.OleHelper.CreateIconPart(oleSlidePart, properties);

        // 4. Dimensions. width/height must be >= 1 EMU — zero / negative would
        // produce a <a:ext cx="0"/> which Office silently rejects on open and
        // throws the whole shape away (matches shape size validation).
        long oleCx = properties.TryGetValue("width", out var wv)
            ? ParseEmu(wv) : OfficeCli.Core.OleHelper.DefaultOleWidthEmu;
        long oleCy = properties.TryGetValue("height", out var hv)
            ? ParseEmu(hv) : OfficeCli.Core.OleHelper.DefaultOleHeightEmu;
        if (oleCx < 1)
            throw new ArgumentException($"Invalid ole width '{properties.GetValueOrDefault("width")}': must be >= 1 EMU.");
        if (oleCy < 1)
            throw new ArgumentException($"Invalid ole height '{properties.GetValueOrDefault("height")}': must be >= 1 EMU.");
        long oleX = properties.TryGetValue("x", out var xv) ? ParseEmu(xv) : 457200;
        long oleY = properties.TryGetValue("y", out var yv) ? ParseEmu(yv) : 457200;

        // 5. Display mode: icon (default) or content. Strict validation —
        // unknown values throw (see OleHelper.NormalizeOleDisplay).
        var oleDisplay = OfficeCli.Core.OleHelper.NormalizeOleDisplay(
            properties.GetValueOrDefault("display", "icon"));
        bool showAsIcon = oleDisplay != "content";

        // 6. Build the GraphicFrame + OleObject subtree. We lean on
        //    strong-typed p:oleObj / p:embed / p:pic from the SDK so
        //    attributes get schema-checked; only the outer GraphicFrame
        //    wrapper uses hand-built OuterXml because GraphicData.Uri is
        //    a string attribute, not a type particle.
        var oleShapeId = AcquireShapeId(oleShapeTree, properties);
        var oleName = properties.GetValueOrDefault("name", $"Object {oleShapeId}");

        var oleObj = new DocumentFormat.OpenXml.Presentation.OleObject
        {
            ShapeId = "",
            Name = oleName,
            ShowAsIcon = showAsIcon,
            Id = embedRelId,
            ImageWidth = (int)oleCx,
            ImageHeight = (int)oleCy,
            ProgId = progId,
        };
        // p:embed followColorScheme="full" — lets PowerPoint paint the
        // icon using the current slide theme accent, matching PPT's own
        // default for embed-mode OLE.
        oleObj.AppendChild(new DocumentFormat.OpenXml.Presentation.OleObjectEmbed
        {
            FollowColorScheme = DocumentFormat.OpenXml.Presentation.OleObjectFollowColorSchemeValues.Full,
        });

        // Inner p:pic holding the icon preview (bound to the image part we
        // just created). Structure mirrors a minimal non-animated picture.
        var olePic = new DocumentFormat.OpenXml.Presentation.Picture();
        olePic.NonVisualPictureProperties = new NonVisualPictureProperties(
            new NonVisualDrawingProperties { Id = 0U, Name = "" },
            new NonVisualPictureDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()
        );
        olePic.BlipFill = new BlipFill(
            new Drawing.Blip { Embed = oleIconRelId },
            new Drawing.Stretch(new Drawing.FillRectangle())
        );
        olePic.ShapeProperties = new ShapeProperties(
            new Drawing.Transform2D(
                new Drawing.Offset { X = oleX, Y = oleY },
                new Drawing.Extents { Cx = oleCx, Cy = oleCy }
            ),
            new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
        );
        oleObj.AppendChild(olePic);

        // 7. Wrap the OleObject in a GraphicFrame with the ole URI.
        var oleGraphicData = new Drawing.GraphicData(oleObj)
        {
            Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole",
        };
        var oleFrame = new GraphicFrame(
            new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = oleShapeId, Name = oleName },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            new Transform(
                new Drawing.Offset { X = oleX, Y = oleY },
                new Drawing.Extents { Cx = oleCx, Cy = oleCy }
            ),
            new Drawing.Graphic(oleGraphicData)
        );

        InsertAtPosition(oleShapeTree, oleFrame, index);
        GetSlide(oleSlidePart).Save();

        // Count OLE frames on this slide for the return path.
        var oleFrames = oleShapeTree.Elements<GraphicFrame>()
            .Count(gf => gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any());
        return $"/slide[{oleSlideIdx}]/ole[{oleFrames}]";
    }

    // Normalize a media trim input into the millisecond-integer string that
    // PowerPoint accepts for p14:trim/@st|@end. Accepted input forms:
    //   - bare ms count  ("200", "200ms")
    //   - seconds        ("0.2s", "1.5s")
    //   - hh:mm:ss(.fff) ("00:00:00.200", "00:01:30")
    // PowerPoint refuses to open a file whose @st is a hh:mm:ss literal
    // (0x80070570). Schema docs the timestamp form so it has to be tolerated
    // on input; the wire form is always ms-int.
    internal static string NormalizeMediaTimeMs(string raw, string propName)
    {
        if (string.IsNullOrWhiteSpace(raw))
            throw new ArgumentException($"Invalid '{propName}' value: empty.");
        var s = raw.Trim();

        // hh:mm:ss(.fff) — at least one colon
        if (s.IndexOf(':') >= 0)
        {
            var parts = s.Split(':');
            if (parts.Length < 2 || parts.Length > 3)
                throw new ArgumentException(
                    $"Invalid '{propName}' value: '{raw}'. Expected ms (e.g. '200'), seconds ('0.2s'), or 'hh:mm:ss.fff'.");
            double h = 0, m, sec;
            try
            {
                if (parts.Length == 3)
                {
                    h = double.Parse(parts[0], System.Globalization.CultureInfo.InvariantCulture);
                    m = double.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture);
                    sec = double.Parse(parts[2], System.Globalization.CultureInfo.InvariantCulture);
                }
                else
                {
                    m = double.Parse(parts[0], System.Globalization.CultureInfo.InvariantCulture);
                    sec = double.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            catch (FormatException)
            {
                throw new ArgumentException(
                    $"Invalid '{propName}' value: '{raw}'. Expected ms (e.g. '200'), seconds ('0.2s'), or 'hh:mm:ss.fff'.");
            }
            var totalMs = (h * 3600 + m * 60 + sec) * 1000;
            if (totalMs < 0)
                throw new ArgumentException($"Invalid '{propName}' value: '{raw}'. Must be non-negative.");
            return ((long)Math.Round(totalMs)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        // suffix: ms / s
        double scale = 1.0;
        var body = s;
        if (s.EndsWith("ms", StringComparison.OrdinalIgnoreCase))
        { body = s[..^2].Trim(); scale = 1.0; }
        else if (s.EndsWith("s", StringComparison.OrdinalIgnoreCase))
        { body = s[..^1].Trim(); scale = 1000.0; }

        if (!double.TryParse(body, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var n) || n < 0)
            throw new ArgumentException(
                $"Invalid '{propName}' value: '{raw}'. Expected ms (e.g. '200'), seconds ('0.2s'), or 'hh:mm:ss.fff'.");
        return ((long)Math.Round(n * scale)).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

}
