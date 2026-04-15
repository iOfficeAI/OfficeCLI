// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;

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

                // Embed image into slide part
                var imagePart = imgSlidePart.AddImagePart(imgPartType);
                imagePart.FeedData(imgStream);
                imgStream.Position = 0;
                var imgRelId = imgSlidePart.GetIdOfPart(imagePart);

                // Dimensions (default: 6in x 4in, with auto aspect-ratio)
                // CONSISTENCY(picture-aspect): when only one dimension is
                // supplied, compute the other from native pixel ratio — same
                // behavior as WordHandler.AddPicture.
                bool hasWidth = properties.TryGetValue("width", out var widthStr);
                bool hasHeight = properties.TryGetValue("height", out var heightStr);
                long cxEmu = hasWidth ? ParseEmu(widthStr!) : 5486400;  // 6 inches fallback
                long cyEmu = hasHeight ? ParseEmu(heightStr!) : 3657600; // 4 inches fallback

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

                var imgShapeId = GenerateUniqueShapeId(imgShapeTree);
                var imgName = properties.GetValueOrDefault("name", $"Picture {imgShapeTree.Elements<Picture>().Count() + 1}");
                var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

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
                picture.BlipFill.AppendChild(new Drawing.Stretch(new Drawing.FillRectangle()));

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

                InsertAtPosition(imgShapeTree, picture, index);
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

                // Parse chart data
                var chartType = properties.FirstOrDefault(kv =>
                    kv.Key.Equals("charttype", StringComparison.OrdinalIgnoreCase)
                    || kv.Key.Equals("type", StringComparison.OrdinalIgnoreCase)).Value
                    ?? "column";
                var chartTitle = properties.GetValueOrDefault("title");
                var categories = ChartHelper.ParseCategories(properties);
                var seriesData = ChartHelper.ParseSeriesData(properties);

                if (seriesData.Count == 0)
                    throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                        "or series1=\"Revenue:100,200,300\"");

                // Position
                long chartX = properties.TryGetValue("x", out var xv) ? ParseEmu(xv) : 838200;     // ~2.3cm
                long chartY = properties.TryGetValue("y", out var yv) ? ParseEmu(yv) : 1825625;     // ~5cm
                long chartCx = properties.TryGetValue("width", out var wv) ? ParseEmu(wv) : 8229600; // ~22.9cm
                long chartCy = properties.TryGetValue("height", out var hv) ? ParseEmu(hv) : 4572000; // ~12.7cm
                var chartId = GenerateUniqueShapeId(chartShapeTree);
                var chartName = properties.GetValueOrDefault("name", chartTitle ?? $"Chart {chartShapeTree.Elements<GraphicFrame>().Count(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any() || IsExtendedChartFrame(gf)) + 1}");

                // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
                if (ChartExBuilder.IsExtendedChartType(chartType))
                {
                    var cxChartSpace = ChartExBuilder.BuildExtendedChartSpace(
                        chartType, chartTitle, categories, seriesData, properties);
                    var extChartPart = chartSlidePart.AddNewPart<ExtendedChartPart>();
                    extChartPart.ChartSpace = cxChartSpace;
                    extChartPart.ChartSpace.Save();

                    var chartGfEx = BuildExtendedChartGraphicFrame(chartSlidePart, extChartPart,
                        chartId, chartName, chartX, chartY, chartCx, chartCy);
                    InsertAtPosition(chartShapeTree, chartGfEx, index);
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
                long mX = properties.TryGetValue("x", out var mxv) ? ParseEmu(mxv) : (mediaSlideW - mCx) / 2;
                long mY = properties.TryGetValue("y", out var myv) ? ParseEmu(myv) : (mediaSlideH - mCy) / 2;

                var mediaId = GenerateUniqueShapeId(mediaShapeTree);
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

                // p14:trim (optional start/end trim in milliseconds)
                properties.TryGetValue("trimstart", out var trimStart);
                properties.TryGetValue("trimend", out var trimEnd);
                if (trimStart != null || trimEnd != null)
                {
                    var trim = new DocumentFormat.OpenXml.Office2010.PowerPoint.MediaTrim();
                    if (trimStart != null) trim.Start = trimStart;
                    if (trimEnd != null) trim.End = trimEnd;
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
                var autoPlay = properties.GetValueOrDefault("autoplay", "false")
                    .Equals("true", StringComparison.OrdinalIgnoreCase);

                AddMediaTimingNode(mediaSlide, mediaId, isVideo, vol, autoPlay);

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

        // 4. Dimensions.
        long oleCx = properties.TryGetValue("width", out var wv)
            ? ParseEmu(wv) : OfficeCli.Core.OleHelper.DefaultOleWidthEmu;
        long oleCy = properties.TryGetValue("height", out var hv)
            ? ParseEmu(hv) : OfficeCli.Core.OleHelper.DefaultOleHeightEmu;
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
        var oleShapeId = GenerateUniqueShapeId(oleShapeTree);
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

}
