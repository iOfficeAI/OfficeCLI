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
                    throw new ArgumentException("'path' or 'src' property is required for picture type");

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

                // Resolve image from file/base64/URL
                var (imgStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
                using var imgStreamDispose = imgStream;

                // Embed image into slide part
                var imagePart = imgSlidePart.AddImagePart(imgPartType);
                imagePart.FeedData(imgStream);
                var imgRelId = imgSlidePart.GetIdOfPart(imagePart);

                // Dimensions (default: 6in x 4in)
                long cxEmu = 5486400; // 6 inches in EMUs
                long cyEmu = 3657600; // 4 inches in EMUs
                if (properties.TryGetValue("width", out var widthStr))
                    cxEmu = ParseEmu(widthStr);
                if (properties.TryGetValue("height", out var heightStr))
                    cyEmu = ParseEmu(heightStr);

                // Position (default: centered on slide)
                var (slideW, slideH) = GetSlideSize();
                long xEmu = (slideW - cxEmu) / 2;
                long yEmu = (slideH - cyEmu) / 2;
                if (properties.TryGetValue("x", out var xStr) || properties.TryGetValue("left", out xStr))
                    xEmu = ParseEmu(xStr);
                if (properties.TryGetValue("y", out var yStr) || properties.TryGetValue("top", out yStr))
                    yEmu = ParseEmu(yStr);

                var imgShapeId = (uint)(imgShapeTree.Elements<Shape>().Count() + imgShapeTree.Elements<Picture>().Count() + 2);
                var imgName = properties.GetValueOrDefault("name", $"Picture {imgShapeId}");
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

                imgShapeTree.AppendChild(picture);
                GetSlide(imgSlidePart).Save();

                var picCount = imgShapeTree.Elements<Picture>().Count();
                return $"/slide[{imgSlideIdx}]/picture[{picCount}]";
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

                // Build chart content BEFORE adding part (invalid type throws, must not leave empty part)
                var chartSpace = ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
                var chartPart = chartSlidePart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = chartSpace;
                chartPart.ChartSpace.Save();

                // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
                var deferredProps = properties
                    .Where(kv => ChartHelper.DeferredAddKeys.Contains(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (deferredProps.Count > 0)
                    ChartHelper.SetChartProperties(chartPart, deferredProps);

                // Position
                long chartX = properties.TryGetValue("x", out var xv) ? ParseEmu(xv) : 838200;     // ~2.3cm
                long chartY = properties.TryGetValue("y", out var yv) ? ParseEmu(yv) : 1825625;     // ~5cm
                long chartCx = properties.TryGetValue("width", out var wv) ? ParseEmu(wv) : 8229600; // ~22.9cm
                long chartCy = properties.TryGetValue("height", out var hv) ? ParseEmu(hv) : 4572000; // ~12.7cm

                var chartId = (uint)(chartShapeTree.ChildElements.Count + 2);
                var chartName = properties.GetValueOrDefault("name", chartTitle ?? $"Chart {chartId}");

                var chartGf = BuildChartGraphicFrame(chartSlidePart, chartPart, chartId, chartName,
                    chartX, chartY, chartCx, chartCy);
                chartShapeTree.AppendChild(chartGf);
                GetSlide(chartSlidePart).Save();

                var chartCount = chartShapeTree.Elements<GraphicFrame>()
                    .Count(gf => gf.Descendants<C.ChartReference>().Any());
                return $"/slide[{chartSlideIdx}]/chart[{chartCount}]";
    }


    private string AddMedia(string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
                var mediaSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!mediaSlideMatch.Success)
                    throw new ArgumentException("Media must be added to a slide: /slide[N]");

                if (!properties.TryGetValue("path", out var mediaPath))
                    throw new ArgumentException("'path' property required for media type");
                if (!File.Exists(mediaPath))
                    throw new FileNotFoundException($"Media file not found: {mediaPath}");

                var mediaSlideIdx = int.Parse(mediaSlideMatch.Groups[1].Value);
                var mediaSlideParts = GetSlideParts().ToList();
                if (mediaSlideIdx < 1 || mediaSlideIdx > mediaSlideParts.Count)
                    throw new ArgumentException($"Slide {mediaSlideIdx} not found (total: {mediaSlideParts.Count})");

                var mediaSlidePart = mediaSlideParts[mediaSlideIdx - 1];
                var mediaShapeTree = GetSlide(mediaSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var ext = Path.GetExtension(mediaPath).ToLowerInvariant();
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
                using (var mediaStream = File.OpenRead(mediaPath))
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

                var mediaId = (uint)(mediaShapeTree.ChildElements.Count + 2);
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

                mediaShapeTree.AppendChild(mediaPic);

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

                var picCount = mediaShapeTree.Elements<Picture>().Count();
                return $"/slide[{mediaSlideIdx}]/picture[{picCount}]";
    }


}
