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
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        switch (type.ToLowerInvariant())
        {
            case "slide":
                var presentationPart = _doc.PresentationPart
                    ?? throw new InvalidOperationException("Presentation not found");
                var presentation = presentationPart.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                var slideIdList = presentation.GetFirstChild<SlideIdList>()
                    ?? presentation.AppendChild(new SlideIdList());

                var newSlidePart = presentationPart.AddNewPart<SlidePart>();

                // Link slide to slideLayout (required by PowerPoint)
                var slideLayoutPart = ResolveSlideLayout(
                    presentationPart, properties.GetValueOrDefault("layout"));
                if (slideLayoutPart != null)
                    newSlidePart.AddPart(slideLayoutPart);

                newSlidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties { Id = 1, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties()
                        )
                    )
                );

                // Add title shape if text provided (ID starts at 2 since ShapeTree group uses ID=1)
                uint nextShapeId = 2;
                if (properties.TryGetValue("title", out var titleText))
                {
                    var titleShape = CreateTextShape(nextShapeId++, "Title", titleText, true);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(titleShape);
                }

                // Add content text if provided
                if (properties.TryGetValue("text", out var contentText))
                {
                    var textShape = CreateTextShape(nextShapeId++, "Content", contentText, false);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(textShape);
                }

                // Apply background if provided
                if (properties.TryGetValue("background", out var bgValue))
                    ApplySlideBackground(newSlidePart, bgValue);

                // Apply transition if provided
                if (properties.TryGetValue("transition", out var transValue))
                {
                    ApplyTransition(newSlidePart, transValue);
                    if (transValue.StartsWith("morph", StringComparison.OrdinalIgnoreCase))
                        AutoPrefixMorphNames(newSlidePart);
                }
                if (properties.TryGetValue("advancetime", out var advTime) || properties.TryGetValue("advanceTime", out advTime))
                    SetAdvanceTime(newSlidePart.Slide, advTime);
                if (properties.TryGetValue("advanceclick", out var advClick) || properties.TryGetValue("advanceClick", out advClick))
                    SetAdvanceClick(newSlidePart.Slide, IsTruthy(advClick));

                newSlidePart.Slide.Save();

                var maxId = slideIdList.Elements<SlideId>().Any()
                    ? slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255) + 1
                    : 256;
                var relId = presentationPart.GetIdOfPart(newSlidePart);

                if (index.HasValue && index.Value < slideIdList.Elements<SlideId>().Count())
                {
                    var refSlide = slideIdList.Elements<SlideId>().ElementAtOrDefault(index.Value);
                    if (refSlide != null)
                        slideIdList.InsertBefore(new SlideId { Id = maxId, RelationshipId = relId }, refSlide);
                    else
                        slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }
                else
                {
                    slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }

                presentation.Save();
                // Find the actual position of the inserted slide
                var slideIds = slideIdList.Elements<SlideId>().ToList();
                var insertedIdx = slideIds.FindIndex(s => s.RelationshipId?.Value == relId) + 1;
                return $"/slide[{insertedIdx}]";

            case "shape" or "textbox":
                var slideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException($"Shapes must be added to a slide: /slide[N]");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

                var slidePart = slideParts[slideIdx - 1];
                var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var text = properties.GetValueOrDefault("text", "");
                // Use max existing ID + 1 to avoid collisions after element deletion
                var maxExistingId = shapeTree.ChildElements
                    .Select(e => e.Descendants<NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value ?? 0)
                    .DefaultIfEmpty(1U)
                    .Max();
                var shapeId = maxExistingId + 1;
                var shapeName = properties.GetValueOrDefault("name", $"TextBox {shapeId}");

                // Auto-add !! prefix if the slide (or the next slide) has a morph transition
                if (!shapeName.StartsWith("!!") && !shapeName.StartsWith("TextBox ") && !shapeName.StartsWith("Content ") && shapeName != "")
                {
                    if (SlideHasMorphContext(slidePart, slideParts))
                        shapeName = "!!" + shapeName;
                }

                var newShape = CreateTextShape(shapeId, shapeName, text, false);

                if (properties.TryGetValue("size", out var sizeStr))
                {
                    var sizeVal = (int)Math.Round(ParseFontSize(sizeStr) * 100);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                }
                if (properties.TryGetValue("bold", out var boldStr))
                {
                    var isBold = IsTruthy(boldStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                }
                if (properties.TryGetValue("italic", out var italicStr))
                {
                    var isItalic = IsTruthy(italicStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                }
                if (properties.TryGetValue("color", out var colorVal))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var solidFill = BuildSolidFill(colorVal);
                        if (rProps is OpenXmlCompositeElement composite)
                        {
                            if (!composite.AddChild(solidFill, throwOnError: false))
                                rProps.AppendChild(solidFill);
                        }
                        else
                        {
                            rProps.AppendChild(solidFill);
                        }
                    }
                }

                // Schema order: font (latin/ea) after fill
                if (properties.TryGetValue("font", out var font))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Append(new Drawing.LatinFont { Typeface = font });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = font });
                    }
                }

                // Text margin (padding inside shape)
                if (properties.TryGetValue("margin", out var marginVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                        ApplyTextMargin(bodyPr, marginVal);
                }

                // Text alignment (horizontal)
                if (properties.TryGetValue("align", out var alignVal))
                {
                    var alignment = ParseTextAlignment(alignVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = alignment;
                    }
                }

                // Vertical alignment
                if (properties.TryGetValue("valign", out var valignVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                    {
                        bodyPr.Anchor = valignVal.ToLowerInvariant() switch
                        {
                            "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                            "center" or "middle" or "c" or "m" => Drawing.TextAnchoringTypeValues.Center,
                            "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                            _ => throw new ArgumentException($"Invalid valign: {valignVal}. Use top/center/bottom")
                        };
                    }
                }

                // Rotation
                if (properties.TryGetValue("rotation", out var rotStr) || properties.TryGetValue("rotate", out rotStr))
                {
                    // Will be set on Transform2D below
                }

                // Underline
                if (properties.TryGetValue("underline", out var ulVal))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Underline = ulVal.ToLowerInvariant() switch
                        {
                            "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                            "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                            "heavy" => Drawing.TextUnderlineValues.Heavy,
                            "dotted" => Drawing.TextUnderlineValues.Dotted,
                            "dash" => Drawing.TextUnderlineValues.Dash,
                            "wavy" => Drawing.TextUnderlineValues.Wavy,
                            "false" or "none" => Drawing.TextUnderlineValues.None,
                            _ => throw new ArgumentException($"Invalid underline value: '{ulVal}'. Valid values: single, double, heavy, dotted, dash, wavy, none.")
                        };
                    }
                }

                // Strikethrough
                if (properties.TryGetValue("strikethrough", out var stVal) || properties.TryGetValue("strike", out stVal))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Strike = stVal.ToLowerInvariant() switch
                        {
                            "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                            "double" => Drawing.TextStrikeValues.DoubleStrike,
                            "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                            _ => throw new ArgumentException($"Invalid strikethrough value: '{stVal}'. Valid values: single, double, none.")
                        };
                    }
                }

                // Line spacing
                if (properties.TryGetValue("lineSpacing", out var lsVal) || properties.TryGetValue("linespacing", out lsVal))
                {
                    var (lsInternal, lsIsPercent) = SpacingConverter.ParsePptLineSpacing(lsVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.LineSpacing>();
                        if (lsIsPercent)
                            pProps.AppendChild(new Drawing.LineSpacing(
                                new Drawing.SpacingPercent { Val = lsInternal }));
                        else
                            pProps.AppendChild(new Drawing.LineSpacing(
                                new Drawing.SpacingPoints { Val = lsInternal }));
                    }
                }

                // Space before/after
                if (properties.TryGetValue("spaceBefore", out var sbVal) || properties.TryGetValue("spacebefore", out sbVal))
                {
                    var sbInternal = SpacingConverter.ParsePptSpacing(sbVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                        pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = sbInternal }));
                    }
                }
                if (properties.TryGetValue("spaceAfter", out var saVal) || properties.TryGetValue("spaceafter", out saVal))
                {
                    var saInternal = SpacingConverter.ParsePptSpacing(saVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                        pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = saInternal }));
                    }
                }

                // AutoFit
                if (properties.TryGetValue("autofit", out var afVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                    {
                        switch (afVal.ToLowerInvariant())
                        {
                            case "true" or "normal": bodyPr.AppendChild(new Drawing.NormalAutoFit()); break;
                            case "shape": bodyPr.AppendChild(new Drawing.ShapeAutoFit()); break;
                            case "false" or "none": bodyPr.AppendChild(new Drawing.NoAutoFit()); break;
                        }
                    }
                }

                // Position and size (in EMU, 1cm = 360000 EMU; or parse as cm/in)
                {
                    long xEmu = 0, yEmu = 0;
                    long cxEmu = 9144000, cyEmu = 742950; // default: ~25.4cm x ~2.06cm
                    if (properties.TryGetValue("x", out var xStr)) xEmu = ParseEmu(xStr);
                    if (properties.TryGetValue("y", out var yStr)) yEmu = ParseEmu(yStr);
                    if (properties.TryGetValue("width", out var wStr)) cxEmu = ParseEmu(wStr);
                    if (properties.TryGetValue("height", out var hStr)) cyEmu = ParseEmu(hStr);

                    var xfrm = new Drawing.Transform2D
                    {
                        Offset = new Drawing.Offset { X = xEmu, Y = yEmu },
                        Extents = new Drawing.Extents { Cx = cxEmu, Cy = cyEmu }
                    };
                    if (properties.TryGetValue("rotation", out var rotVal) || properties.TryGetValue("rotate", out rotVal))
                    {
                        if (!double.TryParse(rotVal, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var rotDbl) || double.IsNaN(rotDbl) || double.IsInfinity(rotDbl))
                            throw new ArgumentException($"Invalid 'rotation' value: '{rotVal}'. Expected a finite number in degrees (e.g. 45, -90, 180.5).");
                        xfrm.Rotation = (int)(rotDbl * 60000);
                    }
                    newShape.ShapeProperties!.Transform2D = xfrm;

                    var presetName = properties.GetValueOrDefault("preset", "rect");
                    newShape.ShapeProperties.AppendChild(
                        new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = ParsePresetShape(presetName) }
                    );
                }

                // Shape fill (after xfrm and prstGeom to maintain schema order)
                if (properties.TryGetValue("fill", out var fillVal))
                {
                    ApplyShapeFill(newShape.ShapeProperties!, fillVal);
                }

                // Gradient fill
                if (properties.TryGetValue("gradient", out var gradVal))
                {
                    ApplyGradientFill(newShape.ShapeProperties!, gradVal);
                }

                // Opacity (alpha on fill) — like POI XSLFColor uses <a:alpha val="N"/>
                // Must come after gradient so it can apply to gradient stops too
                if (properties.TryGetValue("opacity", out var opacityVal))
                {
                    if (double.TryParse(opacityVal, System.Globalization.CultureInfo.InvariantCulture, out var alphaNum))
                    {
                        if (alphaNum > 1.0) alphaNum /= 100.0; // treat >1 as percentage (e.g. 30 → 0.30)
                        var alphaPct = (int)(alphaNum * 100000);
                        var solidFill = newShape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>();
                        if (solidFill != null)
                        {
                            var colorEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                                ?? solidFill.GetFirstChild<Drawing.SchemeColor>();
                            if (colorEl != null)
                            {
                                colorEl.RemoveAllChildren<Drawing.Alpha>();
                                colorEl.AppendChild(new Drawing.Alpha { Val = alphaPct });
                            }
                        }
                        var gradientFill = newShape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
                        if (gradientFill != null)
                        {
                            foreach (var stop in gradientFill.Descendants<Drawing.GradientStop>())
                            {
                                var stopColor = stop.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                                    ?? stop.GetFirstChild<Drawing.SchemeColor>();
                                if (stopColor != null)
                                {
                                    stopColor.RemoveAllChildren<Drawing.Alpha>();
                                    stopColor.AppendChild(new Drawing.Alpha { Val = alphaPct });
                                }
                            }
                        }
                    }
                }

                // Line/border (after fill per schema: xfrm → prstGeom → fill → ln)
                if (properties.TryGetValue("line", out var lineColor) || properties.TryGetValue("linecolor", out lineColor) || properties.TryGetValue("lineColor", out lineColor) || properties.TryGetValue("line.color", out lineColor))
                {
                    var outline = EnsureOutline(newShape.ShapeProperties!);
                    if (lineColor.Equals("none", StringComparison.OrdinalIgnoreCase))
                        outline.AppendChild(new Drawing.NoFill());
                    else
                        outline.AppendChild(BuildSolidFill(lineColor));
                }
                if (properties.TryGetValue("linewidth", out var lwStr) || properties.TryGetValue("lineWidth", out lwStr) || properties.TryGetValue("line.width", out lwStr))
                {
                    var outline = EnsureOutline(newShape.ShapeProperties!);
                    outline.Width = Core.EmuConverter.ParseLineWidth(lwStr);
                }

                // List style (bullet/numbered)
                if (properties.TryGetValue("list", out var listVal) || properties.TryGetValue("liststyle", out listVal))
                {
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        ApplyListStyle(pProps, listVal);
                    }
                }

                shapeTree.AppendChild(newShape);

                // Hyperlink on shape
                if (properties.TryGetValue("link", out var linkVal))
                    ApplyShapeHyperlink(slidePart, newShape, linkVal);

                // lineDash, effects, 3D, flip — delegate to SetRunOrShapeProperties
                var effectKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "linedash", "line.dash", "shadow", "glow", "reflection",
                      "softedge", "fliph", "flipv", "rot3d", "rotation3d",
                      "rotx", "roty", "rotz", "bevel", "beveltop", "bevelbottom",
                      "depth", "extrusion", "material", "lighting", "lightrig",
                      "spacing", "charspacing", "letterspacing",
                      "indent", "marginleft", "marl", "marginright", "marr",
                      "textfill", "textgradient", "geometry",
                      "baseline", "superscript", "subscript",
                      "textwarp", "wordart", "autofit",
                      "lineopacity", "line.opacity" };
                var effectProps = properties
                    .Where(kv => effectKeys.Contains(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (effectProps.Count > 0)
                    SetRunOrShapeProperties(effectProps, GetAllRuns(newShape), newShape);

                // Animation
                if (properties.TryGetValue("animation", out var animVal) ||
                    properties.TryGetValue("animate", out animVal))
                    ApplyShapeAnimation(slidePart, newShape, animVal);

                GetSlide(slidePart).Save();
                var shapeCount = shapeTree.Elements<Shape>().Count();
                return $"/slide[{slideIdx}]/shape[{shapeCount}]";

            case "picture" or "image" or "img":
            {
                if (!properties.TryGetValue("path", out var imgPath)
                    && !properties.TryGetValue("src", out imgPath))
                    throw new ArgumentException("'path' or 'src' property is required for picture type");
                if (!File.Exists(imgPath))
                    throw new FileNotFoundException($"Image file not found: {imgPath}");

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

                // Determine image type
                var imgExtension = Path.GetExtension(imgPath).ToLowerInvariant();
                var imgPartType = imgExtension switch
                {
                    ".png" => ImagePartType.Png,
                    ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                    ".gif" => ImagePartType.Gif,
                    ".bmp" => ImagePartType.Bmp,
                    ".tif" or ".tiff" => ImagePartType.Tiff,
                    ".emf" => ImagePartType.Emf,
                    ".wmf" => ImagePartType.Wmf,
                    ".svg" => ImagePartType.Svg,
                    _ => throw new ArgumentException($"Unsupported image format: {imgExtension}")
                };

                // Embed image into slide part
                var imagePart = imgSlidePart.AddImagePart(imgPartType);
                using (var imgStream = File.OpenRead(imgPath))
                    imagePart.FeedData(imgStream);
                var imgRelId = imgSlidePart.GetIdOfPart(imagePart);

                // Dimensions (default: 6in x 4in)
                long cxEmu = 5486400; // 6 inches in EMUs
                long cyEmu = 3657600; // 4 inches in EMUs
                if (properties.TryGetValue("width", out var widthStr))
                    cxEmu = ParseEmu(widthStr);
                if (properties.TryGetValue("height", out var heightStr))
                    cyEmu = ParseEmu(heightStr);

                // Position (default: centered on standard 10x7.5 inch slide)
                long xEmu = (9144000 - cxEmu) / 2;
                long yEmu = (6858000 - cyEmu) / 2;
                if (properties.TryGetValue("x", out var xStr))
                    xEmu = ParseEmu(xStr);
                if (properties.TryGetValue("y", out var yStr))
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
                picture.ShapeProperties.AppendChild(
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                );

                imgShapeTree.AppendChild(picture);
                GetSlide(imgSlidePart).Save();

                var picCount = imgShapeTree.Elements<Picture>().Count();
                return $"/slide[{imgSlideIdx}]/picture[{picCount}]";
            }

            case "chart":
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

            case "table":
            {
                var tblSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!tblSlideMatch.Success)
                    throw new ArgumentException("Tables must be added to a slide: /slide[N]");

                var tblSlideIdx = int.Parse(tblSlideMatch.Groups[1].Value);
                var tblSlideParts = GetSlideParts().ToList();
                if (tblSlideIdx < 1 || tblSlideIdx > tblSlideParts.Count)
                    throw new ArgumentException($"Slide {tblSlideIdx} not found (total: {tblSlideParts.Count})");

                var tblSlidePart = tblSlideParts[tblSlideIdx - 1];
                var tblShapeTree = GetSlide(tblSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var rowsStr = properties.GetValueOrDefault("rows", "3");
                var colsStr = properties.GetValueOrDefault("cols", "3");
                if (!int.TryParse(rowsStr, out var rows))
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Expected a positive integer.");
                if (!int.TryParse(colsStr, out var cols))
                    throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Expected a positive integer.");
                if (rows < 1 || cols < 1)
                    throw new ArgumentException("rows and cols must be >= 1");

                // Position & size
                long tblX = properties.TryGetValue("x", out var txStr) ? ParseEmu(txStr) : 457200; // ~1.27cm
                long tblY = properties.TryGetValue("y", out var tyStr) ? ParseEmu(tyStr) : 1600200; // ~4.44cm
                long tblCx = properties.TryGetValue("width", out var twStr) ? ParseEmu(twStr) : 8229600; // ~22.86cm
                long tblCy = properties.TryGetValue("height", out var thStr) ? ParseEmu(thStr) : (long)(rows * 370840); // ~1.03cm per row
                long colWidth = tblCx / cols;
                long rowHeight = tblCy / rows;

                var tblId = (uint)(tblShapeTree.ChildElements.Count + 2);

                // Build GraphicFrame
                var graphicFrame = new GraphicFrame();
                graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = tblId, Name = properties.GetValueOrDefault("name", $"Table {tblId}") },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                graphicFrame.Transform = new Transform(
                    new Drawing.Offset { X = tblX, Y = tblY },
                    new Drawing.Extents { Cx = tblCx, Cy = tblCy }
                );

                // Build table
                var table = new Drawing.Table();
                var tblProps = new Drawing.TableProperties { FirstRow = true, BandRow = true };

                // Apply table style if specified
                if (properties.TryGetValue("style", out var tblStyleVal))
                {
                    var styleId = tblStyleVal.ToLowerInvariant() switch
                    {
                        "medium1" or "mediumstyle1" => "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
                        "medium2" or "mediumstyle2" => "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
                        "medium3" or "mediumstyle3" => "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
                        "medium4" or "mediumstyle4" => "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
                        "light1" or "lightstyle1" => "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
                        "light2" or "lightstyle2" => "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
                        "light3" or "lightstyle3" => "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
                        "dark1" or "darkstyle1" => "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
                        "dark2" or "darkstyle2" => "{125E5076-3810-47DD-B79F-674D7AD40C01}",
                        "none" => "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
                        _ when tblStyleVal.StartsWith("{") => tblStyleVal,
                        _ => tblStyleVal
                    };
                    tblProps.AppendChild(new Drawing.TableStyleId(styleId));
                }

                table.Append(tblProps);

                var tableGrid = new Drawing.TableGrid();
                for (int c = 0; c < cols; c++)
                    tableGrid.Append(new Drawing.GridColumn { Width = colWidth });
                table.Append(tableGrid);

                for (int r = 0; r < rows; r++)
                {
                    var tableRow = new Drawing.TableRow { Height = rowHeight };
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = new Drawing.TableCell();
                        cell.Append(new Drawing.TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(new Drawing.EndParagraphRunProperties { Language = "en-US" })
                        ));
                        cell.Append(new Drawing.TableCellProperties());
                        tableRow.Append(cell);
                    }
                    table.Append(tableRow);
                }

                var graphic = new Drawing.Graphic(
                    new Drawing.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
                );
                graphicFrame.Append(graphic);
                tblShapeTree.AppendChild(graphicFrame);
                GetSlide(tblSlidePart).Save();

                var tblCount = tblShapeTree.Elements<GraphicFrame>()
                    .Count(gf => gf.Descendants<Drawing.Table>().Any());
                return $"/slide[{tblSlideIdx}]/table[{tblCount}]";
            }

            case "equation" or "formula" or "math":
            {
                if (!properties.TryGetValue("formula", out var eqFormula))
                    throw new ArgumentException("'formula' property is required for equation type");

                var eqSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!eqSlideMatch.Success)
                    throw new ArgumentException($"Equations must be added to a slide: /slide[N]");

                var eqSlideIdx = int.Parse(eqSlideMatch.Groups[1].Value);
                var eqSlideParts = GetSlideParts().ToList();
                if (eqSlideIdx < 1 || eqSlideIdx > eqSlideParts.Count)
                    throw new ArgumentException($"Slide {eqSlideIdx} not found (total: {eqSlideParts.Count})");

                var eqSlidePart = eqSlideParts[eqSlideIdx - 1];
                var eqShapeTree = GetSlide(eqSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var eqShapeId = (uint)(eqShapeTree.Elements<Shape>().Count() + eqShapeTree.Elements<Picture>().Count() + 2);
                var eqShapeName = properties.GetValueOrDefault("name", $"Equation {eqShapeId}");

                // Parse formula to OMML
                var mathContent = FormulaParser.Parse(eqFormula);
                M.OfficeMath oMath;
                if (mathContent is M.OfficeMath directMath)
                    oMath = directMath;
                else
                    oMath = new M.OfficeMath(mathContent.CloneNode(true));

                // Build the a14:m wrapper element via raw XML
                // PPT equations are embedded as: a:p > a14:m > m:oMathPara > m:oMath
                var mathPara = new M.Paragraph(oMath);
                var a14mXml = $"<a14:m xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\">{mathPara.OuterXml}</a14:m>";

                // Create shape with equation paragraph
                var eqShape = new Shape();
                eqShape.NonVisualShapeProperties = new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = eqShapeId, Name = eqShapeName },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                var eqSpPr = new ShapeProperties();
                {
                    long eqX = 838200, eqY = 2743200;        // default: ~2.33cm, ~7.62cm
                    long eqCx = 10515600, eqCy = 2743200;    // default: ~29.21cm, ~7.62cm
                    if (properties.TryGetValue("x", out var exStr)) eqX = ParseEmu(exStr);
                    if (properties.TryGetValue("y", out var eyStr)) eqY = ParseEmu(eyStr);
                    if (properties.TryGetValue("width", out var ewStr)) eqCx = ParseEmu(ewStr);
                    if (properties.TryGetValue("height", out var ehStr)) eqCy = ParseEmu(ehStr);
                    eqSpPr.Transform2D = new Drawing.Transform2D
                    {
                        Offset = new Drawing.Offset { X = eqX, Y = eqY },
                        Extents = new Drawing.Extents { Cx = eqCx, Cy = eqCy }
                    };
                }
                eqShape.ShapeProperties = eqSpPr;

                // Create text body with math paragraph
                var bodyProps = new Drawing.BodyProperties();
                var listStyle = new Drawing.ListStyle();
                var drawingPara = new Drawing.Paragraph();

                // Build mc:AlternateContent > mc:Choice(Requires="a14") > a14:m > m:oMathPara
                var a14mElement = new OpenXmlUnknownElement("a14", "m", "http://schemas.microsoft.com/office/drawing/2010/main");
                a14mElement.AppendChild(mathPara.CloneNode(true));

                var choice = new AlternateContentChoice();
                choice.Requires = "a14";
                choice.AppendChild(a14mElement);

                // Fallback: readable text for older versions
                var fallback = new AlternateContentFallback();
                var fallbackRun = new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text(FormulaParser.ToReadableText(mathPara))
                );
                fallback.AppendChild(fallbackRun);

                var altContent = new AlternateContent();
                altContent.AppendChild(choice);
                altContent.AppendChild(fallback);
                drawingPara.AppendChild(altContent);

                eqShape.TextBody = new TextBody(bodyProps, listStyle, drawingPara);
                eqShapeTree.AppendChild(eqShape);

                // Ensure slide root has xmlns:a14 and mc:Ignorable="a14" so PowerPoint accepts the equation
                var eqSlide = GetSlide(eqSlidePart);
                if (eqSlide.LookupNamespace("a14") == null)
                    eqSlide.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                if (eqSlide.LookupNamespace("mc") == null)
                    eqSlide.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                var currentIgnorable = eqSlide.MCAttributes?.Ignorable?.Value ?? "";
                if (!currentIgnorable.Contains("a14"))
                {
                    var newVal = string.IsNullOrEmpty(currentIgnorable) ? "a14" : $"{currentIgnorable} a14";
                    eqSlide.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = newVal };
                }
                eqSlide.Save();

                var eqShapeCount = eqShapeTree.Elements<Shape>().Count();
                return $"/slide[{eqSlideIdx}]/shape[{eqShapeCount}]";
            }

            case "notes":
            {
                var notesSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!notesSlideMatch.Success)
                    throw new ArgumentException("Notes must be added to a slide: /slide[N]");
                var notesSlideIdx = int.Parse(notesSlideMatch.Groups[1].Value);
                var notesSlideParts = GetSlideParts().ToList();
                if (notesSlideIdx < 1 || notesSlideIdx > notesSlideParts.Count)
                    throw new ArgumentException($"Slide {notesSlideIdx} not found (total: {notesSlideParts.Count})");
                var notesSlidePart = EnsureNotesSlidePart(notesSlideParts[notesSlideIdx - 1]);
                if (properties.TryGetValue("text", out var notesText))
                    SetNotesText(notesSlidePart, notesText);
                return $"/slide[{notesSlideIdx}]/notes";
            }

            case "video" or "audio" or "media":
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
                var posterPart = mediaSlidePart.AddImagePart(ImagePartType.Png);
                if (properties.TryGetValue("poster", out var posterPath) && File.Exists(posterPath))
                {
                    using var posterStream = File.OpenRead(posterPath);
                    posterPart.FeedData(posterStream);
                }
                else
                {
                    // Minimal 1x1 transparent PNG placeholder
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
                long mX = properties.TryGetValue("x", out var mxv) ? ParseEmu(mxv) : 1524000;
                long mY = properties.TryGetValue("y", out var myv) ? ParseEmu(myv) : 857250;
                long mCx = properties.TryGetValue("width", out var mwv) ? ParseEmu(mwv) : 9144000;
                long mCy = properties.TryGetValue("height", out var mhv) ? ParseEmu(mhv) : 5143500;

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
                    vol = (int)(volDbl * 1000); // 0-100 → 0-100000
                }
                var autoPlay = properties.GetValueOrDefault("autoplay", "false")
                    .Equals("true", StringComparison.OrdinalIgnoreCase);

                AddMediaTimingNode(mediaSlide, mediaId, isVideo, vol, autoPlay);

                mediaSlide.Save();

                var picCount = mediaShapeTree.Elements<Picture>().Count();
                return $"/slide[{mediaSlideIdx}]/picture[{picCount}]";
            }

            case "connector" or "connection":
            {
                var cxnSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!cxnSlideMatch.Success)
                    throw new ArgumentException("Connectors must be added to a slide: /slide[N]");

                var cxnSlideIdx = int.Parse(cxnSlideMatch.Groups[1].Value);
                var cxnSlideParts = GetSlideParts().ToList();
                if (cxnSlideIdx < 1 || cxnSlideIdx > cxnSlideParts.Count)
                    throw new ArgumentException($"Slide {cxnSlideIdx} not found (total: {cxnSlideParts.Count})");

                var cxnSlidePart = cxnSlideParts[cxnSlideIdx - 1];
                var cxnShapeTree = GetSlide(cxnSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var cxnId = (uint)(cxnShapeTree.ChildElements.Count + 2);
                var cxnName = properties.GetValueOrDefault("name", $"Connector {cxnId}");

                // Position: x1,y1 → x2,y2 or x,y,width,height
                long cxnX = properties.TryGetValue("x", out var cx1) ? ParseEmu(cx1) : 2000000;
                long cxnY = properties.TryGetValue("y", out var cy1) ? ParseEmu(cy1) : 3000000;
                long cxnCx = properties.TryGetValue("width", out var cw) ? ParseEmu(cw) : 4000000;
                long cxnCy = properties.TryGetValue("height", out var ch) ? ParseEmu(ch) : 0;

                var connector = new ConnectionShape();
                var cxnNvProps = new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = cxnId, Name = cxnName },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );

                // Connect to shapes if specified
                var cxnDrawProps = cxnNvProps.NonVisualConnectorShapeDrawingProperties!;
                if (properties.TryGetValue("startshape", out var startId))
                {
                    if (!uint.TryParse(startId, out var startIdVal))
                        throw new ArgumentException($"Invalid 'startshape' value: '{startId}'. Expected a positive integer (shape ID).");
                    cxnDrawProps.StartConnection = new Drawing.StartConnection { Id = startIdVal, Index = 0 };
                }
                if (properties.TryGetValue("endshape", out var endId))
                {
                    if (!uint.TryParse(endId, out var endIdVal))
                        throw new ArgumentException($"Invalid 'endshape' value: '{endId}'. Expected a positive integer (shape ID).");
                    cxnDrawProps.EndConnection = new Drawing.EndConnection { Id = endIdVal, Index = 0 };
                }

                connector.NonVisualConnectionShapeProperties = cxnNvProps;
                connector.ShapeProperties = new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = cxnX, Y = cxnY },
                        new Drawing.Extents { Cx = cxnCx, Cy = cxnCy }
                    ),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    {
                        Preset = properties.GetValueOrDefault("preset", "straightConnector1").ToLowerInvariant() switch
                        {
                            "straight" or "straightconnector1" => Drawing.ShapeTypeValues.StraightConnector1,
                            "elbow" or "bentconnector3" => Drawing.ShapeTypeValues.BentConnector3,
                            "curve" or "curvedconnector3" => Drawing.ShapeTypeValues.CurvedConnector3,
                            _ => throw new ArgumentException($"Invalid connector preset: '{properties.GetValueOrDefault("preset", "straightConnector1")}'. Valid values: straight, elbow, curve.")
                        }
                    }
                );

                // Line style
                var cxnOutline = new Drawing.Outline { Width = 12700 }; // 1pt default
                if (properties.TryGetValue("lineColor", out var cxnColor2) || properties.TryGetValue("linecolor", out cxnColor2)
                    || properties.TryGetValue("line", out cxnColor2))
                    cxnOutline.AppendChild(BuildSolidFill(cxnColor2));
                else
                    cxnOutline.AppendChild(BuildSolidFill("000000"));
                if (properties.TryGetValue("linewidth", out var lwVal) || properties.TryGetValue("lineWidth", out lwVal))
                    cxnOutline.Width = Core.EmuConverter.ParseLineWidth(lwVal);
                if (properties.TryGetValue("lineDash", out var cxnDash) || properties.TryGetValue("linedash", out cxnDash))
                {
                    cxnOutline.AppendChild(new Drawing.PresetDash
                    {
                        Val = cxnDash.ToLowerInvariant() switch
                        {
                            "solid" => Drawing.PresetLineDashValues.Solid,
                            "dot" => Drawing.PresetLineDashValues.Dot,
                            "dash" => Drawing.PresetLineDashValues.Dash,
                            "dashdot" => Drawing.PresetLineDashValues.DashDot,
                            "longdash" => Drawing.PresetLineDashValues.LargeDash,
                            "longdashdot" => Drawing.PresetLineDashValues.LargeDashDot,
                            "sysdot" => Drawing.PresetLineDashValues.SystemDot,
                            "sysdash" => Drawing.PresetLineDashValues.SystemDash,
                            _ => Drawing.PresetLineDashValues.Solid
                        }
                    });
                }
                if (properties.TryGetValue("rotation", out var cxnRot))
                {
                    if (int.TryParse(cxnRot, out var rotDeg))
                        connector.ShapeProperties.Transform2D!.Rotation = rotDeg * 60000;
                }
                connector.ShapeProperties.AppendChild(cxnOutline);

                cxnShapeTree.AppendChild(connector);
                GetSlide(cxnSlidePart).Save();

                var cxnCount = cxnShapeTree.Elements<ConnectionShape>().Count();
                return $"/slide[{cxnSlideIdx}]/connector[{cxnCount}]";
            }

            case "group":
            {
                var grpSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!grpSlideMatch.Success)
                    throw new ArgumentException("Groups must be added to a slide: /slide[N]");

                var grpSlideIdx = int.Parse(grpSlideMatch.Groups[1].Value);
                var grpSlideParts = GetSlideParts().ToList();
                if (grpSlideIdx < 1 || grpSlideIdx > grpSlideParts.Count)
                    throw new ArgumentException($"Slide {grpSlideIdx} not found (total: {grpSlideParts.Count})");

                var grpSlidePart = grpSlideParts[grpSlideIdx - 1];
                var grpShapeTree = GetSlide(grpSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var grpId = (uint)(grpShapeTree.ChildElements.Count + 2);
                var grpName = properties.GetValueOrDefault("name", $"Group {grpId}");

                // Parse shape paths to group: shapes="1,2,3" (shape indices)
                if (!properties.TryGetValue("shapes", out var shapesStr))
                    throw new ArgumentException("'shapes' property required: comma-separated shape indices to group (e.g. shapes=1,2,3)");

                var shapeParts = shapesStr.Split(',');
                var shapeIndices = new List<int>();
                foreach (var sp in shapeParts)
                {
                    if (!int.TryParse(sp.Trim(), out var idx))
                        throw new ArgumentException($"Invalid 'shapes' value: '{sp.Trim()}' is not a valid integer. Expected comma-separated shape indices (e.g. shapes=1,2,3).");
                    shapeIndices.Add(idx);
                }
                var allShapes = grpShapeTree.Elements<Shape>().ToList();

                // Collect shapes to group (in reverse order to maintain indices during removal)
                var toGroup = new List<Shape>();
                foreach (var si in shapeIndices.OrderBy(i => i))
                {
                    if (si < 1 || si > allShapes.Count)
                        throw new ArgumentException($"Shape {si} not found (total: {allShapes.Count})");
                    toGroup.Add(allShapes[si - 1]);
                }

                // Calculate bounding box
                long minX = long.MaxValue, minY = long.MaxValue, maxX = long.MinValue, maxY = long.MinValue;
                bool hasTransform = false;
                foreach (var s in toGroup)
                {
                    var xfrm = s.ShapeProperties?.Transform2D;
                    if (xfrm?.Offset == null || xfrm.Extents == null) continue;
                    hasTransform = true;
                    long sx = xfrm.Offset.X ?? 0;
                    long sy = xfrm.Offset.Y ?? 0;
                    long scx = xfrm.Extents.Cx ?? 0;
                    long scy = xfrm.Extents.Cy ?? 0;
                    if (sx < minX) minX = sx;
                    if (sy < minY) minY = sy;
                    if (sx + scx > maxX) maxX = sx + scx;
                    if (sy + scy > maxY) maxY = sy + scy;
                }
                if (!hasTransform) { minX = 0; minY = 0; maxX = 0; maxY = 0; }

                var groupShape = new GroupShape();
                groupShape.NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = grpId, Name = grpName },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                groupShape.GroupShapeProperties = new GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = minX, Y = minY },
                        new Drawing.Extents { Cx = maxX - minX, Cy = maxY - minY },
                        new Drawing.ChildOffset { X = minX, Y = minY },
                        new Drawing.ChildExtents { Cx = maxX - minX, Cy = maxY - minY }
                    )
                );

                // Move shapes into group
                foreach (var s in toGroup)
                {
                    s.Remove();
                    groupShape.AppendChild(s);
                }

                grpShapeTree.AppendChild(groupShape);
                GetSlide(grpSlidePart).Save();

                var grpCount = grpShapeTree.Elements<GroupShape>().Count();
                return $"/slide[{grpSlideIdx}]/group[{grpCount}]";
            }

            case "row" or "tr":
            {
                // Resolve parent table via logical path
                var rowLogical = ResolveLogicalPath(parentPath);
                if (!rowLogical.HasValue || rowLogical.Value.element is not Drawing.Table rowTable)
                    throw new ArgumentException("Rows can only be added to a table: /slide[N]/table[M]");

                var rowSlidePart = rowLogical.Value.slidePart;

                // Determine column count from existing grid
                var existingColCount = rowTable.Elements<Drawing.TableGrid>().FirstOrDefault()
                    ?.Elements<Drawing.GridColumn>().Count() ?? 1;
                int newColCount = existingColCount;
                if (properties.TryGetValue("cols", out var rcVal))
                {
                    if (!int.TryParse(rcVal, out newColCount))
                        throw new ArgumentException($"Invalid 'cols' value: '{rcVal}'. Expected a positive integer.");
                }

                // Row height: default from first existing row, or 370840 EMU (~1cm)
                long newRowHeight = properties.TryGetValue("height", out var rhVal)
                    ? ParseEmu(rhVal)
                    : rowTable.Elements<Drawing.TableRow>().FirstOrDefault()?.Height?.Value ?? 370840;

                var newTblRow = new Drawing.TableRow { Height = newRowHeight };
                for (int c = 0; c < newColCount; c++)
                {
                    var newTblCell = new Drawing.TableCell();
                    var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
                    var bodyProps = new Drawing.BodyProperties();
                    var listStyle = new Drawing.ListStyle();
                    var cellPara = new Drawing.Paragraph();
                    if (!string.IsNullOrEmpty(cellText))
                        cellPara.Append(new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US" },
                            new Drawing.Text(cellText)));
                    else
                        cellPara.Append(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                    newTblCell.Append(new Drawing.TextBody(bodyProps, listStyle, cellPara));
                    newTblCell.Append(new Drawing.TableCellProperties());
                    newTblRow.Append(newTblCell);
                }

                if (index.HasValue)
                {
                    var existingRows = rowTable.Elements<Drawing.TableRow>().ToList();
                    if (index.Value < existingRows.Count)
                        rowTable.InsertBefore(newTblRow, existingRows[index.Value]);
                    else
                        rowTable.AppendChild(newTblRow);
                }
                else
                {
                    rowTable.AppendChild(newTblRow);
                }

                GetSlide(rowSlidePart).Save();
                var rowIdx = rowTable.Elements<Drawing.TableRow>().ToList().IndexOf(newTblRow) + 1;
                return $"{parentPath}/tr[{rowIdx}]";
            }

            case "cell" or "tc":
            {
                // Resolve parent row via logical path
                var cellLogical = ResolveLogicalPath(parentPath);
                if (!cellLogical.HasValue || cellLogical.Value.element is not Drawing.TableRow cellRow)
                    throw new ArgumentException("Cells can only be added to a table row: /slide[N]/table[M]/tr[R]");

                var cellSlidePart = cellLogical.Value.slidePart;

                var newCell = new Drawing.TableCell();
                var cBodyProps = new Drawing.BodyProperties();
                var cListStyle = new Drawing.ListStyle();
                var cPara = new Drawing.Paragraph();
                if (properties.TryGetValue("text", out var cText) && !string.IsNullOrEmpty(cText))
                    cPara.Append(new Drawing.Run(
                        new Drawing.RunProperties { Language = "en-US" },
                        new Drawing.Text(cText)));
                else
                    cPara.Append(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                newCell.Append(new Drawing.TextBody(cBodyProps, cListStyle, cPara));
                newCell.Append(new Drawing.TableCellProperties());

                if (index.HasValue)
                {
                    var existingCells = cellRow.Elements<Drawing.TableCell>().ToList();
                    if (index.Value < existingCells.Count)
                        cellRow.InsertBefore(newCell, existingCells[index.Value]);
                    else
                        cellRow.AppendChild(newCell);
                }
                else
                {
                    cellRow.AppendChild(newCell);
                }

                GetSlide(cellSlidePart).Save();
                var cellIdx = cellRow.Elements<Drawing.TableCell>().ToList().IndexOf(newCell) + 1;
                return $"{parentPath}/tc[{cellIdx}]";
            }

            case "animation" or "animate":
            {
                // Add animation to a shape: parentPath must be /slide[N]/shape[M]
                var animMatch = System.Text.RegularExpressions.Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
                if (!animMatch.Success)
                    throw new ArgumentException("Animations must be added to a shape: /slide[N]/shape[M]");

                var animSlideIdx = int.Parse(animMatch.Groups[1].Value);
                var animShapeIdx = int.Parse(animMatch.Groups[2].Value);
                var (animSlidePart, animShape) = ResolveShape(animSlideIdx, animShapeIdx);

                // Build animation value string from properties
                var effect = properties.GetValueOrDefault("effect", "fade");
                var cls = properties.GetValueOrDefault("class", "entrance");
                var duration = properties.GetValueOrDefault("duration", "500");
                var trigger = properties.GetValueOrDefault("trigger", "onclick");

                // Map trigger property to animation format
                var triggerPart = trigger.ToLowerInvariant() switch
                {
                    "onclick" or "click" => "click",
                    "after" or "afterprevious" => "after",
                    "with" or "withprevious" => "with",
                    _ => throw new ArgumentException($"Invalid animation trigger: '{trigger}'. Valid values: onclick, click, after, afterprevious, with, withprevious.")
                };

                var animValue = $"{effect}-{cls}-{duration}-{triggerPart}";
                ApplyShapeAnimation(animSlidePart, animShape, animValue);
                GetSlide(animSlidePart).Save();

                // Count animations on this shape
                var animShapeId = animShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0;
                var timing = GetSlide(animSlidePart).GetFirstChild<Timing>();
                var animCount = timing?.Descendants<ShapeTarget>()
                    .Count(st => st.ShapeId?.Value == animShapeId.ToString()) ?? 0;
                return $"{parentPath}/animation[{animCount}]";
            }

            case "paragraph" or "para":
            {
                // Add a paragraph to an existing shape: /slide[N]/shape[M]
                var paraParentMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
                if (!paraParentMatch.Success)
                    throw new ArgumentException("Paragraphs must be added to a shape: /slide[N]/shape[M]");

                var paraSlideIdx = int.Parse(paraParentMatch.Groups[1].Value);
                var paraShapeIdx = int.Parse(paraParentMatch.Groups[2].Value);
                var (paraSlidePart, paraShape) = ResolveShape(paraSlideIdx, paraShapeIdx);

                var textBody = paraShape.TextBody
                    ?? throw new InvalidOperationException("Shape has no text body");

                var newPara = new Drawing.Paragraph();
                var pProps = new Drawing.ParagraphProperties();

                // Paragraph-level properties
                if (properties.TryGetValue("align", out var pAlign))
                    pProps.Alignment = ParseTextAlignment(pAlign);
                if (properties.TryGetValue("indent", out var pIndent))
                    pProps.Indent = (int)ParseEmu(pIndent);
                if (properties.TryGetValue("marginLeft", out var pMarL) || properties.TryGetValue("marl", out pMarL))
                    pProps.LeftMargin = (int)ParseEmu(pMarL);
                if (properties.TryGetValue("marginRight", out var pMarR) || properties.TryGetValue("marr", out pMarR))
                    pProps.RightMargin = (int)ParseEmu(pMarR);
                if (properties.TryGetValue("list", out var pList) || properties.TryGetValue("liststyle", out pList))
                    ApplyListStyle(pProps, pList);

                newPara.ParagraphProperties = pProps;

                // Create initial run with text and run-level properties
                var paraText = properties.GetValueOrDefault("text", "");
                var newRun = new Drawing.Run();
                var rProps = new Drawing.RunProperties { Language = "en-US" };

                if (properties.TryGetValue("size", out var pSize))
                    rProps.FontSize = (int)Math.Round(ParseFontSize(pSize) * 100);
                if (properties.TryGetValue("bold", out var pBold))
                    rProps.Bold = IsTruthy(pBold);
                if (properties.TryGetValue("italic", out var pItalic))
                    rProps.Italic = IsTruthy(pItalic);
                // Schema order: solidFill before latin/ea
                if (properties.TryGetValue("color", out var pColor))
                    rProps.AppendChild(BuildSolidFill(pColor));
                if (properties.TryGetValue("font", out var pFont))
                {
                    rProps.Append(new Drawing.LatinFont { Typeface = pFont });
                    rProps.Append(new Drawing.EastAsianFont { Typeface = pFont });
                }
                if (properties.TryGetValue("spacing", out var pSpacing) || properties.TryGetValue("charspacing", out pSpacing))
                {
                    if (!double.TryParse(pSpacing, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pSpcVal))
                        throw new ArgumentException($"Invalid 'spacing' value: '{pSpacing}'. Expected a number in points.");
                    rProps.Spacing = (int)(pSpcVal * 100);
                }
                if (properties.TryGetValue("baseline", out var pBaseline))
                {
                    rProps.Baseline = pBaseline.ToLowerInvariant() switch
                    {
                        "super" or "true" => 30000,
                        "sub" => -25000,
                        _ => double.TryParse(pBaseline, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pBlVal) && !double.IsNaN(pBlVal) && !double.IsInfinity(pBlVal)
                            ? (int)(pBlVal * 1000)
                            : throw new ArgumentException($"Invalid 'baseline' value: '{pBaseline}'. Expected 'super', 'sub', or a percentage.")
                    };
                }

                newRun.RunProperties = rProps;
                newRun.Text = new Drawing.Text(paraText.Replace("\\n", "\n"));
                newPara.Append(newRun);

                if (index.HasValue && index.Value >= 0)
                {
                    var existingParas = textBody.Elements<Drawing.Paragraph>().ToList();
                    if (index.Value < existingParas.Count)
                        textBody.InsertBefore(newPara, existingParas[index.Value]);
                    else
                        textBody.Append(newPara);
                }
                else
                {
                    textBody.Append(newPara);
                }

                var paraCount = textBody.Elements<Drawing.Paragraph>().Count();
                GetSlide(paraSlidePart).Save();
                return $"/slide[{paraSlideIdx}]/shape[{paraShapeIdx}]/paragraph[{paraCount}]";
            }

            case "run":
            {
                // Add a run to a paragraph: /slide[N]/shape[M]/paragraph[P] or /slide[N]/shape[M]
                var runParaMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\](?:/paragraph\[(\d+)\])?$");
                if (!runParaMatch.Success)
                    throw new ArgumentException("Runs must be added to a shape or paragraph: /slide[N]/shape[M] or /slide[N]/shape[M]/paragraph[P]");

                var runSlideIdx = int.Parse(runParaMatch.Groups[1].Value);
                var runShapeIdx = int.Parse(runParaMatch.Groups[2].Value);
                var (runSlidePart, runShape) = ResolveShape(runSlideIdx, runShapeIdx);

                var runTextBody = runShape.TextBody
                    ?? throw new InvalidOperationException("Shape has no text body");

                Drawing.Paragraph targetPara;
                int targetParaIdx;
                if (runParaMatch.Groups[3].Success)
                {
                    targetParaIdx = int.Parse(runParaMatch.Groups[3].Value);
                    var paras = runTextBody.Elements<Drawing.Paragraph>().ToList();
                    if (targetParaIdx < 1 || targetParaIdx > paras.Count)
                        throw new ArgumentException($"Paragraph {targetParaIdx} not found");
                    targetPara = paras[targetParaIdx - 1];
                }
                else
                {
                    // Append to last paragraph
                    var paras = runTextBody.Elements<Drawing.Paragraph>().ToList();
                    targetPara = paras.LastOrDefault()
                        ?? throw new InvalidOperationException("Shape has no paragraphs");
                    targetParaIdx = paras.Count;
                }

                var runText = properties.GetValueOrDefault("text", "");
                var newRun = new Drawing.Run();
                var rProps = new Drawing.RunProperties { Language = "en-US" };

                if (properties.TryGetValue("size", out var rSize))
                    rProps.FontSize = (int)Math.Round(ParseFontSize(rSize) * 100);
                if (properties.TryGetValue("bold", out var rBold))
                    rProps.Bold = IsTruthy(rBold);
                if (properties.TryGetValue("italic", out var rItalic))
                    rProps.Italic = IsTruthy(rItalic);
                if (properties.TryGetValue("underline", out var rUnderline))
                    rProps.Underline = rUnderline.ToLowerInvariant() switch
                    {
                        "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                        "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                        "heavy" => Drawing.TextUnderlineValues.Heavy,
                        "dotted" => Drawing.TextUnderlineValues.Dotted,
                        "dash" => Drawing.TextUnderlineValues.Dash,
                        "wavy" => Drawing.TextUnderlineValues.Wavy,
                        "false" or "none" => Drawing.TextUnderlineValues.None,
                        _ => throw new ArgumentException($"Invalid underline value: '{rUnderline}'. Valid values: single, double, heavy, dotted, dash, wavy, none.")
                    };
                if (properties.TryGetValue("strikethrough", out var rStrike) || properties.TryGetValue("strike", out rStrike))
                    rProps.Strike = rStrike.ToLowerInvariant() switch
                    {
                        "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                        "double" => Drawing.TextStrikeValues.DoubleStrike,
                        "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                        _ => throw new ArgumentException($"Invalid strikethrough value: '{rStrike}'. Valid values: single, double, none.")
                    };
                // Schema order: solidFill before latin/ea
                if (properties.TryGetValue("color", out var rColor))
                    rProps.AppendChild(BuildSolidFill(rColor));
                if (properties.TryGetValue("font", out var rFont))
                {
                    rProps.Append(new Drawing.LatinFont { Typeface = rFont });
                    rProps.Append(new Drawing.EastAsianFont { Typeface = rFont });
                }
                if (properties.TryGetValue("spacing", out var rSpacing) || properties.TryGetValue("charspacing", out rSpacing))
                    rProps.Spacing = (int)(ParseHelpers.SafeParseDouble(rSpacing, "charspacing") * 100);
                if (properties.TryGetValue("baseline", out var rBaseline))
                {
                    rProps.Baseline = rBaseline.ToLowerInvariant() switch
                    {
                        "super" or "true" => 30000,
                        "sub" => -25000,
                        "none" or "false" or "0" => 0,
                        _ => (int)(ParseHelpers.SafeParseDouble(rBaseline, "baseline") * 1000)
                    };
                }
                else if (properties.TryGetValue("superscript", out var rSuper))
                    rProps.Baseline = IsTruthy(rSuper) ? 30000 : 0;
                else if (properties.TryGetValue("subscript", out var rSub))
                    rProps.Baseline = IsTruthy(rSub) ? -25000 : 0;

                newRun.RunProperties = rProps;
                newRun.Text = new Drawing.Text(runText.Replace("\\n", "\n"));

                // Append run to paragraph (before EndParagraphRunProperties if present)
                var endParaRun = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                if (endParaRun != null)
                    targetPara.InsertBefore(newRun, endParaRun);
                else
                    targetPara.Append(newRun);

                var runCount = targetPara.Elements<Drawing.Run>().Count();
                GetSlide(runSlidePart).Save();
                return $"/slide[{runSlideIdx}]/shape[{runShapeIdx}]/paragraph[{targetParaIdx}]/run[{runCount}]";
            }

            case "zoom" or "slidezoom" or "slide-zoom":
            {
                var zmSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!zmSlideMatch.Success)
                    throw new ArgumentException("Zoom must be added to a slide: /slide[N]");

                // Target slide (required)
                if (!properties.TryGetValue("target", out var targetStr) && !properties.TryGetValue("slide", out targetStr))
                    throw new ArgumentException("'target' property required for zoom type (target slide number, e.g. target=2)");
                if (!int.TryParse(targetStr, out var targetSlideNum))
                    throw new ArgumentException($"Invalid 'target' value: '{targetStr}'. Expected a slide number.");

                var zmSlideIdx = int.Parse(zmSlideMatch.Groups[1].Value);
                var zmSlideParts = GetSlideParts().ToList();
                if (zmSlideIdx < 1 || zmSlideIdx > zmSlideParts.Count)
                    throw new ArgumentException($"Slide {zmSlideIdx} not found (total: {zmSlideParts.Count})");
                if (targetSlideNum < 1 || targetSlideNum > zmSlideParts.Count)
                    throw new ArgumentException($"Target slide {targetSlideNum} not found (total: {zmSlideParts.Count})");

                var zmSlidePart = zmSlideParts[zmSlideIdx - 1];
                var zmShapeTree = GetSlide(zmSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");
                var targetSlidePart = zmSlideParts[targetSlideNum - 1];

                // Get target slide's SlideId from presentation.xml
                var zmPresentation = _doc.PresentationPart?.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                var zmSlideIdList = zmPresentation.GetFirstChild<SlideIdList>()
                    ?? throw new InvalidOperationException("No slides");
                var zmSlideIds = zmSlideIdList.Elements<SlideId>().ToList();
                var targetSldId = zmSlideIds[targetSlideNum - 1].Id!.Value;

                // Position and size (default: 8cm x 4.5cm, centered)
                long zmCx = 3048000; // ~8cm
                long zmCy = 1714500; // ~4.5cm
                if (properties.TryGetValue("width", out var zmW)) zmCx = ParseEmu(zmW);
                if (properties.TryGetValue("height", out var zmH)) zmCy = ParseEmu(zmH);
                long zmX = (12192000 - zmCx) / 2;
                long zmY = (6858000 - zmCy) / 2;
                if (properties.TryGetValue("x", out var zmXStr)) zmX = ParseEmu(zmXStr);
                if (properties.TryGetValue("y", out var zmYStr)) zmY = ParseEmu(zmYStr);

                var returnToParent = properties.TryGetValue("returntoparent", out var rtp) && IsTruthy(rtp) ? "1" : "0";
                var transitionDur = properties.GetValueOrDefault("transitiondur", "1000");

                // Generate shape IDs
                var zmShapeId = (uint)(zmShapeTree.ChildElements.Count + 2);
                var zmName = properties.GetValueOrDefault("name", $"Slide Zoom {zmShapeId}");
                var zmGuid = Guid.NewGuid().ToString("B").ToUpperInvariant();
                var zmCreationId = Guid.NewGuid().ToString("B").ToUpperInvariant();

                // Create a minimal 1x1 gray placeholder PNG (PowerPoint regenerates the thumbnail on open)
                byte[] placeholderPng = GenerateZoomPlaceholderPng();
                var zmImagePart = zmSlidePart.AddImagePart(ImagePartType.Png);
                using (var ms = new MemoryStream(placeholderPng))
                    zmImagePart.FeedData(ms);
                var zmImageRelId = zmSlidePart.GetIdOfPart(zmImagePart);

                // Create slide-to-slide relationship for fallback hyperlink
                var zmSlideRelId = zmSlidePart.CreateRelationshipToPart(targetSlidePart);

                // Build mc:AlternateContent programmatically (same pattern as morph transition)
                var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                var pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
                var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
                var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var pslzNs = "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom";
                var p166Ns = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
                var a16Ns = "http://schemas.microsoft.com/office/drawing/2014/main";

                var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);

                // === mc:Choice (for clients that support Slide Zoom) ===
                var choiceElement = new OpenXmlUnknownElement("mc", "Choice", mcNs);
                choiceElement.SetAttribute(new OpenXmlAttribute("", "Requires", null!, "pslz"));
                choiceElement.AddNamespaceDeclaration("pslz", pslzNs);

                var gfElement = new OpenXmlUnknownElement("p", "graphicFrame", pNs);
                gfElement.AddNamespaceDeclaration("a", aNs);
                gfElement.AddNamespaceDeclaration("r", rNs);

                // nvGraphicFramePr
                var nvGfPr = new OpenXmlUnknownElement("p", "nvGraphicFramePr", pNs);
                var cNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
                cNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmShapeId.ToString()));
                cNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, zmName));
                // creationId extension
                var extLst = new OpenXmlUnknownElement("a", "extLst", aNs);
                var ext = new OpenXmlUnknownElement("a", "ext", aNs);
                ext.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
                var creationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
                creationId.SetAttribute(new OpenXmlAttribute("", "id", null!, zmCreationId));
                ext.AppendChild(creationId);
                extLst.AppendChild(ext);
                cNvPr.AppendChild(extLst);
                nvGfPr.AppendChild(cNvPr);

                var cNvGfSpPr = new OpenXmlUnknownElement("p", "cNvGraphicFramePr", pNs);
                var gfLocks = new OpenXmlUnknownElement("a", "graphicFrameLocks", aNs);
                gfLocks.SetAttribute(new OpenXmlAttribute("", "noChangeAspect", null!, "1"));
                cNvGfSpPr.AppendChild(gfLocks);
                nvGfPr.AppendChild(cNvGfSpPr);
                nvGfPr.AppendChild(new OpenXmlUnknownElement("p", "nvPr", pNs));
                gfElement.AppendChild(nvGfPr);

                // xfrm (position/size)
                var gfXfrm = new OpenXmlUnknownElement("p", "xfrm", pNs);
                var gfOff = new OpenXmlUnknownElement("a", "off", aNs);
                gfOff.SetAttribute(new OpenXmlAttribute("", "x", null!, zmX.ToString()));
                gfOff.SetAttribute(new OpenXmlAttribute("", "y", null!, zmY.ToString()));
                var gfExt = new OpenXmlUnknownElement("a", "ext", aNs);
                gfExt.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                gfExt.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                gfXfrm.AppendChild(gfOff);
                gfXfrm.AppendChild(gfExt);
                gfElement.AppendChild(gfXfrm);

                // graphic > graphicData > pslz:sldZm
                var graphic = new OpenXmlUnknownElement("a", "graphic", aNs);
                var graphicData = new OpenXmlUnknownElement("a", "graphicData", aNs);
                graphicData.SetAttribute(new OpenXmlAttribute("", "uri", null!, pslzNs));

                var sldZm = new OpenXmlUnknownElement("pslz", "sldZm", pslzNs);
                var sldZmObj = new OpenXmlUnknownElement("pslz", "sldZmObj", pslzNs);
                sldZmObj.SetAttribute(new OpenXmlAttribute("", "sldId", null!, targetSldId.ToString()));
                sldZmObj.SetAttribute(new OpenXmlAttribute("", "cId", null!, "0"));

                var zmPr = new OpenXmlUnknownElement("pslz", "zmPr", pslzNs);
                zmPr.AddNamespaceDeclaration("p166", p166Ns);
                zmPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmGuid));
                zmPr.SetAttribute(new OpenXmlAttribute("", "returnToParent", null!, returnToParent));
                zmPr.SetAttribute(new OpenXmlAttribute("", "transitionDur", null!, transitionDur));

                // blipFill (thumbnail)
                var blipFill = new OpenXmlUnknownElement("p166", "blipFill", p166Ns);
                var blip = new OpenXmlUnknownElement("a", "blip", aNs);
                blip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, zmImageRelId));
                blipFill.AppendChild(blip);
                var stretch = new OpenXmlUnknownElement("a", "stretch", aNs);
                stretch.AppendChild(new OpenXmlUnknownElement("a", "fillRect", aNs));
                blipFill.AppendChild(stretch);
                zmPr.AppendChild(blipFill);

                // spPr (shape properties inside zoom)
                var zmSpPr = new OpenXmlUnknownElement("p166", "spPr", p166Ns);
                var zmSpXfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
                var zmSpOff = new OpenXmlUnknownElement("a", "off", aNs);
                zmSpOff.SetAttribute(new OpenXmlAttribute("", "x", null!, "0"));
                zmSpOff.SetAttribute(new OpenXmlAttribute("", "y", null!, "0"));
                var zmSpExt = new OpenXmlUnknownElement("a", "ext", aNs);
                zmSpExt.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                zmSpExt.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                zmSpXfrm.AppendChild(zmSpOff);
                zmSpXfrm.AppendChild(zmSpExt);
                zmSpPr.AppendChild(zmSpXfrm);
                var prstGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
                prstGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
                prstGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
                zmSpPr.AppendChild(prstGeom);
                var zmLn = new OpenXmlUnknownElement("a", "ln", aNs);
                zmLn.SetAttribute(new OpenXmlAttribute("", "w", null!, "3175"));
                var zmLnFill = new OpenXmlUnknownElement("a", "solidFill", aNs);
                var zmLnClr = new OpenXmlUnknownElement("a", "prstClr", aNs);
                zmLnClr.SetAttribute(new OpenXmlAttribute("", "val", null!, "ltGray"));
                zmLnFill.AppendChild(zmLnClr);
                zmLn.AppendChild(zmLnFill);
                zmSpPr.AppendChild(zmLn);
                zmPr.AppendChild(zmSpPr);

                sldZmObj.AppendChild(zmPr);
                sldZm.AppendChild(sldZmObj);
                graphicData.AppendChild(sldZm);
                graphic.AppendChild(graphicData);
                gfElement.AppendChild(graphic);
                choiceElement.AppendChild(gfElement);

                // === mc:Fallback (pic + hyperlink for older clients) ===
                var fallbackElement = new OpenXmlUnknownElement("mc", "Fallback", mcNs);
                var fbPic = new OpenXmlUnknownElement("p", "pic", pNs);
                fbPic.AddNamespaceDeclaration("a", aNs);
                fbPic.AddNamespaceDeclaration("r", rNs);

                var fbNvPicPr = new OpenXmlUnknownElement("p", "nvPicPr", pNs);
                var fbCNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
                fbCNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmShapeId.ToString()));
                fbCNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, zmName));
                var hlinkClick = new OpenXmlUnknownElement("a", "hlinkClick", aNs);
                hlinkClick.SetAttribute(new OpenXmlAttribute("r", "id", rNs, zmSlideRelId));
                hlinkClick.SetAttribute(new OpenXmlAttribute("", "action", null!, "ppaction://hlinksldjump"));
                fbCNvPr.AppendChild(hlinkClick);
                // Same creationId
                var fbExtLst = new OpenXmlUnknownElement("a", "extLst", aNs);
                var fbExt = new OpenXmlUnknownElement("a", "ext", aNs);
                fbExt.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
                var fbCreationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
                fbCreationId.SetAttribute(new OpenXmlAttribute("", "id", null!, zmCreationId));
                fbExt.AppendChild(fbCreationId);
                fbExtLst.AppendChild(fbExt);
                fbCNvPr.AppendChild(fbExtLst);
                fbNvPicPr.AppendChild(fbCNvPr);

                var fbCNvPicPr = new OpenXmlUnknownElement("p", "cNvPicPr", pNs);
                var picLocks = new OpenXmlUnknownElement("a", "picLocks", aNs);
                foreach (var lockAttr in new[] { "noGrp", "noRot", "noChangeAspect", "noMove", "noResize",
                    "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType" })
                    picLocks.SetAttribute(new OpenXmlAttribute("", lockAttr, null!, "1"));
                fbCNvPicPr.AppendChild(picLocks);
                fbNvPicPr.AppendChild(fbCNvPicPr);
                fbNvPicPr.AppendChild(new OpenXmlUnknownElement("p", "nvPr", pNs));
                fbPic.AppendChild(fbNvPicPr);

                // Fallback blipFill
                var fbBlipFill = new OpenXmlUnknownElement("p", "blipFill", pNs);
                var fbBlip = new OpenXmlUnknownElement("a", "blip", aNs);
                fbBlip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, zmImageRelId));
                fbBlipFill.AppendChild(fbBlip);
                var fbStretch = new OpenXmlUnknownElement("a", "stretch", aNs);
                fbStretch.AppendChild(new OpenXmlUnknownElement("a", "fillRect", aNs));
                fbBlipFill.AppendChild(fbStretch);
                fbPic.AppendChild(fbBlipFill);

                // Fallback spPr
                var fbSpPr = new OpenXmlUnknownElement("p", "spPr", pNs);
                var fbXfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
                var fbOff = new OpenXmlUnknownElement("a", "off", aNs);
                fbOff.SetAttribute(new OpenXmlAttribute("", "x", null!, zmX.ToString()));
                fbOff.SetAttribute(new OpenXmlAttribute("", "y", null!, zmY.ToString()));
                var fbExtSz = new OpenXmlUnknownElement("a", "ext", aNs);
                fbExtSz.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                fbExtSz.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                fbXfrm.AppendChild(fbOff);
                fbXfrm.AppendChild(fbExtSz);
                fbSpPr.AppendChild(fbXfrm);
                var fbGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
                fbGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
                fbGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
                fbSpPr.AppendChild(fbGeom);
                var fbLn = new OpenXmlUnknownElement("a", "ln", aNs);
                fbLn.SetAttribute(new OpenXmlAttribute("", "w", null!, "3175"));
                var fbLnFill = new OpenXmlUnknownElement("a", "solidFill", aNs);
                var fbLnClr = new OpenXmlUnknownElement("a", "prstClr", aNs);
                fbLnClr.SetAttribute(new OpenXmlAttribute("", "val", null!, "ltGray"));
                fbLnFill.AppendChild(fbLnClr);
                fbLn.AppendChild(fbLnFill);
                fbSpPr.AppendChild(fbLn);
                fbPic.AppendChild(fbSpPr);

                fallbackElement.AppendChild(fbPic);

                acElement.AppendChild(choiceElement);
                acElement.AppendChild(fallbackElement);
                zmShapeTree.AppendChild(acElement);
                GetSlide(zmSlidePart).Save();

                var zmCount = zmShapeTree.Elements<OpenXmlUnknownElement>()
                    .Count(e => e.LocalName == "AlternateContent");
                return $"/slide[{zmSlideIdx}]/zoom[{zmCount}]";
            }

            default:
            {
                // Try resolving logical paths (table/placeholder) first
                var logicalResult = ResolveLogicalPath(parentPath);
                SlidePart fbSlidePart;
                OpenXmlElement fbParent;

                if (logicalResult.HasValue)
                {
                    fbSlidePart = logicalResult.Value.slidePart;
                    fbParent = logicalResult.Value.element;
                }
                else
                {
                    // Generic fallback: navigate by XML localName
                    var allSegments = GenericXmlQuery.ParsePathSegments(parentPath);
                    if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                        throw new ArgumentException($"Generic add requires a path starting with /slide[N]: {parentPath}");

                    var fbSlideIdx = allSegments[0].Index!.Value;
                    var fbSlideParts = GetSlideParts().ToList();
                    if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                        throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

                    fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                    fbParent = GetSlide(fbSlidePart);
                    var remaining = allSegments.Skip(1).ToList();
                    if (remaining.Count > 0)
                    {
                        fbParent = GenericXmlQuery.NavigateByPath(fbParent, remaining)
                            ?? throw new ArgumentException($"Parent element not found: {parentPath}");
                    }
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                GetSlide(fbSlidePart).Save();

                // Build result path
                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }

    public void Remove(string path)
    {
        var slideMatch = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!slideMatch.Success)
            throw new ArgumentException($"Invalid path: {path}");

        var slideIdx = int.Parse(slideMatch.Groups[1].Value);

        if (!slideMatch.Groups[2].Success)
        {
            // Remove entire slide
            var presentationPart = _doc.PresentationPart
                ?? throw new InvalidOperationException("Presentation not found");
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");

            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideIds.Count})");

            var slideId = slideIds[slideIdx - 1];
            var relId = slideId.RelationshipId?.Value;
            slideId.Remove();
            if (relId != null)
                presentationPart.DeletePart(presentationPart.GetPartById(relId));
            presentation.Save();
            return;
        }

        // Remove element from slide
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shapes");

        var elementType = slideMatch.Groups[2].Value;
        var elementIdx = int.Parse(slideMatch.Groups[3].Value);

        if (elementType == "shape")
        {
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found");
            var shapeToRemove = shapes[elementIdx - 1];
            var shapeId = shapeToRemove.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0;
            if (shapeId > 0)
                RemoveShapeAnimations(GetSlide(slidePart), (uint)shapeId);
            shapeToRemove.Remove();
        }
        else if (elementType is "picture" or "pic" or "video" or "audio")
        {
            List<Picture> pics;
            if (elementType is "video")
                pics = shapeTree.Elements<Picture>()
                    .Where(p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<Drawing.VideoFromFile>() != null).ToList();
            else if (elementType is "audio")
                pics = shapeTree.Elements<Picture>()
                    .Where(p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<Drawing.AudioFromFile>() != null).ToList();
            else
                pics = shapeTree.Elements<Picture>().ToList();

            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {pics.Count})");

            var pic = pics[elementIdx - 1];
            RemovePictureWithCleanup(slidePart, shapeTree, pic);
        }
        else if (elementType == "table")
        {
            var tables = shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > tables.Count)
                throw new ArgumentException($"Table {elementIdx} not found");
            tables[elementIdx - 1].Remove();
        }
        else if (elementType == "chart")
        {
            var charts = shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > charts.Count)
                throw new ArgumentException($"Chart {elementIdx} not found");
            var chartGf = charts[elementIdx - 1];
            // Clean up ChartPart
            var chartRef = chartGf.Descendants<C.ChartReference>().FirstOrDefault();
            if (chartRef?.Id?.Value != null)
            {
                try { slidePart.DeletePart(chartRef.Id.Value); } catch { }
            }
            chartGf.Remove();
        }
        else if (elementType is "connector" or "connection")
        {
            var connectors = shapeTree.Elements<ConnectionShape>().ToList();
            if (elementIdx < 1 || elementIdx > connectors.Count)
                throw new ArgumentException($"Connector {elementIdx} not found");
            connectors[elementIdx - 1].Remove();
        }
        else if (elementType == "group")
        {
            // Ungroup: move children back to parent shape tree, then remove group
            var groups = shapeTree.Elements<GroupShape>().ToList();
            if (elementIdx < 1 || elementIdx > groups.Count)
                throw new ArgumentException($"Group {elementIdx} not found");
            var group = groups[elementIdx - 1];
            // Recursively clean up any pictures inside the group before ungrouping
            var children = group.ChildElements
                .Where(e => e is Shape or Picture or ConnectionShape or GraphicFrame or GroupShape)
                .ToList();
            foreach (var child in children)
            {
                child.Remove();
                shapeTree.AppendChild(child);
            }
            group.Remove();
        }
        else if (elementType is "zoom" or "slidezoom")
        {
            var zoomElements = GetZoomElements(shapeTree);
            if (elementIdx < 1 || elementIdx > zoomElements.Count)
                throw new ArgumentException($"Zoom {elementIdx} not found (total: {zoomElements.Count})");
            var zmAc = zoomElements[elementIdx - 1];
            // Clean up image relationship if not referenced by other elements
            var zmBlip = zmAc.Descendants().FirstOrDefault(d => d.LocalName == "blip");
            if (zmBlip != null)
            {
                var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var embedAttr = zmBlip.GetAttribute("embed", rNs);
                if (!string.IsNullOrEmpty(embedAttr.Value))
                {
                    var relId = embedAttr.Value;
                    // Check if any other element references this image
                    zmAc.Remove();
                    var slideXml = GetSlide(slidePart).OuterXml;
                    if (!slideXml.Contains(relId))
                    {
                        try { slidePart.DeletePart(relId); } catch { }
                    }
                    GetSlide(slidePart).Save();
                    return;
                }
            }
            zmAc.Remove();
        }
        else
        {
            throw new ArgumentException($"Unknown element type: {elementType}. Supported: shape, picture, video, audio, table, chart, connector/connection, group, zoom");
        }

        GetSlide(slidePart).Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 1: Move entire slide (reorder)
        var slideOnlyMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success)
        {
            var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
            var movePresentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = movePresentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideIds.Count})");

            var slideId = slideIds[slideIdx - 1];
            slideId.Remove();

            if (index.HasValue)
            {
                var remaining = slideIdList.Elements<SlideId>().ToList();
                if (index.Value >= 0 && index.Value < remaining.Count)
                    remaining[index.Value].InsertBeforeSelf(slideId);
                else
                    slideIdList.AppendChild(slideId);
            }
            else
            {
                slideIdList.AppendChild(slideId);
            }

            movePresentation.Save();
            var newSlideIds = slideIdList.Elements<SlideId>().ToList();
            var newIdx = newSlideIds.IndexOf(slideId) + 1;
            return $"/slide[{newIdx}]";
        }

        // Case 2: Move element within/across slides
        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);

        // Determine target
        string effectiveParentPath;
        SlidePart tgtSlidePart;
        ShapeTree tgtShapeTree;

        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within same parent
            tgtSlidePart = srcSlidePart;
            tgtShapeTree = GetSlide(srcSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
            var srcSlideIdx = slideParts.IndexOf(srcSlidePart) + 1;
            effectiveParentPath = $"/slide[{srcSlideIdx}]";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
            if (!tgtSlideMatch.Success)
                throw new ArgumentException($"Target must be a slide: /slide[N]");
            var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
            if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {tgtSlideIdx} not found (total: {slideParts.Count})");
            tgtSlidePart = slideParts[tgtSlideIdx - 1];
            tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
        }

        // Copy relationships BEFORE removing from source (so rel IDs are still accessible)
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(srcElement, srcSlidePart, tgtSlidePart);

        srcElement.Remove();

        InsertAtPosition(tgtShapeTree, srcElement, index);

        GetSlide(srcSlidePart).Save();
        if (srcSlidePart != tgtSlidePart)
            GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(effectiveParentPath, srcElement, tgtShapeTree);
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 1: Swap two slides
        var slide1Match = Regex.Match(path1, @"^/slide\[(\d+)\]$");
        var slide2Match = Regex.Match(path2, @"^/slide\[(\d+)\]$");
        if (slide1Match.Success && slide2Match.Success)
        {
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            var idx1 = int.Parse(slide1Match.Groups[1].Value);
            var idx2 = int.Parse(slide2Match.Groups[1].Value);
            if (idx1 < 1 || idx1 > slideIds.Count) throw new ArgumentException($"Slide {idx1} not found (total: {slideIds.Count})");
            if (idx2 < 1 || idx2 > slideIds.Count) throw new ArgumentException($"Slide {idx2} not found (total: {slideIds.Count})");
            if (idx1 == idx2) return (path1, path2);

            SwapXmlElements(slideIds[idx1 - 1], slideIds[idx2 - 1]);
            presentation.Save();
            return ($"/slide[{idx2}]", $"/slide[{idx1}]");
        }

        // Case 2: Swap two elements within the same slide
        var (slide1Part, elem1) = ResolveSlideElement(path1, slideParts);
        var (slide2Part, elem2) = ResolveSlideElement(path2, slideParts);
        if (slide1Part != slide2Part)
            throw new ArgumentException("Cannot swap elements on different slides");

        SwapXmlElements(elem1, elem2);
        GetSlide(slide1Part).Save();

        var slideIdx = slideParts.IndexOf(slide1Part) + 1;
        var parentPath = $"/slide[{slideIdx}]";
        var shapeTree = GetSlide(slide1Part).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");
        var newPath1 = ComputeElementPath(parentPath, elem1, shapeTree);
        var newPath2 = ComputeElementPath(parentPath, elem2, shapeTree);
        return (newPath1, newPath2);
    }

    internal static void SwapXmlElements(OpenXmlElement a, OpenXmlElement b)
    {
        if (a == b || a.Parent == null || b.Parent == null) return;
        var parent = a.Parent;
        var aNext = a.NextSibling();
        var bNext = b.NextSibling();

        a.Remove();
        b.Remove();

        if (aNext == b)
        {
            // A was directly before B: [... A B ...] → [... B A ...]
            if (bNext != null)
                bNext.InsertBeforeSelf(b);
            else
                parent.AppendChild(b);
            b.InsertAfterSelf(a);
        }
        else if (bNext == a)
        {
            // B was directly before A: [... B A ...] → [... A B ...]
            if (aNext != null)
                aNext.InsertBeforeSelf(a);
            else
                parent.AppendChild(a);
            a.InsertBeforeSelf(b);
        }
        else
        {
            // Non-adjacent: insert each where the other was
            if (aNext != null)
                aNext.InsertBeforeSelf(b);
            else
                parent.AppendChild(b);
            if (bNext != null)
                bNext.InsertBeforeSelf(a);
            else
                parent.AppendChild(a);
        }
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var slideParts = GetSlideParts().ToList();

        // Whole-slide clone: --from /slide[N] to /
        var slideCloneMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideCloneMatch.Success && (targetParentPath is "/" or "" or "/presentation"))
        {
            return CloneSlide(slideCloneMatch, slideParts, index);
        }

        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);
        var clone = srcElement.CloneNode(true);

        var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
        if (!tgtSlideMatch.Success)
            throw new ArgumentException($"Target must be a slide: /slide[N]");
        var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
        if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {tgtSlideIdx} not found (total: {slideParts.Count})");

        var tgtSlidePart = slideParts[tgtSlideIdx - 1];
        var tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // Copy relationships if across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(clone, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, clone, index);
        GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(targetParentPath, clone, tgtShapeTree);
    }

    /// <summary>
    /// Clone an entire slide with all its content, relationships (images, charts, media),
    /// layout link, background, notes, and transitions.
    /// Pattern follows POI's createSlide(layout) + importContent(srcSlide).
    /// </summary>
    private string CloneSlide(Match slideMatch, List<SlidePart> slideParts, int? index)
    {
        var srcSlideIdx = int.Parse(slideMatch.Groups[1].Value);
        if (srcSlideIdx < 1 || srcSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {srcSlideIdx} not found (total: {slideParts.Count})");

        var srcSlidePart = slideParts[srcSlideIdx - 1];
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var presentation = presentationPart.Presentation
            ?? throw new InvalidOperationException("No presentation");

        // 1. Create new SlidePart
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // 2. Copy slide layout relationship (link to same layout as source)
        var srcLayoutPart = srcSlidePart.SlideLayoutPart;
        if (srcLayoutPart != null)
            newSlidePart.AddPart(srcLayoutPart);

        // 3. Deep-clone the Slide XML
        var srcSlide = GetSlide(srcSlidePart);
        newSlidePart.Slide = (Slide)srcSlide.CloneNode(true);

        // 4. Copy all referenced parts (images, charts, embedded objects, media)
        CopySlideParts(srcSlidePart, newSlidePart);

        // 5. Copy notes slide if present
        if (srcSlidePart.NotesSlidePart != null)
        {
            var srcNotesPart = srcSlidePart.NotesSlidePart;
            var newNotesPart = newSlidePart.AddNewPart<NotesSlidePart>();
            newNotesPart.NotesSlide = srcNotesPart.NotesSlide != null
                ? (NotesSlide)srcNotesPart.NotesSlide.CloneNode(true)
                : new NotesSlide();
            // Link notes to the new slide
            newNotesPart.AddPart(newSlidePart);
        }

        newSlidePart.Slide.Save();

        // 6. Register in SlideIdList at the correct position
        var slideIdList = presentation.GetFirstChild<SlideIdList>()
            ?? presentation.AppendChild(new SlideIdList());
        var maxId = slideIdList.Elements<SlideId>().Any()
            ? slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255) + 1
            : 256;
        var relId = presentationPart.GetIdOfPart(newSlidePart);
        var newSlideId = new SlideId { Id = maxId, RelationshipId = relId };

        if (index.HasValue && index.Value < slideIdList.Elements<SlideId>().Count())
        {
            var refSlide = slideIdList.Elements<SlideId>().ElementAtOrDefault(index.Value);
            if (refSlide != null)
                slideIdList.InsertBefore(newSlideId, refSlide);
            else
                slideIdList.AppendChild(newSlideId);
        }
        else
        {
            slideIdList.AppendChild(newSlideId);
        }

        presentation.Save();

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        var insertedIdx = slideIds.FindIndex(s => s.RelationshipId?.Value == relId) + 1;
        return $"/slide[{insertedIdx}]";
    }

    /// <summary>
    /// Copy all sub-parts (images, charts, media, etc.) from source to target slide,
    /// remapping relationship IDs in the cloned XML.
    /// </summary>
    private static void CopySlideParts(SlidePart source, SlidePart target)
    {
        // Build a map of old rId → new rId for all parts that need copying
        var rIdMap = new Dictionary<string, string>();

        foreach (var part in source.Parts)
        {
            // Skip SlideLayoutPart (already linked above)
            if (part.OpenXmlPart is SlideLayoutPart) continue;
            // Skip NotesSlidePart (handled separately)
            if (part.OpenXmlPart is NotesSlidePart) continue;

            try
            {
                // Try to add the same part (shares the underlying data)
                var newRelId = target.CreateRelationshipToPart(part.OpenXmlPart);
                if (newRelId != part.RelationshipId)
                    rIdMap[part.RelationshipId] = newRelId;
            }
            catch
            {
                // If sharing fails, deep-copy the part data
                try
                {
                    var newPart = target.AddNewPart<OpenXmlPart>(part.OpenXmlPart.ContentType, part.RelationshipId);
                    using var stream = part.OpenXmlPart.GetStream();
                    newPart.FeedData(stream);
                }
                catch { /* Best effort — some parts may not be copyable */ }
            }
        }

        // Also copy external relationships (hyperlinks, media links)
        foreach (var extRel in source.ExternalRelationships)
        {
            try
            {
                target.AddExternalRelationship(extRel.RelationshipType, extRel.Uri, extRel.Id);
            }
            catch { }
        }
        foreach (var hyperRel in source.HyperlinkRelationships)
        {
            try
            {
                target.AddHyperlinkRelationship(hyperRel.Uri, hyperRel.IsExternal, hyperRel.Id);
            }
            catch { }
        }

        // Remap any changed relationship IDs in the slide XML
        if (rIdMap.Count > 0 && target.Slide != null)
        {
            RemapRelationshipIds(target.Slide, rIdMap);
            target.Slide.Save();
        }
    }

    /// <summary>
    /// Update all r:id references in the XML tree when relationship IDs changed during copy.
    /// </summary>
    private static void RemapRelationshipIds(OpenXmlElement root, Dictionary<string, string> rIdMap)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        foreach (var el in root.Descendants().Prepend(root).ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri || attr.Value == null) continue;
                if (rIdMap.TryGetValue(attr.Value, out var newId))
                {
                    el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newId));
                }
            }
        }
    }

    private (SlidePart slidePart, OpenXmlElement element) ResolveSlideElement(string path, List<SlidePart> slideParts)
    {
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\]$");
        if (!match.Success)
            throw new ArgumentException($"Invalid element path: {path}. Expected /slide[N]/element[M]");

        var slideIdx = int.Parse(match.Groups[1].Value);
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);

        OpenXmlElement element = elementType switch
        {
            "shape" => shapeTree.Elements<Shape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Shape {elementIdx} not found"),
            "picture" or "pic" => shapeTree.Elements<Picture>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Picture {elementIdx} not found"),
            "connector" or "connection" => shapeTree.Elements<ConnectionShape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Connector {elementIdx} not found"),
            "table" => shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Table {elementIdx} not found"),
            "chart" => shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()).ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Chart {elementIdx} not found"),
            "group" => shapeTree.Elements<GroupShape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Group {elementIdx} not found"),
            _ => shapeTree.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase))
                .ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"{elementType} {elementIdx} not found")
        };

        return (slidePart, element);
    }

    private static void CopyRelationships(OpenXmlElement element, SlidePart sourcePart, SlidePart targetPart)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var allElements = element.Descendants().Prepend(element);

        foreach (var el in allElements.ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri) continue;

                var oldRelId = attr.Value;
                if (string.IsNullOrEmpty(oldRelId)) continue;

                try
                {
                    var referencedPart = sourcePart.GetPartById(oldRelId);
                    string newRelId;
                    try
                    {
                        newRelId = targetPart.GetIdOfPart(referencedPart);
                    }
                    catch (ArgumentException)
                    {
                        newRelId = targetPart.CreateRelationshipToPart(referencedPart);
                    }

                    if (newRelId != oldRelId)
                    {
                        el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRelId));
                    }
                }
                catch (ArgumentOutOfRangeException) { /* Not a valid relationship ID, skip */ }
            }
        }
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue && parent is ShapeTree)
        {
            // Skip structural elements (nvGrpSpPr, grpSpPr) that must stay at the beginning
            var contentChildren = parent.ChildElements
                .Where(e => e is not NonVisualGroupShapeProperties && e is not GroupShapeProperties)
                .ToList();
            if (index.Value >= 0 && index.Value < contentChildren.Count)
                contentChildren[index.Value].InsertBeforeSelf(element);
            else if (contentChildren.Count > 0)
                contentChildren.Last().InsertAfterSelf(element);
            else
                parent.AppendChild(element);
        }
        else if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }

    private static string ComputeElementPath(string parentPath, OpenXmlElement element, ShapeTree shapeTree)
    {
        // Map back to semantic type names
        string typeName;
        int typeIdx;
        if (element is Shape)
        {
            typeName = "shape";
            typeIdx = shapeTree.Elements<Shape>().ToList().IndexOf((Shape)element) + 1;
        }
        else if (element is Picture)
        {
            typeName = "picture";
            typeIdx = shapeTree.Elements<Picture>().ToList().IndexOf((Picture)element) + 1;
        }
        else if (element is ConnectionShape)
        {
            typeName = "connector";
            typeIdx = shapeTree.Elements<ConnectionShape>().ToList().IndexOf((ConnectionShape)element) + 1;
        }
        else if (element is GroupShape)
        {
            typeName = "group";
            typeIdx = shapeTree.Elements<GroupShape>().ToList().IndexOf((GroupShape)element) + 1;
        }
        else if (element is GraphicFrame gf)
        {
            if (gf.Descendants<Drawing.Table>().Any())
            {
                typeName = "table";
                typeIdx = shapeTree.Elements<GraphicFrame>()
                    .Where(f => f.Descendants<Drawing.Table>().Any())
                    .ToList().IndexOf(gf) + 1;
            }
            else if (gf.Descendants<C.ChartReference>().Any())
            {
                typeName = "chart";
                typeIdx = shapeTree.Elements<GraphicFrame>()
                    .Where(f => f.Descendants<C.ChartReference>().Any())
                    .ToList().IndexOf(gf) + 1;
            }
            else
            {
                typeName = element.LocalName;
                typeIdx = shapeTree.ChildElements
                    .Where(e => e.LocalName == element.LocalName)
                    .ToList().IndexOf(element) + 1;
            }
        }
        else
        {
            typeName = element.LocalName;
            typeIdx = shapeTree.ChildElements
                .Where(e => e.LocalName == element.LocalName)
                .ToList().IndexOf(element) + 1;
        }
        return $"{parentPath}/{typeName}[{typeIdx}]";
    }
}
