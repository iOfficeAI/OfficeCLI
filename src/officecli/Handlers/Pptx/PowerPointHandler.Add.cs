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

                // Add title shape if text provided
                if (properties.TryGetValue("title", out var titleText))
                {
                    var titleShape = CreateTextShape(1, "Title", titleText, true);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(titleShape);
                }

                // Add content text if provided
                if (properties.TryGetValue("text", out var contentText))
                {
                    var textShape = CreateTextShape(2, "Content", contentText, false);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(textShape);
                }

                // Apply background if provided
                if (properties.TryGetValue("background", out var bgValue))
                    ApplySlideBackground(newSlidePart, bgValue);

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
                var slideCount = slideIdList.Elements<SlideId>().Count();
                return $"/slide[{slideCount}]";

            case "shape" or "textbox":
                var slideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException($"Shapes must be added to a slide: /slide[N]");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide {slideIdx} not found");

                var slidePart = slideParts[slideIdx - 1];
                var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var text = properties.GetValueOrDefault("text", "");
                var shapeName = properties.GetValueOrDefault("name", $"TextBox {shapeTree.Elements<Shape>().Count() + 1}");
                var shapeId = (uint)(shapeTree.Elements<Shape>().Count() + shapeTree.Elements<Picture>().Count() + 2);

                var newShape = CreateTextShape(shapeId, shapeName, text, false);

                if (properties.TryGetValue("font", out var font))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Append(new Drawing.LatinFont { Typeface = font });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = font });
                    }
                }
                if (properties.TryGetValue("size", out var sizeStr))
                {
                    var sizeVal = int.Parse(sizeStr) * 100;
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                }
                if (properties.TryGetValue("bold", out var boldStr))
                {
                    var isBold = bool.Parse(boldStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                }
                if (properties.TryGetValue("italic", out var italicStr))
                {
                    var isItalic = bool.Parse(italicStr);
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
                        var solidFill = new Drawing.SolidFill();
                        solidFill.Append(new Drawing.RgbColorModelHex { Val = colorVal.ToUpperInvariant() });
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
                            "false" or "none" => Drawing.TextUnderlineValues.None,
                            _ => Drawing.TextUnderlineValues.Single
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
                            _ => Drawing.TextStrikeValues.SingleStrike
                        };
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
                        xfrm.Rotation = (int)(double.Parse(rotVal) * 60000);
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

                // Line/border (after fill per schema: xfrm → prstGeom → fill → ln)
                if (properties.TryGetValue("line", out var lineColor) || properties.TryGetValue("linecolor", out lineColor) || properties.TryGetValue("line.color", out lineColor))
                {
                    var outline = newShape.ShapeProperties!.GetFirstChild<Drawing.Outline>() ?? newShape.ShapeProperties.AppendChild(new Drawing.Outline());
                    if (lineColor.Equals("none", StringComparison.OrdinalIgnoreCase))
                        outline.AppendChild(new Drawing.NoFill());
                    else
                        outline.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = lineColor.TrimStart('#').ToUpperInvariant() }));
                }
                if (properties.TryGetValue("linewidth", out var lwStr) || properties.TryGetValue("line.width", out lwStr))
                {
                    var outline = newShape.ShapeProperties!.GetFirstChild<Drawing.Outline>() ?? newShape.ShapeProperties.AppendChild(new Drawing.Outline());
                    outline.Width = (int)ParseEmu(lwStr);
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

                // lineDash, effects (shadow/glow/reflection) — delegate to SetRunOrShapeProperties
                var effectKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "linedash", "line.dash", "shadow", "glow", "reflection" };
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
                if (!properties.TryGetValue("path", out var imgPath))
                    throw new ArgumentException("'path' property is required for picture type");
                if (!File.Exists(imgPath))
                    throw new FileNotFoundException($"Image file not found: {imgPath}");

                var imgSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!imgSlideMatch.Success)
                    throw new ArgumentException($"Pictures must be added to a slide: /slide[N]");

                var imgSlideIdx = int.Parse(imgSlideMatch.Groups[1].Value);
                var imgSlideParts = GetSlideParts().ToList();
                if (imgSlideIdx < 1 || imgSlideIdx > imgSlideParts.Count)
                    throw new ArgumentException($"Slide {imgSlideIdx} not found");

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
                    throw new ArgumentException($"Slide {chartSlideIdx} not found");

                var chartSlidePart = chartSlideParts[chartSlideIdx - 1];
                var chartShapeTree = GetSlide(chartSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                // Parse chart data
                var chartType = properties.GetValueOrDefault("charttype",
                    properties.GetValueOrDefault("type", "column"));
                var chartTitle = properties.GetValueOrDefault("title");
                var categories = ParseCategories(properties);
                var seriesData = ParseSeriesData(properties);

                if (seriesData.Count == 0)
                    throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                        "or series1=\"Revenue:100,200,300\"");

                // Create ChartPart and build chart
                var chartPart = chartSlidePart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
                chartPart.ChartSpace.Save();

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
                    throw new ArgumentException($"Slide {tblSlideIdx} not found");

                var tblSlidePart = tblSlideParts[tblSlideIdx - 1];
                var tblShapeTree = GetSlide(tblSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                int rows = int.Parse(properties.GetValueOrDefault("rows", "3"));
                int cols = int.Parse(properties.GetValueOrDefault("cols", "3"));
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
                            new Drawing.Paragraph(new Drawing.EndParagraphRunProperties { Language = "zh-CN" })
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
                    throw new ArgumentException($"Slide {eqSlideIdx} not found");

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
                    throw new ArgumentException($"Slide {notesSlideIdx} not found");
                var notesSlidePart = EnsureNotesSlidePart(notesSlideParts[notesSlideIdx - 1]);
                if (properties.TryGetValue("text", out var notesText))
                    SetNotesText(notesSlidePart, notesText);
                return $"/slide[{notesSlideIdx}]/notes";
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
                        throw new ArgumentException($"Slide {fbSlideIdx} not found");

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
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slideId = slideIds[slideIdx - 1];
            var relId = slideId.RelationshipId?.Value;
            slideId.Remove();
            if (relId != null)
                presentationPart.DeletePart(presentationPart.GetPartById(relId));
            presentation.Save();
            return;
        }

        // Remove shape or picture from slide
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

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
            shapes[elementIdx - 1].Remove();
        }
        else if (elementType is "picture" or "pic")
        {
            var pics = shapeTree.Elements<Picture>().ToList();
            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"Picture {elementIdx} not found");
            pics[elementIdx - 1].Remove();
        }
        else
        {
            throw new ArgumentException($"Unknown element type: {elementType}");
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
                throw new ArgumentException($"Slide {slideIdx} not found");

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
                throw new ArgumentException($"Slide {tgtSlideIdx} not found");
            tgtSlidePart = slideParts[tgtSlideIdx - 1];
            tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
        }

        srcElement.Remove();

        // Copy relationships if moving across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(srcElement, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, srcElement, index);

        GetSlide(srcSlidePart).Save();
        if (srcSlidePart != tgtSlidePart)
            GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(effectiveParentPath, srcElement, tgtShapeTree);
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var slideParts = GetSlideParts().ToList();

        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);
        var clone = srcElement.CloneNode(true);

        var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
        if (!tgtSlideMatch.Success)
            throw new ArgumentException($"Target must be a slide: /slide[N]");
        var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
        if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {tgtSlideIdx} not found");

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

    private (SlidePart slidePart, OpenXmlElement element) ResolveSlideElement(string path, List<SlidePart> slideParts)
    {
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\]$");
        if (!match.Success)
            throw new ArgumentException($"Invalid element path: {path}. Expected /slide[N]/element[M]");

        var slideIdx = int.Parse(match.Groups[1].Value);
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

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
                    catch
                    {
                        newRelId = targetPart.CreateRelationshipToPart(referencedPart);
                    }

                    if (newRelId != oldRelId)
                    {
                        el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRelId));
                    }
                }
                catch { /* Not a valid relationship, skip */ }
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
