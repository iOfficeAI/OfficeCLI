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
    private string AddShape(string parentPath, int? index, Dictionary<string, string> properties)
    {
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
                    var (titleSlideW, _) = GetSlideSize();
                    long cxEmu = titleSlideW, cyEmu = 1800000; // default: slide width x 5cm
                    if (properties.TryGetValue("x", out var xStr) || properties.TryGetValue("left", out xStr)) xEmu = ParseEmu(xStr);
                    if (properties.TryGetValue("y", out var yStr) || properties.TryGetValue("top", out yStr)) yEmu = ParseEmu(yStr);
                    if (properties.TryGetValue("width", out var wStr) || properties.TryGetValue("w", out wStr))
                    {
                        cxEmu = ParseEmu(wStr);
                        if (cxEmu < 0) throw new ArgumentException($"Negative width is not allowed: '{wStr}'.");
                    }
                    if (properties.TryGetValue("height", out var hStr) || properties.TryGetValue("h", out hStr))
                    {
                        cyEmu = ParseEmu(hStr);
                        if (cyEmu < 0) throw new ArgumentException($"Negative height is not allowed: '{hStr}'.");
                    }

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
                if (properties.TryGetValue("line", out var lineColor) || properties.TryGetValue("linecolor", out lineColor) || properties.TryGetValue("lineColor", out lineColor) || properties.TryGetValue("line.color", out lineColor) || properties.TryGetValue("border", out lineColor) || properties.TryGetValue("border.color", out lineColor))
                {
                    var outline = EnsureOutline(newShape.ShapeProperties!);
                    if (lineColor.Equals("none", StringComparison.OrdinalIgnoreCase))
                        outline.AppendChild(new Drawing.NoFill());
                    else
                        outline.AppendChild(BuildSolidFill(lineColor));
                }
                if (properties.TryGetValue("linewidth", out var lwStr) || properties.TryGetValue("lineWidth", out lwStr) || properties.TryGetValue("line.width", out lwStr) || properties.TryGetValue("border.width", out lwStr))
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
                      "lineopacity", "line.opacity",
                      "image", "imagefill" };
                var effectProps = properties
                    .Where(kv => effectKeys.Contains(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (effectProps.Count > 0)
                    SetRunOrShapeProperties(effectProps, GetAllRuns(newShape), newShape, slidePart);

                // Animation
                if (properties.TryGetValue("animation", out var animVal) ||
                    properties.TryGetValue("animate", out animVal))
                    ApplyShapeAnimation(slidePart, newShape, animVal);

                GetSlide(slidePart).Save();
                var shapeCount = shapeTree.Elements<Shape>().Count();
                return $"/slide[{slideIdx}]/shape[{shapeCount}]";
    }


}
