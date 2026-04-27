// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

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
                var shapeId = GenerateUniqueShapeId(shapeTree);
                var shapeName = properties.GetValueOrDefault("name", $"TextBox {shapeTree.Elements<Shape>().Count() + 1}");

                // Auto-add !! prefix if the slide (or the next slide) has a morph transition
                if (!shapeName.StartsWith("!!") && !shapeName.StartsWith("TextBox ") && !shapeName.StartsWith("Content ") && shapeName != "")
                {
                    if (SlideHasMorphContext(slidePart, slideParts))
                        shapeName = "!!" + shapeName;
                }

                var newShape = CreateTextShape(shapeId, shapeName, text, false);

                // CONSISTENCY(font-dotted-alias): mirror Set's font.<attr> aliases
                // (commit 80fb739e). Without these, `add shape --prop font.name=Arial`
                // silently dropped while `set --prop font.name=Arial` succeeded.
                if (properties.TryGetValue("size", out var sizeStr)
                    || properties.TryGetValue("fontSize", out sizeStr)
                    || properties.TryGetValue("fontsize", out sizeStr)
                    || properties.TryGetValue("font.size", out sizeStr))
                {
                    var sizeVal = (int)Math.Round(ParseFontSize(sizeStr) * 100);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                }
                if (properties.TryGetValue("bold", out var boldStr)
                    || properties.TryGetValue("font.bold", out boldStr))
                {
                    var isBold = IsTruthy(boldStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                }
                if (properties.TryGetValue("italic", out var italicStr)
                    || properties.TryGetValue("font.italic", out italicStr))
                {
                    var isItalic = IsTruthy(italicStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                }
                if (properties.TryGetValue("color", out var colorVal)
                    || properties.TryGetValue("font.color", out colorVal))
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
                if (properties.TryGetValue("font", out var font)
                    || properties.TryGetValue("font.name", out font))
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
                if (properties.TryGetValue("underline", out var ulVal)
                    || properties.TryGetValue("font.underline", out ulVal))
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
                if (properties.TryGetValue("strikethrough", out var stVal)
                    || properties.TryGetValue("strike", out stVal)
                    || properties.TryGetValue("font.strike", out stVal)
                    || properties.TryGetValue("font.strikethrough", out stVal))
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

                // Caps (allCaps / smallCaps / cap=all|small|none)
                // CONSISTENCY(allcaps-alias): mirror Word commit ccaed17a;
                // accept allCaps/allcaps/smallCaps/smallcaps as run-level rPr cap.
                {
                    string? capValue = null;
                    if (properties.TryGetValue("cap", out var rawCap)) capValue = rawCap;
                    else if (properties.TryGetValue("allCaps", out var allCaps)
                          || properties.TryGetValue("allcaps", out allCaps))
                        capValue = (allCaps is "0" or "false" or "False" or "none") ? "none" : "all";
                    else if (properties.TryGetValue("smallCaps", out var smallCaps)
                          || properties.TryGetValue("smallcaps", out smallCaps))
                        capValue = (smallCaps is "0" or "false" or "False" or "none") ? "none" : "small";

                    if (capValue != null)
                    {
                        foreach (var run in newShape.Descendants<Drawing.Run>())
                        {
                            var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rProps.SetAttribute(new OpenXmlAttribute("", "cap", "", capValue));
                        }
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
                    long cxEmu = 3600000, cyEmu = 1800000; // default: 10cm x 5cm (avoid full-slide overlap when width unspecified)
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

                    var presetName = properties.TryGetValue("preset", out var pn) ? pn
                        : properties.TryGetValue("geometry", out pn) ? pn
                        : properties.GetValueOrDefault("shape", "rect");
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

                // Pattern fill (mutually exclusive with fill/gradient — last one wins, following fill/gradient convention)
                if (properties.TryGetValue("pattern", out var patternVal))
                {
                    ApplyPatternFill(newShape.ShapeProperties!, patternVal);
                }

                // Opacity (alpha on fill) — like POI XSLFColor uses <a:alpha val="N"/>
                // Must come after gradient so it can apply to gradient stops too.
                // Alpha must attach to a color element inside a fill carrier; if
                // the caller gave 'opacity' without any fill/gradient/pattern,
                // the value has nothing to bind to. Per schemas/help/pptx/shape.json
                // 'opacity.requires: ["fill"]', reject rather than silently drop.
                if (properties.TryGetValue("opacity", out var opacityVal))
                {
                    var hasFillCarrier =
                        properties.ContainsKey("fill") ||
                        properties.ContainsKey("gradient") ||
                        properties.ContainsKey("pattern") ||
                        (newShape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>() != null) ||
                        (newShape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null) ||
                        (newShape.ShapeProperties?.GetFirstChild<Drawing.PatternFill>() != null);
                    if (!hasFillCarrier)
                        throw new ArgumentException(
                            $"'opacity'='{opacityVal}' requires a fill carrier. Provide one of 'fill' / 'gradient' / 'pattern' " +
                            "so the alpha value has a color element to attach to.");
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

                InsertAtPosition(shapeTree, newShape, index);

                // Hyperlink on shape
                if (properties.TryGetValue("link", out var linkVal))
                {
                    var tooltipVal = properties.GetValueOrDefault("tooltip");
                    ApplyShapeHyperlink(slidePart, newShape, linkVal, tooltipVal);
                }

                // lineDash, effects, 3D, flip — delegate to SetRunOrShapeProperties
                var effectKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "linedash", "line.dash", "shadow", "glow", "reflection",
                      "softedge", "blur", "fliph", "flipv", "rot3d", "rotation3d",
                      "rotx", "roty", "rotz", "bevel", "beveltop", "bevelbottom",
                      "depth", "extrusion", "material", "lighting", "lightrig",
                      "spacing", "charspacing", "letterspacing",
                      "indent", "marginleft", "marl", "marginright", "marr",
                      "textfill", "textgradient", "geometry",
                      "baseline", "superscript", "subscript",
                      "textwarp", "wordart", "autofit",
                      "lineopacity", "line.opacity",
                      "image", "imagefill",
                      // CONSISTENCY(rpr-attr-fallback / R21-fuzzer-1+2): drawingML
                      // run-property attributes must reach SetRunOrShapeProperties
                      // so the long-tail rPr-attribute branch routes them to the
                      // first run instead of dropping them on the <p:sp> element.
                      "lang", "altLang", "altlang", "spc", "kern", "cap",
                      "kumimoji", "normalizeH", "normalizeh", "noProof", "noproof",
                      "dirty", "smtClean", "smtclean", "smtId", "smtid", "err" };
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
                return $"/slide[{slideIdx}]/{BuildElementPathSegment("shape", newShape, shapeTree.Elements<Shape>().Count())}";
    }


}
