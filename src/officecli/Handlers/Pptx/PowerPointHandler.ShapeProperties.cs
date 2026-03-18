// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private static List<Drawing.Run> GetAllRuns(Shape shape)
    {
        return shape.TextBody?.Elements<Drawing.Paragraph>()
            .SelectMany(p => p.Elements<Drawing.Run>()).ToList()
            ?? new List<Drawing.Run>();
    }

    private static List<string> SetRunOrShapeProperties(
        Dictionary<string, string> properties, List<Drawing.Run> runs, Shape shape, OpenXmlPart? part = null)
    {
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                {
                    var textLines = value.Replace("\\n", "\n").Split('\n');
                    if (runs.Count == 1 && textLines.Length == 1)
                    {
                        // Single run, single line: just replace its text
                        runs[0].Text = new Drawing.Text(textLines[0]);
                    }
                    else
                    {
                        // Shape-level: replace all text, preserve first run formatting
                        var textBody = shape.TextBody;
                        if (textBody != null)
                        {
                            var firstRun = textBody.Descendants<Drawing.Run>().FirstOrDefault();
                            var runProps = firstRun?.RunProperties?.CloneNode(true) as Drawing.RunProperties;

                            textBody.RemoveAllChildren<Drawing.Paragraph>();

                            foreach (var textLine in textLines)
                            {
                                var newPara = new Drawing.Paragraph();
                                var newRun = new Drawing.Run();
                                if (runProps != null)
                                    newRun.RunProperties = runProps.CloneNode(true) as Drawing.RunProperties;
                                newRun.Text = new Drawing.Text(textLine);
                                newPara.Append(newRun);
                                textBody.Append(newPara);
                            }
                        }
                    }
                    break;
                }

                case "font":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                    }
                    break;

                case "size":
                    var sizeVal = (int)Math.Round(ParseFontSize(value) * 100);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                    break;

                case "bold":
                    var isBold = IsTruthy(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                    break;

                case "italic":
                    var isItalic = IsTruthy(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                    break;

                case "color":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var colorFill = BuildSolidFill(value);
                        if (rProps is OpenXmlCompositeElement composite)
                        {
                            if (!composite.AddChild(colorFill, throwOnError: false))
                                rProps.AppendChild(colorFill);
                        }
                        else
                        {
                            rProps.AppendChild(colorFill);
                        }
                    }
                    break;

                case "underline":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Underline = value.ToLowerInvariant() switch
                        {
                            "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                            "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                            "heavy" => Drawing.TextUnderlineValues.Heavy,
                            "dotted" => Drawing.TextUnderlineValues.Dotted,
                            "dash" => Drawing.TextUnderlineValues.Dash,
                            "wavy" => Drawing.TextUnderlineValues.Wavy,
                            "false" or "none" => Drawing.TextUnderlineValues.None,
                            _ => Drawing.TextUnderlineValues.Single
                        };
                    }
                    break;

                case "strikethrough" or "strike":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Strike = value.ToLowerInvariant() switch
                        {
                            "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                            "double" => Drawing.TextStrikeValues.DoubleStrike,
                            "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                            _ => Drawing.TextStrikeValues.SingleStrike
                        };
                    }
                    break;

                case "fill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyShapeFill(spPr, value);
                    break;
                }

                case "gradient":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyGradientFill(spPr, value);
                    break;
                }

                case "liststyle" or "list":
                {
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        ApplyListStyle(pProps, value);
                    }
                    break;
                }

                case "margin" or "inset":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    ApplyTextMargin(bodyPr, value);
                    break;
                }

                case "align":
                {
                    var alignment = ParseTextAlignment(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = alignment;
                    }
                    break;
                }

                case "valign":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.Anchor = value.ToLowerInvariant() switch
                    {
                        "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                        "center" or "middle" or "c" or "m" => Drawing.TextAnchoringTypeValues.Center,
                        "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                        _ => throw new ArgumentException($"Invalid valign: {value}. Use top/center/bottom")
                    };
                    break;
                }

                case "preset":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var existingGeom = spPr.GetFirstChild<Drawing.PresetGeometry>();
                    if (existingGeom != null)
                        existingGeom.Preset = ParsePresetShape(value);
                    else
                        spPr.AppendChild(new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = ParsePresetShape(value) });
                    break;
                }

                case "line" or "linecolor" or "line.color":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = spPr.GetFirstChild<Drawing.Outline>() ?? spPr.AppendChild(new Drawing.Outline());
                    outline.RemoveAllChildren<Drawing.SolidFill>();
                    outline.RemoveAllChildren<Drawing.NoFill>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        outline.AppendChild(new Drawing.NoFill());
                    else
                        outline.AppendChild(BuildSolidFill(value));
                    break;
                }

                case "linewidth" or "line.width":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = spPr.GetFirstChild<Drawing.Outline>() ?? spPr.AppendChild(new Drawing.Outline());
                    outline.Width = Core.EmuConverter.ParseEmuAsInt(value);
                    break;
                }

                case "linedash" or "line.dash":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = spPr.GetFirstChild<Drawing.Outline>() ?? spPr.AppendChild(new Drawing.Outline());
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    outline.AppendChild(new Drawing.PresetDash { Val = value.ToLowerInvariant() switch
                    {
                        "solid" => Drawing.PresetLineDashValues.Solid,
                        "dot" => Drawing.PresetLineDashValues.Dot,
                        "dash" => Drawing.PresetLineDashValues.Dash,
                        "dashdot" or "dash_dot" => Drawing.PresetLineDashValues.DashDot,
                        "longdash" or "lgdash" or "lg_dash" => Drawing.PresetLineDashValues.LargeDash,
                        "longdashdot" or "lgdashdot" or "lg_dash_dot" => Drawing.PresetLineDashValues.LargeDashDot,
                        _ => Drawing.PresetLineDashValues.Solid
                    }});
                    break;
                }

                case "lineopacity" or "line.opacity":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = spPr.GetFirstChild<Drawing.Outline>() ?? spPr.AppendChild(new Drawing.Outline());
                    var solidFillLn = outline.GetFirstChild<Drawing.SolidFill>();
                    if (solidFillLn != null)
                    {
                        var color = solidFillLn.GetFirstChild<Drawing.RgbColorModelHex>();
                        if (color != null)
                        {
                            color.RemoveAllChildren<Drawing.Alpha>();
                            var pct = (int)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 100000); // 0.0-1.0 → 0-100000
                            color.AppendChild(new Drawing.Alpha { Val = pct });
                        }
                    }
                    break;
                }

                case "rotation" or "rotate":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    xfrm.Rotation = (int)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 60000); // degrees to 60000ths
                    break;
                }

                case "opacity":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var solidFill = spPr.GetFirstChild<Drawing.SolidFill>();
                    if (solidFill != null)
                    {
                        var color = solidFill.GetFirstChild<Drawing.RgbColorModelHex>();
                        if (color != null)
                        {
                            color.RemoveAllChildren<Drawing.Alpha>();
                            var pct = (int)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 100000); // 0.0-1.0 → 0-100000
                            color.AppendChild(new Drawing.Alpha { Val = pct });
                        }
                    }
                    break;
                }

                case "image" or "imagefill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null || part is not SlidePart slidePart) { unsupported.Add(key); break; }
                    ApplyShapeImageFill(spPr, value, slidePart);
                    break;
                }

                case "linespacing" or "line.spacing":
                {
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.LineSpacing>();
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPercent { Val = (int)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 100000) })); // e.g. 1.5 → 150000 (150%)
                    }
                    break;
                }

                case "spacebefore" or "space.before":
                {
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                        pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = (int)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 100) })); // pt
                    }
                    break;
                }

                case "spaceafter" or "space.after":
                {
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                        pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = (int)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 100) })); // pt
                    }
                    break;
                }

                case "textwarp" or "wordart":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.RemoveAllChildren<Drawing.PresetTextWarp>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var warpName = value.StartsWith("text") ? value : $"text{char.ToUpper(value[0])}{value[1..]}";
                        bodyPr.AppendChild(new Drawing.PresetTextWarp(
                            new Drawing.AdjustValueList()
                        ) { Preset = new Drawing.TextShapeValues(warpName) });
                    }
                    break;
                }

                case "autofit":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.RemoveAllChildren<Drawing.NormalAutoFit>();
                    bodyPr.RemoveAllChildren<Drawing.ShapeAutoFit>();
                    bodyPr.RemoveAllChildren<Drawing.NoAutoFit>();
                    switch (value.ToLowerInvariant())
                    {
                        case "true" or "normal": bodyPr.AppendChild(new Drawing.NormalAutoFit()); break;
                        case "shape": bodyPr.AppendChild(new Drawing.ShapeAutoFit()); break;
                        case "false" or "none": bodyPr.AppendChild(new Drawing.NoAutoFit()); break;
                    }
                    break;
                }

                case "x" or "y" or "width" or "height":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    var offset = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                    var extents = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                    var emu = ParseEmu(value);
                    switch (key.ToLowerInvariant())
                    {
                        case "x": offset.X = emu; break;
                        case "y": offset.Y = emu; break;
                        case "width": extents.Cx = emu; break;
                        case "height": extents.Cy = emu; break;
                    }
                    break;
                }

                case "shadow":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyShadow(spPr, value);
                    break;
                }

                case "glow":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyGlow(spPr, value);
                    break;
                }

                case "reflection":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyReflection(spPr, value);
                    break;
                }

                default:
                    if (!GenericXmlQuery.SetGenericAttribute(shape, key, value))
                        unsupported.Add(key);
                    break;
            }
        }

        return unsupported;
    }

    private static List<string> SetTableCellProperties(Drawing.TableCell cell, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                {
                    var textBody = cell.TextBody;
                    var lines = value.Replace("\\n", "\n").Split('\n');
                    if (textBody == null)
                    {
                        textBody = new Drawing.TextBody(
                            new Drawing.BodyProperties(), new Drawing.ListStyle());
                        foreach (var line in lines)
                        {
                            textBody.AppendChild(new Drawing.Paragraph(new Drawing.Run(
                                new Drawing.RunProperties { Language = "en-US" },
                                new Drawing.Text(line))));
                        }
                        cell.PrependChild(textBody);
                    }
                    else
                    {
                        var firstRun = textBody.Descendants<Drawing.Run>().FirstOrDefault();
                        var runProps = firstRun?.RunProperties?.CloneNode(true) as Drawing.RunProperties;
                        textBody.RemoveAllChildren<Drawing.Paragraph>();
                        foreach (var line in lines)
                        {
                            var newRun = new Drawing.Run();
                            if (runProps != null) newRun.RunProperties = runProps.CloneNode(true) as Drawing.RunProperties;
                            else newRun.RunProperties = new Drawing.RunProperties { Language = "en-US" };
                            newRun.Text = new Drawing.Text(line);
                            textBody.Append(new Drawing.Paragraph(newRun));
                        }
                    }
                    break;
                }
                case "font":
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                    }
                    break;
                case "size":
                    var sz = (int)Math.Round(ParseFontSize(value) * 100);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sz;
                    }
                    break;
                case "bold":
                    var b = IsTruthy(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = b;
                    }
                    break;
                case "italic":
                    var it = IsTruthy(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = it;
                    }
                    break;
                case "color":
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        rProps.AppendChild(BuildSolidFill(value));
                    }
                    break;
                case "fill":
                {
                    var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                    if (tcPr == null)
                    {
                        tcPr = new Drawing.TableCellProperties();
                        cell.Append(tcPr);
                    }
                    tcPr.RemoveAllChildren<Drawing.SolidFill>();
                    tcPr.RemoveAllChildren<Drawing.NoFill>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        tcPr.Append(new Drawing.NoFill());
                    }
                    else
                    {
                        tcPr.Append(BuildSolidFill(value));
                    }
                    break;
                }
                case "align" or "alignment":
                {
                    var para = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                    if (para != null)
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = ParseTextAlignment(value);
                    }
                    break;
                }
                case "valign":
                {
                    var tcPrV = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    tcPrV.Anchor = value.ToLowerInvariant() switch
                    {
                        "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                        "middle" or "center" or "ctr" => Drawing.TextAnchoringTypeValues.Center,
                        "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                        _ => Drawing.TextAnchoringTypeValues.Top
                    };
                    break;
                }
                case "gridspan" or "colspan":
                    cell.GridSpan = new DocumentFormat.OpenXml.Int32Value(int.Parse(value));
                    break;
                case "rowspan":
                    cell.RowSpan = new DocumentFormat.OpenXml.Int32Value(int.Parse(value));
                    break;
                case "vmerge":
                    cell.VerticalMerge = new DocumentFormat.OpenXml.BooleanValue(IsTruthy(value));
                    break;
                case "hmerge":
                    cell.HorizontalMerge = new DocumentFormat.OpenXml.BooleanValue(IsTruthy(value));
                    break;
                case "underline":
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Underline = value.Equals("none", StringComparison.OrdinalIgnoreCase)
                            ? Drawing.TextUnderlineValues.None
                            : new Drawing.TextUnderlineValues(value);
                    }
                    break;
                case "strikethrough" or "strike":
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Strike = value.Equals("none", StringComparison.OrdinalIgnoreCase) || !IsTruthy(value)
                            ? Drawing.TextStrikeValues.NoStrike
                            : new Drawing.TextStrikeValues(value == "true" ? "sngStrike" : value);
                    }
                    break;
                case var k when k.StartsWith("border"):
                {
                    var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                    if (tcPr == null)
                    {
                        tcPr = new Drawing.TableCellProperties();
                        cell.Append(tcPr);
                    }

                    // Parse value: "FF0000", "1pt solid FF0000", "2pt dash 0000FF"
                    var borderParts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    string? borderColor = null;
                    long? borderWidth = null;
                    string? borderDash = null;
                    foreach (var bp in borderParts)
                    {
                        if (bp.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ||
                            bp.EndsWith("cm", StringComparison.OrdinalIgnoreCase) ||
                            bp.EndsWith("px", StringComparison.OrdinalIgnoreCase))
                            borderWidth = Core.EmuConverter.ParseEmu(bp);
                        else if (bp is "solid" or "dot" or "dash" or "lgDash" or "dashDot" or "sysDot" or "sysDash")
                            borderDash = bp;
                        else if (bp.Length >= 3 && !bp.Equals("none", StringComparison.OrdinalIgnoreCase))
                            borderColor = bp.TrimStart('#').ToUpperInvariant();
                    }

                    // Build line properties following POI's setBorderDefaults pattern
                    void ApplyBorderLine(OpenXmlCompositeElement lineProps)
                    {
                        // Remove NoFill if present (POI: setBorderDefaults line 265)
                        lineProps.RemoveAllChildren<Drawing.NoFill>();
                        // Set width (default 12700 EMU = 1pt like POI)
                        if (borderWidth.HasValue)
                        {
                            var wAttr = lineProps.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
                            if (wAttr.LocalName != null) lineProps.SetAttribute(new OpenXmlAttribute("", "w", null, borderWidth.Value.ToString()));
                            else lineProps.SetAttribute(new OpenXmlAttribute("", "w", null, borderWidth.Value.ToString()));
                        }
                        // Set color
                        if (borderColor != null)
                        {
                            lineProps.RemoveAllChildren<Drawing.SolidFill>();
                            lineProps.RemoveAllChildren<Drawing.NoFill>();
                            lineProps.AppendChild(BuildSolidFill(borderColor));
                        }
                        // Set dash style (default: solid)
                        if (borderDash != null)
                        {
                            lineProps.RemoveAllChildren<Drawing.PresetDash>();
                            lineProps.AppendChild(new Drawing.PresetDash
                            {
                                Val = borderDash switch
                                {
                                    "dot" => Drawing.PresetLineDashValues.Dot,
                                    "dash" => Drawing.PresetLineDashValues.Dash,
                                    "lgDash" => Drawing.PresetLineDashValues.LargeDash,
                                    "dashDot" => Drawing.PresetLineDashValues.DashDot,
                                    "sysDot" => Drawing.PresetLineDashValues.SystemDot,
                                    "sysDash" => Drawing.PresetLineDashValues.SystemDash,
                                    _ => Drawing.PresetLineDashValues.Solid
                                }
                            });
                        }
                    }

                    var edges = k switch
                    {
                        "border.left" => new[] { "left" },
                        "border.right" => new[] { "right" },
                        "border.top" => new[] { "top" },
                        "border.bottom" => new[] { "bottom" },
                        _ => new[] { "left", "right", "top", "bottom" }  // "border" or "border.all"
                    };

                    foreach (var edge in edges)
                    {
                        switch (edge)
                        {
                            case "left":
                                var lnL = tcPr.LeftBorderLineProperties ?? (tcPr.LeftBorderLineProperties = new Drawing.LeftBorderLineProperties());
                                ApplyBorderLine(lnL);
                                break;
                            case "right":
                                var lnR = tcPr.RightBorderLineProperties ?? (tcPr.RightBorderLineProperties = new Drawing.RightBorderLineProperties());
                                ApplyBorderLine(lnR);
                                break;
                            case "top":
                                var lnT = tcPr.TopBorderLineProperties ?? (tcPr.TopBorderLineProperties = new Drawing.TopBorderLineProperties());
                                ApplyBorderLine(lnT);
                                break;
                            case "bottom":
                                var lnB = tcPr.BottomBorderLineProperties ?? (tcPr.BottomBorderLineProperties = new Drawing.BottomBorderLineProperties());
                                ApplyBorderLine(lnB);
                                break;
                        }
                    }
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                        unsupported.Add(key);
                    break;
            }
        }
        return unsupported;
    }
}
