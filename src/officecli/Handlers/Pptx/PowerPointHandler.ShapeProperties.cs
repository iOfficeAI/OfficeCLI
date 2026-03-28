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
            if (value is null) { unsupported.Add(key); continue; }
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
                        // Shape-level: replace all text, preserve first run and paragraph formatting
                        var textBody = shape.TextBody;
                        if (textBody != null)
                        {
                            var firstPara = textBody.Elements<Drawing.Paragraph>().FirstOrDefault();
                            var firstRun = textBody.Descendants<Drawing.Run>().FirstOrDefault();
                            var runProps = firstRun?.RunProperties?.CloneNode(true) as Drawing.RunProperties;
                            var paraProps = firstPara?.ParagraphProperties?.CloneNode(true) as Drawing.ParagraphProperties;

                            textBody.RemoveAllChildren<Drawing.Paragraph>();

                            foreach (var textLine in textLines)
                            {
                                var newPara = new Drawing.Paragraph();
                                if (paraProps != null)
                                    newPara.ParagraphProperties = paraProps.CloneNode(true) as Drawing.ParagraphProperties;
                                var newRun = new Drawing.Run();
                                if (runProps != null)
                                    newRun.RunProperties = runProps.CloneNode(true) as Drawing.RunProperties;
                                newRun.Text = new Drawing.Text(textLine);
                                newPara.Append(newRun);
                                textBody.Append(newPara);
                            }
                        }
                    }
                    // Refresh runs list so subsequent properties target the new runs
                    runs.Clear();
                    runs.AddRange(GetAllRuns(shape));

                    // Check for text overflow and warn if text may exceed shape bounds
                    CheckTextOverflow(shape, value, unsupported);
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
                {
                    // Build fill before removing old one (atomic: no data loss on invalid color)
                    var colorFill = BuildSolidFill(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        rProps.RemoveAllChildren<Drawing.GradientFill>();
                        var fill = (Drawing.SolidFill)colorFill.CloneNode(true);
                        if (rProps is OpenXmlCompositeElement composite)
                        {
                            if (!composite.AddChild(fill, throwOnError: false))
                                rProps.AppendChild(fill);
                        }
                        else
                        {
                            rProps.AppendChild(fill);
                        }
                    }
                    break;
                }

                case "textfill" or "textgradient":
                {
                    // Build fill before removing old one (atomic: no data loss on invalid value)
                    OpenXmlElement newTextFill = value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        ? new Drawing.NoFill()
                        : BuildGradientFill(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        rProps.RemoveAllChildren<Drawing.GradientFill>();
                        rProps.RemoveAllChildren<Drawing.NoFill>();
                        InsertFillInRunProperties(rProps, newTextFill.CloneNode(true));
                    }
                    break;
                }

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
                            _ => throw new ArgumentException($"Invalid underline value: '{value}'. Valid values: single, double, heavy, dotted, dash, wavy, none.")
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
                            _ => throw new ArgumentException($"Invalid strikethrough value: '{value}'. Valid values: single, double, none.")
                        };
                    }
                    break;

                case "baseline" or "superscript" or "subscript":
                {
                    // Baseline offset: positive = superscript, negative = subscript
                    // Value in percent (e.g. "30" = 30% superscript, "-25" = 25% subscript)
                    // OOXML stores as 1/1000ths of percent (30000 = 30%)
                    // Shortcuts: "super"/"true" = 30%, "sub" = -25%, "none"/"false" = 0
                    int baselineVal;
                    if (key.ToLowerInvariant() == "superscript")
                        baselineVal = IsTruthy(value) ? 30000 : 0;
                    else if (key.ToLowerInvariant() == "subscript")
                        baselineVal = IsTruthy(value) ? -25000 : 0;
                    else
                    {
                        baselineVal = value.ToLowerInvariant() switch
                        {
                            "super" or "true" => 30000,
                            "sub" => -25000,
                            "none" or "false" or "0" => 0,
                            _ => double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var blVal) && !double.IsNaN(blVal) && !double.IsInfinity(blVal)
                                ? (int)(blVal * 1000)
                                : throw new ArgumentException($"Invalid 'baseline' value: '{value}'. Expected 'super', 'sub', 'none', or a percentage (e.g. 30 for superscript 30%).")
                        };
                    }
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Baseline = baselineVal;
                    }
                    break;
                }

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
                    // Remove any existing geometry (preset or custom) before setting new one
                    spPr.RemoveAllChildren<Drawing.CustomGeometry>();
                    var existingGeom = spPr.GetFirstChild<Drawing.PresetGeometry>();
                    if (existingGeom != null)
                        existingGeom.Preset = ParsePresetShape(value);
                    else
                        spPr.AppendChild(new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = ParsePresetShape(value) });
                    break;
                }

                case "geometry" or "path" when key.ToLowerInvariant() != "path" || shape.ShapeProperties != null:
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    // Check if value is a preset shape name (no spaces, no commas, simple identifier)
                    if (!value.Contains(' ') && !value.Contains(',') && !value.Contains('M'))
                    {
                        // Treat as preset shape name
                        spPr.RemoveAllChildren<Drawing.CustomGeometry>();
                        var existingGeom = spPr.GetFirstChild<Drawing.PresetGeometry>();
                        if (existingGeom != null)
                            existingGeom.Preset = ParsePresetShape(value);
                        else
                            spPr.AppendChild(new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = ParsePresetShape(value) });
                    }
                    else
                    {
                        // Custom geometry path:
                        // Format: "M x,y L x,y L x,y C x1,y1 x2,y2 x,y Z" (SVG-like path syntax)
                        spPr.RemoveAllChildren<Drawing.PresetGeometry>();
                        spPr.RemoveAllChildren<Drawing.CustomGeometry>();
                        // Insert after xfrm (OOXML requires geometry before fill/line)
                        var xfrm = spPr.GetFirstChild<Drawing.Transform2D>();
                        var custGeom = ParseCustomGeometry(value);
                        if (xfrm != null)
                            xfrm.InsertAfterSelf(custGeom);
                        else
                            spPr.PrependChild(custGeom);
                    }
                    break;
                }

                case "line" or "linecolor" or "line.color":
                {
                    // Build fill before removing old one (atomic)
                    OpenXmlElement newLineFill = value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        ? new Drawing.NoFill()
                        : BuildSolidFill(value);
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.SolidFill>();
                    outline.RemoveAllChildren<Drawing.NoFill>();
                    // CT_LineProperties schema: fill (solidFill/noFill/gradFill/pattFill) → prstDash → ...
                    var prstDash = outline.GetFirstChild<Drawing.PresetDash>();
                    if (prstDash != null)
                        outline.InsertBefore(newLineFill, prstDash);
                    else
                        outline.AppendChild(newLineFill);
                    break;
                }

                case "linewidth" or "line.width":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.Width = Core.EmuConverter.ParseLineWidth(value);
                    break;
                }

                case "linedash" or "line.dash":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    outline.AppendChild(new Drawing.PresetDash { Val = value.ToLowerInvariant() switch
                    {
                        "solid" => Drawing.PresetLineDashValues.Solid,
                        "dot" => Drawing.PresetLineDashValues.Dot,
                        "dash" => Drawing.PresetLineDashValues.Dash,
                        "dashdot" or "dash_dot" => Drawing.PresetLineDashValues.DashDot,
                        "longdash" or "lgdash" or "lg_dash" => Drawing.PresetLineDashValues.LargeDash,
                        "longdashdot" or "lgdashdot" or "lg_dash_dot" => Drawing.PresetLineDashValues.LargeDashDot,
                        _ => throw new ArgumentException($"Invalid 'lineDash' value: '{value}'. Valid values: solid, dot, dash, dashdot, longdash, longdashdot.")
                    }});
                    break;
                }

                case "lineopacity" or "line.opacity":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var lnOpacity) || double.IsNaN(lnOpacity) || double.IsInfinity(lnOpacity))
                        throw new ArgumentException($"Invalid 'lineopacity' value: '{value}'. Expected a finite decimal 0.0-1.0 (e.g. 0.5 = 50% opacity).");
                    var outline = EnsureOutline(spPr);
                    var solidFillLn = outline.GetFirstChild<Drawing.SolidFill>();
                    if (solidFillLn == null)
                    {
                        // Auto-create a black line fill (matching Apache POI behavior)
                        solidFillLn = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" });
                        outline.PrependChild(solidFillLn);
                    }
                    {
                        var colorEl = solidFillLn.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                            ?? solidFillLn.GetFirstChild<Drawing.SchemeColor>();
                        if (colorEl != null)
                        {
                            colorEl.RemoveAllChildren<Drawing.Alpha>();
                            var pct = (int)(lnOpacity * 100000); // 0.0-1.0 → 0-100000
                            colorEl.AppendChild(new Drawing.Alpha { Val = pct });
                        }
                    }
                    break;
                }

                case "rotation" or "rotate":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var rotVal) || double.IsNaN(rotVal) || double.IsInfinity(rotVal))
                        throw new ArgumentException($"Invalid 'rotation' value: '{value}'. Expected a finite number in degrees (e.g. 45, -90, 180.5).");
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    xfrm.Rotation = (int)(rotVal * 60000); // degrees to 60000ths
                    break;
                }

                case "opacity":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var opacityVal) || double.IsNaN(opacityVal) || double.IsInfinity(opacityVal))
                        throw new ArgumentException($"Invalid 'opacity' value: '{value}'. Expected a finite decimal 0.0-1.0 (e.g. 0.5 = 50% opacity).");
                    if (opacityVal > 1.0) opacityVal /= 100.0; // treat >1 as percentage (e.g. 30 → 0.30)
                    var alphaPct = (int)(opacityVal * 100000); // 0.0-1.0 → 0-100000

                    // Apply alpha to gradient fill stops if present
                    var gradFill = spPr.GetFirstChild<Drawing.GradientFill>();
                    if (gradFill != null)
                    {
                        var gradStops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>();
                        if (gradStops != null)
                        {
                            foreach (var stop in gradStops)
                            {
                                var stopColorEl = stop.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                                    ?? stop.GetFirstChild<Drawing.SchemeColor>();
                                if (stopColorEl != null)
                                {
                                    stopColorEl.RemoveAllChildren<Drawing.Alpha>();
                                    stopColorEl.AppendChild(new Drawing.Alpha { Val = alphaPct });
                                }
                            }
                        }
                        break;
                    }

                    var solidFill = spPr.GetFirstChild<Drawing.SolidFill>();
                    if (solidFill == null)
                    {
                        // Auto-create a white fill (matching Apache POI behavior)
                        spPr.RemoveAllChildren<Drawing.NoFill>();
                        solidFill = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "FFFFFF" });
                        InsertFillElement(spPr, solidFill);
                    }
                    {
                        var colorEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                            ?? solidFill.GetFirstChild<Drawing.SchemeColor>();
                        if (colorEl != null)
                        {
                            colorEl.RemoveAllChildren<Drawing.Alpha>();
                            colorEl.AppendChild(new Drawing.Alpha { Val = alphaPct });
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

                case "spacing" or "charspacing" or "letterspacing":
                {
                    // Character spacing in points (e.g. "2" = +2pt, "-1" = -1pt)
                    // Stored as 1/100th of a point in OOXML (POI: setSpc((int)(100*spc)))
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var spcDbl) || double.IsNaN(spcDbl) || double.IsInfinity(spcDbl))
                        throw new ArgumentException($"Invalid 'charspacing' value: '{value}'. Expected a finite number in points (e.g. 2, -1, 0.5).");
                    var spcVal = (int)(spcDbl * 100);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Spacing = spcVal;
                    }
                    break;
                }

                case "indent":
                {
                    var indentEmu = (int)ParseEmu(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Indent = indentEmu;
                    }
                    break;
                }

                case "marginleft" or "marl":
                {
                    var mlEmu = (int)ParseEmu(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.LeftMargin = mlEmu;
                    }
                    break;
                }

                case "marginright" or "marr":
                {
                    var mrEmu = (int)ParseEmu(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RightMargin = mrEmu;
                    }
                    break;
                }

                case "linespacing" or "line.spacing":
                {
                    var (lsIntVal, lsIsPct) = SpacingConverter.ParsePptLineSpacing(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.LineSpacing>();
                        var lnSpcElem = lsIsPct
                            ? new Drawing.LineSpacing(new Drawing.SpacingPercent { Val = lsIntVal })
                            : new Drawing.LineSpacing(new Drawing.SpacingPoints { Val = lsIntVal });
                        // CT_TextParagraphProperties schema: lnSpc → spcBef → spcAft
                        var spcBef = pProps.GetFirstChild<Drawing.SpaceBefore>();
                        var spcAft = pProps.GetFirstChild<Drawing.SpaceAfter>();
                        var insertBefore = spcBef ?? (OpenXmlElement?)spcAft;
                        if (insertBefore != null)
                            pProps.InsertBefore(lnSpcElem, insertBefore);
                        else
                            pProps.AppendChild(lnSpcElem);
                    }
                    break;
                }

                case "spacebefore" or "space.before":
                {
                    var sbIntVal = SpacingConverter.ParsePptSpacing(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                        var spcBefElem = new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = sbIntVal });
                        // CT_TextParagraphProperties schema: lnSpc → spcBef → spcAft
                        var spcAftRef = pProps.GetFirstChild<Drawing.SpaceAfter>();
                        if (spcAftRef != null)
                            pProps.InsertBefore(spcBefElem, spcAftRef);
                        else
                            pProps.AppendChild(spcBefElem);
                    }
                    break;
                }

                case "spaceafter" or "space.after":
                {
                    var saIntVal = SpacingConverter.ParsePptSpacing(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                        pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = saIntVal }));
                    }
                    break;
                }

                case "textwarp" or "wordart":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.RemoveAllChildren<Drawing.PresetTextWarp>();
                    if (!string.IsNullOrWhiteSpace(value) && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var warpName = value.StartsWith("text") ? value : $"text{char.ToUpper(value[0])}{value[1..]}";
                        var warpEnum = new Drawing.TextShapeValues(warpName);
                        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
                        var testWarp = new Drawing.PresetTextWarp(new Drawing.AdjustValueList()) { Preset = warpEnum };
                        var errors = validator.Validate(testWarp);
                        if (errors.Any())
                            throw new ArgumentException($"Invalid textwarp preset: '{value}'. Use full preset names like 'textArchUp', 'textWave1', 'textInflate', etc.");
                        bodyPr.AppendChild(testWarp);
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
                        case "true" or "normal" or "normautofit" or "auto" or "shrink": bodyPr.AppendChild(new Drawing.NormalAutoFit()); break;
                        case "shape" or "spautofit" or "resize": bodyPr.AppendChild(new Drawing.ShapeAutoFit()); break;
                        case "false" or "none": bodyPr.AppendChild(new Drawing.NoAutoFit()); break;
                        default: throw new ArgumentException($"Invalid autofit value: '{value}'. Valid values: true/normal/shrink, shape/resize, false/none.");
                    }
                    break;
                }

                case "x" or "y" or "width" or "height":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    TryApplyPositionSize(key.ToLowerInvariant(), value,
                        xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                        xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                    break;
                }

                case "shadow":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var shadowVal = value;
                    if (IsTruthy(shadowVal)) shadowVal = "000000";
                    if (IsNoFillShape(spPr) && runs.Count > 0)
                        foreach (var run in runs) ApplyTextShadow(run, shadowVal);
                    else
                        ApplyShadow(spPr, shadowVal);
                    break;
                }

                case "glow":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var glowVal = value;
                    if (IsTruthy(glowVal)) glowVal = "4472C4";
                    if (IsNoFillShape(spPr) && runs.Count > 0)
                        foreach (var run in runs) ApplyTextGlow(run, glowVal);
                    else
                        ApplyGlow(spPr, glowVal);
                    break;
                }

                case "reflection":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    if (IsNoFillShape(spPr) && runs.Count > 0)
                        foreach (var run in runs) ApplyTextReflection(run, value);
                    else
                        ApplyReflection(spPr, value);
                    break;
                }

                case "softedge":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    if (IsNoFillShape(spPr) && runs.Count > 0)
                        foreach (var run in runs) ApplyTextSoftEdge(run, value);
                    else
                        ApplySoftEdge(spPr, value);
                    break;
                }

                case "fliph":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    xfrm.HorizontalFlip = IsTruthy(value);
                    break;
                }

                case "flipv":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    xfrm.VerticalFlip = IsTruthy(value);
                    break;
                }

                case "rot3d" or "rotation3d" or "3drotation" or "3d.rotation":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    Apply3DRotation(spPr, value);
                    break;
                }

                case "rotx":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    Apply3DRotationAxis(spPr, "x", value);
                    break;
                }

                case "roty":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    Apply3DRotationAxis(spPr, "y", value);
                    break;
                }

                case "rotz":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    Apply3DRotationAxis(spPr, "z", value);
                    break;
                }

                case "bevel" or "beveltop":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyBevel(spPr, value, top: true);
                    break;
                }

                case "bevelbottom":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyBevel(spPr, value, top: false);
                    break;
                }

                case "depth" or "extrusion" or "3ddepth" or "3d.depth":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    Apply3DDepth(spPr, value);
                    break;
                }

                case "material":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    Apply3DMaterial(spPr, value);
                    break;
                }

                case "lighting" or "lightrig":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyLightRig(spPr, value);
                    break;
                }

                case "name":
                {
                    var nvPr = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                    if (nvPr != null) nvPr.Name = value;
                    else unsupported.Add(key);
                    break;
                }

                default:
                    if (!GenericXmlQuery.SetGenericAttribute(shape, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid shape props: text, bold, italic, underline, color, fill, size, font, gradient, line, opacity, align, valign, x, y, width, height, rotation, name, link, animation)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }

        return unsupported;
    }

    /// <summary>Ensure the cell has at least one Drawing.Run, creating one if needed.</summary>
    private static void EnsureTableCellHasRun(Drawing.TableCell cell)
    {
        if (cell.Descendants<Drawing.Run>().Any()) return;
        var textBody = cell.TextBody;
        if (textBody == null)
        {
            textBody = new Drawing.TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle());
            cell.PrependChild(textBody);
        }
        var para = textBody.Elements<Drawing.Paragraph>().FirstOrDefault();
        if (para == null)
        {
            para = new Drawing.Paragraph();
            textBody.Append(para);
        }
        var run = new Drawing.Run(
            new Drawing.RunProperties { Language = "en-US" },
            new Drawing.Text(""));
        para.Append(run);
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
                    EnsureTableCellHasRun(cell);
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
                    EnsureTableCellHasRun(cell);
                    var sz = (int)Math.Round(ParseFontSize(value) * 100);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sz;
                    }
                    break;
                case "bold":
                    EnsureTableCellHasRun(cell);
                    var b = IsTruthy(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = b;
                    }
                    break;
                case "italic":
                    EnsureTableCellHasRun(cell);
                    var it = IsTruthy(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = it;
                    }
                    break;
                case "color":
                {
                    // Build fill before removing old one (atomic)
                    EnsureTableCellHasRun(cell);
                    var cellColorFill = BuildSolidFill(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        rProps.AppendChild((Drawing.SolidFill)cellColorFill.CloneNode(true));
                    }
                    break;
                }
                case "fill":
                {
                    // Build new fill element BEFORE removing old one (atomic: no data loss on invalid color)
                    OpenXmlElement newCellFill;
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        newCellFill = new Drawing.NoFill();
                    }
                    else if (value.Contains('-'))
                    {
                        // Gradient fill: "FF0000-0000FF" or "FF0000-0000FF-90"
                        var gradParts = value.Split('-');
                        var colors = gradParts.ToList();
                        double degree = 0;
                        if (colors.Count >= 2 && double.TryParse(colors.Last(),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var angleDeg)
                            && colors.Last().Length <= 3)
                        {
                            degree = angleDeg;
                            colors.RemoveAt(colors.Count - 1);
                        }
                        if (colors.Count < 2) colors.Add(colors[0]);

                        // Validate that all segments look like hex colors
                        foreach (var c in colors)
                        {
                            var hex = c.TrimStart('#');
                            if (hex.Length < 3 || !hex.All(ch => char.IsAsciiHexDigit(ch)))
                                Console.Error.WriteLine($"Warning: '{c}' does not look like a hex color. Gradient format: COLOR1-COLOR2[-ANGLE] e.g. FF0000-0000FF-90");
                        }

                        var gradFill = new Drawing.GradientFill();
                        var gsList = new Drawing.GradientStopList();
                        for (int gi = 0; gi < colors.Count; gi++)
                        {
                            var pos = colors.Count == 1 ? 0 : gi * 100000 / (colors.Count - 1);
                            var (cRgb, cAlpha) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(colors[gi]);
                            var cEl = new Drawing.RgbColorModelHex { Val = cRgb };
                            if (cAlpha.HasValue) cEl.AppendChild(new Drawing.Alpha { Val = cAlpha.Value });
                            gsList.Append(new Drawing.GradientStop(cEl) { Position = pos });
                        }
                        gradFill.Append(gsList);
                        gradFill.Append(new Drawing.LinearGradientFill { Angle = (int)(degree * 60000), Scaled = true });
                        newCellFill = gradFill;
                    }
                    else
                    {
                        newCellFill = BuildSolidFill(value);
                    }

                    var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                    if (tcPr == null)
                    {
                        tcPr = new Drawing.TableCellProperties();
                        cell.Append(tcPr);
                    }
                    tcPr.RemoveAllChildren<Drawing.SolidFill>();
                    tcPr.RemoveAllChildren<Drawing.NoFill>();
                    tcPr.RemoveAllChildren<Drawing.GradientFill>();
                    tcPr.RemoveAllChildren<Drawing.BlipFill>();
                    // Insert fill after border line elements to maintain CT_TableCellProperties schema order
                    var lastBorder = tcPr.ChildElements.LastOrDefault(c =>
                        c is Drawing.LeftBorderLineProperties
                        or Drawing.RightBorderLineProperties
                        or Drawing.TopBorderLineProperties
                        or Drawing.BottomBorderLineProperties
                        or Drawing.TopLeftToBottomRightBorderLineProperties
                        or Drawing.BottomLeftToTopRightBorderLineProperties);
                    if (lastBorder != null)
                        lastBorder.InsertAfterSelf(newCellFill);
                    else
                        tcPr.Append(newCellFill);
                    break;
                }
                case "align" or "alignment":
                {
                    foreach (var para in cell.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
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
                        _ => throw new ArgumentException($"Invalid valign value: '{value}'. Valid values: top, middle, center, bottom.")
                    };
                    break;
                }
                case "gridspan" or "colspan":
                    cell.GridSpan = new DocumentFormat.OpenXml.Int32Value(ParseHelpers.SafeParseInt(value, "gridspan"));
                    break;
                case "rowspan":
                    cell.RowSpan = new DocumentFormat.OpenXml.Int32Value(ParseHelpers.SafeParseInt(value, "rowspan"));
                    break;
                case "vmerge":
                    cell.VerticalMerge = new DocumentFormat.OpenXml.BooleanValue(IsTruthy(value));
                    break;
                case "hmerge":
                    cell.HorizontalMerge = new DocumentFormat.OpenXml.BooleanValue(IsTruthy(value));
                    break;
                case "merge.right":
                {
                    // Convenience: merge.right=N sets gridSpan on this cell and hMerge on next N cells
                    var span = ParseHelpers.SafeParseInt(value, "merge.right") + 1;
                    cell.GridSpan = new DocumentFormat.OpenXml.Int32Value(span);
                    var row = cell.Parent as Drawing.TableRow;
                    if (row != null)
                    {
                        var cells = row.Elements<Drawing.TableCell>().ToList();
                        var idx = cells.IndexOf(cell);
                        for (int mi = idx + 1; mi < idx + span && mi < cells.Count; mi++)
                            cells[mi].HorizontalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                    }
                    break;
                }
                case "merge.down":
                {
                    // Convenience: merge.down=N sets rowSpan on this cell and vMerge on cells below
                    var rSpan = ParseHelpers.SafeParseInt(value, "merge.down") + 1;
                    cell.RowSpan = new DocumentFormat.OpenXml.Int32Value(rSpan);
                    var row = cell.Parent as Drawing.TableRow;
                    var table = row?.Parent;
                    if (table != null && row != null)
                    {
                        var rows = table.Elements<Drawing.TableRow>().ToList();
                        var rowIdx = rows.IndexOf(row);
                        var cells = row.Elements<Drawing.TableCell>().ToList();
                        var colIdx = cells.IndexOf(cell);
                        for (int ri = rowIdx + 1; ri < rowIdx + rSpan && ri < rows.Count; ri++)
                        {
                            var belowCells = rows[ri].Elements<Drawing.TableCell>().ToList();
                            if (colIdx < belowCells.Count)
                                belowCells[colIdx].VerticalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                        }
                    }
                    break;
                }
                case "underline":
                    EnsureTableCellHasRun(cell);
                    foreach (var run in cell.Descendants<Drawing.Run>())
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
                            _ => throw new ArgumentException($"Invalid underline value: '{value}'. Valid values: single, double, heavy, dotted, dash, wavy, none.")
                        };
                    }
                    break;
                case "strikethrough" or "strike":
                    EnsureTableCellHasRun(cell);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Strike = value.ToLowerInvariant() switch
                        {
                            "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                            "double" => Drawing.TextStrikeValues.DoubleStrike,
                            "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                            _ => throw new ArgumentException($"Invalid strikethrough value: '{value}'. Valid values: single, double, none.")
                        };
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

                    // Handle "none" — remove border by adding NoFill
                    bool isNone = value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase);

                    // Parse value: "FF0000", "1pt solid FF0000", "2pt dash 0000FF", or "style;width;color;dash"
                    string? borderColor = null;
                    long? borderWidth = null;
                    string? borderDash = null;
                    if (!isNone)
                    {
                        if (value.Contains(';'))
                        {
                            // Semicolon format: style;width;color[;dash]
                            var scParts = value.Split(';');
                            // Part 0: style (ignored for table border — used for Word only)
                            // Part 1: width (in pt/EMU)
                            if (scParts.Length > 1 && !string.IsNullOrEmpty(scParts[1]))
                            {
                                var wStr = scParts[1];
                                if (!wStr.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                                    wStr += "pt";
                                borderWidth = Core.EmuConverter.ParseEmu(wStr);
                            }
                            // Part 2: color
                            if (scParts.Length > 2 && !string.IsNullOrEmpty(scParts[2]))
                                borderColor = scParts[2].TrimStart('#').ToUpperInvariant();
                            // Part 3: dash style
                            if (scParts.Length > 3)
                            {
                                var d = scParts[3].ToLowerInvariant();
                                if (d is "solid" or "dot" or "dash" or "lgdash" or "dashdot" or "sysdot" or "sysdash")
                                    borderDash = d;
                                else
                                    throw new ArgumentException($"Invalid border dash value: '{scParts[3]}'. Valid values: solid, dot, dash, lgDash, dashDot, sysDot, sysDash.");
                            }
                        }
                        else
                        {
                            // Space-separated format: "2pt dash FF0000"
                            var borderParts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                            foreach (var bp in borderParts)
                            {
                                if (bp.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ||
                                    bp.EndsWith("cm", StringComparison.OrdinalIgnoreCase) ||
                                    bp.EndsWith("px", StringComparison.OrdinalIgnoreCase))
                                    borderWidth = Core.EmuConverter.ParseEmu(bp);
                                else if (bp.ToLowerInvariant() is "solid" or "dot" or "dash" or "lgdash" or "dashdot" or "sysdot" or "sysdash")
                                    borderDash = bp.ToLowerInvariant();
                                else if (bp.Length >= 3 && !bp.Equals("none", StringComparison.OrdinalIgnoreCase))
                                    borderColor = bp.TrimStart('#').ToUpperInvariant();
                            }
                        }
                    }

                    // Build line properties following POI's setBorderDefaults pattern
                    void ApplyBorderLine(OpenXmlCompositeElement lineProps)
                    {
                        if (isNone)
                        {
                            // Remove border: clear all children and add NoFill
                            lineProps.RemoveAllChildren<Drawing.SolidFill>();
                            lineProps.RemoveAllChildren<Drawing.PresetDash>();
                            lineProps.RemoveAllChildren<Drawing.NoFill>();
                            lineProps.AppendChild(new Drawing.NoFill());
                            return;
                        }
                        // Remove NoFill if present (POI: setBorderDefaults line 265)
                        lineProps.RemoveAllChildren<Drawing.NoFill>();
                        // Set width (default 12700 EMU = 1pt like POI)
                        if (borderWidth.HasValue)
                        {
                            var wAttr = lineProps.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
                            lineProps.SetAttribute(new OpenXmlAttribute("", "w", null!, borderWidth.Value.ToString()));
                        }
                        // Set color (build before removing for atomicity)
                        if (borderColor != null)
                        {
                            var borderFill = BuildSolidFill(borderColor);
                            lineProps.RemoveAllChildren<Drawing.SolidFill>();
                            lineProps.RemoveAllChildren<Drawing.NoFill>();
                            lineProps.AppendChild(borderFill);
                        }
                        // Set dash style (default: solid)
                        if (borderDash != null)
                        {
                            lineProps.RemoveAllChildren<Drawing.PresetDash>();
                            lineProps.AppendChild(new Drawing.PresetDash
                            {
                                Val = borderDash.ToLowerInvariant() switch
                                {
                                    "dot" => Drawing.PresetLineDashValues.Dot,
                                    "dash" => Drawing.PresetLineDashValues.Dash,
                                    "lgdash" => Drawing.PresetLineDashValues.LargeDash,
                                    "dashdot" => Drawing.PresetLineDashValues.DashDot,
                                    "sysdot" => Drawing.PresetLineDashValues.SystemDot,
                                    "sysdash" => Drawing.PresetLineDashValues.SystemDash,
                                    "solid" => Drawing.PresetLineDashValues.Solid,
                                    _ => throw new ArgumentException($"Invalid border dash value: '{borderDash}'. Valid values: solid, dot, dash, lgDash, dashDot, sysDot, sysDash.")
                                }
                            });
                        }
                    }

                    var edges = k switch
                    {
                        "border.left" or "border.left.width" or "border.left.color" or "border.left.dash" => new[] { "left" },
                        "border.right" or "border.right.width" or "border.right.color" or "border.right.dash" => new[] { "right" },
                        "border.top" or "border.top.width" or "border.top.color" or "border.top.dash" => new[] { "top" },
                        "border.bottom" or "border.bottom.width" or "border.bottom.color" or "border.bottom.dash" => new[] { "bottom" },
                        "border.tl2br" or "border.tl2br.width" or "border.tl2br.color" or "border.tl2br.dash" => new[] { "tl2br" },
                        "border.tr2bl" or "border.tr2bl.width" or "border.tr2bl.color" or "border.tr2bl.dash" => new[] { "tr2bl" },
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
                            case "tl2br":
                                var lnTl = tcPr.TopLeftToBottomRightBorderLineProperties ?? (tcPr.TopLeftToBottomRightBorderLineProperties = new Drawing.TopLeftToBottomRightBorderLineProperties());
                                ApplyBorderLine(lnTl);
                                break;
                            case "tr2bl":
                                var lnTr = tcPr.BottomLeftToTopRightBorderLineProperties ?? (tcPr.BottomLeftToTopRightBorderLineProperties = new Drawing.BottomLeftToTopRightBorderLineProperties());
                                ApplyBorderLine(lnTr);
                                break;
                        }
                    }
                    break;
                }
                case "margin" or "padding":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    var parts = value.Split(',');
                    if (parts.Length == 1)
                    {
                        var emu = (int)ParseEmu(parts[0].Trim());
                        tcPrM.LeftMargin = emu;
                        tcPrM.RightMargin = emu;
                        tcPrM.TopMargin = emu;
                        tcPrM.BottomMargin = emu;
                    }
                    else if (parts.Length == 4)
                    {
                        tcPrM.LeftMargin = (int)ParseEmu(parts[0].Trim());
                        tcPrM.TopMargin = (int)ParseEmu(parts[1].Trim());
                        tcPrM.RightMargin = (int)ParseEmu(parts[2].Trim());
                        tcPrM.BottomMargin = (int)ParseEmu(parts[3].Trim());
                    }
                    else if (parts.Length == 2)
                    {
                        var h = (int)ParseEmu(parts[0].Trim());
                        var v = (int)ParseEmu(parts[1].Trim());
                        tcPrM.LeftMargin = h;
                        tcPrM.RightMargin = h;
                        tcPrM.TopMargin = v;
                        tcPrM.BottomMargin = v;
                    }
                    break;
                }
                case "margin.left" or "padding.left":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    tcPrM.LeftMargin = (int)ParseEmu(value);
                    break;
                }
                case "margin.right" or "padding.right":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    tcPrM.RightMargin = (int)ParseEmu(value);
                    break;
                }
                case "margin.top" or "padding.top":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    tcPrM.TopMargin = (int)ParseEmu(value);
                    break;
                }
                case "margin.bottom" or "padding.bottom":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    tcPrM.BottomMargin = (int)ParseEmu(value);
                    break;
                }
                case "textdirection" or "textdir" or "vert":
                {
                    var tcPrTd = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    tcPrTd.Vertical = value.ToLowerInvariant() switch
                    {
                        "horizontal" or "horz" or "none" => Drawing.TextVerticalValues.Horizontal,
                        "vertical" or "vert" or "vert270" => Drawing.TextVerticalValues.Vertical270,
                        "vertical270" => Drawing.TextVerticalValues.Vertical270,
                        "vertical90" or "vert90" => Drawing.TextVerticalValues.Vertical,
                        "stacked" or "wordartvert" => Drawing.TextVerticalValues.WordArtVertical,
                        _ => throw new ArgumentException($"Invalid textDirection: '{value}'. Valid: horizontal, vertical, vertical90, vertical270, stacked.")
                    };
                    break;
                }
                case "wordwrap" or "wrap":
                {
                    var bodyProps = cell.TextBody?.GetFirstChild<Drawing.BodyProperties>();
                    if (bodyProps == null)
                    {
                        var tb = cell.TextBody ?? cell.PrependChild(new Drawing.TextBody(
                            new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph()));
                        bodyProps = tb.GetFirstChild<Drawing.BodyProperties>()!;
                    }
                    bodyProps.Wrap = IsTruthy(value) ? Drawing.TextWrappingValues.Square : Drawing.TextWrappingValues.None;
                    break;
                }
                case "linespacing":
                {
                    var (spcVal, isPct) = OfficeCli.Core.SpacingConverter.ParsePptLineSpacing(value);
                    foreach (var para in cell.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.LineSpacing>();
                        var ls = new Drawing.LineSpacing();
                        if (isPct) ls.AppendChild(new Drawing.SpacingPercent { Val = spcVal });
                        else ls.AppendChild(new Drawing.SpacingPoints { Val = spcVal });
                        // CT_TextParagraphProperties schema: lnSpc → spcBef → spcAft
                        var insertBefore = (OpenXmlElement?)pProps.GetFirstChild<Drawing.SpaceBefore>()
                            ?? pProps.GetFirstChild<Drawing.SpaceAfter>();
                        if (insertBefore != null) pProps.InsertBefore(ls, insertBefore);
                        else pProps.AppendChild(ls);
                    }
                    break;
                }
                case "spacebefore":
                {
                    var sbVal = OfficeCli.Core.SpacingConverter.ParsePptSpacing(value);
                    foreach (var para in cell.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                        var sb = new Drawing.SpaceBefore();
                        sb.AppendChild(new Drawing.SpacingPoints { Val = sbVal });
                        // CT_TextParagraphProperties schema: lnSpc → spcBef → spcAft
                        var spcAftRef = pProps.GetFirstChild<Drawing.SpaceAfter>();
                        if (spcAftRef != null) pProps.InsertBefore(sb, spcAftRef);
                        else pProps.AppendChild(sb);
                    }
                    break;
                }
                case "spaceafter":
                {
                    var saVal = OfficeCli.Core.SpacingConverter.ParsePptSpacing(value);
                    foreach (var para in cell.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                        var sa = new Drawing.SpaceAfter();
                        sa.AppendChild(new Drawing.SpacingPoints { Val = saVal });
                        pProps.AppendChild(sa); // spcAft is last, append is correct
                    }
                    break;
                }
                case "opacity" or "fill.opacity" or "alpha" or "fill.alpha":
                {
                    // Set fill opacity on the cell's existing fill element
                    var tcPrO = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                    if (tcPrO != null)
                    {
                        var opacityVal = ParseHelpers.SafeParseDouble(value, "opacity");
                        if (opacityVal > 1.0) opacityVal /= 100.0; // treat >1 as percentage (e.g. 50 → 0.50)
                        var alphaVal = (int)Math.Round(opacityVal * 100000); // 0.0-1.0 → 0-100000
                        alphaVal = Math.Max(0, Math.Min(100000, alphaVal));
                        var solidFill = tcPrO.GetFirstChild<Drawing.SolidFill>();
                        if (solidFill != null)
                        {
                            var colorEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()
                                ?? solidFill.GetFirstChild<Drawing.SchemeColor>() as OpenXmlElement;
                            if (colorEl != null)
                            {
                                colorEl.RemoveAllChildren<Drawing.Alpha>();
                                colorEl.AppendChild(new Drawing.Alpha { Val = alphaVal });
                            }
                        }
                    }
                    break;
                }
                case "bevel" or "cell3d":
                {
                    // Cell3D bevel gives a subtle rounded/embossed look
                    var tcPrB = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        tcPrB.RemoveAllChildren<Drawing.Cell3DProperties>();
                    }
                    else
                    {
                        var cell3d = tcPrB.GetFirstChild<Drawing.Cell3DProperties>();
                        if (cell3d == null)
                        {
                            cell3d = new Drawing.Cell3DProperties();
                            // CT_TableCellProperties schema: borders → cell3D → fill → extLst
                            var insertBefore = (OpenXmlElement?)tcPrB.GetFirstChild<Drawing.SolidFill>()
                                ?? (OpenXmlElement?)tcPrB.GetFirstChild<Drawing.NoFill>()
                                ?? (OpenXmlElement?)tcPrB.GetFirstChild<Drawing.GradientFill>()
                                ?? (OpenXmlElement?)tcPrB.GetFirstChild<Drawing.BlipFill>()
                                ?? (OpenXmlElement?)tcPrB.GetFirstChild<Drawing.PatternFill>()
                                ?? tcPrB.GetFirstChild<Drawing.ExtensionList>();
                            if (insertBefore != null) tcPrB.InsertBefore(cell3d, insertBefore);
                            else tcPrB.AppendChild(cell3d);
                        }
                        cell3d.RemoveAllChildren<Drawing.Bevel>();

                        // Parse: "circle" or "circle-6-6" (preset-width-height in pt)
                        var bevelParts = value.Split('-');
                        var preset = bevelParts[0].ToLowerInvariant() switch
                        {
                            "circle" => Drawing.BevelPresetValues.Circle,
                            "relaxedinset" => Drawing.BevelPresetValues.RelaxedInset,
                            "cross" => Drawing.BevelPresetValues.Cross,
                            "coolslant" => Drawing.BevelPresetValues.CoolSlant,
                            "angle" => Drawing.BevelPresetValues.Angle,
                            "softround" => Drawing.BevelPresetValues.SoftRound,
                            "convex" => Drawing.BevelPresetValues.Convex,
                            "slope" => Drawing.BevelPresetValues.Slope,
                            "artdeco" => Drawing.BevelPresetValues.ArtDeco,
                            _ => Drawing.BevelPresetValues.Circle
                        };
                        var bevel = new Drawing.Bevel { Preset = preset };
                        if (bevelParts.Length >= 2)
                            bevel.Width = (long)(ParseHelpers.SafeParseDouble(bevelParts[1], "bevel width") * 12700); // pt to EMU
                        if (bevelParts.Length >= 3)
                            bevel.Height = (long)(ParseHelpers.SafeParseDouble(bevelParts[2], "bevel height") * 12700);
                        cell3d.AppendChild(bevel);
                    }
                    break;
                }
                case "image":
                {
                    // Validate before modifying (atomic: no data loss on invalid input)
                    if (!File.Exists(value))
                        throw new FileNotFoundException($"Image file not found: {value}");

                    // Image fill on table cell (like POI CTBlipFillProperties on CTTableCellProperties)
                    var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                    if (tcPr == null) { tcPr = new Drawing.TableCellProperties(); cell.Append(tcPr); }
                    tcPr.RemoveAllChildren<Drawing.SolidFill>();
                    tcPr.RemoveAllChildren<Drawing.NoFill>();
                    tcPr.RemoveAllChildren<Drawing.GradientFill>();
                    tcPr.RemoveAllChildren<Drawing.BlipFill>();
                    var (cellImgStream, cellImgType) = OfficeCli.Core.ImageSource.Resolve(value);
                    using var cellImgDispose = cellImgStream;
                    // Find the SlidePart — the method is called from Set which has the slidePart context
                    var rootElement = cell.Ancestors<OpenXmlElement>().LastOrDefault() ?? cell;
                    var ownerPart = rootElement is DocumentFormat.OpenXml.Presentation.Slide slide
                        ? slide.SlidePart : null;
                    if (ownerPart == null) { unsupported.Add(key); break; }

                    var imgPart = ownerPart.AddImagePart(cellImgType);
                    imgPart.FeedData(cellImgStream);
                    var relId = ownerPart.GetIdOfPart(imgPart);

                    tcPr.Append(new Drawing.BlipFill(
                        new Drawing.Blip { Embed = relId },
                        new Drawing.Stretch(new Drawing.FillRectangle())
                    ));
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid cell props: text, bold, italic, underline, color, fill, size, font, align, valign, border, colspan, rowspan, margin)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        return unsupported;
    }

    /// <summary>
    /// Estimates whether the given text will overflow the shape bounds and adds a warning if so.
    /// This is a rough heuristic — actual rendering depends on font metrics, kerning, etc.
    /// </summary>
    private static void CheckTextOverflow(Shape shape, string text, List<string> unsupported)
    {
        var spPr = shape.ShapeProperties;
        var xfrm = spPr?.Transform2D;
        var extents = xfrm?.Extents;
        if (extents?.Cx == null || extents?.Cy == null) return;

        long cx = extents.Cx!.Value;  // width in EMU
        long cy = extents.Cy!.Value;  // height in EMU

        // Convert EMU to points (1pt = 12700 EMU)
        const double emuPerPt = 12700.0;
        double shapeWidthPt = cx / emuPerPt;
        double shapeHeightPt = cy / emuPerPt;

        // Account for internal margins (default PPT margins: 0.1in left/right, 0.05in top/bottom)
        // 1in = 72pt
        double marginLeftRight = 0.1 * 72 * 2; // ~14.4pt total
        double marginTopBottom = 0.05 * 72 * 2; // ~7.2pt total

        // Check for custom margins on the text body
        var textBody = shape.TextBody;
        if (textBody?.BodyProperties != null)
        {
            var bp = textBody.BodyProperties;
            if (bp.LeftInset != null || bp.RightInset != null)
            {
                long leftEmu = bp.LeftInset?.Value ?? 91440;  // default 0.1in
                long rightEmu = bp.RightInset?.Value ?? 91440;
                marginLeftRight = (leftEmu + rightEmu) / emuPerPt;
            }
            if (bp.TopInset != null || bp.BottomInset != null)
            {
                long topEmu = bp.TopInset?.Value ?? 45720;  // default 0.05in
                long bottomEmu = bp.BottomInset?.Value ?? 45720;
                marginTopBottom = (topEmu + bottomEmu) / emuPerPt;
            }
        }

        double usableWidth = shapeWidthPt - marginLeftRight;
        double usableHeight = shapeHeightPt - marginTopBottom;
        if (usableWidth <= 0 || usableHeight <= 0) return;

        // Get font size from first run (hundredths of a point), default 18pt
        double fontSizePt = 18.0;
        var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
        if (firstRun?.RunProperties?.FontSize != null)
            fontSizePt = firstRun.RunProperties.FontSize.Value / 100.0;

        // Estimate character width and line height
        double charWidth = fontSizePt * 0.6;
        double lineHeight = fontSizePt * 1.4;

        if (charWidth <= 0) return;
        int charsPerLine = (int)(usableWidth / charWidth);
        if (charsPerLine <= 0) charsPerLine = 1;

        // Calculate total lines needed
        var textLines = text.Replace("\\n", "\n").Split('\n');
        int totalLines = 0;
        foreach (var line in textLines)
        {
            if (line.Length == 0)
                totalLines += 1;
            else
                totalLines += (int)Math.Ceiling((double)line.Length / charsPerLine);
        }

        double estimatedHeight = totalLines * lineHeight;
        if (estimatedHeight > usableHeight)
        {
            unsupported.Add($"text may overflow shape bounds (est. {estimatedHeight:F0}pt text in {usableHeight:F0}pt shape)");
        }
    }
}
