// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private static List<Drawing.Run> GetAllRuns(Shape shape)
    {
        return shape.TextBody?.Elements<Drawing.Paragraph>()
            .SelectMany(p => p.Elements<Drawing.Run>()).ToList()
            ?? new List<Drawing.Run>();
    }

    // drawingML CT_TextCharacterProperties attribute set (rPr attrs).
    // Long-tail run-context Set in SetRunOrShapeProperties uses this to
    // distinguish attribute-pattern keys (set as XML attributes on rPr) from
    // child-pattern keys (route through TryCreateTypedChild). Symmetric with
    // FillUnknownRunProps in NodeBuilder.cs which surfaces these via Get.
    // Source: ECMA-376 Part 1, 21.1.2.3.9 (a:rPr).
    private static readonly System.Collections.Generic.HashSet<string> DrawingRunPropertyAttrs =
        new(System.StringComparer.Ordinal)
    {
        "kumimoji", "lang", "altLang", "sz", "b", "i", "u", "strike",
        "kern", "cap", "spc", "normalizeH", "baseline", "noProof",
        "dirty", "err", "smtClean", "smtId", "bmk",
    };

    // Schema-typed sub-sets used for value validation in run-context Set.
    // Without these, an out-of-domain value for any typed attribute (e.g.
    // kern=abc, u=GARBAGE) would be silently written as invalid OOXML — the
    // file then fails strict validation downstream. Source: ECMA-376 Part 1
    // 21.1.2.3.9 (a:rPr).
    private static readonly System.Collections.Generic.HashSet<string> DrawingRunIntAttrs =
        new(System.StringComparer.Ordinal) { "sz", "kern", "spc", "baseline", "smtId" };
    private static readonly System.Collections.Generic.HashSet<string> DrawingRunBoolAttrs =
        new(System.StringComparer.Ordinal) { "b", "i", "noProof", "normalizeH", "dirty", "err", "smtClean", "kumimoji" };

    // ST_TextUnderlineType — full enumeration per ECMA-376 §21.1.10.82.
    private static readonly System.Collections.Generic.HashSet<string> DrawingUnderlineEnum =
        new(System.StringComparer.Ordinal)
    {
        "none", "words", "sng", "dbl", "heavy", "dotted", "dottedHeavy",
        "dash", "dashHeavy", "dashLong", "dashLongHeavy",
        "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy",
        "wavy", "wavyHeavy", "wavyDbl",
    };
    // ST_TextStrikeType per ECMA-376 §21.1.10.78.
    private static readonly System.Collections.Generic.HashSet<string> DrawingStrikeEnum =
        new(System.StringComparer.Ordinal) { "noStrike", "sngStrike", "dblStrike" };
    // ST_TextCapsType per ECMA-376 §21.1.10.7.
    private static readonly System.Collections.Generic.HashSet<string> DrawingCapsEnum =
        new(System.StringComparer.Ordinal) { "none", "small", "all" };

    // Tolerant BCP-47 shape: starts with letter, allows letters/digits/hyphens.
    // Stricter than xsd:language but loose enough to accept all real-world tags
    // (zh-Hant-TW, en-US, x-private, ...). Rejects whitespace and special chars.
    private static readonly System.Text.RegularExpressions.Regex Bcp47Shape =
        new(@"^[A-Za-z][A-Za-z0-9-]*$", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static bool IsValidDrawingRunAttrValue(string key, string value)
    {
        if (DrawingRunIntAttrs.Contains(key)) return int.TryParse(value, out _);
        if (DrawingRunBoolAttrs.Contains(key))
            return value is "0" or "1" or "true" or "false" or "True" or "False";
        if (key == "u") return DrawingUnderlineEnum.Contains(value);
        if (key == "strike") return DrawingStrikeEnum.Contains(value);
        if (key == "cap") return DrawingCapsEnum.Contains(value);
        if (key is "lang" or "altLang") return string.IsNullOrEmpty(value) || Bcp47Shape.IsMatch(value);
        return true; // remaining string attrs (kumimoji handled above; bmk arbitrary string)
    }

    // runContext=true when the caller is a run-targeted Set path (e.g.
    // /slide[N]/shape[K]/r[R] or /slide[N]/shape[K]/p[P]/r[R]). Affects the
    // default branch only: long-tail unknown keys are routed to each run's
    // RunProperties (attribute or child) instead of the shape element.
    // Curated cases keep their existing per-key targeting (some still write
    // to shape regardless of context — fill, geometry, etc.).
    private static List<string> SetRunOrShapeProperties(
        Dictionary<string, string> properties, List<Drawing.Run> runs, Shape shape, OpenXmlPart? part = null,
        bool runContext = false)
    {
        var unsupported = new List<string>();

        // CONSISTENCY(allcaps-alias): map allCaps/smallCaps onto OOXML's `cap`
        // attribute so users mirroring CSS / Word vocabulary don't see UNSUPPORTED.
        // Mirrors WordHandler.Helpers.cs allcaps→Caps fix (commit ccaed17a).
        // Boolean-truthy → "all" / "small" ; explicit "none"/"false" → cap="none".
        if (!properties.ContainsKey("cap"))
        {
            string? capsKey = properties.Keys.FirstOrDefault(k =>
                k.Equals("allCaps", StringComparison.OrdinalIgnoreCase)
                || k.Equals("allcaps", StringComparison.OrdinalIgnoreCase));
            if (capsKey != null)
            {
                var v = properties[capsKey];
                properties = new Dictionary<string, string>(properties, properties.Comparer);
                properties.Remove(capsKey);
                properties["cap"] = (v is "0" or "false" or "False" or "none") ? "none" : "all";
            }
            string? smallCapsKey = properties.Keys.FirstOrDefault(k =>
                k.Equals("smallCaps", StringComparison.OrdinalIgnoreCase)
                || k.Equals("smallcaps", StringComparison.OrdinalIgnoreCase));
            if (smallCapsKey != null && !properties.ContainsKey("cap"))
            {
                var v = properties[smallCapsKey];
                properties = new Dictionary<string, string>(properties, properties.Comparer);
                properties.Remove(smallCapsKey);
                properties["cap"] = (v is "0" or "false" or "False" or "none") ? "none" : "small";
            }
        }

        // CONSISTENCY(prop-order): fill carriers (fill/gradient/pattern) must run
        // before modifier props (opacity attaches alpha to the resulting solidFill);
        // otherwise opacity auto-creates a white fill that fill= then overwrites.
        // Mirrors the implicit ordering in Add.Shape.cs which processes fill first.
        var orderedKeys = properties.Keys
            .OrderBy(k => k.ToLowerInvariant() switch
            {
                "fill" or "gradient" or "pattern" => 0,
                _ => 1
            })
            .ToList();

        foreach (var key in orderedKeys)
        {
            var value = properties[key];
            if (value is null) { unsupported.Add(key); continue; }
            switch (key.ToLowerInvariant())
            {
                case "cap":
                {
                    // Apply rPr/cap to every run in the shape (or to runs when in run context).
                    if (!DrawingCapsEnum.Contains(value))
                    {
                        unsupported.Add($"cap (value '{value}' must be one of: none, small, all)");
                        break;
                    }
                    var targetRuns = runs.Count > 0 ? runs : shape.Descendants<Drawing.Run>().ToList();
                    foreach (var run in targetRuns)
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.SetAttribute(new OpenXmlAttribute("", "cap", "", value));
                    }
                    break;
                }
                case "text":
                {
                    var textLines = value.Replace("\\n", "\n").Split('\n');
                    if (runs.Count == 1 && textLines.Length == 1)
                    {
                        // Single run, single line: just replace its text
                        runs[0].Text = new Drawing.Text { Text = textLines[0] };
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
                                newRun.Text = new Drawing.Text { Text = textLine };
                                newPara.Append(newRun);
                                textBody.Append(newPara);
                            }
                        }
                    }
                    // Refresh runs list so subsequent properties target the new runs
                    runs.Clear();
                    runs.AddRange(GetAllRuns(shape));

                    break;
                }

                case "font":
                case "font.name":
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
                case "fontSize":
                case "fontsize":
                case "font.size":
                    var sizeVal = (int)Math.Round(ParseFontSize(value) * 100);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                    break;

                case "bold":
                case "font.bold":
                    var isBold = IsTruthy(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                    break;

                case "italic":
                case "font.italic":
                    var isItalic = IsTruthy(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                    break;

                case "color":
                case "font.color":
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
                case "font.underline":
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

                case "strikethrough" or "strike" or "font.strike" or "font.strikethrough":
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

                case "pattern":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyPatternFill(spPr, value);
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
                    // R10: reject out-of-range opacity instead of writing invalid OOXML
                    // (a:alpha/@val must be in [0, 100000]). Negative input was producing
                    // <a:alpha val="-100000"/> which corrupts the file.
                    if (opacityVal < 0.0 || opacityVal > 1.0)
                        throw new ArgumentException($"Invalid 'opacity' value: '{value}'. Expected 0.0-1.0 (or 0-100 as percent).");
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

                case "spacing" or "charspacing" or "letterspacing" or "spc":
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
                        // Resolve ambiguous shorthands before applying the "text" prefix
                        var resolved = value.ToLowerInvariant() switch
                        {
                            "wave" => "textWave1",
                            "arch" => "textArchUp",
                            "circle" => "textCircle",
                            "button" => "textButton",
                            _ => value
                        };
                        var warpName = resolved.StartsWith("text", StringComparison.OrdinalIgnoreCase) ? resolved : $"text{char.ToUpper(resolved[0])}{resolved[1..]}";
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
                    if (IsValidBooleanString(shadowVal) && IsTruthy(shadowVal)) shadowVal = "000000";
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
                    if (IsValidBooleanString(glowVal) && IsTruthy(glowVal)) glowVal = "4472C4";
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

                case "blur":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyBlur(spPr, value);
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

                case "alt" or "alttext" or "description":
                {
                    var nvPr = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                    if (nvPr != null) nvPr.Description = value;
                    else unsupported.Add(key);
                    break;
                }

                case "formula":
                {
                    // Replace equation content in shape (a14:m > m:oMathPara > m:oMath)
                    var textBody = shape.TextBody;
                    if (textBody == null) { unsupported.Add(key); break; }

                    var mathContent = FormulaParser.Parse(value);
                    M.OfficeMath oMath = mathContent is M.OfficeMath dm
                        ? dm : new M.OfficeMath(mathContent.CloneNode(true));
                    var mathPara = new M.Paragraph(oMath);

                    // Find existing AlternateContent (equation container) or create one
                    var existingAlt = textBody.Descendants<AlternateContent>().FirstOrDefault();
                    if (existingAlt != null)
                    {
                        // Replace existing equation: update Choice (a14:m) and Fallback
                        var choice = existingAlt.GetFirstChild<AlternateContentChoice>();
                        if (choice != null)
                        {
                            choice.RemoveAllChildren();
                            choice.Requires = "a14";
                            var a14m = new OpenXmlUnknownElement("a14", "m", "http://schemas.microsoft.com/office/drawing/2010/main");
                            a14m.AppendChild(mathPara.CloneNode(true));
                            choice.AppendChild(a14m);
                        }
                        var fallback = existingAlt.GetFirstChild<AlternateContentFallback>();
                        if (fallback != null)
                        {
                            fallback.RemoveAllChildren();
                            var fbRun = new Drawing.Run(
                                new Drawing.RunProperties { Language = "en-US" },
                                new Drawing.Text { Text = FormulaParser.ToReadableText(mathPara) }
                            );
                            fallback.AppendChild(fbRun);
                        }
                    }
                    else
                    {
                        // No existing equation — build full structure
                        var a14m = new OpenXmlUnknownElement("a14", "m", "http://schemas.microsoft.com/office/drawing/2010/main");
                        a14m.AppendChild(mathPara.CloneNode(true));
                        var choice = new AlternateContentChoice { Requires = "a14" };
                        choice.AppendChild(a14m);
                        var fallback = new AlternateContentFallback();
                        fallback.AppendChild(new Drawing.Run(
                            new Drawing.RunProperties { Language = "en-US" },
                            new Drawing.Text { Text = FormulaParser.ToReadableText(mathPara) }
                        ));
                        var altContent = new AlternateContent();
                        altContent.AppendChild(choice);
                        altContent.AppendChild(fallback);

                        // Clear text body paragraphs and add equation paragraph
                        textBody.RemoveAllChildren<Drawing.Paragraph>();
                        var drawingPara = new Drawing.Paragraph();
                        drawingPara.AppendChild(altContent);
                        textBody.AppendChild(drawingPara);
                    }
                    break;
                }

                default:
                {
                    // Long-tail OOXML fallback. In run-context (e.g. set on
                    // /slide[N]/shape[K]/r[R]), drawingML rPr stores most
                    // properties as attributes on rPr itself (kern, spc,
                    // baseline, lang, dirty, smtClean, normalizeH, ...), with
                    // a few child-pattern props (effectLst, hlinkClick).
                    // Try attribute-setting first against the known
                    // drawingML CT_TextCharacterProperties attribute set; fall
                    // back to TryCreateTypedChild for child-pattern keys.
                    bool handledByRun = false;
                    // CONSISTENCY(rpr-attr-fallback): drawingML run-property
                    // attributes (spc, lang, kern, cap, baseline, ...) must
                    // route to rPr regardless of runContext. Shape-level Set
                    // applies to all runs (mirrors how bold/size/font work
                    // above); run-level Set applies to the targeted run only.
                    // Without this, shape-level spc/lang silently fell through
                    // to SetGenericAttribute(sp, ...) and wrote attributes onto
                    // the <p:sp> element, which Office ignores.
                    if (runs.Count > 0 && DrawingRunPropertyAttrs.Contains(key))
                    {
                        if (!IsValidDrawingRunAttrValue(key, value))
                        {
                            unsupported.Add($"{key} (value '{value}' is not valid for OOXML rPr/{key} type)");
                            break;
                        }
                        handledByRun = true;
                        foreach (var run in runs)
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rPr.SetAttribute(new OpenXmlAttribute("", key, "", value));
                        }
                    }
                    if (handledByRun) break;
                    if (runContext && runs.Count > 0)
                    {
                        // Child-pattern fallback (rare in rPr but exists for
                        // hlinkClick etc.). Symmetric with Word.
                        handledByRun = true;
                        foreach (var run in runs)
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            if (!GenericXmlQuery.TryCreateTypedChild(rPr, key, value))
                            {
                                handledByRun = false;
                                break;
                            }
                        }
                    }
                    if (handledByRun) break;
                    if (!GenericXmlQuery.SetGenericAttribute(shape, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid shape props: text, bold, italic, underline, color, fill, size, font, gradient, line, opacity, align, valign, x, y, width, height, rotation, name, link, animation, formula, geometry, preset, shadow, glow, reflection, softEdge, pattern, flip, flipH, flipV)");
                        else
                            unsupported.Add(key);
                    }
                    break;
                }
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
            new Drawing.Text { Text = "" });
        para.Append(run);
    }

    /// <summary>
    /// Replace the text content of a table cell's first paragraph with the given value.
    /// Removes any existing runs/breaks and preserves EndParagraphRunProperties ordering
    /// (schema requires Run before EndParagraphRunProperties).
    /// </summary>
    private static void ReplaceCellText(Drawing.TableCell cell, string value)
    {
        var txBody = cell.TextBody;
        if (txBody == null)
        {
            txBody = new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph());
            cell.AppendChild(txBody);
        }
        var para = txBody.Elements<Drawing.Paragraph>().FirstOrDefault()
            ?? txBody.AppendChild(new Drawing.Paragraph());
        para.RemoveAllChildren<Drawing.Run>();
        para.RemoveAllChildren<Drawing.Break>();
        var savedEndParaRPr = para.Elements<Drawing.EndParagraphRunProperties>().FirstOrDefault();
        if (savedEndParaRPr != null)
            savedEndParaRPr.Remove();
        if (!string.IsNullOrEmpty(value))
        {
            var newRun = new Drawing.Run(
                new Drawing.RunProperties { Language = "en-US" },
                new Drawing.Text { Text = value });
            para.AppendChild(newRun);
        }
        if (savedEndParaRPr != null)
            para.AppendChild(savedEndParaRPr);
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
                                new Drawing.Text { Text = line })));
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
                            newRun.Text = new Drawing.Text { Text = line };
                            textBody.Append(new Drawing.Paragraph(newRun));
                        }
                    }
                    break;
                }
                case "font":
                case "font.name":
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
                case "font.size":
                    EnsureTableCellHasRun(cell);
                    var sz = (int)Math.Round(ParseFontSize(value) * 100);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sz;
                    }
                    break;
                case "bold":
                case "font.bold":
                    EnsureTableCellHasRun(cell);
                    var b = IsTruthy(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = b;
                    }
                    break;
                case "italic":
                case "font.italic":
                    EnsureTableCellHasRun(cell);
                    var it = IsTruthy(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = it;
                    }
                    break;
                case "color":
                case "font.color":
                {
                    // Build fill before removing old one (atomic)
                    EnsureTableCellHasRun(cell);
                    var cellColorFill = BuildSolidFill(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        rProps.RemoveAllChildren<Drawing.GradientFill>();
                        InsertFillInRunProperties(rProps, (Drawing.SolidFill)cellColorFill.CloneNode(true));
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
                case "font.underline":
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
                case "strikethrough" or "strike" or "font.strike" or "font.strikethrough":
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

        // Ensure DrawingML CT_TextCharacterProperties child order (B-R9-2 / B-R13-2).
        // Our switch arms append children independently (solidFill, latin, ea, ...),
        // which produces a mixed order that OpenXmlValidator flags as schema violations
        // and PowerPoint silently drops out-of-order elements. Reorder once at the end.
        foreach (var rPr in cell.Descendants<Drawing.RunProperties>())
            ReorderDrawingRunProperties(rPr);
        foreach (var endRPr in cell.Descendants<Drawing.EndParagraphRunProperties>())
            ReorderDrawingRunProperties(endRPr);

        return unsupported;
    }

    /// <summary>
    /// Public entry point: resolve shape by path and check for text overflow.
    /// </summary>
    public string? CheckShapeTextOverflow(string path)
    {
        // Parse /slide[N]/shape[M] from path
        var match = System.Text.RegularExpressions.Regex.Match(path, @"/slide\[(\d+)\]/shape\[(\d+)\]");
        if (!match.Success) return null;
        int slideIdx = int.Parse(match.Groups[1].Value);
        int shapeIdx = int.Parse(match.Groups[2].Value);
        var slideParts = _doc.PresentationPart?.SlideParts?.ToList();
        if (slideParts == null || slideIdx < 1 || slideIdx > slideParts.Count) return null;
        var shapeTree = slideParts[slideIdx - 1].Slide?.CommonSlideData?.ShapeTree;
        var shapes = shapeTree?.Elements<Shape>().ToList();
        if (shapes == null || shapeIdx < 1 || shapeIdx > shapes.Count) return null;
        return CheckTextOverflow(shapes[shapeIdx - 1]);
    }

    /// <summary>
    /// Estimates whether the given text will overflow the shape bounds.
    /// Uses per-character width estimation (CJK vs Latin) and reads actual line spacing from the shape.
    /// Returns a warning message if overflow is detected, null otherwise.
    /// </summary>
    internal static string? CheckTextOverflow(Shape shape)
    {
        var text = GetShapeText(shape);
        if (string.IsNullOrEmpty(text)) return null;
        var spPr = shape.ShapeProperties;
        var xfrm = spPr?.Transform2D;
        var extents = xfrm?.Extents;
        if (extents?.Cx == null || extents?.Cy == null) return null;

        long cx = extents.Cx!.Value;  // width in EMU
        long cy = extents.Cy!.Value;  // height in EMU

        const double emuPerPt = 12700.0;
        double shapeWidthPt = cx / emuPerPt;
        double shapeHeightPt = cy / emuPerPt;

        // Read actual margins from BodyProperties, falling back to PPT defaults (0.1in L/R, 0.05in T/B)
        const long defaultLRInset = 91440;   // 0.1in in EMU
        const long defaultTBInset = 45720;   // 0.05in in EMU
        long leftEmu = defaultLRInset, rightEmu = defaultLRInset;
        long topEmu = defaultTBInset, bottomEmu = defaultTBInset;

        var textBody = shape.TextBody;
        var bp = textBody?.BodyProperties;
        if (bp != null)
        {
            if (bp.LeftInset != null) leftEmu = bp.LeftInset.Value;
            if (bp.RightInset != null) rightEmu = bp.RightInset.Value;
            if (bp.TopInset != null) topEmu = bp.TopInset.Value;
            if (bp.BottomInset != null) bottomEmu = bp.BottomInset.Value;
        }

        double usableWidth = shapeWidthPt - (leftEmu + rightEmu) / emuPerPt;
        double usableHeight = shapeHeightPt - (topEmu + bottomEmu) / emuPerPt;
        // If usable area is negative/zero, shape is too small for even its own margins
        double marginPt = (topEmu + bottomEmu) / emuPerPt;
        if (usableWidth <= 0 || usableHeight <= 0)
        {
            // Need at least margins + one line of default text (18pt)
            double defaultLinePt = 18.0;
            double needPt = marginPt + defaultLinePt;
            double minHeightCm = needPt / 72.0 * 2.54;
            // Round up to 0.05cm for cleaner values
            minHeightCm = Math.Ceiling(minHeightCm * 20) / 20.0;
            long minHeightEmu = (long)Math.Round(minHeightCm * 360000.0);
            return $"text overflow: need ≥{defaultLinePt:F0}pt, usable 0pt (shape {shapeHeightPt:F0}pt < margins {marginPt:F0}pt). suggest.height={EmuConverter.FormatEmu(minHeightEmu)}";
        }

        // Collect font size from each paragraph's runs; track the max for line height calculation
        var paragraphs = textBody?.Elements<Drawing.Paragraph>().ToList();
        if (paragraphs == null || paragraphs.Count == 0) return null;

        // Read line spacing from the first paragraph (SpacingPercent as percentage×1000, SpacingPoints as pt×100)
        double lineSpacingMultiplier = 1.0; // default: single spacing (PPT default is 100000 = 1.0x)
        double? fixedLineSpacingPt = null;
        var firstParaProps = paragraphs[0].ParagraphProperties;
        var lsEl = firstParaProps?.GetFirstChild<Drawing.LineSpacing>();
        if (lsEl != null)
        {
            var pct = lsEl.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (pct.HasValue)
                lineSpacingMultiplier = pct.Value / 100000.0;
            var pts = lsEl.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (pts.HasValue)
                fixedLineSpacingPt = pts.Value / 100.0;
        }

        // Read spaceBefore/spaceAfter from first paragraph
        double spaceBeforePt = 0, spaceAfterPt = 0;
        var sbEl = firstParaProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
        if (sbEl.HasValue) spaceBeforePt = sbEl.Value / 100.0;
        var saEl = firstParaProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
        if (saEl.HasValue) spaceAfterPt = saEl.Value / 100.0;

        // Resolve font size: explicit run FontSize → paragraph defRPr → fallback 18pt (PPT default for textboxes)
        double fontSizePt = 0;
        foreach (var para in paragraphs)
        {
            foreach (var run in para.Elements<Drawing.Run>())
            {
                if (run.RunProperties?.FontSize?.HasValue == true)
                {
                    double sz = run.RunProperties.FontSize.Value / 100.0;
                    if (sz > fontSizePt) fontSizePt = sz;
                }
            }
            // Check paragraph default run properties
            var defRp = para.ParagraphProperties?.GetFirstChild<Drawing.DefaultRunProperties>();
            if (defRp?.FontSize?.HasValue == true)
            {
                double sz = defRp.FontSize.Value / 100.0;
                if (sz > fontSizePt) fontSizePt = sz;
            }
        }
        // Also check text body list style level 1 default
        if (fontSizePt <= 0)
        {
            var lstDefRp = textBody?.GetFirstChild<Drawing.ListStyle>()
                ?.GetFirstChild<Drawing.Level1ParagraphProperties>()
                ?.GetFirstChild<Drawing.DefaultRunProperties>();
            if (lstDefRp?.FontSize?.HasValue == true)
                fontSizePt = lstDefRp.FontSize.Value / 100.0;
        }
        if (fontSizePt <= 0) fontSizePt = 18.0; // PPT default for new textboxes

        // Line height: fixed spacing overrides multiplier
        double lineHeight = fixedLineSpacingPt ?? fontSizePt * lineSpacingMultiplier;
        if (lineHeight <= 0) return null;

        // Estimate text width per line using per-character measurement
        var textLines = text.Replace("\\n", "\n").Split('\n');
        int totalLines = 0;
        foreach (var line in textLines)
        {
            if (line.Length == 0)
            {
                totalLines += 1;
                continue;
            }
            // Walk characters, accumulate width, wrap when exceeding usable width
            int linesForSegment = 1;
            double currentLineWidth = 0;
            foreach (char ch in line)
            {
                double charWidth = ParseHelpers.IsCjkOrFullWidth(ch) ? fontSizePt : fontSizePt * 0.55;
                if (currentLineWidth + charWidth > usableWidth && currentLineWidth > 0)
                {
                    linesForSegment++;
                    currentLineWidth = charWidth;
                }
                else
                {
                    currentLineWidth += charWidth;
                }
            }
            totalLines += linesForSegment;
        }

        double estimatedHeight = totalLines * lineHeight
            + spaceBeforePt + spaceAfterPt * Math.Max(textLines.Length - 1, 0);
        if (estimatedHeight > usableHeight * 1.05) // 5% tolerance for rounding
        {
            // Calculate minimum height: estimated text height + margins, converted to cm
            double minHeightCm = (estimatedHeight + marginPt) / 72.0 * 2.54;
            // Round up to 0.05cm for cleaner values
            minHeightCm = Math.Ceiling(minHeightCm * 20) / 20.0;
            long minHeightEmu = (long)Math.Round(minHeightCm * 360000.0);
            return $"text overflow: {totalLines} lines at {fontSizePt:F1}pt need {estimatedHeight:F0}pt, usable {usableHeight:F0}pt. suggest.height={EmuConverter.FormatEmu(minHeightEmu)}";
        }
        return null;
    }

}
