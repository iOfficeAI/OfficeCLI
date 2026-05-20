// Copyright 2025 OfficeCLI (officecli.ai)
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

    // Split documented compound 'line=color[:width[:style]]' form (e.g.
    // "FF0000:1.5:dash") into its parts. The split-key form (line=,
    // lineWidth=, lineDash=) is the underlying canonical; this helper just
    // unpacks the compound surface listed in schemas/help/_shared/shape.json
    // so the documented example works on Add and Set.
    //
    // Inputs without ':' return (value, null, null) unchanged — including
    // "none", named colors, hex (#RRGGBB), scheme tokens (accent1), rgb(...)
    // (commas, not colons). The compound form is unambiguous because no
    // accepted color literal contains ':'.
    private static (string color, string? width, string? dash) SplitCompoundLineValue(string value)
    {
        if (string.IsNullOrEmpty(value) || value.IndexOf(':') < 0)
            return (value, null, null);
        var parts = value.Split(':');
        var color = parts[0];
        var width = parts.Length >= 2 ? parts[1] : null;
        var dash = parts.Length >= 3 ? parts[2] : null;
        return (color, width, dash);
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

    // CONSISTENCY(bcp47-validation): shape regex lives in Core/Bcp47LanguageTag.cs
    // so docx and pptx share one validator. `lang` and `altLang` are the only
    // BCP-47-shaped attrs in rPr; the rest of the long-tail string attrs
    // (kumimoji, bmk, …) stay free-form.

    private static bool IsValidDrawingRunAttrValue(string key, string value)
    {
        if (DrawingRunIntAttrs.Contains(key))
        {
            if (!int.TryParse(value, out var iv)) return false;
            // OOXML ST_TextNonNegativePoint refuses negative kern. Writing
            // kern=-100 produces a file PowerPoint silently rewrites on open.
            // Upper bound mirrors ST_TextPoint's 400000 hundredths-of-a-point
            // ceiling — beyond that PowerPoint clamps on open, so reject up
            // front instead of letting an out-of-band value land on disk.
            if (key == "kern" && (iv < 0 || iv > 400000)) return false;
            // OOXML ST_TextPoint clamps spc to [-400000, 400000] hundredths
            // of a point. Out-of-band values get silently dropped on open.
            if (key == "spc" && (iv < -400000 || iv > 400000)) return false;
            return true;
        }
        if (DrawingRunBoolAttrs.Contains(key))
            return value is "0" or "1" or "true" or "false" or "True" or "False";
        if (key == "u") return DrawingUnderlineEnum.Contains(value);
        if (key == "strike") return DrawingStrikeEnum.Contains(value);
        if (key == "cap") return DrawingCapsEnum.Contains(value);
        if (key is "lang" or "altLang") return OfficeCli.Core.Bcp47LanguageTag.IsValid(value);
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
        bool runContext = false,
        string? unsupportedContextHint = null)
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

        // CONSISTENCY(lang-aliases): Word run rPr has three per-script lang slots
        // (lang.latin / lang.ea / lang.cs). DrawingML CT_TextCharacterProperties
        // exposes only `lang` (and `altLang`) — a single primary-language slot
        // per ECMA-376 §21.1.2.3.9, no per-script split. lang.latin is accepted
        // as an alias for `lang`. lang.ea and lang.cs are explicitly rejected
        // (UNSUPPORTED) rather than silently aliased onto the same attribute,
        // because previously a single Set call with all three keys collapsed
        // to last-write-wins, silently dropping two of the user's values.
        // Users who want CJK/RTL theme fonts should use theme bodyFont.ea/.cs.
        {
            string? latinKey = properties.Keys.FirstOrDefault(k => k.Equals("lang.latin", StringComparison.OrdinalIgnoreCase));
            if (latinKey != null)
            {
                var v = properties[latinKey];
                properties = new Dictionary<string, string>(properties, properties.Comparer);
                properties.Remove(latinKey);
                if (!properties.ContainsKey("lang")) properties["lang"] = v;
            }
        }

        // Raise OOXML short-form attribute names to canonical curated case
        // labels BEFORE dispatch. Without this, short-forms (`sz`, `b`, `i`,
        // `u`, `strike`) fall through to the long-tail attribute writer which
        // writes the raw value verbatim — `sz=14` lands as sz="14" violating
        // ST_TextFontSize (min 100, hundredths of a point) and corrupts the
        // file; `b=true` lands as b="true" instead of the xsd:boolean
        // canonical "1". Mapping early lets the curated cases below handle
        // unit conversion and canonical serialization (FontSize×100, bool→1/0).
        var shortFormMap = new (string Short, string Canonical)[]
        {
            ("sz", "size"),
            ("b", "bold"),
            ("i", "italic"),
            ("u", "underline"),
        };
        foreach (var (shortKey, canonical) in shortFormMap)
        {
            string? matched = properties.Keys.FirstOrDefault(k => k.Equals(shortKey, StringComparison.Ordinal));
            if (matched == null || properties.ContainsKey(canonical)) continue;
            var v = properties[matched];
            properties = new Dictionary<string, string>(properties, properties.Comparer);
            properties.Remove(matched);
            properties[canonical] = v;
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
                    // ST_TextCapsType enum is lowercase; normalize so mixed-case
                    // input ("SMALL", "ALL") does not produce schema-invalid OOXML.
                    var capValue = value.ToLowerInvariant();
                    if (!DrawingCapsEnum.Contains(capValue))
                    {
                        unsupported.Add($"cap (value '{value}' must be one of: none, small, all)");
                        break;
                    }
                    var targetRuns = runs.Count > 0 ? runs : shape.Descendants<Drawing.Run>().ToList();
                    foreach (var run in targetRuns)
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.SetAttribute(new OpenXmlAttribute("", "cap", "", capValue));
                    }
                    break;
                }
                case "text":
                {
                    XmlTextValidator.ValidateOrThrow(value, "text");
                    // CONSISTENCY(text-escape-boundary): \n / \t resolution at
                    // CLI --prop parse; here value has real newlines/tabs.
                    var textLines = value.Split('\n');
                    if (runs.Count == 1 && textLines.Length == 1 && !textLines[0].Contains('\t'))
                    {
                        // Single run, single line, no tabs: just replace text
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
                                AppendLineWithTabs(newPara, textLine, seg =>
                                {
                                    var r = new Drawing.Run();
                                    if (runProps != null)
                                        r.RunProperties = runProps.CloneNode(true) as Drawing.RunProperties;
                                    r.Text = new Drawing.Text { Text = seg };
                                    return r;
                                });
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
                    // Bare 'font' targets Latin + EastAsian (and clears any
                    // prior CS so users get a single coherent typeface).
                    // For per-script control use 'font.latin' / 'font.ea' /
                    // 'font.cs' below (Japanese / Korean / Arabic etc).
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;

                case "font.latin":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;

                case "font.ea" or "font.eastasia" or "font.eastasian":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;

                case "font.cs" or "font.complexscript" or "font.complex":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        rProps.Append(new Drawing.ComplexScriptFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
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
                        var ulMapped = value.ToLowerInvariant() switch
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
                        rProps.Underline = ulMapped;
                        // When the user clears the underline (none/false), any
                        // previously-attached uFill / uFillTx children describe
                        // the colour of a stroke that no longer exists. Leave
                        // them in place and PowerPoint silently renders the
                        // run as underlined again on next open. Strip them so
                        // "underline=none" actually means "no underline".
                        if (ulMapped == Drawing.TextUnderlineValues.None)
                        {
                            rProps.RemoveAllChildren<Drawing.UnderlineFill>();
                            rProps.RemoveAllChildren<Drawing.UnderlineFillText>();
                        }
                    }
                    break;

                case "underlineColor":
                case "underlinecolor":
                case "underline.color":
                case "font.underline.color":
                {
                    // DrawingML: <a:uFill><a:solidFill><a:srgbClr val="…"/></a:solidFill></a:uFill>
                    // Sits between a:uLn and a:latin in CT_TextCharacterProperties
                    // (schema order bucket 6 — see DrawingRunPropChildOrder).
                    // ReorderDrawingRunProperties at the end of this method's
                    // existing post-set cleanup keeps the element in order.
                    var ulHex = ParseHelpers.SanitizeColorForOoxml(value).Rgb;
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.UnderlineFill>();
                        rProps.RemoveAllChildren<Drawing.UnderlineFillText>();
                        var uFill = new Drawing.UnderlineFill(
                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = ulHex }));
                        rProps.AppendChild(uFill);
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;
                }

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

                case "direction" or "dir" or "rtl":
                {
                    // Paragraph reading direction + textbox column direction.
                    // <a:pPr rtl="1"/> reverses character order inside each
                    // paragraph; <a:bodyPr rtlCol="1"/> reverses the column
                    // flow of the text body itself. POI / PowerPoint's UI set
                    // both when the user toggles "Right-to-left text direction"
                    // on a shape, so a single 'direction=rtl' here mirrors the
                    // same intent end-to-end.
                    bool rtl = key.ToLowerInvariant() == "rtl"
                        ? IsTruthy(value)
                        : ParsePptDirectionRtl(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        // Clear semantics: direction=ltr removes the rtl attribute
                        // entirely rather than writing rtl="0" (the schema default
                        // is ltr; an explicit "0" pollutes every freshly-added
                        // paragraph). Mirror Word direction=ltr clear behavior.
                        if (rtl)
                            pProps.RightToLeft = true;
                        else
                            pProps.RightToLeft = null;
                    }
                    var dirBodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    // OpenXml SDK doesn't expose rtlCol as a typed property on
                    // BodyProperties — set the attribute directly. "1"/"0" is
                    // the only canonical xsd:boolean form Office tooling reads.
                    // For ltr (the schema default), strip the attribute rather
                    // than writing rtlCol="0" so a rtl→ltr toggle leaves no
                    // stale explicit-default noise in the XML.
                    if (dirBodyPr != null)
                    {
                        if (rtl)
                            dirBodyPr.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "rtlCol", "", "1"));
                        else
                            dirBodyPr.RemoveAttribute("rtlCol", "");
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
                        {
                            var newGeom = EnsurePresetGeometry(spPr);
                            newGeom.AppendChild(new Drawing.AdjustValueList());
                            newGeom.Preset = ParsePresetShape(value);
                        }
                    break;
                }

                case "geometry" or "path" when key.ToLowerInvariant() != "path" || shape.ShapeProperties != null:
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    // Check if value is a preset shape name (no spaces, no commas, simple identifier)
                    if (!value.Contains(' ') && !value.Contains(',') && !value.Contains('M'))
                    {
                        // Treat as preset shape name. Use the strict variant so
                        // an unrecognised name surfaces as unsupported_property
                        // instead of silently rewriting the geometry to a
                        // rectangle (the Add-side fallback's intent — keep a
                        // batch import alive on one bad preset — is wrong for a
                        // single-property Set: the caller asked for a specific
                        // shape and deserves to know the name didn't match).
                        if (!TryParsePresetShape(value, out var preset))
                        {
                            unsupported.Add($"{key}={value} (unknown preset shape name)");
                            break;
                        }
                        spPr.RemoveAllChildren<Drawing.CustomGeometry>();
                        var existingGeom = spPr.GetFirstChild<Drawing.PresetGeometry>();
                        if (existingGeom != null)
                            existingGeom.Preset = preset;
                        else
                            {
                            var newGeom = EnsurePresetGeometry(spPr);
                            newGeom.AppendChild(new Drawing.AdjustValueList());
                            newGeom.Preset = preset;
                        }
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
                    // Schema documents compound form 'color[:width[:style]]'
                    // (schemas/help/_shared/shape.json) — split here and
                    // fall through the existing single-part code paths so
                    // there's one place doing the OOXML mutation.
                    var (lineColor, lineWidthPart, lineDashPart) = SplitCompoundLineValue(value);
                    // Build fill before removing old one (atomic)
                    OpenXmlElement newLineFill = lineColor.Equals("none", StringComparison.OrdinalIgnoreCase)
                        ? new Drawing.NoFill()
                        : BuildSolidFill(lineColor);
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
                    if (lineWidthPart != null)
                        outline.Width = Core.EmuConverter.ParseLineWidth(lineWidthPart);
                    if (lineDashPart != null)
                    {
                        outline.RemoveAllChildren<Drawing.PresetDash>();
                        outline.AppendChild(new Drawing.PresetDash { Val = ParseLineDashValue(lineDashPart) });
                    }
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

                case "line.gradient" or "linegradient":
                {
                    // Gradient stroke. Reader emits this for any <a:ln> whose
                    // child is GradientFill; without a setter the round trip
                    // dropped the gradient and replayed as a bare <a:ln/>
                    // (theme thin black stroke). Use the same gradient-spec
                    // grammar the shape fill accepts ("color1-color2[:angle]"
                    // or full multi-stop form).
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.SolidFill>();
                    outline.RemoveAllChildren<Drawing.NoFill>();
                    outline.RemoveAllChildren<Drawing.GradientFill>();
                    var grad = BuildGradientFill(value);
                    // CT_LineProperties schema: fill (solidFill/noFill/gradFill/pattFill) → prstDash → ...
                    var prstDashAnchor = outline.GetFirstChild<Drawing.PresetDash>();
                    if (prstDashAnchor != null)
                        outline.InsertBefore(grad, prstDashAnchor);
                    else
                        outline.PrependChild(grad);
                    break;
                }

                case "linedash" or "line.dash":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    outline.AppendChild(new Drawing.PresetDash { Val = ParseLineDashValue(value) });
                    break;
                }

                // lineCap → <a:ln cap="..."> attribute (was silently dropped).
                case "linecap" or "line.cap":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.CapType = value.ToLowerInvariant() switch
                    {
                        "round" or "rnd" => Drawing.LineCapValues.Round,
                        "flat" => Drawing.LineCapValues.Flat,
                        "square" or "sq" => Drawing.LineCapValues.Square,
                        _ => throw new ArgumentException($"Invalid 'lineCap' value: '{value}'. Valid values: round, flat, square.")
                    };
                    break;
                }
                // lineJoin → child element <a:round/>|<a:bevel/>|<a:miter/> (was silently dropped).
                case "linejoin" or "line.join":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.Round>();
                    outline.RemoveAllChildren<Drawing.LineJoinBevel>();
                    outline.RemoveAllChildren<Drawing.Miter>();
                    OpenXmlElement joinEl = value.ToLowerInvariant() switch
                    {
                        "round" => new Drawing.Round(),
                        "bevel" => new Drawing.LineJoinBevel(),
                        "miter" => new Drawing.Miter(),
                        _ => throw new ArgumentException($"Invalid 'lineJoin' value: '{value}'. Valid values: round, bevel, miter.")
                    };
                    // CT_LineProperties schema: ... → prstDash → (round|bevel|miter) → headEnd → tailEnd
                    var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                    if (headEnd != null) outline.InsertBefore(joinEl, headEnd);
                    else
                    {
                        var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                        if (tailEnd != null) outline.InsertBefore(joinEl, tailEnd);
                        else outline.AppendChild(joinEl);
                    }
                    break;
                }
                // cmpd → <a:ln cmpd="..."> attribute (was silently dropped).
                case "cmpd" or "compoundline" or "line.compound":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.CompoundLineType = value switch
                    {
                        var s when s.Equals("sng", StringComparison.OrdinalIgnoreCase) || s.Equals("single", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Single,
                        var s when s.Equals("dbl", StringComparison.OrdinalIgnoreCase) || s.Equals("double", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Double,
                        var s when s.Equals("thickThin", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.ThickThin,
                        var s when s.Equals("thinThick", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.ThinThick,
                        var s when s.Equals("tri", StringComparison.OrdinalIgnoreCase) || s.Equals("triple", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Triple,
                        _ => throw new ArgumentException($"Invalid 'cmpd' value: '{value}'. Valid values: sng, dbl, thickThin, thinThick, tri.")
                    };
                    break;
                }
                // lineAlign → <a:ln algn="..."> attribute (was silently dropped).
                case "linealign" or "line.align":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.Alignment = value.ToLowerInvariant() switch
                    {
                        "ctr" or "center" => Drawing.PenAlignmentValues.Center,
                        "in" or "inset" => Drawing.PenAlignmentValues.Insert,
                        _ => throw new ArgumentException($"Invalid 'lineAlign' value: '{value}'. Valid values: ctr, in.")
                    };
                    break;
                }
                // head/tail end arrowheads on shape outlines (CT_LineProperties allows them
                // on any outline, not just connectors). Previously dropped.
                case "headend" or "arrowstart":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.HeadEnd>();
                    var newHeadEnd = new Drawing.HeadEnd { Type = ParseLineEndType(value) };
                    // CT_LineProperties: ... → headEnd → tailEnd
                    var existingTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                    if (existingTailEnd != null) outline.InsertBefore(newHeadEnd, existingTailEnd);
                    else outline.AppendChild(newHeadEnd);
                    break;
                }
                case "tailend" or "arrowend":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.TailEnd>();
                    outline.AppendChild(new Drawing.TailEnd { Type = ParseLineEndType(value) });
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
                    // The percentage shorthand (>1 treated as 0-100 percent)
                    // was silently accepting ambiguous values in the (1, 2)
                    // range: opacity=1.5 → divided to 0.015, written as
                    // alpha=1500 (≈1.5% visible) instead of being rejected
                    // outright. 1.5 isn't a meaningful percentage (a user
                    // typing "1.5" almost certainly meant the decimal form,
                    // which is out of range) AND isn't a meaningful decimal
                    // (>1). Treat the gap as a clear input error rather than
                    // a silent /100 division.
                    if (opacityVal > 1.0 && opacityVal < 2.0)
                        throw new ArgumentException($"Invalid 'opacity' value: '{value}'. Expected 0.0-1.0 as decimal or 2-100 as percent (use 0-1 for the decimal form; values in (1, 2) are ambiguous).");
                    if (opacityVal > 1.0) opacityVal /= 100.0; // treat >=2 as percentage (e.g. 30 → 0.30)
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
                    // OOXML ST_TextPoint: hundredths of a point, range
                    // [-400000, 400000] (== [-4000pt, 4000pt]). PowerPoint
                    // silently rewrites out-of-band values to default on open.
                    if (spcVal < -400000 || spcVal > 400000)
                        throw new ArgumentException($"Invalid 'charspacing' value: '{value}': OOXML ST_TextPoint range is [-4000pt, 4000pt].");
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
                        // CONSISTENCY(schema-order-pptx): pPr children must follow
                        // CT_TextParagraphProperties order or PowerPoint silently
                        // drops them. See PowerPointHandler.Helpers.cs.
                        InsertPPrChild(pProps, lnSpcElem);
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
                        InsertPPrChild(pProps, new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = sbIntVal }));
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
                        InsertPPrChild(pProps, new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = saIntVal }));
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
                        case "true" or "normal" or "normautofit" or "auto": bodyPr.AppendChild(new Drawing.NormalAutoFit()); break;
                        case "shape" or "spautofit" or "resize": bodyPr.AppendChild(new Drawing.ShapeAutoFit()); break;
                        case "false" or "none": bodyPr.AppendChild(new Drawing.NoAutoFit()); break;
                        // 'shrink' previously aliased to 'normal' (same plain
                        // <a:normAutofit/> with no fontScale/lnSpcReduction
                        // attributes), so callers asking for shrink-on-overflow
                        // got identical XML to autofit=normal. Reject the alias
                        // instead of silently lying — implementing real shrink
                        // requires picking fontScale/lnSpcReduction values per
                        // overflow geometry, which is its own feature. Callers
                        // wanting shrink-on-overflow today should use
                        // autofit=normal and tune fontScale via raw-set.
                        case "shrink": throw new ArgumentException($"Invalid autofit value: 'shrink'. PowerPoint's shrink-on-overflow requires fontScale/lnSpcReduction attributes that officecli does not synthesize; use autofit=normal (plain normAutofit) and tune via raw-set if needed. Valid values: true/normal, shape/resize, false/none.");
                        default: throw new ArgumentException($"Invalid autofit value: '{value}'. Valid values: true/normal, shape/resize, false/none.");
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
                    if (nvPr != null)
                    {
                        XmlTextValidator.ValidateOrThrow(value, "name");
                        nvPr.Name = value;
                    }
                    else unsupported.Add(key);
                    break;
                }

                case "alt" or "alttext" or "description":
                {
                    var nvPr = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                    if (nvPr != null)
                    {
                        XmlTextValidator.ValidateOrThrow(value, "alttext");
                        nvPr.Description = value;
                    }
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
                            // Invalid value for a typed OOXML rPr attribute (kern=abc,
                            // u=GARBAGE, b=2, etc.) — throw rather than collecting
                            // into `unsupported`, which is reserved for unknown keys
                            // (handler-doesn't-implement). Invalid values silently
                            // accepted would corrupt the document and fail strict
                            // OOXML validation downstream.
                            // CONSISTENCY(bcp47-error): mirror the docx lang error
                            // shape so agents see one message across handlers
                            // (WordHandler.Helpers.cs ~1671).
                            if (key is "lang" or "altLang")
                                throw new ArgumentException(
                                    $"Invalid BCP-47 language tag for {key}: '{value}'. Expected a tag like 'en-US', 'ja-JP', or 'ar-SA' (RFC 5646: <= {OfficeCli.Core.Bcp47LanguageTag.MaxLength} chars, primary subtag 2-3 letters, then hyphen-separated subtags).");
                            if (key == "kern" && int.TryParse(value, out var kv) && kv < 0)
                                throw new ArgumentException(
                                    $"Invalid kern '{value}': OOXML ST_TextNonNegativePoint requires kern >= 0 (hundredths of a point).");
                            if (key == "spc" && int.TryParse(value, out var sv) && (sv < -400000 || sv > 400000))
                                throw new ArgumentException(
                                    $"Invalid spc '{value}': OOXML ST_TextPoint range is [-400000, 400000] hundredths of a point.");
                            throw new ArgumentException(
                                $"Invalid value for OOXML rPr/{key}: '{value}'.");
                        }
                        handledByRun = true;
                        // CONSISTENCY(lang-clear): empty lang/altLang clears the
                        // attribute entirely (mirrors Word lang.latin="" semantics).
                        // Writing lang="" produces invalid OOXML — Office and
                        // BCP-47 require either a non-empty tag or no attribute.
                        bool clearAttr = (key.Equals("lang", StringComparison.OrdinalIgnoreCase)
                                          || key.Equals("altLang", StringComparison.OrdinalIgnoreCase))
                                         && string.IsNullOrEmpty(value);
                        foreach (var run in runs)
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            if (clearAttr)
                                rPr.RemoveAttribute(key, "");
                            else
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
                        {
                            // Context-aware guidance: run/paragraph callers route
                            // here via fallback but the prop list they accept is a
                            // subset of shape's. Without the hint the error
                            // misleadingly cites x/y/width/height/etc.
                            var msg = unsupportedContextHint
                                ?? "valid shape props: text, bold, italic, underline, color, fill, size, font, gradient, line, opacity, align, valign, x, y, width, height, rotation, name, link, animation, formula, geometry, preset, shadow, glow, reflection, softEdge, pattern, flip, flipH, flipV";
                            unsupported.Add($"{key} ({msg})");
                        }
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
        // CT_TextParagraph schema: pPr? (br | r | fld)* endParaRPr? — endParaRPr,
        // when present, must be last. AddTable seeds empty cells with just an
        // <a:endParaRPr/>, so a naive Append lands the new run AFTER it and
        // produces Sch_UnexpectedElementContentExpectingComplex.
        var endParaRPr = para.GetFirstChild<Drawing.EndParagraphRunProperties>();
        if (endParaRPr != null)
            para.InsertBefore(run, endParaRPr);
        else
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
                    XmlTextValidator.ValidateOrThrow(value, "text");
                    var textBody = cell.TextBody;
                    // CONSISTENCY(text-escape-boundary): see CommandBuilder.
                    var lines = value.Split('\n');
                    if (textBody == null)
                    {
                        textBody = new Drawing.TextBody(
                            new Drawing.BodyProperties(), new Drawing.ListStyle());
                        foreach (var line in lines)
                        {
                            var para = new Drawing.Paragraph();
                            AppendLineWithTabs(para, line, seg => new Drawing.Run(
                                new Drawing.RunProperties { Language = "en-US" },
                                new Drawing.Text { Text = seg }));
                            textBody.AppendChild(para);
                        }
                        cell.PrependChild(textBody);
                    }
                    else
                    {
                        var firstRun = textBody.Descendants<Drawing.Run>().FirstOrDefault();
                        var runProps = firstRun?.RunProperties?.CloneNode(true) as Drawing.RunProperties;
                        // Snapshot the existing first paragraph's properties
                        // (algn, lvl, marL, indent, …) so a single set call
                        // that bundles `align=center` with `text='X'` doesn't
                        // lose the alignment when text rebuilds the
                        // paragraph tree. Iteration order on a Dictionary is
                        // insertion order on .NET but callers shouldn't have
                        // to know that — preserve align by cloning the
                        // existing pPr BEFORE wiping paragraphs, then
                        // re-attach on each rebuilt paragraph.
                        var firstPara = textBody.GetFirstChild<Drawing.Paragraph>();
                        var savedPPr = firstPara?.ParagraphProperties?.CloneNode(true) as Drawing.ParagraphProperties;
                        textBody.RemoveAllChildren<Drawing.Paragraph>();
                        foreach (var line in lines)
                        {
                            var para = new Drawing.Paragraph();
                            if (savedPPr != null)
                                para.ParagraphProperties = savedPPr.CloneNode(true) as Drawing.ParagraphProperties;
                            AppendLineWithTabs(para, line, seg =>
                            {
                                var r = new Drawing.Run();
                                r.RunProperties = runProps != null
                                    ? runProps.CloneNode(true) as Drawing.RunProperties
                                    : new Drawing.RunProperties { Language = "en-US" };
                                r.Text = new Drawing.Text { Text = seg };
                                return r;
                            });
                            textBody.Append(para);
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
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;
                case "font.latin":
                    EnsureTableCellHasRun(cell);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;
                case "font.ea" or "font.eastasia" or "font.eastasian":
                    EnsureTableCellHasRun(cell);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;
                case "font.cs" or "font.complexscript" or "font.complex":
                    EnsureTableCellHasRun(cell);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        rProps.Append(new Drawing.ComplexScriptFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
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
                case "background":
                case "gradient":
                {
                    // CONSISTENCY(fill-gradient-shorthand): accept linear
                    // ("C1-C2[-angle]"), radial ("radial:C1-C2[-focus]"),
                    // path ("path:C1-C2[-focus]"), and "LINEAR;C1;C2;angle"
                    // shorthand directly on fill= — matches the shape and
                    // slide-background contract via the shared
                    // NormalizeGradientValue / IsGradientColorString /
                    // BuildGradientFill helpers in Fill.cs.
                    // `gradient=` is the canonical key shape-level uses
                    // (shape Set dispatches to ApplyGradientFill); mirror
                    // it on cells so dump/replay and direct callers work.
                    // Build new fill element BEFORE removing old one (atomic: no data loss on invalid color)
                    var normalizedCellFill = NormalizeGradientValue(value);
                    OpenXmlElement newCellFill;
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        newCellFill = new Drawing.NoFill();
                    }
                    else if (normalizedCellFill.StartsWith("radial:", StringComparison.OrdinalIgnoreCase)
                          || normalizedCellFill.StartsWith("path:", StringComparison.OrdinalIgnoreCase)
                          || IsGradientColorString(normalizedCellFill))
                    {
                        newCellFill = BuildGradientFill(normalizedCellFill);
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
                case "align" or "alignment" or "halign":
                {
                    foreach (var para in cell.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = ParseTextAlignment(value);
                    }
                    break;
                }
                case "direction" or "dir" or "rtl":
                {
                    // Mirror the shape-level direction handler: cascade
                    // <a:pPr rtl="1"/> to every paragraph in the cell.
                    // bodyPr/rtlCol is not relevant for table cells (each
                    // cell has its own txBody but no column-flow attribute).
                    bool rtl = key.ToLowerInvariant() == "rtl"
                        ? IsTruthy(value)
                        : ParsePptDirectionRtl(value);
                    foreach (var para in cell.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        // Clear semantics: direction=ltr strips the attribute.
                        if (rtl) pProps.RightToLeft = true;
                        else pProps.RightToLeft = null;
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
                {
                    // CONSISTENCY(merge-continuation): a CT_TableCell with
                    // gridSpan=N is only a valid horizontal merge if the next
                    // (N-1) cells in the same row carry hMerge=true. Without
                    // them PowerPoint renders the row un-merged. Mirror the
                    // merge.right case (below) so plain `gridSpan=N` produces
                    // a working merge instead of a half-applied one.
                    var span = ParseHelpers.SafeParseInt(value, "gridspan");
                    // BUG-R6-B: validate span ≥ 1 and not exceeding row width.
                    if (span < 1)
                        throw new ArgumentException($"Invalid colspan: '{value}'. Must be >= 1.");
                    if (cell.Parent is Drawing.TableRow gsRowChk)
                    {
                        var gsCellsChk = gsRowChk.Elements<Drawing.TableCell>().ToList();
                        var gsIdxChk = gsCellsChk.IndexOf(cell);
                        var remaining = gsCellsChk.Count - gsIdxChk;
                        if (span > remaining)
                            throw new ArgumentException($"Invalid colspan: {span} exceeds remaining columns ({remaining}) from this cell.");
                    }
                    if (span > 1)
                    {
                        cell.GridSpan = new DocumentFormat.OpenXml.Int32Value(span);
                        if (cell.Parent is Drawing.TableRow gsRow)
                        {
                            var gsCells = gsRow.Elements<Drawing.TableCell>().ToList();
                            var gsIdx = gsCells.IndexOf(cell);
                            for (int mi = gsIdx + 1; mi < gsIdx + span && mi < gsCells.Count; mi++)
                                gsCells[mi].HorizontalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                            // BUG-R5-table-merge BUG-8: when the anchor cell
                            // already has rowSpan>1, the corner cells in each
                            // continuation row need both hMerge=true (covered
                            // by this gridSpan) and vMerge=true (covered by
                            // the prior rowSpan). CONSISTENCY(table-merge-2d).
                            int gsAnchorRowSpan = cell.RowSpan?.Value ?? 1;
                            if (gsAnchorRowSpan > 1 && gsRow.Parent is Drawing.Table gsAnchorTbl)
                            {
                                var gsRows = gsAnchorTbl.Elements<Drawing.TableRow>().ToList();
                                var gsRowIdx = gsRows.IndexOf(gsRow);
                                for (int ri = gsRowIdx + 1; ri < gsRowIdx + gsAnchorRowSpan && ri < gsRows.Count; ri++)
                                {
                                    var rowCells = gsRows[ri].Elements<Drawing.TableCell>().ToList();
                                    for (int ci = gsIdx + 1; ci < gsIdx + span && ci < rowCells.Count; ci++)
                                    {
                                        rowCells[ci].HorizontalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                                        rowCells[ci].VerticalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // colspan=1 → un-merge: drop any prior gridSpan attribute (omitted = 1)
                        // and clear hMerge on the cells previously covered by this anchor.
                        var prevSpan = cell.GridSpan?.Value ?? 1;
                        cell.GridSpan = null;
                        if (prevSpan > 1 && cell.Parent is Drawing.TableRow gsRow1)
                        {
                            var gsCells1 = gsRow1.Elements<Drawing.TableCell>().ToList();
                            var gsIdx1 = gsCells1.IndexOf(cell);
                            for (int mi = gsIdx1 + 1; mi < gsIdx1 + prevSpan && mi < gsCells1.Count; mi++)
                                gsCells1[mi].HorizontalMerge = null;
                        }
                    }
                    break;
                }
                case "rowspan":
                {
                    var rsSpan = ParseHelpers.SafeParseInt(value, "rowspan");
                    // BUG-R6-B: validate rowspan ≥ 1 and not exceeding remaining rows.
                    if (rsSpan < 1)
                        throw new ArgumentException($"Invalid rowspan: '{value}'. Must be >= 1.");
                    if (cell.Parent is Drawing.TableRow rsRowChk && rsRowChk.Parent is Drawing.Table rsTblChk)
                    {
                        var rsRows = rsTblChk.Elements<Drawing.TableRow>().ToList();
                        var rsRowIdx = rsRows.IndexOf(rsRowChk);
                        var remainingRows = rsRows.Count - rsRowIdx;
                        if (rsSpan > remainingRows)
                            throw new ArgumentException($"Invalid rowspan: {rsSpan} exceeds remaining rows ({remainingRows}) from this cell.");
                    }
                    cell.RowSpan = new DocumentFormat.OpenXml.Int32Value(rsSpan);
                    // BUG-R1-table-merge: rowSpan on the anchor cell is not
                    // sufficient — every continuation cell directly below
                    // must carry vMerge=true or PowerPoint treats the cells
                    // as independent. CONSISTENCY(table-merge-anchor):
                    // mirrors merge.down case below.
                    if (rsSpan > 1
                        && cell.Parent is Drawing.TableRow rsAnchorRow
                        && rsAnchorRow.Parent is Drawing.Table rsAnchorTbl)
                    {
                        var rsRows2 = rsAnchorTbl.Elements<Drawing.TableRow>().ToList();
                        var rsRowIdx2 = rsRows2.IndexOf(rsAnchorRow);
                        var rsCells2 = rsAnchorRow.Elements<Drawing.TableCell>().ToList();
                        var rsColIdx2 = rsCells2.IndexOf(cell);
                        // BUG-R5-table-merge BUG-8: when anchor already has
                        // gridSpan>1, corner continuation cells in each
                        // below-row need both vMerge (this rowSpan) and
                        // hMerge (the prior gridSpan). CONSISTENCY(table-merge-2d).
                        int rsAnchorGridSpan = cell.GridSpan?.Value ?? 1;
                        for (int ri = rsRowIdx2 + 1; ri < rsRowIdx2 + rsSpan && ri < rsRows2.Count; ri++)
                        {
                            var belowCells = rsRows2[ri].Elements<Drawing.TableCell>().ToList();
                            if (rsColIdx2 < belowCells.Count)
                                belowCells[rsColIdx2].VerticalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                            for (int ci = rsColIdx2 + 1; ci < rsColIdx2 + rsAnchorGridSpan && ci < belowCells.Count; ci++)
                            {
                                belowCells[ci].HorizontalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                                belowCells[ci].VerticalMerge = new DocumentFormat.OpenXml.BooleanValue(true);
                            }
                        }
                    }
                    break;
                }
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
                case "underlineColor":
                case "underlinecolor":
                case "underline.color":
                case "font.underline.color":
                {
                    EnsureTableCellHasRun(cell);
                    var ulHex = ParseHelpers.SanitizeColorForOoxml(value).Rgb;
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.UnderlineFill>();
                        rProps.RemoveAllChildren<Drawing.UnderlineFillText>();
                        var uFill = new Drawing.UnderlineFill(
                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = ulHex }));
                        rProps.AppendChild(uFill);
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;
                }
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
                    // Sub-key axis selectors: border.width / border.color /
                    // border.dash (and the edge-qualified .left.width etc).
                    // Without this routing, "border.width=-5" fell through to
                    // the loose space-branch and set borderColor="-5" — a
                    // silent corruption. Detect the sub-key suffix and route
                    // the value to the matching axis, then reject negatives
                    // for width to match OOXML ST_LineWidth's non-negative
                    // requirement (mirrors line.width's ParseEmuAsInt guard).
                    bool isWidthOnly = k.EndsWith(".width", StringComparison.Ordinal);
                    bool isColorOnly = k.EndsWith(".color", StringComparison.Ordinal);
                    bool isDashOnly = k.EndsWith(".dash", StringComparison.Ordinal);
                    if (!isNone && (isWidthOnly || isColorOnly || isDashOnly))
                    {
                        if (isWidthOnly)
                        {
                            // ParseLineWidth treats bare numbers as pt,
                            // routes through ParseEmuAsInt which rejects
                            // negatives and INT32 overflow. Catch the
                            // rejection and re-raise with a border-shaped
                            // message so the caller sees the key, not a
                            // generic "dimension" diagnostic.
                            try
                            {
                                borderWidth = Core.EmuConverter.ParseLineWidth(value);
                            }
                            catch (ArgumentException)
                            {
                                throw new ArgumentException($"Invalid border width: '{value}' (must be >= 0).");
                            }
                        }
                        else if (isColorOnly)
                        {
                            borderColor = value.TrimStart('#').ToUpperInvariant();
                        }
                        else
                        {
                            var d = value.ToLowerInvariant();
                            if (d is "solid" or "dot" or "dash" or "lgdash" or "dashdot" or "sysdot" or "sysdash")
                                borderDash = d;
                            else
                                throw new ArgumentException($"Invalid border dash value: '{value}'. Valid values: solid, dot, dash, lgDash, dashDot, sysDot, sysDash.");
                        }
                    }
                    else if (!isNone)
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
                                // OOXML ST_LineWidth requires >= 0. ParseEmu
                                // returns a signed long with no sign guard;
                                // reject negatives here so border.width=-5 no
                                // longer silently writes a negative w
                                // attribute that downstream readers truncate.
                                // Mirrors line.width's ParseLineWidth path
                                // (ParseEmuAsInt rejects negatives) and the
                                // padding/margin guards below.
                                if (borderWidth.Value < 0)
                                    throw new ArgumentException($"Invalid border width: '{scParts[1]}' (must be >= 0).");
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
                                {
                                    borderWidth = Core.EmuConverter.ParseEmu(bp);
                                    if (borderWidth.Value < 0)
                                        throw new ArgumentException($"Invalid border width: '{bp}' (must be >= 0).");
                                }
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
                    // BUG-R6-E: cell padding/margin must be >= 0 (OOXML schema requirement).
                    static int NonNegEmu(string v, string side)
                    {
                        var e = (int)ParseEmu(v.Trim());
                        if (e < 0) throw new ArgumentException($"Invalid cell {side}: '{v.Trim()}' (must be >= 0).");
                        return e;
                    }
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    var parts = value.Split(',');
                    if (parts.Length == 1)
                    {
                        var emu = NonNegEmu(parts[0], "padding");
                        tcPrM.LeftMargin = emu;
                        tcPrM.RightMargin = emu;
                        tcPrM.TopMargin = emu;
                        tcPrM.BottomMargin = emu;
                    }
                    else if (parts.Length == 4)
                    {
                        tcPrM.LeftMargin = NonNegEmu(parts[0], "padding.left");
                        tcPrM.TopMargin = NonNegEmu(parts[1], "padding.top");
                        tcPrM.RightMargin = NonNegEmu(parts[2], "padding.right");
                        tcPrM.BottomMargin = NonNegEmu(parts[3], "padding.bottom");
                    }
                    else if (parts.Length == 2)
                    {
                        var h = NonNegEmu(parts[0], "padding.horizontal");
                        var v = NonNegEmu(parts[1], "padding.vertical");
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
                    var v = (int)ParseEmu(value);
                    if (v < 0) throw new ArgumentException($"Invalid cell padding.left: '{value}' (must be >= 0).");
                    tcPrM.LeftMargin = v;
                    break;
                }
                case "margin.right" or "padding.right":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    var v = (int)ParseEmu(value);
                    if (v < 0) throw new ArgumentException($"Invalid cell padding.right: '{value}' (must be >= 0).");
                    tcPrM.RightMargin = v;
                    break;
                }
                case "margin.top" or "padding.top":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    var v = (int)ParseEmu(value);
                    if (v < 0) throw new ArgumentException($"Invalid cell padding.top: '{value}' (must be >= 0).");
                    tcPrM.TopMargin = v;
                    break;
                }
                case "margin.bottom" or "padding.bottom":
                {
                    var tcPrM = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    var v = (int)ParseEmu(value);
                    if (v < 0) throw new ArgumentException($"Invalid cell padding.bottom: '{value}' (must be >= 0).");
                    tcPrM.BottomMargin = v;
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
                        // CONSISTENCY(schema-order-pptx): see Helpers.InsertPPrChild.
                        InsertPPrChild(pProps, ls);
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
                        InsertPPrChild(pProps, sb);
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
                        InsertPPrChild(pProps, sa);
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
                        // CONSISTENCY(opacity-clamp): values in (1, 2) are
                        // ambiguous — explicit-reject upfront before the /100
                        // would silently coerce 1.5 to 0.015. See the shape
                        // opacity path for the full rationale.
                        if (opacityVal > 1.0 && opacityVal < 2.0)
                            throw new ArgumentException($"Invalid 'opacity' value: '{value}'. Expected 0.0-1.0 as decimal or 2-100 as percent (values in (1, 2) are ambiguous).");
                        if (opacityVal > 1.0) opacityVal /= 100.0; // treat >=2 as percentage (e.g. 50 → 0.50)
                        if (opacityVal < 0.0 || opacityVal > 1.0)
                            throw new ArgumentException($"Invalid 'opacity' value: '{value}'. Expected 0.0-1.0 (or 0-100 as percent).");
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

            // autoFit=normal: PowerPoint shrinks text (fontScale) at render
            // time to fit the shape — there is no real overflow.
            // autoFit=shape: the shape resizes to fit the text — also no
            // overflow. Skip the size-based check in both cases.
            if (bp.GetFirstChild<Drawing.NormalAutoFit>() != null
                || bp.GetFirstChild<Drawing.ShapeAutoFit>() != null)
            {
                return null;
            }
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
        // CONSISTENCY(escape-sequences): both \n and \t are interpreted in text=
        // properties cross-handler; resolve here so width estimation matches what
        // PowerPoint will actually render.
        var textLines = text.Split('\n');
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
