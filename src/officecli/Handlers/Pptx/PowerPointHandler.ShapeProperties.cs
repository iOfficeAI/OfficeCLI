// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    // CONSISTENCY(merge-bool-form): real PowerPoint emits xsd:boolean attrs
    // hMerge / vMerge as "1" / "0". OpenXml SDK's BooleanValue(true)
    // serialises as "true" by default — diff-friendly for tests but a visible
    // drift versus the source XML. Wrap construction in a helper that pins
    // the lexical form to "1" so dump→replay round-trips the canonical
    // attribute value (also accepted by PowerPoint, but minimises noise).
    private static DocumentFormat.OpenXml.BooleanValue OneOnBool()
        => new DocumentFormat.OpenXml.BooleanValue(true) { InnerText = "1" };

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
    // Stamp a <a:normAutofit>'s fontScale / lnSpcReduction from sibling props.
    // PowerPoint authors these on "shrink text on overflow" boxes (e.g.
    // fontScale="92500" = render at 92.5%); without round-tripping them the box
    // rebuilt at 100%, so text overflowed/re-flowed across the whole deck.
    // Values are OOXML thousandths-of-percent (92500); a trailing "%" form
    // ("92.5%") is also accepted.
    private static Drawing.NormalAutoFit ApplyNormalAutoFitScale(Drawing.NormalAutoFit naf, Dictionary<string, string> properties)
    {
        if ((properties.TryGetValue("fontScale", out var fs) || properties.TryGetValue("fontscale", out fs))
            && TryParseScalePerMille(fs, out var fsv))
            naf.FontScale = fsv;
        if ((properties.TryGetValue("lnSpcReduction", out var lr) || properties.TryGetValue("lnspcreduction", out lr)
                || properties.TryGetValue("lineSpaceReduction", out lr) || properties.TryGetValue("linespacereduction", out lr)
                || properties.TryGetValue("lineSpacingReduction", out lr) || properties.TryGetValue("linespacingreduction", out lr))
            && TryParseScalePerMille(lr, out var lrv))
            naf.LineSpaceReduction = lrv;
        return naf;
    }

    private static bool TryParseScalePerMille(string? s, out int val)
    {
        val = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s.Trim();
        if (s.EndsWith("%"))
            return double.TryParse(s.TrimEnd('%').Trim(), System.Globalization.NumberStyles.Float,
                       System.Globalization.CultureInfo.InvariantCulture, out var d)
                   && (val = (int)Math.Round(d * 1000)) >= 0;
        if (!int.TryParse(s, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out val)) return false;
        // R10-3 (LEAD: Option B): a bare integer in 0..100 is a PERCENT (75 → 75000),
        // matching the user-facing percent vocabulary. Values >100 are already raw
        // OOXML thousandths-of-percent (e.g. dump→replay of fontScale="92500").
        if (val >= 0 && val <= 100) val *= 1000;
        return true;
    }

    private static List<string> SetRunOrShapeProperties(
        Dictionary<string, string> properties, List<Drawing.Run> runs, Shape shape, OpenXmlPart? part = null,
        bool runContext = false,
        string? unsupportedContextHint = null,
        ICollection<string>? unrecognizedLatex = null)
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

        // RC1: an EMPTY placeholder (no runs — common on layout/master title and
        // body placeholders that only carry prompt text via inheritance) used to
        // silently drop run-format props (size/color/bold/…): every per-key
        // `foreach (var run in runs)` loop ran zero iterations, yet Set still
        // reported "Updated". Fix: when there are no runs but run-format props are
        // present, run the existing per-case logic against a detached SCRATCH run,
        // then transplant the resulting RunProperties children onto the first
        // paragraph's <a:defRPr> (the OOXML home for default run formatting on a
        // runless placeholder). This reuses every existing case unchanged.
        Drawing.Run? defRPrScratchRun = null;
        Drawing.Paragraph? defRPrTargetPara = null;
        if (runs.Count == 0 && properties.Keys.Any(k => RunFormatDefRPrKeys.Contains(k.ToLowerInvariant())))
        {
            var txBody = shape.TextBody;
            if (txBody == null)
            {
                txBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                    new Drawing.BodyProperties(), new Drawing.ListStyle());
                shape.TextBody = txBody;
            }
            defRPrTargetPara = txBody.GetFirstChild<Drawing.Paragraph>();
            if (defRPrTargetPara == null)
            {
                defRPrTargetPara = new Drawing.Paragraph();
                txBody.AppendChild(defRPrTargetPara);
            }
            defRPrScratchRun = new Drawing.Run(new Drawing.RunProperties(), new Drawing.Text(string.Empty));
            runs = new List<Drawing.Run> { defRPrScratchRun };
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
                    // Text-less shape: fall back to the first paragraph's endParaRPr
                    // so cap on an empty shape isn't silently dropped (mirrors the
                    // RunPropTargets endParaRPr fallback used by bold/italic in Add).
                    if (targetRuns.Count == 0)
                    {
                        var firstParaCap = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                        if (firstParaCap != null)
                        {
                            var endRPr = firstParaCap.GetFirstChild<Drawing.EndParagraphRunProperties>();
                            if (endRPr == null)
                            {
                                endRPr = new Drawing.EndParagraphRunProperties { Language = "en-US" };
                                firstParaCap.AppendChild(endRPr);
                            }
                            endRPr.SetAttribute(new OpenXmlAttribute("", "cap", "", capValue));
                        }
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
                        runs[0].Text = MakePreservingText(textLines[0]);
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
                                    r.Text = MakePreservingText(seg);
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
                        // Empty value clears the override (removes the elements) so
                        // the run inherits the theme/placeholder font, rather than
                        // pinning a literal typeface="" (which is not "default").
                        if (!string.IsNullOrEmpty(value))
                        {
                            rProps.Append(new Drawing.LatinFont { Typeface = value });
                            rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                        }
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;

                case "font.latin":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        if (!string.IsNullOrEmpty(value))
                            rProps.Append(new Drawing.LatinFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;

                case "font.ea" or "font.eastasia" or "font.eastasian":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        if (!string.IsNullOrEmpty(value))
                            rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;

                case "font.cs" or "font.complexscript" or "font.complex":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        if (!string.IsNullOrEmpty(value))
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

                case "highlight":
                {
                    // CONSISTENCY(highlight): same a:highlight write as the
                    // find/replace formatting path (ApplyPptRunFormatting in
                    // Helpers.RunFormat.cs). ReorderDrawingRunProperties pins
                    // the schema slot — CT_TextCharacterProperties requires
                    // highlight after effectLst and before uLn/uFill/latin;
                    // PowerPoint silently drops out-of-order children.
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.Highlight>();
                        if (!string.Equals(value, "none", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
                        {
                            var hl = new Drawing.Highlight();
                            hl.AppendChild(BuildSolidFillColor(value));
                            rProps.AppendChild(hl);
                            ReorderDrawingRunProperties(rProps);
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

                // R61 bt-1: <a:ln> on rPr — text outline / glyph stroke. Distinct
                // from shape-level line= (which strokes the shape edge on spPr).
                // Compound form `textOutline=width:color` mirrors SplitCompoundLineValue
                // (the `line=` parser); split keys `textOutline.width` and
                // `textOutline.color` allow additive Set without overwriting
                // the other half. Schema order bucket 1 (ln) — ReorderDrawingRunProperties
                // moves it before solidFill/latin/etc.
                case "textFillRaw" or "textfillraw":
                {
                    // Verbatim <a:blipFill> glyph fill (WordArt picture fill,
                    // sample10). The referenced ImagePart rides separately via
                    // add-part image with the pinned rId.
                    foreach (var run in runs)
                    {
                        var rp = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rp.RemoveAllChildren<Drawing.SolidFill>();
                        rp.RemoveAllChildren<Drawing.GradientFill>();
                        rp.RemoveAllChildren<Drawing.BlipFill>();
                        rp.RemoveAllChildren<Drawing.PatternFill>();
                        if (!string.IsNullOrWhiteSpace(value))
                            InsertFillInRunProperties(rp, new Drawing.BlipFill(value));
                    }
                    break;
                }

                case "textOutlineRaw" or "textoutlineraw":
                {
                    // Verbatim run <a:ln> — dash / gradient stroke / cap-join
                    // forms the width:color compound can't express (sample10).
                    foreach (var run in runs)
                    {
                        var rp = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rp.RemoveAllChildren<Drawing.Outline>();
                        if (!string.IsNullOrWhiteSpace(value))
                            rp.PrependChild(new Drawing.Outline(value));
                    }
                    break;
                }

                case "textOutline" or "textoutline":
                {
                    // "none" / "false" → strip; mirrors text underline=none clearing.
                    if (value.Equals("none", System.StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", System.StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var run in runs)
                            run.RunProperties?.RemoveAllChildren<Drawing.Outline>();
                        break;
                    }
                    // Compound is width:color (Get emit form mirrors the
                    // canonical width-first dotted keys textOutline.width /
                    // textOutline.color). SplitCompoundLineValue returns
                    // (first, second, _) positions — name-shadows the line=
                    // (color, width, dash) layout because the underlying
                    // split is position-only.
                    var (toWidthPart, toColorPart, _) = SplitCompoundLineValue(value);
                    long? widthEmu = null;
                    // Carry the full color string (incl. +shade/+alpha/+lumMod
                    // transform chain and #RRGGBBAA alpha) through BuildSolidFill
                    // so the dump round-trip form survives. SanitizeColorForOoxml
                    // only returns the bare RGB — it strips the transform suffix,
                    // which made Get's emit form ("#4F81BD11+shade2") un-replayable.
                    string? colorValue = null;
                    if (toColorPart != null)
                    {
                        widthEmu = Core.EmuConverter.ParseLineWidth(toWidthPart);
                        colorValue = toColorPart.Equals("none", System.StringComparison.OrdinalIgnoreCase)
                            ? null : toColorPart;
                    }
                    else
                    {
                        // Single-part: try width first (bare "2pt", "0.5pt",
                        // numeric EMU). Falls through to colour parse if not.
                        try { widthEmu = Core.EmuConverter.ParseLineWidth(value); }
                        catch { widthEmu = null; }
                        if (widthEmu == null && !value.Equals("true", System.StringComparison.OrdinalIgnoreCase))
                            colorValue = value;
                    }
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.Outline>();
                        var ln = new Drawing.Outline();
                        if (widthEmu.HasValue) ln.Width = (int)widthEmu.Value;
                        if (colorValue != null)
                            ln.AppendChild(BuildSolidFill(colorValue));
                        rProps.AppendChild(ln);
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;
                }

                case "textOutline.width" or "textoutline.width":
                {
                    var widthEmu = Core.EmuConverter.ParseLineWidth(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        var ln = rProps.GetFirstChild<Drawing.Outline>();
                        if (ln == null)
                        {
                            ln = new Drawing.Outline();
                            rProps.PrependChild(ln);
                            ReorderDrawingRunProperties(rProps);
                        }
                        ln.Width = (int)widthEmu;
                    }
                    break;
                }

                case "textOutline.color" or "textoutline.color":
                {
                    // BuildSolidFill carries the +shade/+alpha/+lumMod transform
                    // chain and #RRGGBBAA alpha that Get emits; SanitizeColorForOoxml
                    // returned the bare RGB only, dropping the round-trip suffix.
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        var ln = rProps.GetFirstChild<Drawing.Outline>();
                        if (ln == null)
                        {
                            ln = new Drawing.Outline();
                            rProps.PrependChild(ln);
                            ReorderDrawingRunProperties(rProps);
                        }
                        ln.RemoveAllChildren<Drawing.SolidFill>();
                        ln.AppendChild(BuildSolidFill(value));
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
                    // Value in percent (e.g. "30" or "30%" = 30% superscript, "-25"
                    // or "-25%" = 25% subscript). OOXML stores as 1/1000ths of
                    // percent (30000 = 30%). Shortcuts: "super"/"true" = 30%,
                    // "sub" = -25%, "none"/"false" = 0. R56 bt-3: accept the
                    // canonical `%` suffix the Get reader now emits.
                    int baselineVal;
                    if (key.ToLowerInvariant() == "superscript")
                        baselineVal = IsTruthy(value) ? 30000 : 0;
                    else if (key.ToLowerInvariant() == "subscript")
                        baselineVal = IsTruthy(value) ? -25000 : 0;
                    else
                    {
                        var blNorm = value.Trim().TrimEnd('%').Trim();
                        baselineVal = blNorm.ToLowerInvariant() switch
                        {
                            "super" or "true" => 30000,
                            "sub" => -25000,
                            "none" or "false" or "0" => 0,
                            _ => double.TryParse(blNorm, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var blVal) && !double.IsNaN(blVal) && !double.IsInfinity(blVal)
                                ? (int)(blVal * 1000)
                                : throw new ArgumentException($"Invalid 'baseline' value: '{value}'. Expected 'super', 'sub', 'none', or a percentage (e.g. 30 or 30% for superscript 30%).")
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

                case "gradientraw":
                {
                    // bt-B2 / bt-7 dump→replay passthrough. Value is the
                    // captured <a:gradFill ...> verbatim including flip= and
                    // any <a:tileRect/> child — attributes BuildGradientFill
                    // never re-emits. Delegates to ApplyGradientRaw so AddShape
                    // and SetShape share one parser; previously the Set side
                    // inlined the XmlReader walk while AddShape silently
                    // dropped the key — the helper closed that gap.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    if (!ApplyGradientRaw(spPr, value))
                        unsupported.Add(key);
                    break;
                }

                case "pattern":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyPatternFill(spPr, value);
                    break;
                }

                case "liststyle" or "list" or "bullet":
                {
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        ApplyListStyle(pProps, value, preserveIndent: properties.ContainsKey("indent") || properties.ContainsKey("marginLeft") || properties.ContainsKey("marginleft") || properties.ContainsKey("marL") || properties.ContainsKey("marl"));
                    }
                    break;
                }

                case "bulletraw" or "bulletRaw":
                {
                    // Full bullet group (buClr/buFont/buSzPct/buChar/…) verbatim.
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        ApplyBulletRaw(pProps, value);
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

                case "align" or "alignment" or "halign":
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
                    // flow of the text body itself. PowerPoint's UI sets
                    // both when the user toggles "Right-to-left text direction"
                    // on a shape, so a single 'direction=rtl' here mirrors the
                    // same intent end-to-end.
                    bool rtl = key.ToLowerInvariant() == "rtl"
                        ? IsTruthy(value)
                        : ParsePptDirectionRtl(value);
                    // CONSISTENCY(run-context-explicit): when the caller targeted
                    // a run path, write <a:rPr rtl="1"/> on the run only and
                    // leave pPr/bodyPr alone. Mirrors the shadow/glow/reflection
                    // run-context branches and matches the OOXML schema, which
                    // allows the rtl attribute on CT_TextCharacterProperties too.
                    if (runContext && runs.Count > 0)
                    {
                        foreach (var run in runs)
                        {
                            var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            // Drawing.RunProperties does not expose `rtl` as a
                            // typed property even though CT_TextCharacterProperties
                            // declares it; set the raw attribute (unqualified ns
                            // matches DrawingML's <a:rPr rtl="1"/>).
                            if (rtl)
                                rProps.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "rtl", "", "1"));
                            else
                                rProps.RemoveAttribute("rtl", "");
                        }
                        break;
                    }
                    // R64 bt-2 / bt-4: explicit Set writes the attribute on
                    // BOTH rtl and ltr instead of stripping for ltr. An explicit
                    // shape-level Set is the caller's "override inheritance"
                    // signal — a textbox inside a master/layout with rtl=1 /
                    // rtlCol=1 silently inherits RTL when we strip the attribute
                    // on ltr, so the Set looks Updated but the persisted XML
                    // reads as a no-op (`<a:pPr/>` / `<a:bodyPr/>` with no rtl /
                    // rtlCol attr). Writing "0" pins ltr regardless of inherited
                    // cascade. (Add path keeps strip-on-ltr so a freshly built
                    // ltr shape stays free of explicit-default noise.)
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RightToLeft = rtl;
                    }
                    var dirBodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    // OpenXml SDK doesn't expose rtlCol as a typed property on
                    // BodyProperties — set the attribute directly. "1"/"0" is
                    // the only canonical xsd:boolean form Office tooling reads.
                    if (dirBodyPr != null)
                    {
                        dirBodyPr.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "rtlCol", "", rtl ? "1" : "0"));
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

                case "columns" or "numcol":
                {
                    // <a:bodyPr numCol="N"/> lays the text body out in N columns.
                    // Mirrors the valign/textdirection bodyPr-attr setters above.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var nCol) || nCol < 1 || nCol > 16)
                        throw new ArgumentException($"Invalid columns: '{value}'. Use an integer 1-16.");
                    bodyPr.ColumnCount = nCol;
                    break;
                }

                case "columnspacing" or "spccol":
                {
                    // <a:bodyPr spcCol="EMU"/> — gap between columns. Bare numbers
                    // are points (CONSISTENCY(pptx-bare-as-points)).
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.ColumnSpacing = (int)Math.Round(SpacingConverter.ParsePointsSigned(value) * EmuConverter.EmuPerPointF);
                    break;
                }

                case "textdirection" or "textdir":
                {
                    // CONSISTENCY(textdir-shape): <a:bodyPr vert="…"/> is valid
                    // on shapes/textboxes, not just table cells. Mirrors the cell
                    // helper's textdirection case (Set.Table cell context) so the
                    // same vocabulary works at the shape surface.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    // OOXML semantics:
                    //   Vertical    = 90° CCW (bottom-to-top — what most users call "vert")
                    //   Vertical270 = 270° CCW (top-to-bottom — the rarer rotation)
                    // The old switch collapsed "vert" and "vert270" both to Vertical270, so
                    // textDirection=vert silently produced the wrong rotation. Split into
                    // distinct cases and add "eavert" (East Asian vertical) which OOXML
                    // supports as its own enum member.
                    bodyPr.Vertical = value.ToLowerInvariant() switch
                    {
                        "horizontal" or "horz" or "none" => Drawing.TextVerticalValues.Horizontal,
                        "vertical" or "vert" or "vertical90" or "vert90" => Drawing.TextVerticalValues.Vertical,
                        "vertical270" or "vert270" => Drawing.TextVerticalValues.Vertical270,
                        // Note: SDK enum member spelling is "EastAsianVetical" (typo
                        // present in DocumentFormat.OpenXml 3.x); serialized XML is "eaVert".
                        "eavert" or "eavertical" => Drawing.TextVerticalValues.EastAsianVetical,
                        "stacked" or "wordartvert" => Drawing.TextVerticalValues.WordArtVertical,
                        _ => throw new ArgumentException($"Invalid textDirection: '{value}'. Valid: horizontal, vertical (=vert / vertical90, 90° CCW), vertical270 (=vert270, 270° CCW), eaVert, stacked.")
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

                case "adj":
                {
                    // CONSISTENCY(preset-adj-handles): set the avLst on the
                    // shape's PresetGeometry. The shape must already carry
                    // a preset (Set adj on a custGeom is meaningless — its
                    // own avLst is part of the custom path); if there's
                    // none yet, materialize an empty preset rectangle so
                    // the caller's adj values land somewhere predictable
                    // rather than throwing.
                    var spPrAdj = shape.ShapeProperties ?? (shape.ShapeProperties = new ShapeProperties());
                    var prstAdj = spPrAdj.GetFirstChild<Drawing.PresetGeometry>();
                    if (prstAdj == null)
                    {
                        prstAdj = EnsurePresetGeometry(spPrAdj);
                        prstAdj.Preset = Drawing.ShapeTypeValues.Rectangle;
                    }
                    var avLst = prstAdj.GetFirstChild<Drawing.AdjustValueList>()
                        ?? prstAdj.AppendChild(new Drawing.AdjustValueList())!;
                    ApplyAdjustHandles(avLst, value, prstAdj.Preset?.Value);
                    break;
                }

                case "geometry" or "path" when key.ToLowerInvariant() != "path" || shape.ShapeProperties != null:
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    // Distinguish preset shape name from SVG-like custom path.
                    // SVG paths always have whitespace-separated commands and
                    // comma-separated coordinates ("M 0,0 L 100,0 Z"); preset
                    // names are bare camelCase identifiers. The previous
                    // `!value.Contains('M')` heuristic misfired on legitimate
                    // preset names containing 'M' — flowChartMultidocument,
                    // flowChartMerge, flowChartManualInput — routing them
                    // through ParseCustomGeometry which produced an empty
                    // <a:pathLst><a:path/></a:pathLst> and a blank render.
                    if (!value.Contains(' ') && !value.Contains(','))
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
                    // styledLine: the emitter signals that a <p:style> raw-set
                    // will follow and no explicit line colour was dumped — the
                    // stroke colour comes from lnRef, so injecting the default
                    // black here would override the theme tint (stress013's
                    // grey/orange borders replayed black). Mirrors the
                    // connector styledLine contract.
                    if (!properties.ContainsKey("styledLine") && !properties.ContainsKey("styledline"))
                        EnsureOutlineHasFill(outline);
                    break;
                }

                case "styledline":
                    // Signal-only key (see linewidth above) — consumed so the
                    // handler-as-truth tracker doesn't report it unsupported.
                    break;

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
                    var grad = BuildGradientFill(NormalizeLineGradientSpec(value));
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
                    outline.RemoveAllChildren<Drawing.CustomDash>();
                    outline.AppendChild(new Drawing.PresetDash { Val = ParseLineDashValue(value) });
                    break;
                }

                // R64 bt-3: lineDashRaw — verbatim <a:custDash> passthrough on
                // shape Set. Mirrors connector Set: clears any preset/custom
                // dash and appends a fresh Drawing.CustomDash rebuilt from the
                // source XML. Empty value removes the dash entirely.
                case "linedashraw" or "line.dashraw":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    outline.RemoveAllChildren<Drawing.CustomDash>();
                    if (!string.IsNullOrWhiteSpace(value))
                        outline.AppendChild(BuildCustomDashFromRaw(value));
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
                // R61 bt-2: accept compound form "miter:<lim>" so a single CLI key can
                // carry both the join token and the miter limit; standalone miterLimit
                // case below also extends a pre-existing <a:miter> with the lim attr.
                case "linejoin" or "line.join":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var outline = EnsureOutline(spPr);
                    // BUGFIX (CompanionInterferenceScanTests): preserve a
                    // previously-set miter limit when re-affirming lineJoin=miter
                    // without an inline limit token. Same rebuild-drops-sibling
                    // family as the autoFit/legend/labelPos fixes — RemoveAllChildren
                    // + fresh <a:miter> used to wipe the lim attribute set via the
                    // standalone miterLimit= property.
                    var prevMiterLimit = outline.GetFirstChild<Drawing.Miter>()?.Limit;
                    outline.RemoveAllChildren<Drawing.Round>();
                    outline.RemoveAllChildren<Drawing.LineJoinBevel>();
                    outline.RemoveAllChildren<Drawing.Miter>();
                    var joinValue = value;
                    int? compoundMiterLimit = null;
                    var colonIdx = value.IndexOf(':');
                    if (colonIdx > 0)
                    {
                        joinValue = value.Substring(0, colonIdx);
                        var limTok = value.Substring(colonIdx + 1).Trim();
                        if (!int.TryParse(limTok, System.Globalization.NumberStyles.Integer,
                                System.Globalization.CultureInfo.InvariantCulture, out var limParsed))
                            throw new ArgumentException($"Invalid 'lineJoin' miter limit token: '{limTok}'. Expected integer (1000ths of a percent, e.g. 800000 = 800%).");
                        compoundMiterLimit = limParsed;
                    }
                    OpenXmlElement joinEl = joinValue.ToLowerInvariant() switch
                    {
                        "round" => new Drawing.Round(),
                        "bevel" => new Drawing.LineJoinBevel(),
                        "miter" => compoundMiterLimit.HasValue
                            ? new Drawing.Miter { Limit = compoundMiterLimit.Value }
                            : (prevMiterLimit != null ? new Drawing.Miter { Limit = prevMiterLimit } : new Drawing.Miter()),
                        _ => throw new ArgumentException($"Invalid 'lineJoin' value: '{joinValue}'. Valid values: round, bevel, miter.")
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
                // R61 bt-2: miterLimit → <a:miter lim="N"/> attribute. Extends an
                // existing <a:miter/> with the lim attribute, or auto-creates the
                // miter join if none was set (matches PowerPoint behavior — lim
                // is meaningless without miter as the join). Value is OOXML
                // 1000ths-of-a-percent (e.g. 800000 = 800%).
                case "miterlimit" or "miter.limit" or "line.miterlimit":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    if (!int.TryParse(value, System.Globalization.NumberStyles.Integer,
                            System.Globalization.CultureInfo.InvariantCulture, out var limVal))
                        throw new ArgumentException($"Invalid 'miterLimit' value: '{value}'. Expected integer (1000ths of a percent, e.g. 800000 = 800%).");
                    // BUGFIX (NumericBoundaryScanTests): <a:miter lim> is a
                    // non-negative percentage; a negative value is schema-invalid.
                    if (limVal < 0)
                        throw new ArgumentException($"Invalid 'miterLimit' value: '{value}'. Must be >= 0 (1000ths of a percent, e.g. 800000 = 800%).");
                    var outline = EnsureOutline(spPr);
                    var miterEl = outline.GetFirstChild<Drawing.Miter>();
                    if (miterEl == null)
                    {
                        outline.RemoveAllChildren<Drawing.Round>();
                        outline.RemoveAllChildren<Drawing.LineJoinBevel>();
                        miterEl = new Drawing.Miter { Limit = limVal };
                        var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                        if (headEnd != null) outline.InsertBefore(miterEl, headEnd);
                        else
                        {
                            var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                            if (tailEnd != null) outline.InsertBefore(miterEl, tailEnd);
                            else outline.AppendChild(miterEl);
                        }
                    }
                    else
                    {
                        miterEl.Limit = limVal;
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
                    // BUGFIX (NumericBoundaryScanTests): enforce the stated 0.0-1.0
                    // range. Out-of-range values produced an <a:alpha> outside
                    // [0,100000] → schema-invalid file PowerPoint refuses to open.
                    if (lnOpacity < 0.0 || lnOpacity > 1.0)
                        throw new ArgumentException($"Invalid 'lineopacity' value: '{value}'. Expected a decimal in 0.0-1.0 (e.g. 0.5 = 50% opacity).");
                    var outline = EnsureOutline(spPr);
                    var solidFillLn = outline.GetFirstChild<Drawing.SolidFill>();
                    if (solidFillLn == null)
                    {
                        // Auto-create a black line fill
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
                                    // 100000 = 100% = OOXML default; omit the element.
                                    if (alphaPct < 100000)
                                        stopColorEl.AppendChild(new Drawing.Alpha { Val = alphaPct });
                                }
                            }
                        }
                        break;
                    }

                    var solidFill = spPr.GetFirstChild<Drawing.SolidFill>();
                    if (solidFill == null)
                    {
                        // Auto-create a white fill
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
                            // 100000 = 100% = OOXML default; omit the element.
                            if (alphaPct < 100000)
                                colorEl.AppendChild(new Drawing.Alpha { Val = alphaPct });
                        }
                    }
                    break;
                }

                case "image" or "imagefill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null || part is not SlidePart slidePart) { unsupported.Add(key); break; }
                    // image=none/clear removes the blip fill (mirrors fill=none → NoFill).
                    // Guard before any file resolve so ImageSource.Resolve never throws on "none".
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("clear", StringComparison.OrdinalIgnoreCase))
                    {
                        spPr.RemoveAllChildren<Drawing.BlipFill>();
                        break;
                    }
                    // Pass any sibling fillRect= / srcRect= so the image fill's
                    // framing (stretch insets / crop) round-trips with it.
                    string? frSpec = properties.TryGetValue("fillRect", out var frv) ? frv
                        : properties.TryGetValue("fillrect", out frv) ? frv : null;
                    string? srSpec = properties.TryGetValue("srcRect", out var srv) ? srv
                        : properties.TryGetValue("srcrect", out srv) ? srv : null;
                    ApplyShapeImageFill(spPr, value, slidePart, frSpec, srSpec);
                    break;
                }

                case "fillRect" or "fillrect" or "srcRect" or "srcrect":
                {
                    // Blip-fill framing. Normally consumed as a sibling of image=
                    // (above). Handle the standalone case too — a Set that adjusts
                    // the stretch insets / crop on a shape that already carries a
                    // blip fill — by patching the existing <a:blipFill>. No image
                    // fill present → nothing to frame, leave as a no-op rather than
                    // a spurious unsupported_property.
                    var spPr = shape.ShapeProperties;
                    var existingBlip = spPr?.GetFirstChild<Drawing.BlipFill>();
                    if (existingBlip == null) break;
                    var rect = ParsePerMilleRect(value);
                    if (!rect.HasValue) break;
                    bool isSrc = key.Equals("srcRect", StringComparison.OrdinalIgnoreCase)
                                 || key.Equals("srcrect", StringComparison.OrdinalIgnoreCase);
                    if (isSrc)
                    {
                        existingBlip.RemoveAllChildren<Drawing.SourceRectangle>();
                        var blipEl = existingBlip.GetFirstChild<Drawing.Blip>();
                        var sr = new Drawing.SourceRectangle { Left = rect.Value.L, Top = rect.Value.T, Right = rect.Value.R, Bottom = rect.Value.B };
                        if (blipEl != null) existingBlip.InsertAfter(sr, blipEl); else existingBlip.PrependChild(sr);
                    }
                    else
                    {
                        var stretch = existingBlip.GetFirstChild<Drawing.Stretch>();
                        if (stretch == null) { stretch = new Drawing.Stretch(); existingBlip.AppendChild(stretch); }
                        stretch.RemoveAllChildren<Drawing.FillRectangle>();
                        stretch.AppendChild(new Drawing.FillRectangle { Left = rect.Value.L, Top = rect.Value.T, Right = rect.Value.R, Bottom = rect.Value.B });
                    }
                    break;
                }

                case "spacing" or "charspacing" or "letterspacing" or "spc":
                {
                    // Character spacing in points (e.g. "2" = +2pt, "-1" = -1pt)
                    // Stored as 1/100th of a point in OOXML
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
                    // CONSISTENCY(pptx-bare-as-points): mirror AddParagraph / Set.Shape.
                    var indentEmu = (int)Math.Round(SpacingConverter.ParsePointsSigned(value) * EmuConverter.EmuPerPointF);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Indent = indentEmu;
                    }
                    break;
                }

                case "marginleft" or "marl":
                {
                    // CONSISTENCY(pptx-bare-as-points): mirror AddParagraph / Set.Shape.
                    var mlEmu = (int)Math.Round(SpacingConverter.ParsePointsSigned(value) * EmuConverter.EmuPerPointF);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.LeftMargin = mlEmu;
                    }
                    break;
                }

                case "marginright" or "marr":
                {
                    var mrEmu = (int)Math.Round(SpacingConverter.ParsePointsSigned(value) * EmuConverter.EmuPerPointF);
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

                case "textwarpraw":
                {
                    // Verbatim <a:prstTxWarp> (round-trips the avLst adjust
                    // values the semantic textwarp= preset-name form loses).
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.RemoveAllChildren<Drawing.PresetTextWarp>();
                    if (!string.IsNullOrWhiteSpace(value))
                        bodyPr.InsertAt(new Drawing.PresetTextWarp(value), 0); // schema: first bodyPr child
                    break;
                }

                case "textscene3draw":
                {
                    // Verbatim <a:scene3d> INSIDE bodyPr — 3D text camera/light
                    // (distinct from the shape-level scene3d on spPr).
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.RemoveAllChildren<Drawing.Scene3DType>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        // Schema order: …autofit, scene3d, sp3d — insert before
                        // an existing sp3d, else append.
                        var sp3dSibling = bodyPr.GetFirstChild<Drawing.Shape3DType>();
                        var scene = new Drawing.Scene3DType(value);
                        if (sp3dSibling != null) bodyPr.InsertBefore(scene, sp3dSibling);
                        else bodyPr.AppendChild(scene);
                    }
                    break;
                }

                case "textsp3draw":
                {
                    // Verbatim <a:sp3d> INSIDE bodyPr — 3D text extrusion/bevel.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.RemoveAllChildren<Drawing.Shape3DType>();
                    if (!string.IsNullOrWhiteSpace(value))
                        bodyPr.AppendChild(new Drawing.Shape3DType(value));
                    break;
                }

                case "vertoverflow":
                {
                    // <a:bodyPr vertOverflow="clip|ellipsis|overflow">.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.VerticalOverflow = value.ToLowerInvariant() switch
                    {
                        "clip" => Drawing.TextVerticalOverflowValues.Clip,
                        "ellipsis" => Drawing.TextVerticalOverflowValues.Ellipsis,
                        "overflow" => Drawing.TextVerticalOverflowValues.Overflow,
                        _ => throw new ArgumentException($"Invalid 'vertOverflow' value: '{value}'. Valid values: clip, ellipsis, overflow.")
                    };
                    break;
                }

                case "horzoverflow":
                {
                    // <a:bodyPr horzOverflow="clip|overflow">.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.HorizontalOverflow = value.ToLowerInvariant() switch
                    {
                        "clip" => Drawing.TextHorizontalOverflowValues.Clip,
                        "overflow" => Drawing.TextHorizontalOverflowValues.Overflow,
                        _ => throw new ArgumentException($"Invalid 'horzOverflow' value: '{value}'. Valid values: clip, overflow.")
                    };
                    break;
                }

                case "anchorctr" or "anchorcenter":
                {
                    // <a:bodyPr anchorCtr="0|1"> — horizontal centering of the
                    // whole text block (distinct from paragraph align).
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.AnchorCenter = IsTruthy(value);
                    break;
                }

                case "upright":
                {
                    // <a:bodyPr upright="1"> — glyphs stay upright inside a
                    // rotated text body.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.UpRight = IsTruthy(value);
                    break;
                }

                case "wrap" or "wordwrap":
                {
                    // Shape-level <a:bodyPr @wrap = "square" | "none">.
                    // NodeBuilder surfaces this as Format["wrap"] = true|false,
                    // mirror the table-cell wrap case higher up. Without a
                    // shape-level handler, dump->replay silently lost
                    // `wrap="square"` on every textbox emit.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.Wrap = IsTruthy(value)
                        ? Drawing.TextWrappingValues.Square
                        : Drawing.TextWrappingValues.None;
                    break;
                }

                case "fontscale" or "fontScale" or "lnspcreduction" or "lnSpcReduction"
                    or "linespacereduction" or "lineSpaceReduction"
                    or "linespacingreduction" or "lineSpacingReduction":
                {
                    // Consumed as siblings of autofit= (ApplyNormalAutoFitScale).
                    // Handle standalone too: patch the existing <a:normAutofit>, or
                    // synthesize one when the body has no autofit child yet (fontScale/
                    // lnSpcReduction are normAutofit attributes — setting them implies
                    // shrink-to-fit mode). normAutofit is mutually exclusive with
                    // spAutoFit/noAutofit, so a bare body needs one created.
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    var naf = bodyPr.GetFirstChild<Drawing.NormalAutoFit>();
                    if (naf == null && bodyPr.GetFirstChild<Drawing.ShapeAutoFit>() == null
                        && bodyPr.GetFirstChild<Drawing.NoAutoFit>() == null)
                        naf = bodyPr.AppendChild(new Drawing.NormalAutoFit());
                    if (naf != null) ApplyNormalAutoFitScale(naf, properties);
                    else unsupported.Add(key);
                    break;
                }
                case "autofit":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    // BUGFIX (CompanionInterferenceScanTests): capture the existing
                    // normAutofit's scale attributes before removing it, so
                    // re-affirming the autofit mode (autoFit=normal) doesn't wipe a
                    // previously-set fontScale / lnSpcReduction. Same rebuild-drops-
                    // sibling family as the chart legend/labelPos fixes.
                    var prevNaf = bodyPr.GetFirstChild<Drawing.NormalAutoFit>();
                    var prevFontScale = prevNaf?.FontScale;
                    var prevLnSpcReduction = prevNaf?.LineSpaceReduction;
                    bodyPr.RemoveAllChildren<Drawing.NormalAutoFit>();
                    bodyPr.RemoveAllChildren<Drawing.ShapeAutoFit>();
                    bodyPr.RemoveAllChildren<Drawing.NoAutoFit>();
                    switch (value.ToLowerInvariant())
                    {
                        // R10-4: 'shrink' and 'true' alias normAutofit. PowerPoint's
                        // <a:normAutofit> IS the shrink-text-on-overflow mode; the
                        // optional fontScale/lnSpcReduction attributes carry the
                        // computed shrink ratio (callers may tune via fontScale=).
                        case "true" or "shrink" or "normal" or "normautofit" or "auto":
                        {
                            var naf = ApplyNormalAutoFitScale(new Drawing.NormalAutoFit(), properties);
                            // Carry over the prior scale when this Set didn't supply one.
                            if (naf.FontScale == null && prevFontScale != null) naf.FontScale = prevFontScale;
                            if (naf.LineSpaceReduction == null && prevLnSpcReduction != null) naf.LineSpaceReduction = prevLnSpcReduction;
                            bodyPr.AppendChild(naf);
                            break;
                        }
                        case "shape" or "spautofit" or "resize": bodyPr.AppendChild(new Drawing.ShapeAutoFit()); break;
                        case "false" or "none": bodyPr.AppendChild(new Drawing.NoAutoFit()); break;
                        default: throw new ArgumentException($"Invalid autofit value: '{value}'. Valid values: true/shrink/normal, shape/resize, false/none.");
                    }
                    break;
                }

                case "x" or "y" or "left" or "top" or "width" or "height":
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
                    // CONSISTENCY(run-context-explicit): when the caller explicitly
                    // targeted a run path, write to the run's <a:rPr><a:effectLst>
                    // unconditionally. The IsNoFillShape heuristic only makes sense
                    // for whole-shape Set; an explicit run-path Set must not be
                    // hijacked to the shape level when the shape has a fill.
                    if (runContext && runs.Count > 0)
                        foreach (var run in runs) ApplyTextShadow(run, shadowVal);
                    else if (IsNoFillShape(spPr) && runs.Count > 0)
                        foreach (var run in runs) ApplyTextShadow(run, shadowVal);
                    else
                        ApplyShadow(spPr, shadowVal);
                    break;
                }

                case "shadowraw":
                {
                    // bt-2 dump→replay path. Value is the verbatim
                    // <a:outerShdw .../> element captured by NodeBuilder
                    // when sx/sy/kx/ky/algn/rotWithShape deviate from
                    // ApplyShadow's compressed-form defaults. Re-install
                    // verbatim so source-authored scale/skew survives the
                    // round-trip instead of collapsing to the COLOR-BLUR-
                    // ANGLE-DIST-OPACITY tuple. Mirrors reflectionRaw shape.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var effectList = EnsureEffectList(spPr);
                    effectList.RemoveAllChildren<Drawing.OuterShadow>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        try
                        {
                            var raw = value.Contains("xmlns:a=")
                                ? value
                                : value.Replace("<a:outerShdw",
                                    "<a:outerShdw xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
                            var shadow = new Drawing.OuterShadow();
                            shadow.InnerXml = "";
                            // OpenXml SDK supports OuterXml-style construction via
                            // OuterXml setter only on some types; for OuterShadow
                            // we set InnerXml + lift attributes manually.
                            using var sr = new System.IO.StringReader(raw);
                            using var xr = System.Xml.XmlReader.Create(sr);
                            xr.MoveToContent();
                            if (xr.HasAttributes)
                            {
                                while (xr.MoveToNextAttribute())
                                {
                                    if (xr.Prefix == "xmlns" || xr.Name == "xmlns") continue;
                                    shadow.SetAttribute(new OpenXmlAttribute(
                                        xr.Prefix, xr.LocalName, xr.NamespaceURI, xr.Value));
                                }
                                xr.MoveToElement();
                            }
                            // Lift inner XML (color child + its alpha) so the
                            // shadow color round-trips through the raw path.
                            if (!xr.IsEmptyElement)
                            {
                                var subtreeXml = xr.ReadInnerXml();
                                if (!string.IsNullOrWhiteSpace(subtreeXml))
                                {
                                    var inner = subtreeXml.Contains("xmlns:a=")
                                        ? subtreeXml
                                        : "<wrap xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                                          + subtreeXml + "</wrap>";
                                    // Cheap approach: set InnerXml on the shadow
                                    // directly — OpenXml SDK accepts a-namespaced
                                    // children when the namespace is declared on
                                    // the parent already (shape's a:graphic root).
                                    shadow.InnerXml = subtreeXml;
                                }
                            }
                            DrawingEffectsHelper.InsertEffectInSchemaOrder(effectList, shadow);
                        }
                        catch
                        {
                            unsupported.Add(key);
                        }
                    }
                    break;
                }

                case "effectdagraw":
                {
                    // R52 bt-2: <a:effectDag> verbatim passthrough. effectDag
                    // is a sibling to effectLst on spPr (CT_EffectProperties
                    // choice — schema accepts at most one of effectLst /
                    // effectDag, though real decks compose both in document
                    // order). The cont/sib nesting + per-leaf blur radius /
                    // fillOverlay blend modes have no compressible form, so
                    // mirror shadowRaw/fillOverlayRaw: lift attrs + InnerXml,
                    // install as a fresh effectDag on spPr.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    spPr.RemoveAllChildren<Drawing.EffectDag>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        try
                        {
                            var raw = value.Contains("xmlns:a=")
                                ? value
                                : value.Replace("<a:effectDag",
                                    "<a:effectDag xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
                            var dag = new Drawing.EffectDag();
                            using var sr = new System.IO.StringReader(raw);
                            using var xr = System.Xml.XmlReader.Create(sr);
                            xr.MoveToContent();
                            if (xr.HasAttributes)
                            {
                                while (xr.MoveToNextAttribute())
                                {
                                    if (xr.Prefix == "xmlns" || xr.Name == "xmlns") continue;
                                    dag.SetAttribute(new OpenXmlAttribute(
                                        xr.Prefix, xr.LocalName, xr.NamespaceURI, xr.Value));
                                }
                                xr.MoveToElement();
                            }
                            if (!xr.IsEmptyElement)
                            {
                                var inner = xr.ReadInnerXml();
                                if (!string.IsNullOrWhiteSpace(inner))
                                    dag.InnerXml = inner;
                            }
                            // Schema order on spPr: ...fill, ln, effectLst /
                            // effectDag, scene3d, sp3d, extLst. Insert after
                            // effectLst (or after any fill/ln) so PowerPoint
                            // accepts the result.
                            var afterEl = (OpenXmlElement?)spPr.GetFirstChild<Drawing.EffectList>()
                                ?? (OpenXmlElement?)spPr.GetFirstChild<Drawing.Outline>()
                                ?? spPr.GetFirstChild<Drawing.SolidFill>() as OpenXmlElement
                                ?? spPr.GetFirstChild<Drawing.GradientFill>() as OpenXmlElement;
                            if (afterEl != null) spPr.InsertAfter(dag, afterEl);
                            else spPr.AppendChild(dag);
                        }
                        catch
                        {
                            unsupported.Add(key);
                        }
                    }
                    break;
                }

                case "effectsraw":
                {
                    // R58 bt-2: <a:effectLst> verbatim passthrough. The
                    // compressed shadow/innerShadow/glow/fillOverlay/reflection/
                    // softEdge/blur readers handle the well-known children,
                    // but real decks carry tint / lum / hsl / alphaModFix /
                    // clrChange / duotone / biLevel / xfrm / relOff children
                    // on the spPr effectLst (the reported case was
                    // <a:effectLst><a:tint amt="50000"/></a:effectLst>). These
                    // have no compressible string surface, so the only round-
                    // trip-safe form is the verbatim OuterXml. Mirrors
                    // effectDagRaw / fillOverlayRaw — lift attrs + InnerXml,
                    // install a fresh <a:effectLst> in schema-correct position
                    // (which EnsureEffectList already handles). Wins over any
                    // compressed sibling keys emitted from the same source
                    // effectLst: replaces the whole element wholesale.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    spPr.RemoveAllChildren<Drawing.EffectList>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        try
                        {
                            var raw = value.Contains("xmlns:a=")
                                ? value
                                : value.Replace("<a:effectLst",
                                    "<a:effectLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
                            var freshList = new Drawing.EffectList();
                            using var sr = new System.IO.StringReader(raw);
                            using var xr = System.Xml.XmlReader.Create(sr);
                            xr.MoveToContent();
                            if (xr.HasAttributes)
                            {
                                while (xr.MoveToNextAttribute())
                                {
                                    if (xr.Prefix == "xmlns" || xr.Name == "xmlns") continue;
                                    freshList.SetAttribute(new OpenXmlAttribute(
                                        xr.Prefix, xr.LocalName, xr.NamespaceURI, xr.Value));
                                }
                                xr.MoveToElement();
                            }
                            if (!xr.IsEmptyElement)
                            {
                                var inner = xr.ReadInnerXml();
                                if (!string.IsNullOrWhiteSpace(inner))
                                    freshList.InnerXml = inner;
                            }
                            // EnsureEffectList knows the schema-correct
                            // insertion slot (between ln and scene3d/sp3d/
                            // extLst). Re-use that helper rather than open-
                            // coding the InsertBefore chain.
                            var inserted = EnsureEffectList(spPr);
                            // EnsureEffectList created an empty placeholder;
                            // replace it with the fresh, populated list.
                            inserted.InsertAfterSelf(freshList);
                            inserted.Remove();
                        }
                        catch
                        {
                            unsupported.Add(key);
                        }
                    }
                    break;
                }

                case "filloverlayraw":
                {
                    // bt-1 dump→replay path. Value is the verbatim
                    // <a:fillOverlay blend=…>…</a:fillOverlay>. ApplyShadow /
                    // BuildGlow have no equivalent for fillOverlay; without
                    // raw passthrough the composited tint is dropped from
                    // the shape's effectLst on every round-trip.
                    //
                    // R62 bt-5: run-level <a:rPr><a:fillOverlay> needs the same
                    // passthrough — NodeBuilder now emits fillOverlayRaw on run
                    // nodes too. Honor runContext so a /paragraph[N]/run[K]
                    // path writes to the run's own rPr/effectLst instead of
                    // the shape's spPr/effectLst (which would over-broad apply
                    // the overlay to the shape body). Mirror the shadow/glow/
                    // reflection routing at lines 1410, 1667, 1747.
                    if (runContext && runs.Count > 0)
                    {
                        foreach (var run in runs)
                            ApplyRunFillOverlayRaw(run, value);
                        break;
                    }
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var effectList = EnsureEffectList(spPr);
                    effectList.RemoveAllChildren<Drawing.FillOverlay>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        try
                        {
                            var overlay = BuildFillOverlayFromRaw(value);
                            DrawingEffectsHelper.InsertEffectInSchemaOrder(effectList, overlay);
                        }
                        catch
                        {
                            unsupported.Add(key);
                        }
                    }
                    break;
                }

                case "glow":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var glowVal = value;
                    if (IsValidBooleanString(glowVal) && IsTruthy(glowVal)) glowVal = "4472C4";
                    // CONSISTENCY(run-context-explicit): see shadow case above.
                    if (runContext && runs.Count > 0)
                        foreach (var run in runs) ApplyTextGlow(run, glowVal);
                    else if (IsNoFillShape(spPr) && runs.Count > 0)
                        foreach (var run in runs) ApplyTextGlow(run, glowVal);
                    else
                        ApplyGlow(spPr, glowVal);
                    break;
                }

                case "innershadow":
                {
                    // bt-1: NodeBuilder ignored <a:innerShdw> on dump, so any
                    // shape effectLst carrying an inner shadow lost it on
                    // dump→replay. Mirror the outer-shadow case — same
                    // input vocabulary, separate effectLst child.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var innerVal = value;
                    if (IsValidBooleanString(innerVal) && IsTruthy(innerVal)) innerVal = "000000";
                    ApplyInnerShadow(spPr, innerVal);
                    break;
                }

                case "innershadowraw":
                {
                    // R56 bt-2 dump→replay path. Value is the verbatim
                    // <a:innerShdw .../> element captured by NodeBuilder
                    // when the color child carries lumMod/lumOff/shade/tint
                    // transforms that the compressed innerShadow= form can
                    // only express via the undocumented `accent1+lumMod50+
                    // lumOff50-BLUR-ANGLE-DIST-OPACITY` mixed syntax.
                    // Mirrors shadowRaw — lift attrs + InnerXml, install
                    // a fresh <a:innerShdw> on the effectLst.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var effectList = EnsureEffectList(spPr);
                    effectList.RemoveAllChildren<Drawing.InnerShadow>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        try
                        {
                            var raw = value.Contains("xmlns:a=")
                                ? value
                                : value.Replace("<a:innerShdw",
                                    "<a:innerShdw xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
                            var shadow = new Drawing.InnerShadow();
                            using var sr = new System.IO.StringReader(raw);
                            using var xr = System.Xml.XmlReader.Create(sr);
                            xr.MoveToContent();
                            if (xr.HasAttributes)
                            {
                                while (xr.MoveToNextAttribute())
                                {
                                    if (xr.Prefix == "xmlns" || xr.Name == "xmlns") continue;
                                    shadow.SetAttribute(new OpenXmlAttribute(
                                        xr.Prefix, xr.LocalName, xr.NamespaceURI, xr.Value));
                                }
                                xr.MoveToElement();
                            }
                            if (!xr.IsEmptyElement)
                            {
                                var subtreeXml = xr.ReadInnerXml();
                                if (!string.IsNullOrWhiteSpace(subtreeXml))
                                    shadow.InnerXml = subtreeXml;
                            }
                            DrawingEffectsHelper.InsertEffectInSchemaOrder(effectList, shadow);
                        }
                        catch
                        {
                            unsupported.Add(key);
                        }
                    }
                    break;
                }

                case "reflection":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    // CONSISTENCY(run-context-explicit): see shadow case above.
                    if (runContext && runs.Count > 0)
                        foreach (var run in runs) ApplyTextReflection(run, value);
                    else if (IsNoFillShape(spPr) && runs.Count > 0)
                        foreach (var run in runs) ApplyTextReflection(run, value);
                    else
                        ApplyReflection(spPr, value);
                    break;
                }

                case "reflectionraw":
                {
                    // bt-B1 dump→replay path. Value is the verbatim
                    // <a:reflection .../> element captured by NodeBuilder
                    // when the source-authored attrs (blurRad, stA, endA,
                    // dist, dir, …) deviate from ApplyReflection's preset
                    // shape. Re-install verbatim so the round-trip carries
                    // the user's tuning intact instead of collapsing to the
                    // nearest preset bucket.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var effectList = EnsureEffectList(spPr);
                    effectList.RemoveAllChildren<Drawing.Reflection>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        try
                        {
                            // OOXML SDK Reflection.OuterXml is read-only; the
                            // public mutation is via Read-Through with
                            // OpenXmlReader and Element-build. The captured
                            // slice may or may not carry xmlns:a — inject
                            // defensively so the standalone parse succeeds.
                            var raw = value.Contains("xmlns:a=")
                                ? value
                                : value.Replace("<a:reflection",
                                    "<a:reflection xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
                            using var sr = new System.IO.StringReader(raw);
                            using var xr = System.Xml.XmlReader.Create(sr);
                            xr.MoveToContent();
                            var refl = new Drawing.Reflection();
                            // Lift attributes off the parsed root.
                            if (xr.HasAttributes)
                            {
                                while (xr.MoveToNextAttribute())
                                {
                                    if (xr.Prefix == "xmlns" || xr.Name == "xmlns") continue;
                                    refl.SetAttribute(new OpenXmlAttribute(
                                        xr.Prefix, xr.LocalName, xr.NamespaceURI, xr.Value));
                                }
                                xr.MoveToElement();
                            }
                            DrawingEffectsHelper.InsertEffectInSchemaOrder(effectList, refl);
                        }
                        catch
                        {
                            unsupported.Add(key);
                        }
                    }
                    break;
                }

                case "softedge":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    // CONSISTENCY(run-context-explicit): see shadow case above.
                    if (runContext && runs.Count > 0)
                        foreach (var run in runs) ApplyTextSoftEdge(run, value);
                    else if (IsNoFillShape(spPr) && runs.Count > 0)
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

                case "lightingdir" or "lightrigdir":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyLightRigDirection(spPr, value);
                    break;
                }

                case "lightingrot" or "lightrigrot":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyLightRigRotation(spPr, value);
                    break;
                }

                case "extrusioncolor" or "extrusionclr":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplySp3DColor(spPr, value, isExtrusion: true);
                    break;
                }

                case "contourcolor" or "contourclr":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplySp3DColor(spPr, value, isExtrusion: false);
                    break;
                }

                case "camera" or "camerapreset" or "cameraprst":
                {
                    // a:scene3d/a:camera/@prst. Accept the raw OOXML name
                    // (orthographicFront, perspectiveContrastingRightFacing,
                    // isometricTopUp, …) — there are 62 presets and
                    // maintaining a lowercase-alias table would just lag the
                    // schema. PresetCameraValues' string ctor accepts the
                    // OOXML inner text directly.
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyCameraPreset(spPr, value);
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

                    // R3-fuzz-1: lenient parse (warn + exit 2 + placeholder)
                    // instead of throwing on a too-deep/unparseable formula.
                    var mathContent = FormulaParser.ParseLenient(value, unrecognizedLatex);
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
                        // CONSISTENCY(rpr-bool-form): drawingML rPr xsd:boolean
                        // attrs (b/i/noProof/normalizeH/dirty/err/smtClean/kumimoji)
                        // must serialise as the lexical "1"/"0" PowerPoint
                        // authors — passing through "true"/"false" produces
                        // a file whose attrs PowerPoint accepts on read but
                        // textual byte-diffs treat as drift. Get normalises
                        // both wire forms to "true"/"false" for cross-handler
                        // vocabulary parity; pin the write-side here so
                        // dump→Get→replay round-trips the canonical form
                        // instead of leaking the canonical-readback string
                        // back onto disk. Mirrors the OneOnBool() pin on
                        // hMerge / vMerge (R43 779099bc).
                        string writeValue = value;
                        if (DrawingRunBoolAttrs.Contains(key))
                            writeValue = (value is "1" or "true" or "True") ? "1" : "0";
                        foreach (var run in runs)
                        {
                            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            if (clearAttr)
                                rPr.RemoveAttribute(key, "");
                            else
                                rPr.SetAttribute(new OpenXmlAttribute("", key, "", writeValue));
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
                            // Derive the valid-prop list dynamically from the pptx
                            // shape schema (verb=set) so the hint never drifts from
                            // the actually-supported props (font.latin, strike, adj,
                            // lineSpacing, spaceBefore, cap, autoFit, …). The schema
                            // is kept in sync with the handler via schema-verify
                            // tooling, so this stays accurate without hand-editing a
                            // literal string. Fall back to a short literal only if
                            // the schema can't be loaded.
                            var dynProps = OfficeCli.Help.SchemaHelpLoader.ListProperties("pptx", "shape", "set");
                            var msg = unsupportedContextHint
                                ?? (dynProps.Count > 0
                                    ? "valid shape props: " + string.Join(", ", dynProps)
                                    : "valid shape props: text, bold, italic, underline, color, fill, size, font, gradient, line, opacity, align, valign, x, y, width, height, rotation, name, link, animation, formula, geometry, preset, shadow, glow, reflection, softEdge, pattern, flip, flipH, flipV");
                            unsupported.Add($"{key} ({msg})");
                        }
                        else
                            unsupported.Add(key);
                    }
                    break;
                }
            }
        }

        // RC1: transplant the scratch run's resolved RunProperties onto the
        // target paragraph's <a:pPr><a:defRPr>. defRPr shares CT_TextCharacter-
        // Properties with rPr, so the children (solidFill, latin, etc.) copy
        // 1:1; only the element name differs. Schema order is enforced by
        // ReorderDrawingRunProperties on the source rPr before the copy.
        if (defRPrScratchRun != null && defRPrTargetPara != null)
        {
            var scratchProps = defRPrScratchRun.RunProperties;
            if (scratchProps != null && (scratchProps.HasChildren || scratchProps.HasAttributes))
            {
                var pPr = defRPrTargetPara.GetFirstChild<Drawing.ParagraphProperties>();
                if (pPr == null)
                {
                    pPr = new Drawing.ParagraphProperties();
                    defRPrTargetPara.InsertAt(pPr, 0);
                }
                pPr.RemoveAllChildren<Drawing.DefaultRunProperties>();
                var defRPr = new Drawing.DefaultRunProperties();
                // Copy attributes (sz, b, i, …) and child elements (solidFill,
                // latin font, …) from the scratch rPr onto the defRPr.
                foreach (var attr in scratchProps.GetAttributes())
                    defRPr.SetAttribute(attr);
                foreach (var child in scratchProps.ChildElements)
                    defRPr.AppendChild((OpenXmlElement)child.CloneNode(true));
                // defRPr must be the LAST child of a:pPr per CT_TextParagraphProperties.
                pPr.AppendChild(defRPr);
            }
        }

        return unsupported;
    }

    // RC1: run-format keys that, on a runless placeholder, are written to the
    // paragraph's <a:defRPr> instead of being silently dropped. Shape-level
    // props (fill, geometry, x/y, …) are intentionally excluded — they target
    // the shape regardless of run count and must not trigger defRPr injection.
    private static readonly HashSet<string> RunFormatDefRPrKeys = new(StringComparer.Ordinal)
    {
        "size", "fontsize", "font.size",
        "color", "font.color",
        "bold", "font.bold", "italic", "font.italic",
        "underline", "font.underline", "u",
        "strike", "font.strike",
        "font", "font.name", "font.latin", "font.ea", "font.eastasia", "font.eastasian",
        "font.cs", "font.complexscript", "font.complex",
        "highlight", "spc", "kern", "lang", "altlang", "baseline",
        "sz", "b", "i",
        // NOTE: "cap" intentionally EXCLUDED — it already has its own textless-
        // shape fallback (writes to the first paragraph's endParaRPr, see the
        // `case "cap"` below). Routing it through defRPr would break that
        // established behavior (AuditPptxAddSet*Cap* tests).
    };

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
        // Drop any extra paragraphs left by a previous multi-line value so the
        // cell rebuilds cleanly from the first paragraph.
        foreach (var extra in txBody.Elements<Drawing.Paragraph>().Skip(1).ToList())
            extra.Remove();
        para.RemoveAllChildren<Drawing.Run>();
        para.RemoveAllChildren<Drawing.Break>();
        var savedEndParaRPr = para.Elements<Drawing.EndParagraphRunProperties>().FirstOrDefault();
        if (savedEndParaRPr != null)
            savedEndParaRPr.Remove();
        if (!string.IsNullOrEmpty(value))
        {
            // CONSISTENCY(text-escape-boundary): \n → paragraph break, \t → tab,
            // exactly like the cell `text=` case and pptx shape text=. The row
            // c1…cN shortcut routed here must not diverge from them.
            Drawing.Run RunFactory(string seg) => new Drawing.Run(
                new Drawing.RunProperties { Language = "en-US" }, MakePreservingText(seg));
            var lines = value.Split('\n');
            AppendLineWithTabs(para, lines[0], RunFactory);
            for (int li = 1; li < lines.Length; li++)
            {
                var extraPara = new Drawing.Paragraph();
                AppendLineWithTabs(extraPara, lines[li], RunFactory);
                txBody.AppendChild(extraPara);
            }
        }
        if (savedEndParaRPr != null)
            txBody.Elements<Drawing.Paragraph>().Last().AppendChild(savedEndParaRPr);
    }

    private static List<string> SetTableCellProperties(Drawing.TableCell cell, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "txbodyraw":
                {
                    // Verbatim cell text-body re-injection. The plain text=
                    // rebuild produces bare paragraphs (no pPr/lstStyle/rPr
                    // richness); the captured OuterXml restores the source's
                    // full <a:txBody> — bodyPr, lstStyle, every paragraph's
                    // pPr (lnSpc/spc/bu*/tabLst/defRPr) and every run's rPr
                    // (ea/latin/solidFill) and endParaRPr. Replace the cell's
                    // entire text body with the parsed verbatim element.
                    // The emitter suppresses the companion text= op when this
                    // key is present, so there is no clobber ordering hazard.
                    if (string.IsNullOrWhiteSpace(value)) break;
                    // Re-inject xml:space="preserve" on whitespace-bearing <a:t>
                    // before the SDK reparses, or the parser drops space-only /
                    // edge-whitespace run text (PowerPoint authors these spacer
                    // runs without the attribute). See PreserveWhitespaceInRawText.
                    var parsedBody = new Drawing.TextBody(PreserveWhitespaceInRawText(value));
                    var existingBody = cell.TextBody;
                    if (existingBody != null)
                    {
                        existingBody.InsertAfterSelf(parsedBody);
                        existingBody.Remove();
                    }
                    else
                    {
                        // txBody is the first child of a:tc (before a:tcPr).
                        cell.PrependChild(parsedBody);
                    }
                    break;
                }
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
                                MakePreservingText(seg)));
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
                                r.Text = MakePreservingText(seg);
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
                        // Empty clears the override (see run-level path above).
                        if (!string.IsNullOrEmpty(value))
                        {
                            rProps.Append(new Drawing.LatinFont { Typeface = value });
                            rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                        }
                        ReorderDrawingRunProperties(rProps);
                    }
                    break;
                case "font.latin":
                    EnsureTableCellHasRun(cell);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        if (!string.IsNullOrEmpty(value))
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
                        if (!string.IsNullOrEmpty(value))
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
                        if (!string.IsNullOrEmpty(value))
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
                case "horzoverflow":
                {
                    // R56 bt-4: a:tcPr @horzOverflow (overflow|clip). Typed SDK
                    // enum; vocabulary mirrors NodeBuilder readback.
                    var tcPrHov = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    tcPrHov.HorizontalOverflow = value.ToLowerInvariant() switch
                    {
                        "overflow" => Drawing.TextHorizontalOverflowValues.Overflow,
                        "clip" => Drawing.TextHorizontalOverflowValues.Clip,
                        _ => throw new ArgumentException($"Invalid horzOverflow value: '{value}'. Valid values: overflow, clip.")
                    };
                    break;
                }
                case "locktext":
                {
                    // R56 bt-4: a:tcPr @lockText — non-standard MS Office
                    // extension (not in the SDK enum). Set/clear via raw
                    // attribute manipulation so dump→batch round-trips.
                    var tcPrLt = cell.TableCellProperties ?? (cell.TableCellProperties = new Drawing.TableCellProperties());
                    bool lockOn = IsTruthy(value);
                    // Remove any prior occurrence first (idempotent). OpenXmlAttribute
                    // is a struct so FirstOrDefault returns default (empty LocalName)
                    // when absent — check non-empty before RemoveAttribute.
                    var existing = tcPrLt.GetAttributes().FirstOrDefault(a => a.LocalName == "lockText");
                    if (!string.IsNullOrEmpty(existing.LocalName))
                        tcPrLt.RemoveAttribute(existing.LocalName, existing.NamespaceUri);
                    if (lockOn)
                        tcPrLt.SetAttribute(new OpenXmlAttribute("lockText", "", "1"));
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
                                gsCells[mi].HorizontalMerge = OneOnBool();
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
                                        rowCells[ci].HorizontalMerge = OneOnBool();
                                        rowCells[ci].VerticalMerge = OneOnBool();
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
                                belowCells[rsColIdx2].VerticalMerge = OneOnBool();
                            for (int ci = rsColIdx2 + 1; ci < rsColIdx2 + rsAnchorGridSpan && ci < belowCells.Count; ci++)
                            {
                                belowCells[ci].HorizontalMerge = OneOnBool();
                                belowCells[ci].VerticalMerge = OneOnBool();
                            }
                        }
                    }
                    break;
                }
                // vmerge / hmerge are get-only per schema (set=false). Removed
                // from Set so the key falls through to default → unsupported
                // warning + exit 2, matching schema intent. Without this, the
                // case consumed the key and called IsTruthy() which threw on
                // non-boolean inputs like "restart" (R7 minor-2). Users
                // wanting to merge should use merge.down=N / merge.right=N.
                case "merge.right":
                {
                    // Convenience: merge.right=N sets gridSpan on this cell and hMerge on next N cells.
                    // CONSISTENCY(merge-clamp): clamp gridSpan to the cells that
                    // actually exist on this row so a high `merge.right=N` can't
                    // produce a corrupt file (PowerPoint silently misrenders
                    // gridSpan values that exceed the row's cell count).
                    var span = ParseHelpers.SafeParseInt(value, "merge.right") + 1;
                    var row = cell.Parent as Drawing.TableRow;
                    if (row != null)
                    {
                        var cells = row.Elements<Drawing.TableCell>().ToList();
                        var idx = cells.IndexOf(cell);
                        span = System.Math.Max(1, System.Math.Min(span, cells.Count - idx));
                        cell.GridSpan = new DocumentFormat.OpenXml.Int32Value(span);
                        for (int mi = idx + 1; mi < idx + span && mi < cells.Count; mi++)
                            cells[mi].HorizontalMerge = OneOnBool();
                    }
                    else
                    {
                        cell.GridSpan = new DocumentFormat.OpenXml.Int32Value(System.Math.Max(1, span));
                    }
                    break;
                }
                case "merge.down":
                {
                    // Convenience: merge.down=N sets rowSpan on this cell and vMerge on cells below.
                    // CONSISTENCY(merge-clamp): mirror the merge.right clamp so
                    // rowSpan never exceeds (rowCount - rowIdx).
                    var rSpan = ParseHelpers.SafeParseInt(value, "merge.down") + 1;
                    var row = cell.Parent as Drawing.TableRow;
                    var table = row?.Parent;
                    if (table != null && row != null)
                    {
                        var rows = table.Elements<Drawing.TableRow>().ToList();
                        var rowIdx = rows.IndexOf(row);
                        var cells = row.Elements<Drawing.TableCell>().ToList();
                        var colIdx = cells.IndexOf(cell);
                        rSpan = System.Math.Max(1, System.Math.Min(rSpan, rows.Count - rowIdx));
                        cell.RowSpan = new DocumentFormat.OpenXml.Int32Value(rSpan);
                        for (int ri = rowIdx + 1; ri < rowIdx + rSpan && ri < rows.Count; ri++)
                        {
                            var belowCells = rows[ri].Elements<Drawing.TableCell>().ToList();
                            if (colIdx < belowCells.Count)
                                belowCells[colIdx].VerticalMerge = OneOnBool();
                        }
                    }
                    else
                    {
                        cell.RowSpan = new DocumentFormat.OpenXml.Int32Value(System.Math.Max(1, rSpan));
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

                    // Verbatim border-line re-injection (border.<edge>.raw). The
                    // dump readback captures each lnL/lnR/lnT/lnB/lnTlToBr/lnBlToTr
                    // OuterXml; the granular color/width/dash path can't represent
                    // an invisible <a:noFill/> border (it skips it) nor cap/algn/
                    // prstDash/round/head-tail-end. Splice the parsed element in
                    // verbatim, replacing whatever the granular path may have built.
                    // tcPr enforces CT_TableCellProperties child order internally
                    // when the typed property setter is used.
                    if (k.EndsWith(".raw", StringComparison.Ordinal)
                        && !string.IsNullOrWhiteSpace(value))
                    {
                        switch (k)
                        {
                            case "border.left.raw":
                                tcPr.LeftBorderLineProperties = new Drawing.LeftBorderLineProperties(value);
                                break;
                            case "border.right.raw":
                                tcPr.RightBorderLineProperties = new Drawing.RightBorderLineProperties(value);
                                break;
                            case "border.top.raw":
                                tcPr.TopBorderLineProperties = new Drawing.TopBorderLineProperties(value);
                                break;
                            case "border.bottom.raw":
                                tcPr.BottomBorderLineProperties = new Drawing.BottomBorderLineProperties(value);
                                break;
                            case "border.tl2br.raw":
                                tcPr.TopLeftToBottomRightBorderLineProperties = new Drawing.TopLeftToBottomRightBorderLineProperties(value);
                                break;
                            case "border.tr2bl.raw":
                                tcPr.BottomLeftToTopRightBorderLineProperties = new Drawing.BottomLeftToTopRightBorderLineProperties(value);
                                break;
                            default:
                                throw new ArgumentException($"Unknown border raw key: '{k}'.");
                        }
                        break;
                    }

                    // Handle "none" — remove border by adding NoFill
                    bool isNone = value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase);

                    // Parse value: "FF0000", "1pt solid FF0000", "2pt dash 0000FF", or "style;width;color;dash"
                    string? borderColor = null;
                    long? borderWidth = null;
                    string? borderDash = null;
                    Drawing.CompoundLineValues? borderCompound = null;
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
                    // Compound line style sub-key (.compound). The NodeBuilder
                    // readback emits the raw OOXML cmpd attr value (sng/dbl/
                    // thickThin/thinThick/tri). Without this routing it fell
                    // through to the space-split path, where "sng" failed the
                    // pt/dash checks and the color parser uppercased it to
                    // "SNG" → "Invalid color value: 'SNG'", aborting the op.
                    bool isCompoundOnly = k.EndsWith(".compound", StringComparison.Ordinal);
                    if (!isNone && (isWidthOnly || isColorOnly || isDashOnly || isCompoundOnly))
                    {
                        if (isCompoundOnly)
                        {
                            borderCompound = value.ToLowerInvariant() switch
                            {
                                "sng" or "single" => Drawing.CompoundLineValues.Single,
                                "dbl" or "double" => Drawing.CompoundLineValues.Double,
                                "thickthin" => Drawing.CompoundLineValues.ThickThin,
                                "thinthick" => Drawing.CompoundLineValues.ThinThick,
                                "tri" or "triple" => Drawing.CompoundLineValues.Triple,
                                _ => throw new ArgumentException($"Invalid border compound value: '{value}'. Valid values: sng, dbl, thickThin, thinThick, tri.")
                            };
                        }
                        else if (isWidthOnly)
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
                        else if (value.Contains(':'))
                        {
                            // Colon-delimited form: width:color[:dash], mirroring
                            // the shape `line=` compound form. Without this branch
                            // "2pt:FF0000" landed in the space-split path as a
                            // single token, failed the pt/cm/px check, and the
                            // color parser uppercased the whole thing into
                            // "2PT:FF0000" — surfacing as "Invalid color value".
                            var coParts = value.Split(':');
                            if (coParts.Length > 0 && !string.IsNullOrEmpty(coParts[0]))
                            {
                                var wStr = coParts[0];
                                if (!wStr.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                                    && !wStr.EndsWith("cm", StringComparison.OrdinalIgnoreCase)
                                    && !wStr.EndsWith("px", StringComparison.OrdinalIgnoreCase)
                                    && !wStr.EndsWith("emu", StringComparison.OrdinalIgnoreCase))
                                    wStr += "pt";
                                borderWidth = Core.EmuConverter.ParseEmu(wStr);
                                if (borderWidth.Value < 0)
                                    throw new ArgumentException($"Invalid border width: '{coParts[0]}' (must be >= 0).");
                            }
                            if (coParts.Length > 1 && !string.IsNullOrEmpty(coParts[1]))
                                borderColor = coParts[1].TrimStart('#').ToUpperInvariant();
                            if (coParts.Length > 2)
                            {
                                var d = coParts[2].ToLowerInvariant();
                                if (d is "solid" or "dot" or "dash" or "lgdash" or "dashdot" or "sysdot" or "sysdash")
                                    borderDash = d;
                                else
                                    throw new ArgumentException($"Invalid border dash value: '{coParts[2]}'. Valid values: solid, dot, dash, lgDash, dashDot, sysDot, sysDash.");
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

                    // Build line properties
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
                        // Remove NoFill if present
                        lineProps.RemoveAllChildren<Drawing.NoFill>();
                        // Set width (default 12700 EMU = 1pt)
                        if (borderWidth.HasValue)
                        {
                            var wAttr = lineProps.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
                            lineProps.SetAttribute(new OpenXmlAttribute("", "w", null!, borderWidth.Value.ToString()));
                        }
                        // Set compound line style (cmpd attr on the line element).
                        if (borderCompound.HasValue && lineProps is Drawing.LinePropertiesType lpCmpd)
                        {
                            lpCmpd.CompoundLineType = borderCompound.Value;
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

                    // CONSISTENCY(border-edge-aliases): accept the OOXML SDK
                    // property-name forms (borderTopLeftToBottomRight /
                    // borderBottomLeftToTopRight) alongside the canonical
                    // border.tl2br / border.tr2bl. The verbose forms match
                    // the SDK's TopLeftToBottomRightBorderLineProperties /
                    // BottomLeftToTopRightBorderLineProperties property
                    // names — agents and tools that introspect via
                    // reflection naturally produce them. Without these
                    // aliases the key (border-prefixed, .color suffix)
                    // matched the outer `border` case but no specific edge,
                    // and fell through to the default `new[] { left, right,
                    // top, bottom }` arm, silently writing the diagonal
                    // border to all four straight edges instead.
                    var edges = k switch
                    {
                        "border.left" or "border.left.width" or "border.left.color" or "border.left.dash" or "border.left.compound" => new[] { "left" },
                        "border.right" or "border.right.width" or "border.right.color" or "border.right.dash" or "border.right.compound" => new[] { "right" },
                        "border.top" or "border.top.width" or "border.top.color" or "border.top.dash" or "border.top.compound" => new[] { "top" },
                        "border.bottom" or "border.bottom.width" or "border.bottom.color" or "border.bottom.dash" or "border.bottom.compound" => new[] { "bottom" },
                        "border.tl2br" or "border.tl2br.width" or "border.tl2br.color" or "border.tl2br.dash" or "border.tl2br.compound" => new[] { "tl2br" },
                        "border.tr2bl" or "border.tr2bl.width" or "border.tr2bl.color" or "border.tr2bl.dash" or "border.tr2bl.compound" => new[] { "tr2bl" },
                        "bordertoplefttobottomright"
                          or "bordertoplefttobottomright.width"
                          or "bordertoplefttobottomright.color"
                          or "bordertoplefttobottomright.dash" => new[] { "tl2br" },
                        "borderbottomlefttotopright"
                          or "borderbottomlefttotopright.width"
                          or "borderbottomlefttotopright.color"
                          or "borderbottomlefttotopright.dash" => new[] { "tr2bl" },
                        // PowerPoint UI naming: diagDown = top-left → bottom-right
                        // (line slopes downward L→R); diagUp = bottom-left → top-right.
                        // Without these arms, "border.diagdown" fell through to the
                        // 4-side fallback and applied to every orthogonal border.
                        "border.diagdown" or "border.diagdown.width"
                          or "border.diagdown.color" or "border.diagdown.dash" => new[] { "tl2br" },
                        "border.diagup" or "border.diagup.width"
                          or "border.diagup.color" or "border.diagup.dash" => new[] { "tr2bl" },
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
                    // OOXML semantics: Vertical = 90° CCW (bottom-to-top), Vertical270 =
                    // 270° CCW (top-to-bottom). Old code collapsed both "vert" and
                    // "vert270" to Vertical270 (wrong rotation for "vert"). Split, and
                    // accept eaVert (East Asian vertical) consistent with shape-level set.
                    tcPrTd.Vertical = value.ToLowerInvariant() switch
                    {
                        "horizontal" or "horz" or "none" => Drawing.TextVerticalValues.Horizontal,
                        "vertical" or "vert" or "vertical90" or "vert90" => Drawing.TextVerticalValues.Vertical,
                        "vertical270" or "vert270" => Drawing.TextVerticalValues.Vertical270,
                        // Note: SDK enum spelling is EastAsianVetical (typo); XML is "eaVert".
                        "eavert" or "eavertical" => Drawing.TextVerticalValues.EastAsianVetical,
                        "stacked" or "wordartvert" => Drawing.TextVerticalValues.WordArtVertical,
                        _ => throw new ArgumentException($"Invalid textDirection: '{value}'. Valid: horizontal, vertical (=vert / vertical90, 90° CCW), vertical270 (=vert270, 270° CCW), eaVert, stacked.")
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
                            bevel.Width = (long)(ParseHelpers.SafeParseDouble(bevelParts[1], "bevel width") * EmuConverter.EmuPerPoint); // pt to EMU
                        if (bevelParts.Length >= 3)
                            bevel.Height = (long)(ParseHelpers.SafeParseDouble(bevelParts[2], "bevel height") * EmuConverter.EmuPerPoint);
                        cell3d.AppendChild(bevel);
                    }
                    break;
                }
                case "image":
                {
                    // Validate before modifying (atomic: no data loss on invalid input)
                    if (!File.Exists(value))
                        throw new FileNotFoundException($"Image file not found: {value}");

                    // Image fill on table cell
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
        var shapeTree = slideParts[PathIndex.ToArrayIndex(slideIdx)].Slide?.CommonSlideData?.ShapeTree;
        var shapes = shapeTree?.Elements<Shape>().ToList();
        if (shapes == null || shapeIdx < 1 || shapeIdx > shapes.Count) return null;
        return CheckTextOverflow(shapes[PathIndex.ToArrayIndex(shapeIdx)]);
    }

    /// <summary>
    /// Per-preset inscribed text rectangle, as width/height fractions of the
    /// shape's bounding box. Sourced from the OOXML presetShapeDefinitions
    /// &lt;rect&gt; element (the region PowerPoint actually lays text into):
    /// a diamond's text area is the centered w/2 × h/2 rectangle, an ellipse's
    /// is the cos45° inscribed rect, etc. Presets not listed use the full box.
    /// </summary>
    // Keyed by the preset's OOXML InnerText: SDK v3 enum-structs render as
    // "ShapeTypeValues { }" under ToString(), so InnerText is the stable key.
    private static readonly Dictionary<string, (double W, double H)> PresetTextRectFactors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["diamond"] = (0.50, 0.50),
        ["ellipse"] = (0.7071, 0.7071),
        ["triangle"] = (0.50, 0.50),
        ["rtTriangle"] = (0.50, 0.3333),
        ["pentagon"] = (0.60, 0.6667),
        ["hexagon"] = (0.7071, 0.7071),
        ["octagon"] = (0.7071, 0.7071),
        ["cloud"] = (0.60, 0.55),
        ["teardrop"] = (0.7071, 0.7071),
    };

    /// <summary>
    /// Estimates whether the given text will overflow the shape bounds.
    /// Uses per-character width estimation (CJK vs Latin) and reads actual line spacing from the shape.
    /// Non-rectangular presets shrink the usable area to their inscribed text
    /// rect (a diamond holds ~half the width/height its bounding box suggests).
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

        const double emuPerPt = EmuConverter.EmuPerPointF;
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

        // Non-rectangular presets: text lives in the preset's inscribed rect,
        // not the bounding box — shrink the box before subtracting insets.
        var presetName = spPr?.GetFirstChild<Drawing.PresetGeometry>()?.Preset?.InnerText;
        string presetNote = "";
        double heightFactor = 1.0;
        if (presetName != null && PresetTextRectFactors.TryGetValue(presetName, out var f))
        {
            shapeWidthPt *= f.W;
            shapeHeightPt *= f.H;
            heightFactor = f.H;
            presetNote = $" ({presetName} text area is the inscribed {f.W * 100:F0}%×{f.H * 100:F0}% rect of the bounding box)";
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
            long minHeightEmu = (long)Math.Round(minHeightCm * EmuConverter.EmuPerCmF);
            return $"text overflow: need ≥{defaultLinePt:F0}pt, usable 0pt (shape {shapeHeightPt:F0}pt < margins {marginPt:F0}pt){presetNote}. suggest.height={EmuConverter.FormatEmu(minHeightEmu)}";
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
            var pct = lsEl.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
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

        // Line height: fixed spacing overrides multiplier. Percent/default
        // spacing is relative to the font's line pitch (~1.2× the font size for
        // Latin faces, ~1.32× for CJK faces per hhea/OS2 metrics), NOT 1.0× —
        // estimating at 1.0× made the check miss real overflows that PowerPoint
        // visibly clips (issue #236). Resolve the ratio from the first run
        // typeface; unlocatable fonts fall back to the 1.2 Latin norm.
        double singleRatio = SingleSpacingPitch(
            paragraphs.SelectMany(p => p.Elements<Drawing.Run>()).FirstOrDefault()?.RunProperties);
        double lineHeight = fixedLineSpacingPt ?? fontSizePt * lineSpacingMultiplier * singleRatio;
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
            // Minimum height: text + margins must fit the INSCRIBED rect, so a
            // shrunk preset scales the suggested bounding-box height back up.
            double minHeightCm = (estimatedHeight + marginPt) / heightFactor / 72.0 * 2.54;
            // Round up to 0.05cm for cleaner values
            minHeightCm = Math.Ceiling(minHeightCm * 20) / 20.0;
            long minHeightEmu = (long)Math.Round(minHeightCm * EmuConverter.EmuPerCmF);
            return $"text overflow: {totalLines} lines at {fontSizePt:F1}pt need {estimatedHeight:F0}pt, usable {usableHeight:F0}pt{presetNote}. suggest.height={EmuConverter.FormatEmu(minHeightEmu)}";
        }
        return null;
    }

}
