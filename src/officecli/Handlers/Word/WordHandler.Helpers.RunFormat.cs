// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{

    /// <summary>
    /// P1-7: detect the "Nlines" suffix on a spacing value and convert it to
    /// hundredths of a line (the unit of `<w:spacing w:beforeLines/afterLines>`).
    /// Returns false when the value lacks the "lines" suffix so the caller can
    /// fall through to the points/twips path.
    /// </summary>
    internal static bool TryParseLinesSuffix(string value, out string hundredthsOfLine)
    {
        hundredthsOfLine = "";
        var trimmed = (value ?? "").Trim();
        if (!trimmed.EndsWith("lines", StringComparison.OrdinalIgnoreCase) &&
            !trimmed.EndsWith("line", StringComparison.OrdinalIgnoreCase))
            return false;
        var num = trimmed.EndsWith("lines", StringComparison.OrdinalIgnoreCase) ? trimmed[..^5] : trimmed[..^4];
        if (!double.TryParse(num.Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var n) || double.IsNaN(n) || double.IsInfinity(n) || n < 0)
            throw new ArgumentException($"Invalid lines value '{value}'. Expected a non-negative number with 'lines' suffix (e.g. '0.5lines', '1lines').");
        hundredthsOfLine = ((int)Math.Round(n * 100)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        return true;
    }

    /// <summary>
    /// Parse a lineRule prop value (auto / exact / atLeast) into the OOXML
    /// enum. BUG-019 — needed to distinguish AtLeast from Exact since
    /// SpacingConverter.FormatWordLineSpacing serializes both as "Npt".
    /// </summary>
    internal static LineSpacingRuleValues ParseLineRule(string value)
    {
        var v = (value ?? "").Trim().ToLowerInvariant();
        return v switch
        {
            "auto" => LineSpacingRuleValues.Auto,
            "exact" => LineSpacingRuleValues.Exact,
            "atleast" => LineSpacingRuleValues.AtLeast,
            _ => throw new ArgumentException(
                $"Invalid lineRule '{value}'. Expected: auto, exact, or atLeast."),
        };
    }

    /// <summary>
    /// Normalize a user-provided underline token to a valid Word OOXML UnderlineValues enum string.
    /// Accepts common aliases (wavy → wave, dashdot → dotDash, etc.) plus truthy/none.
    /// </summary>
    internal static string NormalizeUnderlineValue(string value)
    {
        var v = (value ?? "").Trim();
        var mapped = v.ToLowerInvariant() switch
        {
            "true" or "single" or "1" => "single",
            "false" or "none" or "0" or "" => "none",
            "double" => "double",
            "thick" => "thick",
            "dotted" => "dotted",
            "dottedheavy" or "dotted-heavy" or "dotted_heavy" => "dottedHeavy",
            "dash" or "dashed" => "dash",
            "dashedheavy" or "dashheavy" => "dashedHeavy",
            "dashlong" or "longdash" => "dashLong",
            "dashlongheavy" or "longdashheavy" => "dashLongHeavy",
            // Word uses "dotDash" and "dashDotHeavy" (note asymmetric casing in OOXML spec).
            "dotdash" or "dashdot" => "dotDash",
            "dotdashheavy" or "dashdotheavy" => "dashDotHeavy",
            "dotdotdash" or "dashdotdot" => "dotDotDash",
            "dotdotdashheavy" or "dashdotdotheavy" => "dashDotDotHeavy",
            "wave" or "wavy" => "wave",
            "waveheavy" or "wavyheavy" => "wavyHeavy",
            "wavedouble" or "wavydouble" or "doublewave" => "wavyDouble",
            "words" or "word" => "words",
            _ => v  // pass-through for already-valid OOXML tokens
        };
        // CONSISTENCY(allowlist): mirror tab val/leader allowlist (R1 a1554d59) and
        // ParseJustification — validate before handing off to OpenXML SDK to avoid
        // leaking "specified value is not valid according to the specified enum type".
        if (!ValidUnderlineValues.Contains(mapped))
            throw new ArgumentException(
                $"Invalid underline value: '{value}'. Valid values: single, double, thick, dotted, " +
                "dottedHeavy, dash, dashedHeavy, dashLong, dashLongHeavy, dotDash, dashDotHeavy, " +
                "dotDotDash, dashDotDotHeavy, wave, wavyHeavy, wavyDouble, words, none.");
        return mapped;
    }

    private static readonly HashSet<string> ValidUnderlineValues = new(StringComparer.Ordinal)
    {
        "single", "double", "thick", "dotted", "dottedHeavy",
        "dash", "dashedHeavy", "dashLong", "dashLongHeavy",
        "dotDash", "dashDotHeavy", "dotDotDash", "dashDotDotHeavy",
        "wave", "wavyHeavy", "wavyDouble", "words", "none"
    };

    /// <summary>
    /// Apply a <c>tabs=POS:ALIGN[:LEADER],POS:ALIGN[:LEADER]...</c>
    /// shorthand to a paragraph properties container (paragraph
    /// <c>w:pPr</c> or style <c>w:pPr</c> alike). Each segment becomes a
    /// <c>w:tab</c> child of the container's <c>w:tabs</c> element. Existing
    /// <c>w:tabs</c> is replaced wholesale so a new shorthand defines the
    /// definitive tab strip — partial-merge would surprise callers who
    /// expect "set tabs=…" to mean "this is the tab strip now".
    ///
    /// <para>Supported forms (case-insensitive):</para>
    /// <list type="bullet">
    ///   <item><c>9360:right</c></item>
    ///   <item><c>9360:right:dot</c></item>
    ///   <item><c>2880:center,5760:decimal,9360:right:dot</c></item>
    ///   <item><c>5cm:left</c> / <c>2in:right</c> (unit suffix on pos)</item>
    /// </list>
    ///
    /// <para>
    /// <c>ALIGN</c>: left, center, right, decimal, bar, clear, num,
    /// start, end. <c>LEADER</c>: none, dot, heavy, hyphen, middleDot,
    /// underscore.
    /// </para>
    /// </summary>
    internal static void ApplyTabsShorthand(OpenXmlCompositeElement pPr, string tabsStr)
    {
        if (string.IsNullOrWhiteSpace(tabsStr))
        {
            // Empty value clears any existing tab strip — useful for
            // overriding inherited tabs from basedOn.
            pPr.RemoveAllChildren<Tabs>();
            return;
        }

        var newTabs = new Tabs();
        foreach (var rawSeg in tabsStr.Split(','))
        {
            var seg = rawSeg.Trim();
            if (seg.Length == 0) continue;
            var parts = seg.Split(':');
            if (parts.Length < 1 || string.IsNullOrWhiteSpace(parts[0]))
                throw new ArgumentException(
                    $"Invalid tabs segment '{seg}'. Expected POS[:ALIGN[:LEADER]] (e.g. 9360:right or 5cm:right:dot).");

            // pos: allow negative twips for hanging-tab positions, accept
            // bare twips OR unit suffix (pt/cm/in). Same parser as `add
            // /body/p[N] --type tab --prop pos=…`.
            int posTwips;
            try { posTwips = ParseSignedTwips(parts[0]); }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Invalid tab pos '{parts[0]}' in tabs segment '{seg}': {ex.Message}");
            }

            var tabStop = new TabStop { Position = posTwips };

            if (parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[1]))
            {
                var alignNorm = parts[1].Trim().ToLowerInvariant();
                var knownTabVals = new[] { "left", "center", "right", "decimal", "bar", "clear", "num", "start", "end" };
                if (!knownTabVals.Contains(alignNorm))
                    throw new ArgumentException(
                        $"Invalid tab align '{parts[1]}' in tabs segment '{seg}'. Valid: {string.Join(", ", knownTabVals)}.");
                tabStop.Val = new EnumValue<TabStopValues>(new TabStopValues(alignNorm));
            }
            else
            {
                tabStop.Val = TabStopValues.Left;
            }

            if (parts.Length >= 3 && !string.IsNullOrWhiteSpace(parts[2]))
            {
                var leaderNorm = parts[2].Trim().ToLowerInvariant();
                tabStop.Leader = leaderNorm switch
                {
                    "none"       => TabStopLeaderCharValues.None,
                    "dot"        => TabStopLeaderCharValues.Dot,
                    "heavy"      => TabStopLeaderCharValues.Heavy,
                    "hyphen"     => TabStopLeaderCharValues.Hyphen,
                    "middledot"  => TabStopLeaderCharValues.MiddleDot,
                    "underscore" => TabStopLeaderCharValues.Underscore,
                    _ => throw new ArgumentException(
                        $"Invalid tab leader '{parts[2]}' in tabs segment '{seg}'. Valid: none, dot, heavy, hyphen, middleDot, underscore."),
                };
            }

            newTabs.Append(tabStop);
        }

        // Replace any existing tabs strip with the new one. Schema places
        // <w:tabs> early in pPr; PrependChild keeps schema order without
        // having to compute the exact slot.
        pPr.RemoveAllChildren<Tabs>();
        pPr.PrependChild(newTabs);
    }

    private static JustificationValues ParseJustification(string value) =>
        value.ToLowerInvariant() switch
        {
            "left" => JustificationValues.Left,
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "justify" or "both" => JustificationValues.Both,
            // BUG-R7-04 (F-4): w:jc="distribute" stretches every line
            // (including the last) — used in CJK/Thai documents to fill
            // the column. Was rejected by the white-list even though
            // OOXML / Word accept it (see HtmlPreview.Css distribute
            // branch). Mirror Word's tolerant parser for the rest of the
            // ECMA-376 ST_Jc enum: highKashida/mediumKashida/lowKashida
            // (Arabic), thaiDistribute, numTab.
            "distribute" => JustificationValues.Distribute,
            "thaidistribute" => JustificationValues.ThaiDistribute,
            "highkashida" => JustificationValues.HighKashida,
            "mediumkashida" => JustificationValues.MediumKashida,
            "lowkashida" => JustificationValues.LowKashida,
            "numtab" => JustificationValues.NumTab,
            "start" => JustificationValues.Left, // bidi-aware alias
            "end" => JustificationValues.Right,
            _ => throw new ArgumentException($"Invalid alignment value: '{value}'. Valid values: left, center, right, justify, distribute, thaiDistribute, start, end.")
        };

    /// <summary>
    /// Sanitize a hex color for Word OOXML (ST_HexColorRGB = exactly 6-char RGB).
    /// Strips # prefix, uppercases, and handles 8-char AARRGGBB by extracting RGB portion.
    /// </summary>
    private static string SanitizeHex(string value)
    {
        var (rgb, alphaPercent) = ParseHelpers.SanitizeColorForOoxml(value);
        // BUG-R6-07: ARGB input (e.g. `80FF0000`) was silently truncated to
        // RGB. OOXML's w:color stores only 6-digit RGB so the alpha
        // channel cannot be preserved here. Emit a stderr warning so
        // callers know the input was lossy rather than rejected.
        if (alphaPercent.HasValue)
        {
            try
            {
                Console.Error.WriteLine(
                    $"WARNING: color value '{value}' has an alpha component which OOXML w:color cannot store. Stored as #{rgb} (alpha discarded).");
            }
            catch { /* best effort — never fail the operation over a warning */ }
        }
        return rgb;
    }

    /// <summary>
    /// Sanitize a font name input for the per-script font slots. Strips
    /// a leading BOM (U+FEFF) — font names are token-like strings, and
    /// a stray BOM (commonly produced by Windows clipboard / shell
    /// quoting paths) breaks Word's font lookup and round-trips back
    /// into OOXML as a literal U+FEFF byte attached to the typeface
    /// name. Surrounding ASCII whitespace is trimmed as well.
    /// </summary>
    private static string SanitizeFontTokenInput(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var s = value!;
        while (s.Length > 0 && s[0] == '﻿') s = s.Substring(1);
        while (s.Length > 0 && s[s.Length - 1] == '﻿') s = s.Substring(0, s.Length - 1);
        return s.Trim();
    }

    /// <summary>
    /// True when a w:rFonts element carries no value-bearing attribute and
    /// can be safely removed from its parent rPr / rPrChange.
    /// </summary>
    private static bool RunFontsIsEmpty(RunFonts rf) =>
        string.IsNullOrEmpty(rf.Ascii?.Value)
        && string.IsNullOrEmpty(rf.HighAnsi?.Value)
        && string.IsNullOrEmpty(rf.EastAsia?.Value)
        && string.IsNullOrEmpty(rf.ComplexScript?.Value)
        && string.IsNullOrEmpty(rf.AsciiTheme?.InnerText)
        && string.IsNullOrEmpty(rf.HighAnsiTheme?.InnerText)
        && string.IsNullOrEmpty(rf.EastAsiaTheme?.InnerText)
        && string.IsNullOrEmpty(rf.ComplexScriptTheme?.InnerText)
        && string.IsNullOrEmpty(rf.Hint?.InnerText);

    /// <summary>
    /// Parse a highlight color name, throwing ArgumentException with valid options on failure.
    /// </summary>
    private static readonly HashSet<string> ValidHighlightColors = new(StringComparer.OrdinalIgnoreCase)
    {
        "yellow", "green", "cyan", "magenta", "blue", "red",
        "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
        "darkGray", "lightGray", "black", "white", "none"
    };

    private static HighlightColorValues ParseHighlightColor(string value)
    {
        if (!ValidHighlightColors.Contains(value))
            throw new ArgumentException(
                $"Invalid 'highlight' value '{value}'. Valid values: yellow, green, cyan, magenta, blue, red, " +
                $"darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, white, none.");
        return new HighlightColorValues(value);
    }

    /// <summary>
    /// Warn if a value that should be a shading pattern name looks like a hex color instead.
    /// </summary>
    private static void WarnIfShadingOrderWrong(string patternSegment)
    {
        var trimmed = patternSegment.TrimStart('#');
        if (trimmed.Length >= 6 && trimmed.All(char.IsAsciiHexDigit))
            Console.Error.WriteLine($"Warning: '{patternSegment}' looks like a color, but is in the pattern position. "
                + "Shading format: FILL (single value) or PATTERN;FILL[;COLOR] e.g. clear;FF0000");
    }

    private static double ParseFontSize(string value) =>
        ParseHelpers.ParseFontSize(value);

    // CONSISTENCY(run-special-content): true when <paramref name="key"/>
    // names a typography property that has no glyph to apply on a ptab /
    // fieldChar / instrText / tab / break run. Used by SetElementRun to
    // reject cosmetic writes on these runs, mirroring the readback strip.
    private static bool IsTypographyOnlyKey(string key)
    {
        var k = key.ToLowerInvariant();
        return k is "font" or "font.ascii" or "font.eastasia" or "font.hansi" or "font.cs"
            or "size"
            or "bold" or "italic"
            or "color"
            or "underline" or "underline.color"
            or "strike" or "dstrike" or "highlight"
            or "caps" or "smallcaps" or "vanish"
            or "outline" or "shadow" or "emboss" or "imprint"
            or "noproof" or "rtl"
            or "superscript" or "subscript"
            or "charspacing" or "shading";
    }

    // CONSISTENCY(run-special-content): typography-only Format keys that
    // get scrubbed from runs whose Type was upgraded to ptab / fieldChar /
    // instrText / tab / break. These properties are valid in the underlying
    // <w:rPr> but have no glyph to apply to on these specialized runs, so
    // surfacing them is noise that primes audit tools to misread cosmetic
    // styling on a structural marker as meaningful.
    private static readonly string[] TypographyOnlyKeys =
    {
        "font.ascii", "font.eastAsia", "font.hAnsi", "font.cs",
        "size", "bold", "italic", "color",
        "underline", "underline.color",
        "strike", "dstrike", "highlight",
        "caps", "smallcaps", "vanish",
        "outline", "shadow", "emboss", "imprint",
        "noproof", "rtl", "superscript", "subscript",
        "charSpacing", "shading",
        "effective.size", "effective.size.src",
        "effective.font.ascii", "effective.font.ascii.src",
        "effective.font.eastAsia", "effective.font.eastAsia.src",
        "effective.font.hAnsi", "effective.font.hAnsi.src",
        "effective.font.cs", "effective.font.cs.src",
        "effective.bold", "effective.bold.src",
        "effective.italic", "effective.italic.src",
        "effective.color", "effective.color.src",
        "effective.underline", "effective.underline.src",
    };

    // BUG-DUMP-R46-FFSIZE: typography keys a FORM-FIELD begin fieldChar keeps
    // (instead of shedding as noise) so the field-run face/size round-trips into
    // `add formfield` and a <w:sizeAuto/> checkbox renders at its source size.
    // Mirrors the keys AddFormField's field-run formatting honors.
    private static readonly HashSet<string> FieldRunFormatKeepKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "size", "bold", "italic", "color", "underline", "strike",
        "font.ascii", "font.hAnsi", "font.eastAsia", "font.cs",
    };

    // CONSISTENCY(run-special-content): canonical parsers for the run-internal
    // structural types (ptab / fldChar / break) shared by Add and Set.
    // Lowercase XML attribute values are the canonical input; legacy
    // synonyms (`line`→TextWrapping) are accepted for ergonomics.
    private static EnumValue<AbsolutePositionTabAlignmentValues> ParsePtabAlignment(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "left" => AbsolutePositionTabAlignmentValues.Left,
            "center" => AbsolutePositionTabAlignmentValues.Center,
            "right" => AbsolutePositionTabAlignmentValues.Right,
            _ => throw new ArgumentException(
                $"Invalid ptab alignment '{s}'. Valid: left, center, right.")
        };
    }

    private static EnumValue<AbsolutePositionTabPositioningBaseValues> ParsePtabRelativeTo(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "margin" => AbsolutePositionTabPositioningBaseValues.Margin,
            "indent" => AbsolutePositionTabPositioningBaseValues.Indent,
            _ => throw new ArgumentException(
                $"Invalid ptab relativeTo '{s}'. Valid: margin, indent.")
        };
    }

    private static EnumValue<AbsolutePositionTabLeaderCharValues> ParsePtabLeader(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "none" => AbsolutePositionTabLeaderCharValues.None,
            "dot" => AbsolutePositionTabLeaderCharValues.Dot,
            "hyphen" => AbsolutePositionTabLeaderCharValues.Hyphen,
            "middledot" => AbsolutePositionTabLeaderCharValues.MiddleDot,
            "underscore" => AbsolutePositionTabLeaderCharValues.Underscore,
            _ => throw new ArgumentException(
                $"Invalid ptab leader '{s}'. Valid: none, dot, hyphen, middleDot, underscore.")
        };
    }

    private static EnumValue<FieldCharValues> ParseFieldCharType(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "begin" => FieldCharValues.Begin,
            "separate" => FieldCharValues.Separate,
            "end" => FieldCharValues.End,
            _ => throw new ArgumentException(
                $"Invalid fieldCharType '{s}'. Valid: begin, separate, end.")
        };
    }

    private static EnumValue<BreakValues> ParseBreakType(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "page" => BreakValues.Page,
            "column" => BreakValues.Column,
            "textwrapping" or "line" => BreakValues.TextWrapping,
            _ => throw new ArgumentException(
                $"Invalid break type '{s}'. Valid: page, column, line, textwrapping.")
        };
    }

    private static string? GetRunFont(Run run)
    {
        var fonts = run.RunProperties?.RunFonts;
        return fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value ?? fonts?.EastAsia?.Value;
    }

    private static string? GetRunFontSize(Run run)
    {
        var size = run.RunProperties?.FontSize?.Val?.Value;
        if (size == null) return null;
        return $"{int.Parse(size) / 2.0:0.##}pt"; // stored as half-points
    }

    private string GetRunFormatDescription(Run run, Paragraph? para = null)
    {
        var parts = new List<string>();

        RunProperties? rProps;
        if (para != null)
        {
            rProps = ResolveEffectiveRunProperties(run, para);
        }
        else
        {
            rProps = run.RunProperties;
        }
        if (rProps == null) return "(default)";

        var font = GetFontFromProperties(rProps);
        if (font != null) parts.Add(font);

        var size = GetSizeFromProperties(rProps);
        if (size != null) parts.Add(size);

        if (rProps.Bold != null) parts.Add("bold");
        if (rProps.Italic != null) parts.Add("italic");
        if (rProps.Underline != null) parts.Add("underline");
        if (rProps.Strike != null) parts.Add("strikethrough");

        return parts.Count > 0 ? string.Join(" ", parts) : "(default)";
    }

    private static RunProperties EnsureRunProperties(Run run)
    {
        return run.RunProperties ?? run.PrependChild(new RunProperties());
    }

    /// <summary>
    /// Return the paragraph-mark run properties, creating them IN CT_PPr SCHEMA
    /// ORDER when absent. The strongly-typed <see cref="ParagraphProperties.ParagraphMarkRunProperties"/>
    /// setter inserts the element at its schema position — crucially BEFORE any
    /// <c>&lt;w:sectPr&gt;</c> (rPr precedes sectPr in CT_PPr). Plain
    /// <c>AppendChild(new ParagraphMarkRunProperties())</c> appends at the end,
    /// which lands the rPr AFTER an already-present sectPr and fails schema
    /// validation — exactly the section-break-paragraph case (BUG-DUMP-R37-1),
    /// where the rebuilt section paragraph already holds its sectPr when Set
    /// re-applies the paragraph-mark formatting.
    /// </summary>
    private static ParagraphMarkRunProperties EnsureMarkRunProperties(ParagraphProperties pProps)
    {
        return pProps.ParagraphMarkRunProperties
            ??= new ParagraphMarkRunProperties();
    }

    /// <summary>
    /// Parse a w:shd value string ("fill", "val;fill", "val;fill;color") into a Shading element.
    /// Shared by paragraph-level, run-level, and pmrp shading handlers.
    /// </summary>
    // OOXML ST_Shd enum values. Kept as a static set so ParseShadingValue can
    // reject anything else before the SDK's opaque EnumValue ctor leaks at
    // save time.
    private static readonly HashSet<string> s_knownShadingPatternValues = new(StringComparer.Ordinal)
    {
        "clear", "solid", "nil",
        "pct5", "pct10", "pct12", "pct15", "pct20", "pct25", "pct30", "pct35",
        "pct37", "pct40", "pct45", "pct50", "pct55", "pct60", "pct62", "pct65",
        "pct70", "pct75", "pct80", "pct85", "pct87", "pct90", "pct95",
        "diagStripe", "reverseDiagStripe", "thinDiagStripe", "thinReverseDiagStripe",
        "horzStripe", "vertStripe", "thinHorzStripe", "thinVertStripe",
        "diagCross", "thinDiagCross", "horzCross", "thinHorzCross",
    };
    private static bool IsKnownShadingPatternValue(string v) => s_knownShadingPatternValues.Contains(v);

    private static Shading ParseShadingValue(string value)
    {
        // BUG-DUMP-R41-4: peel off any theme key=val tail
        // (themeFill=…/themeFillShade=…/themeFillTint=… and
        // themeColor/themeShade/themeTint) before positional VAL;FILL;COLOR
        // parsing, then re-stamp the theme attrs via ApplyShadingTheme. A
        // non-themed shading string contains no '=' and is parsed verbatim.
        var (posValue, shdTheme) = ExtractThemeTail(value);
        value = posValue;
        var shdParts = value.Split(';');
        var shd = new Shading();
        if (shdParts.Length == 1)
        {
            shd.Val = ShadingPatternValues.Clear;
            shd.Fill = SanitizeHex(shdParts[0]);
        }
        else if (shdParts.Length >= 2)
        {
            var firstAsHex = shdParts[0].TrimStart('#');
            if (firstAsHex.Length >= 6 && firstAsHex.All(char.IsAsciiHexDigit))
            {
                shd.Val = ShadingPatternValues.Clear;
                shd.Fill = SanitizeHex(shdParts[0]);
            }
            else
            {
                WarnIfShadingOrderWrong(shdParts[0]);
                // OOXML ST_Shd enum: clear/solid/nil + many pct*/stripe/cross
                // variants. SDK's `new ShadingPatternValues(raw)` happily
                // accepts any string and throws an opaque internal exception
                // at serialize time. Validate up-front so callers see a
                // friendly ArgumentException.
                if (!IsKnownShadingPatternValue(shdParts[0]))
                    throw new ArgumentException(
                        $"Invalid 'shading' pattern value: '{shdParts[0]}'. Valid ST_Shd values: clear, solid, nil, pct5/10/12/15/20/25/30/35/37/40/45/50/55/60/62/65/70/75/80/85/87/90/95, diagStripe, reverseDiagStripe, thinDiagStripe, thinReverseDiagStripe, horzStripe, vertStripe, thinHorzStripe, thinVertStripe, diagCross, thinDiagCross, horzCross, thinHorzCross.");
                shd.Val = new ShadingPatternValues(shdParts[0]);
                shd.Fill = SanitizeHex(shdParts[1]);
                if (shdParts.Length >= 3 && !string.IsNullOrEmpty(shdParts[2])) shd.Color = SanitizeHex(shdParts[2]);
            }
        }
        // BUG-DUMP-R41-4: stamp the theme-linkage attrs harvested above.
        ApplyShadingTheme(shd, shdTheme);
        return shd;
    }

    /// <summary>
    /// CONSISTENCY(format-inherit): drop one explicit run property so it falls
    /// back to style / docDefaults inheritance (issue #274). Returns true when the
    /// key has a removal mapping — the caller must then skip its own parse of the
    /// token — and false when the key has no mapping, so the caller parses
    /// normally.
    ///
    /// A key that carries no explicit value is a no-op, never an error, and only
    /// the rPr child is touched: the run itself (and any w:ins / w:del wrapper
    /// around it) is left exactly as it was.
    ///
    /// Deliberately excluded, because the token is a legal value for that key
    /// (verified against the released binary):
    ///   * highlight=none — <c>&lt;w:highlight w:val="none"/&gt;</c> is an
    ///     explicit "no highlight", which is the opposite of inheriting one;
    ///   * underline=none — <c>&lt;w:u w:val="none"/&gt;</c> likewise;
    ///   * color=none — writes <c>w:val="auto" w:themeColor="none"</c>;
    ///   * the font.* slots — an empty value already clears those slots, so
    ///     <c>font=None</c> keeps meaning "a typeface named None";
    ///   * bdr — its STYLE;SIZE;COLOR parser already treats none/false/empty as
    ///     a wipe.
    ///
    /// The on/off toggles are mapped too, where the removal spelling is new
    /// behavior rather than a restatement: <c>bold=false</c> keeps writing
    /// <c>&lt;w:b w:val="0"/&gt;</c> (explicit off) and <c>bold=clear</c> used to
    /// fail with "Invalid boolean value" — it now removes the element, which is
    /// what lets a style-level bold show through again.
    ///
    /// Scope: this only applies where the key is already routed to
    /// ApplyRunFormatting. A target that owns the key in its own cascade (e.g.
    /// paragraph-level <c>rtl</c>, which the paragraph path sends through the
    /// direction cascade) keeps its current handling.
    /// </summary>
    private static bool TryRemoveRunFormatting(OpenXmlCompositeElement props, string key, string value)
    {
        var token = value.Trim().ToLowerInvariant();
        switch (key.ToLowerInvariant())
        {
            case "size" or "fontsize" or "font.size":
                // <w:sz w:val="half-points"/> — the explicit size that made a
                // redefined Normal style invisible in #274.
                props.RemoveAllChildren<FontSize>();
                return true;
            case "size.cs" or "font.size.cs" or "sizecs":
                props.RemoveAllChildren<FontSizeComplexScript>();
                return true;
            case "color" or "font.color":
                if (token == "none") return false;   // themeColor="none" is legal
                props.RemoveAllChildren<Color>();
                return true;
            case "highlight":
                if (token == "none") return false;   // <w:highlight w:val="none"/>
                props.RemoveAllChildren<Highlight>();
                return true;
            case "underline" or "font.underline":
                if (token == "none") return false;   // <w:u w:val="none"/>
                props.RemoveAllChildren<Underline>();
                return true;
            case "bold" or "font.bold":
                props.RemoveAllChildren<Bold>();
                return true;
            case "bold.cs" or "font.bold.cs" or "boldcs":
                props.RemoveAllChildren<BoldComplexScript>();
                return true;
            case "italic" or "font.italic":
                props.RemoveAllChildren<Italic>();
                return true;
            case "italic.cs" or "font.italic.cs" or "italiccs":
                props.RemoveAllChildren<ItalicComplexScript>();
                return true;
            case "strike" or "strikethrough" or "font.strike" or "font.strikethrough":
                props.RemoveAllChildren<Strike>();
                return true;
            case "dstrike":
                props.RemoveAllChildren<DoubleStrike>();
                return true;
            case "caps" or "allcaps":
                props.RemoveAllChildren<Caps>();
                return true;
            case "smallcaps":
                props.RemoveAllChildren<SmallCaps>();
                return true;
            case "vanish":
                props.RemoveAllChildren<Vanish>();
                return true;
            case "specvanish":
                props.RemoveAllChildren<SpecVanish>();
                return true;
            case "outline":
                props.RemoveAllChildren<Outline>();
                return true;
            case "shadow":
                props.RemoveAllChildren<Shadow>();
                return true;
            case "emboss":
                props.RemoveAllChildren<Emboss>();
                return true;
            case "imprint":
                props.RemoveAllChildren<Imprint>();
                return true;
            case "noproof":
                props.RemoveAllChildren<NoProof>();
                return true;
            case "rtl":
            case "direction" or "dir":
                // rtl=false / direction=ltr keep writing the explicit val="0"
                // override; only the removal spellings drop the element.
                props.RemoveAllChildren<RightToLeftText>();
                return true;
            case "superscript":
            case "subscript":
            case "vertalign":
                props.RemoveAllChildren<VerticalTextAlignment>();
                return true;
            case "charspacing" or "letterspacing" or "spacing":
                props.RemoveAllChildren<Spacing>();
                return true;
            case "w" or "charscale":
                props.RemoveAllChildren<CharacterScale>();
                return true;
            case "kern":
                props.RemoveAllChildren<Kern>();
                return true;
            case "position":
                props.RemoveAllChildren<Position>();
                return true;
            case "shading" or "shd" or "fill":
                if (token == "clear") return false;  // ST_Shd pattern is a value
                props.RemoveAllChildren<Shading>();
                return true;
            case "rstyle":
                // <w:rStyle w:val="..."/> — drop the character-style binding so
                // the paragraph style's character formatting applies again.
                props.RemoveAllChildren<RunStyle>();
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Apply a run-level (rPr-style) property to any container that holds rPr children:
    /// <c>RunProperties</c>, <c>ParagraphMarkRunProperties</c>, or <c>StyleRunProperties</c>.
    /// Uses <see cref="OpenXmlCompositeElement"/> + RemoveAllChildren+InsertRunPropInSchemaOrder
    /// so the same logic works across all three despite their different typed property surfaces.
    /// Returns true if the key was handled, false if caller should fall through.
    /// </summary>
    private static bool ApplyRunFormatting(OpenXmlCompositeElement props, string key, string? value)
    {
        if (value is null) return false;
        // CONSISTENCY(format-inherit): an explicit removal spelling drops the
        // rPr child that carries the setting instead of being parsed as a value,
        // so the run inherits it from the style / docDefaults again (issue #274).
        if (IsFormatRemovalToken(value) && TryRemoveRunFormatting(props, key, value)) return true;
        switch (key.ToLowerInvariant())
        {
            case "size":
            case "fontsize":
            case "font.size":
                var existingFs = props.GetFirstChild<FontSize>();
                if (existingFs != null) existingFs.Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString();
                else InsertRunPropInSchemaOrder(props, new FontSize { Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString() });
                return true;
            case "font":
            case "font.name":
                // Bare 'font' targets ASCII+HighAnsi+EastAsia. Use 'font.latin',
                // 'font.ea', 'font.cs' for per-script control (e.g. Japanese,
                // Korean, Arabic — the CS slot owns Arabic/Hebrew typefaces).
                {
                    var fv = SanitizeFontTokenInput(value);
                    var existingRf = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(fv))
                    {
                        if (existingRf != null)
                        {
                            existingRf.Ascii = null; existingRf.HighAnsi = null; existingRf.EastAsia = null;
                            if (RunFontsIsEmpty(existingRf)) existingRf.Remove();
                        }
                    }
                    else if (existingRf != null) { existingRf.Ascii = fv; existingRf.HighAnsi = fv; existingRf.EastAsia = fv; }
                    else InsertRunPropInSchemaOrder(props, new RunFonts { Ascii = fv, HighAnsi = fv, EastAsia = fv });
                }
                return true;
            case "font.latin":
                {
                    var fv = SanitizeFontTokenInput(value);
                    var rfLatin = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(fv))
                    {
                        if (rfLatin != null)
                        {
                            rfLatin.Ascii = null; rfLatin.HighAnsi = null;
                            if (RunFontsIsEmpty(rfLatin)) rfLatin.Remove();
                        }
                    }
                    else if (rfLatin != null) { rfLatin.Ascii = fv; rfLatin.HighAnsi = fv; }
                    else InsertRunPropInSchemaOrder(props, new RunFonts { Ascii = fv, HighAnsi = fv });
                }
                return true;
            // CONSISTENCY(font-slot-asymmetric): Navigation surfaces ascii and
            // hAnsi separately when their values disagree (font.ascii vs
            // font.hAnsi); the bare `font.latin` shorthand only handles the
            // symmetric case. Without these, sources with asymmetric Latin
            // fonts (e.g. ascii=黑体 hAnsi=宋体, common in CJK docs to fork
            // ASCII vs extended-Latin glyphs) round-trip with the keys lost.
            case "font.ascii":
                {
                    var fv = SanitizeFontTokenInput(value);
                    var rfA = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(fv))
                    {
                        if (rfA != null) { rfA.Ascii = null; if (RunFontsIsEmpty(rfA)) rfA.Remove(); }
                    }
                    else if (rfA != null) rfA.Ascii = fv;
                    else InsertRunPropInSchemaOrder(props, new RunFonts { Ascii = fv });
                }
                return true;
            case "font.hansi" or "font.hAnsi":
                {
                    var fv = SanitizeFontTokenInput(value);
                    var rfHA = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(fv))
                    {
                        if (rfHA != null) { rfHA.HighAnsi = null; if (RunFontsIsEmpty(rfHA)) rfHA.Remove(); }
                    }
                    else if (rfHA != null) rfHA.HighAnsi = fv;
                    else InsertRunPropInSchemaOrder(props, new RunFonts { HighAnsi = fv });
                }
                return true;
            case "font.ea" or "font.eastasia" or "font.eastasian":
                {
                    var fv = SanitizeFontTokenInput(value);
                    var rfEa = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(fv))
                    {
                        if (rfEa != null)
                        {
                            rfEa.EastAsia = null;
                            if (RunFontsIsEmpty(rfEa)) rfEa.Remove();
                        }
                    }
                    else if (rfEa != null) { rfEa.EastAsia = fv; }
                    else InsertRunPropInSchemaOrder(props, new RunFonts { EastAsia = fv });
                }
                return true;
            // <w:rFonts w:hint="eastAsia|cs|default"/> — selects which font slot
            // renders ambiguous characters. Curated (not a font typeface slot)
            // so it round-trips alongside font.ea/font.latin. Dropping it
            // rewraps tight CJK lines (the renderer picks a different-width font).
            case "font.hint":
                {
                    var rfHint = props.GetFirstChild<RunFonts>();
                    var hv = (value ?? "").Trim().ToLowerInvariant();
                    FontTypeHintValues? hint = hv switch
                    {
                        "eastasia" => FontTypeHintValues.EastAsia,
                        "cs" or "complexscript" or "complex" => FontTypeHintValues.ComplexScript,
                        "default" => FontTypeHintValues.Default,
                        _ => (FontTypeHintValues?)null,
                    };
                    if (!hint.HasValue)
                    {
                        if (rfHint != null) { rfHint.Hint = null; if (RunFontsIsEmpty(rfHint)) rfHint.Remove(); }
                    }
                    else if (rfHint != null) rfHint.Hint = hint.Value;
                    else InsertRunPropInSchemaOrder(props, new RunFonts { Hint = hint.Value });
                }
                return true;
            // CONSISTENCY(font-theme-slot): theme-font slots bind to a theme
            // major/minor face instead of a literal typeface. Mirrors the
            // text-run additions in AddRun/AddParagraph but routed through
            // ApplyRunFormatting so Set paragraph (and any other call site
            // that funnels through this helper) honours them too.
            case "font.asciitheme" or "font.asciiTheme":
                {
                    var rfAT = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(value))
                    {
                        if (rfAT != null) { rfAT.AsciiTheme = null; if (RunFontsIsEmpty(rfAT)) rfAT.Remove(); }
                    }
                    else
                    {
                        var enumAT = new EnumValue<ThemeFontValues>(new ThemeFontValues(value));
                        if (rfAT != null) rfAT.AsciiTheme = enumAT;
                        else InsertRunPropInSchemaOrder(props, new RunFonts { AsciiTheme = enumAT });
                    }
                }
                return true;
            case "font.hansitheme" or "font.hAnsiTheme":
                {
                    var rfHAT = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(value))
                    {
                        if (rfHAT != null) { rfHAT.HighAnsiTheme = null; if (RunFontsIsEmpty(rfHAT)) rfHAT.Remove(); }
                    }
                    else
                    {
                        var enumHAT = new EnumValue<ThemeFontValues>(new ThemeFontValues(value));
                        if (rfHAT != null) rfHAT.HighAnsiTheme = enumHAT;
                        else InsertRunPropInSchemaOrder(props, new RunFonts { HighAnsiTheme = enumHAT });
                    }
                }
                return true;
            case "font.eatheme" or "font.eaTheme" or "font.eastasiatheme":
                {
                    var rfEAT = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(value))
                    {
                        if (rfEAT != null) { rfEAT.EastAsiaTheme = null; if (RunFontsIsEmpty(rfEAT)) rfEAT.Remove(); }
                    }
                    else
                    {
                        var enumEAT = new EnumValue<ThemeFontValues>(new ThemeFontValues(value));
                        if (rfEAT != null) rfEAT.EastAsiaTheme = enumEAT;
                        else InsertRunPropInSchemaOrder(props, new RunFonts { EastAsiaTheme = enumEAT });
                    }
                }
                return true;
            case "font.cstheme" or "font.csTheme":
                {
                    var rfCST = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(value))
                    {
                        if (rfCST != null) { rfCST.ComplexScriptTheme = null; if (RunFontsIsEmpty(rfCST)) rfCST.Remove(); }
                    }
                    else
                    {
                        var enumCST = new EnumValue<ThemeFontValues>(new ThemeFontValues(value));
                        if (rfCST != null) rfCST.ComplexScriptTheme = enumCST;
                        else InsertRunPropInSchemaOrder(props, new RunFonts { ComplexScriptTheme = enumCST });
                    }
                }
                return true;
            case "font.cs" or "font.complexscript" or "font.complex":
                {
                    var fv = SanitizeFontTokenInput(value);
                    var rfCs = props.GetFirstChild<RunFonts>();
                    if (string.IsNullOrEmpty(fv))
                    {
                        // CONSISTENCY(empty-clears): empty value clears the
                        // attribute, mirroring direction=. Stub <w:rFonts cs=""/>
                        // is invalid OOXML and confuses Get readback.
                        if (rfCs != null)
                        {
                            rfCs.ComplexScript = null;
                            if (RunFontsIsEmpty(rfCs)) rfCs.Remove();
                        }
                    }
                    else if (rfCs != null) { rfCs.ComplexScript = fv; }
                    else InsertRunPropInSchemaOrder(props, new RunFonts { ComplexScript = fv });
                }
                return true;
            case "bold":
            case "font.bold":
                props.RemoveAllChildren<Bold>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Bold { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Bold());
                return true;
            case "bold.cs" or "font.bold.cs" or "boldcs":
                props.RemoveAllChildren<BoldComplexScript>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new BoldComplexScript { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new BoldComplexScript());
                return true;
            case "italic":
            case "font.italic":
                props.RemoveAllChildren<Italic>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Italic { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Italic());
                return true;
            case "italic.cs" or "font.italic.cs" or "italiccs":
                props.RemoveAllChildren<ItalicComplexScript>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new ItalicComplexScript { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new ItalicComplexScript());
                return true;
            case "size.cs" or "font.size.cs" or "sizecs":
                // Complex-script font size (<w:szCs/>, half-points). When set,
                // Arabic / Hebrew renders at this size; <w:sz/> only affects
                // Latin runs. Bare 'size' continues to write <w:sz/> only —
                // see CONSISTENCY(cs-explicit) in the bare-size case above.
                props.RemoveAllChildren<FontSizeComplexScript>();
                InsertRunPropInSchemaOrder(props, new FontSizeComplexScript { Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString() });
                return true;
            case "color":
            case "font.color":
                props.RemoveAllChildren<Color>();
                // Scheme colors (e.g. accent1, dark2, hyperlink) write to the
                // ThemeColor attribute instead of Val; Val is left at "auto"
                // per ECMA-376 §17.3.2.6 (Excel rejects Val=accent1).
                {
                    // BUG-DUMP-R44-1: split the ';themeColor=…' tail first so a
                    // direct run color carrying both an explicit hex and a theme
                    // linkage (e.g. "FFFFFF;themeColor=background1") keeps the hex
                    // as w:val AND stamps the theme attrs — instead of collapsing
                    // to val="auto" (invisible). Mirrors the style-color path.
                    var (posValue, colorTheme) = ExtractThemeTail(value);
                    Color colorEl;
                    // Bare "auto" is a legal Color val per ECMA-376 §17.3.2.6 —
                    // it tells Word to use the document's automatic text color.
                    // SchemeColorNames includes "auto" for the cross-handler
                    // input lenience pass, but new ThemeColorValues("auto")
                    // throws (no such enum). Short-circuit before the scheme
                    // branch so dump-emitted color=auto round-trips correctly.
                    if (string.Equals(posValue, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        colorEl = new Color { Val = "auto" };
                    }
                    else if (posValue.Length == 0 && colorTheme.Count > 0)
                    {
                        // Theme-only linkage with no positional hex — Val=auto.
                        colorEl = new Color { Val = "auto" };
                    }
                    else
                    {
                        var schemeName = OfficeCli.Core.ParseHelpers.NormalizeSchemeColorName(posValue);
                        if (schemeName != null)
                        {
                            colorEl = new Color { Val = "auto", ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(schemeName)) };
                        }
                        else
                        {
                            colorEl = new Color { Val = SanitizeHex(posValue) };
                        }
                    }
                    // Stamp themeColor/themeShade/themeTint from the tail (no-op
                    // when the value carried none). When a scheme name already
                    // set ThemeColor above, an explicit tail themeColor wins.
                    ApplyColorTheme(colorEl, colorTheme);
                    InsertRunPropInSchemaOrder(props, colorEl);
                }
                return true;
            case "highlight":
                props.RemoveAllChildren<Highlight>();
                InsertRunPropInSchemaOrder(props, new Highlight { Val = ParseHighlightColor(value) });
                return true;
            case "underline":
            case "font.underline":
            {
                // CONSISTENCY(underline-color-preserve): snapshot any existing
                // <w:u w:color="…"/> attribute before rebuilding the element,
                // so toggling the style ("single" → "double") does not silently
                // drop a previously-set underline colour. The dedicated
                // "underline.color" case rebuilds the Underline element from
                // scratch and would otherwise be the only path that keeps
                // colour through a style change.
                var existingUl = props.GetFirstChild<Underline>();
                var preservedColor = existingUl?.Color?.Value;
                var preservedThemeColor = existingUl?.ThemeColor?.Value;
                var preservedThemeTint = existingUl?.ThemeTint?.Value;
                var preservedThemeShade = existingUl?.ThemeShade?.Value;
                props.RemoveAllChildren<Underline>();
                var ulMapped = NormalizeUnderlineValue(value);
                var newUl = new Underline { Val = new UnderlineValues(ulMapped) };
                if (preservedColor != null) newUl.Color = preservedColor;
                if (preservedThemeColor != null) newUl.ThemeColor = preservedThemeColor;
                if (preservedThemeTint != null) newUl.ThemeTint = preservedThemeTint;
                if (preservedThemeShade != null) newUl.ThemeShade = preservedThemeShade;
                InsertRunPropInSchemaOrder(props, newUl);
                return true;
            }
            case "underline.color":
            case "underlinecolor":
            case "underlineColor":
            case "font.underline.color":
            {
                // CONSISTENCY(underline-color): Get emits canonical
                // 'underline.color' (see Navigation.cs L1815 etc.). Set
                // accepts dotted form plus camelCase aliases. The OOXML
                // shape is <w:u w:val="…" w:color="RRGGBB"/> — color is an
                // attribute on the existing Underline element, not a child
                // element. Preserve any existing val — including its absence:
                // <w:u w:color="…"/> with no w:val is a legal source shape
                // that renders NO underline, and defaulting it to single here
                // underlined entire documents on dump→batch replay. Callers
                // that want a visible underline pass `underline=` alongside.
                var existingUl = props.GetFirstChild<Underline>();
                var ulVal = existingUl?.Val;
                props.RemoveAllChildren<Underline>();
                var hex = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(value).Rgb;
                var ulEl = new Underline { Color = hex };
                if (ulVal != null) ulEl.Val = ulVal;
                InsertRunPropInSchemaOrder(props, ulEl);
                return true;
            }
            case "strike" or "strikethrough" or "font.strike" or "font.strikethrough":
                props.RemoveAllChildren<Strike>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Strike { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Strike());
                return true;
            case "dstrike":
                props.RemoveAllChildren<DoubleStrike>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new DoubleStrike { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new DoubleStrike());
                return true;
            case "outline":
                props.RemoveAllChildren<Outline>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Outline { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Outline());
                return true;
            case "shadow":
                props.RemoveAllChildren<Shadow>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Shadow { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Shadow());
                return true;
            case "emboss":
                props.RemoveAllChildren<Emboss>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Emboss { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Emboss());
                return true;
            case "imprint":
                props.RemoveAllChildren<Imprint>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Imprint { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Imprint());
                return true;
            case "noproof":
                props.RemoveAllChildren<NoProof>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new NoProof { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new NoProof());
                return true;
            case "rtl":
            case "direction" or "dir":
                // 'direction=rtl|ltr' is the canonical key (mirrors paragraph
                // and PPT); 'rtl=true|false' kept as legacy boolean alias.
                props.RemoveAllChildren<RightToLeftText>();
                bool isLegacyRtlKey = key.ToLowerInvariant() == "rtl";
                bool rtlOn = isLegacyRtlKey
                    ? IsTruthy(value)
                    : value.ToLowerInvariant() switch
                    {
                        "rtl" or "righttoleft" or "right-to-left" or "true" or "1" => true,
                        "ltr" or "lefttoright" or "left-to-right" or "false" or "0" or "" => false,
                        _ => throw new ArgumentException($"Invalid direction value: '{value}'. Valid values: rtl, ltr.")
                    };
                if (rtlOn)
                {
                    InsertRunPropInSchemaOrder(props, new RightToLeftText());
                }
                else
                {
                    // Both legacy 'rtl=false' and canonical 'direction=ltr'
                    // emit <w:rtl w:val="0"/> as an explicit override of an
                    // inherited RTL paragraph / style / docDefaults — without
                    // it the LTR override has no render-time effect.
                    InsertRunPropInSchemaOrder(props, new RightToLeftText { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) });
                }
                return true;
            case "charspacing" or "letterspacing" or "spacing":
                // w:spacing native unit is twips. With `pt` suffix convert to
                // twips; bare number is taken as twips so dump→batch round-trips
                // and `Get`'s "<n>pt" readback (twips/20) stays consistent.
                int csTwips = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? (int)Math.Round(ParseHelpers.SafeParseDouble(value[..^2], "charspacing") * 20, MidpointRounding.AwayFromZero)
                    : (int)Math.Round(ParseHelpers.SafeParseDouble(value, "charspacing"), MidpointRounding.AwayFromZero);
                props.RemoveAllChildren<Spacing>();
                InsertRunPropInSchemaOrder(props, new Spacing { Val = csTwips });
                return true;
            case "shading" or "shd" or "fill":
                // CONSISTENCY(shd-canonical-fill): `fill` is the canonical key
                // Get emits for a solid run <w:shd>; accept it as a Set/Add alias
                // alongside the legacy shading/shd so the dump→batch round-trip
                // (which now carries `fill`) replays. ParseShadingValue treats a
                // bare #RRGGBB as <w:shd w:val="clear" w:fill="…"/>.
                props.RemoveAllChildren<Shading>();
                InsertRunPropInSchemaOrder(props, ParseShadingValue(value));
                return true;
            case "rstyle":
                // <w:rStyle w:val="…"/> — character style binding. Needed by
                // the markRPr.* dispatch (markRPr.rStyle) and generic callers;
                // run/hyperlink Adds keep their own explicit handling.
                props.RemoveAllChildren<RunStyle>();
                if (!string.IsNullOrEmpty(value))
                    InsertRunPropInSchemaOrder(props, new RunStyle { Val = value });
                return true;
            case "w" or "charscale":
                // <w:w w:val="…"/> — horizontal character scale percentage.
                props.RemoveAllChildren<CharacterScale>();
                InsertRunPropInSchemaOrder(props, new CharacterScale
                {
                    Val = (int)Math.Round(ParseHelpers.SafeParseDouble(value.TrimEnd('%'), "charscale"), MidpointRounding.AwayFromZero),
                });
                return true;
            case "superscript":
                props.RemoveAllChildren<VerticalTextAlignment>();
                if (IsTruthy(value))
                    InsertRunPropInSchemaOrder(props, new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
                return true;
            case "subscript":
                props.RemoveAllChildren<VerticalTextAlignment>();
                if (IsTruthy(value))
                    InsertRunPropInSchemaOrder(props, new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
                return true;
            case "vertalign":
                // Canonical run-level vertAlign Set. Without this case the
                // generic TypedAttributeFallback silently accepted invalid
                // values (banana → SDK EnumValue null → Word silently treats
                // as baseline). Mirror AddRun's vocabulary.
                props.RemoveAllChildren<VerticalTextAlignment>();
                InsertRunPropInSchemaOrder(props, new VerticalTextAlignment
                {
                    Val = value.ToLowerInvariant() switch
                    {
                        "superscript" or "super" => VerticalPositionValues.Superscript,
                        "subscript" or "sub" => VerticalPositionValues.Subscript,
                        "baseline" => VerticalPositionValues.Baseline,
                        _ => throw new ArgumentException($"Invalid 'vertAlign' value: '{value}'. Valid values: superscript, subscript, baseline."),
                    }
                });
                return true;
            case "caps":
            case "allcaps":
                props.RemoveAllChildren<Caps>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Caps { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Caps());
                return true;
            case "smallcaps":
                props.RemoveAllChildren<SmallCaps>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new SmallCaps { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new SmallCaps());
                return true;
            case "vanish":
                props.RemoveAllChildren<Vanish>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new Vanish { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Vanish());
                return true;
            case "specvanish":
                // <w:specVanish/> — Word's style-separator marker (e.g. the
                // hidden ¶ that lets a heading share a line with body text).
                props.RemoveAllChildren<SpecVanish>();
                if (IsExplicitFalseAddOverride(value))
                    InsertRunPropInSchemaOrder(props, new SpecVanish { Val = OnOffValue.FromBoolean(false) });
                else if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new SpecVanish());
                return true;
            case "bdr":
                // BUG-R7-06: character border <w:bdr/> — round-trip captured
                // it from real docs but Add/Set rejected it as UNSUPPORTED.
                // Accept the same colon-encoded form as paragraph borders
                // (STYLE[;SIZE[;COLOR[;SPACE]]]). Empty/none/false clears.
                props.RemoveAllChildren<Border>();
                if (!string.IsNullOrEmpty(value)
                    && !string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
                {
                    var (bStyle, bSize, bColor, bSpace) = ParseBorderValue(value);
                    var bdr = new Border { Val = bStyle, Size = bSize };
                    if (bSpace.HasValue) bdr.Space = bSpace.Value;
                    if (bColor != null) bdr.Color = bColor;
                    // BUG-DUMP-R36-1: round-trip w:shadow / w:frame (segments 5/6
                    // of STYLE;SIZE;COLOR;SPACE;SHADOW;FRAME); stamp only when the
                    // source carried them so a plain border stays plain.
                    var (bShadow, bFrame) = ParseBorderShadowFrame(value);
                    if (bShadow.HasValue) bdr.Shadow = bShadow.Value;
                    if (bFrame.HasValue) bdr.Frame = bFrame.Value;
                    InsertRunPropInSchemaOrder(props, bdr);
                }
                return true;
            case "kern":
                // BUG-R7-06: <w:kern w:val="N"/> (kerning threshold in
                // half-points). Get exposes it; Add/Set silently dropped.
                // Docx kern unit is half-points (raw uint per ST_HpsMeasure);
                // unlike pptx kern (100ths of pt) we deliberately do NOT
                // accept a "pt" suffix here — pass an integer half-points
                // value. An empty value clears the element; an invalid
                // value (e.g. "14pt", "abc") returns false so the dispatch
                // surfaces invalid_value rather than silently no-op'ing.
                props.RemoveAllChildren<Kern>();
                if (string.IsNullOrEmpty(value))
                    return true;
                if (!uint.TryParse(value, out var kernVal))
                    throw new ArgumentException(
                        $"Invalid kern value '{value}'. Pass an integer in half-points (e.g. 28 = 14pt threshold); 'pt' suffix is not accepted on docx kern.");
                // OOXML ST_HpsMeasure MaxInclusive=3277 half-points (~164pt).
                if (kernVal > 3277)
                    throw new ArgumentException(
                        $"Invalid kern value '{value}'. Must be 0-3277 (OOXML ST_HpsMeasure MaxInclusive=3277 half-points = ~164pt).");
                InsertRunPropInSchemaOrder(props, new Kern { Val = kernVal });
                return true;
            case "position":
                // <w:position w:val="N"/> — vertical raise/lower in
                // half-points (ST_SignedHpsMeasure). Positive = raise,
                // negative = lower, 0 = baseline. Distinct from
                // vertAlign=super|sub which is the typographic toggle
                // (renders at ~58% size); position keeps full size and
                // shifts by exact half-points. Doc/x .doc carries this
                // as sprmCHpsPos (0x4845). Get exposes "position" already
                // (see WordHandler.Navigation.cs); Add/Set previously
                // routed to tab-stop SetElementTabStop only.
                props.RemoveAllChildren<Position>();
                if (!string.IsNullOrEmpty(value))
                {
                    int posVal = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                        ? (int)Math.Round(ParseHelpers.SafeParseDouble(value[..^2], "position") * 2, MidpointRounding.AwayFromZero)
                        : (int)Math.Round(ParseHelpers.SafeParseDouble(value, "position"), MidpointRounding.AwayFromZero);
                    InsertRunPropInSchemaOrder(props, new Position { Val = posVal.ToString(System.Globalization.CultureInfo.InvariantCulture) });
                }
                return true;
            case "lang" or "lang.latin" or "lang.val":
            case "lang.ea" or "lang.eastasia" or "lang.eastasian":
            case "lang.cs" or "lang.complexscript" or "lang.bidi":
            {
                // <w:lang w:val=".." w:eastAsia=".." w:bidi=".."/> — three slots
                // for Latin / EastAsian / ComplexScript scripts. Mirrors the
                // font.latin/font.ea/font.cs vocabulary.
                // CONSISTENCY(bcp47-validation): match the PPTX shape lang
                // validator (Bcp47Shape) — reject malformed tags up front
                // rather than writing them into <w:lang> and producing an
                // unloadable document. Empty value clears the slot (and
                // removes the <w:lang> element if all three slots end up
                // empty); literal "null" is rejected as a stray sentinel.
                bool clearSlot = string.IsNullOrEmpty(value);
                if (!clearSlot)
                {
                    if (string.Equals(value, "null", StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException($"Invalid BCP-47 language tag for {key}: '{value}'. Expected a tag like 'en-US', 'ja-JP', or 'ar-SA'.");
                    if (!LangBcp47IsValid(value))
                        throw new ArgumentException($"Invalid BCP-47 language tag for {key}: '{value}'. Expected a tag like 'en-US', 'ja-JP', or 'ar-SA' (RFC 5646: <= {LangBcp47MaxLength} chars, primary subtag 2-3 letters, then hyphen-separated subtags).");
                }
                var lang = props.GetFirstChild<Languages>();
                if (lang == null)
                {
                    if (clearSlot) return true;
                    lang = new Languages();
                    InsertRunPropInSchemaOrder(props, lang);
                }
                switch (key.ToLowerInvariant())
                {
                    case "lang":
                    case "lang.latin":
                    case "lang.val":
                        if (clearSlot) lang.Val = null; else lang.Val = value;
                        break;
                    case "lang.ea":
                    case "lang.eastasia":
                    case "lang.eastasian":
                        if (clearSlot) lang.EastAsia = null; else lang.EastAsia = value;
                        break;
                    case "lang.cs":
                    case "lang.complexscript":
                    case "lang.bidi":
                        if (clearSlot) lang.Bidi = null; else lang.Bidi = value;
                        break;
                }
                // Remove the <w:lang> element entirely when all three slots
                // are empty — leaves no stale empty-attr noise after clears.
                if (lang.Val?.Value is null && lang.EastAsia?.Value is null && lang.Bidi?.Value is null)
                    lang.Remove();
                return true;
            }
            case "snaptogrid":
                // <w:snapToGrid> run/¶-mark toggle (CT_OnOff). On a doc with a
                // <w:docGrid>, snapToGrid="0" tells Word NOT to snap this run (or
                // the paragraph mark) to the line grid — which changes the line's
                // height. The OFF form is the meaningful one to round-trip: a
                // ¶-mark <w:rPr><w:snapToGrid w:val="0"/> set the terminating
                // line's height, and dropping it re-snapped the line to the grid,
                // shifting metrics and reflowing the page. The run-level toggle
                // already round-trips via the generic CT_OnOff reader + this case;
                // it's also reachable through the markRPr.* dispatch (Add.Text.cs)
                // which previously fell to `default: return false` and dropped it.
                props.RemoveAllChildren<SnapToGrid>();
                if (IsTruthy(value))
                    InsertRunPropInSchemaOrder(props, new SnapToGrid());
                else
                    InsertRunPropInSchemaOrder(props, new SnapToGrid { Val = OnOffValue.FromBoolean(false) });
                return true;
            // BUG-DUMP-RPR-CONTAINER: <w:em> (East-Asian emphasis mark — the dots/
            // circles painted above/below CJK glyphs). Previously had NO curated
            // case, so it only round-tripped on the plain-run path via AddRun's
            // generic schema-reflection fallback; runs in a hyperlink / field-result /
            // footnote (which apply rPr through this curated applier) silently lost it.
            case "em":
            case "emphasismark":
            case "emphasis":
            {
                props.RemoveAllChildren<Emphasis>();
                EmphasisMarkValues? emv = value?.Trim().ToLowerInvariant() switch
                {
                    "dot" => EmphasisMarkValues.Dot,
                    "comma" => EmphasisMarkValues.Comma,
                    "circle" => EmphasisMarkValues.Circle,
                    "underdot" or "under-dot" => EmphasisMarkValues.UnderDot,
                    "none" or "" or null => EmphasisMarkValues.None,
                    _ => null
                };
                if (emv.HasValue)
                    InsertRunPropInSchemaOrder(props, new Emphasis { Val = emv.Value });
                return true;
            }
            default:
                return false;
        }
    }

    /// <summary>
    /// Insert a run property element in the correct CT_RPr schema position.
    /// Delegates to <see cref="Core.SchemaOrder"/>, which reads the SDK's own
    /// compiled-particle order — complete by construction. The previous
    /// hand-maintained type→index switch silently dropped any unlisted type
    /// (kern, w, position, …) to "append at end", producing out-of-order rPr
    /// that strict validators reject.
    /// </summary>
    // BUG-DUMP-R71-RPR-ORDER: re-seat every standard (non-w14) child of a
    // run-property container (CT_RPr / CT_ParaRPr) into its schema slot. The
    // rPr is built across mixed paths — SDK typed setters, ApplyRunFormatting
    // (schema-ordered), but also raw AppendChild (e.g. AddRun's rFonts/sz) and
    // TypedAttributeFallback (appends at the tail) — so a child can land out of
    // CT_RPr order (sz after u, ins after rStyle). Running each standard child
    // back through InsertRunPropInSchemaOrder once at the end converges on the
    // correct order regardless of how it got there; its Place + w14 hoist also
    // keep the w14 extension block last. Content-preserving (only reorders).
    private static void NormalizeRunPropsSchemaOrder(OpenXmlCompositeElement? props)
    {
        if (props == null) return;
        foreach (var child in props.ChildElements.Where(c => c.NamespaceUri != W14Ns).ToList())
        {
            child.Remove();
            InsertRunPropInSchemaOrder(props, child);
        }
    }

    private static void InsertRunPropInSchemaOrder(OpenXmlCompositeElement props, OpenXmlElement elem)
    {
        props.AppendChild(elem);
        Core.SchemaOrder.Place(props, elem);

        // The w14 run-property extensions (glow / shadow / … / textFill) are
        // appended at the very END of rPr, but the SDK particle comparator
        // treats them as unknown and sorts unknowns to the FRONT — so Place
        // judges a just-added STANDARD child (e.g. <w:lang>) to come *after* the
        // w14 block and leaves it stranded behind them, which is schema-invalid
        // (every CT_RPr child must precede the w14 extension block). If elem is a
        // standard element that ended up after a w14 sibling, hoist it before the
        // first w14 child.
        if (elem.NamespaceUri != W14Ns)
        {
            OpenXmlElement? firstW14 = null;
            bool sawElem = false;
            foreach (var c in props.ChildElements)
            {
                if (ReferenceEquals(c, elem)) { sawElem = true; }
                else if (c.NamespaceUri == W14Ns) { firstW14 = c; break; }
            }
            if (firstW14 != null && !sawElem)
            {
                elem.Remove();
                props.InsertBefore(elem, firstW14);
            }
        }
    }
}
