// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private Dictionary<string, string>? _themeColors;
    private Dictionary<string, string>? _themeFonts;

    // OOXML theme font axes: major{Ascii|HAnsi|EastAsia|Bidi} +
    // minor{Ascii|HAnsi|EastAsia|Bidi}. The 8 keys map a w:asciiTheme /
    // w:hAnsiTheme / w:eastAsiaTheme / w:cstheme attribute value (after
    // normalization to one of these enum strings) to the resolved typeface
    // declared in theme1.xml's <a:fontScheme>. asciiTheme and hAnsiTheme
    // both point at the latin face — Word treats them as one slot.
    // Modeled after ThemeHandler::resolveMajorMinorTypeFace.
    private Dictionary<string, string> GetThemeFonts()
    {
        if (_themeFonts != null) return _themeFonts;
        _themeFonts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        DocumentFormat.OpenXml.Drawing.FontScheme? fs = null;
        try { fs = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme; }
        catch (System.Xml.XmlException) { return _themeFonts; }
        if (fs == null) return _themeFonts;

        void Put(string key, string? typeface)
        {
            if (!string.IsNullOrEmpty(typeface)) _themeFonts[key] = typeface;
        }
        if (fs.MajorFont is { } maj)
        {
            Put("majorAscii", maj.LatinFont?.Typeface?.Value);
            Put("majorHAnsi", maj.LatinFont?.Typeface?.Value);
            Put("majorEastAsia", maj.EastAsianFont?.Typeface?.Value);
            Put("majorBidi", maj.ComplexScriptFont?.Typeface?.Value);
        }
        if (fs.MinorFont is { } min)
        {
            Put("minorAscii", min.LatinFont?.Typeface?.Value);
            Put("minorHAnsi", min.LatinFont?.Typeface?.Value);
            Put("minorEastAsia", min.EastAsianFont?.Typeface?.Value);
            Put("minorBidi", min.ComplexScriptFont?.Typeface?.Value);
        }
        return _themeFonts;
    }

    // OOXML theme attribute values are an enum of {majorAscii, majorHAnsi,
    // majorEastAsia, majorBidi, minorAscii, minorHAnsi, minorEastAsia,
    // minorBidi}. Returns null when the theme part is missing or the
    // requested axis isn't declared.
    private string? ResolveThemeFont(string? themeAttr)
    {
        if (string.IsNullOrEmpty(themeAttr)) return null;
        return GetThemeFonts().TryGetValue(themeAttr, out var face) ? face : null;
    }

    // CONSISTENCY(office-default-palette): when the doc has no <a:theme>
    // part, fall back to the canonical Office palette so
    // w:themeColor="accent1" resolves instead of silently dropping.
    private static readonly Dictionary<string, string> _officeDefaultThemeFallback = OfficeDefaultThemeColors.BuildAliasMap();

    private Dictionary<string, string> GetThemeColors()
    {
        if (_themeColors != null) return _themeColors;

        // A malformed theme1.xml (any XML error) throws XmlException on
        // lazy access deep inside the first reader. Fall back to the Office
        // default palette rather than tainting the whole preview. Same
        // approach used for styles/footnotes below.
        DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme = null;
        try { colorScheme = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme; }
        catch (System.Xml.XmlException) { }
        _themeColors = ThemeColorResolver.BuildColorMap(colorScheme, includePptAliases: false);

        // Fill in any missing standard names from the Office default theme so
        // themeColor references resolve even when the docx has no theme part.
        foreach (var (name, hex) in _officeDefaultThemeFallback)
        {
            if (!_themeColors.ContainsKey(name))
                _themeColors[name] = hex;
        }
        return _themeColors;
    }

    private string? ResolveSchemeColor(OpenXmlElement schemeColor)
    {
        var schemeName = schemeColor.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (schemeName == null) return null;

        var themeColors = GetThemeColors();
        if (!themeColors.TryGetValue(schemeName, out var hex)) return null;

        // Extract transform values from child elements
        var tint = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "tint");
        var shade = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "shade");
        var lumMod = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "lumMod");
        var lumOff = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "lumOff");

        var hasTint = tint != null ? (int?)GetLongAttr(tint, "val") : null;
        var hasShade = shade != null ? (int?)GetLongAttr(shade, "val") : null;
        var hasLumMod = lumMod != null ? (int?)GetLongAttr(lumMod, "val") : null;
        var hasLumOff = lumOff != null ? (int?)GetLongAttr(lumOff, "val") : null;

        // No transforms needed — return raw hex
        if (hasTint == null && hasShade == null && hasLumMod == null && hasLumOff == null)
            return $"#{hex}";

        return ColorMath.ApplyTransforms(hex,
            tint: hasTint, shade: hasShade, lumMod: hasLumMod, lumOff: hasLumOff);
    }

    // CONSISTENCY(shape-fill-css): this solidFill/gradFill/pattFill → CSS mapping
    // is ~70% structurally duplicated by PowerPointHandler.GetShapeFillCss
    // (Pptx/PowerPointHandler.HtmlPreview.Css.cs). They diverge on element access
    // (untyped LocalName scan here vs SDK-typed GetFirstChild there) and ride
    // different tint/shade extraction before both delegate to ColorMath. Deferred
    // Core consolidation (e.g. Core/ShapeFillCss) — do NOT land as a one-handler
    // special case; unify cross-handler in one pass once the docx fix-storm settles.
    private string ResolveShapeFillCss(OpenXmlElement? spPr)
    {
        if (spPr == null) return "";

        // No fill
        if (spPr.Elements().Any(e => e.LocalName == "noFill")) return "";

        // Solid fill
        var solidFill = spPr.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        if (solidFill != null)
        {
            var rgb = solidFill.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
            if (rgb != null)
            {
                var val = rgb.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                if (val != null && IsHexColor(val)) return $"background-color:#{val}";
            }
            var scheme = solidFill.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
            if (scheme != null)
            {
                var color = ResolveSchemeColor(scheme);
                if (color != null) return $"background-color:{color}";
            }
        }

        // Gradient fill → CSS linear-gradient. OOXML stores stops as <a:gsLst>
        // with each <a:gs pos="N"/> (in 1/1000 of a percent). Direction comes
        // from <a:lin ang="N"/> (in 60000ths of a degree).
        var gradFill = spPr.Elements().FirstOrDefault(e => e.LocalName == "gradFill");
        if (gradFill != null)
        {
            var gsLst = gradFill.Elements().FirstOrDefault(e => e.LocalName == "gsLst");
            if (gsLst != null)
            {
                var stops = new List<string>();
                foreach (var gs in gsLst.Elements().Where(e => e.LocalName == "gs"))
                {
                    var posAttr = gs.GetAttributes().FirstOrDefault(a => a.LocalName == "pos").Value;
                    double pct = int.TryParse(posAttr, out var posVal) ? posVal / 1000.0 : 0;
                    string? color = null;
                    var gsRgb = gs.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
                    if (gsRgb != null)
                        color = "#" + gsRgb.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
                    var gsScheme = gs.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
                    if (gsScheme != null) color = ResolveSchemeColor(gsScheme);
                    if (color != null)
                        stops.Add($"{color} {pct:0.##}%");
                }
                if (stops.Count > 0)
                {
                    // ang: 60000ths of a degree; CSS linear-gradient uses "to <dir>" or "<deg>"
                    // OOXML 0 = left→right; CSS 0deg = bottom→top. Convert OOXML → CSS:
                    // CSS angle = (OOXML angle / 60000 + 90) % 360
                    var lin = gradFill.Elements().FirstOrDefault(e => e.LocalName == "lin");
                    double cssAngleDeg = 90;
                    var angAttr = lin?.GetAttributes().FirstOrDefault(a => a.LocalName == "ang").Value;
                    if (long.TryParse(angAttr, out var angVal))
                        cssAngleDeg = (angVal / 60000.0 + 90) % 360;
                    return $"background:linear-gradient({cssAngleDeg:0.##}deg,{string.Join(",", stops)})";
                }
            }
        }

        // Pattern fill → approximate the hatch with a CSS repeating-linear-gradient
        // alternating <a:fgClr> and <a:bgClr>. CSS can't reproduce every OOXML
        // preset (dkHorizontal, diagCross, …) so the angle is chosen from the
        // preset family (vertical / horizontal / diagonal); the result conveys
        // "patterned, not empty". Falls back to a solid fgClr when colors are
        // partially present, never to transparent.
        var pattFill = spPr.Elements().FirstOrDefault(e => e.LocalName == "pattFill");
        if (pattFill != null)
        {
            var fg = ResolvePatternColor(pattFill.Elements().FirstOrDefault(e => e.LocalName == "fgClr"));
            var bg = ResolvePatternColor(pattFill.Elements().FirstOrDefault(e => e.LocalName == "bgClr"));
            var prst = pattFill.GetAttributes().FirstOrDefault(a => a.LocalName == "prst").Value;

            if (fg != null && bg != null)
            {
                var angle = prst switch
                {
                    var p when p != null && p.Contains("Vertical", StringComparison.OrdinalIgnoreCase) => "90deg",
                    var p when p != null && p.Contains("Horizontal", StringComparison.OrdinalIgnoreCase) => "0deg",
                    _ => "45deg", // diagonal / cross / dotted / default
                };
                return $"background:repeating-linear-gradient({angle},{fg} 0 3px,{bg} 3px 6px)";
            }
            // Partial color info → solid fallback (existence over transparency).
            if (fg != null) return $"background-color:{fg}";
            if (bg != null) return $"background-color:{bg}";
        }

        return "";
    }

    /// <summary>Resolve an a:fgClr/a:bgClr wrapper to a CSS color, or null.</summary>
    private string? ResolvePatternColor(OpenXmlElement? clr)
    {
        if (clr == null) return null;
        var rgb = clr.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        if (rgb != null)
        {
            var val = rgb.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (val != null && IsHexColor(val)) return $"#{val}";
        }
        var scheme = clr.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
        if (scheme != null) return ResolveSchemeColor(scheme);
        return null;
    }

    private string ResolveShapeBorderCss(OpenXmlElement? spPr)
    {
        if (spPr == null) return "";
        var ln = spPr.Elements().FirstOrDefault(e => e.LocalName == "ln");
        if (ln == null) return "";
        if (ln.Elements().Any(e => e.LocalName == "noFill")) return "border:none";

        var solidFill = ln.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        if (solidFill == null) return "";

        string? color = null;
        var rgb = solidFill.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        if (rgb != null) {
            var rv = rgb.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (rv != null && IsHexColor(rv)) color = $"#{rv}";
        }
        var scheme = solidFill.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
        if (scheme != null) color = ResolveSchemeColor(scheme);

        var w = ln.GetAttributes().FirstOrDefault(a => a.LocalName == "w").Value;
        var widthPx = w != null && long.TryParse(w, out var emu) ? Math.Max(1, emu / EmuConverter.EmuPerPointF) : 1;

        var style = ResolveBorderDashStyle(ln);
        return $"border:{widthPx:0.#}px {style} {color ?? "#000"}";
    }

    /// <summary>
    /// Map an a:ln's a:prstDash preset to a CSS border-style. CSS has only
    /// solid/dashed/dotted; the OOXML dash family collapses accordingly.
    /// </summary>
    private static string ResolveBorderDashStyle(OpenXmlElement ln)
    {
        var prstDash = ln.Elements().FirstOrDefault(e => e.LocalName == "prstDash");
        var val = prstDash?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return val switch
        {
            "dot" or "sysDot" => "dotted",
            "dash" or "sysDash" or "lgDash"
                or "dashDot" or "lgDashDot" or "sysDashDot" or "sysDashDotDot" or "lgDashDotDot" => "dashed",
            _ => "solid", // "solid", null, or unknown
        };
    }

    // ==================== Color Math Helpers ====================

    /// <summary>Apply themeTint/themeShade to a base theme color hex.</summary>
    private static string ApplyTintShade(string hex, string? tintHex, string? shadeHex)
    {
        if (hex.Length < 6) return $"#{hex}";
        var (r, g, b) = ColorMath.HexToRgb(hex);

        // themeTint: blend toward white (tint value is hex 00-FF)
        if (tintHex != null && int.TryParse(tintHex, System.Globalization.NumberStyles.HexNumber, null, out var tint))
        {
            var t = tint / 255.0;
            r = (int)(r * t + 255 * (1 - t));
            g = (int)(g * t + 255 * (1 - t));
            b = (int)(b * t + 255 * (1 - t));
        }

        // themeShade: blend toward black
        if (shadeHex != null && int.TryParse(shadeHex, System.Globalization.NumberStyles.HexNumber, null, out var shade))
        {
            var s = shade / 255.0;
            r = (int)(r * s);
            g = (int)(g * s);
            b = (int)(b * s);
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private static long GetLongAttr(OpenXmlElement? el, string attrName, long defaultVal = 0)
    {
        if (el == null) return defaultVal;
        var val = el.GetAttributes().FirstOrDefault(a => a.LocalName == attrName).Value;
        return val != null && long.TryParse(val, out var v) ? v : defaultVal;
    }


    // ==================== Inline CSS ====================

    // True for the document's first body block-level paragraph (top of page 1).
    // Word suppresses spaceBefore at the top of a page; we render it flush.
    // Excludes paragraphs nested in tables/headers/footers (Parent != Body) and
    // any paragraph preceded by a sibling paragraph or table.
    private bool IsFirstBodyParagraph(Paragraph para)
    {
        if (para.Parent is not Body body) return false;
        return body.Elements()
            .FirstOrDefault(e => e is Paragraph || e is Table) == para;
    }

    private string GetParagraphInlineCss(Paragraph para, bool isListItem = false)
    {
        var parts = new List<string>();

        // Set paragraph font-size and font-family to match the first run.
        // This keeps the paragraph's anonymous inline box (strut) sized in the
        // same metrics as the actual text spans, preventing line-box inflation
        // when the page-level defaults differ from the run.
        // For empty paragraphs (no text-bearing run) Word stores the
        // would-be content's font/size on pPr/rPr (the paragraph mark's run
        // properties), so synthesize a Run from those props and run it
        // through the same resolver — the strut metrics then match what Word
        // would have rendered if there had been content.
        Run? probeRun = para.Elements<Run>().FirstOrDefault(r =>
            r.ChildElements.Any(c => c is Text t && !string.IsNullOrEmpty(t.Text)));
        if (probeRun == null)
        {
            var markProps = para.ParagraphProperties?.ParagraphMarkRunProperties;
            if (markProps != null)
            {
                var synthRPr = new RunProperties();
                foreach (var child in markProps.ChildElements)
                    synthRPr.AppendChild(child.CloneNode(true));
                probeRun = new Run(synthRPr);
            }
        }
        double? paraFontSizePt = null;
        if (probeRun != null)
        {
            var rProps = ResolveEffectiveRunProperties(probeRun, para);
            var sz = rProps.FontSize?.Val?.Value;
            if (sz != null && int.TryParse(sz, out var hp))
            {
                parts.Add($"font-size:{hp / 2.0:0.##}pt");
                paraFontSizePt = hp / 2.0;
            }

            var fonts = rProps.RunFonts;
            var paraFont = fonts?.EastAsia?.Value ?? ResolveThemeFont(fonts?.EastAsiaTheme?.InnerText)
                ?? fonts?.Ascii?.Value ?? ResolveThemeFont(fonts?.AsciiTheme?.InnerText)
                ?? fonts?.HighAnsi?.Value ?? ResolveThemeFont(fonts?.HighAnsiTheme?.InnerText);
            if (!string.IsNullOrEmpty(paraFont)
                && !paraFont.StartsWith("+", StringComparison.Ordinal)
                && !string.Equals(paraFont, ReadDocDefaults().Font, StringComparison.Ordinal))
            {
                var fallback = GetChineseFontFallback(paraFont);
                var generic = GenericFontFamily(paraFont);
                parts.Add(fallback != null
                    ? $"font-family:'{CssSanitize(paraFont)}',{fallback},{generic}"
                    : $"font-family:'{CssSanitize(paraFont)}',{generic}");
            }
        }

        var pProps = para.ParagraphProperties;
        if (pProps == null)
        {
            var styleCss = ResolveParagraphStyleCss(para);
            if (parts.Count > 0 && !string.IsNullOrEmpty(styleCss))
                return string.Join(";", parts) + ";" + styleCss;
            if (parts.Count > 0) return string.Join(";", parts);
            return styleCss;
        }

        // Style ID for fallback lookups
        var styleId = pProps.ParagraphStyleId?.Val?.Value;

        // Alignment (direct or from style chain)
        var jc = pProps.Justification?.Val;
        if (jc == null) jc = ResolveJustificationFromStyle(styleId);
        if (jc != null)
        {
            var jcVal = jc.InnerText;
            var align = jcVal switch
            {
                "center" => "center",
                "right" or "end" => "right",
                "both" or "distribute" => "justify",
                _ => (string?)null
            };
            if (align != null) parts.Add($"text-align:{align}");
            // w:jc="distribute" stretches EVERY line (including single/last)
            // to full width with inter-character spacing. Plain CSS justify
            // leaves the last line unstretched, so add text-align-last
            // and text-justify hints for closer fidelity.
            if (jcVal == "distribute")
                parts.Add("text-align-last:justify;text-justify:inter-character");
        }

        // Paragraph-level RTL (w:bidi) — flips the paragraph direction
        if (pProps.BiDi != null && (pProps.BiDi.Val == null || pProps.BiDi.Val.Value))
        {
            parts.Add("direction:rtl");
            // Word right-aligns an RTL paragraph that has no explicit jc
            // (start edge = right). The preview's global `p{text-align:left}`
            // rule otherwise forces it left, so emit text-align:right when no
            // explicit alignment was resolved above. Paragraphs with an
            // explicit jc (center/right/justify) already added text-align and
            // must not be overridden.
            if (jc == null)
                parts.Add("text-align:right");
        }

        // Drop cap detection — used to suppress text-indent
        var framePrForIndent = pProps.GetFirstChild<FrameProperties>();
        var hasDropCap = framePrForIndent != null &&
            framePrForIndent.GetAttributes().FirstOrDefault(a => a.LocalName == "dropCap").Value is "drop" or "margin";

        // Indentation (skip for list items — handled by list nesting)
        if (!isListItem)
        {
            // Indentation — merge direct properties with style chain fallback
            var directInd = pProps.Indentation;
            var styleInd = ResolveIndentationFromStyle(styleId);
            var indLeft = directInd?.Left?.Value ?? styleInd?.Left?.Value;
            var indRight = directInd?.Right?.Value ?? styleInd?.Right?.Value;
            var indFirstLine = directInd?.FirstLine?.Value ?? styleInd?.FirstLine?.Value;
            var indHanging = directInd?.Hanging?.Value ?? styleInd?.Hanging?.Value;
            // *Chars variants: indentation expressed as 100ths of an East-Asian
            // character width. Convert against the paragraph's effective font
            // size (fallback 10.5pt = Normal default) when the twips counterpart
            // is absent. Direct overrides win; otherwise inherit style chain.
            var indLeftChars = directInd?.LeftChars?.Value ?? styleInd?.LeftChars?.Value;
            var indRightChars = directInd?.RightChars?.Value ?? styleInd?.RightChars?.Value;
            var indFirstLineChars = directInd?.FirstLineChars?.Value ?? styleInd?.FirstLineChars?.Value;
            var indHangingChars = directInd?.HangingChars?.Value ?? styleInd?.HangingChars?.Value;
            double charWidthPt = paraFontSizePt ?? 10.5;

            // Hanging indent needs left padding/margin equal to the hanging
            // amount to produce the visual effect (first line at 0, follow
            // lines indented). When only `hanging` is set without `left`,
            // use hanging as the left margin too.
            double? hangPt = null;
            if (indHanging is string hpTwips && hpTwips != "0")
                hangPt = Units.TwipsToPt(hpTwips);
            else if (indHangingChars is int hpChars && hpChars != 0)
                hangPt = hpChars / 100.0 * charWidthPt;
            double leftPt = 0;
            if (indLeft is string leftTwips && leftTwips != "0")
                leftPt = Units.TwipsToPt(leftTwips);
            else if (indLeftChars is int leftChars && leftChars != 0)
                leftPt = leftChars / 100.0 * charWidthPt;
            // When hanging is set and left is 0, promote hanging into left
            // margin so subsequent lines visibly indent.
            if (hangPt.HasValue && leftPt == 0) leftPt = hangPt.Value;
            if (leftPt != 0)
                parts.Add($"margin-left:{leftPt:0.##}pt");
            if (indRight is string rightTwips && rightTwips != "0")
                parts.Add($"margin-right:{Units.TwipsToPt(rightTwips):0.##}pt");
            else if (indRightChars is int rightChars && rightChars != 0)
                parts.Add($"margin-right:{rightChars / 100.0 * charWidthPt:0.##}pt");
            if (!hasDropCap)
            {
                if (indFirstLine is string firstLineTwips && firstLineTwips != "0")
                    parts.Add($"text-indent:{Units.TwipsToPt(firstLineTwips):0.##}pt");
                else if (indFirstLineChars is int firstLineChars && firstLineChars != 0)
                    parts.Add($"text-indent:{firstLineChars / 100.0 * charWidthPt:0.##}pt");
                if (hangPt.HasValue)
                    parts.Add($"text-indent:-{hangPt.Value:0.##}pt");
            }
        }

        // Spacing — direct properties first, fallback to style chain per-property
        var spacing = pProps.SpacingBetweenLines;
        var styleSpacing = ResolveSpacingFromStyle(styleId);
        if (spacing == null)
            spacing = styleSpacing;

        // Paragraph before/after spacing always renders OUTSIDE the border box
        // (verified against real Word): the border hugs the text — its internal
        // text-to-border gap comes solely from the border's w:space, emitted as
        // padding — and spaceBefore/spaceAfter sit above/below the border as
        // margin. Mapping spacing to padding when borders are present put the
        // space INSIDE the box (tall border, no gap below) and additionally
        // collided with the w:space padding on the same side (last-wins).
        var vSpacingPropBefore = "margin-top";
        var vSpacingPropAfter = "margin-bottom";

        // Continuous-shaded-box margin suppression. When consecutive paragraphs
        // share an identical pBdr (the OOXML §17.3.1.24 border-merge condition
        // handled below at the border block) AND each carries a paragraph-level
        // shd fill, Word renders them as ONE continuous shaded box with no
        // internal gap — the fill of one paragraph abuts the next. HTML paints
        // a paragraph's background only inside its content/padding box, never
        // into vertical margins, so any spaceBefore/spaceAfter (here typically
        // an inherited docDefaults `w:after`) opens a white band between the
        // strips and visually shreds the box. Mirror Word by zeroing the
        // inter-paragraph margin on the joined edge — the same mechanism
        // contextualSpacing uses — so the shaded strips touch. Scoped to the
        // shd+identical-pBdr pair (a lone shaded paragraph, or shaded paragraphs
        // with differing/absent borders, keeps its normal margin), so it does
        // not perturb normal spacing, R66 border merge, or R80 table-style shd
        // (table cells, not body paragraphs).
        bool continuousShadeBefore = ParagraphJoinsShadedBox(para, para.PreviousSibling() as Paragraph);
        bool continuousShadeAfter = ParagraphJoinsShadedBox(para, para.NextSibling() as Paragraph);

        if (spacing != null)
        {
            // contextualSpacing: when enabled and adjacent paragraph has the same style,
            // spaceBefore/spaceAfter between them is suppressed (set to zero).
            // w:contextualSpacing is an on/off toggle: present-with-no-val means ON,
            // but w:val="0"/"false"/"off" means explicitly OFF (do not suppress).
            var hasContextualSpacing = IsContextualSpacingOn(pProps.ContextualSpacing)
                ?? ResolveContextualSpacingFromStyle(styleId);
            var prevPara = para.PreviousSibling<Paragraph>();
            var nextPara = para.NextSibling<Paragraph>();
            var prevStyleId = prevPara?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var nextStyleId = nextPara?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            bool suppressBefore = (hasContextualSpacing && prevPara != null
                && (prevStyleId ?? "") == (styleId ?? "")) || continuousShadeBefore;
            bool suppressAfter = (hasContextualSpacing && nextPara != null
                && (nextStyleId ?? "") == (styleId ?? "")) || continuousShadeAfter;

            // Before/after spacing: w:before is in twips; w:beforeLines is in
            // hundredths of a line. Per ECMA-376 §17.3.1.33 beforeLines
            // OVERRIDES before when both are present. The "1 line" base unit
            // is implementation-defined; (and Word) anchor it to
            // 240 twips = 12pt FIXED, not the paragraph's font line.
            const double LineUnitPt = 12.0;

            static double? ResolveSpacingPt(string? twips, int? lines)
            {
                if (lines is int n) return n / 100.0 * LineUnitPt;  // beforeLines wins
                if (twips != null && int.TryParse(twips, out var tw)) return tw / 20.0;
                return null;
            }

            // OOXML §17.3.1.5 beforeAutospacing / §17.3.1.4 afterAutospacing:
            // when set, the spec's "application-determined autospacing"
            // substitutes a 280-twip (14pt) baseline for the literal
            // Before/After before margin collapse. Common in HTML-imported
            // docx where the flag mirrors browser <p>-margin defaults.
            //
            // Suppression in table cells: the cell boundary (tcMar) already
            // provides the visual gap, so autospacing is fully suppressed
            // for paragraphs directly inside a TableCell — both for adjacent
            // pairs (cell-internal collapse) and for first/last paragraphs
            // in the cell (cell-edge collapse).
            const string AutospacingTwips = "280";
            var inTableCell = para.Parent is TableCell;
            var prevInSameCell = inTableCell;
            var nextInSameCell = inTableCell;

            var beforeAutoRaw = (pProps.SpacingBetweenLines?.BeforeAutoSpacing?.Value
                                 ?? styleSpacing?.BeforeAutoSpacing?.Value) == true;
            var beforeAuto = beforeAutoRaw && !prevInSameCell;
            var beforeVal = beforeAuto ? AutospacingTwips
                : (beforeAutoRaw && prevInSameCell ? "0"
                   : (pProps.SpacingBetweenLines?.Before?.Value
                      ?? styleSpacing?.Before?.Value));
            var beforeLinesVal = beforeAuto || beforeAutoRaw ? null
                : (pProps.SpacingBetweenLines?.BeforeLines?.Value
                   ?? styleSpacing?.BeforeLines?.Value);

            // Word collapses adjacent spaceBefore/spaceAfter to max(prev.after, cur.before)
            // instead of adding them. The HTML paragraphs are normal block-flow siblings,
            // so their vertical margins ALSO collapse (CSS takes the max of adjacent
            // margins). We therefore emit each paragraph's OWN spaceBefore/spaceAfter in
            // full and let CSS margin-collapse reproduce Word's max() naturally.
            // (Subtracting the previous sibling's spaceAfter here was wrong: collapse
            // takes the max, not the sum, so the subtraction UNDERSTATED the gap.)

            // Word suppresses spaceBefore at the TOP of a page: the document's
            // first body paragraph renders flush at the top margin (verified
            // against real Word). Mirrors the PowerPoint first-paragraph fix.
            // Scope to the body's first block-level paragraph only — paragraphs
            // inside tables/headers/footers keep their spaceBefore.
            if (!suppressBefore && IsFirstBodyParagraph(para))
                suppressBefore = true;

            if (suppressBefore)
            {
                parts.Add($"{vSpacingPropBefore}:0");
            }
            else
            {
                var beforePt = ResolveSpacingPt(beforeVal, beforeLinesVal);
                if (beforePt is double bp && bp > 0)
                    parts.Add($"{vSpacingPropBefore}:{bp:0.##}pt");
            }

            var afterAutoRaw = (pProps.SpacingBetweenLines?.AfterAutoSpacing?.Value
                                ?? styleSpacing?.AfterAutoSpacing?.Value) == true;
            var afterAuto = afterAutoRaw && !nextInSameCell;
            var afterVal = afterAuto ? AutospacingTwips
                : (afterAutoRaw && nextInSameCell ? "0"
                   : (pProps.SpacingBetweenLines?.After?.Value
                      ?? styleSpacing?.After?.Value));
            var afterLinesVal = afterAuto || afterAutoRaw ? null
                : (pProps.SpacingBetweenLines?.AfterLines?.Value
                   ?? styleSpacing?.AfterLines?.Value);
            if (suppressAfter)
            {
                parts.Add($"{vSpacingPropAfter}:0");
            }
            else
            {
                var afterPt = ResolveSpacingPt(afterVal, afterLinesVal);
                if (afterPt is double ap)
                    parts.Add($"{vSpacingPropAfter}:{ap:0.##}pt");
            }

            // Line: try direct, then style fallback
            var lineVal = pProps.SpacingBetweenLines?.Line?.Value
                          ?? styleSpacing?.Line?.Value;
            if (lineVal is string lv)
            {
                var rule = pProps.SpacingBetweenLines?.LineRule?.InnerText
                           ?? styleSpacing?.LineRule?.InnerText;
                if (rule == "auto" || rule == null)
                {
                    if (int.TryParse(lv, out var lvNum) && lvNum > 0)
                    {
                        // OOXML §17.3.1.33 "auto" rule: line value is in
                        // 240ths of a line. Final line-height multiplies
                        // the font's natural single-line ratio by the
                        // per-paragraph (lvNum/240) factor.
                        // CSS unitless: line-height = (lvNum/240) × natural_ratio
                        var paraFont = ResolveParaFontForLineHeight(para);
                        var ratio = FontMetricsReader.GetRatio(paraFont);
                        var lh = ratio * (lvNum / 240.0);
                        parts.Add($"line-height:{lh:0.####}");
                    }
                }
                else if (rule == "exact" || rule == "atLeast")
                {
                    var linePt = Units.TwipsToPt(lv);
                    // OOXML §17.3.1.33 atLeast: floor only. When the
                    // paragraph's natural single-line height exceeds the
                    // floor, the natural value applies.
                    var emitPt = rule == "atLeast" ? ResolveAtLeastPt(linePt, para) : linePt;
                    parts.Add($"line-height:{emitPt:0.##}pt");
                    // lineRule=exact pins the line box to a fixed height, but
                    // Word still shows the text — it does NOT erase a line whose
                    // content is taller than the exact box; over-tall glyphs are
                    // visually clipped at the box edge, not blanked out. The
                    // earlier overflow:hidden (on a fixed box) blanked whole
                    // labels/list-rows when the natural content height exceeded
                    // the exact value (content loss). line-height alone reproduces
                    // the fixed leading while keeping content visible; we no
                    // longer emit overflow:hidden here. Priority: content visible
                    // over strict exact height (R49/R31 don't-clip-content rule).
                }
            }

            // If no explicit line-height was set, use font metrics ratio
            if (!parts.Any(p => p.StartsWith("line-height")))
            {
                var paraFont = ResolveParaFontForLineHeight(para);
                var ratio = FontMetricsReader.GetRatio(paraFont);
                if (ratio > 1.01 || ratio < 0.99) // only if meaningfully different from 1.0
                    parts.Add($"line-height:{ratio:0.####}");
            }

        }
        else
        {
            // No explicit <w:spacing> on paragraph or anywhere in its style chain.
            // Word may still apply baked-in defaults from Normal.dotm — but only
            // when the doc actually carries Normal defaults (Normal style defined
            // OR docDefaults/pPrDefault populated). When neither is present (rare
            // in real-world docs, common in synthetic fixtures), Word emits zero
            // spacing; mirroring that keeps cli aligned without needing the user
            // to put explicit <w:spacing> on every paragraph.
            var builtIn = ResolveBuiltInStyleDefaults(styleId);
            if (builtIn == null && DocCarriesNormalDefaults())
                builtIn = BuiltInStyleDefaults["Normal"];

            // contextualSpacing must suppress before/after between same-style
            // siblings even when the resolved spacing comes from BuiltInStyleDefaults
            // (typical for ListParagraph: built-in After=10pt, but contextualSpacing
            // on the style should collapse it to 0 between adjacent bullets).
            var hasContextualSpacing = IsContextualSpacingOn(pProps.ContextualSpacing)
                ?? ResolveContextualSpacingFromStyle(styleId);
            var prevPara = para.PreviousSibling<Paragraph>();
            var nextPara = para.NextSibling<Paragraph>();
            var prevStyleId = prevPara?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var nextStyleId = nextPara?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            bool suppressBefore = (hasContextualSpacing && prevPara != null
                && (prevStyleId ?? "") == (styleId ?? "")) || continuousShadeBefore;
            bool suppressAfter = (hasContextualSpacing && nextPara != null
                && (nextStyleId ?? "") == (styleId ?? "")) || continuousShadeAfter;

            // Word collapses adjacent spaceBefore/spaceAfter to max(prev.after, cur.before).
            // The HTML paragraphs are normal block-flow siblings, so their vertical margins
            // ALSO collapse to the max — we therefore emit each paragraph's own spaceBefore
            // in full and let CSS margin-collapse reproduce Word's max() naturally. (The old
            // subtraction of the previous sibling's spaceAfter understated the gap.)

            var paraFontDef = ResolveParaFontForLineHeight(para);
            var ratioDef = FontMetricsReader.GetRatio(paraFontDef);

            if (builtIn != null)
            {
                var beforePt = suppressBefore ? 0 : builtIn.Before;
                if (beforePt > 0)
                    parts.Add($"{vSpacingPropBefore}:{beforePt:0.##}pt");
                var afterPt = suppressAfter ? 0 : builtIn.After;
                if (afterPt > 0)
                    parts.Add($"{vSpacingPropAfter}:{afterPt:0.##}pt");
                // Use built-in line multiplier, but raise to font metric ratio when the
                // font's natural ascent+descent exceeds it (CJK / glyph-tall fonts).
                var lhDef = Math.Max(builtIn.Line, ratioDef);
                parts.Add($"line-height:{lhDef:0.####}");
            }
            else
            {
                // Doc carries no Normal defaults. Emit no margin — let the line
                // box pure-stack at the natural single-line height. Still emit
                // CJK ratio so SimSun/etc. render at their full em height.
                if (ratioDef > 1.01 || ratioDef < 0.99)
                    parts.Add($"line-height:{ratioDef:0.####}");
            }

            // NOTE: do not emit font-size/bold/color from BuiltInStyleDefaults here.
            // Per ECMA-376, when a paragraph references a style that is undefined
            // in the doc, Word renders as if no style applied — it does NOT pull
            // font-size/bold/color from Normal.dotm. Those Normal.dotm built-ins
            // are template-specific, not standard. Verified against formulas.docx:
            // Heading1/Heading2 referenced without styles.xml render as plain 11pt
            // black in real Word. Only spacing/line-height are kept here because
            // Word still applies Normal-equivalent paragraph defaults regardless.
        }

        // docGrid snap: when type="lines" and paragraph doesn't opt out via snapToGrid=false,
        // snap line-height to the nearest multiple of linePitch that fits the text.
        {
            var snapToGrid = pProps.SnapToGrid?.Val?.Value ?? true;
            if (snapToGrid)
            {
                var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
                var dg = sectPr?.GetFirstChild<DocGrid>();
                if ((dg?.Type?.Value == DocGridValues.Lines || dg?.Type?.Value == DocGridValues.LinesAndChars)
                    && dg.LinePitch?.Value is int lp && lp > 0)
                {
                    double gridPitchPt = lp / 20.0;
                    var gFont = ResolveParaFontForLineHeight(para);
                    var gRatio = FontMetricsReader.GetRatio(gFont);
                    double gSizePt = 0;
                    var gFirstRun = para.Elements<Run>().FirstOrDefault(r =>
                        r.ChildElements.Any(c => c is Text t && !string.IsNullOrEmpty(t.Text)));
                    if (gFirstRun != null)
                    {
                        var grProps = ResolveEffectiveRunProperties(gFirstRun, para);
                        if (grProps.FontSize?.Val?.Value is string gsz && int.TryParse(gsz, out var ghp))
                            gSizePt = ghp / 2.0;
                    }
                    if (gSizePt <= 0) gSizePt = 12.0;

                    double fontHeightPt = gSizePt * gRatio;
                    double snappedPt = Math.Ceiling(fontHeightPt / gridPitchPt) * gridPitchPt;
                    parts.RemoveAll(p => p.StartsWith("line-height"));
                    parts.Add($"line-height:{snappedPt:0.##}pt");
                }
            }
        }

        // Shading / background (direct or from style)
        var shading = pProps.Shading;
        var fillColor = ResolveShadingFill(shading);
        if (fillColor != null)
            parts.Add($"background-color:{fillColor}");
        else
        {
            // Try to resolve from paragraph style
            var bgFromStyle = ResolveParagraphShadingFromStyle(para);
            if (bgFromStyle != null) parts.Add($"background-color:{bgFromStyle}");
        }

        // Borders — pBdr on the paragraph itself wins; otherwise fall through
        // the pStyle chain (e.g. the `Title` style ships a bottom border that
        // the para never re-declares, so without this fallback the blue rule
        // under a title is silently dropped).
        var pBdr = pProps.ParagraphBorders
            ?? ResolveStyleParagraphBorders(pProps.ParagraphStyleId?.Val?.Value);
        if (pBdr != null)
        {
            // OOXML §17.3.1.24 border merging: when consecutive paragraphs carry
            // an identical pBdr (same val/color/sz/space on each side, no explicit
            // w:between), Word renders them as ONE continuous box — no internal
            // top/bottom rule between the stacked paragraphs. HTML emits per-para
            // borders, so without suppression the shared box shows a doubled
            // horizontal divider that splits the logical box into stacked
            // sub-boxes. Suppress the inner edge when the adjacent sibling shares
            // the same pBdr: drop border-top if the previous sibling matches, drop
            // border-bottom if the next sibling matches. Left/right always emit.
            var prevBdr = ResolveSiblingParagraphBorders(para.PreviousSibling() as Paragraph);
            var nextSiblingBdr = ResolveSiblingParagraphBorders(para.NextSibling() as Paragraph);
            var suppressTop = pBdr.BetweenBorder == null && ParagraphBordersEqual(pBdr, prevBdr);
            var suppressBottom = pBdr.BetweenBorder == null && ParagraphBordersEqual(pBdr, nextSiblingBdr);

            if (!suppressTop) RenderBorderCss(parts, pBdr.TopBorder, "border-top");
            if (!suppressBottom) RenderBorderCss(parts, pBdr.BottomBorder, "border-bottom");
            RenderBorderCss(parts, pBdr.LeftBorder, "border-left");
            RenderBorderCss(parts, pBdr.RightBorder, "border-right");
            // w:between draws a rule BETWEEN consecutive paragraphs that share
            // the same pBdr (OOXML §17.3.1.24). HTML has no native "between"
            // border, so approximate as a bottom-border on the upper paragraph
            // when the following sibling paragraph also carries a matching
            // pBdr — and only when no explicit w:bottom already painted that
            // edge (an explicit bottom wins on the para's own outer box).
            if (pBdr.BetweenBorder != null && pBdr.BottomBorder == null
                && para.NextSibling() is Paragraph nextPara
                && (nextPara.ParagraphProperties?.ParagraphBorders
                    ?? ResolveStyleParagraphBorders(nextPara.ParagraphProperties?.ParagraphStyleId?.Val?.Value))
                   is ParagraphBorders nextBdr
                && (nextBdr.BetweenBorder != null || nextBdr.TopBorder != null))
            {
                RenderBorderCss(parts, pBdr.BetweenBorder, "border-bottom");
            }
        }

        // Page break before
        if (pProps.PageBreakBefore?.Val?.Value != false && pProps.PageBreakBefore != null)
            parts.Add("page-break-before:always");

        // Drop cap (framePr with dropCap attribute)
        var framePr = pProps.GetFirstChild<FrameProperties>();
        if (framePr != null)
        {
            var dropCap = framePr.GetAttributes().FirstOrDefault(a => a.LocalName == "dropCap").Value;
            if (dropCap == "drop" || dropCap == "margin")
            {
                // OOXML §17.3.1.36 framePr/dropCap: the cap glyph renders at the
                // run's effective font size from the rPr cascade (run → style →
                // docDefaults). The framed paragraph hosts a float whose box
                // bounds the cap glyph's visible vertical extent so wrap text
                // flows alongside. Container height = font_size × full-ascent
                // ratio (top of the inline strut; reaches above cap-height for
                // accented capitals) so the ink isn't clipped and no trailing
                // whitespace runs past the visible glyph.
                var dropCapSizePt = ResolveParaPrincipalSizePt(para) ?? 11.0;
                var dropCapFont = ResolveParaFontForLineHeight(para);
                var (dropCapAscPct, _) = Core.FontMetricsReader.GetSplitAscDscOverride(dropCapFont);
                var ascRatio = dropCapAscPct > 0 ? dropCapAscPct / 100.0 : 0.95;
                var dropCapHeight = dropCapSizePt * ascRatio;
                var hSpaceAttr = framePr.GetAttributes().FirstOrDefault(a => a.LocalName == "hSpace").Value;
                var hSpacePt = hSpaceAttr != null && int.TryParse(hSpaceAttr, out var hsTwips) ? hsTwips / 20.0 : 0;
                parts.Add("float:left");
                parts.Add($"line-height:{dropCapHeight:0.#}pt");
                // Clip the float so the cap glyph's natural strut can't push
                // the box taller than the visible cap.
                parts.Add($"height:{dropCapHeight:0.#}pt");
                parts.Add("overflow:hidden");
                parts.Add($"padding-right:{hSpacePt:0.#}pt");
                parts.Add($"margin:0");
            }
        }

        return string.Join(";", parts);
    }

    /// <summary>
    /// Resolve paragraph background shading from the style chain.
    /// </summary>
    private string? ResolveParagraphShadingFromStyle(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = FindStyleById(currentStyleId);
            if (style == null) break;

            var shading = style.StyleParagraphProperties?.Shading;
            var sFill = ResolveShadingFill(shading);
            if (sFill != null) return sFill;

            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>
    /// Resolve Justification from the style chain.
    /// </summary>
    private JustificationValues? ResolveJustificationFromStyle(string? styleId)
    {
        if (styleId == null) return null;
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = FindStyleById(currentStyleId);
            if (style == null) break;
            var jc = style.StyleParagraphProperties?.Justification?.Val;
            if (jc != null) return jc;
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>
    /// Resolve PageBreakBefore from the style chain.
    /// Falls back to Word built-in defaults for latent styles not defined in styles.xml.
    /// </summary>
    private PageBreakBefore? ResolvePageBreakBeforeFromStyle(string? styleId)
    {
        if (styleId == null) return null;
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = FindStyleById(currentStyleId);
            if (style == null)
            {
                // Word built-in TOCHeading has pageBreakBefore=true by default
                if (currentStyleId == "TOCHeading")
                    return new PageBreakBefore();
                break;
            }
            var pgBB = style.StyleParagraphProperties?.PageBreakBefore;
            if (pgBB != null) return pgBB;
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>
    /// Resolve SpacingBetweenLines from the style chain (basedOn walk).
    /// </summary>
    private IEnumerable<TabStop>? ResolveTabStopsFromStyle(string? styleId)
    {
        if (styleId == null) return null;
        // Word ACCUMULATES w:tabs across the basedOn chain (verified against
        // real Word): a child style's tab list adds to the parent's rather
        // than replacing it. A derived declaration at the same position wins
        // (and w:val="clear" removes the inherited stop at that position).
        var byPos = new Dictionary<int, TabStop>();
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = FindStyleById(currentStyleId);
            if (style == null) break;
            var tabs = style.StyleParagraphProperties?.Tabs?.Elements<TabStop>();
            if (tabs != null)
                foreach (var t in tabs)
                {
                    var pos = t.Position?.Value;
                    if (pos == null) continue;
                    // Derived styles are visited first — first declaration wins.
                    if (!byPos.ContainsKey(pos.Value)) byPos[pos.Value] = t;
                }
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        if (byPos.Count == 0) return null;
        var mergedTabs = byPos
            .Where(kv => kv.Value.Val == null || kv.Value.Val.InnerText != "clear")
            .OrderBy(kv => kv.Key)
            .Select(kv => kv.Value)
            .ToList();
        return mergedTabs.Count > 0 ? mergedTabs : null;
    }

    /// <summary>Word built-in style defaults (Office 2010+ Normal.dotm baseline).
    /// Used when the style is referenced but undefined in the doc, OR defined
    /// without these properties — Word fills in baked-in values regardless.
    /// Progressive — covers spacing/line/size/bold/color. Italic/keepWithNext
    /// still missing. Terminal goal is full-fidelity built-in style table.</summary>
    private record BuiltInStyleDefault(
        double Before, double After, double Line,
        double? SizePt, bool Bold, string? ColorHex);

    private static readonly System.Collections.Generic.Dictionary<string, BuiltInStyleDefault> BuiltInStyleDefaults
        = new(System.StringComparer.OrdinalIgnoreCase)
    {
        // Normal: Office 2010 baseline (10pt after, 1.15 line). Office 2013+ uses
        // 8pt/1.08; we keep 2010 values for consistency with global else-branch fallback.
        ["Normal"]       = new(0,  10, 1.15, null, false, null),
        ["Heading1"]     = new(12,  0, 1.08, 16,   true,  "#2E74B5"),
        ["Heading2"]     = new( 2,  0, 1.08, 13,   true,  "#2E74B5"),
        ["Heading3"]     = new( 2,  0, 1.08, 12,   true,  "#1F3864"),
        ["Heading4"]     = new( 2,  0, 1.08, 11,   true,  "#2E74B5"),
        ["Heading5"]     = new( 2,  0, 1.08, 11,   false, "#2E74B5"),
        ["Heading6"]     = new( 2,  0, 1.08, 11,   false, "#1F3864"),
        ["Heading7"]     = new( 2,  0, 1.08, 11,   false, "#1F3864"),
        ["Heading8"]     = new( 2,  0, 1.08, 11,   false, "#2E74B5"),
        ["Heading9"]     = new( 2,  0, 1.08, 11,   false, "#2E74B5"),
        ["Title"]        = new( 0,  0, 1.0,  28,   false, null),
        ["Subtitle"]     = new( 0,  0, 1.15, 11,   false, "#5A5A5A"),
        ["ListParagraph"]= new( 0, 10, 1.15, null, false, null),  // contextualSpacing handled separately
        ["Quote"]        = new( 0,  0, 1.15, null, false, null),
        ["IntenseQuote"] = new( 0,  0, 1.15, null, true,  "#2E74B5"),
    };

    /// <summary>Walk the style chain and return Word's built-in defaults for the
    /// first style that (1) is actually defined in the doc and (2) matches a known
    /// built-in name, OR is referenced as the doc's default Normal-equivalent.
    /// Per ECMA-376, when a style is referenced but undefined, Word treats the
    /// paragraph as styleless — it does NOT inherit Normal.dotm's Heading1
    /// built-ins. Verified against formulas.docx: pStyle="Heading1" without
    /// styles.xml renders as plain 11pt black, no 12pt spaceBefore.
    /// Returns null when no defined style in the chain matches a built-in.</summary>
    private BuiltInStyleDefault? ResolveBuiltInStyleDefaults(string? styleId)
    {
        if (styleId == null) return null;
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = FindStyleById(current);
            if (style == null) return null;  // Undefined style → no built-in inheritance.
            if (BuiltInStyleDefaults.TryGetValue(current, out var defaults))
                return defaults;
            current = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private bool? _docCarriesNormalDefaultsCache;
    /// <summary>
    /// Whether this doc carries Normal-style paragraph defaults. True when EITHER
    /// the doc's styles.xml defines a Normal-equivalent paragraph style (a style
    /// named "Normal" or one with default="1"), OR docDefaults/pPrDefault carries
    /// a spacing element. False when the doc has no Normal style and an empty
    /// pPrDefault (synthetic test fixtures, raw XML hand-built docs) — Word
    /// renders such paragraphs with no implicit Normal.dotm baseline, so cli
    /// shouldn't inject one either.
    /// </summary>
    private bool DocCarriesNormalDefaults()
    {
        if (_docCarriesNormalDefaultsCache.HasValue) return _docCarriesNormalDefaultsCache.Value;
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        bool result = false;
        if (styles != null)
        {
            // (1) styles.xml defines Normal or another paragraph style flagged default="1"
            foreach (var s in styles.Elements<Style>())
            {
                if (s.Type?.Value != StyleValues.Paragraph) continue;
                if (string.Equals(s.StyleId?.Value, "Normal", StringComparison.OrdinalIgnoreCase)
                    || s.Default?.Value == true)
                {
                    result = true;
                    break;
                }
            }
            // (2) docDefaults/pPrDefault carries a <w:spacing> element
            if (!result)
            {
                var pPrDef = styles.GetFirstChild<DocDefaults>()?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
                if (pPrDef?.SpacingBetweenLines != null)
                    result = true;
            }
        }
        _docCarriesNormalDefaultsCache = result;
        return result;
    }

    private SpacingBetweenLines? ResolveSpacingFromStyle(string? styleId)
    {
        // Per OOXML, each attribute on <w:spacing> inherits independently
        // through the basedOn chain. A derived style overriding only `after`
        // must still pick up `before`/`beforeLines`/`line`/`lineRule` from
        // its base. Element-level resolution (returning the first non-null
        // sp in the walk) loses inherited attributes that aren't restated
        // on the derived style.
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return null;

        var merged = new SpacingBetweenLines();
        bool anySet = false;

        void MergeFrom(SpacingBetweenLines? sp)
        {
            if (sp == null) return;
            if (merged.Before == null && sp.Before != null) { merged.Before = sp.Before.Value; anySet = true; }
            if (merged.BeforeLines == null && sp.BeforeLines != null) { merged.BeforeLines = sp.BeforeLines.Value; anySet = true; }
            if (merged.BeforeAutoSpacing == null && sp.BeforeAutoSpacing != null) { merged.BeforeAutoSpacing = sp.BeforeAutoSpacing.Value; anySet = true; }
            if (merged.After == null && sp.After != null) { merged.After = sp.After.Value; anySet = true; }
            if (merged.AfterLines == null && sp.AfterLines != null) { merged.AfterLines = sp.AfterLines.Value; anySet = true; }
            if (merged.AfterAutoSpacing == null && sp.AfterAutoSpacing != null) { merged.AfterAutoSpacing = sp.AfterAutoSpacing.Value; anySet = true; }
            if (merged.Line == null && sp.Line != null) { merged.Line = sp.Line.Value; anySet = true; }
            if (merged.LineRule == null && sp.LineRule != null) { merged.LineRule = sp.LineRule.Value; anySet = true; }
        }

        // Resolve starting style: explicit styleId or document's default paragraph style.
        var startStyleId = styleId;
        if (startStyleId == null)
        {
            var defaultStyle = styles.Elements<Style>()
                .FirstOrDefault(s => s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
            startStyleId = defaultStyle?.StyleId?.Value;
        }

        // Walk basedOn chain derived → base, merging attributes not yet set.
        var visited = new HashSet<string>();
        var currentStyleId = startStyleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            MergeFrom(style.StyleParagraphProperties?.SpacingBetweenLines);
            currentStyleId = style.BasedOn?.Val?.Value;
        }

        // Final fallback: docDefaults pPrDefault — fills any attribute the
        // style chain left unset. Without this, a doc whose only spacing
        // declaration is in <w:pPrDefault> emits zero margin and the
        // before/after collapse computes incorrectly for adjacent paras.
        MergeFrom(styles.DocDefaults?.ParagraphPropertiesDefault
            ?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines);

        return anySet ? merged : null;
    }

    /// <summary>Resolve contextualSpacing from the style chain, with docDefaults fallback.</summary>
    private bool ResolveContextualSpacingFromStyle(string? styleId)
    {
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return false;

        var startStyleId = styleId;
        if (startStyleId == null)
        {
            var defaultStyle = styles.Elements<Style>()
                .FirstOrDefault(s => s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
            startStyleId = defaultStyle?.StyleId?.Value;
        }

        var visited = new HashSet<string>();
        var currentStyleId = startStyleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            var styleCs = IsContextualSpacingOn(style.StyleParagraphProperties?.ContextualSpacing);
            if (styleCs != null) return styleCs.Value;
            currentStyleId = style.BasedOn?.Val?.Value;
        }

        // Fallback: docDefaults pPrDefault.
        return IsContextualSpacingOn(styles.DocDefaults?.ParagraphPropertiesDefault
            ?.ParagraphPropertiesBaseStyle?.ContextualSpacing) ?? false;
    }

    /// <summary>
    /// Evaluate a w:contextualSpacing on/off toggle. Returns null when the element
    /// is absent (caller should fall back to the style chain); true when present
    /// and on (no val, or val=1/true/on); false when present with val=0/false/off.
    /// </summary>
    private static bool? IsContextualSpacingOn(ContextualSpacing? cs)
    {
        if (cs == null) return null;
        var v = cs.Val;
        return v == null || v.Value;
    }

    /// <summary>
    /// Effective left indent (pt) of a paragraph — direct w:ind/@w:left, else
    /// the style chain. Used by the positional-tab renderer so the first tab
    /// segment's box width compensates for the paragraph's left padding and the
    /// following text lands on the absolute tab position. Mirrors the indent
    /// resolution in GetParagraphInlineCss (direct ?? style).
    /// </summary>
    private double GetParagraphLeftIndentPt(Paragraph para)
    {
        var pProps = para.ParagraphProperties;
        var styleId = pProps?.ParagraphStyleId?.Val?.Value;
        var indLeft = pProps?.Indentation?.Left?.Value
            ?? ResolveIndentationFromStyle(styleId)?.Left?.Value;
        if (indLeft is string twips && twips != "0")
            return Units.TwipsToPt(twips);
        return 0;
    }

    /// <summary>
    /// Resolve Indentation from the style chain (basedOn walk).
    /// </summary>
    private Indentation? ResolveIndentationFromStyle(string? styleId)
    {
        // Attribute-level inheritance through basedOn (mirrors
        // ResolveSpacingFromStyle): each indentation attribute inherits
        // independently. A derived style overriding only `firstLine` must
        // still pick up `left`/`right`/`hanging` from its base.
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return null;

        if (styleId == null)
        {
            var defaultStyle = styles.Elements<Style>()
                .FirstOrDefault(s => s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
            return defaultStyle?.StyleParagraphProperties?.Indentation;
        }

        var merged = new Indentation();
        bool anySet = false;
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            var ind = style.StyleParagraphProperties?.Indentation;
            if (ind != null)
            {
                if (merged.Left == null && ind.Left != null) { merged.Left = ind.Left.Value; anySet = true; }
                if (merged.Right == null && ind.Right != null) { merged.Right = ind.Right.Value; anySet = true; }
                if (merged.FirstLine == null && ind.FirstLine != null) { merged.FirstLine = ind.FirstLine.Value; anySet = true; }
                if (merged.Hanging == null && ind.Hanging != null) { merged.Hanging = ind.Hanging.Value; anySet = true; }
                if (merged.Start == null && ind.Start != null) { merged.Start = ind.Start.Value; anySet = true; }
                if (merged.End == null && ind.End != null) { merged.End = ind.End.Value; anySet = true; }
                if (merged.LeftChars == null && ind.LeftChars != null) { merged.LeftChars = ind.LeftChars.Value; anySet = true; }
                if (merged.RightChars == null && ind.RightChars != null) { merged.RightChars = ind.RightChars.Value; anySet = true; }
                if (merged.FirstLineChars == null && ind.FirstLineChars != null) { merged.FirstLineChars = ind.FirstLineChars.Value; anySet = true; }
                if (merged.HangingChars == null && ind.HangingChars != null) { merged.HangingChars = ind.HangingChars.Value; anySet = true; }
            }
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return anySet ? merged : null;
    }

    /// <summary>
    /// Resolve paragraph CSS from style chain when no direct paragraph properties.
    /// </summary>
    private string ResolveParagraphStyleCss(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null)
        {
            // Fall back to default paragraph style (Normal)
            var defaultStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
            styleId = defaultStyle?.StyleId?.Value;
            if (styleId == null) return "";
        }

        var parts = new List<string>();
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = FindStyleById(currentStyleId);
            if (style == null) break;

            var pPr = style.StyleParagraphProperties;
            if (pPr != null)
            {
                var jc = pPr.Justification?.Val;
                if (jc != null && !parts.Any(p => p.StartsWith("text-align")))
                {
                    var align = jc.InnerText switch { "center" => "center", "right" or "end" => "right", "both" => "justify", _ => (string?)null };
                    if (align != null) parts.Add($"text-align:{align}");
                }

                var spacing = pPr.SpacingBetweenLines;
                if (spacing != null)
                {
                    // beforeLines/afterLines override before/after per
                    // ECMA-376 §17.3.1.33; "1 line" = 240 twips = 12pt fixed
                    // (matches Word's single-line spacing).
                    const double LineUnitPt = 12.0;
                    if (!parts.Any(p => p.StartsWith("margin-top")))
                    {
                        if (spacing.BeforeLines?.Value is int bl && bl != 0)
                            parts.Add($"margin-top:{bl / 100.0 * LineUnitPt:0.##}pt");
                        else if (spacing.Before?.Value is string b && b != "0")
                            parts.Add($"margin-top:{Units.TwipsToPt(b):0.##}pt");
                    }
                    if (!parts.Any(p => p.StartsWith("margin-bottom")))
                    {
                        if (spacing.AfterLines?.Value is int al && al != 0)
                            parts.Add($"margin-bottom:{al / 100.0 * LineUnitPt:0.##}pt");
                        else if (spacing.After?.Value is string a)
                            parts.Add($"margin-bottom:{Units.TwipsToPt(a):0.##}pt");
                    }
                    if (spacing.Line?.Value is string lv && !parts.Any(p => p.StartsWith("line-height")))
                    {
                        var rule = spacing.LineRule?.InnerText;
                        if ((rule == "auto" || rule == null) && int.TryParse(lv, out var val) && val > 0)
                        {
                            // OOXML §17.3.1.33 "auto" rule: see paragraph
                            // path above. line-height = (val/240) × natural_ratio.
                            var paraFont = ResolveParaFontForLineHeight(para);
                            var ratio = FontMetricsReader.GetRatio(paraFont);
                            parts.Add($"line-height:{ratio * (val / 240.0):0.####}");
                        }
                        else if (rule == "exact" || rule == "atLeast")
                        {
                            // §17.3.1.33 atLeast acts as a floor; use the
                            // paragraph natural single-line height when it
                            // exceeds the floor.
                            var linePt = Units.TwipsToPt(lv);
                            var emitPt = rule == "atLeast" ? ResolveAtLeastPt(linePt, para) : linePt;
                            parts.Add($"line-height:{emitPt:0.##}pt");
                            // exact pins the leading but keeps content visible;
                            // no overflow:hidden (would blank over-tall content —
                            // see content-loss note in the paragraph path above).
                        }
                    }
                }

                // Indentation
                var ind = pPr.Indentation;
                if (ind != null)
                {
                    if (ind.Left?.Value is string leftTwips && leftTwips != "0" && !parts.Any(p => p.StartsWith("margin-left")))
                        parts.Add($"margin-left:{Units.TwipsToPt(leftTwips):0.##}pt");
                    if (ind.Right?.Value is string rightTwips && rightTwips != "0" && !parts.Any(p => p.StartsWith("margin-right")))
                        parts.Add($"margin-right:{Units.TwipsToPt(rightTwips):0.##}pt");
                    if (ind.FirstLine?.Value is string fl && fl != "0" && !parts.Any(p => p.StartsWith("text-indent")))
                        parts.Add($"text-indent:{Units.TwipsToPt(fl):0.##}pt");
                    if (ind.Hanging?.Value is string hg && hg != "0" && !parts.Any(p => p.StartsWith("text-indent")))
                        parts.Add($"text-indent:-{Units.TwipsToPt(hg):0.##}pt");
                }

                var shadingFill = ResolveShadingFill(pPr.Shading);
                if (shadingFill != null && !parts.Any(p => p.StartsWith("background")))
                    parts.Add($"background-color:{shadingFill}");
            }

            currentStyleId = style.BasedOn?.Val?.Value;
        }

        // docDefaults pPrDefault fallback: when the entire style chain left
        // spacing/indent unset, pick up <w:pPrDefault> values. Without this,
        // a paragraph with no <w:pPr> in a doc whose only spacing source is
        // pPrDefault (typical of synthetic / cli-authored docs) emits zero
        // margin-bottom and the next paragraph's spaceBefore-vs-prev.spaceAfter
        // collapse computes incorrectly.
        var defPPr = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
            ?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        if (defPPr != null)
        {
            const double LineUnitPt = 12.0;
            var spacing = defPPr.SpacingBetweenLines;
            if (spacing != null)
            {
                if (!parts.Any(p => p.StartsWith("margin-top")))
                {
                    if (spacing.BeforeLines?.Value is int bl && bl != 0)
                        parts.Add($"margin-top:{bl / 100.0 * LineUnitPt:0.##}pt");
                    else if (spacing.Before?.Value is string b && b != "0")
                        parts.Add($"margin-top:{Units.TwipsToPt(b):0.##}pt");
                }
                if (!parts.Any(p => p.StartsWith("margin-bottom")))
                {
                    if (spacing.AfterLines?.Value is int al && al != 0)
                        parts.Add($"margin-bottom:{al / 100.0 * LineUnitPt:0.##}pt");
                    else if (spacing.After?.Value is string a)
                        parts.Add($"margin-bottom:{Units.TwipsToPt(a):0.##}pt");
                }
                if (spacing.Line?.Value is string lv && !parts.Any(p => p.StartsWith("line-height")))
                {
                    var rule = spacing.LineRule?.InnerText;
                    if ((rule == "auto" || rule == null) && int.TryParse(lv, out var val) && val > 0)
                    {
                        // OOXML §17.3.1.33 "auto" rule (see paragraph path
                        // above). line-height = (val/240) × natural_ratio.
                        var paraFont = ResolveParaFontForLineHeight(para);
                        var ratio = FontMetricsReader.GetRatio(paraFont);
                        parts.Add($"line-height:{ratio * (val / 240.0):0.####}");
                    }
                    else if (rule == "exact" || rule == "atLeast")
                    {
                        // §17.3.1.33 atLeast: floor only; substitute natural
                        // line-height when the paragraph's content exceeds it.
                        var linePt = Units.TwipsToPt(lv);
                        var emitPt = rule == "atLeast" ? ResolveAtLeastPt(linePt, para) : linePt;
                        parts.Add($"line-height:{emitPt:0.##}pt");
                        // exact pins the leading but keeps content visible;
                        // no overflow:hidden (would blank over-tall content —
                        // see content-loss note in the paragraph path above).
                    }
                }
            }
            var ind = defPPr.Indentation;
            if (ind != null)
            {
                if (ind.Left?.Value is string leftTwips && leftTwips != "0" && !parts.Any(p => p.StartsWith("margin-left")))
                    parts.Add($"margin-left:{Units.TwipsToPt(leftTwips):0.##}pt");
                if (ind.Right?.Value is string rightTwips && rightTwips != "0" && !parts.Any(p => p.StartsWith("margin-right")))
                    parts.Add($"margin-right:{Units.TwipsToPt(rightTwips):0.##}pt");
                if (ind.FirstLine?.Value is string fl && fl != "0" && !parts.Any(p => p.StartsWith("text-indent")))
                    parts.Add($"text-indent:{Units.TwipsToPt(fl):0.##}pt");
                if (ind.Hanging?.Value is string hg && hg != "0" && !parts.Any(p => p.StartsWith("text-indent")))
                    parts.Add($"text-indent:-{Units.TwipsToPt(hg):0.##}pt");
            }
        }

        return string.Join(";", parts);
    }

    /// <summary>Apply OOXML §17.3.1.33 atLeast semantics: the value is a
    /// floor, not a fixed line-height. When the paragraph's natural
    /// single-line height (font ratio × principal size) exceeds the
    /// floor, the natural value is used instead.</summary>
    private double ResolveAtLeastPt(double floorPt, Paragraph para)
    {
        var paraFont = ResolveParaFontForLineHeight(para);
        var ratio = FontMetricsReader.GetRatio(paraFont);
        var paraSizePt = ResolveParaPrincipalSizePt(para) ?? 11.0;
        return Math.Max(floorPt, ratio * paraSizePt);
    }

    /// <summary>Effective spacing-after for a paragraph in pt, cascading
    /// through direct pPr → pStyle chain → docDefaults pPrDefault. Returns
    /// 0 when nothing in the cascade sets it.</summary>
    private double ResolveParaAfterSpacingPt(Paragraph para)
    {
        var pProps = para.ParagraphProperties;
        var styleId = pProps?.ParagraphStyleId?.Val?.Value;
        var styleSpacing = ResolveSpacingFromStyle(styleId);
        var afterTwips = pProps?.SpacingBetweenLines?.After?.Value
                         ?? styleSpacing?.After?.Value;
        if (afterTwips == null)
        {
            afterTwips = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults
                ?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle
                ?.SpacingBetweenLines?.After?.Value;
        }
        if (afterTwips != null && int.TryParse(afterTwips, out var tw))
            return tw / 20.0;
        return 0;
    }

    /// <summary>Read the paragraph's principal font size (in pt), the same
    /// value GetParagraphInlineCss emits on the &lt;p&gt; element.</summary>
    private double? ResolveParaPrincipalSizePt(Paragraph para)
    {
        Run? probeRun = para.Elements<Run>().FirstOrDefault(r =>
            r.ChildElements.Any(c => c is Text t && !string.IsNullOrEmpty(t.Text)));
        if (probeRun == null)
        {
            var markProps = para.ParagraphProperties?.ParagraphMarkRunProperties;
            if (markProps != null)
            {
                var synthRPr = new RunProperties();
                foreach (var child in markProps.ChildElements)
                    synthRPr.AppendChild(child.CloneNode(true));
                probeRun = new Run(synthRPr);
            }
        }
        if (probeRun == null) return null;
        var rProps = ResolveEffectiveRunProperties(probeRun, para);
        var sz = rProps.FontSize?.Val?.Value;
        if (sz != null && int.TryParse(sz, out var hp))
            return hp / 2.0;
        return null;
    }

    /// <summary>Compute a line-height CSS value for a single run-level span.
    /// CSS 2.1 §10.8.1: line-box height = max over each inline of its own
    /// line-height. When the run's font-size matches the paragraph's
    /// principal size, the run emits the unitless multiplier 1 so the
    /// paragraph's own line-height dominates the line-box; this also caps
    /// the inline box at the run's font-size, preventing a font variant
    /// whose intrinsic inline metrics exceed the paragraph's line-height
    /// from extending the line-box. When the run's size differs from the
    /// paragraph's principal size, the run mirrors the paragraph's
    /// line-height rule using its own font's natural ratio so mixed-size
    /// paragraphs render at the correct max-line-height per OOXML
    /// §17.3.1.33.</summary>
    private string ResolveRunLineHeightCss(string? runFontName, double? runSizePt, Paragraph para)
    {
        var pProps = para.ParagraphProperties;
        var styleId = pProps?.ParagraphStyleId?.Val?.Value;
        var styleSpacing = ResolveSpacingFromStyle(styleId);
        var hasSpacing = pProps?.SpacingBetweenLines != null || styleSpacing != null;
        var lineVal = pProps?.SpacingBetweenLines?.Line?.Value ?? styleSpacing?.Line?.Value;
        var rule = pProps?.SpacingBetweenLines?.LineRule?.InnerText ?? styleSpacing?.LineRule?.InnerText;

        // §17.3.1.33 exact: the paragraph pins the line box to a fixed height
        // and clips over-tall glyphs (the paragraph path emits the fixed
        // line-height + overflow:hidden). The run span must NOT emit its own
        // line-height — line-height:1 on an over-tall run resolves to the
        // run's font-size and would defeat the exact box. Inherit instead so
        // the fixed value dominates regardless of run-vs-paragraph size match.
        if (rule == "exact" && lineVal != null)
            return "line-height:inherit";

        var paraSizePt = ResolveParaPrincipalSizePt(para);
        bool sizeMatches = runSizePt == null
            || (paraSizePt != null && Math.Abs(runSizePt.Value - paraSizePt.Value) < 0.01);
        if (sizeMatches) return "line-height:1";

        var font = runFontName ?? ResolveParaFontForLineHeight(para);
        var ratio = FontMetricsReader.GetRatio(font);

        if (hasSpacing)
        {
            if (lineVal != null)
            {
                if ((rule == "auto" || rule == null)
                    && int.TryParse(lineVal, out var lvNum) && lvNum > 0)
                    return $"line-height:{ratio * (lvNum / 240.0):0.####}";
                // rule == "exact" handled at top (inherit, paragraph clips).
                if (rule == "atLeast")
                {
                    // §17.3.1.33 atLeast: floor; this run's natural single
                    // line height (ratio × runSize) substitutes when greater.
                    var floorPt = Units.TwipsToPt(lineVal);
                    var runNaturalPt = ratio * runSizePt!.Value;
                    return $"line-height:{Math.Max(floorPt, runNaturalPt):0.##}pt";
                }
            }
            return $"line-height:{ratio:0.####}";
        }

        var builtIn = ResolveBuiltInStyleDefaults(styleId);
        if (builtIn == null && DocCarriesNormalDefaults())
            builtIn = BuiltInStyleDefaults["Normal"];
        if (builtIn != null)
            return $"line-height:{Math.Max(builtIn.Line, ratio):0.####}";
        return $"line-height:{ratio:0.####}";
    }

    private string GetRunInlineCss(RunProperties? rProps, Paragraph? para = null)
    {
        if (rProps == null) return "";
        var parts = new List<string>();

        // Font
        var fonts = rProps.RunFonts;
        // CS slot priority for RTL runs (Arabic / Hebrew). When the run is
        // tagged <w:rtl/>, ComplexScript is the script-correct face — without
        // this, ar/he runs that only carry rFonts/@w:cs (the LocaleFontRegistry
        // default for ar="Arabic Typesetting") rendered in the body's default
        // Latin font. EA-priority is preserved for the default LTR path so CJK
        // runs continue to read rFonts/@w:eastAsia.
        var isRtlRun = rProps.RightToLeftText != null
            && (rProps.RightToLeftText.Val == null || rProps.RightToLeftText.Val.Value);
        // Plain rFonts attributes win when present; otherwise resolve the
        // matching *Theme attribute against theme1.xml. This is what
        // styles like Title (rFonts asciiTheme="majorHAnsi") rely on —
        // without it the run silently falls back to the body default.
        var font = isRtlRun
            ? (fonts?.ComplexScript?.Value ?? ResolveThemeFont(fonts?.ComplexScriptTheme?.InnerText)
               ?? fonts?.Ascii?.Value ?? ResolveThemeFont(fonts?.AsciiTheme?.InnerText)
               ?? fonts?.HighAnsi?.Value ?? ResolveThemeFont(fonts?.HighAnsiTheme?.InnerText))
            : (fonts?.EastAsia?.Value ?? ResolveThemeFont(fonts?.EastAsiaTheme?.InnerText)
               ?? fonts?.Ascii?.Value ?? ResolveThemeFont(fonts?.AsciiTheme?.InnerText)
               ?? fonts?.HighAnsi?.Value ?? ResolveThemeFont(fonts?.HighAnsiTheme?.InnerText));
        // Skip the legacy "+mn-lt" / "+mj-ea" shorthand syntax (rare, predates
        // the typed *Theme attributes — and the typed path above already
        // handled the modern equivalent). Also skip when the resolved font
        // matches the document default — body-level CSS already declares
        // font-family there, so duplicating it on every run span only bloats
        // the HTML and obscures real per-run overrides.
        // Complex-script slot (cs/csTheme). On the LTR path the primary `font`
        // above never reads it, so a run that carries a cs face (Arabic / Hebrew
        // typesetting) for embedded RTL spans dropped that face entirely. Resolve
        // it separately and append it as a fallback after the primary/Latin faces
        // so the browser uses it for complex-script glyphs the others lack. LTR
        // runs only; the RTL path already resolves cs as the primary `font`.
        var csFont = isRtlRun ? null
            : (fonts?.ComplexScript?.Value ?? ResolveThemeFont(fonts?.ComplexScriptTheme?.InnerText));
        if (font != null
            && !font.StartsWith("+", StringComparison.Ordinal)
            && !string.Equals(font, ReadDocDefaults().Font, StringComparison.Ordinal))
        {
            var fallback = GetChineseFontFallback(font);
            // Always append a generic family so the run still renders with the right
            // serif/sans-serif class when neither the primary nor the CJK fallback
            // is installed (matters in headless browsers like Playwright).
            var generic = GenericFontFamily(font);
            // Latin slot (ascii/hAnsi). When a run carries BOTH a Latin face and a
            // distinct EastAsia face, Word renders ASCII with the Latin face and
            // CJK with the EastAsia face. The EA-priority resolution above picked
            // the EastAsia face as `font`, dropping the Latin one — prepend it so
            // the browser uses Latin first and falls back to EastAsia (+ its CJK
            // chain) for glyphs the Latin face lacks. LTR runs only; the RTL path
            // already resolves CS/Latin and never wants an EastAsia prefix.
            var latinFont = isRtlRun ? null
                : (fonts?.Ascii?.Value ?? ResolveThemeFont(fonts?.AsciiTheme?.InnerText)
                   ?? fonts?.HighAnsi?.Value ?? ResolveThemeFont(fonts?.HighAnsiTheme?.InnerText));
            var latinPrefix = (latinFont != null
                && !latinFont.StartsWith("+", StringComparison.Ordinal)
                && !string.Equals(latinFont, font, StringComparison.Ordinal))
                ? $"'{CssSanitize(latinFont)}',"
                : "";
            var csSuffix = (csFont != null
                && !csFont.StartsWith("+", StringComparison.Ordinal)
                && !string.Equals(csFont, font, StringComparison.Ordinal)
                && !string.Equals(csFont, latinFont, StringComparison.Ordinal))
                ? $",'{CssSanitize(csFont)}'"
                : "";
            // Latin-led run (distinct ascii face, EastAsia kept only as the
            // CJK-glyph provider): insert the synth-bold-capable generic right
            // after the Latin face so that when the Latin font isn't installed
            // (headless/Playwright), the browser reaches the generic — which
            // synthesizes bold — BEFORE the EastAsia face. EA faces like
            // "MS PGothic" carry no bold instance AND block synthetic bold, so
            // when they lead the Latin glyphs they silently neutralize a run's
            // <w:b/>. The EastAsia font + its CJK fallback chain still trail the
            // generic, so CJK glyphs (absent from Latin/generic) continue to
            // resolve to the EA face per CSS per-glyph matching. Pure CJK runs
            // (no distinct Latin face → empty latinPrefix) keep the old order:
            // EA leads, generic last — unchanged.
            var latinLedGeneric = latinPrefix.Length > 0 ? $"{generic}," : "";
            parts.Add(fallback != null
                ? $"font-family:{latinPrefix}{latinLedGeneric}'{CssSanitize(font)}',{fallback}{csSuffix},{generic}"
                : $"font-family:{latinPrefix}{latinLedGeneric}'{CssSanitize(font)}'{csSuffix},{generic}");
        }
        else if (csFont != null
            && !csFont.StartsWith("+", StringComparison.Ordinal)
            && !string.Equals(csFont, ReadDocDefaults().Font, StringComparison.Ordinal))
        {
            // cs-only LTR run (no Latin/EastAsia slot resolved a non-default face):
            // the complex-script face is the only one declared, so it leads the
            // stack. Without this the span emitted no font-family at all.
            var generic = GenericFontFamily(csFont);
            parts.Add($"font-family:'{CssSanitize(csFont)}',{generic}");
        }

        // Size (stored as half-points)
        var size = rProps.FontSize?.Val?.Value;
        if (size != null && int.TryParse(size, out var halfPts))
            parts.Add($"font-size:{halfPts / 2.0:0.##}pt");

        // Bold (w:b with no val or val="true"/"1" means bold; val="false"/"0" means not bold)
        if (rProps.Bold != null && (rProps.Bold.Val == null || rProps.Bold.Val.Value))
            parts.Add("font-weight:bold");

        // Italic (same logic as bold)
        if (rProps.Italic != null && (rProps.Italic.Val == null || rProps.Italic.Val.Value))
            parts.Add("font-style:italic");

        // Underline: map OOXML variants to CSS text-decoration-style / thickness.
        // OOXML vals: single, double, thick, dotted, dottedHeavy, dash, dashedHeavy,
        //   dashLong, dashLongHeavy, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy,
        //   wave, wavyHeavy, wavyDouble, words, none
        if (rProps.Underline?.Val != null)
        {
            var ulVal = rProps.Underline.Val.InnerText;
            if (ulVal != "none")
            {
                parts.Add("text-decoration:underline");
                // Map to text-decoration-style
                string? style = ulVal switch
                {
                    "double" or "wavyDouble" => "double",
                    "dotted" or "dottedHeavy" => "dotted",
                    "dash" or "dashedHeavy" or "dashLong" or "dashLongHeavy"
                        or "dotDash" or "dotDashHeavy" or "dotDotDash" or "dotDotDashHeavy" => "dashed",
                    "wave" or "wavyHeavy" => "wavy",
                    _ => null,
                };
                if (style != null)
                    parts.Add($"text-decoration-style:{style}");
                // Thickness: "thick" and any *Heavy variant
                if (ulVal == "thick" || (ulVal?.EndsWith("Heavy") ?? false))
                    parts.Add("text-decoration-thickness:2px");
                // Per-underline color via w:u w:color="RRGGBB"
                var ulColor = rProps.Underline.Color?.Value;
                if (!string.IsNullOrEmpty(ulColor) && !ulColor.Equals("auto", StringComparison.OrdinalIgnoreCase)
                    && IsHexColor(ulColor))
                    parts.Add($"text-decoration-color:#{ulColor}");
            }
            else
            {
                // Explicit w:u val="none" must suppress the inherited underline.
                // Inside a hyperlink the run is wrapped in an <a>, whose UA
                // default underline persists unless the span overrides it, so
                // emit text-decoration:none to match real Word (no underline).
                parts.Add("text-decoration:none");
            }
        }

        // Strikethrough (single or double)
        var hasSingleStrike = rProps.Strike != null && (rProps.Strike.Val == null || rProps.Strike.Val.Value);
        var hasDoubleStrike = rProps.DoubleStrike != null && (rProps.DoubleStrike.Val == null || rProps.DoubleStrike.Val.Value);
        if (hasSingleStrike || hasDoubleStrike)
        {
            var existing = parts.FirstOrDefault(p => p.StartsWith("text-decoration:"));
            if (existing != null && existing != "text-decoration:none")
            {
                parts.Remove(existing);
                parts.Add(existing + " line-through");
            }
            else
            {
                // "text-decoration:none" (explicit underline off) must not be
                // concatenated into "none line-through" (invalid); replace it.
                if (existing == "text-decoration:none") parts.Remove(existing);
                parts.Add("text-decoration:line-through");
            }
            // Double-strike renders via text-decoration-style: double (CSS3, broad support)
            if (hasDoubleStrike)
                parts.Add("text-decoration-style:double");
        }

        // Character spacing (w:spacing val in twips = 1/20 pt, can be negative)
        if (rProps.Spacing?.Val?.HasValue == true)
        {
            var sp = rProps.Spacing.Val.Value;
            if (sp != 0)
                parts.Add($"letter-spacing:{sp / 20.0:0.##}pt");
        }

        // Character scale (w:w, horizontal stretch as a percentage). Render via a bare
        // transform:scaleX — NOT display:inline-block. CSS transforms are paint-only and
        // never affect layout, so inline-block bought no width reservation here (the scaled
        // glyphs overflow the box at any ratio != 1 regardless of display); its only effect
        // was to shrink-wrap the run and trim its leading/trailing whitespace. inline-block
        // boxes drop trailing whitespace at the box edge (Chromium hangs it with zero
        // advance, irrespective of white-space:pre/pre-wrap/break-spaces), so a run ending
        // in a space ("Once the ministry has ") butted against the next run and rendered as
        // "hasreviewed". A bare inline transform keeps the run inline, so its own boundary
        // spaces survive and word gaps are preserved; overflow at large w:w is equal to or
        // milder than the inline-block path (the inline space buffers the next run). Only the
        // w:w branch is touched — unscaled runs are untouched. Default/unit 100% → skip.
        var charScale = rProps.CharacterScale?.Val?.Value;
        if (charScale.HasValue && charScale.Value > 0 && charScale.Value != 100)
        {
            var ratio = charScale.Value / 100.0;
            parts.Add($"transform:scaleX({ratio:0.##});transform-origin:left");
        }

        // Color: w:color val + themeColor with tint/shade. Route through
        // ResolveRunColor for consistency with conditional-format and border
        // paths. Val wins if not "auto"; else fall through to themeColor.
        var resolvedColor = ResolveRunColor(rProps.Color);
        if (resolvedColor != null)
        {
            parts.Add($"color:{resolvedColor}");
        }
        else
        {
            // No explicit/theme run color → Word's automatic color: pick black
            // or white by the run's effective background luminance. Word renders
            // color=auto text as white on a dark fill (the deep-blue title bars
            // in this corpus) and black on light/no fill. The browser default is
            // unconditional black, so without this the title bars read black-on-
            // dark-blue. Only auto runs are touched; explicit black stays black.
            var bgHex = ResolveEffectiveBackgroundForRun(rProps, para);
            // White (or absent) backdrop → black text: this is the prior
            // behavior, so don't emit a redundant color for the common case.
            if (bgHex != null && IsColorDark(bgHex))
                parts.Add("color:#FFFFFF");
            // R102-2: an EXPLICIT run-level w:color val="auto" (Word's
            // "Automatic" = black on a light backdrop) must beat any inherited
            // color — notably the global `a { color:#2B579A }` rule when this
            // run lives inside a <w:hyperlink> with rStyle="Hyperlink". OOXML
            // direct run color (incl. auto) wins over the rStyle character
            // style's color, so the email link renders black like Word, not the
            // style's blue. Emitting nothing here would let the ancestor <a>
            // blue leak through. Only fire on an explicit auto value: a plain
            // Hyperlink run with NO run-level color (rProps.Color == null) keeps
            // the inherited blue; the dark-bg reverse-video branch above already
            // claimed the white case.
            else if (rProps.Color?.Val?.Value == "auto")
                parts.Add("color:#000000");
        }

        // Highlight
        var highlight = rProps.Highlight?.Val?.InnerText;
        if (highlight != null)
        {
            var hlColor = HighlightToCssColor(highlight);
            if (hlColor != null) parts.Add($"background-color:{hlColor}");
        }

        // Superscript / Subscript per OOXML §17.3.2.42: the surrounding
        // line-box keeps the base-font height; only the affected run's glyph
        // is repositioned. Browser default vertical-align:super/sub raises
        // the inline box which participates in line-height aggregation and
        // expands the parent line; position:relative shifts the visual glyph
        // without affecting line geometry. Reduced font size and baseline
        // offset come from the run's own font (OS/2 ySub/SuperscriptYSize
        // and ySub/SuperscriptYOffset); a font-agnostic fallback applies
        // when those fields are unreadable.
        var vertAlign = rProps.VerticalTextAlignment?.Val;
        if (vertAlign != null)
        {
            var ssFont = font ?? (para != null ? ResolveParaFontForLineHeight(para) : null);
            var ss = ssFont != null
                ? FontMetricsReader.GetSuperSubMetrics(ssFont)
                : default;
            if (vertAlign.InnerText == "superscript")
            {
                if (!ss.IsEmpty && ss.SuperSizeEm > 0 && ss.SuperOffsetEm > 0)
                    parts.Add($"vertical-align:baseline;position:relative;top:-{ss.SuperOffsetEm:0.###}em;font-size:{ss.SuperSizeEm * 100:0.#}%");
                else
                    parts.Add("vertical-align:baseline;position:relative;top:-0.35em;font-size:smaller");
            }
            else if (vertAlign.InnerText == "subscript")
            {
                if (!ss.IsEmpty && ss.SubSizeEm > 0 && ss.SubOffsetEm > 0)
                    parts.Add($"vertical-align:baseline;position:relative;top:{ss.SubOffsetEm:0.###}em;font-size:{ss.SubSizeEm * 100:0.#}%");
                else
                    parts.Add("vertical-align:baseline;position:relative;top:0.15em;font-size:smaller");
            }
        }

        // w:position (OOXML §17.3.2.24) — "raised/lowered text by N points"
        // character property, distinct from super/subscript: the glyph is
        // shifted vertically WITHOUT changing the font size. Val is in
        // HALF-POINTS, positive = raised, negative = lowered. Unlike
        // super/subscript (which intentionally keeps the line box fixed),
        // Word EXPANDS the line height to contain raised/lowered text so it
        // doesn't overlap adjacent paragraphs. CSS vertical-align with a
        // length shifts the inline box AND grows the line box to contain it,
        // matching Word. (position:relative shifts only the glyph and leaves
        // the line box at base height, causing overlap.) val/2 = pt; positive
        // raises (positive vertical-align), negative lowers.
        var posVal = rProps.Position?.Val?.Value;
        if (!string.IsNullOrEmpty(posVal) && int.TryParse(posVal, out var posHalfPt) && posHalfPt != 0)
        {
            // Use position:relative;top: rather than vertical-align:<length>:
            // a length-valued vertical-align expands the inline line box by the
            // shift amount (an 8pt raise on 11pt text grows the row to ~1000px in
            // some browsers), which doesn't match Word's rendering. Matches the
            // super/sub handling above. positive posHalfPt = raise → negative top.
            var offsetPt = Math.Abs(posHalfPt) / 2.0;
            var sign = posHalfPt > 0 ? "-" : "";
            parts.Add($"position:relative;top:{sign}{offsetPt:0.###}pt");
        }

        // SmallCaps / AllCaps
        if (rProps.SmallCaps != null && (rProps.SmallCaps.Val == null || rProps.SmallCaps.Val.Value))
            parts.Add("font-variant:small-caps");
        if (rProps.Caps != null && (rProps.Caps.Val == null || rProps.Caps.Val.Value))
            parts.Add("text-transform:uppercase");

        // Run shading (w:shd) — background color on text (e.g. inverse video)
        var runShd = rProps.Shading;
        if (runShd != null && highlight == null) // don't override highlight
        {
            // val=solid → 100% foreground (w:color, black if absent); fill is hidden.
            if (runShd.Val != null && runShd.Val.Value == ShadingPatternValues.Solid)
            {
                var color = runShd.Color?.Value;
                parts.Add(color != null && color != "auto" && IsHexColor(color)
                    ? $"background-color:#{color}"
                    : "background-color:#000000");
            }
            else
            {
                var fill = runShd.Fill?.Value;
                if (fill != null && fill != "auto" && IsHexColor(fill))
                    parts.Add($"background-color:#{fill}");
            }
        }

        // Run border (w:bdr) — border around text (e.g. "box" text)
        var runBdr = rProps.GetFirstChild<Border>();
        if (runBdr != null)
        {
            var bdrVal = runBdr.Val?.InnerText;
            if (bdrVal != null && bdrVal != "none" && bdrVal != "nil")
            {
                var bdrSz = runBdr.Size?.Value ?? 4;
                var bdrColor = runBdr.Color?.Value;
                var px = Math.Max(1, bdrSz / 8.0);
                var color = (bdrColor != null && bdrColor != "auto" && IsHexColor(bdrColor)) ? $"#{bdrColor}" : "#000";
                parts.Add($"border:{px:0.#}px solid {color};padding:0 2px");
            }
        }

        // RTL text direction — use unicode-bidi:embed so Arabic/Hebrew
        // contextual shaping + Unicode BiDi algorithm still apply.
        // bidi-override would force reversal, corrupting Arabic glyph order.
        if (rProps.RightToLeftText != null && (rProps.RightToLeftText.Val == null || rProps.RightToLeftText.Val.Value))
        {
            parts.Add("direction:rtl;unicode-bidi:embed");
        }
        else if (para?.ParagraphProperties?.BiDi is { } paraBiDi
            && (paraBiDi.Val == null || paraBiDi.Val.Value))
        {
            // LTR run inside an RTL paragraph (e.g. "100 USD" embedded in
            // Arabic): the paragraph's direction:rtl base would let the
            // browser's BiDi algorithm split a "number space letters"
            // sequence across the line ("100 ... USD"). unicode-bidi:isolate
            // pins the LTR run as a single self-contained directional island
            // so it renders left-to-right as one unit, symmetric to the
            // embed treatment given to RTL runs above. Only emitted in the
            // RTL-paragraph context — a plain LTR paragraph needs no extra
            // direction declaration on its runs.
            parts.Add("direction:ltr;unicode-bidi:isolate");
        }

        // East Asian emphasis mark (w:em val=dot/comma/circle/underDot)
        // → CSS text-emphasis-style, widely supported (including -webkit- prefix)
        var emVal = rProps.Emphasis?.Val?.InnerText;
        if (emVal != null && emVal != "none")
        {
            string css = emVal switch
            {
                "dot" => "filled dot",
                "comma" => "filled sesame",
                "circle" => "filled circle",
                "underDot" => "filled dot",
                _ => "filled",
            };
            var pos = emVal == "underDot" ? "under" : "over";
            parts.Add($"text-emphasis:{css};text-emphasis-position:{pos};-webkit-text-emphasis:{css};-webkit-text-emphasis-position:{pos}");
        }

        // w14 text effects (textFill, textOutline, glow, shadow, reflection)
        AppendW14CssEffects(rProps, parts);

        // CSS 2.1 §10.8.1 — line-box height = max over each inline of
        // its own line-height. ResolveRunLineHeightCss picks "1" when the
        // run's font-size matches the paragraph's principal size (the
        // paragraph's own line-height dominates the line-box, and the
        // run's inline box is capped to its font-size so a heavier font
        // variant can't extend it); when the run's size differs, the
        // run's line-height mirrors the paragraph's rule using its own
        // font's natural ratio so the bigger inline box drives the
        // line-box to max-of-fonts × max-size × multi.
        double? runSizePt = (size != null && int.TryParse(size, out var hp))
            ? hp / 2.0 : (double?)null;
        if (parts.Count > 0 && para != null)
            parts.Add(ResolveRunLineHeightCss(font, runSizePt, para));

        return string.Join(";", parts);
    }

    private static string HexToRgba(string hexColor, double opacity)
    {
        if (hexColor.Length == 7 && int.TryParse(hexColor.AsSpan(1),
            System.Globalization.NumberStyles.HexNumber, null, out var rgb))
            return $"rgba({(rgb >> 16) & 0xFF},{(rgb >> 8) & 0xFF},{rgb & 0xFF},{opacity:0.##})";
        return hexColor;
    }

    private static void AppendW14CssEffects(RunProperties rProps, List<string> parts)
    {
        var textShadows = new List<string>();

        foreach (var child in rProps.ChildElements)
        {
            if (child.NamespaceUri != W14Ns) continue;

            switch (child.LocalName)
            {
                case "textFill":
                {
                    var innerXml = child.InnerXml;
                    if (innerXml.Contains("gradFill"))
                    {
                        var colors = new List<string>();
                        foreach (System.Text.RegularExpressions.Match m in
                            System.Text.RegularExpressions.Regex.Matches(innerXml, @"val=""([0-9A-Fa-f]{6})"""))
                            colors.Add($"#{m.Groups[1].Value}");

                        if (colors.Count >= 2)
                        {
                            var isRadial = innerXml.Contains("<w14:path");
                            var angleMatch = System.Text.RegularExpressions.Regex.Match(innerXml, @"ang=""(\d+)""");
                            var angle = angleMatch.Success ? int.Parse(angleMatch.Groups[1].Value) / 60000.0 : 0.0;

                            parts.RemoveAll(p => p.StartsWith("color:"));

                            if (isRadial)
                            {
                                // CONSISTENCY(radial-gradient-extent): closest-side so gradient reaches shape edge (matches PPTX R2 fix).
                                parts.Add($"background:radial-gradient(circle closest-side,{colors[0]},{colors[1]})");
                            }
                            else
                            {
                                // OOXML: 0°=left→right, 90°=top→bottom
                                // CSS:   0°=bottom→top,  90°=left→right, 180°=top→bottom
                                var cssAngle = angle + 90;
                                parts.Add($"background:linear-gradient({cssAngle:0.##}deg,{colors[0]},{colors[1]})");
                            }
                            parts.Add("-webkit-background-clip:text");
                            parts.Add("background-clip:text");
                            parts.Add("-webkit-text-fill-color:transparent");
                        }
                        else if (colors.Count == 1)
                        {
                            parts.RemoveAll(p => p.StartsWith("color:"));
                            parts.Add($"color:{colors[0]}");
                        }
                    }
                    else if (innerXml.Contains("solidFill"))
                    {
                        var colorMatch = System.Text.RegularExpressions.Regex.Match(
                            innerXml, @"val=""([0-9A-Fa-f]{6})""");
                        if (colorMatch.Success)
                        {
                            parts.RemoveAll(p => p.StartsWith("color:"));
                            parts.Add($"color:#{colorMatch.Groups[1].Value}");
                        }
                    }
                    break;
                }
                case "textOutline":
                {
                    var wAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
                    var widthEmu = long.TryParse(wAttr.Value, out var w) ? w : 0;
                    var widthPt = Math.Max(0.5, widthEmu / EmuConverter.EmuPerPointF);
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? $"#{colorMatch.Groups[1].Value}" : "currentColor";
                    parts.Add($"-webkit-text-stroke:{widthPt:0.##}pt {color}");
                    break;
                }
                case "shadow":
                {
                    var attrs = child.GetAttributes().ToDictionary(a => a.LocalName, a => a.Value);
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? $"#{colorMatch.Groups[1].Value}" : "#000000";
                    var blurEmu = attrs.TryGetValue("blurRad", out var br) && long.TryParse(br, out var blurVal) ? blurVal : 0;
                    // Word renders w14:shadow on body text far more subtly than a
                    // literal EMU→px translation suggests: the default preset
                    // (blurRad=38100, dist=19050, dk1 @ full alpha) is barely
                    // visible behind glyphs, not the heavy "1.4px 1.4px 4px"
                    // smudge a direct mapping produces. Cap offset/blur to small
                    // values and clamp opacity low so a document full of default
                    // shadows reads clean instead of dirty/embossed.
                    var blurPx = Math.Min(blurEmu / EmuConverter.EmuPerPointF * 1.333, 1.5);
                    var distEmu = attrs.TryGetValue("dist", out var dist) && long.TryParse(dist, out var distLong) ? distLong : 0;
                    var dirVal = attrs.TryGetValue("dir", out var dir) && long.TryParse(dir, out var dirLong) ? dirLong : 0;
                    var angleRad = dirVal / 60000.0 * Math.PI / 180.0;
                    var distPx = Math.Min(distEmu / EmuConverter.EmuPerPointF * 1.333, 0.8);
                    var xPx = distPx * Math.Sin(angleRad);
                    var yPx = distPx * Math.Cos(angleRad);
                    var alphaMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"alpha[^>]*val=""(\d+)""");
                    // Author alpha (if any) is a *ceiling*; Word never shows the
                    // body-text shadow at full strength, so clamp to <=0.30.
                    var shadowAlpha = alphaMatch.Success && double.TryParse(alphaMatch.Groups[1].Value, out var alphaVal)
                        ? alphaVal / 100000.0
                        : 1.0;
                    color = HexToRgba(color, Math.Min(shadowAlpha, 0.30));
                    textShadows.Add($"{xPx:0.#}px {yPx:0.#}px {blurPx:0.#}px {color}");
                    break;
                }
                case "glow":
                {
                    var radAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "rad");
                    var radiusEmu = long.TryParse(radAttr.Value, out var r) ? r : 0;
                    var radiusPx = radiusEmu / EmuConverter.EmuPerPointF * 1.333;
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? $"#{colorMatch.Groups[1].Value}" : "#000000";
                    var alphaMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"alpha[^>]*val=""(\d+)""");
                    var alpha = alphaMatch.Success && double.TryParse(alphaMatch.Groups[1].Value, out var av) ? av / 100000.0 : 1.0;
                    // Multiple stacked text-shadow layers to approximate Word glow spread
                    // Word glow is a soft halo that extends from text edges; simulate with
                    // tight + medium + wide shadow layers at decreasing opacity
                    var c1 = HexToRgba(color, Math.Min(1.0, alpha * 0.9));
                    var c2 = HexToRgba(color, Math.Min(1.0, alpha * 0.8));
                    var c3 = HexToRgba(color, Math.Min(1.0, alpha * 0.5));
                    var c4 = HexToRgba(color, Math.Min(1.0, alpha * 0.25));
                    textShadows.Add($"0 0 {Math.Max(1, radiusPx * 0.15):0.#}px {c1}");
                    textShadows.Add($"0 0 {Math.Max(2, radiusPx * 0.5):0.#}px {c2}");
                    textShadows.Add($"0 0 {Math.Max(4, radiusPx * 1.0):0.#}px {c3}");
                    textShadows.Add($"0 0 {Math.Max(8, radiusPx * 2.0):0.#}px {c4}");
                    break;
                }
                case "reflection":
                    // Reflection handled at paragraph level via GetW14ReflectionCss()
                    // because -webkit-box-reflect on inline spans overlaps content below
                    break;
            }
        }

        // Legacy w: namespace run effects (w:shadow/outline/emboss/imprint).
        // These boolean RunProperties live in the w: namespace, so the W14Ns
        // loop above skips them — they previously produced no CSS at all. Each
        // is "on" when present and not explicitly false (Val == null = true per
        // OOXML), matching the rProps.Bold/Strike read pattern used above.
        if (rProps.Shadow != null && (rProps.Shadow.Val == null || rProps.Shadow.Val.Value))
            textShadows.Add("1px 1px 0 rgba(0,0,0,0.5)");
        if (rProps.Emboss != null && (rProps.Emboss.Val == null || rProps.Emboss.Val.Value))
        {
            // 3D raised: light edge on top, dark edge below.
            textShadows.Add("0 1px 0 rgba(255,255,255,.7)");
            textShadows.Add("0 -1px 0 rgba(0,0,0,.4)");
        }
        if (rProps.Imprint != null && (rProps.Imprint.Val == null || rProps.Imprint.Val.Value))
        {
            // 3D engraved: reversed vs emboss.
            textShadows.Add("0 -1px 0 rgba(255,255,255,.7)");
            textShadows.Add("0 1px 0 rgba(0,0,0,.4)");
        }
        if (rProps.Outline != null && (rProps.Outline.Val == null || rProps.Outline.Val.Value))
        {
            // Hollow/outline text: stroke the glyph edge AND make the fill
            // transparent so the interior shows through (white-centre + edge =
            // hollow outline, matching Word). Stroke alone only thickened the
            // glyph, leaving it solid. -webkit-text-fill-color overrides the
            // fill independently of `color`, which still drives the stroke
            // (currentColor). Chromium (Playwright preview) honours both.
            parts.Add("-webkit-text-stroke:0.5pt currentColor");
            parts.Add("-webkit-text-fill-color:transparent");
        }

        if (textShadows.Count > 0)
            parts.Add($"text-shadow:{string.Join(",", textShadows)}");
    }

    private static bool HasW14Reflection(Paragraph para)
    {
        foreach (var run in para.Elements<Run>())
        {
            var rProps = run.RunProperties;
            if (rProps == null) continue;
            if (rProps.ChildElements.Any(c => c.NamespaceUri == W14Ns && c.LocalName == "reflection"))
                return true;
        }
        return false;
    }

    /// <summary>
    /// If any run in the paragraph has w14:reflection, appends a flipped duplicate
    /// block element below the original to simulate the reflection effect.
    /// This approach reserves proper layout space (unlike -webkit-box-reflect).
    /// </summary>
    private void AppendW14ReflectionBlock(StringBuilder sb, Paragraph para, string tag, string? baseStyle)
    {
        // Find the first run with w14:reflection
        OpenXmlElement? reflectionEl = null;
        foreach (var run in para.Elements<Run>())
        {
            var rProps = run.RunProperties;
            if (rProps == null) continue;
            foreach (var child in rProps.ChildElements)
            {
                if (child.NamespaceUri == W14Ns && child.LocalName == "reflection")
                { reflectionEl = child; break; }
            }
            if (reflectionEl != null) break;
        }
        if (reflectionEl == null) return;

        var attrs = reflectionEl.GetAttributes().ToDictionary(a => a.LocalName, a => a.Value);
        var stA = attrs.TryGetValue("stA", out var sa) && int.TryParse(sa, out var saVal) ? saVal / 1000.0 : 50.0;
        var endA = attrs.TryGetValue("endA", out var ea) && int.TryParse(ea, out var eaVal) ? eaVal / 1000.0 : 0.0;
        var endPos = attrs.TryGetValue("endPos", out var ep) && int.TryParse(ep, out var epVal) ? epVal / 1000.0 : 90.0;
        var distEmu = attrs.TryGetValue("dist", out var d) && long.TryParse(d, out var dVal) ? dVal : 0;
        var blurEmu = attrs.TryGetValue("blurRad", out var br) && long.TryParse(br, out var brVal) ? brVal : 0;
        var distPx = distEmu / EmuConverter.EmuPerPointF * 1.333;
        var blurPx = blurEmu / EmuConverter.EmuPerPointF * 1.333;

        // Build the reflection element: flipped, fading, non-interactive
        var reflectStyle = new List<string>();
        if (!string.IsNullOrEmpty(baseStyle)) reflectStyle.Add(baseStyle);
        reflectStyle.Add("transform:scaleY(-1)");
        reflectStyle.Add("margin:0");
        reflectStyle.Add($"padding-top:{distPx:0.#}px");
        reflectStyle.Add("overflow:hidden");
        reflectStyle.Add("pointer-events:none");
        reflectStyle.Add("user-select:none");
        reflectStyle.Add("text-shadow:none");
        // Gradient mask: opaque at bottom (nearest to original text) → transparent at top
        // Since the element is scaleY(-1) with transform-origin:top, the visual top is the
        // reflected bottom of the text (closest to original). Mask goes from fully opaque
        // at bottom to transparent at top in the element's own coordinate space.
        var maskPct = 100.0 - endPos;  // where full transparency starts
        reflectStyle.Add($"-webkit-mask-image:linear-gradient(to top,rgba(0,0,0,{stA / 100.0:0.##}) {maskPct:0.#}%,rgba(0,0,0,{endA / 100.0:0.###}) 100%)");
        reflectStyle.Add($"mask-image:linear-gradient(to top,rgba(0,0,0,{stA / 100.0:0.##}) {maskPct:0.#}%,rgba(0,0,0,{endA / 100.0:0.###}) 100%)");
        if (blurPx > 0)
            reflectStyle.Add($"filter:blur({blurPx:0.#}px)");

        sb.Append($"<{tag} aria-hidden=\"true\" style=\"{string.Join(";", reflectStyle)}\">");
        RenderParagraphContentHtml(sb, para);
        sb.AppendLine($"</{tag}>");
    }

    /// <summary>
    /// Read a dxa width value leniently. The SDK's typed Int16/Int32
    /// <c>.Value</c> getter throws <see cref="System.FormatException"/> on
    /// decimal dxa strings such as <c>"108.0"</c> that some non-Word generators
    /// emit — and an uncaught throw aborted the ENTIRE HTML render (zero output).
    /// Reading the raw <c>InnerText</c> never parses, so we truncate the
    /// fractional part ourselves. Returns null for absent/"auto"/unparseable.
    /// </summary>
    private static int? LenientDxa(DocumentFormat.OpenXml.OpenXmlSimpleType? widthVal)
    {
        var raw = widthVal?.InnerText;
        if (string.IsNullOrEmpty(raw)) return null;
        if (int.TryParse(raw, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var i))
            return i;
        if (double.TryParse(raw, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var d))
            return (int)System.Math.Round(d);
        return null;
    }

    private string GetTableCellInlineCss(TableCell cell, bool tableBordersNone, TableBorders? tblBorders = null,
        Dictionary<string, TableConditionalFormat>? condFormats = null, List<string>? condTypes = null,
        int rowIdx = 0, int colIdx = 0, int totalRows = 1, int totalCols = 1,
        double? exactRowHeightPt = null, TableCellMarginDefault? tblCellMar = null,
        string? tableStyleCellFill = null)
    {
        var parts = new List<string>();
        var tcPr = cell.TableCellProperties;

        // Table-style base cell shading (<w:style><w:tcPr><w:shd>) — the
        // whole-table cell fill, lowest priority. Seeded first so a conditional
        // format (tblStylePr) shd or a direct cell shd (both below) can override
        // it via the RemoveAll(background-color:) path. Without this a dark-list
        // table's blue fill never appears and its white run color (applied on the
        // <table>) renders the cell labels as an invisible empty frame.
        if (tableStyleCellFill != null)
            parts.Add($"background-color:{tableStyleCellFill}");

        // Apply table-level borders: outer borders only on table edges, insideH/V on inner edges
        if (!tableBordersNone && tblBorders != null)
        {
            var hInner = !IsBorderNone(tblBorders.InsideHorizontalBorder) ? (OpenXmlElement)tblBorders.InsideHorizontalBorder! : null;
            var vInner = !IsBorderNone(tblBorders.InsideVerticalBorder) ? (OpenXmlElement)tblBorders.InsideVerticalBorder! : null;

            // Top edge: outer border if first row, insideH if inner row
            RenderBorderCss(parts, rowIdx == 0 ? tblBorders.TopBorder : hInner, "border-top");
            // Bottom edge: outer border if last row, insideH if inner row
            RenderBorderCss(parts, rowIdx == totalRows - 1 ? tblBorders.BottomBorder : hInner, "border-bottom");
            // Left edge: outer border if first col, insideV if inner col
            RenderBorderCss(parts, colIdx == 0 ? tblBorders.LeftBorder : vInner, "border-left");
            // Right edge: outer border if last col, insideV if inner col
            RenderBorderCss(parts, colIdx == totalCols - 1 ? tblBorders.RightBorder : vInner, "border-right");
        }

        // Apply conditional formatting from table style (priority order: banding < col < row)
        if (condFormats != null && condTypes != null)
        {
            foreach (var condType in condTypes)
            {
                if (!condFormats.TryGetValue(condType, out var fmt)) continue;

                // Cell shading / background
                var condFill = ResolveShadingFill(fmt.Shading);
                if (condFill != null)
                {
                    parts.RemoveAll(p => p.StartsWith("background-color:"));
                    parts.Add($"background-color:{condFill}");
                }

                // Border overrides from conditional format
                if (fmt.Borders != null)
                {
                    var cb = fmt.Borders;
                    // Apply or clear each border edge from conditional format
                    // val=nil/none means explicitly REMOVE the border
                    ApplyCondBorder(parts, cb.TopBorder, "border-top");
                    ApplyCondBorder(parts, cb.BottomBorder, "border-bottom");
                    ApplyCondBorder(parts, cb.LeftBorder, "border-left");
                    ApplyCondBorder(parts, cb.RightBorder, "border-right");
                    // insideH/insideV only apply to edges NOT already set by explicit top/bottom/left/right
                    if (cb.InsideHorizontalBorder != null)
                    {
                        if (cb.TopBorder == null) ApplyCondBorder(parts, cb.InsideHorizontalBorder, "border-top");
                        if (cb.BottomBorder == null) ApplyCondBorder(parts, cb.InsideHorizontalBorder, "border-bottom");
                    }
                    if (cb.InsideVerticalBorder != null)
                    {
                        if (cb.LeftBorder == null) ApplyCondBorder(parts, cb.InsideVerticalBorder, "border-left");
                        if (cb.RightBorder == null) ApplyCondBorder(parts, cb.InsideVerticalBorder, "border-right");
                    }
                }

                // Text formatting from conditional format (bold, color, font-size)
                if (fmt.RunProperties != null)
                {
                    var rPr = fmt.RunProperties;
                    if (rPr.Bold != null && (rPr.Bold.Val == null || rPr.Bold.Val.Value))
                        parts.Add("font-weight:bold");
                    if (rPr.Italic != null && (rPr.Italic.Val == null || rPr.Italic.Val.Value))
                        parts.Add("font-style:italic");
                    var condColor = ResolveRunColor(rPr.Color);
                    if (condColor != null)
                        parts.Add($"color:{condColor}");
                    if (rPr.FontSize?.Val?.Value is string fsz && int.TryParse(fsz, out var fhp))
                    {
                        parts.Add($"font-size:{fhp / 2.0}pt");
                        parts.Add("__TSF__"); // marker for table style font-size override
                    }
                }
            }
        }

        if (tcPr == null)
        {
            // No cell properties at all: the global `th,td { padding:0 5.4pt }`
            // CSS rule already supplies the Word default, so we normally emit
            // nothing. But an explicit table-level tblCellMar (incl. 0) must
            // still win over that CSS default — emit it inline so a tblCellMar
            // L/R=0 table with bare cells doesn't fall back to 5.4pt.
            if (tblCellMar != null)
            {
                var tTop = LenientDxa(tblCellMar.TopMargin?.Width);
                var tBot = LenientDxa(tblCellMar.BottomMargin?.Width);
                var tLeft = LenientDxa(tblCellMar.TableCellLeftMargin?.Width);
                var tRight = LenientDxa(tblCellMar.TableCellRightMargin?.Width);
                var pTop = tTop != null ? $"{Units.TwipsToPt(tTop.Value):0.#}pt" : "0pt";
                var pBot = tBot != null ? $"{Units.TwipsToPt(tBot.Value):0.#}pt" : "0pt";
                var pLeft = tLeft != null ? $"{Units.TwipsToPt(tLeft.Value):0.#}pt" : "5.4pt";
                var pRight = tRight != null ? $"{Units.TwipsToPt(tRight.Value):0.#}pt" : "5.4pt";
                parts.Add($"padding:{pTop} {pRight} {pBot} {pLeft}");
            }
            return string.Join(";", parts);
        }

        // Shading / fill (supports theme colors) — direct cell shading overrides conditional
        var cellFill = ResolveShadingFill(tcPr.Shading);
        if (cellFill != null)
        {
            parts.RemoveAll(p => p.StartsWith("background-color:"));
            parts.Add($"background-color:{cellFill}");
        }

        // Vertical alignment
        var vAlign = tcPr.TableCellVerticalAlignment?.Val;
        if (vAlign != null)
        {
            var va = vAlign.InnerText switch
            {
                "center" => "middle",
                "bottom" => "bottom",
                _ => (string?)null
            };
            if (va != null) parts.Add($"vertical-align:{va}");
        }

        // Cell-level borders override table-level and conditional
        var tcBorders = tcPr.TableCellBorders;
        if (tcBorders != null)
        {
            // A PRESENT cell border side always overrides the inherited
            // table-level border, including when it is an explicit nil/none:
            // OOXML treats <w:* w:val="nil"/> as "suppress the table border on
            // this side", not "inherit it". So remove the inherited border-<side>
            // whenever the side element exists, then paint the cell border only
            // when it actually has a value. An ABSENT side (null) still inherits.
            if (tcBorders.TopBorder != null) { parts.RemoveAll(p => p.StartsWith("border-top:")); if (!IsBorderNone(tcBorders.TopBorder)) RenderBorderCss(parts, tcBorders.TopBorder, "border-top"); }
            if (tcBorders.BottomBorder != null) { parts.RemoveAll(p => p.StartsWith("border-bottom:")); if (!IsBorderNone(tcBorders.BottomBorder)) RenderBorderCss(parts, tcBorders.BottomBorder, "border-bottom"); }
            if (tcBorders.LeftBorder != null) { parts.RemoveAll(p => p.StartsWith("border-left:")); if (!IsBorderNone(tcBorders.LeftBorder)) RenderBorderCss(parts, tcBorders.LeftBorder, "border-left"); }
            if (tcBorders.RightBorder != null) { parts.RemoveAll(p => p.StartsWith("border-right:")); if (!IsBorderNone(tcBorders.RightBorder)) RenderBorderCss(parts, tcBorders.RightBorder, "border-right"); }

            // Diagonal cell borders (w:tl2br / w:tr2bl) render as an absolutely
            // positioned SVG overlay inside the <td> (HTML has no diagonal
            // border). The <td> must become position:relative for the overlay
            // to anchor — added only when a diagonal is actually present to
            // minimize CSS regression surface. Mirrors the Excel/PPTX cell-diag
            // idiom. The SVG itself is prepended to the cell content in
            // RenderTableHtml via TryBuildCellDiagonalSvg.
            if (TryBuildCellDiagonalSvg(cell) != null)
                parts.Add("position:relative");
        }

        // Cell width
        var width = tcPr.TableCellWidth?.Width?.Value;
        if (width != null && int.TryParse(width, out var w))
        {
            var type = tcPr.TableCellWidth?.Type?.InnerText;
            if (type == "dxa")
                parts.Add($"width:{w / 20.0:0.##}pt");
            else if (type == "pct")
                parts.Add($"width:{w / 50.0:0.#}%");
        }

        // Cell text direction (tcDir): rotate text 90° or 270° via CSS writing-mode + transform
        // Common values: btLr (bottom→top, left→right = 90° CCW), tbRl (top→bottom, right→left = 90° CW)
        var tcDir = tcPr.GetFirstChild<TextDirection>()?.Val?.InnerText;
        if (tcDir != null)
        {
            var wm = tcDir switch
            {
                "btLr" => "vertical-lr",                            // read bottom-up (left-to-right column axis)
                "tbRl" => "vertical-rl",                            // read top-down
                "lrTb" or null => null,                             // default horizontal
                _ => null,
            };
            if (wm != null) parts.Add($"writing-mode:{wm}");
        }

        // Cell noWrap — prevents content wrapping within the cell. Pair with
        // overflow:hidden so that under a fixed-layout table (table-layout:fixed,
        // where the column width is a hard cap) over-long single-line content is
        // clipped at the cell's own edge instead of visually bleeding across the
        // neighbouring columns. In autofit tables the column grows to fit the
        // nowrap content, so the overflow guard never triggers there.
        if (tcPr.NoWrap != null)
        {
            parts.Add("white-space:nowrap");
            parts.Add("overflow:hidden");
        }

        // #7a0: vertical-writing cell + noWrap interaction. When both are
        // present, flex alignment + min-height otherwise position text in
        // the cell's middle; Word anchors it at the inline-start edge and
        // fills the declared trHeight. Force flex-start + stretch so the
        // text column runs from top (or right, in vertical-rl) of the cell.
        if (tcDir != null && tcPr.NoWrap != null)
        {
            parts.Add("justify-content:flex-start");
            parts.Add("align-items:stretch");
        }

        // Padding resolution mirrors Word's per-edge cell-margin cascade:
        //   cell tcMar slot (incl. 0) > table tblCellMar slot (incl. 0) > Word
        //   TableNormal default (top=0 left=108(=5.4pt) bottom=0 right=108).
        // The earlier code consulted only the cell-level tcMar and fell back to
        // the hardcoded 5.4pt L/R default whenever a cell lacked its own tcMar —
        // ignoring an explicit table-level tblCellMar of 0 and stealing 10.8pt
        // of horizontal content width per column under table-layout:fixed +
        // box-sizing:border-box (header/number wrap+clip). A document that
        // declares tblCellMar L/R=0 must yield td padding L/R=0. (The older
        // CellPadVComp=3pt vertical compensation for line-height:1 ascender
        // clipping is no longer needed since cli emits unitless line-height.)
        var margins = tcPr?.TableCellMargin;
        {
            // top/bottom: TopMargin/BottomMargin on both tcMar and tblCellMar.
            var topVal = LenientDxa(margins?.TopMargin?.Width) ?? LenientDxa(tblCellMar?.TopMargin?.Width);
            var botVal = LenientDxa(margins?.BottomMargin?.Width) ?? LenientDxa(tblCellMar?.BottomMargin?.Width);
            // left/right: tcMar exposes Left/Start + Right/End; tblCellMar uses
            // the distinct TableCellLeftMargin / TableCellRightMargin children.
            var leftVal = LenientDxa(margins?.LeftMargin?.Width) ?? LenientDxa(margins?.StartMargin?.Width)
                          ?? LenientDxa(tblCellMar?.TableCellLeftMargin?.Width);
            var rightVal = LenientDxa(margins?.RightMargin?.Width) ?? LenientDxa(margins?.EndMargin?.Width)
                           ?? LenientDxa(tblCellMar?.TableCellRightMargin?.Width);
            var padTop = topVal != null ? $"{Units.TwipsToPt(topVal.Value):0.#}pt" : "0pt";
            var padBot = botVal != null ? $"{Units.TwipsToPt(botVal.Value):0.#}pt" : "0pt";
            var padLeft = leftVal != null ? $"{Units.TwipsToPt(leftVal.Value):0.#}pt" : "5.4pt";
            var padRight = rightVal != null ? $"{Units.TwipsToPt(rightVal.Value):0.#}pt" : "5.4pt";
            parts.Add($"padding:{padTop} {padRight} {padBot} {padLeft}");
        }

        // hRule="exact": Word pins the row to the exact height but still SHOWS
        // the cell text — it does not blank a cell whose content is taller than
        // the exact value. The earlier fixed height + max-height + overflow:hidden
        // hard-clipped over-tall cells to empty (lost evaluation labels, list
        // rows). Emit the exact value as a min-height floor instead: normal cells
        // (content ≤ exact) keep the exact height unchanged, while over-tall cells
        // grow to show their content rather than going blank. Priority: content
        // visible over strict exact height (R49/R31 don't-clip-content rule).
        if (exactRowHeightPt is double exH)
        {
            parts.Add($"min-height:{exH:0.#}pt");
        }

        return string.Join(";", parts);
    }

    // ==================== CSS Helpers ====================

    /// <summary>
    /// If the cell carries a diagonal border (w:tl2br / w:tr2bl with a non-nil
    /// style), return an absolutely-positioned inline SVG that draws the
    /// diagonal line(s) inside the TD — HTML has no native diagonal border.
    /// tl2br = top-left (0,0) → bottom-right (100%,100%);
    /// tr2bl = top-right (100%,0) → bottom-left (0,100%). Both may be present.
    /// Honors w:sz (eighths-of-pt → pt) and w:color. Returns null when the cell
    /// has no diagonal. Mirrors the Excel/PPTX cell-diag overlay idiom. The TD
    /// must be position:relative for the overlay to anchor (set in
    /// GetTableCellInlineCss).
    /// </summary>
    private string? TryBuildCellDiagonalSvg(TableCell? cell)
    {
        var tcBorders = cell?.TableCellProperties?.TableCellBorders;
        if (tcBorders == null) return null;

        var tlBr = tcBorders.TopLeftToBottomRightCellBorder;
        var trBl = tcBorders.TopRightToBottomLeftCellBorder;
        bool hasTlBr = !IsBorderNone(tlBr);
        bool hasTrBl = !IsBorderNone(trBl);
        if (!hasTlBr && !hasTrBl) return null;

        var lines = new StringBuilder();
        if (hasTlBr)
        {
            var (color, widthPt) = ResolveDiagonalLine(tlBr!);
            lines.Append($"<line x1=\"0\" y1=\"0\" x2=\"100%\" y2=\"100%\" stroke=\"{color}\" stroke-width=\"{widthPt:0.##}\"/>");
        }
        if (hasTrBl)
        {
            var (color, widthPt) = ResolveDiagonalLine(trBl!);
            lines.Append($"<line x1=\"0\" y1=\"100%\" x2=\"100%\" y2=\"0\" stroke=\"{color}\" stroke-width=\"{widthPt:0.##}\"/>");
        }

        return $"<svg class=\"cell-diag\" width=\"100%\" height=\"100%\" style=\"position:absolute;inset:0;pointer-events:none;overflow:visible\" preserveAspectRatio=\"none\">{lines}</svg>";
    }

    /// <summary>Resolve a diagonal cell border's color + stroke width (pt) the
    /// same way RenderBorderCss resolves box borders (sz eighths-of-pt, hex or
    /// themeColor with tint/shade, fallback black).</summary>
    /// <summary>
    /// Resolve a literal-or-theme color to a CSS color string, handling the
    /// "#hex (unless auto) else themeColor + themeTint/themeShade else null"
    /// chain shared by cell/diagonal borders and run color. Callers supply the
    /// literal color and theme name (their attribute names differ — borders use
    /// w:color/w:themeColor, run color uses w:val/typed ThemeColor) and the
    /// fallback for the null case. themeTint/themeShade are read generically
    /// off <paramref name="element"/>.
    /// </summary>
    private string? ResolveThemeAwareColor(OpenXmlElement element, string? literalColor, string? themeName)
    {
        if (literalColor != null && !literalColor.Equals("auto", StringComparison.OrdinalIgnoreCase) && IsHexColor(literalColor))
            return $"#{literalColor}";
        if (themeName != null && GetThemeColors().TryGetValue(themeName, out var tcHex))
        {
            var tint = element.GetAttributes().FirstOrDefault(a => a.LocalName == "themeTint").Value;
            var shade = element.GetAttributes().FirstOrDefault(a => a.LocalName == "themeShade").Value;
            return ApplyTintShade(tcHex, tint, shade);
        }
        return null;
    }

    private (string color, double widthPt) ResolveDiagonalLine(OpenXmlElement border)
    {
        var sz = border.GetAttributes().FirstOrDefault(a => a.LocalName == "sz").Value;
        var color = border.GetAttributes().FirstOrDefault(a => a.LocalName == "color").Value;
        var themeColor = border.GetAttributes().FirstOrDefault(a => a.LocalName == "themeColor").Value;
        var widthPt = sz != null && int.TryParse(sz, out var s) ? Math.Max(0.5, s / 8.0) : 1.0;

        var cssColor = ResolveThemeAwareColor(border, color, themeColor) ?? "#000";
        return (cssColor, widthPt);
    }

    private void RenderBorderCss(List<string> parts, OpenXmlElement? border, string cssProp)
    {
        if (border == null) return;
        var val = border.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (val == null || val == "nil" || val == "none") return;

        var sz = border.GetAttributes().FirstOrDefault(a => a.LocalName == "sz").Value;
        var color = border.GetAttributes().FirstOrDefault(a => a.LocalName == "color").Value;

        var style = val switch
        {
            "single" => "solid",
            "thick" => "solid",
            "double" => "double",
            "triple" => "double",  // CSS has no 3-line; double is closest
            "dashed" or "dashSmallGap" => "dashed",
            "dashDotStroked" or "dashDotHeavy" => "dashed",
            "dotted" => "dotted",
            "dotDash" or "dotDotDash" => "dashed",
            "wave" or "doubleWave" => "solid",  // CSS has no wave border
            _ => "solid"
        };
        // OOXML border sz is in 1/8 of a point (8 = 1pt, 24 = 3pt, etc.)
        var widthPt = sz != null && int.TryParse(sz, out var s) ? Math.Max(0.5, s / 8.0) : 1.0;
        // CSS double border style needs at least ~2.25pt (≈3px) to show two visible lines
        if (style == "double" && widthPt < 2.25) widthPt = 2.25;
        var width = $"{widthPt:0.##}pt";

        // Resolve color: try direct color, then themeColor with tint/shade
        var themeColor = border.GetAttributes().FirstOrDefault(a => a.LocalName == "themeColor").Value;
        var cssColor = ResolveThemeAwareColor(border, color, themeColor) ?? "#000";

        parts.Add($"{cssProp}:{width} {style} {cssColor}");

        // Border spacing (w:space) → padding on the corresponding side
        var space = border.GetAttributes().FirstOrDefault(a => a.LocalName == "space").Value;
        if (space != null && int.TryParse(space, out var spacePt) && spacePt > 0)
        {
            var paddingSide = cssProp.Replace("border-", "padding-");
            parts.Add($"{paddingSide}:{spacePt}pt");
        }
    }

    /// <summary>Resolve a run Color element to a CSS color string, handling themeColor + tint/shade.</summary>
    private string? ResolveRunColor(DocumentFormat.OpenXml.Wordprocessing.Color? color)
    {
        if (color == null) return null;
        return ResolveThemeAwareColor(color, color.Val?.Value, color.ThemeColor?.InnerText);
    }

    /// <summary>
    /// Effective background color (#RRGGBB) behind a run, for automatic-color
    /// (color=auto) text contrast. Priority mirrors Word's shading cascade:
    /// run shd (w:rPr/w:shd) > paragraph shd (direct or style) > nearest
    /// ancestor table-cell shd. Returns null when no opaque fill applies
    /// (backdrop is the page/white) — callers then keep black auto text.
    /// </summary>
    private string? ResolveEffectiveBackgroundForRun(RunProperties? rProps, Paragraph? para)
    {
        // 1) Run-level shading (inverse-video spans set this directly).
        var runFill = ResolveShadingFill(rProps?.Shading);
        if (runFill != null) return runFill;

        if (para != null)
        {
            // 2) Paragraph shading — direct, else via the pStyle chain (the
            //    deep-blue title bars carry pPr/shd w:fill="1F3864").
            var paraFill = ResolveShadingFill(para.ParagraphProperties?.Shading)
                ?? ResolveParagraphShadingFromStyle(para);
            if (paraFill != null) return paraFill;

            // 3) Nearest ancestor table cell's shading.
            var cell = para.Ancestors<TableCell>().FirstOrDefault();
            var cellFill = ResolveShadingFill(cell?.TableCellProperties?.Shading);
            if (cellFill != null) return cellFill;
        }
        return null;
    }

    /// <summary>
    /// True when a #RRGGBB color is dark enough that automatic text should be
    /// white. Standard relative-luminance approximation, threshold 128/255.
    /// Mirrors the pptx <c>IsColorDark</c> helper.
    /// </summary>
    private static bool IsColorDark(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length < 6) return false;
        var (r, g, b) = ColorMath.HexToRgb(hex);
        return (r * 0.299 + g * 0.587 + b * 0.114) < 128;
    }

    // Unit conversions moved to shared Units class (Core/Units.cs).

    private static string? HighlightToCssColor(string highlight) => highlight.ToLowerInvariant() switch
    {
        "yellow" => "#FFFF00",
        "green" => "#00FF00",
        "cyan" => "#00FFFF",
        "magenta" => "#FF00FF",
        "blue" => "#0000FF",
        "red" => "#FF0000",
        "darkblue" => "#000080",
        "darkcyan" => "#008080",
        "darkgreen" => "#008000",
        "darkmagenta" => "#800080",
        "darkred" => "#800000",
        "darkyellow" => "#808000",
        "darkgray" => "#808080",
        "lightgray" => "#C0C0C0",
        "black" => "#000000",
        "white" => "#FFFFFF",
        _ => null
    };

    /// <summary>
    /// Heuristic: does this typeface name belong to the serif family?
    /// Used to pick the generic CSS fallback (serif vs sans-serif) when neither
    /// the primary font nor the CJK fallback is installed.
    /// </summary>
    private static bool IsLikelySerif(string font)
    {
        var f = font.ToLowerInvariant();
        // Western serif faces
        if (f.Contains("times") || f.Contains("serif") || f.Contains("georgia")
            || f.Contains("cambria") || f.Contains("garamond") || f.Contains("palatino")
            || f.Contains("book antiqua") || f.Contains("constantia") || f.Contains("didot")
            || f.Contains("baskerville") || f.Contains("minion"))
            return true;
        // CJK serif (宋体 / Song / Ming / Mincho)
        if (f.Contains("song") || f.Contains("ming") || f.Contains("mincho")
            || f.Contains("fangsong") || font.Contains("宋") || font.Contains("仿宋")
            || font.Contains("明朝"))
            return true;
        return false;
    }

    /// <summary>
    /// Heuristic: does this typeface name belong to the monospace (fixed-width)
    /// family? Picks the <c>monospace</c> generic fallback so code/columns stay
    /// aligned when the named font is unavailable.
    /// </summary>
    private static bool IsLikelyMonospace(string font)
    {
        var f = font.ToLowerInvariant();
        return f.Contains("courier") || f.Contains("consolas")
            || f.Contains("lucida console") || f.Contains("monaco")
            || f.Contains("menlo") || f.Contains("cascadia")
            || f.Contains("mono") || f.Contains("sf mono")
            || f.Contains("monospace");
    }

    /// <summary>
    /// Pick the generic CSS family (monospace / serif / sans-serif) to terminate
    /// a font-family list, so the run still renders in the right class when the
    /// named font and any CJK fallback are unavailable.
    /// </summary>
    private static string GenericFontFamily(string font)
        => IsLikelyMonospace(font) ? "monospace"
            : IsLikelySerif(font) ? "serif"
            : "sans-serif";

    /// <summary>
    /// Returns CSS fallback fonts for common Windows Chinese fonts that are unavailable on Mac.
    /// </summary>
    private string? GetChineseFontFallback(string font)
    {
        var result = font switch
        {
            "仿宋_GB2312" => "'仿宋',FangSong,STFangsong",
            "楷体_GB2312" => "'楷体',KaiTi,STKaiti",
            "长城小标宋体" => "'华文中宋',STZhongsong,'宋体',SimSun",
            "黑体" => "'Heiti SC',STHeiti",
            _ => null
        };
        if (result != null) return result;
        // Fall back to CJK font mapping for western fonts
        var cjk = GetCjkFontFallback(font, _eastAsiaLang, _themeCjkFont);
        return string.IsNullOrEmpty(cjk) ? null : cjk.TrimStart(',', ' ');
    }

    /// <summary>Resolve font size from a style chain by styleId. Returns e.g. "10pt" or null.</summary>
    /// <summary>Resolve the dominant font for line-height calculation from a paragraph's runs.</summary>
    /// <remarks>
    /// Word's line height = max ratio across fonts that actually have glyphs
    /// in the line. EastAsia is only counted when at least one CJK char is
    /// present; setting rFonts.eastAsia on a Latin-only run does not enlarge
    /// the line. We scan Ascii / HighAnsi (always) and EastAsia (only when
    /// the paragraph has any CJK char) across all runs and return the font
    /// with the highest ratio. CSS unitless line-height inheritance then
    /// scales it per-span by each run's own font-size.
    /// </remarks>
    /// <summary>
    /// Inline style for a footnote/endnote-reference &lt;sup&gt;. Resolves the
    /// run's own font first (when its rFonts pin one) and falls back to the
    /// paragraph's principal font; reads the font's OS/2 sub/superscript
    /// fields to size and position the reference glyph the way the run's
    /// own font would. See [Css.cs vertAlign emit](WordHandler.HtmlPreview.Css.cs)
    /// for the symmetrical inline-run path; the font-agnostic fallback is
    /// the same CSS browsers use for &lt;sup&gt; when OS/2 data isn't queried.
    /// </summary>
    private string ResolveNoteRefSupStyle(Run run, Paragraph para)
    {
        var rProps = run.RunProperties;
        var fonts = rProps?.RunFonts;
        string? font = fonts?.Ascii?.Value
            ?? ResolveThemeFont(fonts?.AsciiTheme?.InnerText)
            ?? fonts?.HighAnsi?.Value
            ?? ResolveThemeFont(fonts?.HighAnsiTheme?.InnerText)
            ?? ResolveParaFontForLineHeight(para);
        return BuildSupStyleFromFont(font);
    }

    /// <summary>
    /// Inline style for the footnote / endnote list-marker &lt;sup&gt; emitted
    /// inside the page-bottom notes area. The note text typically renders in
    /// the FootnoteText / EndnoteText style font; falls back to the document
    /// default when that style omits an explicit typeface.
    /// </summary>
    private string ResolveNoteListSupStyle(string styleId)
    {
        var font = ResolveStyleFontName(styleId) ?? ReadDocDefaults().Font;
        return BuildSupStyleFromFont(font);
    }

    private static string BuildSupStyleFromFont(string? font)
    {
        var ss = !string.IsNullOrEmpty(font)
            ? FontMetricsReader.GetSuperSubMetrics(font)
            : default;
        if (!ss.IsEmpty && ss.SuperSizeEm > 0 && ss.SuperOffsetEm > 0)
            return $"vertical-align:baseline;position:relative;top:-{ss.SuperOffsetEm:0.###}em;font-size:{ss.SuperSizeEm * 100:0.#}%;text-decoration:none";
        return "vertical-align:baseline;position:relative;top:-0.35em;font-size:smaller;text-decoration:none";
    }

    private string ResolveParaFontForLineHeight(Paragraph para)
    {
        bool paraHasCjk = para.Elements<Run>()
            .SelectMany(r => r.Descendants<Text>())
            .SelectMany(t => t.Text ?? string.Empty)
            .Any(IsCjkCodepoint);

        string? best = null;
        double bestRatio = 0;

        void Consider(RunProperties rProps, bool includeEastAsia)
        {
            var fonts = rProps.RunFonts;
            if (fonts == null) return;
            // OOXML §17.3.2.27: each rFonts slot may carry either a literal
            // typeface OR a *Theme reference (asciiTheme/hAnsiTheme/...).
            // When only the theme attribute is set, resolve it via theme1.xml
            // so the line-height calculation sees the same effective font
            // the renderer uses.
            var slots = new List<string?>
            {
                fonts.Ascii?.Value ?? ResolveThemeFont(fonts.AsciiTheme?.InnerText),
                fonts.HighAnsi?.Value ?? ResolveThemeFont(fonts.HighAnsiTheme?.InnerText),
            };
            if (includeEastAsia)
                slots.Add(fonts.EastAsia?.Value ?? ResolveThemeFont(fonts.EastAsiaTheme?.InnerText));
            foreach (var f in slots)
            {
                if (string.IsNullOrEmpty(f)) continue;
                var r = FontMetricsReader.GetRatio(f);
                if (r > bestRatio) { bestRatio = r; best = f; }
            }
        }

        foreach (var run in para.Elements<Run>())
            Consider(ResolveEffectiveRunProperties(run, para), paraHasCjk);

        // Empty paragraphs carry their would-be font on pPr/rPr (the mark
        // properties). EastAsia is honored unconditionally here — without
        // any actual text we can't gate by CJK content, but the writer
        // setting eastAsia signals intent for that font's metrics to apply.
        if (best == null)
        {
            var markProps = para.ParagraphProperties?.ParagraphMarkRunProperties;
            var synthRPr = new RunProperties();
            if (markProps != null)
            {
                foreach (var child in markProps.ChildElements)
                    synthRPr.AppendChild(child.CloneNode(true));
            }
            // Even when the paragraph mark carries no rPr (truly bare empty
            // paragraph), still run the synthetic run through the style
            // cascade so the default paragraph style's rFonts apply. Per
            // OOXML §17.7.5.2 rPrDefault and §17.3.1 paragraph-mark rPr,
            // the empty-paragraph mark inherits through the same docDefaults
            // → default style → direct chain that content runs traverse, so
            // the empty paragraph resolves to the same effective font.
            var synthRun = new Run(synthRPr);
            Consider(ResolveEffectiveRunProperties(synthRun, para), includeEastAsia: true);
        }
        if (best != null) return best;

        var defRFonts = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
            ?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.RunFonts;
        var defFont = defRFonts?.Ascii?.Value
            ?? ResolveThemeFont(defRFonts?.AsciiTheme?.InnerText)
            ?? defRFonts?.HighAnsi?.Value
            ?? ResolveThemeFont(defRFonts?.HighAnsiTheme?.InnerText);
        return defFont ?? GetThemeMinorLatinFont() ?? OfficeDefaultFonts.MinorLatin;
    }

    /// <summary>True when c falls in any CJK Unicode block: Unified Ideographs +
    /// Extension A, kana, Hangul syllables, CJK Symbols & Punctuation, CJK
    /// Compatibility, Halfwidth/Fullwidth Forms.</summary>
    private static bool IsCjkCodepoint(char c) =>
        (c >= 0x3000 && c <= 0x30FF) ||  // CJK Symbols & Punct, kana
        (c >= 0x3400 && c <= 0x4DBF) ||  // CJK Unified Extension A
        (c >= 0x4E00 && c <= 0x9FFF) ||  // CJK Unified Ideographs
        (c >= 0xAC00 && c <= 0xD7AF) ||  // Hangul Syllables
        (c >= 0xF900 && c <= 0xFAFF) ||  // CJK Compatibility
        (c >= 0xFF00 && c <= 0xFFEF);    // Halfwidth/Fullwidth Forms

    /// <summary>Read theme1.xml's <c>a:fontScheme/a:minorFont/a:latin/@typeface</c>.</summary>
    private string? GetThemeMinorLatinFont()
    {
        try
        {
            return _doc.MainDocumentPart?.ThemePart?.Theme?
                .ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
        }
        catch (System.Xml.XmlException) { return null; }
    }

    private string? ResolveStyleFontSize(string styleId)
    {
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = FindStyleById(current);
            if (style == null) break;
            var sz = style.StyleRunProperties?.FontSize?.Val?.Value;
            if (sz != null && int.TryParse(sz, out var halfPts))
                return $"{halfPts / 2.0:0.##}pt";
            current = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private string? ResolveStyleFontName(string styleId)
    {
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = FindStyleById(current);
            if (style == null) break;
            var rf = style.StyleRunProperties?.RunFonts;
            var name = rf?.Ascii?.Value
                ?? ResolveThemeFont(rf?.AsciiTheme?.InnerText)
                ?? rf?.HighAnsi?.Value
                ?? ResolveThemeFont(rf?.HighAnsiTheme?.InnerText);
            if (!string.IsNullOrEmpty(name)) return name;
            current = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private string? ResolveStyleColor(string styleId)
    {
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = FindStyleById(current);
            if (style == null) break;
            var cv = style.StyleRunProperties?.Color?.Val?.Value;
            if (cv != null && cv != "auto" && IsHexColor(cv)) return $"#{cv}";
            var tc = style.StyleRunProperties?.Color?.ThemeColor?.InnerText;
            if (tc != null && GetThemeColors().TryGetValue(tc, out var tcHex)) return $"#{tcHex}";
            current = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private ParagraphBorders? ResolveStyleParagraphBorders(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return null;
        // Word merges w:pBdr PER SIDE across the basedOn chain: a child style
        // declaring only w:bottom keeps the parent's top/left/right (verified
        // against real Word — unlike w:tblBorders, which replaces wholesale).
        // Walk derived→base and keep the most-derived declaration of each side.
        ParagraphBorders? merged = null;
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = FindStyleById(current);
            if (style == null) break;
            // GetFirstChild — Open XML SDK doesn't always surface less-common
            // pPr children as typed properties on StyleParagraphProperties.
            var pBdr = style.StyleParagraphProperties?.GetFirstChild<ParagraphBorders>();
            if (pBdr != null)
            {
                if (merged == null)
                    merged = (ParagraphBorders)pBdr.CloneNode(true);
                else
                    foreach (var side in pBdr.ChildElements)
                        if (!merged.ChildElements.Any(c => c.LocalName == side.LocalName))
                            merged.AppendChild(side.CloneNode(true));
            }
            current = style.BasedOn?.Val?.Value;
        }
        return merged;
    }

    // Resolve a paragraph's effective shd fill (direct shd, else style chain).
    private string? ResolveParagraphShadeFill(Paragraph? para)
    {
        if (para == null) return null;
        return ResolveShadingFill(para.ParagraphProperties?.Shading)
            ?? ResolveParagraphShadingFromStyle(para);
    }

    // True when `para` and an adjacent `sibling` form ONE continuous shaded box
    // (OOXML §17.3.1.24 border merge over a paragraph-shd fill): both carry a
    // resolved shd fill AND an identical four-side pBdr with no w:between. This
    // is the precondition for suppressing the inter-paragraph margin so the
    // shaded strips abut (HTML never paints background into a vertical margin).
    // The pBdr-equality + no-between gate matches the border-merge suppression
    // in GetParagraphInlineCss, so the margin join and the border join stay in
    // lockstep. A lone shaded paragraph, or shaded paragraphs whose borders
    // differ/are absent, returns false and keeps its normal margin.
    private bool ParagraphJoinsShadedBox(Paragraph para, Paragraph? sibling)
    {
        if (sibling == null) return false;
        if (ResolveParagraphShadeFill(para) == null) return false;
        if (ResolveParagraphShadeFill(sibling) == null) return false;
        var pBdr = para.ParagraphProperties?.ParagraphBorders
            ?? ResolveStyleParagraphBorders(para.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        if (pBdr == null || pBdr.BetweenBorder != null) return false;
        var sibBdr = ResolveSiblingParagraphBorders(sibling);
        if (sibBdr?.BetweenBorder != null) return false;
        return ParagraphBordersEqual(pBdr, sibBdr);
    }

    // Resolve a sibling paragraph's effective pBdr (direct pBdr, else style
    // chain) — same resolution as the main pBdr lookup, for border-merge
    // comparison against the current paragraph.
    private ParagraphBorders? ResolveSiblingParagraphBorders(Paragraph? sibling)
    {
        if (sibling == null) return null;
        return sibling.ParagraphProperties?.ParagraphBorders
            ?? ResolveStyleParagraphBorders(sibling.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    // Two pBdr blocks are "the same continuous box" (OOXML §17.3.1.24 merge)
    // when their four outer sides each match on val/color/sz/space. A null
    // sibling pBdr never matches. Border elements are compared by their
    // material attributes (not OuterXml) so namespace/attribute-order noise
    // doesn't defeat the match.
    private static bool ParagraphBordersEqual(ParagraphBorders? a, ParagraphBorders? b)
    {
        if (a == null || b == null) return false;
        return BorderAttrsEqual(a.TopBorder, b.TopBorder)
            && BorderAttrsEqual(a.BottomBorder, b.BottomBorder)
            && BorderAttrsEqual(a.LeftBorder, b.LeftBorder)
            && BorderAttrsEqual(a.RightBorder, b.RightBorder);
    }

    private static bool BorderAttrsEqual(OpenXmlElement? x, OpenXmlElement? y)
    {
        if (x == null && y == null) return true;
        if (x == null || y == null) return false;
        static string? Attr(OpenXmlElement e, string name) =>
            e.GetAttributes().FirstOrDefault(at => at.LocalName == name).Value;
        return Attr(x, "val") == Attr(y, "val")
            && Attr(x, "color") == Attr(y, "color")
            && Attr(x, "themeColor") == Attr(y, "themeColor")
            && Attr(x, "sz") == Attr(y, "sz")
            && Attr(x, "space") == Attr(y, "space");
    }

    // Resolved bold state for a pStyle chain: true → chain explicitly bold,
    // false → chain explicitly NOT bold, null → unspecified. Distinguishing
    // the three matters for headings: the Word `Title` style ships no <w:b/>
    // (renders thin), but the browser default `<h1>{font-weight:bold}` would
    // force it bold unless the renderer explicitly emits `font-weight:normal`.
    private bool? ResolveStyleBold(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return null;
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = FindStyleById(current);
            if (style == null) break;
            var b = style.StyleRunProperties?.Bold;
            if (b != null) return b.Val == null || b.Val.Value;
            current = style.BasedOn?.Val?.Value;
        }
        // No <w:b/> anywhere in the resolved chain → the explicit declarations
        // don't decide weight. Fall back to Word's built-in style table so a
        // heading style that ships no <w:b/> still reports its real weight: the
        // `Title` style renders THIN (Bold=false) but `<h1>`'s browser default
        // would force it bold unless we report false here. Heading1-4 / Subtitle
        // (Bold=true in the table, but any <w:b/> above already short-circuited)
        // stay bold; Heading5-9 / Title report false → caller emits
        // font-weight:normal. Genuinely-unresolvable styles still return null
        // (ResolveBuiltInStyleDefaults bails when a chain style is undefined),
        // deferring to the browser default rather than stomping built-in bold.
        var builtIn = ResolveBuiltInStyleDefaults(styleId);
        if (builtIn != null) return builtIn.Bold;
        return null;
    }

    private string? ResolveStyleIndent(string styleId)
    {
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = FindStyleById(current);
            if (style == null) break;
            var ind = style.StyleParagraphProperties?.Indentation;
            if (ind?.Left?.Value is string lv && int.TryParse(lv, out var twips))
                return $"{twips / 20.0:0.#}pt";
            if (ind?.FirstLine?.Value is string flv && int.TryParse(flv, out var flTwips))
                return $"{flTwips / 20.0:0.#}pt";
            current = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    // Strip every character that isn't a valid CSS identifier-ish character
    // for font names. OOXML rFonts/theme attrs are attacker-controlled, so
    // CssSanitize not only removes the obvious breakouts (" ' ; { } < > & \)
    // but also parens, colons, slashes, and anything non-alpha so a name like
    // `Arial";background:url(javascript:)//` can't appear as substring inside
    // the inline style (a CSS parser would treat it as a font name there, but
    // downstream safety checks still grep for the substring).
    private static string CssSanitize(string value)
    {
        if (string.IsNullOrEmpty(value)) return value;
        var sb = new StringBuilder(value.Length);
        foreach (var c in value)
            if (char.IsLetterOrDigit(c) || c == ' ' || c == '-' || c == '_' || c == '.')
                sb.Append(c);
        return sb.ToString();
    }

    private static string JsStringLiteral(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "\"\"";
        var sb = new StringBuilder("\"");
        foreach (var c in text)
        {
            switch (c)
            {
                case '\\': sb.Append("\\\\"); break;
                case '"': sb.Append("\\\""); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '<': sb.Append("\\x3c"); break;
                case '>': sb.Append("\\x3e"); break;
                default: sb.Append(c); break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    private static string HtmlEncode(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        var encoded = text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
        // Preserve consecutive spaces (HTML collapses them by default)
        // Replace runs of 2+ spaces: keep first as normal space, rest as &nbsp;
        encoded = Regex.Replace(encoded, @"  +", m =>
            " " + new string('\u00A0', m.Length - 1)); // space + (n-1) × &nbsp;
        return encoded;
    }

    /// <summary>HTML-encode for attribute values without nbsp conversion (used for LaTeX formulas).</summary>
    private static string HtmlEncodeAttr(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    // ==================== CSS Stylesheet ====================

    /// <summary>Check if document uses linked styles (w:linkStyles in settings).
    /// When true, Word applies default spaceAfter=10pt and lineSpacing=115% for Normal.</summary>
    private bool HasLinkedStyles()
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        return settings?.Descendants<DocumentFormat.OpenXml.Wordprocessing.LinkStyles>().Any() == true;
    }

    private string GenerateWordCss(PageLayout pg, DocDef dd)
    {
        // Use pt units (twips/20) for pixel-perfect accuracy — no cm→px conversion loss
        var mL = $"{pg.MarginLeftPt:0.#}pt";
        var mR = $"{pg.MarginRightPt:0.#}pt";
        var mT = $"{pg.MarginTopPt:0.#}pt";
        var mB = $"{pg.MarginBottomPt:0.#}pt";

        // Honor document-level auto-hyphenation setting. CSS `hyphens: auto`
        // requires the element (or ancestor) to specify a `lang` attribute;
        // browsers use the language-specific hyphenation dictionaries.
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var hyphensCss = settings?.Descendants<AutoHyphenation>().Any() == true
            ? "hyphens: auto; -webkit-hyphens: auto;"
            : "";
        // Build font fallback chain: document font → locale-aware CJK equivalents → generic.
        // GetCjkFontFallback already weaves in the locale's CJK chain (or empty if
        // the document is locale-neutral); we terminate with -apple-system + sans-serif
        // so the OS picks a system default rather than a hardcoded script.
        var docFont = CssSanitize(dd.Font);
        var cjkFallback = GetCjkFontFallback(docFont, _eastAsiaLang, _themeCjkFont);
        var font = $"\'{docFont}\'{cjkFallback}, -apple-system, sans-serif";
        var pageH = $"{pg.HeightPt:0.#}pt";
        var pageW = $"{pg.WidthPt:0.#}pt";
        var sz = $"{dd.SizePt:0.##}pt";
        // Use docGrid linePitch as line-height when available (CJK snap-to-grid)
        var lh = dd.GridLinePitchPt > 0 ? $"{dd.GridLinePitchPt:0.##}pt" : $"{dd.LineHeight:0.##}";

        return $@"
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ background: #f0f0f0; font-family: {font}; color: {dd.Color}; padding: 20px; }}
        .page-wrapper {{ margin: 0 auto 40px; transition: width 0.15s ease, height 0.15s ease; }}
        .page {{ margin: 0 auto; padding: {mT} {mR} {mB} {mL};
            box-shadow: 0 2px 8px rgba(0,0,0,0.15); border-radius: 4px;
            min-height: {pageH}; line-height: {lh}; font-size: {sz}; position: relative; overflow-x: auto;
            display: flex; flex-direction: column; font-kerning: none; letter-spacing: 0;
            transform-origin: left top; transition: transform 0.15s ease;
            isolation: isolate;
            }}
        /* The white page fill lives on a pseudo-element behind everything so a
           behind-text float (z-index:-1) paints ON the page, not under it. A
           background directly on .page would sit at the stacking-context root and
           hide any negative-z-index child (watermark/behind-doc image). */
        .page::before {{ content: ''; position: absolute; inset: 0; background: white;
            border-radius: 4px; z-index: -2; }}
        /* break-word (not anywhere): a Latin word is only broken when it cannot
           fit on a line BY ITSELF; an oversized word beside a float first wraps
           to the next line. anywhere would break mid-word (produc-t) whenever
           the word does not fit the current inline gap, which is wrong for Latin.
           Table cells still need anywhere (see th,td rule below) so the R32
           fixed-grid column min-content collapses and long content wraps inside
           its column instead of overflowing the page. */
        .page-body {{ flex: 1; display: flex; flex-direction: column; text-autospace: ideograph-alpha ideograph-numeric; overflow-wrap: break-word; {hyphensCss} }}
        /* Multi-column sections: flex ignores column-count; switch to block. */
        .page-body[style*=""column-count""] {{ display: block; }}
        /* A table is typically full text-column width; inside a multi-column
           section it cannot fit one narrow column and would overflow into and
           overprint the adjacent column. Let tables span all columns (Word
           renders a full-width table across the section, with body text
           flowing in columns above/below it). */
        [style*=""column-count""] > table {{ column-span: all; }}
        /* Continuation page-bodies (created by pagination JS when content
           overflows): the segment leader was already at its computed offset
           in the source body, so its server-rendered margin-top must be
           zeroed when it becomes :first-child of a new page-body. The
           ORIGINAL page-body (which holds the document's first paragraph)
           is intentionally not matched here, so its first-paragraph
           spaceBefore renders the way Word emits it. */
        .page-body-cont > :first-child {{ margin-top: 0 !important; }}
        .page-body > img + h1, .page-body > img + img + h1 {{ margin-top: 0 !important; }}
        .doc-header, .doc-footer {{ font-size: {dd.SizePt:0.##}pt; }}
        /* Word paints the header/footer in a layer BEHIND the main body text
           (they are background bands, not foreground content). The header/footer
           is position:absolute, so without a z-index it would paint ABOVE the
           in-flow .page-body (positioned elements paint over non-positioned
           siblings at the same z-auto level). A full-bleed cover banner floated
           into the header would then occlude the body text on every page. Pin
           the band to z-index:-1 so body text (z-auto) paints on top of it, yet
           it stays ABOVE the white page fill (.page::before at z-index:-2). This
           also makes the cover-page white title overlay the banner correctly. */
        .doc-header {{ position: absolute; top: {pg.HeaderDistancePt:0.#}pt; left: {mL}; right: {mR};
            padding-bottom: 0.3em; z-index: -1; }}
        .doc-footer {{ position: absolute; bottom: {pg.FooterDistancePt:0.#}pt; left: {mL}; right: {mR};
            padding-top: 0.3em; z-index: -1; }}
        h1, h2, h3, h4, h5, h6 {{ line-height: {FontMetricsReader.GetRatio(dd.Font) * dd.LineHeight:0.####}; }}
        p {{ margin: 0; margin-bottom: {(dd.SpaceAfterPt > 0 ? $"{dd.SpaceAfterPt:0.##}pt" : "0")}; line-height: {FontMetricsReader.GetRatio(dd.Font) * dd.LineHeight:0.####}; text-align: {dd.DefaultAlign};{(dd.DefaultAlign == "justify" ? " text-justify: inter-character;" : "")} text-autospace: ideograph-alpha ideograph-numeric; }}
        a {{ color: #2B579A; }} a:hover {{ color: #1a3c6e; }}
        .toc {{ display: flex; text-indent: 0 !important; }}
        .toc a {{ color: inherit; text-decoration: none; display: flex; flex: 1; }}
        .toc a span {{ color: inherit !important; text-decoration: none !important; }}
        /* TOC entries authored as a fldChar field (HYPERLINK \l ... between
           begin/separate/end) render as plain spans, NOT wrapped in <a>. Word
           does not apply the Hyperlink character-style color/underline to a
           TOC field's internal links — entries take the toc-N paragraph color
           (black/auto by default). Mirror the .toc a span suppression for the
           un-wrapped case so the Hyperlink rStyle blue does not leak through.
           color:inherit recovers an explicit toc-N paragraph/style color (e.g.
           a styled toc2) since that color lands on the .toc <p> itself. */
        .toc > span {{ color: inherit !important; text-decoration: none !important; }}
        .dot-leader {{ flex: 1; border-bottom: 1px dotted #000; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        .hyphen-leader {{ flex: 1; border-bottom: 1px dashed #000; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        .underscore-leader {{ flex: 1; border-bottom: 1px solid #000; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        .middledot-leader {{ flex: 1; border-bottom: 2px dotted #555; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        /* CONSISTENCY(run-special-content): w:ptab anchors header/footer
           left/center/right alignment regions. The paragraph carrying
           ptabs becomes a flex container so .ptab-spacer (and the leader
           variants above) can flex-grow to push siblings apart. */
        p.has-ptab, div.has-ptab {{ display: flex; align-items: baseline; flex-wrap: wrap; }}
        /* TOC-style <w:tab> paragraphs (center/right tab + dot leader): need a
           flex container so the .dot-leader span (flex:1) stretches and the
           trailing page-number segment lands at the right edge. */
        p.has-leader-tab, div.has-leader-tab {{ display: flex; align-items: baseline; }}
        /* Three-part Left-tab-Center-tab-Right header/paragraph: the
           paragraph is a no-wrap flex row and each .atab-band flex-grows,
           text-aligned (left/center/right) per its own tab stop's Val. nowrap
           keeps all bands on one line (Word never wraps these).
           flex-basis is `auto` (band's intrinsic content width), not `0`
           (forced equal thirds): when every band is short the free space splits
           ~evenly (grow:1 each) so Center/Right still land mid/right exactly
           like a three-part header, but a long band (a TOC entry's full title)
           grows to fit its content and pushes its neighbours rather than being
           capped to a third. No overflow:hidden / text-overflow:ellipsis — a
           tab advances the pen to AT LEAST the stop and over-long content
           simply extends past it; Word never clips at a tab stop. Same
           ''don't clip body text at a tab'' principle as the positional-tab
           min-width path. */
        p.has-aligned-tab, div.has-aligned-tab {{ display: flex; align-items: baseline; flex-wrap: nowrap; }}
        .atab-band {{ flex: 1 1 auto; min-width: 0; white-space: nowrap; }}
        .ptab-spacer {{ flex: 1; min-width: 1em; }}
        ul, ol {{ padding-left: 2em; margin: 0; }}
        ul {{ list-style-type: disc; }}
        li {{ margin: 0; }}
        /* OOXML §17.3.1.36 dropCap: the framed <p> hosts a float clipped
           to the visible cap-glyph height; the inner run <span>'s own
           line-height must defer so its natural strut doesn't expand
           the float past that clip. */
        .dropcap-wrap > p:first-child > span {{ line-height: inherit !important; }}
        .equation {{ text-align: center; padding: 0.5em 0; overflow-x: auto; }}
        img {{ max-width: 100%; height: auto; }}
        img {{ writing-mode: horizontal-tb; }}
        .img-error {{ color: #999; font-style: italic; }}
        table {{ border-collapse: collapse; font-size: {sz}; }}
        td.tsf span, td.tsf div {{ font-size: inherit !important; color: inherit !important; text-align: inherit !important; }}
        .wg {{ margin: 0.3em 0; }}
        .wg p {{ padding: 0; margin: 0.05em 0; }}
        table.borderless {{ border: none; }}
        table.borderless td, table.borderless th {{ border: none; padding: 2px 6px; }}
        /* Default tcMar: Word's TableNormal style is top=0 left=108 bottom=0
           right=108 (twips), so 0pt T/B and 5.4pt L/R. Per-cell tcMar (read
           from tcPr/tcMar) overrides this via inline style. */
        th, td {{ border: none; padding: 0 5.4pt; text-align: inherit; vertical-align: top; break-inside: auto; overflow-wrap: anywhere; }}
        tr {{ break-inside: auto; }}
        th {{ font-weight: 600; }}
        /* #342: comments — highlight the commented range + a hover-able marker. */
        .w-comment-range {{ background: #FFF3CD; border-bottom: 1px dotted #E0A800; }}
        .w-comment-ref {{ font-size: 0.7em; color: #B8860B; cursor: help; margin: 0 1px; user-select: none; }}
        @media print {{ body {{ background: white; padding: 0; }}
            .page {{ box-shadow: none; margin: 0; max-width: none; transform: none !important; }}
            hr.page-break {{ page-break-after: always; border: none; margin: 0; }} }}";
    }

    /// <summary>
    /// Get a platform-specific CJK font fallback fragment for the given
    /// document font. Returned string is prefixed with ", " when non-empty,
    /// so callers can append it directly after the primary font.
    ///
    /// Resolution order:
    ///   1. Style-specific match on the font name itself (e.g. 宋体 → Songti SC).
    ///      These mappings preserve the typographic style across platforms.
    ///   2. Theme's CJK font (from supplemental font list) — if present.
    ///   3. Locale-driven CJK chain via <see cref="LocaleFontRegistry"/>:
    ///      uses <paramref name="eastAsiaLang"/> if declared, otherwise
    ///      tries to detect locale from the font name itself.
    ///   4. Empty — let the OS pick (the body CSS terminates with sans-serif).
    /// </summary>
    private static string GetCjkFontFallback(string docFont, string? eastAsiaLang = null, string? themeCjkFont = null)
    {
        var lower = docFont.ToLowerInvariant();
        // Style-specific Chinese matches — preserve serif/sans/handwriting style.
        if (lower.Contains("宋") || lower.Contains("song") || lower == "simsun")
            return ", 'Songti SC', 'STSong'";
        if (lower.Contains("黑") || lower.Contains("hei") || lower == "simhei")
            return ", 'PingFang SC', 'STHeiti'";
        if (lower.Contains("楷") || lower.Contains("kai"))
            return ", 'Kaiti SC', 'STKaiti'";
        if (lower.Contains("仿宋") || lower.Contains("fangsong"))
            return ", 'STFangsong'";
        // Style-specific Japanese matches.
        if (lower.Contains("明朝") || lower.Contains("mincho"))
            return ", 'Hiragino Mincho ProN', 'Yu Mincho', 'MS Mincho'";
        if (lower.Contains("ゴシック") || lower.Contains("gothic") || lower == "ms gothic" || lower == "yu gothic")
            return ", 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Yu Gothic'";
        // Style-specific Korean matches.
        if (lower.Contains("바탕") || lower == "batang" || lower == "batangche")
            return ", 'Apple SD Gothic Neo', 'Malgun Gothic', 'Batang'";
        if (lower.Contains("굴림") || lower == "gulim" || lower == "dotum" || lower == "malgun gothic")
            return ", 'Apple SD Gothic Neo', 'Malgun Gothic'";

        // Generic Latin/western fonts — use locale (declared or detected) to
        // pick the appropriate CJK fallback chain. Without a locale signal,
        // return empty so the body's terminal sans-serif handles it.
        bool isWestern = lower is "calibri" or "arial" or "helvetica" or "verdana" or "segoe ui"
            or "tahoma" or "trebuchet ms" or "times new roman" or "cambria" or "georgia"
            or "garamond" or "book antiqua" or "palatino linotype";
        if (!isWestern) return "";

        // Theme-resolved CJK font (from supplemental font list) goes first.
        // CssSanitize is required: theme1.xml is attacker-controlled and the
        // value interpolates into font-family.
        var safeTheme = !string.IsNullOrEmpty(themeCjkFont) ? CssSanitize(themeCjkFont) : "";
        var prefix = !string.IsNullOrEmpty(safeTheme) ? $", '{safeTheme}'" : "";

        // Resolve locale: explicit eastAsia lang wins; otherwise probe the
        // theme font name (zh themes typically declare a Chinese typeface).
        var locale = eastAsiaLang;
        if (string.IsNullOrEmpty(locale))
            locale = LocaleFontRegistry.DetectLocaleFromCjkFontName(themeCjkFont);

        var chain = LocaleFontRegistry.GetCjkCssFallback(locale);
        return string.IsNullOrEmpty(chain) ? prefix : prefix + ", " + chain;
    }
}
