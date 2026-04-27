// Copyright 2025 OfficeCli (officecli.ai)
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

        return "";
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
        var widthPx = w != null && long.TryParse(w, out var emu) ? Math.Max(1, emu / 12700.0) : 1;

        return $"border:{widthPx:0.#}px solid {color ?? "#000"}";
    }

    // ==================== Color Math Helpers ====================

    /// <summary>Apply themeTint/themeShade to a base theme color hex.</summary>
    private static string ApplyTintShade(string hex, string? tintHex, string? shadeHex)
    {
        if (hex.Length < 6) return $"#{hex}";
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

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

    private string GetParagraphInlineCss(Paragraph para, bool isListItem = false)
    {
        var parts = new List<string>();

        // Set paragraph font-size to match the first run's resolved font-size.
        // This prevents the CSS "strut" (block container's anonymous inline box) from inflating
        // the line box when .page font-size differs from the actual text span font-size.
        var firstRun = para.Elements<Run>().FirstOrDefault(r =>
            r.ChildElements.Any(c => c is Text t && !string.IsNullOrEmpty(t.Text)));
        if (firstRun != null)
        {
            var rProps = ResolveEffectiveRunProperties(firstRun, para);
            var sz = rProps.FontSize?.Val?.Value;
            if (sz != null && int.TryParse(sz, out var hp))
                parts.Add($"font-size:{hp / 2.0:0.##}pt");
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
            parts.Add("direction:rtl");

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

            // Hanging indent needs left padding/margin equal to the hanging
            // amount to produce the visual effect (first line at 0, follow
            // lines indented). When only `hanging` is set without `left`,
            // use hanging as the left margin too.
            double? hangPt = null;
            if (indHanging is string hpTwips && hpTwips != "0")
                hangPt = Units.TwipsToPt(hpTwips);
            double leftPt = 0;
            if (indLeft is string leftTwips && leftTwips != "0")
                leftPt = Units.TwipsToPt(leftTwips);
            // When hanging is set and left is 0, promote hanging into left
            // margin so subsequent lines visibly indent.
            if (hangPt.HasValue && leftPt == 0) leftPt = hangPt.Value;
            if (leftPt != 0)
                parts.Add($"margin-left:{leftPt:0.##}pt");
            if (indRight is string rightTwips && rightTwips != "0")
                parts.Add($"margin-right:{Units.TwipsToPt(rightTwips):0.##}pt");
            if (!hasDropCap)
            {
                if (indFirstLine is string firstLineTwips && firstLineTwips != "0")
                    parts.Add($"text-indent:{Units.TwipsToPt(firstLineTwips):0.##}pt");
                if (hangPt.HasValue)
                    parts.Add($"text-indent:-{hangPt.Value:0.##}pt");
            }
        }

        // Spacing — direct properties first, fallback to style chain per-property
        var spacing = pProps.SpacingBetweenLines;
        var styleSpacing = ResolveSpacingFromStyle(styleId);
        if (spacing == null)
            spacing = styleSpacing;

        // In Word, paragraph before/after spacing is rendered INSIDE borders.
        // Use padding instead of margin when the paragraph has borders.
        var hasBorders = pProps.ParagraphBorders != null;
        var vSpacingPropBefore = hasBorders ? "padding-top" : "margin-top";
        var vSpacingPropAfter = hasBorders ? "padding-bottom" : "margin-bottom";

        if (spacing != null)
        {
            // contextualSpacing: when enabled and adjacent paragraph has the same style,
            // spaceBefore/spaceAfter between them is suppressed (set to zero).
            var hasContextualSpacing = pProps.ContextualSpacing != null
                || ResolveContextualSpacingFromStyle(styleId);
            var prevPara = para.PreviousSibling<Paragraph>();
            var nextPara = para.NextSibling<Paragraph>();
            var prevStyleId = prevPara?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var nextStyleId = nextPara?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            bool suppressBefore = hasContextualSpacing && prevPara != null
                && (prevStyleId ?? "") == (styleId ?? "");
            bool suppressAfter = hasContextualSpacing && nextPara != null
                && (nextStyleId ?? "") == (styleId ?? "");

            // Before: try direct, then style fallback (before in twips, beforeLines in hundredths of a line)
            var beforeVal = pProps.SpacingBetweenLines?.Before?.Value
                            ?? styleSpacing?.Before?.Value;
            var beforeLinesVal = pProps.SpacingBetweenLines?.BeforeLines?.Value
                                 ?? styleSpacing?.BeforeLines?.Value;

            // Word collapses adjacent spaceBefore/spaceAfter: max(prev.after, cur.before)
            // instead of adding them. CSS flexbox doesn't collapse margins, so we subtract
            // the overlap from spaceBefore when the previous sibling has spaceAfter.
            double prevSpaceAfterPt = 0;
            if (prevPara != null && !suppressBefore)
            {
                var prevPProps = prevPara.ParagraphProperties;
                var prevSId = prevPProps?.ParagraphStyleId?.Val?.Value;
                var prevStyleSpacing = ResolveSpacingFromStyle(prevSId);
                var prevAfter = prevPProps?.SpacingBetweenLines?.After?.Value
                                ?? prevStyleSpacing?.After?.Value;
                if (prevAfter is string pa && int.TryParse(pa, out var paTwips))
                    prevSpaceAfterPt = paTwips / 20.0;
            }

            if (suppressBefore)
                parts.Add($"{vSpacingPropBefore}:0");
            else if (beforeVal is string beforeTwips)
            {
                double beforePt = Units.TwipsToPt(beforeTwips);
                // Collapse: effective spaceBefore = max(0, spaceBefore - prevSpaceAfter)
                if (prevSpaceAfterPt > 0)
                    beforePt = Math.Max(0, beforePt - prevSpaceAfterPt);
                if (beforePt > 0)
                    parts.Add($"{vSpacingPropBefore}:{beforePt:0.##}pt");
            }
            else if (beforeLinesVal is int beforeLines)
                parts.Add($"{vSpacingPropBefore}:{beforeLines / 100.0:0.##}em");

            // After: try direct, then style fallback (after in twips, afterLines in hundredths of a line)
            var afterVal = pProps.SpacingBetweenLines?.After?.Value
                           ?? styleSpacing?.After?.Value;
            var afterLinesVal = pProps.SpacingBetweenLines?.AfterLines?.Value
                                ?? styleSpacing?.AfterLines?.Value;
            if (suppressAfter)
                parts.Add($"{vSpacingPropAfter}:0");
            else if (afterVal is string afterTwips)
                parts.Add($"{vSpacingPropAfter}:{Units.TwipsToPt(afterTwips):0.##}pt");
            else if (afterLinesVal is int afterLines)
                parts.Add($"{vSpacingPropAfter}:{afterLines / 100.0:0.##}em");

            // Line: try direct, then style fallback
            var lineVal = pProps.SpacingBetweenLines?.Line?.Value
                          ?? styleSpacing?.Line?.Value;
            if (lineVal is string lv)
            {
                var rule = pProps.SpacingBetweenLines?.LineRule?.InnerText
                           ?? styleSpacing?.LineRule?.InnerText;
                if (rule == "auto" || rule == null)
                {
                    if (int.TryParse(lv, out var lvNum))
                    {
                        // Correct for font metrics ratio
                        var paraFont = ResolveParaFontForLineHeight(para);
                        var ratio = FontMetricsReader.GetRatio(paraFont);
                        parts.Add($"line-height:{lvNum / 240.0 * ratio:0.##}");
                    }
                }
                else if (rule == "exact" || rule == "atLeast")
                {
                    var linePt = Units.TwipsToPt(lv);
                    parts.Add($"line-height:{linePt:0.##}pt");
                    // #7b0001: when lineRule=exact pins the line box below
                    // ~120% of the paragraph's font size, Word clips
                    // over-tall glyphs. Emit overflow:hidden so tall glyphs
                    // don't leak into neighboring lines.
                    if (rule == "exact")
                    {
                        var sizeStr = ResolveStyleFontSize(
                            para.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "")
                            ?? $"{ReadDocDefaults().SizePt}pt";
                        // ResolveStyleFontSize returns "Npt"; strip suffix.
                        if (sizeStr.EndsWith("pt", StringComparison.Ordinal)
                            && double.TryParse(sizeStr[..^2],
                                System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out var runSizePt)
                            && runSizePt > 0 && linePt < runSizePt * 1.2)
                            parts.Add("overflow:hidden");
                    }
                }
            }

            // If no explicit line-height was set, use font metrics ratio
            if (!parts.Any(p => p.StartsWith("line-height")))
            {
                var paraFont = ResolveParaFontForLineHeight(para);
                var ratio = FontMetricsReader.GetRatio(paraFont);
                if (ratio > 1.01 || ratio < 0.99) // only if meaningfully different from 1.0
                    parts.Add($"line-height:{ratio:0.##}");
            }

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

        // Borders
        var pBdr = pProps.ParagraphBorders;
        if (pBdr != null)
        {
            RenderBorderCss(parts, pBdr.TopBorder, "border-top");
            RenderBorderCss(parts, pBdr.BottomBorder, "border-bottom");
            RenderBorderCss(parts, pBdr.LeftBorder, "border-left");
            RenderBorderCss(parts, pBdr.RightBorder, "border-right");
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
                var lines = framePr.GetAttributes().FirstOrDefault(a => a.LocalName == "lines").Value;
                var lineCount = lines != null && int.TryParse(lines, out var lc) ? lc : 3;
                // Don't override font-size — let the run's actual size (e.g. 58.5pt) apply
                // Set line-height to match lineCount lines of body text
                // Estimate body line height from document defaults
                var defSz = para.Ancestors<Body>().FirstOrDefault()
                    ?.GetFirstChild<SectionProperties>() != null ? 11.0 : 11.0; // fallback
                var rPr = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults
                    ?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                if (rPr?.FontSize?.Val?.Value is string dsz && double.TryParse(dsz, out var dhp))
                    defSz = dhp / 2.0;
                var defSpacing = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults
                    ?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines;
                var lineHMult = 1.15;
                if (defSpacing?.Line?.Value is string dlv && double.TryParse(dlv, out var dlvi)
                    && defSpacing.LineRule?.InnerText is "auto" or null)
                    lineHMult = dlvi / 240.0;
                var bodyLineH = defSz * lineHMult;
                var dropCapHeight = lineCount * bodyLineH;
                // Read hSpace from framePr (OOXML spec default: 0)
                var hSpaceAttr = framePr.GetAttributes().FirstOrDefault(a => a.LocalName == "hSpace").Value;
                var hSpacePt = hSpaceAttr != null && int.TryParse(hSpaceAttr, out var hsTwips) ? hsTwips / 20.0 : 0;
                parts.Add("float:left");
                parts.Add($"line-height:{dropCapHeight:0.#}pt");
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
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
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
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
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
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
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
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            var tabs = style.StyleParagraphProperties?.Tabs?.Elements<TabStop>();
            if (tabs != null && tabs.Any()) return tabs;
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private SpacingBetweenLines? ResolveSpacingFromStyle(string? styleId)
    {
        // If no explicit style, use the default paragraph style (Normal)
        if (styleId == null)
        {
            var defaultStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
            if (defaultStyle?.StyleParagraphProperties?.SpacingBetweenLines != null)
                return defaultStyle.StyleParagraphProperties.SpacingBetweenLines;
            return null;
        }
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            var sp = style.StyleParagraphProperties?.SpacingBetweenLines;
            if (sp != null) return sp;
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>Resolve contextualSpacing from the style chain.</summary>
    private bool ResolveContextualSpacingFromStyle(string? styleId)
    {
        if (styleId == null)
        {
            var defaultStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
            return defaultStyle?.StyleParagraphProperties?.ContextualSpacing != null;
        }
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            if (style.StyleParagraphProperties?.ContextualSpacing != null) return true;
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return false;
    }

    /// <summary>
    /// Resolve Indentation from the style chain (basedOn walk).
    /// </summary>
    private Indentation? ResolveIndentationFromStyle(string? styleId)
    {
        if (styleId == null)
        {
            var defaultStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
            return defaultStyle?.StyleParagraphProperties?.Indentation;
        }
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            var ind = style.StyleParagraphProperties?.Indentation;
            if (ind != null) return ind;
            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
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
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
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
                    if (!parts.Any(p => p.StartsWith("margin-top")))
                    {
                        if (spacing.Before?.Value is string b && b != "0")
                            parts.Add($"margin-top:{Units.TwipsToPt(b):0.##}pt");
                        else if (spacing.BeforeLines?.Value is int bl && bl != 0)
                            parts.Add($"margin-top:{bl / 100.0:0.##}em");
                    }
                    if (!parts.Any(p => p.StartsWith("margin-bottom")))
                    {
                        if (spacing.After?.Value is string a)
                            parts.Add($"margin-bottom:{Units.TwipsToPt(a):0.##}pt");
                        else if (spacing.AfterLines?.Value is int al)
                            parts.Add($"margin-bottom:{al / 100.0:0.##}em");
                    }
                    if (spacing.Line?.Value is string lv && !parts.Any(p => p.StartsWith("line-height")))
                    {
                        var rule = spacing.LineRule?.InnerText;
                        if ((rule == "auto" || rule == null) && int.TryParse(lv, out var val))
                            parts.Add($"line-height:{val / 240.0:0.##}");
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
        return string.Join(";", parts);
    }

    private string GetRunInlineCss(RunProperties? rProps)
    {
        if (rProps == null) return "";
        var parts = new List<string>();

        // Font
        var fonts = rProps.RunFonts;
        var font = fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
        // Skip theme font references (e.g. "+mn-lt", "+mj-ea") — those are shorthand
        // markers, not real font names; the theme-resolved value would already be in
        // AsciiTheme etc. which we don't read here.
        if (font != null && !font.StartsWith("+", StringComparison.Ordinal))
        {
            var fallback = GetChineseFontFallback(font);
            // Always append a generic family so the run still renders with the right
            // serif/sans-serif class when neither the primary nor the CJK fallback
            // is installed (matters in headless browsers like Playwright).
            var generic = IsLikelySerif(font) ? "serif" : "sans-serif";
            parts.Add(fallback != null
                ? $"font-family:'{CssSanitize(font)}',{fallback},{generic}"
                : $"font-family:'{CssSanitize(font)}',{generic}");
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
        }

        // Strikethrough (single or double)
        var hasSingleStrike = rProps.Strike != null && (rProps.Strike.Val == null || rProps.Strike.Val.Value);
        var hasDoubleStrike = rProps.DoubleStrike != null && (rProps.DoubleStrike.Val == null || rProps.DoubleStrike.Val.Value);
        if (hasSingleStrike || hasDoubleStrike)
        {
            var existing = parts.FirstOrDefault(p => p.StartsWith("text-decoration:"));
            if (existing != null)
            {
                parts.Remove(existing);
                parts.Add(existing + " line-through");
            }
            else
            {
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

        // Character scale (w:w, horizontal stretch as a percentage). Use inline-block +
        // transform scaleX so rendering width actually changes — transform alone collapses
        // space reservation. Default/unit value 100% → skip.
        var charScale = rProps.CharacterScale?.Val?.Value;
        if (charScale.HasValue && charScale.Value > 0 && charScale.Value != 100)
        {
            var ratio = charScale.Value / 100.0;
            parts.Add($"display:inline-block;transform:scaleX({ratio:0.##});transform-origin:left");
        }

        // Color: w:color val + themeColor with tint/shade. Route through
        // ResolveRunColor for consistency with conditional-format and border
        // paths. Val wins if not "auto"; else fall through to themeColor.
        var resolvedColor = ResolveRunColor(rProps.Color);
        if (resolvedColor != null)
        {
            parts.Add($"color:{resolvedColor}");
        }

        // Highlight
        var highlight = rProps.Highlight?.Val?.InnerText;
        if (highlight != null)
        {
            var hlColor = HighlightToCssColor(highlight);
            if (hlColor != null) parts.Add($"background-color:{hlColor}");
        }

        // Superscript / Subscript — always shrink to match Word's behavior.
        // Word auto-sizes sub/sup relative to the surrounding run, even when
        // the run has an explicit size. Use font-size:smaller (browser spec
        // default for <sub>/<sup>) so the shrinkage compounds with any
        // explicit size we already emitted for this run.
        var vertAlign = rProps.VerticalTextAlignment?.Val;
        if (vertAlign != null)
        {
            if (vertAlign.InnerText == "superscript")
                parts.Add("vertical-align:super;font-size:smaller");
            else if (vertAlign.InnerText == "subscript")
                parts.Add("vertical-align:sub;font-size:smaller");
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
            var fill = runShd.Fill?.Value;
            if (fill != null && fill != "auto" && IsHexColor(fill))
                parts.Add($"background-color:#{fill}");
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
            parts.Add("direction:rtl;unicode-bidi:embed");

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
                    var widthPt = Math.Max(0.5, widthEmu / 12700.0);
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
                    var blurPx = blurEmu / 12700.0 * 1.333;
                    var distEmu = attrs.TryGetValue("dist", out var dist) && long.TryParse(dist, out var distLong) ? distLong : 0;
                    var dirVal = attrs.TryGetValue("dir", out var dir) && long.TryParse(dir, out var dirLong) ? dirLong : 0;
                    var angleRad = dirVal / 60000.0 * Math.PI / 180.0;
                    var distPx = distEmu / 12700.0 * 1.333;
                    var xPx = distPx * Math.Sin(angleRad);
                    var yPx = distPx * Math.Cos(angleRad);
                    var alphaMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"alpha[^>]*val=""(\d+)""");
                    if (alphaMatch.Success && double.TryParse(alphaMatch.Groups[1].Value, out var alphaVal) && alphaVal < 100000)
                        color = HexToRgba(color, alphaVal / 100000.0);
                    textShadows.Add($"{xPx:0.#}px {yPx:0.#}px {blurPx:0.#}px {color}");
                    break;
                }
                case "glow":
                {
                    var radAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "rad");
                    var radiusEmu = long.TryParse(radAttr.Value, out var r) ? r : 0;
                    var radiusPx = radiusEmu / 12700.0 * 1.333;
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
        var distPx = distEmu / 12700.0 * 1.333;
        var blurPx = blurEmu / 12700.0 * 1.333;

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

    private string GetTableCellInlineCss(TableCell cell, bool tableBordersNone, TableBorders? tblBorders = null,
        Dictionary<string, TableConditionalFormat>? condFormats = null, List<string>? condTypes = null,
        int rowIdx = 0, int colIdx = 0, int totalRows = 1, int totalCols = 1)
    {
        var parts = new List<string>();
        var tcPr = cell.TableCellProperties;

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

        if (tcPr == null) return string.Join(";", parts);

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
            if (!IsBorderNone(tcBorders.TopBorder)) { parts.RemoveAll(p => p.StartsWith("border-top:")); RenderBorderCss(parts, tcBorders.TopBorder, "border-top"); }
            if (!IsBorderNone(tcBorders.BottomBorder)) { parts.RemoveAll(p => p.StartsWith("border-bottom:")); RenderBorderCss(parts, tcBorders.BottomBorder, "border-bottom"); }
            if (!IsBorderNone(tcBorders.LeftBorder)) { parts.RemoveAll(p => p.StartsWith("border-left:")); RenderBorderCss(parts, tcBorders.LeftBorder, "border-left"); }
            if (!IsBorderNone(tcBorders.RightBorder)) { parts.RemoveAll(p => p.StartsWith("border-right:")); RenderBorderCss(parts, tcBorders.RightBorder, "border-right"); }
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
                "btLr" => "vertical-rl;transform:rotate(180deg)", // read bottom-up
                "tbRl" => "vertical-rl",                            // read top-down
                "lrTb" or null => null,                             // default horizontal
                _ => null,
            };
            if (wm != null) parts.Add($"writing-mode:{wm}");
        }

        // Cell noWrap — prevents content wrapping within the cell
        if (tcPr.NoWrap != null)
            parts.Add("white-space:nowrap");

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

        // Padding — add vertical compensation for CSS line-height:1 clipping glyph ascenders
        const double CellPadVComp = 3.0; // pt
        var margins = tcPr?.TableCellMargin;
        {
            var padTop = Units.TwipsToPt(margins?.TopMargin?.Width?.Value ?? "0") + CellPadVComp;
            var padBot = Units.TwipsToPt(margins?.BottomMargin?.Width?.Value ?? "0") + CellPadVComp;
            var leftVal = margins?.LeftMargin?.Width?.Value ?? margins?.StartMargin?.Width?.Value;
            var rightVal = margins?.RightMargin?.Width?.Value ?? margins?.EndMargin?.Width?.Value;
            var padLeft = leftVal != null ? $"{Units.TwipsToPt(leftVal):0.#}pt" : "5.4pt";
            var padRight = rightVal != null ? $"{Units.TwipsToPt(rightVal):0.#}pt" : "5.4pt";
            parts.Add($"padding:{padTop:0.#}pt {padRight} {padBot:0.#}pt {padLeft}");
        }

        return string.Join(";", parts);
    }

    // ==================== CSS Helpers ====================

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
        string cssColor;
        if (color != null && !color.Equals("auto", StringComparison.OrdinalIgnoreCase)
            && IsHexColor(color))
        {
            cssColor = $"#{color}";
        }
        else
        {
            var themeColor = border.GetAttributes().FirstOrDefault(a => a.LocalName == "themeColor").Value;
            if (themeColor != null && GetThemeColors().TryGetValue(themeColor, out var tcHex))
            {
                var tint = border.GetAttributes().FirstOrDefault(a => a.LocalName == "themeTint").Value;
                var shade = border.GetAttributes().FirstOrDefault(a => a.LocalName == "themeShade").Value;
                cssColor = ApplyTintShade(tcHex, tint, shade);
            }
            else
            {
                cssColor = "#000";
            }
        }

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
        var colorVal = color.Val?.Value;
        if (colorVal != null && colorVal != "auto" && IsHexColor(colorVal))
            return $"#{colorVal}";
        var tcName = color.ThemeColor?.InnerText;
        if (tcName != null && GetThemeColors().TryGetValue(tcName, out var tcHex))
        {
            var tint = color.GetAttributes().FirstOrDefault(a => a.LocalName == "themeTint").Value;
            var shade = color.GetAttributes().FirstOrDefault(a => a.LocalName == "themeShade").Value;
            return ApplyTintShade(tcHex, tint, shade);
        }
        return null;
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
        "darkblue" => "#00008B",
        "darkcyan" => "#008B8B",
        "darkgreen" => "#006400",
        "darkmagenta" => "#8B008B",
        "darkred" => "#8B0000",
        "darkyellow" => "#808000",
        "darkgray" => "#A9A9A9",
        "lightgray" => "#D3D3D3",
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
    private string ResolveParaFontForLineHeight(Paragraph para)
    {
        // Use the first run's ascii font; fall back to document default
        var firstRun = para.Elements<Run>().FirstOrDefault();
        if (firstRun != null)
        {
            var rProps = ResolveEffectiveRunProperties(firstRun, para);
            var font = rProps.RunFonts?.Ascii?.Value ?? rProps.RunFonts?.HighAnsi?.Value;
            if (!string.IsNullOrEmpty(font)) return font;
        }
        // Fall back to document default font, then theme MinorFont, then
        // the Office default. The theme step matters for modern Office docs
        // (e.g. Aptos) where styles.xml omits an explicit ascii font but
        // theme1.xml's fontScheme defines it.
        var defFont = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
            ?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.RunFonts?.Ascii?.Value;
        return defFont ?? GetThemeMinorLatinFont() ?? OfficeDefaultFonts.MinorLatin;
    }

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
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == current);
            if (style == null) break;
            var sz = style.StyleRunProperties?.FontSize?.Val?.Value;
            if (sz != null && int.TryParse(sz, out var halfPts))
                return $"{halfPts / 2.0:0.##}pt";
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
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == current);
            if (style == null) break;
            var cv = style.StyleRunProperties?.Color?.Val?.Value;
            if (cv != null && cv != "auto" && IsHexColor(cv)) return $"#{cv}";
            var tc = style.StyleRunProperties?.Color?.ThemeColor?.InnerText;
            if (tc != null && GetThemeColors().TryGetValue(tc, out var tcHex)) return $"#{tcHex}";
            current = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private string? ResolveStyleIndent(string styleId)
    {
        var visited = new HashSet<string>();
        var current = styleId;
        while (current != null && visited.Add(current))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == current);
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
        // Build font fallback chain: document font → platform-specific CJK equivalents → generic
        var docFont = CssSanitize(dd.Font);
        var cjkFallback = GetCjkFontFallback(docFont, _eastAsiaLang, _themeCjkFont);
        var font = $"\'{docFont}\'{cjkFallback}, \'Microsoft YaHei\', -apple-system, \'PingFang SC\', sans-serif";
        var pageH = $"{pg.HeightPt:0.#}pt";
        var pageW = $"{pg.WidthPt:0.#}pt";
        var sz = $"{dd.SizePt:0.##}pt";
        // Use docGrid linePitch as line-height when available (CJK snap-to-grid)
        var lh = dd.GridLinePitchPt > 0 ? $"{dd.GridLinePitchPt:0.##}pt" : $"{dd.LineHeight:0.##}";

        return $@"
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ background: #f0f0f0; font-family: {font}; color: {dd.Color}; padding: 20px; }}
        .page-wrapper {{ margin: 0 auto 40px; transition: width 0.15s ease, height 0.15s ease; }}
        .page {{ background: white; margin: 0 auto; padding: {mT} {mR} {mB} {mL};
            box-shadow: 0 2px 8px rgba(0,0,0,0.15); border-radius: 4px;
            min-height: {pageH}; line-height: {lh}; font-size: {sz}; position: relative; overflow-x: auto;
            display: flex; flex-direction: column; font-kerning: none; letter-spacing: 0;
            transform-origin: left top; transition: transform 0.15s ease;
            }}
        .page-body {{ flex: 1; display: flex; flex-direction: column; text-autospace: ideograph-alpha ideograph-numeric; overflow-wrap: anywhere; {hyphensCss} }}
        /* Multi-column sections: flex ignores column-count; switch to block. */
        .page-body[style*=""column-count""] {{ display: block; }}
        .page-body > :first-child {{ margin-top: 0 !important; }}
        .page-body > img + h1, .page-body > img + img + h1 {{ margin-top: 0 !important; }}
        .doc-header, .doc-footer {{ font-size: {dd.SizePt:0.##}pt; }}
        .doc-header {{ position: absolute; top: {pg.HeaderDistancePt:0.#}pt; left: {mL}; right: {mR};
            padding-bottom: 0.3em; }}
        .doc-footer {{ position: absolute; bottom: {pg.FooterDistancePt:0.#}pt; left: {mL}; right: {mR};
            padding-top: 0.3em; }}
        h1, h2, h3, h4, h5, h6 {{ line-height: {dd.LineHeight * FontMetricsReader.GetRatio(dd.Font):0.##}; }}
        p {{ margin: 0; margin-bottom: {(dd.SpaceAfterPt > 0 ? $"{dd.SpaceAfterPt:0.##}pt" : "0")}; line-height: {dd.LineHeight * FontMetricsReader.GetRatio(dd.Font):0.##}; text-align: {dd.DefaultAlign};{(dd.DefaultAlign == "justify" ? " text-justify: inter-character;" : "")} text-autospace: ideograph-alpha ideograph-numeric; }}
        p.empty {{ margin: 0; min-height: 1em; }}
        a {{ color: #2B579A; }} a:hover {{ color: #1a3c6e; }}
        .toc {{ display: flex; text-indent: 0 !important; }}
        .toc a {{ color: inherit; text-decoration: none; display: flex; flex: 1; }}
        .toc a span {{ color: inherit !important; text-decoration: none !important; }}
        .dot-leader {{ flex: 1; border-bottom: 1px dotted #000; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        .hyphen-leader {{ flex: 1; border-bottom: 1px dashed #000; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        .underscore-leader {{ flex: 1; border-bottom: 1px solid #000; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        .middledot-leader {{ flex: 1; border-bottom: 2px dotted #555; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
        ul, ol {{ padding-left: 2em; margin: 0.2em 0; }}
        ul {{ list-style-type: disc; }}
        li {{ margin: 0.1em 0; }}
        .equation {{ text-align: center; padding: 0.5em 0; overflow-x: auto; }}
        img {{ max-width: 100%; height: auto; }}
        .img-error {{ color: #999; font-style: italic; }}
        table {{ border-collapse: collapse; font-size: {sz}; }}
        td.tsf span, td.tsf div {{ font-size: inherit !important; color: inherit !important; text-align: inherit !important; }}
        .wg {{ margin: 0.3em 0; }}
        .wg p {{ padding: 0; margin: 0.05em 0; }}
        table.borderless {{ border: none; }}
        table.borderless td, table.borderless th {{ border: none; padding: 2px 6px; }}
        th, td {{ border: none; padding: 3pt 5.4pt; text-align: inherit; vertical-align: top; break-inside: auto; }}
        tr {{ break-inside: auto; }}
        th {{ font-weight: 600; }}
        @media print {{ body {{ background: white; padding: 0; }}
            .page {{ box-shadow: none; margin: 0; max-width: none; transform: none !important; }}
            hr.page-break {{ page-break-after: always; border: none; margin: 0; }} }}";
    }

    /// <summary>Get platform-specific CJK font fallback for the given document font.</summary>
    /// <param name="docFont">The font name from the document.</param>
    /// <param name="eastAsiaLang">Optional EastAsia language code (e.g. "zh-CN", "ja-JP", "ko-KR").</param>
    private static string GetCjkFontFallback(string docFont, string? eastAsiaLang = null, string? themeCjkFont = null)
    {
        var lower = docFont.ToLowerInvariant();
        // Chinese font names — match directly
        if (lower.Contains("宋") || lower.Contains("song") || lower == "simsun")
            return ", 'Songti SC', 'STSong'";
        if (lower.Contains("黑") || lower.Contains("hei") || lower == "simhei")
            return ", 'PingFang SC', 'STHeiti'";
        if (lower.Contains("楷") || lower.Contains("kai"))
            return ", 'Kaiti SC', 'STKaiti'";
        if (lower.Contains("仿宋") || lower.Contains("fangsong"))
            return ", 'STFangsong'";
        // Japanese font names
        if (lower.Contains("明朝") || lower.Contains("mincho"))
            return ", 'Hiragino Mincho ProN', 'Yu Mincho', 'MS Mincho'";
        if (lower.Contains("ゴシック") || lower.Contains("gothic") || lower == "ms gothic" || lower == "yu gothic")
            return ", 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Yu Gothic'";
        // Korean font names
        if (lower.Contains("바탕") || lower == "batang" || lower == "batangche")
            return ", 'Apple SD Gothic Neo', 'Malgun Gothic', 'Batang'";
        if (lower.Contains("굴림") || lower == "gulim" || lower == "dotum" || lower == "malgun gothic")
            return ", 'Apple SD Gothic Neo', 'Malgun Gothic'";

        // Western fonts — use EastAsia language to pick the right CJK fallback
        bool isWestern = lower is "calibri" or "arial" or "helvetica" or "verdana" or "segoe ui"
            or "tahoma" or "trebuchet ms" or "times new roman" or "cambria" or "georgia"
            or "garamond" or "book antiqua" or "palatino linotype";
        if (isWestern)
        {
            // Prefer theme-resolved CJK font (from supplemental font list)
            // CssSanitize the theme font name — theme1.xml is attacker-
            // controlled and this value interpolates into font-family.
            var safeTheme = !string.IsNullOrEmpty(themeCjkFont) ? CssSanitize(themeCjkFont) : "";
            var prefix = !string.IsNullOrEmpty(safeTheme) ? $", '{safeTheme}'" : "";
            var lang = eastAsiaLang?.ToLowerInvariant() ?? "";
            if (lang.StartsWith("ja"))
                return prefix + ", 'Hiragino Mincho ProN', 'Yu Mincho', 'MS Mincho'";
            if (lang.StartsWith("ko"))
                return prefix + ", 'Apple SD Gothic Neo', 'Malgun Gothic', 'Batang'";
            // Default: Chinese (zh or unspecified)
            return prefix + ", 'SimSun', '宋体', 'Songti SC', 'STSong'";
        }
        return "";
    }
}
