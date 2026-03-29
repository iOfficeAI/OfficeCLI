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

    private Dictionary<string, string> GetThemeColors()
    {
        if (_themeColors != null) return _themeColors;

        _themeColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var theme = _doc.MainDocumentPart?.ThemePart?.Theme;
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return _themeColors;

        void Add(string name, OpenXmlCompositeElement? color)
        {
            if (color == null) return;
            var rgb = color.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            var sys = color.GetFirstChild<A.SystemColor>();
            var srgb = sys?.LastColor?.Value;
            var hex = rgb ?? srgb;
            if (hex != null) _themeColors[name] = hex;
        }

        Add("dk1", colorScheme.Dark1Color);
        Add("dk2", colorScheme.Dark2Color);
        Add("lt1", colorScheme.Light1Color);
        Add("lt2", colorScheme.Light2Color);
        Add("accent1", colorScheme.Accent1Color);
        Add("accent2", colorScheme.Accent2Color);
        Add("accent3", colorScheme.Accent3Color);
        Add("accent4", colorScheme.Accent4Color);
        Add("accent5", colorScheme.Accent5Color);
        Add("accent6", colorScheme.Accent6Color);
        Add("hlink", colorScheme.Hyperlink);
        Add("folHlink", colorScheme.FollowedHyperlinkColor);

        // Aliases
        if (_themeColors.TryGetValue("dk1", out var dk1)) { _themeColors["tx1"] = dk1; _themeColors["dark1"] = dk1; }
        if (_themeColors.TryGetValue("lt1", out var lt1)) { _themeColors["bg1"] = lt1; _themeColors["light1"] = lt1; }
        if (_themeColors.TryGetValue("lt2", out var lt2)) { _themeColors["bg2"] = lt2; _themeColors["light2"] = lt2; }

        return _themeColors;
    }

    private string? ResolveSchemeColor(OpenXmlElement schemeColor)
    {
        var schemeName = schemeColor.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (schemeName == null) return null;

        var themeColors = GetThemeColors();
        if (!themeColors.TryGetValue(schemeName, out var hex)) return null;

        // Apply color transforms (lumMod, lumOff, tint, shade)
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        var lumMod = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "lumMod");
        var lumOff = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "lumOff");
        var tint = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "tint");
        var shade = schemeColor.Elements().FirstOrDefault(e => e.LocalName == "shade");

        if (tint != null)
        {
            var t = GetLongAttr(tint, "val") / 100000.0;
            r = (int)(r + (255 - r) * (1 - t));
            g = (int)(g + (255 - g) * (1 - t));
            b = (int)(b + (255 - b) * (1 - t));
        }

        if (shade != null)
        {
            var s = GetLongAttr(shade, "val") / 100000.0;
            r = (int)(r * s);
            g = (int)(g * s);
            b = (int)(b * s);
        }

        if (lumMod != null || lumOff != null)
        {
            var mod = (lumMod != null ? GetLongAttr(lumMod, "val") : 100000) / 100000.0;
            var off = (lumOff != null ? GetLongAttr(lumOff, "val") : 0) / 100000.0;
            RgbToHsl(r, g, b, out var h, out var s, out var l);
            l = Math.Clamp(l * mod + off, 0, 1);
            HslToRgb(h, s, l, out r, out g, out b);
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
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
                if (val != null) return $"background-color:#{val}";
            }
            var scheme = solidFill.Elements().FirstOrDefault(e => e.LocalName == "schemeClr");
            if (scheme != null)
            {
                var color = ResolveSchemeColor(scheme);
                if (color != null) return $"background-color:{color}";
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
        if (rgb != null) color = $"#{rgb.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value}";
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

    private static void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
    {
        var rf = r / 255.0; var gf = g / 255.0; var bf = b / 255.0;
        var max = Math.Max(rf, Math.Max(gf, bf));
        var min = Math.Min(rf, Math.Min(gf, bf));
        var delta = max - min;
        l = (max + min) / 2.0;
        if (delta < 1e-10) { h = 0; s = 0; return; }
        s = l < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);
        if (Math.Abs(max - rf) < 1e-10) h = ((gf - bf) / delta + (gf < bf ? 6 : 0)) / 6.0;
        else if (Math.Abs(max - gf) < 1e-10) h = ((bf - rf) / delta + 2) / 6.0;
        else h = ((rf - gf) / delta + 4) / 6.0;
    }

    private static void HslToRgb(double h, double s, double l, out int r, out int g, out int b)
    {
        if (s < 1e-10) { r = g = b = (int)Math.Round(l * 255); return; }
        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        r = (int)Math.Round(HueToRgb(p, q, h + 1.0 / 3) * 255);
        g = (int)Math.Round(HueToRgb(p, q, h) * 255);
        b = (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3) * 255);
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1; if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
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
            var align = jc.InnerText switch
            {
                "center" => "center",
                "right" or "end" => "right",
                "both" or "distribute" => "justify",
                _ => (string?)null
            };
            if (align != null) parts.Add($"text-align:{align}");
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

            if (indLeft is string leftTwips && leftTwips != "0")
                parts.Add($"margin-left:{Units.TwipsToPt(leftTwips):0.##}pt");
            if (indRight is string rightTwips && rightTwips != "0")
                parts.Add($"margin-right:{Units.TwipsToPt(rightTwips):0.##}pt");
            if (!hasDropCap)
            {
                if (indFirstLine is string firstLineTwips && firstLineTwips != "0")
                    parts.Add($"text-indent:{Units.TwipsToPt(firstLineTwips):0.##}pt");
                if (indHanging is string hangTwips && hangTwips != "0")
                    parts.Add($"text-indent:-{Units.TwipsToPt(hangTwips):0.##}pt");
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
            // Before: try direct, then style fallback (before in twips, beforeLines in hundredths of a line)
            var beforeVal = pProps.SpacingBetweenLines?.Before?.Value
                            ?? styleSpacing?.Before?.Value;
            var beforeLinesVal = pProps.SpacingBetweenLines?.BeforeLines?.Value
                                 ?? styleSpacing?.BeforeLines?.Value;
            if (beforeVal is string beforeTwips && beforeTwips != "0")
                parts.Add($"{vSpacingPropBefore}:{Units.TwipsToPt(beforeTwips):0.##}pt");
            else if (beforeLinesVal is int beforeLines && beforeLines != 0)
                parts.Add($"{vSpacingPropBefore}:{beforeLines / 100.0:0.##}em");

            // After: try direct, then style fallback (after in twips, afterLines in hundredths of a line)
            var afterVal = pProps.SpacingBetweenLines?.After?.Value
                           ?? styleSpacing?.After?.Value;
            var afterLinesVal = pProps.SpacingBetweenLines?.AfterLines?.Value
                                ?? styleSpacing?.AfterLines?.Value;
            if (afterVal is string afterTwips && afterTwips != "0")
                parts.Add($"{vSpacingPropAfter}:{Units.TwipsToPt(afterTwips):0.##}pt");
            else if (afterLinesVal is int afterLines && afterLines != 0)
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
                    if (int.TryParse(lv, out var lvNum) && lvNum != 240)
                        parts.Add($"line-height:{lvNum / 240.0:0.##}");
                }
                else if (rule == "exact" || rule == "atLeast")
                {
                    parts.Add($"line-height:{Units.TwipsToPt(lv):0.##}pt");
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
                parts.Add("float:left");
                parts.Add($"line-height:{dropCapHeight:0.#}pt");
                parts.Add($"padding-right:6px");
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
            if (style == null) break;
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
                        if (spacing.After?.Value is string a && a != "0")
                            parts.Add($"margin-bottom:{Units.TwipsToPt(a):0.##}pt");
                        else if (spacing.AfterLines?.Value is int al && al != 0)
                            parts.Add($"margin-bottom:{al / 100.0:0.##}em");
                    }
                    if (spacing.Line?.Value is string lv && !parts.Any(p => p.StartsWith("line-height")))
                    {
                        var rule = spacing.LineRule?.InnerText;
                        if ((rule == "auto" || rule == null) && int.TryParse(lv, out var val) && val != 240)
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
        if (font != null)
        {
            var fallback = GetChineseFontFallback(font);
            parts.Add(fallback != null
                ? $"font-family:'{CssSanitize(font)}',{fallback}"
                : $"font-family:'{CssSanitize(font)}'");
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

        // Underline
        if (rProps.Underline?.Val != null)
        {
            var ulVal = rProps.Underline.Val.InnerText;
            if (ulVal != "none")
                parts.Add("text-decoration:underline");
        }

        // Strikethrough (single or double)
        var hasStrike = (rProps.Strike != null && (rProps.Strike.Val == null || rProps.Strike.Val.Value))
            || (rProps.DoubleStrike != null && (rProps.DoubleStrike.Val == null || rProps.DoubleStrike.Val.Value));
        if (hasStrike)
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
        }

        // Color: w:color val is the pre-computed color (already has themeColor+themeTint applied).
        // Use val directly; only fall back to theme resolution if val is missing.
        var colorVal = rProps.Color?.Val?.Value;
        if (colorVal != null && colorVal != "auto")
        {
            parts.Add($"color:#{colorVal}");
        }
        else if (rProps.Color?.ThemeColor?.InnerText is string tcName)
        {
            var tc = GetThemeColors();
            if (tc.TryGetValue(tcName, out var tcHex))
            {
                var tint = rProps.Color?.GetAttributes().FirstOrDefault(a => a.LocalName == "themeTint").Value;
                var shade = rProps.Color?.GetAttributes().FirstOrDefault(a => a.LocalName == "themeShade").Value;
                parts.Add($"color:{ApplyTintShade(tcHex, tint, shade)}");
            }
        }

        // Highlight
        var highlight = rProps.Highlight?.Val?.InnerText;
        if (highlight != null)
        {
            var hlColor = HighlightToCssColor(highlight);
            if (hlColor != null) parts.Add($"background-color:{hlColor}");
        }

        // Superscript / Subscript
        var vertAlign = rProps.VerticalTextAlignment?.Val;
        if (vertAlign != null)
        {
            var hasExplicitSize = rProps.FontSize?.Val?.Value != null;
            if (vertAlign.InnerText == "superscript")
                parts.Add(hasExplicitSize ? "vertical-align:super" : "vertical-align:super;font-size:smaller");
            else if (vertAlign.InnerText == "subscript")
                parts.Add(hasExplicitSize ? "vertical-align:sub" : "vertical-align:sub;font-size:smaller");
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
            if (fill != null && fill != "auto")
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
                var color = (bdrColor != null && bdrColor != "auto") ? $"#{bdrColor}" : "#000";
                parts.Add($"border:{px:0.#}px solid {color};padding:0 2px");
            }
        }

        // RTL text direction
        if (rProps.RightToLeftText != null && (rProps.RightToLeftText.Val == null || rProps.RightToLeftText.Val.Value))
            parts.Add("direction:rtl;unicode-bidi:bidi-override");

        return string.Join(";", parts);
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

        // Padding
        var margins = tcPr.TableCellMargin;
        if (margins != null)
        {
            var padTop = margins.TopMargin?.Width?.Value;
            var padBot = margins.BottomMargin?.Width?.Value;
            var padLeft = margins.LeftMargin?.Width?.Value ?? margins.StartMargin?.Width?.Value;
            var padRight = margins.RightMargin?.Width?.Value ?? margins.EndMargin?.Width?.Value;
            if (padTop != null || padBot != null || padLeft != null || padRight != null)
            {
                parts.Add($"padding:{Units.TwipsToPtStr(padTop ?? "0")} {Units.TwipsToPtStr(padRight ?? "0")} {Units.TwipsToPtStr(padBot ?? "0")} {Units.TwipsToPtStr(padLeft ?? "0")}");
            }
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
            "double" => "double",
            "dashed" or "dashSmallGap" => "dashed",
            "dotted" => "dotted",
            _ => "solid"
        };
        var widthPx = sz != null && int.TryParse(sz, out var s) ? Math.Max(1, s / 8.0) : 1.0;
        // CSS double border needs at least 3px to render two visible lines
        if (style == "double" && widthPx < 3) widthPx = 3;
        var width = $"{widthPx:0.#}px";

        // Resolve color: try direct color, then themeColor with tint/shade
        string cssColor;
        if (color != null && !color.Equals("auto", StringComparison.OrdinalIgnoreCase))
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
        if (colorVal != null && colorVal != "auto")
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
    /// Returns CSS fallback fonts for common Windows Chinese fonts that are unavailable on Mac.
    /// </summary>
    private static string? GetChineseFontFallback(string font) => font switch
    {
        "仿宋_GB2312" => "'仿宋',FangSong,STFangsong",
        "楷体_GB2312" => "'楷体',KaiTi,STKaiti",
        "长城小标宋体" => "'华文中宋',STZhongsong,'宋体',SimSun",
        "黑体" => "'Heiti SC',STHeiti",
        _ => null
    };

    private static string CssSanitize(string value) =>
        Regex.Replace(value, @"[""'\\<>&;{}]", "");

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

    // ==================== CSS Stylesheet ====================

    private static string GenerateWordCss(PageLayout pg, DocDef dd)
    {
        // Use pt units (twips/20) for pixel-perfect accuracy — no cm→px conversion loss
        var mL = $"{pg.MarginLeftPt:0.#}pt";
        var mR = $"{pg.MarginRightPt:0.#}pt";
        var mT = $"{pg.MarginTopPt:0.#}pt";
        var mB = $"{pg.MarginBottomPt:0.#}pt";
        // Build font fallback chain: document font → platform-specific CJK equivalents → generic
        var docFont = CssSanitize(dd.Font);
        var cjkFallback = GetCjkFontFallback(docFont);
        var font = $"\'{docFont}\'{cjkFallback}, \'Microsoft YaHei\', -apple-system, \'PingFang SC\', sans-serif";
        var pageH = $"{pg.HeightPt:0.#}pt";
        var pageW = $"{pg.WidthPt:0.#}pt";
        var sz = $"{dd.SizePt:0.##}pt";
        // Use docGrid linePitch as line-height when available (CJK snap-to-grid)
        var lh = dd.GridLinePitchPt > 0 ? $"{dd.GridLinePitchPt:0.##}pt" : $"{dd.LineHeight:0.##}";

        return $@"
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ background: #f0f0f0; font-family: {font}; color: {dd.Color}; padding: 20px; }}
        .page {{ background: white; margin: 0 auto 40px; padding: {mT} {mR} {mB} {mL};
            box-shadow: 0 2px 8px rgba(0,0,0,0.15); border-radius: 4px;
            min-height: {pageH}; line-height: {lh}; font-size: {sz}; position: relative; overflow-x: auto;
            display: flex; flex-direction: column; font-kerning: none; letter-spacing: 0;
            }}
        .page-body {{ flex: 1; }}
        .page-body > :first-child {{ margin-top: 0 !important; }}
        .doc-header, .doc-footer {{ color: #888; font-size: 9pt; }}
        .doc-header {{ position: absolute; top: {pg.HeaderDistancePt:0.#}pt; left: {mL}; right: {mR};
            padding-bottom: 0.3em; }}
        .doc-footer {{ position: absolute; bottom: {pg.FooterDistancePt:0.#}pt; left: {mL}; right: {mR};
            padding-top: 0.3em; }}
        h1, h2, h3, h4, h5, h6 {{ line-height: 1.4; }}
        p {{ margin: 0; text-align: justify; text-justify: inter-character; }}
        p.empty {{ margin: 0; min-height: 1em; }}
        a {{ color: #2B579A; }} a:hover {{ color: #1a3c6e; }}
        .toc {{ display: flex; text-indent: 0 !important; }}
        .toc a {{ color: inherit; text-decoration: none; display: flex; flex: 1; }}
        .toc a span {{ color: inherit !important; text-decoration: none !important; }}
        .dot-leader {{ flex: 1; border-bottom: 1px dotted #000; margin: 0 4px; min-width: 2em; align-self: flex-end; margin-bottom: 0.25em; }}
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
        th, td {{ border: none; padding: 0 5.4pt; text-align: inherit; vertical-align: top; }}
        th {{ font-weight: 600; }}
        @media print {{ body {{ background: white; padding: 0; }}
            .page {{ box-shadow: none; margin: 0; max-width: none; }}
            hr.page-break {{ page-break-after: always; border: none; margin: 0; }} }}";
    }

    /// <summary>Get platform-specific CJK font fallback for the given document font.</summary>
    private static string GetCjkFontFallback(string docFont)
    {
        var lower = docFont.ToLowerInvariant();
        // Song/宋 → serif CJK fonts (macOS: Songti SC / STSong)
        if (lower.Contains("宋") || lower.Contains("song") || lower == "simsun")
            return ", 'Songti SC', 'STSong'";
        // Hei/黑 → sans-serif CJK fonts (macOS: PingFang SC / STHeiti)
        if (lower.Contains("黑") || lower.Contains("hei") || lower == "simhei")
            return ", 'PingFang SC', 'STHeiti'";
        // Kai/楷 → cursive CJK fonts (macOS: Kaiti SC / STKaiti)
        if (lower.Contains("楷") || lower.Contains("kai"))
            return ", 'Kaiti SC', 'STKaiti'";
        // FangSong/仿宋 → (macOS: STFangsong)
        if (lower.Contains("仿宋") || lower.Contains("fangsong"))
            return ", 'STFangsong'";
        return "";
    }
}
