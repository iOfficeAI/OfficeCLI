// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Text Rendering ====================

    private static void RenderTextBody(StringBuilder sb, OpenXmlElement textBody, Dictionary<string, string> themeColors,
        Shape? placeholderShape = null, OpenXmlPart? placeholderPart = null)
    {
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            // Resolve per-paragraph font size based on paragraph level
            int? defaultFontSizeHundredths = null;
            if (placeholderShape != null && placeholderPart != null)
            {
                int level = para.ParagraphProperties?.Level?.Value ?? 0;
                defaultFontSizeHundredths = ResolvePlaceholderFontSize(placeholderShape, placeholderPart, level);
            }
            var paraStyles = new List<string>();

            var pProps = para.ParagraphProperties;
            if (pProps?.Alignment?.HasValue == true)
            {
                var align = pProps.Alignment.InnerText switch
                {
                    "l" => "left",
                    "ctr" => "center",
                    "r" => "right",
                    "just" => "justify",
                    _ => "left"
                };
                paraStyles.Add($"text-align:{align}");
            }

            // Paragraph spacing
            var sbPts = pProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sbPts.HasValue) paraStyles.Add($"margin-top:{sbPts.Value / 100.0:0.##}pt");
            var saPts = pProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (saPts.HasValue) paraStyles.Add($"margin-bottom:{saPts.Value / 100.0:0.##}pt");

            // Line spacing
            var lsPct = pProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (lsPct.HasValue) paraStyles.Add($"line-height:{lsPct.Value / 100000.0:0.##}");
            var lsPts = pProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (lsPts.HasValue) paraStyles.Add($"line-height:{lsPts.Value / 100.0:0.##}pt");

            // Indent
            if (pProps?.Indent?.HasValue == true)
                paraStyles.Add($"text-indent:{Units.EmuToPt(pProps.Indent.Value)}pt");
            if (pProps?.LeftMargin?.HasValue == true)
                paraStyles.Add($"margin-left:{Units.EmuToPt(pProps.LeftMargin.Value)}pt");

            // Bullet
            var bulletChar = pProps?.GetFirstChild<Drawing.CharacterBullet>()?.Char?.Value;
            var bulletAuto = pProps?.GetFirstChild<Drawing.AutoNumberedBullet>();
            var hasBullet = bulletChar != null || bulletAuto != null;

            sb.Append($"<div class=\"para\" style=\"{string.Join(";", paraStyles)}\">");

            if (hasBullet)
            {
                var bullet = bulletChar ?? "\u2022";
                var buStyles = new List<string>();

                // Bullet color: explicit buClr > first run color > default (inherit)
                var buClrFill = pProps?.GetFirstChild<Drawing.BulletColor>()
                    ?.GetFirstChild<Drawing.SolidFill>();
                var bulletColor = ResolveFillColor(buClrFill, themeColors);
                if (bulletColor == null)
                {
                    // Follow first run text color (same as LibreOffice/POI behavior)
                    var firstRun = para.Elements<Drawing.Run>().FirstOrDefault();
                    var firstRunFill = firstRun?.RunProperties?.GetFirstChild<Drawing.SolidFill>();
                    bulletColor = ResolveFillColor(firstRunFill, themeColors);
                }
                if (bulletColor != null) buStyles.Add($"color:{bulletColor}");

                // Bullet size: explicit buSzPts/buSzPct > first run size > default size
                var buSzPts = pProps?.GetFirstChild<Drawing.BulletSizePoints>();
                var buSzPct = pProps?.GetFirstChild<Drawing.BulletSizePercentage>();
                if (buSzPts?.Val?.HasValue == true)
                {
                    buStyles.Add($"font-size:{buSzPts.Val.Value / 100.0:0.##}pt");
                }
                else
                {
                    // Determine base font size from first run or default
                    var firstRun = para.Elements<Drawing.Run>().FirstOrDefault();
                    var baseSizeHundredths = firstRun?.RunProperties?.FontSize?.Value ?? defaultFontSizeHundredths;
                    if (baseSizeHundredths.HasValue)
                    {
                        var pct = buSzPct?.Val?.HasValue == true ? buSzPct.Val.Value / 100000.0 : 1.0;
                        buStyles.Add($"font-size:{baseSizeHundredths.Value / 100.0 * pct:0.##}pt");
                    }
                }

                var buStyle = buStyles.Count > 0 ? $" style=\"{string.Join(";", buStyles)}\"" : "";
                sb.Append($"<span class=\"bullet\"{buStyle}>{HtmlEncode(bullet)} </span>");
            }

            // Check for OfficeMath (a14:m inside mc:AlternateContent) in paragraph XML
            var paraXml = para.OuterXml;
            if (paraXml.Contains("oMath"))
            {
                // AlternateContent is opaque to Descendants() — parse from XML
                var mathMatch = System.Text.RegularExpressions.Regex.Match(paraXml,
                    @"<m:oMathPara[^>]*>.*?</m:oMathPara>|<m:oMath[^>]*>.*?</m:oMath>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (mathMatch.Success)
                {
                    var mathXml = $"<wrapper xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">{mathMatch.Value}</wrapper>";
                    try
                    {
                        var wrapper = new OpenXmlUnknownElement("wrapper");
                        wrapper.InnerXml = mathMatch.Value;
                        var oMath = wrapper.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                        if (oMath != null)
                        {
                            var latex = FormulaParser.ToLatex(oMath);
                            sb.Append($"<span class=\"katex-formula\" data-formula=\"{HtmlEncode(latex)}\"></span>");
                        }
                    }
                    catch { }
                }
            }

            var hasMath = paraXml.Contains("oMath");
            var runs = para.Elements<Drawing.Run>().ToList();
            if (runs.Count == 0 && !hasMath)
            {
                // Empty paragraph (line break)
                sb.Append("&nbsp;");
            }
            else
            {
                foreach (var run in runs)
                {
                    RenderRun(sb, run, themeColors, defaultFontSizeHundredths);
                }
            }

            // Line breaks within paragraph
            foreach (var br in para.Elements<Drawing.Break>())
                sb.Append("<br>");

            sb.AppendLine("</div>");
        }
    }

    private static void RenderRun(StringBuilder sb, Drawing.Run run, Dictionary<string, string> themeColors,
        int? defaultFontSizeHundredths = null)
    {
        var text = run.Text?.Text ?? "";
        if (string.IsNullOrEmpty(text)) return;

        var styles = new List<string>();
        var rp = run.RunProperties;

        if (rp != null)
        {
            // Font
            var font = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
            if (font != null && !font.StartsWith("+", StringComparison.Ordinal))
                styles.Add(CssFontFamilyWithFallback(font));

            // Size — use explicit run size, fall back to placeholder default
            if (rp.FontSize?.HasValue == true)
                styles.Add($"font-size:{rp.FontSize.Value / 100.0:0.##}pt");
            else if (defaultFontSizeHundredths.HasValue)
                styles.Add($"font-size:{defaultFontSizeHundredths.Value / 100.0:0.##}pt");

            // Bold
            if (rp.Bold?.Value == true)
                styles.Add("font-weight:bold");

            // Italic
            if (rp.Italic?.Value == true)
                styles.Add("font-style:italic");

            // Underline
            if (rp.Underline?.HasValue == true && rp.Underline.Value != Drawing.TextUnderlineValues.None)
                styles.Add("text-decoration:underline");

            // Strikethrough
            if (rp.Strike?.HasValue == true && rp.Strike.Value != Drawing.TextStrikeValues.NoStrike)
                styles.Add("text-decoration:line-through");

            // Color
            var solidFill = rp.GetFirstChild<Drawing.SolidFill>();
            var color = ResolveFillColor(solidFill, themeColors);
            if (color != null)
                styles.Add($"color:{color}");

            // Gradient text fill
            var gradFill = rp.GetFirstChild<Drawing.GradientFill>();
            if (gradFill != null)
            {
                var gradCss = GradientToCss(gradFill, themeColors);
                if (!string.IsNullOrEmpty(gradCss))
                {
                    styles.Add($"background:{gradCss}");
                    styles.Add("-webkit-background-clip:text");
                    styles.Add("background-clip:text");
                    styles.Add("-webkit-text-fill-color:transparent");
                }
            }

            // Character spacing
            if (rp.Spacing?.HasValue == true)
                styles.Add($"letter-spacing:{rp.Spacing.Value / 100.0:0.##}pt");

            // Superscript/subscript
            if (rp.Baseline?.HasValue == true && rp.Baseline.Value != 0)
            {
                if (rp.Baseline.Value > 0)
                    styles.Add("vertical-align:super;font-size:smaller");
                else
                    styles.Add("vertical-align:sub;font-size:smaller");
            }
        }

        // Hyperlink
        var hlinkClick = run.Parent?.Elements<Drawing.Run>()
            .Where(r => r == run)
            .Select(_ => run.Parent)
            .FirstOrDefault()
            ?.GetFirstChild<Drawing.HyperlinkOnClick>();
        // Actually check run's parent paragraph for hyperlinks on this run
        // Not critical for preview, skip for simplicity

        if (styles.Count > 0)
            sb.Append($"<span style=\"{string.Join(";", styles)}\">{HtmlEncode(text)}</span>");
        else
            sb.Append(HtmlEncode(text));
    }
}
