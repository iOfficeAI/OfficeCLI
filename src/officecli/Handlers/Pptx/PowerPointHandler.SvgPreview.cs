// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// Generate a self-contained SVG for a single slide.
    /// Uses foreignObject to embed the HTML rendering for full fidelity.
    /// </summary>
    public string ViewAsSvg(int slideNum)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideNum < 1 || slideNum > slideParts.Count)
            throw new CliException($"Slide {slideNum} does not exist. This presentation has {slideParts.Count} slide(s).")
            {
                Code = "out_of_range",
                Suggestion = $"Use a slide number between 1 and {slideParts.Count}."
            };

        var slidePart = slideParts[slideNum - 1];
        var (slideWidthEmu, slideHeightEmu) = GetSlideSize();
        double slideWidthCm = slideWidthEmu / 360000.0;
        double slideHeightCm = slideHeightEmu / 360000.0;
        var themeColors = ResolveThemeColorMap();

        var sb = new StringBuilder();

        // SVG header with viewBox in cm
        sb.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"");
        sb.AppendLine($"     width=\"{slideWidthCm:0.###}cm\" height=\"{slideHeightCm:0.###}cm\"");
        sb.AppendLine($"     viewBox=\"0 0 {slideWidthCm:0.###} {slideHeightCm:0.###}\">");

        // Embed CSS styles for the HTML content inside foreignObject
        sb.AppendLine("<defs>");
        sb.AppendLine("<style type=\"text/css\">");
        sb.AppendLine(GenerateSvgCss(slideWidthCm, slideHeightCm));
        sb.AppendLine("</style>");
        sb.AppendLine("</defs>");

        // White background rect
        sb.AppendLine($"<rect width=\"{slideWidthCm:0.###}\" height=\"{slideHeightCm:0.###}\" fill=\"white\"/>");

        // foreignObject with the slide HTML content
        sb.AppendLine($"<foreignObject x=\"0\" y=\"0\" width=\"{slideWidthCm:0.###}\" height=\"{slideHeightCm:0.###}\">");
        sb.AppendLine($"<div xmlns=\"http://www.w3.org/1999/xhtml\" class=\"slide\"");

        // Slide background + text defaults
        var slideStyles = new List<string>();
        var bgStyle = GetSlideBackgroundCss(slidePart, themeColors);
        if (!string.IsNullOrEmpty(bgStyle))
            slideStyles.Add(bgStyle);
        var textDefaults = GetTextDefaults(slidePart, themeColors);
        if (!string.IsNullOrEmpty(textDefaults))
            slideStyles.Add(textDefaults);
        if (slideStyles.Count > 0)
            sb.Append($" style=\"{string.Join("", slideStyles)}\"");
        sb.AppendLine(">");

        // Render layout placeholders + slide elements (reuse existing HTML rendering)
        RenderLayoutPlaceholders(sb, slidePart, themeColors);
        RenderSlideElements(sb, slidePart, slideNum, slideWidthEmu, slideHeightEmu, themeColors);

        sb.AppendLine("</div>");
        sb.AppendLine("</foreignObject>");
        sb.AppendLine("</svg>");

        return sb.ToString();
    }

    /// <summary>
    /// Generate minimal CSS for SVG foreignObject (no page layout, just element styles).
    /// </summary>
    private static string GenerateSvgCss(double slideWidthCm, double slideHeightCm)
    {
        var css = new StringBuilder();

        // Slide container within foreignObject
        css.AppendLine($".slide {{ width: {slideWidthCm:0.###}cm; height: {slideHeightCm:0.###}cm; position: relative; overflow: hidden; background: transparent; }}");

        // Element styles (same as preview.css essentials)
        css.AppendLine(".shape { position: absolute; overflow: visible; white-space: pre-wrap; word-wrap: break-word; }");
        css.AppendLine(".shape.has-fill { overflow: hidden; }");
        css.AppendLine(".shape-text { width: 100%; height: 100%; display: flex; flex-direction: column; }");
        css.AppendLine(".shape-text.valign-top { justify-content: flex-start; }");
        css.AppendLine(".shape-text.valign-center { justify-content: center; }");
        css.AppendLine(".shape-text.valign-bottom { justify-content: flex-end; }");
        css.AppendLine(".para { width: 100%; line-height: 1; }");
        css.AppendLine(".picture { position: absolute; overflow: hidden; }");
        css.AppendLine(".picture img { width: 100%; height: 100%; object-fit: fill; }");
        css.AppendLine(".table-container { position: absolute; overflow: hidden; }");
        css.AppendLine(".slide-table { width: 100%; height: 100%; border-collapse: collapse; table-layout: fixed; }");
        css.AppendLine(".slide-table td { padding: 4px 6px; vertical-align: top; overflow: hidden; font-size: 10pt; color: inherit; }");
        css.AppendLine(".connector { position: absolute; pointer-events: none; }");
        css.AppendLine(".group { position: absolute; }");

        return css.ToString();
    }
}
