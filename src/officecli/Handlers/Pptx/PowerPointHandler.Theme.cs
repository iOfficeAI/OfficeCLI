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
    // ==================== Morph Check ====================

    /// <summary>
    /// Analyse morph compatibility across all slides.
    /// Returns a node with children — one per morph-eligible shape pair or unmatched shape.
    /// A shape participates in Morph if its name starts with "!!" (per OOXML morph matching rules).
    /// Each child node Format:
    ///   status = "matched" | "unmatched"
    ///   name   = shape name (e.g. "!!circle")
    ///   from   = source path (e.g. "/slide[1]/shape[2]")
    ///   to     = target path if matched (e.g. "/slide[2]/shape[3]")
    ///   type   = shape type
    /// </summary>
    private DocumentNode GetMorphCheckNode()
    {
        var root = new DocumentNode { Path = "/morph-check", Type = "morph-check" };
        var children = new List<DocumentNode>();

        var slideParts = GetSlideParts().ToList();

        // Build a per-slide index: shapeName → (shapeIdx, type)
        static List<(string Name, int Idx, string Type)> GetSlideShapes(Slide slide)
        {
            var result = new List<(string, int, string)>();
            var tree = slide.CommonSlideData?.ShapeTree;
            if (tree == null) return result;
            int i = 0;
            foreach (var shape in tree.Elements<Shape>())
            {
                i++;
                var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "";
                result.Add((name, i, IsTitle(shape) ? "title" : "textbox"));
            }
            int pi = 0;
            foreach (var pic in tree.Elements<Picture>())
            {
                pi++;
                var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "";
                result.Add((name, pi, "picture"));
            }
            return result;
        }

        for (int sIdx = 0; sIdx < slideParts.Count; sIdx++)
        {
            var slide = GetSlide(slideParts[sIdx]);
            // Morph transition is stored as mc:AlternateContent wrapping p159:morph (raw XML)
            var slideEl = GetSlide(slideParts[sIdx]);
            bool hasMorphTransition = slideEl.ChildElements.Any(c =>
                c.LocalName == "AlternateContent" &&
                c.Descendants().Any(d => d.LocalName == "morph"));

            var shapes = GetSlideShapes(slide);
            // Shapes eligible for morph: all shapes if the slide has morph transition,
            // plus any shape named !!* anywhere (for the next slide to match)
            var morphCandidates = shapes.Where(s => s.Name.StartsWith("!!", StringComparison.Ordinal)).ToList();

            if (morphCandidates.Count == 0 && !hasMorphTransition) continue;

            // Build lookup for next slide
            List<(string Name, int Idx, string Type)>? nextShapes = null;
            if (sIdx + 1 < slideParts.Count)
                nextShapes = GetSlideShapes(GetSlide(slideParts[sIdx + 1]));

            var nextLookup = nextShapes?
                .Where(s => s.Name.StartsWith("!!", StringComparison.Ordinal))
                .GroupBy(s => s.Name)
                .ToDictionary(g => g.Key, g => g.First());

            // Report all !! shapes on this slide
            foreach (var (name, idx, type) in morphCandidates)
            {
                var child = new DocumentNode
                {
                    Path = $"/slide[{sIdx + 1}]/shape[{idx}]",
                    Type = type
                };
                child.Format["name"] = name;
                child.Format["slide"] = sIdx + 1;

                if (nextLookup != null && nextLookup.TryGetValue(name, out var match))
                {
                    child.Format["status"] = "matched";
                    child.Format["to"] = $"/slide[{sIdx + 2}]/shape[{match.Idx}]";
                }
                else
                {
                    child.Format["status"] = "unmatched";
                }

                children.Add(child);
            }

            // Report morph transition info per slide
            if (hasMorphTransition)
            {
                var slideNode = new DocumentNode
                {
                    Path = $"/slide[{sIdx + 1}]",
                    Type = "slide"
                };
                slideNode.Format["transition"] = "morph";
                // Read morph mode from raw XML (p159:morph option attribute)
                var morphEl = slideEl.Descendants().FirstOrDefault(d => d.LocalName == "morph");
                var mode = morphEl?.GetAttribute("option", "").Value ?? "byObject";
                slideNode.Format["morphMode"] = string.IsNullOrEmpty(mode) ? "byObject" : mode;
                slideNode.Format["morphShapes"] = morphCandidates.Count;
                slideNode.Format["matchedShapes"] = morphCandidates.Count(s =>
                    nextLookup != null && nextLookup.ContainsKey(s.Name));
                children.Add(slideNode);
            }
        }

        root.Children = children;
        root.ChildCount = children.Count;
        root.Preview = children.Count == 0
            ? "No morph-eligible shapes found (name shapes with !! prefix)"
            : $"{children.Count(c => c.Format.TryGetValue("status", out var s) && s?.ToString() == "matched")} matched, "
              + $"{children.Count(c => c.Format.TryGetValue("status", out var s) && s?.ToString() == "unmatched")} unmatched";
        return root;
    }

    // ==================== Theme Color ====================

    /// <summary>
    /// Get the presentation theme's color scheme.
    /// Returns a DocumentNode at path "/theme" with Format keys:
    ///   accent1-6, dk1, dk2, lt1, lt2, hyperlink, followedhyperlink, headingFont, bodyFont
    /// </summary>
    private DocumentNode GetThemeNode()
    {
        var node = new DocumentNode { Path = "/theme", Type = "theme" };
        var scheme = GetColorScheme();
        if (scheme == null) return node;

        static string? ReadSchemeColor(OpenXmlCompositeElement? el)
        {
            if (el == null) return null;
            var rgb = el.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (rgb != null) return rgb;
            var sys = el.GetFirstChild<Drawing.SystemColor>();
            return sys?.LastColor?.Value ?? sys?.Val?.InnerText;
        }

        if (ReadSchemeColor(scheme.Dark1Color) is { } dk1) node.Format["dk1"] = dk1;
        if (ReadSchemeColor(scheme.Light1Color) is { } lt1) node.Format["lt1"] = lt1;
        if (ReadSchemeColor(scheme.Dark2Color) is { } dk2) node.Format["dk2"] = dk2;
        if (ReadSchemeColor(scheme.Light2Color) is { } lt2) node.Format["lt2"] = lt2;
        if (ReadSchemeColor(scheme.Accent1Color) is { } a1) node.Format["accent1"] = a1;
        if (ReadSchemeColor(scheme.Accent2Color) is { } a2) node.Format["accent2"] = a2;
        if (ReadSchemeColor(scheme.Accent3Color) is { } a3) node.Format["accent3"] = a3;
        if (ReadSchemeColor(scheme.Accent4Color) is { } a4) node.Format["accent4"] = a4;
        if (ReadSchemeColor(scheme.Accent5Color) is { } a5) node.Format["accent5"] = a5;
        if (ReadSchemeColor(scheme.Accent6Color) is { } a6) node.Format["accent6"] = a6;
        if (ReadSchemeColor(scheme.Hyperlink) is { } hl) node.Format["hyperlink"] = hl;
        if (ReadSchemeColor(scheme.FollowedHyperlinkColor) is { } fhl) node.Format["followedhyperlink"] = fhl;

        // Font scheme
        var themePart = GetThemePart();
        if (themePart != null)
        {
            var fontScheme = themePart.Theme?.ThemeElements?.FontScheme;
            var majorLatin = fontScheme?.MajorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface;
            var minorLatin = fontScheme?.MinorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface;
            if (!string.IsNullOrEmpty(majorLatin)) node.Format["headingFont"] = majorLatin;
            if (!string.IsNullOrEmpty(minorLatin)) node.Format["bodyFont"] = minorLatin;
            if (scheme.Name?.Value != null) node.Format["name"] = scheme.Name.Value;
        }

        return node;
    }

    /// <summary>
    /// Set theme color scheme properties.
    /// Supported keys: accent1-6, dk1, dk2, lt1, lt2, hyperlink, followedhyperlink,
    ///                 headingFont, bodyFont, name
    /// Values: hex RGB (e.g. "FF6B35") or "default" to reset to Office default.
    /// </summary>
    private List<string> SetThemeProperties(Dictionary<string, string> properties)
    {
        var scheme = GetColorScheme()
            ?? throw new InvalidOperationException("No theme color scheme found in presentation");
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "accent1": SetSchemeColor(scheme.Accent1Color ??= new Drawing.Accent1Color(), value); break;
                case "accent2": SetSchemeColor(scheme.Accent2Color ??= new Drawing.Accent2Color(), value); break;
                case "accent3": SetSchemeColor(scheme.Accent3Color ??= new Drawing.Accent3Color(), value); break;
                case "accent4": SetSchemeColor(scheme.Accent4Color ??= new Drawing.Accent4Color(), value); break;
                case "accent5": SetSchemeColor(scheme.Accent5Color ??= new Drawing.Accent5Color(), value); break;
                case "accent6": SetSchemeColor(scheme.Accent6Color ??= new Drawing.Accent6Color(), value); break;
                case "dk1" or "dark1": SetSchemeColor(scheme.Dark1Color ??= new Drawing.Dark1Color(), value); break;
                case "dk2" or "dark2": SetSchemeColor(scheme.Dark2Color ??= new Drawing.Dark2Color(), value); break;
                case "lt1" or "light1": SetSchemeColor(scheme.Light1Color ??= new Drawing.Light1Color(), value); break;
                case "lt2" or "light2": SetSchemeColor(scheme.Light2Color ??= new Drawing.Light2Color(), value); break;
                case "hyperlink" or "hlink": SetSchemeColor(scheme.Hyperlink ??= new Drawing.Hyperlink(), value); break;
                case "followedhyperlink" or "folhlink":
                    SetSchemeColor(scheme.FollowedHyperlinkColor ??= new Drawing.FollowedHyperlinkColor(), value);
                    break;
                case "name":
                    scheme.Name = value;
                    break;
                case "headingfont" or "headingFont":
                    SetFontScheme(majorTypeface: value);
                    break;
                case "bodyfont" or "bodyFont":
                    SetFontScheme(minorTypeface: value);
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        GetThemePart()?.Theme?.Save();
        return unsupported;
    }

    private static void SetSchemeColor(OpenXmlCompositeElement colorEl, string value)
    {
        // Remove existing color children
        colorEl.RemoveAllChildren<Drawing.RgbColorModelHex>();
        colorEl.RemoveAllChildren<Drawing.SystemColor>();
        colorEl.RemoveAllChildren<Drawing.SchemeColor>();
        colorEl.RemoveAllChildren<Drawing.HslColor>();
        colorEl.RemoveAllChildren<Drawing.PresetColor>();

        var rgb = value.TrimStart('#').ToUpperInvariant();
        if (rgb.Length == 6 && rgb.All(char.IsAsciiHexDigit))
            colorEl.AppendChild(new Drawing.RgbColorModelHex { Val = rgb });
        else
            throw new ArgumentException($"Theme color must be a 6-character hex value (e.g. FF6B35), got: {value}");
    }

    private void SetFontScheme(string? majorTypeface = null, string? minorTypeface = null)
    {
        var themePart = GetThemePart();
        if (themePart?.Theme?.ThemeElements?.FontScheme == null) return;
        var fontScheme = themePart.Theme.ThemeElements.FontScheme;

        if (majorTypeface != null)
        {
            var majorFont = fontScheme.MajorFont ??= new Drawing.MajorFont();
            var latin = majorFont.GetFirstChild<Drawing.LatinFont>();
            if (latin != null) latin.Typeface = majorTypeface;
            else majorFont.PrependChild(new Drawing.LatinFont { Typeface = majorTypeface });
        }
        if (minorTypeface != null)
        {
            var minorFont = fontScheme.MinorFont ??= new Drawing.MinorFont();
            var latin = minorFont.GetFirstChild<Drawing.LatinFont>();
            if (latin != null) latin.Typeface = minorTypeface;
            else minorFont.PrependChild(new Drawing.LatinFont { Typeface = minorTypeface });
        }
    }

    private Drawing.ColorScheme? GetColorScheme()
    {
        return GetThemePart()?.Theme?.ThemeElements?.ColorScheme;
    }

    private DocumentFormat.OpenXml.Packaging.ThemePart? GetThemePart()
    {
        var presentationPart = _doc.PresentationPart;
        if (presentationPart == null) return null;

        // Prefer theme directly on presentationPart
        if (presentationPart.ThemePart != null)
            return presentationPart.ThemePart;

        // Fall back to first slide master's theme
        return presentationPart.SlideMasterParts.FirstOrDefault()?.ThemePart;
    }
}
