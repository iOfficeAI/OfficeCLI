// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Shared Theme Get/Set logic for all document types.
/// Operates on ThemePart which has identical structure across Word/Excel/PowerPoint.
/// </summary>
public static class ThemeHandler
{
    // ColorScheme slot names → accessor pairs
    private static readonly (string Key, Func<A.ColorScheme, A.Color2Type?> Get, Action<A.ColorScheme, string> Set)[] ColorSlots =
    [
        ("dk1", cs => cs.Dark1Color, (cs, v) => SetColorSlot(cs.Dark1Color, v)),
        ("lt1", cs => cs.Light1Color, (cs, v) => SetColorSlot(cs.Light1Color, v)),
        ("dk2", cs => cs.Dark2Color, (cs, v) => SetColorSlot(cs.Dark2Color, v)),
        ("lt2", cs => cs.Light2Color, (cs, v) => SetColorSlot(cs.Light2Color, v)),
        ("accent1", cs => cs.Accent1Color, (cs, v) => SetColorSlot(cs.Accent1Color, v)),
        ("accent2", cs => cs.Accent2Color, (cs, v) => SetColorSlot(cs.Accent2Color, v)),
        ("accent3", cs => cs.Accent3Color, (cs, v) => SetColorSlot(cs.Accent3Color, v)),
        ("accent4", cs => cs.Accent4Color, (cs, v) => SetColorSlot(cs.Accent4Color, v)),
        ("accent5", cs => cs.Accent5Color, (cs, v) => SetColorSlot(cs.Accent5Color, v)),
        ("accent6", cs => cs.Accent6Color, (cs, v) => SetColorSlot(cs.Accent6Color, v)),
        ("hlink", cs => cs.Hyperlink, (cs, v) => SetColorSlot(cs.Hyperlink, v)),
        ("folHlink", cs => cs.FollowedHyperlinkColor, (cs, v) => SetColorSlot(cs.FollowedHyperlinkColor, v)),
    ];

    /// <summary>
    /// Populate Format dictionary with theme properties.
    /// </summary>
    public static void PopulateTheme(ThemePart? themePart, DocumentNode node)
    {
        var theme = themePart?.Theme;
        if (theme == null) return;

        if (theme.Name?.Value != null)
            node.Format["theme.name"] = theme.Name.Value;

        var elements = theme.ThemeElements;
        if (elements == null) return;

        // ColorScheme
        var colorScheme = elements.ColorScheme;
        if (colorScheme != null)
        {
            if (colorScheme.Name?.Value != null)
                node.Format["theme.colorScheme"] = colorScheme.Name.Value;

            foreach (var (key, getter, _) in ColorSlots)
            {
                var slot = getter(colorScheme);
                var hex = ReadColorSlot(slot);
                if (hex != null)
                    node.Format[$"theme.color.{key}"] = ParseHelpers.FormatHexColor(hex);
            }
        }

        // FontScheme
        var fontScheme = elements.FontScheme;
        if (fontScheme != null)
        {
            if (fontScheme.Name?.Value != null)
                node.Format["theme.fontScheme"] = fontScheme.Name.Value;

            if (fontScheme.MajorFont?.LatinFont?.Typeface != null)
                node.Format["theme.font.major.latin"] = fontScheme.MajorFont.LatinFont.Typeface!.Value!;
            if (fontScheme.MajorFont?.EastAsianFont?.Typeface != null)
                node.Format["theme.font.major.eastAsia"] = fontScheme.MajorFont.EastAsianFont.Typeface!.Value!;
            if (fontScheme.MinorFont?.LatinFont?.Typeface != null)
                node.Format["theme.font.minor.latin"] = fontScheme.MinorFont.LatinFont.Typeface!.Value!;
            if (fontScheme.MinorFont?.EastAsianFont?.Typeface != null)
                node.Format["theme.font.minor.eastAsia"] = fontScheme.MinorFont.EastAsianFont.Typeface!.Value!;
        }

        // FormatScheme (Get only — name only, no deep read of fill/line/effect lists)
        var formatScheme = elements.FormatScheme;
        if (formatScheme?.Name?.Value != null)
            node.Format["theme.formatScheme"] = formatScheme.Name.Value;
    }

    /// <summary>
    /// Try to Set a theme.* property. Returns true if handled.
    /// </summary>
    public static bool TrySetTheme(ThemePart? themePart, string key, string value)
    {
        var theme = themePart?.Theme;
        if (theme == null) return false;

        // theme.color.<slot>
        if (key.StartsWith("theme.color."))
        {
            var slotName = key["theme.color.".Length..];
            var colorScheme = theme.ThemeElements?.ColorScheme;
            if (colorScheme == null) return false;

            foreach (var (k, _, setter) in ColorSlots)
            {
                if (string.Equals(k, slotName, StringComparison.OrdinalIgnoreCase))
                {
                    setter(colorScheme, value);
                    theme.Save();
                    return true;
                }
            }
            return false;
        }

        // theme.font.major.latin / theme.font.minor.latin etc.
        if (key.StartsWith("theme.font."))
        {
            var fontScheme = theme.ThemeElements?.FontScheme;
            if (fontScheme == null) return false;

            switch (key)
            {
                case "theme.font.major.latin":
                    if (fontScheme.MajorFont?.LatinFont != null) fontScheme.MajorFont.LatinFont.Typeface = value;
                    break;
                case "theme.font.major.eastasia":
                    if (fontScheme.MajorFont?.EastAsianFont != null) fontScheme.MajorFont.EastAsianFont.Typeface = value;
                    break;
                case "theme.font.minor.latin":
                    if (fontScheme.MinorFont?.LatinFont != null) fontScheme.MinorFont.LatinFont.Typeface = value;
                    break;
                case "theme.font.minor.eastasia":
                    if (fontScheme.MinorFont?.EastAsianFont != null) fontScheme.MinorFont.EastAsianFont.Typeface = value;
                    break;
                default:
                    return false;
            }
            theme.Save();
            return true;
        }

        return false;
    }

    // ==================== Color Slot Helpers ====================

    private static string? ReadColorSlot(A.Color2Type? slot)
    {
        if (slot == null) return null;
        var rgb = slot.GetFirstChild<A.RgbColorModelHex>();
        if (rgb?.Val?.Value != null) return rgb.Val.Value;
        var sys = slot.GetFirstChild<A.SystemColor>();
        if (sys?.LastColor?.Value != null) return sys.LastColor.Value;
        return null;
    }

    private static void SetColorSlot(A.Color2Type? slot, string value)
    {
        if (slot == null) return;
        var result = ParseHelpers.SanitizeColorForOoxml(value);

        // Remove existing children
        slot.RemoveAllChildren();
        slot.AppendChild(new A.RgbColorModelHex { Val = result.Rgb });
    }

    /// <summary>
    /// Get the ThemePart for each document type.
    /// </summary>
    public static ThemePart? GetThemePart(object doc)
    {
        return doc switch
        {
            DocumentFormat.OpenXml.Packaging.WordprocessingDocument w => w.MainDocumentPart?.ThemePart,
            DocumentFormat.OpenXml.Packaging.SpreadsheetDocument s => s.WorkbookPart?.ThemePart,
            DocumentFormat.OpenXml.Packaging.PresentationDocument p => p.PresentationPart?.SlideMasterParts?.FirstOrDefault()?.ThemePart,
            _ => null
        };
    }
}
