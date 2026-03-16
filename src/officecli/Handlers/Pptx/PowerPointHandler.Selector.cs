// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private record ShapeSelector(string? ElementType, int? SlideNum, string? TextContains,
        string? FontEquals, string? FontNotEquals, bool? IsTitle, bool? HasAlt);

    private static ShapeSelector ParseShapeSelector(string selector)
    {
        string? elementType = null;
        int? slideNum = null;
        string? textContains = null;
        string? fontEquals = null;
        string? fontNotEquals = null;
        bool? isTitle = null;
        bool? hasAlt = null;

        // Check for slide prefix
        var slideMatch = Regex.Match(selector, @"slide\[(\d+)\]\s*(.*)");
        if (slideMatch.Success)
        {
            slideNum = int.Parse(slideMatch.Groups[1].Value);
            selector = slideMatch.Groups[2].Value.TrimStart('>', ' ');
        }

        // Element type
        var typeMatch = Regex.Match(selector, @"^(\w+)");
        if (typeMatch.Success)
        {
            var t = typeMatch.Groups[1].Value.ToLowerInvariant();
            if (t is "shape" or "textbox" or "title" or "picture" or "pic" or "equation" or "math" or "formula"
                or "table" or "chart" or "placeholder")
                elementType = t;
        }

        // Attributes
        foreach (Match attrMatch in Regex.Matches(selector, @"\[(\w+)(!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value;
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "font" when op == "=": fontEquals = val; break;
                case "font" when op == "!=": fontNotEquals = val; break;
                case "title": isTitle = val.ToLowerInvariant() != "false"; break;
                case "alt": hasAlt = !string.IsNullOrEmpty(val) && val.ToLowerInvariant() != "false"; break;
            }
        }

        // :contains()
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) textContains = containsMatch.Groups[1].Value;

        // Element type shortcuts
        if (elementType == "title") isTitle = true;

        // :no-alt
        if (selector.Contains(":no-alt")) hasAlt = false;

        return new ShapeSelector(elementType, slideNum, textContains, fontEquals, fontNotEquals, isTitle, hasAlt);
    }

    private static bool MatchesShapeSelector(Shape shape, ShapeSelector selector)
    {
        // Element type filter
        if (selector.ElementType is "picture" or "pic" or "table" or "chart" or "placeholder")
            return false;

        // Title filter
        if (selector.IsTitle.HasValue)
        {
            if (selector.IsTitle.Value != IsTitle(shape)) return false;
        }

        // Text contains
        if (selector.TextContains != null)
        {
            var text = GetShapeText(shape);
            if (!text.Contains(selector.TextContains, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        // Font filter
        var runs = shape.Descendants<Drawing.Run>().ToList();
        if (selector.FontEquals != null)
        {
            bool found = runs.Any(r =>
            {
                var font = r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                return font != null && string.Equals(font, selector.FontEquals, StringComparison.OrdinalIgnoreCase);
            });
            if (!found) return false;
        }

        if (selector.FontNotEquals != null)
        {
            bool hasWrongFont = runs.Any(r =>
            {
                var font = r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                return font != null && !string.Equals(font, selector.FontNotEquals, StringComparison.OrdinalIgnoreCase);
            });
            if (!hasWrongFont) return false;
        }

        return true;
    }

    private static bool MatchesPictureSelector(Picture pic, ShapeSelector selector)
    {
        // Only match if looking for pictures specifically or no type specified
        if (selector.ElementType != null && selector.ElementType != "picture" && selector.ElementType != "pic")
            return false;

        if (selector.IsTitle.HasValue) return false; // Pictures can't be titles

        // Alt text filter
        if (selector.HasAlt.HasValue)
        {
            var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
            bool hasAlt = !string.IsNullOrEmpty(alt);
            if (selector.HasAlt.Value != hasAlt) return false;
        }

        return true;
    }
}
