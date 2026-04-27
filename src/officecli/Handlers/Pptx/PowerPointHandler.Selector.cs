// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private record ShapeSelector(string? ElementType, int? SlideNum, string? TextContains,
        string? FontEquals, string? FontNotEquals, bool? IsTitle, bool? HasAlt,
        Dictionary<string, (string Value, bool Negate)>? Attributes = null);

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
            // CONSISTENCY(query-slide-prefix): strip '>', '/', or ' ' separators
            // so both "slide[1]>ole" and "/slide[1]/ole" resolve the element type.
            selector = slideMatch.Groups[2].Value.TrimStart('>', '/', ' ');
        }
        else
        {
            // CONSISTENCY(query-slide-prefix): also accept unindexed `slide > shape`
            // as "match this child type across all slides" — Word supports child
            // combinators without a specific parent index, so PPTX should too.
            var unindexedSlideMatch = Regex.Match(selector, @"^\s*slide\s*>\s*(.+)$", RegexOptions.IgnoreCase);
            if (unindexedSlideMatch.Success)
                selector = unindexedSlideMatch.Groups[1].Value;
        }

        // Element type
        var typeMatch = Regex.Match(selector, @"^(\w+)");
        if (typeMatch.Success)
        {
            var t = typeMatch.Groups[1].Value.ToLowerInvariant();
            if (t is "shape" or "textbox" or "title" or "picture" or "pic"
                or "video" or "audio"
                or "equation" or "math" or "formula"
                or "table" or "chart" or "placeholder"
                or "connector" or "connection"
                or "group" or "notes"
                or "zoom" or "slidezoom")
                elementType = t;
        }

        // Attributes
        Dictionary<string, (string Value, bool Negate)>? genericAttrs = null;
        foreach (Match attrMatch in Regex.Matches(selector, @"\[(\w+)(~=|\\?!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value.Replace("\\", "");
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "font" when op == "=": fontEquals = val; break;
                case "font" when op == "!=": fontNotEquals = val; break;
                case "title": isTitle = val.ToLowerInvariant() != "false"; break;
                case "alt": hasAlt = !string.IsNullOrEmpty(val) && val.ToLowerInvariant() != "false"; break;
                default:
                    genericAttrs ??= new Dictionary<string, (string, bool)>();
                    if (op == "~=")
                    {
                        // ~= is a "contains" match — store with special prefix
                        // Also handled by AttributeFilter post-filter (idempotent)
                        genericAttrs[key] = ($"\x01~={val}", false);
                    }
                    else
                    {
                        genericAttrs[key] = (val, op == "!=");
                    }
                    break;
            }
        }

        // :contains()
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) textContains = containsMatch.Groups[1].Value;

        // Shorthand: "shape:text" → treat as :contains(text)
        if (textContains == null)
        {
            var shorthandMatch = Regex.Match(selector, @"^(?:\w+)?:(?!contains|empty|no-alt|has)(.+)$");
            if (shorthandMatch.Success) textContains = shorthandMatch.Groups[1].Value;
        }

        // Element type shortcuts
        if (elementType == "title") isTitle = true;

        // :no-alt
        if (selector.Contains(":no-alt")) hasAlt = false;

        return new ShapeSelector(elementType, slideNum, textContains, fontEquals, fontNotEquals, isTitle, hasAlt, genericAttrs);
    }

    private static bool MatchesShapeSelector(Shape shape, ShapeSelector selector)
    {
        // Element type filter
        if (selector.ElementType is "picture" or "pic" or "video" or "audio" or "table" or "chart"
            or "placeholder" or "connector" or "connection" or "group" or "notes" or "zoom")
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

    private static bool MatchesGenericAttributes(DocumentNode node, Dictionary<string, (string Value, bool Negate)>? attributes)
    {
        if (attributes == null || attributes.Count == 0) return true;

        foreach (var (key, (expected, negate)) in attributes)
        {
            // Special case: "text" attribute matches node.Text, not Format["text"]
            var isTextKey = string.Equals(key, "text", StringComparison.OrdinalIgnoreCase);
            var matchedKey = node.Format.Keys.FirstOrDefault(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase));
            var hasKey = matchedKey != null || (isTextKey && node.Text != null);
            object? actual = matchedKey != null ? node.Format[matchedKey!] : (isTextKey ? node.Text : null);
            var actualStr = actual?.ToString() ?? "";

            // Handle ~= (contains) operator
            if (expected.StartsWith("\x01~="))
            {
                var pattern = expected[3..]; // strip "\x01~="
                if (!hasKey) return false;
                if (!actualStr.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return false;
                continue;
            }

            var isNameKey = string.Equals(key, "name", StringComparison.OrdinalIgnoreCase);

            if (negate)
            {
                // [attr!=value]: must not equal
                var matches = isNameKey ? MatchesShapeName(actualStr, expected) : NormalizedEquals(actualStr, expected);
                if (hasKey && matches)
                    return false;
            }
            else
            {
                // [attr=value]: must exist and equal
                if (!hasKey) return false;
                var matches = isNameKey ? MatchesShapeName(actualStr, expected) : NormalizedEquals(actualStr, expected);
                if (!matches)
                {
                    // Special case: boolean properties stored as `true`/`True` matching "true"
                    if (actual is bool b && string.Equals(expected, b.ToString(), StringComparison.OrdinalIgnoreCase))
                        continue;
                    // Special case: dimension values with different units (e.g., "0.07cm" vs "2pt")
                    if (Core.EmuConverter.TryParseEmu(actualStr, out var actualEmu)
                        && Core.EmuConverter.TryParseEmu(expected, out var expectedEmu)
                        && Math.Abs(actualEmu - expectedEmu) <= 500)
                        continue;
                    return false;
                }
            }
        }

        return true;
    }

    /// <summary>
    /// Case-insensitive comparison that also normalizes '#' prefix for color hex values.
    /// "#FF0000" equals "FF0000" and vice versa.
    /// </summary>
    private static bool NormalizedEquals(string a, string b)
    {
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase))
            return true;
        var aNorm = a.TrimStart('#');
        var bNorm = b.TrimStart('#');
        if (aNorm != a || bNorm != b)
            return string.Equals(aNorm, bNorm, StringComparison.OrdinalIgnoreCase);
        return false;
    }

    /// <summary>
    /// Match shape name with !! morph prefix awareness.
    /// "my-box" matches both "my-box" and "!!my-box".
    /// "!!my-box" matches both "!!my-box" and "my-box".
    /// </summary>
    private static bool MatchesShapeName(string? actual, string expected)
    {
        if (actual == null) return false;
        if (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            return true;
        // Strip !! prefix from actual name and compare
        if (actual.StartsWith("!!") && string.Equals(actual[2..], expected, StringComparison.OrdinalIgnoreCase))
            return true;
        // Strip !! prefix from expected and compare
        if (expected.StartsWith("!!") && string.Equals(actual, expected[2..], StringComparison.OrdinalIgnoreCase))
            return true;
        return false;
    }

    private static bool MatchesPictureSelector(Picture pic, ShapeSelector selector)
    {
        // Only match if looking for pictures/video/audio or no type specified
        if (selector.ElementType != null &&
            selector.ElementType is not ("picture" or "pic" or "video" or "audio"))
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
