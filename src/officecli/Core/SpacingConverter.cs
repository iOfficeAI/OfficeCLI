// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;

namespace OfficeCli.Core;

/// <summary>
/// Unified spacing parser/formatter for Word and PowerPoint handlers.
/// Principle: input tolerant (accepts unit-qualified strings), output unified (always with units).
///
/// Supported input formats for spaceBefore / spaceAfter:
///   "12pt"   → 12 points
///   "0.5cm"  → centimeters (1cm = 28.3465pt)
///   "0.5in"  → inches (1in = 72pt)
///   bare number → backward compatible (Word: twips, PPT: points)
///
/// Supported input formats for lineSpacing:
///   "1.5x"   → 1.5× multiplier
///   "150%"   → 150% = 1.5× multiplier
///   "18pt"   → fixed 18pt line spacing
///   "0.5cm"  → fixed, converted to points
///   bare number → backward compatible (Word: twips+Auto, PPT: multiplier)
///
/// Output format:
///   spaceBefore / spaceAfter → "12pt"
///   lineSpacing multiplier   → "1.5x"
///   lineSpacing fixed        → "18pt"
/// </summary>
public static class SpacingConverter
{
    private const double CmPerPoint = 28.3464567; // 1 cm = 28.3465 pt (72 pt/in ÷ 2.54 cm/in)
    private const double PointsPerCm = 72.0 / 2.54; // ~28.3465
    private const double PointsPerInch = 72.0;
    private const int TwipsPerPoint = 20; // 1 pt = 20 twips
    private const int WordAutoLineSpacingUnit = 240; // 240 twips = single line in Auto mode

    // ────────────────────────────────────────────────────────────────
    //  spaceBefore / spaceAfter  →  Word twips
    // ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Parse a spacing value (spaceBefore/spaceAfter) to Word twips (uint).
    /// Accepts: "12pt", "0.5cm", "0.5in", or bare number (treated as twips for backward compat).
    /// </summary>
    public static uint ParseWordSpacing(string value)
    {
        var points = ParseSpacingToPoints(value, bareIsPoints: false);
        if (points < 0)
            throw new ArgumentException($"Invalid spacing value '{value}'. Spacing must be non-negative.");
        return (uint)Math.Round(points * TwipsPerPoint);
    }

    // ────────────────────────────────────────────────────────────────
    //  spaceBefore / spaceAfter  →  PPT hundredths-of-a-point
    // ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Parse a spacing value (spaceBefore/spaceAfter) to PPT hundredths-of-a-point (int).
    /// Accepts: "12pt", "0.5cm", "0.5in", or bare number (treated as points for backward compat).
    /// </summary>
    public static int ParsePptSpacing(string value)
    {
        var points = ParseSpacingToPoints(value, bareIsPoints: true);
        if (points < 0)
            throw new ArgumentException($"Invalid spacing value '{value}'. Spacing must be non-negative.");
        return (int)Math.Round(points * 100);
    }

    // ────────────────────────────────────────────────────────────────
    //  lineSpacing  →  Word (twips + LineRule)
    // ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Parse line spacing for Word. Returns (twips, isMultiplier).
    /// "1.5x" or "150%" → (360, true)  — Auto rule, 240 × multiplier
    /// "18pt"           → (360, true=false) — Exact rule, pt × 20
    /// "0.5cm"          → converted to pt, then Exact
    /// bare number      → (number, true) — Auto rule, backward compat (raw twips)
    /// </summary>
    public static (uint Twips, bool IsMultiplier) ParseWordLineSpacing(string value)
    {
        var trimmed = value.Trim();

        // "1.5x" → multiplier
        if (trimmed.EndsWith("x", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^1], "lineSpacing");
            return ((uint)Math.Round(num * WordAutoLineSpacingUnit), true);
        }

        // "150%" → multiplier
        if (trimmed.EndsWith("%", StringComparison.Ordinal))
        {
            var num = ParseNumber(trimmed[..^1], "lineSpacing");
            return ((uint)Math.Round(num / 100.0 * WordAutoLineSpacingUnit), true);
        }

        // "18pt" → fixed (Exact)
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^2], "lineSpacing");
            return ((uint)Math.Round(num * TwipsPerPoint), false);
        }

        // "0.5cm" → fixed (Exact), convert to points first
        if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^2], "lineSpacing");
            return ((uint)Math.Round(num * PointsPerCm * TwipsPerPoint), false);
        }

        // "0.5in" → fixed (Exact)
        if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^2], "lineSpacing");
            return ((uint)Math.Round(num * PointsPerInch * TwipsPerPoint), false);
        }

        // Bare number → backward compat: raw twips with Auto rule
        var bare = ParseNumber(trimmed, "lineSpacing");
        if (bare < 0)
            throw new ArgumentException($"Invalid 'lineSpacing' value '{value}'. Line spacing must be non-negative.");
        return ((uint)Math.Round(bare), true);
    }

    // ────────────────────────────────────────────────────────────────
    //  lineSpacing  →  PPT (SpacingPercent or SpacingPoints)
    // ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Parse line spacing for PPT. Returns (internalVal, isPercent).
    /// "1.5x" or "150%" → (150000, true)  — SpacingPercent
    /// "18pt"           → (1800, false)    — SpacingPoints (hundredths)
    /// "0.5cm"          → converted to pt, then SpacingPoints
    /// bare number      → (number × 100000, true) — SpacingPercent, backward compat (multiplier)
    /// </summary>
    public static (int Val, bool IsPercent) ParsePptLineSpacing(string value)
    {
        var trimmed = value.Trim();

        // "1.5x" → multiplier → SpacingPercent
        if (trimmed.EndsWith("x", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^1], "lineSpacing");
            return ((int)Math.Round(num * 100000), true);
        }

        // "150%" → multiplier → SpacingPercent
        if (trimmed.EndsWith("%", StringComparison.Ordinal))
        {
            var num = ParseNumber(trimmed[..^1], "lineSpacing");
            return ((int)Math.Round(num * 1000), true);
        }

        // "18pt" → fixed → SpacingPoints
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^2], "lineSpacing");
            return ((int)Math.Round(num * 100), false);
        }

        // "0.5cm" → fixed → SpacingPoints
        if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^2], "lineSpacing");
            return ((int)Math.Round(num * PointsPerCm * 100), false);
        }

        // "0.5in" → fixed → SpacingPoints
        if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseNumber(trimmed[..^2], "lineSpacing");
            return ((int)Math.Round(num * PointsPerInch * 100), false);
        }

        // Bare number → backward compat: multiplier → SpacingPercent
        var bare = ParseNumber(trimmed, "lineSpacing");
        if (bare < 0)
            throw new ArgumentException($"Invalid 'lineSpacing' value '{value}'. Line spacing must be non-negative.");
        return ((int)Math.Round(bare * 100000), true);
    }

    // ────────────────────────────────────────────────────────────────
    //  Output formatting
    // ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Format Word spaceBefore/spaceAfter twips to "Xpt".
    /// </summary>
    public static string FormatWordSpacing(string twipsStr)
    {
        if (!double.TryParse(twipsStr, CultureInfo.InvariantCulture, out var twips))
            return twipsStr;
        var points = twips / TwipsPerPoint;
        return $"{points:0.##}pt";
    }

    /// <summary>
    /// Format PPT spaceBefore/spaceAfter hundredths-of-a-point to "Xpt".
    /// </summary>
    public static string FormatPptSpacing(int hundredths)
    {
        var points = hundredths / 100.0;
        return $"{points:0.##}pt";
    }

    /// <summary>
    /// Format Word lineSpacing from twips + LineRule to "1.5x" or "18pt".
    /// lineRule: "auto" → multiplier (twips / 240), otherwise → fixed (twips / 20 + "pt").
    /// </summary>
    public static string FormatWordLineSpacing(string lineVal, string? lineRule)
    {
        if (!double.TryParse(lineVal, CultureInfo.InvariantCulture, out var twips))
            return lineVal;

        // Auto → multiplier
        if (lineRule == null || lineRule.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            var multiplier = twips / WordAutoLineSpacingUnit;
            return $"{multiplier:0.##}x";
        }

        // Exact or AtLeast → fixed points
        var points = twips / TwipsPerPoint;
        return $"{points:0.##}pt";
    }

    /// <summary>
    /// Format PPT lineSpacing from SpacingPercent val to "1.5x".
    /// </summary>
    public static string FormatPptLineSpacingPercent(int val)
    {
        var multiplier = val / 100000.0;
        return $"{multiplier:0.##}x";
    }

    /// <summary>
    /// Format PPT lineSpacing from SpacingPoints val to "18pt".
    /// </summary>
    public static string FormatPptLineSpacingPoints(int val)
    {
        var points = val / 100.0;
        return $"{points:0.##}pt";
    }

    // ────────────────────────────────────────────────────────────────
    //  Internal helpers
    // ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Parse spacing value to points. If bareIsPoints=true, bare numbers are points;
    /// if false, bare numbers are twips (Word backward compat).
    /// </summary>
    private static double ParseSpacingToPoints(string value, bool bareIsPoints)
    {
        var trimmed = value.Trim();

        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            return ParseNumber(trimmed[..^2], "spacing");

        if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            return ParseNumber(trimmed[..^2], "spacing") * PointsPerCm;

        if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            return ParseNumber(trimmed[..^2], "spacing") * PointsPerInch;

        // Bare number
        var num = ParseNumber(trimmed, "spacing");
        return bareIsPoints ? num : num / TwipsPerPoint; // twips → points if Word
    }

    private static double ParseNumber(string s, string context)
    {
        var trimmed = s.Trim();
        if (!double.TryParse(trimmed, CultureInfo.InvariantCulture, out var result)
            || double.IsNaN(result) || double.IsInfinity(result))
            throw new ArgumentException(
                $"Invalid '{context}' value '{s}'. Expected a finite number with optional unit (e.g. '12pt', '1.5x', '150%').");
        if (result < 0)
            throw new ArgumentException(
                $"Invalid '{context}' value '{s}'. Spacing values must be non-negative.");
        return result;
    }
}
