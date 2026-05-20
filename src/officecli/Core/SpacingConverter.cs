// Copyright 2025 OfficeCLI (officecli.ai)
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
internal static class SpacingConverter
{
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

    /// <summary>
    /// Signed twips variant for OOXML attributes typed `ST_SignedTwipsMeasure`:
    /// w:ind/@w:left, @w:right, @w:start, @w:end, @w:firstLine. Word documents
    /// commonly carry negative indents — e.g. `<w:ind w:right="-46">` so a
    /// table-of-contents page-number column overhangs the right margin. Real
    /// docs (gov.cn corpus) trip ParseWordSpacing's non-negative gate even
    /// though the OOXML schema explicitly allows negatives for these slots.
    /// Use this for indent slots only; spaceBefore/spaceAfter remain on the
    /// strict `ST_TwipsMeasure` path (the non-negative gate there catches
    /// the silent line-collapse failure mode that motivated it).
    /// </summary>
    public static int ParseWordSpacingSigned(string value)
    {
        var points = ParseSpacingToPointsSigned(value, bareIsPoints: false);
        return (int)Math.Round(points * TwipsPerPoint);
    }

    private static double ParseSpacingToPointsSigned(string value, bool bareIsPoints)
    {
        var trimmed = value.Trim();
        bool neg = trimmed.StartsWith("-");
        var rest = neg ? trimmed[1..] : trimmed;
        var pos = ParseSpacingToPoints(rest, bareIsPoints);
        return neg ? -pos : pos;
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
        // BUG-R7-03: PPT stores spaceBefore/spaceAfter in hundredths of a point
        // as a 32-bit signed integer (CT_TextSpacing). Compute in 64-bit and
        // reject values that would silently overflow on cast — the symptom was
        // 999999999pt clamping to int.MaxValue/100 ≈ 21474836.47pt readback.
        var hundredths = (long)Math.Round(points * 100);
        if (hundredths > int.MaxValue)
            throw new ArgumentException(
                $"Invalid spacing value '{value}'. Value too large — exceeds maximum representable spacing (~{int.MaxValue / 100.0}pt).");
        return (int)hundredths;
    }

    /// <summary>
    /// Parse a length value to points. Accepts unit-qualified "12pt", "0.5cm",
    /// "0.5in" or bare number (treated as points). Used for XLSX shape margin
    /// to mirror Get's "Npt" output. CONSISTENCY(spacing-units).
    /// </summary>
    public static double ParsePoints(string value)
    {
        var points = ParseSpacingToPoints(value, bareIsPoints: true);
        if (points < 0)
            throw new ArgumentException($"Invalid length value '{value}'. Must be non-negative.");
        return points;
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

        // BUG-R7-04: lineSpacing must be strictly > 0. Zero produces degenerate
        // OOXML (w:spacing/@line=0 is undefined in MS-DOC) and Office silently
        // collapses to single-spacing — surface the error to the user instead.
        static double RequirePositive(double n, string raw)
        {
            if (n <= 0)
                throw new ArgumentException($"Invalid 'lineSpacing' value '{raw}'. Line spacing must be greater than 0.");
            return n;
        }

        // "1.5x" → multiplier
        if (trimmed.EndsWith("x", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^1], "lineSpacing"), value);
            return ((uint)Math.Round(num * WordAutoLineSpacingUnit), true);
        }

        // "150%" → multiplier
        if (trimmed.EndsWith("%", StringComparison.Ordinal))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^1], "lineSpacing"), value);
            return ((uint)Math.Round(num / 100.0 * WordAutoLineSpacingUnit), true);
        }

        // "18pt" → fixed (Exact)
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^2], "lineSpacing"), value);
            return ((uint)Math.Round(num * TwipsPerPoint), false);
        }

        // "0.5cm" → fixed (Exact), convert to points first
        if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^2], "lineSpacing"), value);
            return ((uint)Math.Round(num * PointsPerCm * TwipsPerPoint), false);
        }

        // "0.5in" → fixed (Exact)
        if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^2], "lineSpacing"), value);
            return ((uint)Math.Round(num * PointsPerInch * TwipsPerPoint), false);
        }

        // Bare number → multiplier under Auto rule, mirrors the "1.5x" path.
        // Word stores Auto line spacing in 240ths of a multiplier (1.0 = 240,
        // 1.5 = 360, 2.0 = 480). Earlier this returned the raw value as twips
        // (`Math.Round(1.5) = 2 twips`), which Word silently treated as a
        // single-spaced line because 2 twips is below any visible threshold.
        var bare = RequirePositive(ParseNumber(trimmed, "lineSpacing"), value);
        return ((uint)Math.Round(bare * WordAutoLineSpacingUnit), true);
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

        // BUG-R7-04: lineSpacing must be strictly > 0. SpacingPercent(0) is
        // degenerate — Office silently renders single-line spacing without
        // any error, masking the user's mistake.
        static double RequirePositive(double n, string raw)
        {
            if (n <= 0)
                throw new ArgumentException($"Invalid 'lineSpacing' value '{raw}'. Line spacing must be greater than 0.");
            return n;
        }

        // "1.5x" → multiplier → SpacingPercent
        if (trimmed.EndsWith("x", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^1], "lineSpacing"), value);
            return ((int)Math.Round(num * 100000), true);
        }

        // "150%" → multiplier → SpacingPercent
        if (trimmed.EndsWith("%", StringComparison.Ordinal))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^1], "lineSpacing"), value);
            return ((int)Math.Round(num * 1000), true);
        }

        // "18pt" → fixed → SpacingPoints
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^2], "lineSpacing"), value);
            return ((int)Math.Round(num * 100), false);
        }

        // "0.5cm" → fixed → SpacingPoints
        if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^2], "lineSpacing"), value);
            return ((int)Math.Round(num * PointsPerCm * 100), false);
        }

        // "0.5in" → fixed → SpacingPoints
        if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = RequirePositive(ParseNumber(trimmed[..^2], "lineSpacing"), value);
            return ((int)Math.Round(num * PointsPerInch * 100), false);
        }

        // Bare number → backward compat: multiplier → SpacingPercent
        var bare = RequirePositive(ParseNumber(trimmed, "lineSpacing"), value);
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
        // SpacingPercent stores multiplier*100000 so the underlying precision
        // is 5 decimal places. "0.##" rounds to 2dp which collapsed any
        // multiplier below 0.005 to "0x", losing the fact that the value was
        // non-zero. Use "0.#####" so small non-zero multipliers round-trip
        // visibly (0.001x -> "0.001x" rather than "0x").
        return $"{multiplier:0.#####}x";
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
