// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Shared parsing helpers for handler property values.
/// Accepts flexible user input (e.g. "true", "yes", "1", "on" for booleans;
/// "24pt" or "24" for font sizes).
/// </summary>
internal static class ParseHelpers
{
    /// <summary>
    /// Map of common CSS/HTML named colors to 6-digit uppercase hex RGB.
    /// </summary>
    private static readonly Dictionary<string, string> NamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["red"] = "FF0000", ["green"] = "008000", ["blue"] = "0000FF",
        ["white"] = "FFFFFF", ["black"] = "000000", ["yellow"] = "FFFF00",
        ["cyan"] = "00FFFF", ["aqua"] = "00FFFF", ["magenta"] = "FF00FF",
        ["fuchsia"] = "FF00FF", ["orange"] = "FFA500", ["purple"] = "800080",
        ["pink"] = "FFC0CB", ["brown"] = "A52A2A", ["gray"] = "808080",
        ["grey"] = "808080", ["silver"] = "C0C0C0", ["gold"] = "FFD700",
        ["navy"] = "000080", ["teal"] = "008080", ["maroon"] = "800000",
        ["olive"] = "808000", ["lime"] = "00FF00", ["coral"] = "FF7F50",
        ["salmon"] = "FA8072", ["tomato"] = "FF6347", ["crimson"] = "DC143C",
        ["indigo"] = "4B0082", ["violet"] = "EE82EE", ["turquoise"] = "40E0D0",
        ["tan"] = "D2B48C", ["khaki"] = "F0E68C", ["beige"] = "F5F5DC",
        ["ivory"] = "FFFFF0", ["lavender"] = "E6E6FA", ["plum"] = "DDA0DD",
        ["orchid"] = "DA70D6", ["chocolate"] = "D2691E", ["sienna"] = "A0522D",
        ["peru"] = "CD853F", ["wheat"] = "F5DEB3", ["linen"] = "FAF0E6",
        ["skyblue"] = "87CEEB", ["steelblue"] = "4682B4", ["slategray"] = "708090",
        ["darkred"] = "8B0000", ["darkgreen"] = "006400", ["darkblue"] = "00008B",
        ["darkcyan"] = "008B8B", ["darkmagenta"] = "8B008B", ["darkorange"] = "FF8C00",
        ["darkviolet"] = "9400D3", ["deeppink"] = "FF1493", ["deepskyblue"] = "00BFFF",
        ["lightgray"] = "D3D3D3", ["lightgreen"] = "90EE90", ["lightblue"] = "ADD8E6",
        ["lightyellow"] = "FFFFE0", ["lightpink"] = "FFB6C1", ["lightcoral"] = "F08080",
        ["darkgray"] = "A9A9A9", ["dimgray"] = "696969",
    };

    /// <summary>
    /// Try to resolve a named color (e.g. "red") or rgb() notation to 6-digit hex.
    /// Returns null if the input is not a named color or rgb() expression.
    /// </summary>
    private static string? TryResolveColorInput(string value)
    {
        var trimmed = value.Trim();

        // Named color lookup
        if (NamedColors.TryGetValue(trimmed, out var hex))
            return hex;

        // rgb(r,g,b) notation
        var m = Regex.Match(trimmed, @"^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$", RegexOptions.IgnoreCase);
        if (m.Success)
        {
            var r = int.Parse(m.Groups[1].Value);
            var g = int.Parse(m.Groups[2].Value);
            var b = int.Parse(m.Groups[3].Value);
            if (r > 255 || g > 255 || b > 255)
                throw new ArgumentException($"Invalid color value: '{value}'. RGB components must be 0-255.");
            return $"{r:X2}{g:X2}{b:X2}";
        }

        return null;
    }

    /// <summary>
    /// Format a raw hex color value for user-facing output.
    /// Adds '#' prefix to 6-digit hex colors. Passes through scheme color names and special values unchanged.
    /// </summary>
    public static string FormatHexColor(string rawValue)
    {
        if (string.IsNullOrEmpty(rawValue)) return rawValue;
        if (rawValue.StartsWith('#')) return rawValue.ToUpperInvariant();
        if (rawValue.Length == 6 && rawValue.All(char.IsAsciiHexDigit))
            return "#" + rawValue.ToUpperInvariant();
        // 8-char ARGB (e.g. "FFFF0000") → strip alpha prefix → "#FF0000"
        if (rawValue.Length == 8 && rawValue.All(char.IsAsciiHexDigit))
            return "#" + rawValue[2..].ToUpperInvariant();
        // Try resolving named colors (e.g. "silver" → "#C0C0C0")
        var resolved = TryResolveColorInput(rawValue);
        if (resolved != null)
            return "#" + resolved.ToUpperInvariant();
        return rawValue; // scheme colors ("accent1"), "none", "auto", etc.
    }

    /// <summary>
    /// Map Excel theme color index to a canonical scheme name.
    /// OOXML theme indices: 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1-6, 10=hlink, 11=folHlink.
    /// </summary>
    public static string? ExcelThemeIndexToName(uint themeIndex) => themeIndex switch
    {
        0 => "lt1",
        1 => "dk1",
        2 => "lt2",
        3 => "dk2",
        4 => "accent1",
        5 => "accent2",
        6 => "accent3",
        7 => "accent4",
        8 => "accent5",
        9 => "accent6",
        10 => "hlink",
        11 => "folHlink",
        _ => null,
    };

    /// <summary>
    /// Returns true if the value is a recognized boolean string and is truthy.
    /// Returns false for null, empty, or recognized falsy values ("false", "0", "no", "off").
    /// Throws <see cref="ArgumentException"/> for non-null values that are not recognized boolean strings.
    /// </summary>
    public static bool IsTruthy(string? value)
    {
        if (value == null) return false;
        return TrimInvisible(value).ToLowerInvariant() switch
        {
            "true" or "1" or "yes" or "on" => true,
            "false" or "0" or "no" or "off" or "" => false,
            _ => throw new ArgumentException(
                $"Invalid boolean value: '{value}'. Expected true/false, yes/no, 1/0, or on/off.")
        };
    }

    // R10: BOM (U+FEFF) and other zero-width / format chars are NOT in
    // char.IsWhiteSpace, so a plain Trim() leaves them in place. R8 added
    // Trim() but tests with `"﻿true"` still threw. Use a stricter
    // predicate that also drops format/control chars.
    private static string TrimInvisible(string s)
    {
        return s.Trim().Trim(s_invisibleChars);
    }

    private static readonly char[] s_invisibleChars =
    {
        '﻿', // BOM / zero-width no-break space
        '​', // zero-width space
        '‌', // zero-width non-joiner
        '‍', // zero-width joiner
        '⁠', // word joiner
        ' ', // non-breaking space (technically whitespace category in some configs but be explicit)
    };

    /// <summary>
    /// Returns true if the value is a recognized truthy string.
    /// Returns false for anything else (null, empty, falsy, or unrecognized values).
    /// Unlike <see cref="IsTruthy"/>, never throws.
    /// </summary>
    public static bool IsTruthySafe(string? value)
    {
        if (value == null) return false;
        return TrimInvisible(value).ToLowerInvariant() is "true" or "1" or "yes" or "on";
    }

    /// <summary>
    /// Returns true if the value is a recognized boolean string (truthy or falsy).
    /// Returns false for null, empty, or non-boolean values (no exception thrown).
    /// </summary>
    public static bool IsValidBooleanString(string? value) =>
        value != null && TrimInvisible(value).ToLowerInvariant() is "true" or "1" or "yes" or "on"
                                                                 or "false" or "0" or "no" or "off";

    /// <summary>
    /// Parse a font size string, stripping optional "pt" suffix.
    /// Supports integers and fractional values (e.g. "24", "10.5", "24pt").
    /// Returns double to preserve fractional sizes for correct unit conversion.
    /// </summary>
    public static double ParseFontSize(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed[..^2].Trim();
        if (trimmed.Contains(','))
            throw new ArgumentException($"Invalid font size: '{value}'. Comma is not allowed — use '.' as decimal separator (e.g., '10.5').");
        if (!double.TryParse(trimmed, CultureInfo.InvariantCulture, out var result) || double.IsNaN(result) || double.IsInfinity(result))
            throw new ArgumentException($"Invalid font size: '{value}'. Expected a finite number (e.g., '12', '10.5', '14pt').");
        if (result <= 0)
            throw new ArgumentException($"Invalid font size: '{value}'. Font size must be greater than 0.");
        return result;
    }

    /// <summary>
    /// Safely parse a string as int, throwing ArgumentException with a clear message on failure.
    /// </summary>
    public static int SafeParseInt(string value, string propertyName)
    {
        if (!int.TryParse(value, CultureInfo.InvariantCulture, out var result))
            throw new ArgumentException($"Invalid '{propertyName}' value '{value}'. Expected an integer.");
        return result;
    }

    /// <summary>
    /// Safely parse a string as double, throwing ArgumentException with a clear message on failure.
    /// </summary>
    public static double SafeParseDouble(string value, string propertyName)
    {
        if (!double.TryParse(value, CultureInfo.InvariantCulture, out var result) || double.IsNaN(result) || double.IsInfinity(result))
            throw new ArgumentException($"Invalid '{propertyName}' value '{value}'. Expected a finite number.");
        return result;
    }

    /// <summary>
    /// Safely parse a string as uint, throwing ArgumentException with a clear message on failure.
    /// </summary>
    public static uint SafeParseUint(string value, string propertyName)
    {
        if (!uint.TryParse(value, CultureInfo.InvariantCulture, out var result))
            throw new ArgumentException($"Invalid '{propertyName}' value '{value}'. Expected a non-negative integer.");
        return result;
    }

    /// <summary>
    /// Safely parse a string as byte, throwing ArgumentException with a clear message on failure.
    /// </summary>
    public static byte SafeParseByte(string value, string propertyName)
    {
        if (!byte.TryParse(value, CultureInfo.InvariantCulture, out var result))
            throw new ArgumentException($"Invalid '{propertyName}' value '{value}'. Expected an integer (0-255).");
        return result;
    }

    /// <summary>
    /// Normalize a hex color string to 8-char ARGB format (e.g. "FFFF0000").
    /// Accepts: "FF0000" (6-char RGB → prepend FF), "#FF0000" (strip #), "F00" (3-char → expand),
    /// "80FF0000" (8-char ARGB → as-is). Always returns uppercase.
    /// </summary>
    public static string NormalizeArgbColor(string value)
    {
        // Try named color / rgb() first
        var resolved = TryResolveColorInput(value);
        if (resolved != null) return "FF" + resolved;

        var hex = value.TrimStart('#').ToUpperInvariant();
        if (hex.Length == 3 && hex.All(char.IsAsciiHexDigit))
        {
            // Expand shorthand: "F00" → "FF0000"
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });
        }
        if (hex.Length == 6 && hex.All(char.IsAsciiHexDigit))
            return "FF" + hex;
        if (hex.Length == 8 && hex.All(char.IsAsciiHexDigit))
            return hex;
        throw new ArgumentException(
            $"Invalid color value: '{value}'. Expected 6-digit hex RGB (e.g. FF0000), " +
            $"8-digit AARRGGBB (e.g. 80FF0000), 3-digit shorthand (e.g. F00), " +
            $"named color (e.g. red), or rgb() notation (e.g. rgb(255,0,0)).");
    }

    /// <summary>
    /// Sanitize a hex color for OOXML srgbClr val (must be exactly 6-char RGB).
    /// If 8-char hex is given, interprets as AARRGGBB (POI convention: alpha first),
    /// strips the leading alpha and returns it separately.
    /// Returns (rgb6, alphaPercent) where alphaPercent is 0-100000 scale or null if fully opaque.
    /// </summary>
    public static (string Rgb, int? AlphaPercent) SanitizeColorForOoxml(string value)
    {
        // "auto" is a legal OOXML value for shading Fill/Color — pass through unchanged
        if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
            return ("auto", null);

        // Try named color / rgb() first
        var resolved = TryResolveColorInput(value);
        if (resolved != null) return (resolved, null);

        var hex = value.TrimStart('#').ToUpperInvariant();
        if (hex.Length == 8 && hex.All(char.IsAsciiHexDigit))
        {
            var alphaByte = Convert.ToByte(hex[..2], 16); // AA portion: 00=transparent, FF=opaque
            var rgb = hex[2..];                            // RRGGBB portion
            if (alphaByte == 0xFF)
                return (rgb, null);
            var alphaPercent = (int)(alphaByte / 255.0 * 100000);
            return (rgb, alphaPercent);
        }
        // Validate: must be exactly 6 hex digits for srgbClr val
        if (hex.Length == 3 && hex.All(char.IsAsciiHexDigit))
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });

        if (hex.Length != 6 || !hex.All(char.IsAsciiHexDigit))
            throw new ArgumentException(
                $"Invalid color value: '{value}'. Expected 6-digit hex RGB (e.g. FF0000), " +
                $"8-digit AARRGGBB (e.g. 80FF0000), named color (e.g. red), " +
                $"rgb() notation (e.g. rgb(255,0,0)), or scheme color name.");

        return (hex, null);
    }

    // ==================== CJK Text Width Estimation ====================

    /// <summary>
    /// Returns true if the character is CJK ideograph, fullwidth, or CJK punctuation.
    /// These characters occupy approximately 1em width (≈ fontSize) vs ~0.55em for Latin.
    /// </summary>
    public static bool IsCjkOrFullWidth(char ch)
    {
        // CJK Unified Ideographs
        if (ch >= 0x4E00 && ch <= 0x9FFF) return true;
        // CJK Extension A
        if (ch >= 0x3400 && ch <= 0x4DBF) return true;
        // CJK Compatibility Ideographs
        if (ch >= 0xF900 && ch <= 0xFAFF) return true;
        // CJK Symbols and Punctuation (。、「」etc.)
        if (ch >= 0x3000 && ch <= 0x303F) return true;
        // Fullwidth Forms (Ａ-Ｚ, ０-９, fullwidth punctuation)
        if (ch >= 0xFF01 && ch <= 0xFF60) return true;
        // Halfwidth Katakana is NOT fullwidth
        // Hiragana
        if (ch >= 0x3040 && ch <= 0x309F) return true;
        // Katakana
        if (ch >= 0x30A0 && ch <= 0x30FF) return true;
        // Hangul Syllables
        if (ch >= 0xAC00 && ch <= 0xD7AF) return true;
        // Bopomofo
        if (ch >= 0x3100 && ch <= 0x312F) return true;
        // Em-dash (U+2014) is fullwidth in CJK contexts
        if (ch == 0x2014) return true;
        return false;
    }

    /// <summary>
    /// Estimate the visual width of a string in "character units" (Latin char = 1.0, CJK/fullwidth = ~1.82).
    /// Useful for Excel column auto-fit where width is measured in character units.
    /// </summary>
    public static double EstimateTextWidthInChars(string text)
    {
        double width = 0;
        foreach (char ch in text)
            width += IsCjkOrFullWidth(ch) ? 1.82 : 1.0;
        return width;
    }
}
