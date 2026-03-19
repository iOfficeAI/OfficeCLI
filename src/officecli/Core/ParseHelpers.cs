// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;

namespace OfficeCli.Core;

/// <summary>
/// Shared parsing helpers for handler property values.
/// Accepts flexible user input (e.g. "true", "yes", "1", "on" for booleans;
/// "24pt" or "24" for font sizes).
/// </summary>
public static class ParseHelpers
{
    /// <summary>
    /// Accepts "true", "1", "yes", "on" (case-insensitive) as truthy.
    /// </summary>
    public static bool IsTruthy(string value) =>
        value.ToLowerInvariant() is "true" or "1" or "yes" or "on";

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
        if (!double.TryParse(trimmed, CultureInfo.InvariantCulture, out var result))
            throw new ArgumentException($"Invalid font size: '{value}'. Expected a number (e.g., '12', '10.5', '14pt').");
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
        if (!double.TryParse(value, CultureInfo.InvariantCulture, out var result))
            throw new ArgumentException($"Invalid '{propertyName}' value '{value}'. Expected a number.");
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
        var hex = value.TrimStart('#').ToUpperInvariant();
        if (hex.Length == 3)
        {
            // Expand shorthand: "F00" → "FF0000"
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });
        }
        if (hex.Length == 6)
            return "FF" + hex;
        return hex; // 8-char ARGB or other (pass through)
    }

    /// <summary>
    /// Sanitize a hex color for OOXML srgbClr val (must be exactly 6-char RGB).
    /// If 8-char hex is given, interprets as AARRGGBB (POI convention: alpha first),
    /// strips the leading alpha and returns it separately.
    /// Returns (rgb6, alphaPercent) where alphaPercent is 0-100000 scale or null if fully opaque.
    /// </summary>
    public static (string Rgb, int? AlphaPercent) SanitizeColorForOoxml(string value)
    {
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
                $"8-digit AARRGGBB (e.g. 80FF0000), or scheme color name.");

        return (hex, null);
    }
}
