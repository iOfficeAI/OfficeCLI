// Copyright 2025 OfficeCLI (officecli.ai)
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
    /// Full SVG / CSS3 named-color set (147 names + the CSS keyword
    /// <c>transparent</c>) mapped to 6-digit uppercase hex RGB. Lookup is
    /// case-insensitive. <c>transparent</c> maps to <c>000000</c> here; the
    /// transparent semantics (alpha = 0) are handled in the alpha-aware
    /// resolver below — callers using only the RGB component get black, which
    /// matches CSS's "transparent = transparent black" definition.
    /// </summary>
    private static readonly Dictionary<string, string> NamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["aliceblue"] = "F0F8FF", ["antiquewhite"] = "FAEBD7", ["aqua"] = "00FFFF",
        ["aquamarine"] = "7FFFD4", ["azure"] = "F0FFFF", ["beige"] = "F5F5DC",
        ["bisque"] = "FFE4C4", ["black"] = "000000", ["blanchedalmond"] = "FFEBCD",
        ["blue"] = "0000FF", ["blueviolet"] = "8A2BE2", ["brown"] = "A52A2A",
        ["burlywood"] = "DEB887", ["cadetblue"] = "5F9EA0", ["chartreuse"] = "7FFF00",
        ["chocolate"] = "D2691E", ["coral"] = "FF7F50", ["cornflowerblue"] = "6495ED",
        ["cornsilk"] = "FFF8DC", ["crimson"] = "DC143C", ["cyan"] = "00FFFF",
        ["darkblue"] = "00008B", ["darkcyan"] = "008B8B", ["darkgoldenrod"] = "B8860B",
        ["darkgray"] = "A9A9A9", ["darkgrey"] = "A9A9A9", ["darkgreen"] = "006400",
        ["darkkhaki"] = "BDB76B", ["darkmagenta"] = "8B008B", ["darkolivegreen"] = "556B2F",
        ["darkorange"] = "FF8C00", ["darkorchid"] = "9932CC", ["darkred"] = "8B0000",
        ["darksalmon"] = "E9967A", ["darkseagreen"] = "8FBC8F", ["darkslateblue"] = "483D8B",
        ["darkslategray"] = "2F4F4F", ["darkslategrey"] = "2F4F4F", ["darkturquoise"] = "00CED1",
        ["darkviolet"] = "9400D3", ["deeppink"] = "FF1493", ["deepskyblue"] = "00BFFF",
        ["dimgray"] = "696969", ["dimgrey"] = "696969", ["dodgerblue"] = "1E90FF",
        ["firebrick"] = "B22222", ["floralwhite"] = "FFFAF0", ["forestgreen"] = "228B22",
        ["fuchsia"] = "FF00FF", ["gainsboro"] = "DCDCDC", ["ghostwhite"] = "F8F8FF",
        ["gold"] = "FFD700", ["goldenrod"] = "DAA520", ["gray"] = "808080",
        ["grey"] = "808080", ["green"] = "008000", ["greenyellow"] = "ADFF2F",
        ["honeydew"] = "F0FFF0", ["hotpink"] = "FF69B4", ["indianred"] = "CD5C5C",
        ["indigo"] = "4B0082", ["ivory"] = "FFFFF0", ["khaki"] = "F0E68C",
        ["lavender"] = "E6E6FA", ["lavenderblush"] = "FFF0F5", ["lawngreen"] = "7CFC00",
        ["lemonchiffon"] = "FFFACD", ["lightblue"] = "ADD8E6", ["lightcoral"] = "F08080",
        ["lightcyan"] = "E0FFFF", ["lightgoldenrodyellow"] = "FAFAD2", ["lightgray"] = "D3D3D3",
        ["lightgrey"] = "D3D3D3", ["lightgreen"] = "90EE90", ["lightpink"] = "FFB6C1",
        ["lightsalmon"] = "FFA07A", ["lightseagreen"] = "20B2AA", ["lightskyblue"] = "87CEFA",
        ["lightslategray"] = "778899", ["lightslategrey"] = "778899", ["lightsteelblue"] = "B0C4DE",
        ["lightyellow"] = "FFFFE0", ["lime"] = "00FF00", ["limegreen"] = "32CD32",
        ["linen"] = "FAF0E6", ["magenta"] = "FF00FF", ["maroon"] = "800000",
        ["mediumaquamarine"] = "66CDAA", ["mediumblue"] = "0000CD", ["mediumorchid"] = "BA55D3",
        ["mediumpurple"] = "9370DB", ["mediumseagreen"] = "3CB371", ["mediumslateblue"] = "7B68EE",
        ["mediumspringgreen"] = "00FA9A", ["mediumturquoise"] = "48D1CC", ["mediumvioletred"] = "C71585",
        ["midnightblue"] = "191970", ["mintcream"] = "F5FFFA", ["mistyrose"] = "FFE4E1",
        ["moccasin"] = "FFE4B5", ["navajowhite"] = "FFDEAD", ["navy"] = "000080",
        ["oldlace"] = "FDF5E6", ["olive"] = "808000", ["olivedrab"] = "6B8E23",
        ["orange"] = "FFA500", ["orangered"] = "FF4500", ["orchid"] = "DA70D6",
        ["palegoldenrod"] = "EEE8AA", ["palegreen"] = "98FB98", ["paleturquoise"] = "AFEEEE",
        ["palevioletred"] = "DB7093", ["papayawhip"] = "FFEFD5", ["peachpuff"] = "FFDAB9",
        ["peru"] = "CD853F", ["pink"] = "FFC0CB", ["plum"] = "DDA0DD",
        ["powderblue"] = "B0E0E6", ["purple"] = "800080", ["rebeccapurple"] = "663399",
        ["red"] = "FF0000", ["rosybrown"] = "BC8F8F", ["royalblue"] = "4169E1",
        ["saddlebrown"] = "8B4513", ["salmon"] = "FA8072", ["sandybrown"] = "F4A460",
        ["seagreen"] = "2E8B57", ["seashell"] = "FFF5EE", ["sienna"] = "A0522D",
        ["silver"] = "C0C0C0", ["skyblue"] = "87CEEB", ["slateblue"] = "6A5ACD",
        ["slategray"] = "708090", ["slategrey"] = "708090", ["snow"] = "FFFAFA",
        ["springgreen"] = "00FF7F", ["steelblue"] = "4682B4", ["tan"] = "D2B48C",
        ["teal"] = "008080", ["thistle"] = "D8BFD8", ["tomato"] = "FF6347",
        ["turquoise"] = "40E0D0", ["violet"] = "EE82EE", ["wheat"] = "F5DEB3",
        ["white"] = "FFFFFF", ["whitesmoke"] = "F5F5F5", ["yellow"] = "FFFF00",
        ["yellowgreen"] = "9ACD32",
        // CSS keyword: transparent black (rgb=000000, alpha=00). Lookup here
        // returns the RGB component only; the alpha is injected by the
        // alpha-aware path in NormalizeArgbColor / SanitizeColorForOoxml.
        ["transparent"] = "000000",
    };

    /// <summary>
    /// Try to resolve a named color, <c>rgb()</c>, <c>rgba()</c>, <c>hsl()</c>,
    /// or <c>hsla()</c> notation to a 6-digit hex RGB plus an optional alpha
    /// byte. Returns <c>null</c> if the input is none of the above (callers
    /// should then fall through to the bare-hex parsers).
    ///
    /// Accepted CSS forms (case-insensitive, leading/trailing whitespace
    /// tolerated):
    ///   rgb(r,g,b)           rgba(r,g,b,a)         — 0–255 ints or 0–100% percentages, mixable
    ///   rgb(r g b)           rgba(r g b / a)       — CSS Level 4 space-separated
    ///   hsl(h,s%,l%)         hsla(h,s%,l%,a)       — h: deg (or unitless = deg), s/l: %
    ///   hsl(h s% l%)         hsla(h s% l% / a)     — CSS Level 4 space-separated
    /// Alpha forms: 0..1 float (CSS default) or 0..100% (CSS Level 4).
    /// </summary>
    private static (string Rgb, byte? Alpha)? TryResolveColorInput(string value)
    {
        var trimmed = value.Trim();

        // Named color lookup. `transparent` is the only entry that carries an
        // implicit alpha (= 0); everything else is opaque.
        if (NamedColors.TryGetValue(trimmed, out var hex))
        {
            if (string.Equals(trimmed, "transparent", StringComparison.OrdinalIgnoreCase))
                return (hex, (byte)0);
            return (hex, null);
        }

        // rgb(...) / rgba(...): three or four comma-separated OR space-
        // separated components (CSS Level 4). The trailing alpha may be
        // separated by `/` in the space-separated form, by `,` otherwise.
        var rgbMatch = Regex.Match(trimmed,
            @"^rgba?\(\s*([^\s,/]+)\s*[,\s]\s*([^\s,/]+)\s*[,\s]\s*([^\s,/]+)\s*(?:[,/]\s*([^\s,/]+)\s*)?\)$",
            RegexOptions.IgnoreCase);
        if (rgbMatch.Success)
        {
            byte r = ParseRgbComponent(rgbMatch.Groups[1].Value, value);
            byte g = ParseRgbComponent(rgbMatch.Groups[2].Value, value);
            byte b = ParseRgbComponent(rgbMatch.Groups[3].Value, value);
            byte? a = rgbMatch.Groups[4].Success
                ? ParseAlphaComponent(rgbMatch.Groups[4].Value, value)
                : (byte?)null;
            return ($"{r:X2}{g:X2}{b:X2}", a);
        }

        // hsl(...) / hsla(...). Hue: number (treated as degrees). Saturation
        // / lightness: percentages. Alpha: 0..1 or 0..100%.
        var hslMatch = Regex.Match(trimmed,
            @"^hsla?\(\s*([^\s,/]+)\s*[,\s]\s*([^\s,/]+)\s*[,\s]\s*([^\s,/]+)\s*(?:[,/]\s*([^\s,/]+)\s*)?\)$",
            RegexOptions.IgnoreCase);
        if (hslMatch.Success)
        {
            double h = ParseHueDegrees(hslMatch.Groups[1].Value, value);
            double s = ParsePercent01(hslMatch.Groups[2].Value, value);
            double l = ParsePercent01(hslMatch.Groups[3].Value, value);
            var (r, g, b) = HslToRgb(h, s, l);
            byte? a = hslMatch.Groups[4].Success
                ? ParseAlphaComponent(hslMatch.Groups[4].Value, value)
                : (byte?)null;
            return ($"{r:X2}{g:X2}{b:X2}", a);
        }

        return null;
    }

    private static byte ParseRgbComponent(string token, string original)
    {
        token = token.Trim();
        if (token.EndsWith('%'))
        {
            if (!double.TryParse(token[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct)
                || pct < 0 || pct > 100)
                throw new ArgumentException($"Invalid color value: '{original}'. RGB percentage components must be 0%-100%.");
            return (byte)Math.Round(pct / 100.0 * 255.0);
        }
        if (!int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) || n < 0 || n > 255)
            throw new ArgumentException($"Invalid color value: '{original}'. RGB components must be 0-255.");
        return (byte)n;
    }

    private static byte ParseAlphaComponent(string token, string original)
    {
        token = token.Trim();
        double a;
        if (token.EndsWith('%'))
        {
            if (!double.TryParse(token[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct)
                || pct < 0 || pct > 100)
                throw new ArgumentException($"Invalid color value: '{original}'. Alpha percentage must be 0%-100%.");
            a = pct / 100.0;
        }
        else
        {
            if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out a) || a < 0 || a > 1)
                throw new ArgumentException($"Invalid color value: '{original}'. Alpha must be 0-1 (e.g. 0.5) or 0%-100%.");
        }
        return (byte)Math.Round(a * 255.0);
    }

    private static double ParseHueDegrees(string token, string original)
    {
        token = token.Trim();
        // Strip an optional `deg` suffix (CSS allows it; other angle units
        // — turn/rad/grad — are out of scope for the input vocabulary we care
        // about and would just need their own multipliers if asked for).
        if (token.EndsWith("deg", StringComparison.OrdinalIgnoreCase))
            token = token[..^3].Trim();
        if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var h))
            throw new ArgumentException($"Invalid color value: '{original}'. Hue must be a number in degrees.");
        h %= 360;
        if (h < 0) h += 360;
        return h;
    }

    private static double ParsePercent01(string token, string original)
    {
        token = token.Trim();
        if (!token.EndsWith('%'))
            throw new ArgumentException($"Invalid color value: '{original}'. HSL saturation/lightness must be expressed as a percentage (e.g. 50%).");
        if (!double.TryParse(token[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct)
            || pct < 0 || pct > 100)
            throw new ArgumentException($"Invalid color value: '{original}'. HSL saturation/lightness must be 0%-100%.");
        return pct / 100.0;
    }

    private static (byte R, byte G, byte B) HslToRgb(double h, double s, double l)
    {
        // Standard HSL → RGB (CSS Color Module Level 3).
        double c = (1 - Math.Abs(2 * l - 1)) * s;
        double hp = h / 60.0;
        double x = c * (1 - Math.Abs(hp % 2 - 1));
        double r1 = 0, g1 = 0, b1 = 0;
        if (hp < 1)      { r1 = c; g1 = x; b1 = 0; }
        else if (hp < 2) { r1 = x; g1 = c; b1 = 0; }
        else if (hp < 3) { r1 = 0; g1 = c; b1 = x; }
        else if (hp < 4) { r1 = 0; g1 = x; b1 = c; }
        else if (hp < 5) { r1 = x; g1 = 0; b1 = c; }
        else             { r1 = c; g1 = 0; b1 = x; }
        double m = l - c / 2;
        return (
            (byte)Math.Round((r1 + m) * 255),
            (byte)Math.Round((g1 + m) * 255),
            (byte)Math.Round((b1 + m) * 255));
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
        // 8-char ARGB (e.g. "FFFF0000"). When alpha == FF (fully opaque), strip the
        // prefix and emit the canonical 6-digit form (#FF0000). When alpha < FF,
        // preserve the 8-digit form (#80FF0000) so partial transparency survives
        // round-tripping through Get. PPTX fill paths already preserve alpha via
        // a:alpha; this plug closes the Excel-side gap.
        if (rawValue.Length == 8 && rawValue.All(char.IsAsciiHexDigit))
        {
            var alpha = rawValue[..2];
            if (string.Equals(alpha, "FF", StringComparison.OrdinalIgnoreCase))
                return "#" + rawValue[2..].ToUpperInvariant();
            // CONSISTENCY(color-input-form): emit CSS #RRGGBBAA on output when
            // the value carries a hash prefix, mirroring the input form accepted
            // by NormalizeArgbColor / SanitizeColorForOoxml. The internal storage
            // stays AARRGGBB (OOXML/POI convention).
            return "#" + rawValue.Substring(2, 6).ToUpperInvariant() + rawValue[..2].ToUpperInvariant();
        }
        // Try resolving named colors (e.g. "silver" → "#C0C0C0")
        var resolved = TryResolveColorInput(rawValue);
        if (resolved is { } r)
            return "#" + r.Rgb.ToUpperInvariant();
        // Sentinel tokens are case-insensitive in OOXML but the canonical
        // Format emit is lowercase — sources occasionally write `w:val="AUTO"`
        // (or `"NONE"`), and a case-only delta poisons dump round-trip
        // diffing. Normalise the known sentinels here; scheme color names
        // (`accent1`, `dark1`, …) pass through unchanged.
        if (string.Equals(rawValue, "auto", StringComparison.OrdinalIgnoreCase)
            || string.Equals(rawValue, "none", StringComparison.OrdinalIgnoreCase))
            return rawValue.ToLowerInvariant();
        return rawValue; // scheme colors ("accent1"), etc.
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
        // OOXML w:sz/w:szCs/w:fontSize are half-points and must be >= 1.
        // Anything below 0.5pt would round to val=0 on write, producing
        // schema-invalid OOXML. Reject up front with the same shape as
        // the "<= 0" guard above.
        if (result < 0.5)
            throw new ArgumentException($"Invalid font size: '{value}'. Minimum font size is 0.5pt (one half-point).");
        // OOXML caps user-entered font size at 1638pt (Word) and Office
        // renderers stop honoring values past ~4000pt anyway. Anything
        // larger silently overflows the int32 the writers cast to (PPTX
        // writes pt × 100, Word writes pt × 2 as half-points), producing
        // negative w:sz / a:rPr@sz values Word rejects on open. Reject
        // up front with the same shape as the lower-bound guards.
        if (result > 4000)
            throw new ArgumentException($"Invalid font size: '{value}'. Maximum font size is 4000pt (Office cap).");
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
        // CONSISTENCY(color-input-whitespace): outer trim BEFORE any
        // hash-strip / hex inspection so " #FF0000 " / " red " / "FF0000\n"
        // (CLI pipes, JSON envelopes, copy-paste) parse the same as the bare
        // value. Inner whitespace ("#FF 0000") remains invalid.
        var trimmedInput = value.Trim();

        // Try named color / rgb()/rgba()/hsl()/hsla() first
        var resolved = TryResolveColorInput(trimmedInput);
        if (resolved is { } cr)
        {
            var alpha = cr.Alpha ?? 0xFF;
            return $"{alpha:X2}{cr.Rgb}";
        }

        var hadHashPrefix = trimmedInput.StartsWith('#');
        var hex = trimmedInput.TrimStart('#').ToUpperInvariant();
        if (hex.Length == 3 && hex.All(char.IsAsciiHexDigit))
        {
            // Expand shorthand: "F00" → "FF0000"
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });
        }
        else if (hadHashPrefix && hex.Length == 4 && hex.All(char.IsAsciiHexDigit)
                 && hex[3] != '0')
        {
            // CSS #RGBA shorthand → #RRGGBBAA. The hash prefix is required
            // because bare 4-hex would be ambiguous with truncated AARRGGBB
            // / RRGGBB-and-a-typo. Same expansion rule as #RGB.
            //
            // The non-zero-alpha guard rejects #XX00 (alpha = 0, fully
            // transparent): the user almost certainly mistyped a 6-digit
            // RGB, not "I want an invisible color via the 4-char shorthand"
            // (the explicit `transparent` keyword or 8-digit #00000000 form
            // exists for that intent and is unambiguous).
            hex = new string(new[]
            {
                hex[0], hex[0], hex[1], hex[1], hex[2], hex[2], hex[3], hex[3],
            });
        }
        if (hex.Length == 6 && hex.All(char.IsAsciiHexDigit))
            return "FF" + hex;
        if (hex.Length == 8 && hex.All(char.IsAsciiHexDigit))
        {
            // CONSISTENCY(color-input-form): #-prefixed 8-hex is CSS RRGGBBAA
            // (alpha last); bare 8-hex stays in OOXML AARRGGBB (alpha first).
            // Mirrors SanitizeColorForOoxml.
            if (hadHashPrefix)
                return hex.Substring(6, 2) + hex[..6];
            return hex;
        }
        throw new ArgumentException(
            $"Invalid color value: '{value}'. Expected 6-digit hex RGB (e.g. FF0000), " +
            $"8-digit AARRGGBB (e.g. 80FF0000), 3-digit shorthand (e.g. F00) or 4-digit #RGBA shorthand (e.g. F00A), " +
            $"named color (e.g. red), rgb()/rgba()/hsl()/hsla() notation, or 'transparent'.");
    }

    /// <summary>
    /// Word/PPT theme scheme color names (ECMA-376 §17.18.97 / §20.1.10.46).
    /// Keep lowercase — input is matched case-insensitively but the canonical
    /// OOXML serialization (and downstream readback) is lowercase.
    /// </summary>
    public static readonly HashSet<string> SchemeColorNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "dark1", "light1", "dark2", "light2",
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
        "hyperlink", "followedHyperlink",
        // Extra variants seen in OOXML: text1/text2/background1/background2 alias dark/light.
        "text1", "text2", "background1", "background2",
        // BUG-R6-06: alternate Word theme color aliases (windowText / windowBackground)
        // are valid OOXML w:themeColor values that map to dark1/light1.
        "windowText", "windowBackground",
        // BUG-R7-01: OOXML internal short forms used by PPT a:schemeClr@val.
        // Accept on input so NormalizeSchemeColorName can map them back to the
        // canonical user-facing names (dk1→dark1, lt1→light1, tx1→dark1, …).
        "dk1", "lt1", "dk2", "lt2", "tx1", "tx2", "bg1", "bg2",
        "hlink", "folHlink",
        "none", "auto",
    };

    /// <summary>
    /// True if <paramref name="value"/> is a recognized OOXML theme scheme
    /// color name (e.g. "accent1", "dark2", "hyperlink"). Comparison is
    /// case-insensitive; the canonical lowercase form is returned via
    /// <see cref="NormalizeSchemeColorName"/>.
    /// </summary>
    public static bool IsSchemeColorName(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        if (value!.StartsWith('#')) return false;
        return SchemeColorNames.Contains(value);
    }

    /// <summary>
    /// Returns the canonical lowercase scheme color name when
    /// <paramref name="value"/> is recognized; otherwise returns null.
    /// </summary>
    public static string? NormalizeSchemeColorName(string? value)
    {
        if (!IsSchemeColorName(value)) return null;
        var v = value!.ToLowerInvariant();
        // Canonicalize the text/background aliases (Excel/PPTX prefer
        // dark1/light1 in writes, but accept both on read).
        return v switch
        {
            // CONSISTENCY(scheme-color-roundtrip): text1/text2/background1/background2
            // are valid w:themeColor values in their own right (Word writes
            // either form depending on origin app). Pass through unchanged so
            // dump round-trip preserves the source's literal value. Only the
            // ambiguous-alias forms (windowText / OOXML short forms) need
            // canonicalisation.
            "text1" => "text1",
            "text2" => "text2",
            "background1" => "background1",
            "background2" => "background2",
            "windowtext" => "dark1",
            "windowbackground" => "light1",
            // OOXML internal short forms (used by PPT a:schemeClr@val).
            // dk/lt collapse to canonical dark/light (no separate user-facing
            // form). tx/bg map to text/background — the SDK distinguishes
            // SchemeColorValues.Text1 / Background1 from Dark1 / Light1, and
            // CLAUDE.md ("scheme colors pass through unchanged") demands the
            // user-supplied form survives Get; collapsing tx1→dark1 would
            // break that contract for text1/text2/background1/background2 set
            // by the user.
            "dk1" => "dark1",
            "dk2" => "dark2",
            "lt1" => "light1",
            "lt2" => "light2",
            "tx1" => "text1",
            "tx2" => "text2",
            "bg1" => "background1",
            "bg2" => "background2",
            "hlink" => "hyperlink",
            "folhlink" => "followedHyperlink",
            _ => v,
        };
    }

    /// <summary>
    /// Sanitize a hex color for OOXML srgbClr val (must be exactly 6-char RGB).
    /// If 8-char hex is given, interprets as AARRGGBB (POI convention: alpha first),
    /// strips the leading alpha and returns it separately.
    /// Returns (rgb6, alphaPercent) where alphaPercent is 0-100000 scale or null if fully opaque.
    /// </summary>
    public static (string Rgb, int? AlphaPercent) SanitizeColorForOoxml(string value)
    {
        // CONSISTENCY(color-input-whitespace): outer trim mirrors NormalizeArgbColor.
        var trimmedInput = value.Trim();

        // "auto" is a legal OOXML value for shading Fill/Color — pass through unchanged
        if (string.Equals(trimmedInput, "auto", StringComparison.OrdinalIgnoreCase))
            return ("auto", null);

        // Try named color / rgb()/rgba()/hsl()/hsla() first
        var resolved = TryResolveColorInput(trimmedInput);
        if (resolved is { } cr)
        {
            if (cr.Alpha is byte a && a != 0xFF)
                return (cr.Rgb, (int)(a / 255.0 * 100000));
            return (cr.Rgb, null);
        }

        // CONSISTENCY(color-input-form): treat the leading '#' as a signal that
        // the input follows the CSS #RRGGBBAA convention (alpha last). Bare
        // 8-hex (no '#') keeps the OOXML/POI AARRGGBB convention (alpha first).
        // Without this distinction, "#FFFFFFAA" was being parsed as AARRGGBB,
        // silently dropping the trailing AA byte and storing rgb=FFFFAA — the
        // user's RGB and alpha were both corrupted.
        var hadHashPrefix = trimmedInput.StartsWith('#');
        var hex = trimmedInput.TrimStart('#').ToUpperInvariant();
        if (hex.Length == 8 && hex.All(char.IsAsciiHexDigit))
        {
            byte alphaByte;
            string rgb;
            if (hadHashPrefix)
            {
                // CSS #RRGGBBAA — alpha is the trailing pair
                rgb = hex[..6];
                alphaByte = Convert.ToByte(hex.Substring(6, 2), 16);
            }
            else
            {
                // OOXML/POI AARRGGBB — alpha is the leading pair
                alphaByte = Convert.ToByte(hex[..2], 16);
                rgb = hex[2..];
            }
            if (alphaByte == 0xFF)
                return (rgb, null);
            var alphaPercent = (int)(alphaByte / 255.0 * 100000);
            return (rgb, alphaPercent);
        }
        // Validate: must be exactly 6 hex digits for srgbClr val
        if (hex.Length == 3 && hex.All(char.IsAsciiHexDigit))
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });
        else if (hadHashPrefix && hex.Length == 4 && hex.All(char.IsAsciiHexDigit)
                 && hex[3] != '0')
        {
            // CSS #RGBA shorthand → #RRGGBBAA; route through the 8-hex path
            // above to share alpha handling. See NormalizeArgbColor for the
            // non-zero-alpha-digit rationale (reject #XX00 as a typo guard).
            var expanded = new string(new[]
            {
                hex[0], hex[0], hex[1], hex[1], hex[2], hex[2], hex[3], hex[3],
            });
            var rgb = expanded[..6];
            var alphaByte = Convert.ToByte(expanded.Substring(6, 2), 16);
            if (alphaByte == 0xFF)
                return (rgb, null);
            return (rgb, (int)(alphaByte / 255.0 * 100000));
        }

        if (hex.Length != 6 || !hex.All(char.IsAsciiHexDigit))
        {
            // Scheme colors (accent1, dark2, hyperlink, …) are not handled
            // here — callers that support theme colors must check
            // IsSchemeColorName first and route to ThemeColor. Surface a
            // hint instead of advertising support we don't provide.
            var schemeHint = IsSchemeColorName(trimmedInput)
                ? " (scheme color names like 'accent1' must be set on properties that accept theme colors)"
                : "";
            throw new ArgumentException(
                $"Invalid color value: '{value}'. Expected 6-digit hex RGB (e.g. FF0000), " +
                $"8-digit AARRGGBB (e.g. 80FF0000), 3-digit shorthand (e.g. F00) or 4-digit #RGBA shorthand (e.g. F00A), " +
                $"named color (e.g. red), rgb()/rgba()/hsl()/hsla() notation, or 'transparent'." + schemeHint);
        }

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

    /// <summary>
    /// Reject XML 1.0 illegal control characters before they reach the OOXML
    /// serializer. Without this, the resident process accepts the value into
    /// the in-memory DOM and only fails at close-time with "save failed —
    /// data may be lost", losing the user's work. Allowed: \t (0x09), \n
    /// (0x0A), \r (0x0D). Rejected: 0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F.
    /// </summary>
    public static void ValidateXmlText(string? value, string propName)
    {
        if (value == null) return;
        for (int i = 0; i < value.Length; i++)
        {
            char c = value[i];
            if (c == '\t' || c == '\n' || c == '\r') continue;
            if (c < 0x20)
                throw new ArgumentException(
                    $"{propName} contains XML-illegal control character U+{(int)c:X4} at position {i}. " +
                    "Allowed control chars: \\t, \\n, \\r.");
        }
    }
}
