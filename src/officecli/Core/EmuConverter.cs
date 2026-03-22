// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;

namespace OfficeCli.Core;

/// <summary>
/// Shared EMU (English Metric Unit) parsing and formatting.
/// 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU, 1 px = 9525 EMU.
/// Accepts: raw EMU integer, or suffixed with cm/in/pt/px.
/// </summary>
public static class EmuConverter
{
    /// <summary>
    /// Parse a dimension string into EMU (long).
    /// Supported formats: "914400" (raw EMU), "2.54cm", "1in", "72pt", "96px".
    /// Throws ArgumentException on invalid input.
    /// </summary>
    public static long ParseEmu(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("EMU value cannot be null or empty.");

        value = value.Trim();

        long result;

        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 360000.0, "cm");
        }
        else if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 914400.0, "in");
        }
        else if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 12700.0, "pt");
        }
        else if (value.EndsWith("px", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 9525.0, "px");
        }
        else if (HasKnownUnitSuffix(value, out var unit))
        {
            throw new ArgumentException($"Unsupported unit '{unit}' in dimension value '{value}'. Supported units: cm, in, pt, px (or raw EMU integer).");
        }
        else
        {
            // Raw EMU integer
            if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
                throw new ArgumentException($"Invalid EMU value '{value}'. Expected a number with optional unit suffix (cm, in, pt, px).");
        }

        if (result < 0)
            throw new ArgumentException($"Negative dimensions are not allowed: '{value}'.");

        return result;
    }

    /// <summary>
    /// Parse EMU and safely cast to int, throwing on overflow.
    /// </summary>
    public static int ParseEmuAsInt(string value)
    {
        long emu = ParseEmu(value);
        if (emu < 0)
            throw new ArgumentException($"Negative dimension value '{value}' is not allowed. This property requires a non-negative value.");
        if (emu > int.MaxValue)
            throw new OverflowException($"EMU value {emu} (from '{value}') exceeds the maximum allowed value of {int.MaxValue}.");
        return (int)emu;
    }

    /// <summary>
    /// Parse line width value into EMU (int). Bare numbers are treated as points (pt),
    /// matching Apache POI's setLineWidth() behavior. Suffixed values (cm/in/pt/px) are
    /// parsed normally via ParseEmu.
    /// </summary>
    public static int ParseLineWidth(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Line width value cannot be null or empty.");

        var trimmed = value.Trim();
        // If bare integer/decimal with no unit suffix, treat as points
        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
            && !HasKnownUnitSuffix(trimmed, out _))
        {
            trimmed += "pt";
        }
        return ParseEmuAsInt(trimmed);
    }

    /// <summary>
    /// Format an EMU value as a human-readable string (e.g., "2.54cm").
    /// </summary>
    public static string FormatEmu(long emu)
    {
        var cm = emu / 360000.0;
        return $"{cm:0.##}cm";
    }

    /// <summary>
    /// Try to parse a dimension string into EMU. Returns false if parsing fails.
    /// </summary>
    public static bool TryParseEmu(string value, out long emu)
    {
        try
        {
            emu = ParseEmu(value);
            return true;
        }
        catch
        {
            emu = 0;
            return false;
        }
    }

    private static long ParseWithUnit(string value, int suffixLen, double factor, string unit)
    {
        var numberPart = value[..^suffixLen];
        if (string.IsNullOrWhiteSpace(numberPart))
            throw new ArgumentException($"Missing numeric value before '{unit}' unit in '{value}'.");

        if (!double.TryParse(numberPart, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || double.IsNaN(number) || double.IsInfinity(number))
            throw new ArgumentException($"Invalid numeric value '{numberPart}' before '{unit}' unit in '{value}'.");

        return (long)Math.Round(number * factor);
    }

    private static bool HasKnownUnitSuffix(string value, out string unit)
    {
        // Check for common but unsupported units
        string[] unsupported = { "mm", "rem", "em", "ex", "pc", "vw", "vh" };
        foreach (var u in unsupported)
        {
            if (value.EndsWith(u, StringComparison.OrdinalIgnoreCase))
            {
                unit = u;
                return true;
            }
        }
        unit = "";
        return false;
    }
}
