// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;

namespace OfficeCli.Core;

/// <summary>
/// Shared EMU (English Metric Unit) parsing and formatting.
/// 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU, 1 px = 9525 EMU.
/// Accepts: raw EMU integer, or suffixed with cm/in/pt/px/emu.
/// </summary>
internal static class EmuConverter
{
    /// <summary>
    /// Parse a dimension/position string into EMU (long).
    /// Supported formats: "914400" (raw EMU), "914400emu", "2.54cm", "1in", "72pt", "96px".
    /// Negative values are allowed (for positions like x, y).
    /// Throws ArgumentException on invalid input.
    /// </summary>
    public static long ParseEmu(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Invalid length value: input is null or empty.");

        value = value.Trim();

        long result;

        // mm = cm/10. CSS-natural sibling of cm; AI assistants reach for mm
        // before pt/in. Test in longer-suffix order so "mm" is matched before
        // any bare-number fallback. ("Q" = mm/4 in CSS, also accepted.)
        if (value.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 36000.0, "mm");
        }
        else if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 360000.0, "cm");
        }
        else if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 914400.0, "in");
        }
        else if (value.EndsWith("pc", StringComparison.OrdinalIgnoreCase))
        {
            // pica = 12pt (CSS / typographic standard).
            result = ParseWithUnit(value, 2, 12.0 * 12700.0, "pc");
        }
        else if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 12700.0, "pt");
        }
        else if (value.EndsWith("px", StringComparison.OrdinalIgnoreCase))
        {
            result = ParseWithUnit(value, 2, 9525.0, "px");
        }
        else if (value.EndsWith("Q", StringComparison.OrdinalIgnoreCase))
        {
            // CSS quarter-millimeter (Q = mm/4). Rare but harmless to support.
            result = ParseWithUnit(value, 1, 9000.0, "Q");
        }
        else if (value.EndsWith("emu", StringComparison.OrdinalIgnoreCase))
        {
            // Explicit emu suffix — symmetric with FormatEmu's tiny-value fallback.
            var numberPart = value[..^3];
            if (string.IsNullOrWhiteSpace(numberPart))
                throw new ArgumentException($"Invalid length value '{value}': missing numeric value before 'emu' unit.");
            if (!long.TryParse(numberPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
                throw new ArgumentException($"Invalid length value '{value}': '{numberPart}' before 'emu' unit is not an integer.");
        }
        else if (HasKnownUnitSuffix(value, out var unit))
        {
            // 'em' / 'ex' / 'rem' depend on font context that ParseEmu does not have.
            throw new ArgumentException(
                $"Invalid length value '{value}': unit '{unit}' is not supported (requires font context). " +
                $"Supported units: cm, mm, in, pt, pc, px, Q, emu (or raw EMU integer).");
        }
        else
        {
            // Raw EMU value: integer form preferred. Bare scientific notation
            // ("1e6", "1.5e2") is also accepted for parity with ParseFontSize
            // — but a plain decimal like "0.0001" is NOT (fractional EMU is
            // meaningless; users wanting sub-EMU sizes have unit suffixes).
            // Reject NaN/Infinity early so downstream emit never produces a
            // "NaNpt" / "Infinitypt" string.
            if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
            {
                // already parsed
            }
            else if ((value.Contains('e') || value.Contains('E'))
                     && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
            {
                if (double.IsNaN(d) || double.IsInfinity(d))
                    throw new ArgumentException(
                        $"Invalid length value '{value}': must be a finite number.");
                if (d > long.MaxValue || d < long.MinValue)
                    throw new ArgumentException(
                        $"Invalid length value '{value}': exceeds the maximum EMU range.");
                result = (long)Math.Round(d);
            }
            else
            {
                throw new ArgumentException(
                    $"Invalid length value '{value}': expected a number with optional unit suffix (cm, mm, in, pt, pc, px, Q, emu).");
            }
        }

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
            throw new ArgumentException("Invalid line width value: input is null or empty.");

        var trimmed = value.Trim();
        // If bare integer/decimal with no unit suffix, treat as points.
        // Reject NaN/Infinity here so we never construct a "NaNpt" / "Infinitypt"
        // string that would slip past ParseEmu's suffix branch with a misleading
        // error tail.
        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var bare)
            && !HasKnownUnitSuffix(trimmed, out _))
        {
            if (double.IsNaN(bare) || double.IsInfinity(bare))
                throw new ArgumentException($"Invalid line width value '{value}': must be a finite number.");
            trimmed += "pt";
        }
        return ParseEmuAsInt(trimmed);
    }

    /// <summary>
    /// Format an EMU value as a human-readable string (e.g., "2.54cm").
    /// </summary>
    public static string FormatEmu(long emu)
    {
        if (emu == 0) return "0cm";
        var cm = emu / 360000.0;
        var cmStr = cm.ToString("0.##", CultureInfo.InvariantCulture);
        // The "0.##" cm format loses precision below ~3600 EMU per side
        // (less than 0.01cm rounds away). For values that round either
        // to "0"/"-0" OR re-parse back to a different EMU than the source,
        // fall back to a `<n>emu` form so Get readback is both non-lossy
        // AND unit-qualified — round-trips through ParseEmu and satisfies
        // the documented length-string readback contract.
        if (cmStr == "0" || cmStr == "-0")
            return emu.ToString(CultureInfo.InvariantCulture) + "emu";
        // Round-trip sanity for sub-0.01cm values: anything under 3600 EMU
        // (= 0.01cm) can't be expressed faithfully in "0.##" cm form (e.g.
        // 1800 EMU → "0.01cm" → re-parses as 3600 EMU, doubling the source).
        // Switch to raw emu only in that narrow band; values ≥ 3600 EMU keep
        // the cm output (existing baselines unchanged), accepting the
        // documented 0.01cm grid quantization for larger sizes.
        if (Math.Abs(emu) < 3600)
            return emu.ToString(CultureInfo.InvariantCulture) + "emu";
        return $"{cmStr}cm";
    }

    /// <summary>
    /// Format an EMU value as points (e.g., "2pt"). Used for line widths and other
    /// thin values where points are more natural than centimeters.
    /// </summary>
    public static string FormatLineWidth(long emu)
    {
        var pt = emu / 12700.0;
        return $"{pt:0.##}pt";
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
        // Font-relative or viewport-relative units that ParseEmu cannot resolve
        // (no font / viewport context at this layer). Listed so the bare-number
        // fallback rejects them with a "not supported" message instead of
        // silently re-interpreting "5em" as raw EMU.
        // Order matters: longer suffixes first so "rem" doesn't shadow "em".
        string[] unsupported = { "rem", "em", "ex", "vw", "vh" };
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
