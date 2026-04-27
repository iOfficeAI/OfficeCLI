// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Precise error hints for Excel cell properties that are genuinely ambiguous
/// when carried over from PPT/Word habits.
///
/// Excel cells use a layered namespace (font.*, border.*, alignment.*, fill).
/// Most common PPT/Word flat keys — `size`, `font`, `halign`, `valign`, `wrap` —
/// are already accepted as aliases by ExcelStyleManager because they have a
/// single unambiguous meaning in cell context.
///
/// This class lists the keys that cannot be safely aliased because they mean
/// two different things. For those we refuse silent mapping and return a
/// precise hint telling the user to pick one explicitly.
/// </summary>
internal static class CellPropHints
{
    private static readonly Dictionary<string, string> AmbiguousKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // `color` in PPT/Word run context means text color, but in Excel cells
        // the user might intuitively expect background color. Force them to
        // pick: `font.color` (text) or `fill` (background).
        ["color"] = "ambiguous in cell context — use 'font.color' for text color or 'fill' for background color",
        // R17 bt-3: `path=` looks plausible (path-like keys exist for picture/ole)
        // but cell uses `ref=` (or `address=`) for the target address. Silently
        // dropping `path` writes the value to the wrong cell — fail loudly.
        ["path"] = "not a cell property — use 'ref' (or 'address') for the cell address, e.g. --prop ref=D5",
    };

    /// <summary>
    /// If the given key is a known ambiguous cell prop, returns a human-readable
    /// hint telling the user to pick an unambiguous alternative. Returns null
    /// otherwise.
    /// </summary>
    public static string? TryGetHint(string key)
    {
        if (!AmbiguousKeys.TryGetValue(key, out var hint))
            return null;

        return $"{key} ({hint})";
    }
}
