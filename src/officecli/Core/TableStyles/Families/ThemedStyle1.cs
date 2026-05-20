// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Themed-Style-1: minimal "no-style" — no fills, no borders, text in tx1
/// when neutral. Accent variants paint a full accent-coloured grid on every
/// edge of every region, plus banded rows at 40% alpha (treated as tint).
/// </summary>
public static class ThemedStyle1
{
    public static TableStyleDefinition Build(string accent)
    {
        if (string.IsNullOrEmpty(accent))
        {
            // Neutral: no borders, no fill, just text colour.
            return new TableStyleDefinition
            {
                WholeTbl = new TableStyleRegion { TextColorRef = "tx1" },
                Band1H = new TableStyleRegion { /* alpha 40% on a non-fill is a no-op */ },
                Band1V = new TableStyleRegion { /* same */ },
            };
        }
        var accentRef = accent.ToLowerInvariant();
        var line = new BorderEdge(accentRef);

        return new TableStyleDefinition
        {
            WholeTbl = new TableStyleRegion
            {
                Top = line, Bottom = line, Left = line, Right = line,
                InsideH = line, InsideV = line,
                TextColorRef = "dk1",
            },
            FirstRow = new TableStyleRegion
            {
                Top = line, Bottom = new BorderEdge("lt1"),
                Left = line, Right = line,
                Fill = new FillSpec(accentRef),
                TextColorRef = "lt1",
            },
            LastRow = new TableStyleRegion
            {
                Top = line, Bottom = line, Left = line, Right = line,
            },
            FirstCol = new TableStyleRegion
            {
                Top = line, Bottom = line, Left = line, Right = line, InsideH = line,
            },
            LastCol = new TableStyleRegion
            {
                Top = line, Bottom = line, Left = line, Right = line, InsideH = line,
            },
            Band1H = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 40000) },
            Band1V = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 40000) },
        };
    }
}
