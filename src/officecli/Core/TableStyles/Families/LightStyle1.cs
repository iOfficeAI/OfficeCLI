// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Light-Style-1: just three horizontal lines (table top, header bottom,
/// table bottom) in accent color; no verticals, no internal cell borders.
/// Banded rows at 20% alpha (treated as tint 20%) provide row separation.
/// </summary>
public static class LightStyle1
{
    public static TableStyleDefinition Build(string accent)
    {
        var accentRef = string.IsNullOrEmpty(accent) ? "tx1" : accent.ToLowerInvariant();
        var line = new BorderEdge(accentRef);

        return new TableStyleDefinition
        {
            WholeTbl = new TableStyleRegion
            {
                Top = line, Bottom = line,
                TextColorRef = "tx1",
            },
            FirstRow = new TableStyleRegion
            {
                Bottom = line,
                TextColorRef = "tx1",
            },
            LastRow = new TableStyleRegion
            {
                Top = line,
            },
            LastCol = new TableStyleRegion
            {
                TextColorRef = "tx1",
            },
            Band1H = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
            Band1V = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
        };
    }
}
