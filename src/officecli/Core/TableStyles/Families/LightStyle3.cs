// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Light-Style-3: full accent grid (all 6 wholeTbl borders + insideH/V),
/// no header fill, accent text colour in header with underline. Banded
/// rows at 20% alpha.
/// </summary>
public static class LightStyle3
{
    public static TableStyleDefinition Build(string accent)
    {
        var accentRef = string.IsNullOrEmpty(accent) ? "tx1" : accent.ToLowerInvariant();
        var line = new BorderEdge(accentRef);

        return new TableStyleDefinition
        {
            WholeTbl = new TableStyleRegion
            {
                Top = line, Bottom = line, Left = line, Right = line,
                InsideH = line, InsideV = line,
                TextColorRef = "tx1",
            },
            FirstRow = new TableStyleRegion
            {
                Bottom = line,
                TextColorRef = accentRef,
            },
            LastRow = new TableStyleRegion { Top = line },
            Band1H = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
            Band1V = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
        };
    }
}
