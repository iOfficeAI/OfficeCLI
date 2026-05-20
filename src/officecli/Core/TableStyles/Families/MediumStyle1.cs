// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Medium-Style-1: accent-colored outer box + insideH lines, white body
/// fill, accent header band (dark on accent, light text), accent-tinted
/// banded rows.
/// </summary>
public static class MediumStyle1
{
    public static TableStyleDefinition Build(string accent)
    {
        var accentRef = string.IsNullOrEmpty(accent) ? "dk1" : accent.ToLowerInvariant();
        var line = new BorderEdge(accentRef);

        return new TableStyleDefinition
        {
            WholeTbl = new TableStyleRegion
            {
                Top = line, Bottom = line, Left = line, Right = line,
                InsideH = line,
                Fill = new FillSpec("lt1"),
                TextColorRef = "dk1",
            },
            FirstRow = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef),
                TextColorRef = "lt1",
            },
            LastRow = new TableStyleRegion
            {
                Top = line,
                Fill = new FillSpec("lt1"),
            },
            Band1H = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
            Band1V = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
        };
    }
}
