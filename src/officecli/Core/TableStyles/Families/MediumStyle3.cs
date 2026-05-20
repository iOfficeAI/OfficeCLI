// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Medium-Style-3: minimal — only outer top/bottom + header underline +
/// last-row top line, all in dk1 (black). No vertical borders, no internal
/// horizontals between body rows; banding (light tint) provides separation.
/// </summary>
public static class MediumStyle3
{
    public static TableStyleDefinition Build(string accent)
    {
        var accentRef = string.IsNullOrEmpty(accent) ? "dk1" : accent.ToLowerInvariant();
        var dkLine = new BorderEdge("dk1");

        return new TableStyleDefinition
        {
            WholeTbl = new TableStyleRegion
            {
                Top = dkLine, Bottom = dkLine,
                Fill = new FillSpec("lt1"),
                TextColorRef = "dk1",
            },
            FirstRow = new TableStyleRegion
            {
                Bottom = dkLine,
                Fill = new FillSpec(accentRef),
                TextColorRef = "lt1",
            },
            LastRow = new TableStyleRegion
            {
                Top = dkLine,
                Fill = new FillSpec("lt1"),
            },
            FirstCol = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef),
                TextColorRef = "lt1",
            },
            LastCol = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef),
                TextColorRef = "lt1",
            },
            Band1H = new TableStyleRegion { Fill = new FillSpec("dk1", Tint: 20000) },
            Band1V = new TableStyleRegion { Fill = new FillSpec("dk1", Tint: 20000) },
        };
    }
}
