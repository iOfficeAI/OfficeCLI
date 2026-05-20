// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Medium-Style-4: accent-colored outer box + insideH/V grid, accent
/// header band (tint 20%), accent-tinted body and banded rows; dark
/// last-row top underline. Looks like a "full grid in accent colour".
/// </summary>
public static class MediumStyle4
{
    public static TableStyleDefinition Build(string accent)
    {
        var accentRef = string.IsNullOrEmpty(accent) ? "dk1" : accent.ToLowerInvariant();
        var accentLine = new BorderEdge(accentRef);
        var darkLine = new BorderEdge("dk1");

        // Tint values match PowerPoint's empirical rendering (half-strength
        // of LibreOffice's predefined-table-styles.cxx reference). Verified
        // by OfficeShot pixel sampling on the bare/dk1 variant.
        return new TableStyleDefinition
        {
            WholeTbl = new TableStyleRegion
            {
                Top = accentLine, Bottom = accentLine,
                Left = accentLine, Right = accentLine,
                InsideH = accentLine, InsideV = accentLine,
                Fill = new FillSpec(accentRef, Tint: 10000),
                TextColorRef = "dk1",
            },
            FirstRow = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef, Tint: 10000),
                TextColorRef = accentRef,
            },
            LastRow = new TableStyleRegion
            {
                Top = darkLine,
                Fill = new FillSpec(accentRef, Tint: 10000),
            },
            Band1H = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
            Band1V = new TableStyleRegion { Fill = new FillSpec(accentRef, Tint: 20000) },
        };
    }
}
