// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Medium-Style-2: the iconic "tile" look — white 1pt borders between every
/// cell ("grout"), header row in raw accent colour with light text, body in
/// 20%-tinted accent, banded rows alternating with 40%-tint. Used by the
/// historical CLI short name `style=medium2` (Accent1 variant by default).
/// </summary>
public static class MediumStyle2
{
    /// <summary>
    /// Build the style definition for the given accent name (e.g. "Accent1"
    /// or "" for the neutral / dk1 variant).
    /// </summary>
    public static TableStyleDefinition Build(string accent)
    {
        // Empty accent → fall back to dk1 (dark scheme colour).
        var accentRef = string.IsNullOrEmpty(accent) ? "dk1" : accent.ToLowerInvariant();

        // White 1pt border used everywhere ("grout").
        var grout = new BorderEdge("lt1");

        return new TableStyleDefinition
        {
            // Whole table: full grid of white borders; body fill = accent
            // at 20% tint (lightened toward white); default text = dk1.
            // Tint values match PowerPoint's empirical rendering of the
            // bare/dk1 variant: wholeTbl ≈ #E7E7E7 (tint 10000), band1H ≈
            // #CBCBCB (tint 20000). LibreOffice's predefined-table-styles.cxx
            // documents these as tint 20000 / 40000 but PowerPoint's actual
            // output is half-strength — verified by OfficeShot pixel sampling.
            WholeTbl = new TableStyleRegion
            {
                Top = grout, Bottom = grout, Left = grout, Right = grout,
                InsideH = grout, InsideV = grout,
                Fill = new FillSpec(accentRef, Tint: 10000),
                TextColorRef = "dk1",
            },
            // Header row: full accent (no tint, the strong colour band);
            // light text for contrast.
            FirstRow = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef),
                TextColorRef = "lt1",
            },
            LastRow = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef),
                TextColorRef = "lt1",
                Top = grout,
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
            // Banded rows: every other body row gets a darker tint than
            // the wholeTbl baseline — accent at 40% tint.
            Band1H = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef, Tint: 20000),
            },
            Band1V = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef, Tint: 20000),
            },
        };
    }
}
