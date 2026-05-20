// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Dark-Style-1: solid accent fills throughout (no internal borders). When
/// no accent is set, fills use dk1 (black) with tint to lighten body /
/// bands. When an accent is set, fills stay strong with SHADE (darkening
/// toward black). White borders separate header from body / first / last
/// columns. Header on dk1 background with lt1 text.
/// </summary>
public static class DarkStyle1
{
    public static TableStyleDefinition Build(string accent)
    {
        bool hasAccent = !string.IsNullOrEmpty(accent);
        var accentRef = hasAccent ? accent.ToLowerInvariant() : "dk1";
        var ltLine = new BorderEdge("lt1");

        // Transform direction depends on whether an accent is set:
        // no accent → tint (lighten dk1 toward white);
        // accent set → shade (darken accent toward black).
        FillSpec accentFill(int amount) =>
            hasAccent ? new FillSpec(accentRef, Shade: amount)
                      : new FillSpec(accentRef, Tint: amount);

        // Tint values match PowerPoint's empirical rendering (half-strength
        // of LibreOffice's predefined-table-styles.cxx reference). Shade
        // (accent-set variants) follow the same halving for consistency.
        return new TableStyleDefinition
        {
            WholeTbl = new TableStyleRegion
            {
                Fill = accentFill(10000),
                TextColorRef = "dk1",
            },
            FirstRow = new TableStyleRegion
            {
                Bottom = ltLine,
                Fill = new FillSpec("dk1"),
                TextColorRef = "lt1",
            },
            LastRow = new TableStyleRegion
            {
                Top = ltLine,
                Fill = new FillSpec(accentRef),
            },
            FirstCol = new TableStyleRegion
            {
                Right = ltLine,
                Fill = accentFill(30000),
            },
            LastCol = new TableStyleRegion
            {
                Left = ltLine,
                Fill = accentFill(30000),
            },
            Band1H = new TableStyleRegion { Fill = accentFill(20000) },
            Band1V = new TableStyleRegion { Fill = accentFill(20000) },
        };
    }
}
