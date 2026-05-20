// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Dark-Style-2: accent-tinted body fills, dark last-row top divider.
/// Header uses a PAIRED accent (Accent1→Accent2, Accent3→Accent4, etc.)
/// for visual contrast against the body. Sparse — only neutral / Accent1
/// / Accent3 / Accent5 variants exist in PowerPoint.
/// </summary>
public static class DarkStyle2
{
    public static TableStyleDefinition Build(string accent)
    {
        var accentRef = string.IsNullOrEmpty(accent) ? "dk1" : accent.ToLowerInvariant();
        var pairedHeader = accent.ToLowerInvariant() switch
        {
            "accent1" => "accent2",
            "accent3" => "accent4",
            "accent5" => "accent6",
            _ => "dk1",
        };
        var darkLine = new BorderEdge("dk1");

        return new TableStyleDefinition
        {
            // Tint values match PowerPoint's empirical rendering on the
            // bare/dk1 variant (half-strength of the spec/LibreOffice
            // reference). Verified by OfficeShot pixel sampling.
            WholeTbl = new TableStyleRegion
            {
                Fill = new FillSpec(accentRef, Tint: 10000),
                TextColorRef = "dk1",
            },
            FirstRow = new TableStyleRegion
            {
                Fill = new FillSpec(pairedHeader),
                TextColorRef = "lt1",
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
