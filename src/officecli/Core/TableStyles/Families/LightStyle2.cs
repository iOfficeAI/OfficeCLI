// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles.Families;

/// <summary>
/// Light-Style-2: full accent-colored outer box, header in accent (full fill,
/// bg1/white text), last-row top accent line. Internal H lines appear between
/// banded rows via Band1HTop/Bottom borders (approximated here as wholeTbl
/// insideH for simplicity — full band1H border support requires extending
/// the data model with band-edge specific borders).
/// </summary>
public static class LightStyle2
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
                Fill = new FillSpec(accentRef),
                TextColorRef = "bg1",
            },
            LastRow = new TableStyleRegion
            {
                Top = line,
            },
        };
    }
}
