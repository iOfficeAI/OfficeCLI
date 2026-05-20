// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core.TableStyles.Families;

namespace OfficeCli.Core.TableStyles;

/// <summary>
/// Resolves PowerPoint built-in table styles to per-cell concrete values
/// (hex fill, hex text colour, four resolved cell-edge borders) given a
/// style id, a cell's position in the table, and the document's theme
/// colour map.
///
/// Region priority (highest wins): lastCol &gt; firstCol &gt; lastRow &gt;
/// firstRow &gt; band1H/band1V &gt; wholeTbl. Each property (Fill, TextColor,
/// each border edge) is taken from the highest-priority region that sets
/// it; lower regions cascade through where higher ones leave gaps.
/// </summary>
public static class TableStyleResolver
{
    /// <summary>
    /// Resolve a (styleIdOrName, cell position, theme) tuple to concrete
    /// values a renderer can emit. Returns null when styleIdOrName is
    /// unknown or maps to a family that has not been ported yet — caller
    /// is expected to fall back to a legacy code path during the
    /// incremental rollout.
    /// </summary>
    public static ResolvedCell? Resolve(
        string? styleIdOrName,
        CellPosition position,
        IReadOnlyDictionary<string, string> themeColors)
    {
        var familyAccent = TableStyleRegistry.Resolve(styleIdOrName);
        if (familyAccent == null) return null;
        var (family, accent) = familyAccent.Value;

        var def = BuildDefinition(family, accent);
        if (def == null) return null;

        return MergeRegions(def, position, themeColors);
    }

    /// <summary>
    /// Dispatch on family name to the appropriate Families/*.cs builder.
    /// Returns null for unknown family names. Add new family branches here
    /// alongside a matching Families/&lt;Name&gt;.cs file.
    /// </summary>
    private static TableStyleDefinition? BuildDefinition(string family, string accent) =>
        family switch
        {
            "Themed-Style-1" => ThemedStyle1.Build(accent),
            "Themed-Style-2" => ThemedStyle2.Build(accent),
            "Light-Style-1"  => LightStyle1.Build(accent),
            "Light-Style-2"  => LightStyle2.Build(accent),
            "Light-Style-3"  => LightStyle3.Build(accent),
            "Medium-Style-1" => MediumStyle1.Build(accent),
            "Medium-Style-2" => MediumStyle2.Build(accent),
            "Medium-Style-3" => MediumStyle3.Build(accent),
            "Medium-Style-4" => MediumStyle4.Build(accent),
            "Dark-Style-1"   => DarkStyle1.Build(accent),
            "Dark-Style-2"   => DarkStyle2.Build(accent),
            _ => null,
        };

    /// <summary>
    /// Walk the regions in priority order (lowest first, highest overrides)
    /// and merge their non-null fields into a ResolvedCell. This is where
    /// "firstRow.Bottom overrides wholeTbl.InsideH" cascades happen.
    /// </summary>
    private static ResolvedCell MergeRegions(
        TableStyleDefinition def,
        CellPosition pos,
        IReadOnlyDictionary<string, string> themeColors)
    {
        // Decide which regions apply to this cell. A cell can be in
        // multiple regions simultaneously — e.g. tc[0,0] of a table with
        // firstRow + firstCol enabled is in BOTH firstRow and firstCol.
        bool isFirstRow = pos.HasFirstRow && pos.RowIndex == 0;
        bool isLastRow  = pos.HasLastRow  && pos.RowIndex == pos.RowCount - 1;
        bool isFirstCol = pos.HasFirstCol && pos.ColIndex == 0;
        bool isLastCol  = pos.HasLastCol  && pos.ColIndex == pos.ColCount - 1;

        // Banded rows: alternate body rows (skipping firstRow if set).
        // Convention: first body row is "band1" (the tinted one); next is
        // unbanded; alternates. This matches PowerPoint's default rendering
        // of the firstRow=true bandedRows=true combo.
        bool isBand1H = false;
        if (pos.HasBandedRows && !isFirstRow && !isLastRow)
        {
            int bodyRowIdx = pos.RowIndex - (pos.HasFirstRow ? 1 : 0);
            isBand1H = bodyRowIdx >= 0 && bodyRowIdx % 2 == 0;
        }
        bool isBand1V = false;
        if (pos.HasBandedCols && !isFirstCol && !isLastCol)
        {
            int bodyColIdx = pos.ColIndex - (pos.HasFirstCol ? 1 : 0);
            isBand1V = bodyColIdx >= 0 && bodyColIdx % 2 == 0;
        }

        // Build a stack of regions in priority order: highest-priority
        // LAST so the merge loop's "last non-null wins" produces the
        // override semantics.
        var stack = new List<TableStyleRegion> { def.WholeTbl };
        if (isBand1V) stack.Add(def.Band1V);
        if (isBand1H) stack.Add(def.Band1H);
        if (isFirstRow) stack.Add(def.FirstRow);
        if (isLastRow)  stack.Add(def.LastRow);
        if (isFirstCol) stack.Add(def.FirstCol);
        if (isLastCol)  stack.Add(def.LastCol);

        FillSpec? fill = null;
        string? textColor = null;
        BorderEdge? top = null, bottom = null, left = null, right = null;
        BorderEdge? insideH = null, insideV = null;
        foreach (var r in stack)
        {
            if (r.Fill != null) fill = r.Fill;
            if (r.TextColorRef != null) textColor = r.TextColorRef;
            if (r.Top != null) top = r.Top;
            if (r.Bottom != null) bottom = r.Bottom;
            if (r.Left != null) left = r.Left;
            if (r.Right != null) right = r.Right;
            if (r.InsideH != null) insideH = r.InsideH;
            if (r.InsideV != null) insideV = r.InsideV;
        }

        // Map the cell's four physical edges to the right region edge:
        // outer rows/cols use Top/Bottom/Left/Right; inner positions use
        // InsideH (between rows) / InsideV (between cols).
        var resolvedTop    = pos.RowIndex == 0                  ? top    : insideH;
        var resolvedBottom = pos.RowIndex == pos.RowCount - 1   ? bottom : insideH;
        var resolvedLeft   = pos.ColIndex == 0                  ? left   : insideV;
        var resolvedRight  = pos.ColIndex == pos.ColCount - 1   ? right  : insideV;

        return new ResolvedCell(
            Fill:      ResolveFillToHex(fill, themeColors),
            TextColor: ResolveColorRefToHex(textColor, themeColors, tint: null),
            Top:       MaterializeBorder(resolvedTop, themeColors),
            Bottom:    MaterializeBorder(resolvedBottom, themeColors),
            Left:      MaterializeBorder(resolvedLeft, themeColors),
            Right:     MaterializeBorder(resolvedRight, themeColors));
    }

    private static ResolvedBorder? MaterializeBorder(
        BorderEdge? edge,
        IReadOnlyDictionary<string, string> themeColors)
    {
        if (edge == null) return null;
        var color = ResolveColorRefToHex(edge.ColorRef, themeColors, tint: null);
        if (color == null) return null;
        // Apply optional lumMod/lumOff (used by some styles for darker/lighter borders).
        if (edge.Lumination is { } lum)
            color = "#" + ColorMath.ApplyLumModOff(color.TrimStart('#'), lum.LumMod, lum.LumOff).TrimStart('#');
        return new ResolvedBorder(color, edge.WidthEmu, edge.Dash);
    }

    private static string? ResolveFillToHex(
        FillSpec? fill,
        IReadOnlyDictionary<string, string> themeColors)
    {
        if (fill == null) return null;
        var baseHex = ResolveColorRefToHex(fill.ColorRef, themeColors, tint: null);
        if (baseHex == null) return null;
        var hex = baseHex.TrimStart('#');
        if (fill.Tint is int tint)
            hex = BlendTowardWhite(hex, tint);
        if (fill.Shade is int shade)
            hex = BlendTowardBlack(hex, shade);
        return "#" + hex.ToUpperInvariant();
    }

    /// <summary>
    /// OOXML &lt;a:tint val="N"/&gt; — linear RGB blend toward white. Per
    /// ECMA-376 §20.1.2.3.30, "tint by N" yields N% of the base color
    /// combined with (100-N)% white: result = base*(N/100000) + 255*(1 - N/100000).
    /// Smaller tint values produce a LIGHTER colour (more white).
    /// </summary>
    private static string BlendTowardWhite(string hex, int tintMilliPercent)
    {
        var (r, g, b) = ParseHex(hex);
        double frac = tintMilliPercent / 100000.0;
        int br = (int)Math.Round(r * frac + 255 * (1 - frac));
        int bg = (int)Math.Round(g * frac + 255 * (1 - frac));
        int bb = (int)Math.Round(b * frac + 255 * (1 - frac));
        return $"{br:X2}{bg:X2}{bb:X2}";
    }

    /// <summary>
    /// OOXML &lt;a:shade val="N"/&gt; — linear RGB blend toward black:
    /// result = base*(N/100000) + 0*(1 - N/100000) = base*(N/100000).
    /// Smaller shade = darker.
    /// </summary>
    private static string BlendTowardBlack(string hex, int shadeMilliPercent)
    {
        var (r, g, b) = ParseHex(hex);
        double frac = shadeMilliPercent / 100000.0;
        int br = (int)Math.Round(r * frac);
        int bg = (int)Math.Round(g * frac);
        int bb = (int)Math.Round(b * frac);
        return $"{br:X2}{bg:X2}{bb:X2}";
    }

    private static (int r, int g, int b) ParseHex(string hex)
    {
        var s = hex.TrimStart('#');
        return (
            Convert.ToInt32(s.Substring(0, 2), 16),
            Convert.ToInt32(s.Substring(2, 2), 16),
            Convert.ToInt32(s.Substring(4, 2), 16));
    }

    /// <summary>
    /// Resolve a colour reference (scheme name like "dk1"/"lt1"/"accent1",
    /// or "#RRGGBB" hex) to a normalised hex string using the document's
    /// theme map. Returns null if reference is unresolvable.
    /// </summary>
    private static string? ResolveColorRefToHex(
        string? colorRef,
        IReadOnlyDictionary<string, string> themeColors,
        int? tint)
    {
        if (string.IsNullOrEmpty(colorRef)) return null;
        if (colorRef.StartsWith("#")) return colorRef.ToUpperInvariant();
        // Scheme color — look up in theme map.
        if (themeColors.TryGetValue(colorRef, out var hex))
            return "#" + hex.TrimStart('#').ToUpperInvariant();
        // Common fallbacks when the theme map does not include the slot.
        return colorRef.ToLowerInvariant() switch
        {
            "lt1" => "#FFFFFF",
            "dk1" => "#000000",
            _ => null,
        };
    }
}
