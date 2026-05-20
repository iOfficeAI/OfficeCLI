// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.TableStyles;

/// <summary>
/// Data types for the PowerPoint built-in table style catalogue.
///
/// PowerPoint ships 74 built-in table styles (11 family templates × accent
/// variants) referenced from pptx files by GUID. The full <a:tblStyle>
/// definitions live in PowerPoint's binary, NOT in any OOXML file — every
/// third-party viewer maintains its own equivalent catalogue. These records
/// mirror the structure of an <a:tblStyle>: a family template parameterised
/// on accent colour, expanded into seven cell regions (wholeTbl + four edge
/// overrides + two banding regions), each carrying optional fill, text
/// colour, and six border edges.
/// </summary>
public static class TableStyleConstants
{
    /// <summary>Default border line width: 12700 EMU = 1pt = 1/72 inch.</summary>
    public const int DefaultBorderEmu = 12700;
}

/// <summary>
/// One edge of a cell border. Null means "no border defined at this region
/// for this edge" — region merging cascades to the next-lower-priority region
/// (e.g. firstRow.Top unset falls back to wholeTbl.Top).
/// </summary>
/// <param name="ColorRef">Theme color reference: scheme name ("lt1", "dk1",
/// "accent1"..) or "#RRGGBB" hex. Renderer resolves via themeColors map.</param>
/// <param name="Lumination">Optional lumMod/lumOff transformation as a
/// (lumMod, lumOff) tuple in 1/1000 percent; null = no transform.</param>
/// <param name="WidthEmu">Line width in EMU (1pt = 12700).</param>
/// <param name="Dash">OOXML dash style: "solid", "dot", "dash", "lgDash",
/// "dashDot", "sysDot", "sysDash", etc. Default "solid".</param>
public record BorderEdge(
    string ColorRef,
    int WidthEmu = TableStyleConstants.DefaultBorderEmu,
    string Dash = "solid",
    (int LumMod, int LumOff)? Lumination = null);

/// <summary>
/// Fill specification for one region. Null TableStyleRegion.Fill means "no
/// fill defined here"; cell falls back to lower-priority region's fill.
/// </summary>
/// <param name="ColorRef">Theme color reference: scheme name or hex.</param>
/// <param name="Tint">OOXML tint transformation in 1/1000 percent
/// (e.g. 20000 = 20% tint towards white). Null = no transform.</param>
/// <param name="Shade">OOXML shade transformation (toward black).</param>
public record FillSpec(string ColorRef, int? Tint = null, int? Shade = null);

/// <summary>
/// One region of a table style — wholeTbl, firstRow, lastRow, firstCol,
/// lastCol, band1H (horizontal banding), or band1V (vertical banding).
/// All fields optional: null means "this region does not override". The
/// resolver merges regions in priority order.
/// </summary>
public record TableStyleRegion
{
    public FillSpec? Fill { get; init; }
    public string? TextColorRef { get; init; }

    // Six per-edge border specs. Outer edges (top/bottom/left/right) apply
    // when the cell sits on the table perimeter. Inside edges (insideH/V)
    // apply between adjacent cells inside the table.
    public BorderEdge? Top { get; init; }
    public BorderEdge? Bottom { get; init; }
    public BorderEdge? Left { get; init; }
    public BorderEdge? Right { get; init; }
    public BorderEdge? InsideH { get; init; }
    public BorderEdge? InsideV { get; init; }
}

/// <summary>
/// Complete definition of a table style — what a single GUID resolves to
/// after accent colour substitution. Built by family constructors (one per
/// family template) and consumed by the resolver to compute per-cell styles.
/// </summary>
public record TableStyleDefinition
{
    public TableStyleRegion WholeTbl { get; init; } = new();
    public TableStyleRegion FirstRow { get; init; } = new();
    public TableStyleRegion LastRow { get; init; } = new();
    public TableStyleRegion FirstCol { get; init; } = new();
    public TableStyleRegion LastCol { get; init; } = new();
    public TableStyleRegion Band1H { get; init; } = new();
    public TableStyleRegion Band1V { get; init; } = new();
}

/// <summary>
/// Where a cell sits in its table. Carries the flags needed to decide which
/// regions apply (and in what priority order).
/// </summary>
/// <param name="RowIndex">0-based row index.</param>
/// <param name="ColIndex">0-based column index.</param>
/// <param name="RowCount">Total rows in the table.</param>
/// <param name="ColCount">Total columns in the table.</param>
/// <param name="HasFirstRow">Table's firstRow flag (header styling enabled).</param>
/// <param name="HasLastRow">Table's lastRow flag (footer styling enabled).</param>
/// <param name="HasFirstCol">Table's firstCol flag.</param>
/// <param name="HasLastCol">Table's lastCol flag.</param>
/// <param name="HasBandedRows">Banded-rows flag.</param>
/// <param name="HasBandedCols">Banded-cols flag.</param>
public record CellPosition(
    int RowIndex,
    int ColIndex,
    int RowCount,
    int ColCount,
    bool HasFirstRow,
    bool HasLastRow,
    bool HasFirstCol,
    bool HasLastCol,
    bool HasBandedRows,
    bool HasBandedCols);

/// <summary>
/// Result of resolving a style + cell position + theme into the concrete
/// values a renderer needs. Hex strings are normalised to "#RRGGBB" without
/// alpha. Null fields mean "no value" — renderer should fall back to its
/// own default (typically transparent fill / no border).
/// </summary>
public record ResolvedCell(
    string? Fill,         // hex "#RRGGBB"
    string? TextColor,    // hex "#RRGGBB"
    ResolvedBorder? Top,
    ResolvedBorder? Bottom,
    ResolvedBorder? Left,
    ResolvedBorder? Right);

/// <summary>One resolved cell-edge border: colour + width + dash.</summary>
public record ResolvedBorder(string Color, int WidthEmu, string Dash);
