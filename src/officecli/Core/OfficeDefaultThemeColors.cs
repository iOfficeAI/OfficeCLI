// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Single source of truth for the Microsoft "Office" default color scheme
/// (the palette Word/Excel/PowerPoint apply when a document has no explicit
/// <c>a:theme</c> part). Used in two contexts:
///
///   1. Blank document creation — emitted into the theme1.xml we write.
///   2. Preview rendering fallback — when reading the doc's theme part
///      yields no <c>ColorScheme</c>, callers fall back to this palette
///      so <c>w:themeColor="accent1"</c> still resolves to a real hex
///      instead of silently dropping.
///
/// Hex values are 6-char OOXML format (no leading <c>#</c>).
/// </summary>
public static class OfficeDefaultThemeColors
{
    public const string Accent1 = "4472C4";
    public const string Accent2 = "ED7D31";
    public const string Accent3 = "A5A5A5";
    public const string Accent4 = "FFC000";
    public const string Accent5 = "5B9BD5";
    public const string Accent6 = "70AD47";

    public const string Dark1 = "000000";
    public const string Dark2 = "44546A";
    public const string Light1 = "FFFFFF";
    public const string Light2 = "E7E6E6";

    public const string Hyperlink = "0563C1";
    public const string FollowedHyperlink = "954F72";

    /// <summary>
    /// Default chart series color rotation when no <c>ColorScheme</c> is
    /// available. Slots 1-6 are the six accent colors; slots 7-12 are the
    /// same accents with <c>lumMod=75000</c> applied (the darker tints
    /// Office cycles through after exhausting the primary accents).
    ///
    /// Hex values are 6-char OOXML format (no leading <c>#</c>). Both the
    /// OOXML chart Builder and the SVG preview Renderer derive from this
    /// array — keep them aligned to avoid the chart-vs-preview drift.
    /// </summary>
    public static readonly string[] DefaultChartSeriesPalette =
    {
        Accent1, Accent2, Accent3, Accent4, Accent5, Accent6,
        "264478", "9E480E", "636363", "997300", "255E91", "43682B",
    };

    /// <summary>
    /// Builds a name→hex dictionary covering the canonical scheme keys plus
    /// the common aliases (dk1/tx1/text1, bg1/lt1/background1, …) that Word
    /// and PowerPoint accept as <c>w:themeColor</c> / <c>a:schemeClr</c>
    /// references. Used by HTML preview fallbacks.
    /// </summary>
    public static Dictionary<string, string> BuildAliasMap() => new(StringComparer.OrdinalIgnoreCase)
    {
        ["accent1"] = Accent1,
        ["accent2"] = Accent2,
        ["accent3"] = Accent3,
        ["accent4"] = Accent4,
        ["accent5"] = Accent5,
        ["accent6"] = Accent6,
        ["dark1"] = Dark1, ["tx1"] = Dark1, ["dk1"] = Dark1, ["text1"] = Dark1,
        ["dark2"] = Dark2, ["tx2"] = Dark2, ["dk2"] = Dark2, ["text2"] = Dark2,
        ["light1"] = Light1, ["bg1"] = Light1, ["lt1"] = Light1, ["background1"] = Light1,
        ["light2"] = Light2, ["bg2"] = Light2, ["lt2"] = Light2, ["background2"] = Light2,
        ["hyperlink"] = Hyperlink,
        ["followedHyperlink"] = FollowedHyperlink,
    };
}
