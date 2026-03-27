// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Chart style presets — curated property combinations for professional chart styling.
/// Applied via Set(chart, { ["preset"] = "minimal" }).
/// </summary>
internal static class ChartPresets
{
    internal static Dictionary<string, string>? GetPreset(string presetName)
    {
        return presetName.ToLowerInvariant() switch
        {
            "minimal" => Minimal,
            "dark" => Dark,
            "corporate" => Corporate,
            "magazine" => Magazine,
            "dashboard" => Dashboard,
            "colorful" => Colorful,
            "monochrome" or "mono" => Monochrome,
            _ => null
        };
    }

    internal static readonly string[] PresetNames =
    {
        "minimal", "dark", "corporate", "magazine", "dashboard", "colorful", "monochrome"
    };

    /// <summary>
    /// Minimal: clean, light, emphasis on data. Thin gray gridlines, no borders, small labels.
    /// </summary>
    private static readonly Dictionary<string, string> Minimal = new()
    {
        ["gridlines"] = "E0E0E0:0.3",
        ["minorGridlines"] = "none",
        ["plotArea.border"] = "none",
        ["chartArea.border"] = "none",
        ["axisLine"] = "none",
        ["majorTickMark"] = "none",
        ["minorTickMark"] = "none",
        ["tickLabelPos"] = "nextTo",
        ["axisFont"] = "9:808080",
        ["legend"] = "bottom",
        ["legendFont"] = "9:808080",
        ["plotFill"] = "none",
        ["chartFill"] = "none",
        ["roundedCorners"] = "false",
        ["colors"] = "4472C4,ED7D31,A5A5A5,FFC000,5B9BD5,70AD47",
    };

    /// <summary>
    /// Dark: dark background, bright data, white text. Suitable for presentations on dark slides.
    /// </summary>
    private static readonly Dictionary<string, string> Dark = new()
    {
        ["chartFill"] = "1E1E1E",
        ["plotFill"] = "2D2D2D",
        ["gridlines"] = "404040:0.3",
        ["minorGridlines"] = "none",
        ["axisLine"] = "555555:0.5",
        ["majorTickMark"] = "none",
        ["axisFont"] = "9:AAAAAA",
        ["legendFont"] = "9:CCCCCC",
        ["legend"] = "bottom",
        ["title.color"] = "FFFFFF",
        ["title.size"] = "16",
        ["plotArea.border"] = "none",
        ["chartArea.border"] = "444444:0.5",
        ["roundedCorners"] = "true",
        ["colors"] = "5B9BD5,FF6B6B,51CF66,FCC419,CC5DE8,22B8CF",
    };

    /// <summary>
    /// Corporate: professional blue-gray palette, clean axes, suitable for business reports.
    /// </summary>
    private static readonly Dictionary<string, string> Corporate = new()
    {
        ["gridlines"] = "D6DCE4:0.4",
        ["minorGridlines"] = "none",
        ["axisLine"] = "8B949E:0.5",
        ["majorTickMark"] = "out",
        ["minorTickMark"] = "none",
        ["axisFont"] = "10:44546A",
        ["legendFont"] = "10:44546A",
        ["legend"] = "right",
        ["title.bold"] = "true",
        ["title.size"] = "14",
        ["title.color"] = "44546A",
        ["plotFill"] = "none",
        ["chartFill"] = "none",
        ["plotArea.border"] = "D6DCE4:0.3",
        ["chartArea.border"] = "none",
        ["roundedCorners"] = "false",
        ["colors"] = "2E75B6,44546A,4472C4,A5A5A5,5B9BD5,264478",
    };

    /// <summary>
    /// Magazine: bold, large title, no axes, direct data labels. Storytelling style.
    /// </summary>
    private static readonly Dictionary<string, string> Magazine = new()
    {
        ["gridlines"] = "none",
        ["minorGridlines"] = "none",
        ["axisVisible"] = "false",
        ["axisLine"] = "none",
        ["majorTickMark"] = "none",
        ["title.bold"] = "true",
        ["title.size"] = "20",
        ["title.color"] = "333333",
        ["plotFill"] = "none",
        ["chartFill"] = "none",
        ["plotArea.border"] = "none",
        ["chartArea.border"] = "none",
        ["legend"] = "none",
        ["datalabels"] = "value",
        ["labelPos"] = "outsideEnd",
        ["labelfont"] = "11:555555",
        ["roundedCorners"] = "false",
        ["colors"] = "E15759,4E79A7,F28E2B,76B7B2,59A14F,EDC948",
    };

    /// <summary>
    /// Dashboard: compact, dense information, thin gridlines, small fonts.
    /// </summary>
    private static readonly Dictionary<string, string> Dashboard = new()
    {
        ["gridlines"] = "EEEEEE:0.2",
        ["minorGridlines"] = "none",
        ["axisLine"] = "CCCCCC:0.3",
        ["majorTickMark"] = "none",
        ["axisFont"] = "8:999999",
        ["legendFont"] = "8:999999",
        ["legend"] = "bottom",
        ["title.size"] = "11",
        ["title.bold"] = "true",
        ["title.color"] = "555555",
        ["plotFill"] = "none",
        ["chartFill"] = "none",
        ["plotArea.border"] = "none",
        ["chartArea.border"] = "E0E0E0:0.3",
        ["roundedCorners"] = "true",
        ["gapWidth"] = "80",
        ["colors"] = "4CAF50,2196F3,FF9800,9C27B0,F44336,00BCD4",
    };

    /// <summary>
    /// Colorful: vibrant, saturated colors with moderate styling. Fun and engaging.
    /// </summary>
    private static readonly Dictionary<string, string> Colorful = new()
    {
        ["gridlines"] = "E8E8E8:0.3",
        ["minorGridlines"] = "none",
        ["axisLine"] = "none",
        ["majorTickMark"] = "none",
        ["axisFont"] = "10:666666",
        ["legendFont"] = "10:444444",
        ["legend"] = "bottom",
        ["plotFill"] = "none",
        ["chartFill"] = "none",
        ["plotArea.border"] = "none",
        ["chartArea.border"] = "none",
        ["roundedCorners"] = "true",
        ["colors"] = "FF6384,36A2EB,FFCE56,4BC0C0,9966FF,FF9F40,C9CBCF,7BC8A4",
    };

    /// <summary>
    /// Monochrome: single-hue progression, elegant and accessible.
    /// </summary>
    private static readonly Dictionary<string, string> Monochrome = new()
    {
        ["gridlines"] = "E0E0E0:0.3",
        ["minorGridlines"] = "none",
        ["axisLine"] = "999999:0.4",
        ["majorTickMark"] = "out",
        ["axisFont"] = "9:666666",
        ["legendFont"] = "9:666666",
        ["legend"] = "bottom",
        ["plotFill"] = "none",
        ["chartFill"] = "none",
        ["plotArea.border"] = "none",
        ["chartArea.border"] = "none",
        ["roundedCorners"] = "false",
        ["colors"] = "1A3A5C,2E6B8A,4A9BBF,7BC0E0,B0D9EF,D6EBF5",
    };
}
