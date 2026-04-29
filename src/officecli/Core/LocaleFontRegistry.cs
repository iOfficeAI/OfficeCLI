// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Locale → default font mapping for fresh blank documents. Mirrors the
/// data-driven approach LibreOffice uses (VCL.xcu): given a locale tag, pick
/// reasonable defaults for the Latin / EastAsian / ComplexScript font slots.
///
/// We deliberately keep this small (one line per locale family) rather than
/// trying to model every Office localization. When no locale is supplied,
/// returning all-empty values lets the host application substitute its own
/// UI-locale defaults — that's the POI-aligned behaviour BlankDocCreator
/// already had after we removed the "宋体" hardcode.
///
/// Font names are chosen for cross-platform availability (typefaces shipped
/// by Microsoft Office on Windows and macOS, plus Apple Sans equivalents).
/// </summary>
public static class LocaleFontRegistry
{
    /// <summary>
    /// Resolve a locale tag (e.g. "zh-CN", "ja", "ar-SA") to a per-script
    /// font triple. Returns (null, null, null) when no locale is supplied
    /// or the tag is unknown — callers should treat that as "leave the
    /// docDefaults blank, let the host application decide".
    /// </summary>
    public static (string? Latin, string? EastAsia, string? ComplexScript) Resolve(string? locale)
    {
        if (string.IsNullOrWhiteSpace(locale)) return (null, null, null);

        // Match on language-only first; full tag lookups (e.g. zh-Hant) are
        // routed through the language-only entry unless a region-specific
        // variant exists.
        var lower = locale.Replace('_', '-').ToLowerInvariant();
        var lang = lower.Split('-')[0];

        // Fully-tagged regional variants take precedence.
        switch (lower)
        {
            case "zh-tw" or "zh-hk" or "zh-mo" or "zh-hant":
                return ("Times New Roman", "新細明體", null);
            case "zh-cn" or "zh-sg" or "zh-hans":
                return ("Times New Roman", "等线", null);
        }

        // Language-only fall-throughs.
        return lang switch
        {
            "zh" => ("Times New Roman", "等线", null),
            "ja" => ("Times New Roman", "游明朝", null),
            "ko" => ("Times New Roman", "맑은 고딕", null),
            "ar" => ("Times New Roman", null, "Arabic Typesetting"),
            "he" => ("Times New Roman", null, "Times New Roman"),
            "th" => ("Times New Roman", null, "Tahoma"),
            "fa" => ("Times New Roman", null, "B Nazanin"),
            "ur" => ("Times New Roman", null, "Jameel Noori Nastaleeq"),
            "hi" => ("Times New Roman", null, "Mangal"),
            "en" or "fr" or "de" or "es" or "it" or "pt" or "nl" or "ru" or "pl"
                => ("Times New Roman", null, null),
            _ => (null, null, null)
        };
    }

    /// <summary>
    /// Returns a CSS font-family fallback fragment for the locale's CJK script,
    /// used by HTML/SVG renderers when the document's declared font isn't
    /// installed on the rendering machine.
    ///
    /// The returned fragment is comma-separated, individually quoted, NOT
    /// prefixed with a comma — callers concatenate as needed. Empty string
    /// for unknown/unspecified locales: callers should fall through to a
    /// neutral generic family (e.g. <c>sans-serif</c>) so the rendering OS
    /// picks a reasonable default rather than forcing one script's glyphs.
    /// </summary>
    public static string GetCjkCssFallback(string? locale)
    {
        if (string.IsNullOrWhiteSpace(locale)) return "";
        var lang = locale.Replace('_', '-').ToLowerInvariant().Split('-')[0];
        return lang switch
        {
            "zh" => "'PingFang SC', 'Microsoft YaHei', 'Noto Sans CJK SC', 'Hiragino Sans GB', 'Songti SC', 'STSong'",
            "ja" => "'Hiragino Sans', 'Hiragino Mincho ProN', 'Yu Gothic', 'Yu Mincho', 'Noto Sans CJK JP', 'MS Gothic'",
            "ko" => "'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans CJK KR', 'Batang'",
            _ => ""
        };
    }

    /// <summary>
    /// Heuristic: detect a CJK locale tag ("zh" / "ja" / "ko") from a font
    /// typeface name. Returns null when the name carries no strong script
    /// signal. Used by renderers to pick the right fallback chain when the
    /// document doesn't declare an explicit eastAsia language tag.
    ///
    /// Order matters: Japanese is checked before Chinese because some JP
    /// font names contain hanzi that overlap with Chinese keywords.
    /// </summary>
    public static string? DetectLocaleFromCjkFontName(string? font)
    {
        if (string.IsNullOrEmpty(font)) return null;
        var lower = font.ToLowerInvariant();

        if (lower.Contains("明朝") || lower.Contains("mincho")
            || lower.Contains("ゴシック") || lower.Contains("hiragino")
            || lower.Contains("yu mincho") || lower.Contains("yu gothic")
            || lower.Contains("ms mincho") || lower.Contains("ms gothic")
            || lower.Contains("meiryo") || lower.Contains("游明朝")
            || lower.Contains("游ゴシック"))
            return "ja";

        if (lower.Contains("바탕") || lower.Contains("굴림") || lower.Contains("돋움")
            || lower.Contains("맑은") || lower == "batang" || lower == "batangche"
            || lower == "gulim" || lower == "dotum" || lower.Contains("malgun")
            || lower.Contains("nanum") || lower.Contains("apple sd gothic"))
            return "ko";

        if (lower.Contains("宋") || lower.Contains("song") || lower.Contains("simsun")
            || lower.Contains("黑") || lower.Contains("hei") || lower.Contains("simhei")
            || lower.Contains("楷") || lower.Contains("kai") || lower.Contains("仿宋")
            || lower.Contains("fangsong") || lower.Contains("pingfang")
            || lower.Contains("yahei") || lower.Contains("等线") || lower.Contains("华文")
            || lower.Contains("方正") || lower.Contains("微软雅黑"))
            return "zh";

        return null;
    }
}
