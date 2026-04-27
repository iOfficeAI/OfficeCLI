// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Targeted hints for Word style props that the curated <c>add /styles</c> /
/// <c>set /styles/&lt;id&gt;</c> surface does not (yet) accept.
///
/// Two design rules:
///   1. Never recommend <c>raw-set</c>. It is an escape hatch, not a normal
///      user path; suggesting it lets users drift out of the curated CLI
///      vocabulary.
///   2. When a curated alternative exists, name it. When one does not,
///      say plainly that the prop is not supported — do not invent
///      workarounds.
/// </summary>
internal static class StyleUnsupportedHints
{
    private static readonly Dictionary<string, string> Hints = new(StringComparer.OrdinalIgnoreCase)
    {
        ["firstLineChars"] = "char-based indent is not supported on styles",
        ["leftChars"]      = "char-based indent is not supported on styles",
        ["rightChars"]     = "char-based indent is not supported on styles",
        ["hangingChars"]   = "char-based indent is not supported on styles",
        ["firstLineIndent"] = "indent attrs on styles are not supported; set indent at paragraph level instead",
        ["leftIndent"]      = "indent attrs on styles are not supported; set indent at paragraph level instead",
        ["rightIndent"]     = "indent attrs on styles are not supported; set indent at paragraph level instead",
        ["hangingIndent"]   = "indent attrs on styles are not supported; set indent at paragraph level instead",

        ["spaceBeforeLines"] = "line-based spacing is not supported; use 'spaceBefore=<twips|pt|cm>' instead",
        ["spaceAfterLines"]  = "line-based spacing is not supported; use 'spaceAfter=<twips|pt|cm>' instead",
        ["lineRule"]         = "use 'lineSpacing=<N>pt' (fixed) or 'lineSpacing=<N>x' (multiplier); lineRule is inferred",

        ["shading"]       = "style-level shading is not supported; set fill at paragraph or run level",
        ["shading.fill"]  = "style-level shading is not supported; set fill at paragraph or run level",
        ["shading.color"] = "style-level shading is not supported; set fill at paragraph or run level",
        ["shading.val"]   = "style-level shading is not supported; set fill at paragraph or run level",

        ["underline.color"] = "underline color is not supported; only 'underline=<single|double|...>' is",

        ["tabs"] = "tabs on styles are not supported",
    };

    /// <summary>
    /// Returns a single-line message of the form
    /// <c>UNSUPPORTED props on &lt;path&gt;: foo (use bar instead), baz (not supported)</c>.
    /// Empty input returns null. <paramref name="scope"/> labels the surface
    /// in the message ("/styles", "/body/p[…]", etc.) so the user knows
    /// where the rejection happened; pass null for a generic phrasing.
    /// </summary>
    public static string? Format(IEnumerable<string> unsupported, string? scope = null)
    {
        var list = unsupported.Where(p => !string.IsNullOrEmpty(p)).Distinct().ToList();
        if (list.Count == 0) return null;

        var parts = list.Select(prop =>
            Hints.TryGetValue(prop, out var hint)
                ? $"{prop} ({hint})"
                : $"{prop} (not supported)");

        var label = string.IsNullOrEmpty(scope) ? "props" : $"props on {scope}";
        return $"UNSUPPORTED {label}: {string.Join(", ", parts)}";
    }
}
