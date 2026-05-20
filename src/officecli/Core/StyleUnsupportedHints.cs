// Copyright 2025 OfficeCLI (officecli.ai)
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
        // firstLineChars / leftChars / rightChars / hangingChars are now wired
        // on /styles (P1-6) — symmetric with firstLineIndent / leftIndent /
        // rightIndent / hangingIndent (BT-5).
        // spaceBeforeLines / spaceAfterLines are now wired on /styles (P1-7) —
        // also on paragraphs alongside the `spaceBefore=Nlines` suffix.
        // shading.* / underline.color now flow through TypedAttributeFallback
        // (via the shading→shd and underline→u aliases) on /styles, paragraph,
        // run, and cell paths — entries removed once verified to write
        // schema-valid <w:shd>/<w:u> XML. `tabs` removed once the curated
        // POS:ALIGN[:LEADER] parser landed.
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
