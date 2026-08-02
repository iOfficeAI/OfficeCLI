// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Helpers for shortening strings that are about to be PRINTED — table
/// columns, previews, error echoes. Not for anything that goes back into a
/// document.
/// </summary>
public static class DisplayText
{
    /// <summary>
    /// Cut <paramref name="s"/> to <paramref name="cutAt"/> UTF-16 units and
    /// append <paramref name="ellipsis"/>, never splitting a surrogate pair —
    /// a cut landing between the halves of an astral character (emoji, CJK
    /// ext-B) leaves a lone surrogate that prints as U+FFFD. Strings at or
    /// under the limit are returned untouched.
    /// </summary>
    public static string Truncate(string? s, int cutAt, string ellipsis = "…")
    {
        s ??= string.Empty;
        if (cutAt < 0 || s.Length <= cutAt) return s;
        var cut = cutAt;
        if (cut > 0 && char.IsHighSurrogate(s[cut - 1])) cut--;
        return s[..cut] + ellipsis;
    }
}
