// Copyright 2026 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Validates user-supplied text against the XML 1.0 character production
/// (https://www.w3.org/TR/xml/#charsets). Open XML files are XML 1.0, so
/// codepoints outside this set cannot be persisted without raising
/// XmlException at save time -- which leaks as `internal_error` and poisons
/// the in-memory document for subsequent writes against a resident handler.
///
/// Call <see cref="ValidateOrThrow"/> at every Add/Set site that accepts a
/// free-form text string BEFORE the value reaches the Open XML SDK. The
/// resulting <see cref="CliException"/> carries `invalid_value` so the CLI
/// envelope surfaces a deterministic, user-fixable error rather than a raw
/// XML parser leak.
///
/// Valid XML 1.0 character range:
///   U+0009 (TAB), U+000A (LF), U+000D (CR),
///   U+0020..U+D7FF,
///   U+E000..U+FFFD,
///   U+10000..U+10FFFF (encoded as surrogate pairs in UTF-16).
/// </summary>
internal static class XmlTextValidator
{
    /// <summary>
    /// Returns null when <paramref name="text"/> contains only XML 1.0
    /// character-data codepoints; otherwise returns a human-readable
    /// description of the first offending codepoint (e.g.
    /// "U+0001 at offset 3"). Null/empty input is treated as valid.
    /// </summary>
    public static string? FindInvalidChar(string? text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];

            // Surrogate pair: consume the low half if well-formed.
            if (char.IsHighSurrogate(c))
            {
                if (i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    int cp = char.ConvertToUtf32(c, text[i + 1]);
                    if (cp is < 0x10000 or > 0x10FFFF)
                        return FormatOffender(cp, i);
                    i++; // skip the low surrogate
                    continue;
                }
                return FormatOffender(c, i); // unpaired high surrogate
            }
            if (char.IsLowSurrogate(c))
                return FormatOffender(c, i); // unpaired low surrogate

            // BMP codepoint.
            if (c == 0x9 || c == 0xA || c == 0xD) continue;
            if (c >= 0x20 && c <= 0xD7FF) continue;
            if (c >= 0xE000 && c <= 0xFFFD) continue;
            return FormatOffender(c, i);
        }
        return null;
    }

    /// <summary>
    /// Throws <see cref="CliException"/> with code `invalid_value` when
    /// <paramref name="text"/> contains a character forbidden in XML 1.0
    /// character data. <paramref name="fieldName"/> appears in the error
    /// message to help callers identify which input was rejected.
    /// </summary>
    public static void ValidateOrThrow(string? text, string fieldName)
    {
        var problem = FindInvalidChar(text);
        if (problem is null) return;
        throw new CliException(
            $"{fieldName}: contains character invalid in XML 1.0 text ({problem}). " +
            "Allowed: U+0009 (tab), U+000A (newline), U+000D (carriage return), and U+0020 and above " +
            "(excluding surrogate code points and U+FFFE/U+FFFF).")
        {
            Code = "invalid_value",
        };
    }

    private static string FormatOffender(int cp, int offset)
        => $"U+{cp.ToString("X4", CultureInfo.InvariantCulture)} at offset {offset}";
}
