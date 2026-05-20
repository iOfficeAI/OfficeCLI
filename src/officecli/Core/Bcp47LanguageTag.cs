// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Shared RFC 5646 / BCP-47 language-tag syntax check used by every handler
/// that writes a `<*:lang>` / `<*:altLang>` attribute. Validates the shape
/// only — does not look up the tag against the IANA subtag registry. That
/// would require shipping the registry and gating new languages on each
/// release; Office itself never refuses a syntactically valid tag, so a
/// pure shape check matches recipient behavior.
///
/// Cross-handler: docx run lang.val/lang.eastAsia/lang.bidi
/// (WordHandler.Helpers.cs) and pptx run/shape lang/altLang
/// (PowerPointHandler.ShapeProperties.cs) both delegate here. Splitting
/// the two regexes would create a "valid in Word, invalid in PowerPoint"
/// drift on the same value — never a useful distinction.
/// </summary>
public static class Bcp47LanguageTag
{
    public const int MaxLength = 35;

    // Shape: primary subtag 2-3 letters with optional hyphenated subtags;
    // 4-8 letter primary requires at least one subtag (reserved/grandfathered
    // range); `x-…` private-use form. Subtags are 1-8 alphanumerics.
    // R18-fuzz-3: prior `^[A-Za-z][A-Za-z0-9-]*$` form let "INVALID" and
    // 1000-char garbage through; this shape rejects both.
    private static readonly Regex Shape = new(
        @"^(?:[A-Za-z]{2,3}(?:-[A-Za-z0-9]{1,8})*|[A-Za-z]{4,8}(?:-[A-Za-z0-9]{1,8})+|x(?:-[A-Za-z0-9]{1,8})+)$",
        RegexOptions.Compiled);

    /// <summary>
    /// True when the value has the BCP-47 shape (or is empty — empty means
    /// "clear the slot", which callers handle separately).
    /// </summary>
    public static bool IsValid(string? value)
    {
        if (string.IsNullOrEmpty(value)) return true;
        return value.Length <= MaxLength && Shape.IsMatch(value);
    }
}
