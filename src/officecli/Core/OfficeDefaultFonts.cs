// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Single source of truth for the Microsoft "Office" default font scheme.
/// These literals appear in two contexts:
///
///   1. Blank document creation — emitted into theme1.xml's fontScheme.
///   2. Preview rendering fallback — when a document lacks any explicit
///      font (no run rPr, no styles.xml docDefaults, no theme part) the
///      HTML preview defaults to these values rather than the browser's
///      generic serif/sans default.
///
/// Note: when a document HAS a theme part, callers should prefer reading
/// <c>theme.fontScheme.MinorFont.LatinFont.Typeface</c> (or MajorFont
/// for headings) before falling back to these constants. The constants
/// are the *last* resort, not the first.
/// </summary>
public static class OfficeDefaultFonts
{
    public const string MajorLatin = "Calibri Light";
    public const string MinorLatin = "Calibri";

    /// <summary>Excel default body font size (pt) when stylesheet Font[0] is missing.</summary>
    public const string ExcelBodySizePt = "11";
}
