// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Single source of truth for Word default page geometry (twips).
/// Used as fallback when a section's pgSz/pgMar is missing — callers
/// must always read the source <c>SectionProperties</c> first and only
/// drop to these defaults when the value is genuinely absent.
/// </summary>
public static class WordPageDefaults
{
    // A4: 210mm × 297mm at 1440 twips/inch (= 567 twips/cm).
    public const uint A4WidthTwips = 11906;
    public const uint A4HeightTwips = 16838;
}
