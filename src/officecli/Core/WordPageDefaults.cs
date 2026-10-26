// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    internal static DocumentFormat.OpenXml.Wordprocessing.PageMargin CreatePageMargin() => new()
    {
        Top = 1440,
        Bottom = 1440,
        Left = 1800U,
        Right = 1800U,
        Header = 720U,
        Footer = 720U,
        Gutter = 0U
    };

    // A4: 210mm × 297mm at 1440 twips/inch (= 567 twips/cm).
    public const uint A4WidthTwips = 11906;
    public const uint A4HeightTwips = 16838;

    // OOXML legal range for w:pgSz/@w:w and @w:h. Word's UI clamps roughly to
    // ~0.4cm–55.9cm; the EcmaSpec defines 1..31680 (22"). Use 240 (1/6") as the
    // lower bound — anything smaller will not produce a renderable page in Word.
    public const uint PageDimMinTwips = 240;
    public const uint PageDimMaxTwips = 31680;

    public static void ValidatePageDim(long twips, string keyName)
    {
        if (twips < PageDimMinTwips || twips > PageDimMaxTwips)
            throw new ArgumentException(
                $"{keyName} must be in range {PageDimMinTwips}–{PageDimMaxTwips} twips " +
                $"(~0.4cm–55.9cm), got {twips}.");
    }
}
