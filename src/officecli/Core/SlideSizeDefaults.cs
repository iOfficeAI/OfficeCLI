// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Presentation;

namespace OfficeCli.Core;

/// <summary>
/// Single source of truth for PowerPoint slide-size presets (EMU).
/// Used as fallback when <c>presentation.xml/sldSz</c> is missing, and
/// as the canonical preset table behind <c>set --prop slidesize=…</c>.
/// </summary>
public static class SlideSizeDefaults
{
    // Office default (also what PowerPoint applies to a brand-new deck).
    public const long Widescreen16x9Cx = 12192000;
    public const long Widescreen16x9Cy = 6858000;

    // Default notes page (portrait, letter-ish).
    public const long NotesPortraitCx = 6858000;
    public const long NotesPortraitCy = 9144000;

    public readonly record struct Preset(long Cx, long Cy, SlideSizeValues Type);

    /// <summary>
    /// Maps the user-facing preset names accepted by <c>set --prop slidesize=…</c>
    /// to the EMU dimensions and matching <c>SlideSizeValues</c> enum.
    /// Lookup is case-insensitive; aliases share an entry.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, Preset> Presets = new Dictionary<string, Preset>(StringComparer.OrdinalIgnoreCase)
    {
        ["16:9"]       = new(Widescreen16x9Cx, Widescreen16x9Cy, SlideSizeValues.Screen16x9),
        ["widescreen"] = new(Widescreen16x9Cx, Widescreen16x9Cy, SlideSizeValues.Screen16x9),
        ["4:3"]        = new(9144000,  6858000,  SlideSizeValues.Screen4x3),
        ["standard"]   = new(9144000,  6858000,  SlideSizeValues.Screen4x3),
        ["16:10"]      = new(12192000, 7620000,  SlideSizeValues.Screen16x10),
        ["a4"]         = new(10692000, 7560000,  SlideSizeValues.A4),
        ["a3"]         = new(15120000, 10692000, SlideSizeValues.A3),
        // Letter = 8.5" × 11" (landscape on slide canvas: 11" × 8.5").
        // 1in = 914400 EMU → 10058400 × 7772400.
        ["letter"]     = new(10058400, 7772400,  SlideSizeValues.Letter),
        ["b4"]         = new(11430000, 8574000,  SlideSizeValues.B4ISO),
        ["b5"]         = new(8208000,  5760000,  SlideSizeValues.B5ISO),
        ["35mm"]       = new(10287000, 6858000,  SlideSizeValues.Film35mm),
        ["overhead"]   = new(9144000,  6858000,  SlideSizeValues.Overhead),
        ["banner"]     = new(7315200,  914400,   SlideSizeValues.Banner),
        ["ledger"]     = new(12192000, 9144000,  SlideSizeValues.Ledger),
    };
}
