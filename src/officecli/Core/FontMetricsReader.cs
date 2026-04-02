// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Lightweight TTF/TTC font metrics reader. Reads only the OS/2 and head tables
/// to extract usWinAscent, usWinDescent, and unitsPerEm — enough to calculate
/// the line-height ratio that Word uses for line spacing.
///
/// Word line spacing formula: actualLineHeight = (winAscent + winDescent) / UPM × fontSize × multiplier
/// CSS line-height is relative to fontSize, so: cssLineHeight = wordMultiplier × ratio
/// where ratio = (winAscent + winDescent) / UPM
/// </summary>
internal static class FontMetricsReader
{
    /// <summary>
    /// Line-height ratio = (usWinAscent + usWinDescent) / unitsPerEm.
    /// Returns 1.0 if the font file cannot be read.
    /// </summary>
    public static double GetLineHeightRatio(string fontFilePath, int fontIndex = 0)
    {
        try
        {
            using var fs = File.OpenRead(fontFilePath);
            using var reader = new BinaryReader(fs);
            var offset = GetFontOffset(reader, fontIndex);
            if (offset < 0) return 1.0;

            var (headOffset, os2Offset, hheaOffset) = FindTables(reader, offset);
            if (headOffset < 0 || os2Offset < 0) return 1.0;

            // head table: unitsPerEm at offset 18 (uint16)
            fs.Position = headOffset + 18;
            var upm = ReadUInt16BE(reader);
            if (upm == 0) return 1.0;

            // OS/2 table: usWinAscent at offset 74, usWinDescent at offset 76 (uint16)
            fs.Position = os2Offset + 74;
            var winAscent = ReadUInt16BE(reader);
            var winDescent = ReadUInt16BE(reader);
            var winTotal = (double)(winAscent + winDescent);

            // hhea table: ascent at offset 4, descent at offset 6 (int16), lineGap at offset 8 (int16)
            // Chrome uses: max(winAscent+winDescent, hheaAscent+|hheaDescent|+hheaLineGap) / UPM
            double hheaTotal = 0;
            if (hheaOffset >= 0)
            {
                fs.Position = hheaOffset + 4;
                var hheaAscent = ReadInt16BE(reader);
                var hheaDescent = ReadInt16BE(reader);
                var hheaLineGap = ReadInt16BE(reader);
                hheaTotal = hheaAscent + Math.Abs(hheaDescent) + Math.Max(0, (int)hheaLineGap);
            }

            return Math.Max(winTotal, hheaTotal) / upm;
        }
        catch
        {
            return 1.0;
        }
    }

    /// <summary>Get the offset to the font directory for TTC collections or plain TTF.</summary>
    private static long GetFontOffset(BinaryReader reader, int fontIndex)
    {
        reader.BaseStream.Position = 0;
        var tag = ReadUInt32BE(reader);

        // TTC header: "ttcf"
        if (tag == 0x74746366)
        {
            reader.BaseStream.Position = 8; // skip version
            var numFonts = (int)ReadUInt32BE(reader);
            if (fontIndex >= numFonts) return -1;
            reader.BaseStream.Position = 12 + fontIndex * 4;
            return ReadUInt32BE(reader);
        }

        // Plain TTF/OTF: tag is 0x00010000 or "OTTO"
        return 0;
    }

    /// <summary>Find head, OS/2, and hhea table offsets from the font directory.</summary>
    private static (long headOffset, long os2Offset, long hheaOffset) FindTables(BinaryReader reader, long fontOffset)
    {
        reader.BaseStream.Position = fontOffset + 4; // skip sfVersion
        var numTables = ReadUInt16BE(reader);
        reader.BaseStream.Position = fontOffset + 12; // skip searchRange, entrySelector, rangeShift

        long headOffset = -1, os2Offset = -1, hheaOffset = -1;
        for (int i = 0; i < numTables; i++)
        {
            var tableTag = ReadUInt32BE(reader);
            reader.BaseStream.Position += 4; // skip checksum
            var tableOffset = (long)ReadUInt32BE(reader);
            reader.BaseStream.Position += 4; // skip length

            if (tableTag == 0x68656164) // "head"
                headOffset = tableOffset;
            else if (tableTag == 0x4F532F32) // "OS/2"
                os2Offset = tableOffset;
            else if (tableTag == 0x68686561) // "hhea"
                hheaOffset = tableOffset;

            if (headOffset >= 0 && os2Offset >= 0 && hheaOffset >= 0) break;
        }
        return (headOffset, os2Offset, hheaOffset);
    }

    private static ushort ReadUInt16BE(BinaryReader r)
    {
        var b = r.ReadBytes(2);
        return (ushort)((b[0] << 8) | b[1]);
    }

    private static short ReadInt16BE(BinaryReader r)
    {
        var b = r.ReadBytes(2);
        return (short)((b[0] << 8) | b[1]);
    }

    private static uint ReadUInt32BE(BinaryReader r)
    {
        var b = r.ReadBytes(4);
        return (uint)((b[0] << 24) | (b[1] << 16) | (b[2] << 8) | b[3]);
    }

    // ==================== Font Discovery ====================

    /// <summary>
    /// Search common font directories for a font file matching the given family name.
    /// Returns the file path, or null if not found.
    /// </summary>
    public static string? FindFontFile(string fontFamily)
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var dirs = new List<string>();
        if (OperatingSystem.IsMacOS())
        {
            dirs.Add(Path.Combine(home, "Library/Fonts"));
            dirs.Add("/Library/Fonts");
            dirs.Add("/System/Library/Fonts");
            dirs.Add("/System/Library/Fonts/Supplemental");
            var officeFonts = "/Applications/Microsoft Word.app/Contents/Resources/DFonts";
            if (Directory.Exists(officeFonts)) dirs.Add(officeFonts);
        }
        else if (OperatingSystem.IsWindows())
        {
            dirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));
            dirs.Add(Path.Combine(home, @"AppData\Local\Microsoft\Windows\Fonts"));
        }
        else
        {
            dirs.Add(Path.Combine(home, ".fonts"));
            dirs.Add("/usr/share/fonts");
            dirs.Add("/usr/local/share/fonts");
        }

        // Normalize: remove spaces, try exact match and lowercase
        var normalized = fontFamily.Replace(" ", "");
        foreach (var dir in dirs)
        {
            if (!Directory.Exists(dir)) continue;
            foreach (var file in Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories))
            {
                var ext = Path.GetExtension(file);
                if (ext is not (".ttf" or ".otf" or ".ttc")) continue;
                var stem = Path.GetFileNameWithoutExtension(file);
                if (stem.Equals(normalized, StringComparison.OrdinalIgnoreCase)
                    || stem.Equals(fontFamily, StringComparison.OrdinalIgnoreCase))
                    return file;
            }
        }
        return null;
    }

    // ==================== Cached Ratio Lookup ====================

    private static readonly Dictionary<string, double> s_ratioCache = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Get the line-height ratio for a font family name. Cached after first lookup.
    /// Returns 1.0 if the font cannot be found or read.
    /// </summary>
    public static double GetRatio(string fontFamily)
    {
        if (s_ratioCache.TryGetValue(fontFamily, out var cached))
            return cached;

        var path = FindFontFile(fontFamily);
        var ratio = path != null ? GetLineHeightRatio(path) : 1.0;
        s_ratioCache[fontFamily] = ratio;
        return ratio;
    }

    // ==================== Ascent/Descent Override ====================

    /// <summary>
    /// Get CSS ascent-override and descent-override percentages for a font.
    /// These tell the browser to distribute line-height space according to
    /// the font's ascent/descent ratio (matching Word's behavior) instead of
    /// the CSS default half-leading model.
    /// Returns (0, 0) if the font cannot be found.
    /// </summary>
    public static (double ascentPct, double descentPct) GetAscentDescentOverride(string fontFamily)
    {
        var path = FindFontFile(fontFamily);
        if (path == null) return (0, 0);

        try
        {
            using var fs = File.OpenRead(path);
            using var reader = new BinaryReader(fs);
            var offset = GetFontOffset(reader, 0);
            if (offset < 0) return (0, 0);

            var (headOffset, os2Offset, _) = FindTables(reader, offset);
            if (headOffset < 0 || os2Offset < 0) return (0, 0);

            fs.Position = headOffset + 18;
            var upm = ReadUInt16BE(reader);
            if (upm == 0) return (0, 0);

            fs.Position = os2Offset + 74;
            var winAscent = ReadUInt16BE(reader);
            var winDescent = ReadUInt16BE(reader);

            return (winAscent * 100.0 / upm, winDescent * 100.0 / upm);
        }
        catch
        {
            return (0, 0);
        }
    }
}
