// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Lightweight TTF/TTC font reader. Extracts the per-font line-height ratio
/// used to size CSS line boxes for paragraph rendering.
///
/// Latin: ratio = (hhea.ascender + |hhea.descender| + hhea.lineGap) / unitsPerEm
/// CJK:   ratio = (asc + dsc + 2 × v7) / UPM
///        v7 = (15 × (asc + dsc) + 50) / 100   [design units]
/// </summary>
internal static class FontMetricsReader
{
    /// <summary>
    /// Line-height ratio for a font file. Returns 1.0 on any read failure.
    /// </summary>
    public static double GetLineHeightRatio(string fontFilePath, int fontIndex = 0)
    {
        try
        {
            using var fs = File.OpenRead(fontFilePath);
            using var reader = new BinaryReader(fs);
            var offset = GetFontOffset(reader, fontIndex);
            if (offset < 0) return 1.0;

            var tables = FindTables(reader, offset);
            if (tables.head < 0 || tables.hhea < 0) return 1.0;

            fs.Position = tables.head + 18;
            var upm = ReadUInt16BE(reader);
            if (upm == 0) return 1.0;

            if (tables.os2 >= 0 && TryReadOs2(reader, tables.os2, out var os2) && os2.IsCjk)
            {
                int asc = os2.UseTypo ? os2.TypoAscent : os2.WinAscent;
                int dsc = os2.UseTypo ? -os2.TypoDescent : os2.WinDescent;
                int v7 = (15 * (asc + dsc) + 50) / 100;
                return (double)(asc + dsc + 2 * v7) / upm;
            }

            fs.Position = tables.hhea + 4;
            var ascender = ReadInt16BE(reader);
            var descender = ReadInt16BE(reader);
            var lineGap = ReadInt16BE(reader);
            int total = ascender + Math.Abs((int)descender) + Math.Max(0, (int)lineGap);
            return (double)total / upm;
        }
        catch
        {
            return 1.0;
        }
    }

    private struct Os2Metrics
    {
        public int WinAscent;
        public int WinDescent;
        public int TypoAscent;
        public int TypoDescent;
        public int TypoLineGap;
        public bool UseTypo;
        public bool IsCjk;
    }

    /// <summary>
    /// CJK detection via OS/2 ulCodePageRange1 bits 17-21:
    /// 17 = JIS Japan, 18 = GB2312 PRC, 19 = Korean Wansung,
    /// 20 = Big5 Taiwan, 21 = Korean Johab.
    /// </summary>
    private static bool TryReadOs2(BinaryReader r, long os2Offset, out Os2Metrics m)
    {
        m = default;
        try
        {
            r.BaseStream.Position = os2Offset;
            ushort version = ReadUInt16BE(r);
            r.BaseStream.Position = os2Offset + 62;
            ushort fsSelection = ReadUInt16BE(r);
            m.UseTypo = (fsSelection & 0x80) != 0;

            r.BaseStream.Position = os2Offset + 68;
            m.TypoAscent = ReadInt16BE(r);
            m.TypoDescent = ReadInt16BE(r);
            m.TypoLineGap = ReadInt16BE(r);
            m.WinAscent = ReadUInt16BE(r);
            m.WinDescent = ReadUInt16BE(r);

            if (version >= 1)
            {
                r.BaseStream.Position = os2Offset + 78;
                uint cp1 = ReadUInt32BE(r);
                const uint cjkMask = (1U << 17) | (1U << 18) | (1U << 19) | (1U << 20) | (1U << 21);
                m.IsCjk = (cp1 & cjkMask) != 0;
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static long GetFontOffset(BinaryReader reader, int fontIndex)
    {
        reader.BaseStream.Position = 0;
        var tag = ReadUInt32BE(reader);

        // TTC collection header
        if (tag == 0x74746366)
        {
            reader.BaseStream.Position = 8;
            var numFonts = (int)ReadUInt32BE(reader);
            if (fontIndex >= numFonts) return -1;
            reader.BaseStream.Position = 12 + fontIndex * 4;
            return ReadUInt32BE(reader);
        }
        return 0;
    }

    private struct TableOffsets
    {
        public long head;
        public long os2;
        public long hhea;
        public long name;
        public long cmap;
    }

    private static TableOffsets FindTables(BinaryReader reader, long fontOffset)
    {
        reader.BaseStream.Position = fontOffset + 4;
        var numTables = ReadUInt16BE(reader);
        reader.BaseStream.Position = fontOffset + 12;

        var t = new TableOffsets { head = -1, os2 = -1, hhea = -1, name = -1, cmap = -1 };
        for (int i = 0; i < numTables; i++)
        {
            var tag = ReadUInt32BE(reader);
            reader.BaseStream.Position += 4;
            var off = (long)ReadUInt32BE(reader);
            reader.BaseStream.Position += 4;

            if (tag == 0x68656164) t.head = off;
            else if (tag == 0x4F532F32) t.os2 = off;
            else if (tag == 0x68686561) t.hhea = off;
            else if (tag == 0x6E616D65) t.name = off;
            else if (tag == 0x636D6170) t.cmap = off;

            if (t.head >= 0 && t.os2 >= 0 && t.hhea >= 0 && t.name >= 0 && t.cmap >= 0) break;
        }
        return t;
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

    // ==================== Font lookup ====================

    private static List<string> GetFontDirs()
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
        return dirs;
    }

    /// <summary>Family-name → (file path, font collection index). Built lazily.</summary>
    private static Dictionary<string, (string path, int idx)>? s_familyIndex;
    private static readonly object s_familyIndexLock = new();

    private static Dictionary<string, (string path, int idx)> BuildFamilyIndex()
    {
        var map = new Dictionary<string, (string, int)>(StringComparer.OrdinalIgnoreCase);
        foreach (var dir in GetFontDirs())
        {
            if (!Directory.Exists(dir)) continue;
            IEnumerable<string> files;
            try { files = Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories); }
            catch { continue; }
            foreach (var file in files)
            {
                var ext = Path.GetExtension(file);
                if (ext is not (".ttf" or ".otf" or ".ttc")) continue;
                try
                {
                    foreach (var (faceIdx, family) in EnumerateFaceFamilies(file))
                    {
                        map.TryAdd(family, (file, faceIdx));
                        map.TryAdd(family.Replace(" ", ""), (file, faceIdx));
                    }
                }
                catch
                {
                    // ignore unreadable file; fall through to stem-based fallback
                }
                // stem fallback for fast lookup of common cases
                var stem = Path.GetFileNameWithoutExtension(file);
                if (!string.IsNullOrEmpty(stem))
                    map.TryAdd(stem, (file, 0));
            }
        }
        return map;
    }

    private static Dictionary<string, (string path, int idx)> GetFamilyIndex()
    {
        if (s_familyIndex != null) return s_familyIndex;
        lock (s_familyIndexLock)
        {
            s_familyIndex ??= BuildFamilyIndex();
            return s_familyIndex;
        }
    }

    private static IEnumerable<(int faceIndex, string family)> EnumerateFaceFamilies(string path)
    {
        using var fs = File.OpenRead(path);
        using var reader = new BinaryReader(fs);
        fs.Position = 0;
        var tag = ReadUInt32BE(reader);
        int faceCount;
        long[] faceOffsets;
        if (tag == 0x74746366)
        {
            fs.Position = 8;
            faceCount = (int)ReadUInt32BE(reader);
            faceOffsets = new long[faceCount];
            fs.Position = 12;
            for (int i = 0; i < faceCount; i++)
                faceOffsets[i] = ReadUInt32BE(reader);
        }
        else
        {
            faceCount = 1;
            faceOffsets = new[] { 0L };
        }

        for (int faceIdx = 0; faceIdx < faceCount; faceIdx++)
        {
            var tables = FindTables(reader, faceOffsets[faceIdx]);
            if (tables.name < 0) continue;
            foreach (var family in ReadFamilyNames(reader, tables.name))
                yield return (faceIdx, family);
        }
    }

    private static IEnumerable<string> ReadFamilyNames(BinaryReader reader, long nameTableOffset)
    {
        var fs = reader.BaseStream;
        fs.Position = nameTableOffset;
        var format = ReadUInt16BE(reader);
        var count = ReadUInt16BE(reader);
        var stringOffset = ReadUInt16BE(reader);

        // Collect candidate (platform/lang priority, raw bytes, encoding) tuples; emit sorted.
        var records = new List<(int priority, byte[] bytes, int encoding)>();
        long recordsStart = fs.Position;

        for (int i = 0; i < count; i++)
        {
            fs.Position = recordsStart + i * 12;
            var platformId = ReadUInt16BE(reader);
            var encodingId = ReadUInt16BE(reader);
            var languageId = ReadUInt16BE(reader);
            var nameId = ReadUInt16BE(reader);
            var length = ReadUInt16BE(reader);
            var strOff = ReadUInt16BE(reader);

            // Family-name name IDs: 1 (family), 16 (preferred family), 4 (full name)
            if (nameId != 1 && nameId != 4 && nameId != 16) continue;

            // Skip languages other than English/Unicode-default
            bool isEnglish =
                (platformId == 3 && (languageId == 0x0409 || languageId == 0)) ||
                (platformId == 0) ||
                (platformId == 1 && languageId == 0);
            if (!isEnglish) continue;

            int priority =
                (nameId == 16 ? 0 : nameId == 1 ? 10 : 20) +
                (platformId == 3 && encodingId == 1 ? 0 :
                 platformId == 3 && encodingId == 10 ? 1 :
                 platformId == 0 ? 2 :
                 platformId == 1 ? 5 : 9);

            var savedPos = fs.Position;
            fs.Position = nameTableOffset + stringOffset + strOff;
            var bytes = reader.ReadBytes(length);
            fs.Position = savedPos;

            int enc =
                (platformId == 3 && (encodingId == 0 || encodingId == 1 || encodingId == 10)) ? 1 :
                (platformId == 0) ? 1 :
                0;

            records.Add((priority, bytes, enc));
        }

        records.Sort((a, b) => a.priority.CompareTo(b.priority));
        foreach (var (_, bytes, enc) in records)
        {
            string s = enc == 1
                ? System.Text.Encoding.BigEndianUnicode.GetString(bytes)
                : System.Text.Encoding.Latin1.GetString(bytes);
            s = s.Trim();
            if (s.Length > 0) yield return s;
        }
    }

    /// <summary>
    /// Look up a font by family name. Returns the file path or null if not present.
    /// </summary>
    public static string? FindFontFile(string fontFamily)
    {
        var hit = FindFont(fontFamily);
        return hit?.path;
    }

    /// <summary>
    /// Look up a font by family name, returning both the file path and the
    /// face index inside a TTC collection.
    /// </summary>
    public static (string path, int idx)? FindFont(string fontFamily)
    {
        if (string.IsNullOrEmpty(fontFamily)) return null;
        var idx = GetFamilyIndex();
        if (idx.TryGetValue(fontFamily, out var hit)) return hit;
        if (idx.TryGetValue(fontFamily.Replace(" ", ""), out hit)) return hit;
        return null;
    }

    // ==================== Cached ratio lookup ====================

    private static readonly Dictionary<string, double> s_ratioCache = new(StringComparer.OrdinalIgnoreCase);

    public static double GetRatio(string fontFamily)
    {
        if (s_ratioCache.TryGetValue(fontFamily, out var cached))
            return cached;

        var hit = FindFont(fontFamily);
        var ratio = hit.HasValue ? GetLineHeightRatio(hit.Value.path, hit.Value.idx) : 1.0;
        s_ratioCache[fontFamily] = ratio;
        return ratio;
    }

    // ==================== Ascent/Descent override ====================

    /// <summary>
    /// Return ascent/descent split (as percentage of em). CJK fonts get
    /// a +round(0.15 × (asc+dsc)) padding on each side. Latin fonts take
    /// the larger of the two ascent/descent pairs available in the font;
    /// the line-gap field, when present, folds into the ascent side.
    /// Returns (0, 0) when the font isn't locatable.
    /// </summary>
    public static (double ascentPctEm, double descentPctEm) GetSplitAscDscOverride(string fontFamily)
    {
        var hit = FindFont(fontFamily);
        if (!hit.HasValue) return (0, 0);

        try
        {
            using var fs = File.OpenRead(hit.Value.path);
            using var reader = new BinaryReader(fs);
            var offset = GetFontOffset(reader, hit.Value.idx);
            if (offset < 0) return (0, 0);

            var tables = FindTables(reader, offset);
            if (tables.head < 0 || tables.hhea < 0) return (0, 0);

            fs.Position = tables.head + 18;
            var upm = ReadUInt16BE(reader);
            if (upm == 0) return (0, 0);

            Os2Metrics os2 = default;
            bool haveOs2 = tables.os2 >= 0 && TryReadOs2(reader, tables.os2, out os2);

            if (haveOs2 && os2.IsCjk)
            {
                int asc = os2.UseTypo ? os2.TypoAscent : os2.WinAscent;
                int dsc = os2.UseTypo ? -os2.TypoDescent : os2.WinDescent;
                int v7 = (15 * (asc + dsc) + 50) / 100;
                return ((asc + v7) * 100.0 / upm, (dsc + v7) * 100.0 / upm);
            }

            fs.Position = tables.hhea + 4;
            var fallbackAsc = ReadInt16BE(reader);
            var fallbackDsc = ReadInt16BE(reader);
            var lineGap = ReadInt16BE(reader);
            int fallbackTotal = fallbackAsc + Math.Abs((int)fallbackDsc) + Math.Max(0, (int)lineGap);

            // Latin split: descent = primary descent; total = larger of
            // the two pairs; ascent = total − descent. A line-gap, when
            // present, folds into the ascent side.
            int primaryAsc = haveOs2 && os2.WinAscent > 0 ? os2.WinAscent : fallbackAsc;
            int primaryDsc = haveOs2 && os2.WinDescent > 0 ? os2.WinDescent : Math.Abs((int)fallbackDsc);
            int total = Math.Max(primaryAsc + primaryDsc, fallbackTotal);
            int ascUnits = total - primaryDsc;
            return (ascUnits * 100.0 / upm, primaryDsc * 100.0 / upm);
        }
        catch
        {
            return (0, 0);
        }
    }

    /// <summary>
    /// Returns true when every codepoint in <paramref name="text"/> maps to
    /// a non-zero glyph in the font's cmap. Used to detect bullet-marker
    /// font fallback at render time — when the rPr-pinned font lacks a
    /// glyph, the renderer switches to a wider fallback face that inflates
    /// the line. Returns false on any read failure so the caller can
    /// default to the conservative "fallback may happen" branch.
    /// </summary>
    public static bool HasGlyphsForChars(string fontFamily, string text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        var hit = FindFont(fontFamily);
        if (!hit.HasValue) return false;

        try
        {
            using var fs = File.OpenRead(hit.Value.path);
            using var reader = new BinaryReader(fs);
            var offset = GetFontOffset(reader, hit.Value.idx);
            if (offset < 0) return false;
            var tables = FindTables(reader, offset);
            if (tables.cmap < 0) return false;

            var sub = SelectCmapSubtable(reader, tables.cmap);
            if (sub.format != 4 && sub.format != 12) return false;

            foreach (var cp in EnumerateCodepoints(text))
            {
                if (!CmapHasCodepoint(reader, sub, cp)) return false;
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    private struct CmapSubtable
    {
        public long offset;
        public int format;
    }

    private static CmapSubtable SelectCmapSubtable(BinaryReader r, long cmapOffset)
    {
        var result = new CmapSubtable { offset = -1, format = 0 };
        r.BaseStream.Position = cmapOffset + 2;
        var numTables = ReadUInt16BE(r);
        long bestF4 = -1, bestF12 = -1;
        for (int i = 0; i < numTables; i++)
        {
            r.BaseStream.Position = cmapOffset + 4 + i * 8;
            ushort platformId = ReadUInt16BE(r);
            ushort encodingId = ReadUInt16BE(r);
            uint subOff = ReadUInt32BE(r);
            // platform 3 + encoding 1 (UCS-2) → format 4
            // platform 3 + encoding 10 (UCS-4) or platform 0 (Unicode) → format 12
            r.BaseStream.Position = cmapOffset + subOff;
            ushort format = ReadUInt16BE(r);
            if (format == 12 && (platformId == 3 || platformId == 0))
                bestF12 = cmapOffset + subOff;
            else if (format == 4 && platformId == 3 && encodingId == 1)
                bestF4 = cmapOffset + subOff;
            else if (format == 4 && bestF4 < 0 && platformId == 0)
                bestF4 = cmapOffset + subOff;
        }
        if (bestF12 >= 0) { result.offset = bestF12; result.format = 12; }
        else if (bestF4 >= 0) { result.offset = bestF4; result.format = 4; }
        return result;
    }

    private static IEnumerable<int> EnumerateCodepoints(string s)
    {
        for (int i = 0; i < s.Length; i++)
        {
            int cp = char.IsHighSurrogate(s[i]) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1])
                ? char.ConvertToUtf32(s[i], s[i + 1])
                : s[i];
            if (cp > 0xFFFF) i++;
            yield return cp;
        }
    }

    private static bool CmapHasCodepoint(BinaryReader r, CmapSubtable sub, int cp)
    {
        if (sub.format == 12)
        {
            r.BaseStream.Position = sub.offset + 12;
            uint numGroups = ReadUInt32BE(r);
            int lo = 0, hi = (int)numGroups - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                r.BaseStream.Position = sub.offset + 16 + mid * 12;
                uint startChar = ReadUInt32BE(r);
                uint endChar = ReadUInt32BE(r);
                uint startGid = ReadUInt32BE(r);
                if ((uint)cp < startChar) hi = mid - 1;
                else if ((uint)cp > endChar) lo = mid + 1;
                else return startGid + ((uint)cp - startChar) != 0;
            }
            return false;
        }
        if (sub.format == 4)
        {
            if (cp > 0xFFFF) return false;
            r.BaseStream.Position = sub.offset + 6;
            int segCount = ReadUInt16BE(r) / 2;
            long endOff = sub.offset + 14;
            long startOff = endOff + segCount * 2 + 2;
            long deltaOff = startOff + segCount * 2;
            long rangeOff = deltaOff + segCount * 2;
            for (int i = 0; i < segCount; i++)
            {
                r.BaseStream.Position = endOff + i * 2;
                ushort end = ReadUInt16BE(r);
                if (end < cp) continue;
                r.BaseStream.Position = startOff + i * 2;
                ushort start = ReadUInt16BE(r);
                if (start > cp) return false;
                r.BaseStream.Position = deltaOff + i * 2;
                short idDelta = ReadInt16BE(r);
                r.BaseStream.Position = rangeOff + i * 2;
                ushort idRangeOff = ReadUInt16BE(r);
                int gid;
                if (idRangeOff == 0)
                {
                    gid = (cp + idDelta) & 0xFFFF;
                }
                else
                {
                    long glyphIdAddr = rangeOff + i * 2 + idRangeOff + (cp - start) * 2;
                    r.BaseStream.Position = glyphIdAddr;
                    int raw = ReadUInt16BE(r);
                    gid = raw == 0 ? 0 : (raw + idDelta) & 0xFFFF;
                }
                return gid != 0;
            }
            return false;
        }
        return false;
    }

    /// <summary>
    /// Return per-font ascent/descent percentages relative to em, suitable for
    /// CSS @font-face overrides. (0,0) when the font cannot be located.
    /// </summary>
    public static (double ascentPct, double descentPct) GetAscentDescentOverride(string fontFamily)
    {
        var hit = FindFont(fontFamily);
        if (!hit.HasValue) return (0, 0);

        try
        {
            using var fs = File.OpenRead(hit.Value.path);
            using var reader = new BinaryReader(fs);
            var offset = GetFontOffset(reader, hit.Value.idx);
            if (offset < 0) return (0, 0);

            var tables = FindTables(reader, offset);
            if (tables.head < 0 || tables.hhea < 0) return (0, 0);

            fs.Position = tables.head + 18;
            var upm = ReadUInt16BE(reader);
            if (upm == 0) return (0, 0);

            if (tables.os2 >= 0 && TryReadOs2(reader, tables.os2, out var os2) && os2.IsCjk)
            {
                int asc = os2.UseTypo ? os2.TypoAscent : os2.WinAscent;
                int dsc = os2.UseTypo ? -os2.TypoDescent : os2.WinDescent;
                int v7 = (15 * (asc + dsc) + 50) / 100;
                return ((asc + v7) * 100.0 / upm, (dsc + v7) * 100.0 / upm);
            }

            fs.Position = tables.hhea + 4;
            var ascender = ReadInt16BE(reader);
            var descender = ReadInt16BE(reader);

            return (ascender * 100.0 / upm, Math.Abs((int)descender) * 100.0 / upm);
        }
        catch
        {
            return (0, 0);
        }
    }
}
