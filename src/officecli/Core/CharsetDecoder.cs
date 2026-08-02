// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Decode a byte payload that carries its own charset declaration — an
/// embedded HTML/text part whose bytes were authored elsewhere and stored
/// verbatim, so they are not necessarily UTF-8.
///
/// Precedence follows what browsers do: a byte-order mark wins, then an
/// explicit charset parameter on the content type, then a &lt;meta charset&gt;
/// in the first 2 KB, then UTF-8.
///
/// .NET ships UTF-8, UTF-16/32, ASCII and Latin-1 only. The one legacy page
/// worth carrying without pulling in a code-page provider is windows-1252,
/// which authoring tools still emit for embedded HTML; it is Latin-1 plus 27
/// printable characters in 0x80-0x9F (curly quotes, dashes, ellipsis — exactly
/// what shows up in prose), so it is decoded as Latin-1 and remapped. Any other
/// declared charset that does not resolve falls back to UTF-8.
/// </summary>
public static class CharsetDecoder
{
    /// <summary>Read <paramref name="stream"/> fully and decode it.</summary>
    public static string Decode(Stream stream, string? declaredCharset)
    {
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Decode(buffer.GetBuffer().AsSpan(0, (int)buffer.Length), declaredCharset);
    }

    public static string Decode(ReadOnlySpan<byte> bytes, string? declaredCharset)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return Encoding.UTF8.GetString(bytes[3..]);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return Encoding.Unicode.GetString(bytes[2..]);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return Encoding.BigEndianUnicode.GetString(bytes[2..]);

        var name = Normalize(declaredCharset) ?? SniffMetaCharset(bytes);
        if (name == "windows-1252")
            return RemapWindows1252(Encoding.Latin1.GetString(bytes));

        var encoding = Resolve(name);
        return (encoding ?? Encoding.UTF8).GetString(bytes);
    }

    /// <summary>
    /// Pull the charset parameter out of a content type
    /// ("text/html; charset=windows-1252" → "windows-1252"). Null when absent.
    /// </summary>
    public static string? ParseCharset(string? contentType)
    {
        if (string.IsNullOrEmpty(contentType)) return null;
        foreach (var part in contentType.Split(';'))
        {
            var p = part.Trim();
            if (!p.StartsWith("charset", StringComparison.OrdinalIgnoreCase)) continue;
            var eq = p.IndexOf('=');
            if (eq < 0) continue;
            return Normalize(p[(eq + 1)..]);
        }
        return null;
    }

    private static string? Normalize(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        var n = name.Trim().Trim('"', '\'').ToLowerInvariant();
        return n.Length == 0 ? null : n switch
        {
            "utf8" => "utf-8",
            "latin1" or "iso8859-1" or "iso_8859-1" or "l1" => "iso-8859-1",
            "cp1252" or "win-1252" or "ansi" => "windows-1252",
            _ => n,
        };
    }

    private static Encoding? Resolve(string? name) => name switch
    {
        null => null,
        "utf-8" => Encoding.UTF8,
        "us-ascii" or "ascii" => Encoding.ASCII,
        "iso-8859-1" => Encoding.Latin1,
        "utf-16" or "utf-16le" or "unicode" => Encoding.Unicode,
        "utf-16be" => Encoding.BigEndianUnicode,
        // Anything else (DBCS pages such as gb2312 / shift_jis) needs a
        // code-page provider this build does not carry — fall back to UTF-8
        // rather than throwing mid-render.
        _ => null,
    };

    /// <summary>
    /// Scan the first 2 KB for `charset=...` (meta http-equiv or HTML5
    /// &lt;meta charset&gt;). ASCII-only scan: every charset this matters for
    /// is ASCII-compatible in that region.
    /// </summary>
    private static string? SniffMetaCharset(ReadOnlySpan<byte> bytes)
    {
        var head = Encoding.ASCII.GetString(bytes[..Math.Min(2048, bytes.Length)]);
        var at = head.IndexOf("charset", StringComparison.OrdinalIgnoreCase);
        if (at < 0) return null;
        var rest = head[(at + "charset".Length)..].TrimStart();
        if (rest.Length == 0 || rest[0] != '=') return null;
        rest = rest[1..].TrimStart();
        // <meta charset="windows-1252"> and charset='…' are both common; the
        // value scan below stops at the first non-name character, so strip the
        // opening quote or it terminates immediately and the sniff finds nothing.
        if (rest.Length > 0 && (rest[0] == '"' || rest[0] == '\'')) rest = rest[1..];
        var end = 0;
        while (end < rest.Length && (char.IsLetterOrDigit(rest[end]) || rest[end] is '-' or '_' or '.' or ':'))
            end++;
        return Normalize(rest[..end]);
    }

    // windows-1252 printable characters occupying 0x80-0x9F, where Latin-1 has
    // C1 controls. Index = byte - 0x80; '\0' marks the five unassigned slots,
    // which are left as decoded.
    private static readonly char[] Cp1252High =
    {
        '€', '\0',     '‚', 'ƒ', '„', '…', '†', '‡',
        'ˆ', '‰', 'Š', '‹', 'Œ', '\0',     'Ž', '\0',
        '\0',     '‘', '’', '“', '”', '•', '–', '—',
        '˜', '™', 'š', '›', 'œ', '\0',     'ž', 'Ÿ',
    };

    private static string RemapWindows1252(string latin1)
    {
        char[]? buf = null;
        for (int i = 0; i < latin1.Length; i++)
        {
            var c = latin1[i];
            if (c < '\u0080' || c > '\u009F') continue;
            var mapped = Cp1252High[c - '\u0080'];
            if (mapped == '\0') continue;
            buf ??= latin1.ToCharArray();
            buf[i] = mapped;
        }
        return buf is null ? latin1 : new string(buf);
    }
}
