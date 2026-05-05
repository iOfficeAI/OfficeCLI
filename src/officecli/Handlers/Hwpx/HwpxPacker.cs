// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

/// <summary>Read/write HWPX ZIP container.</summary>
public static class HwpxPacker
{
    // ==================== Read ====================

    /// <summary>
    /// Read all XML entries from an HWPX ZIP file.
    /// Non-XML entries (images, binaries) are skipped.
    /// Namespace normalization is applied to every entry.
    /// </summary>
    public static Dictionary<string, string> ReadAllEntries(string path)
    {
        if (!File.Exists(path))
            throw new CliException($"File not found: {path}");

        var entries = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && !entry.FullName.EndsWith(".hpf", StringComparison.OrdinalIgnoreCase)
                && !entry.FullName.EndsWith(".opf", StringComparison.OrdinalIgnoreCase))
                continue;

            entries[entry.FullName] = ReadEntry(zip, entry.FullName);
        }

        return entries;
    }

    /// <summary>
    /// Read a single ZIP entry as text, applying namespace normalization.
    /// </summary>
    public static string ReadEntry(ZipArchive zip, string entryName)
    {
        var entry = zip.GetEntry(entryName)
            ?? throw new CliException($"Entry not found in ZIP: {entryName}");

        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xml = reader.ReadToEnd();

        return NormalizeNamespaces(xml);
    }

    private static string NormalizeNamespaces(string xml)
    {
        foreach (var (legacy, canonical) in HwpxNs.LegacyToCanonical)
        {
            xml = xml.Replace(legacy, canonical);
        }

        return xml;
    }

    /// <summary>
    /// Minify XML by removing whitespace between tags.
    /// Hancom Office requires single-line XML — pretty-printed XML renders blank.
    /// </summary>
    public static string MinifyXml(string xml)
    {
        // Remove whitespace between > and < (inter-element whitespace)
        // But preserve whitespace inside text content
        return System.Text.RegularExpressions.Regex.Replace(xml, @">\s+<", "><");
    }

    /// <summary>
    /// Reverse the namespace normalization before saving back to ZIP.
    /// Hancom Office expects the original 2016 namespace URIs to remain intact.
    /// Only restores namespaces that appear as xmlns declarations (not element content).
    /// </summary>
    public static string RestoreOriginalNamespaces(string xml)
    {
        // Fix 1: Restore hp10 namespace declaration (2011→2016)
        xml = xml.Replace(
            "xmlns:hp10=\"http://www.hancom.co.kr/hwpml/2011/paragraph\"",
            "xmlns:hp10=\"http://www.hancom.co.kr/hwpml/2016/paragraph\"");

        // Fix 2: XDocument may have swapped hp: ↔ hp10: prefixes since both resolve
        // to the same URI after normalization. Restore original prefix usage:
        // Elements that should use hp: but got hp10: due to namespace collapse
        xml = xml.Replace("<hp10:p ", "<hp:p ");
        xml = xml.Replace("</hp10:p>", "</hp:p>");
        xml = xml.Replace("<hp10:run ", "<hp:run ");
        xml = xml.Replace("</hp10:run>", "</hp:run>");
        xml = xml.Replace("<hp10:t>", "<hp:t>");
        xml = xml.Replace("</hp10:t>", "</hp:t>");
        xml = xml.Replace("<hp10:tbl ", "<hp:tbl ");
        xml = xml.Replace("</hp10:tbl>", "</hp:tbl>");
        xml = xml.Replace("<hp10:tr>", "<hp:tr>");
        xml = xml.Replace("</hp10:tr>", "</hp:tr>");
        xml = xml.Replace("<hp10:tr ", "<hp:tr ");
        xml = xml.Replace("<hp10:tc ", "<hp:tc ");
        xml = xml.Replace("</hp10:tc>", "</hp:tc>");
        xml = xml.Replace("<hp10:subList ", "<hp:subList ");
        xml = xml.Replace("</hp10:subList>", "</hp:subList>");
        xml = xml.Replace("<hp10:cellAddr ", "<hp:cellAddr ");
        xml = xml.Replace("<hp10:cellSpan ", "<hp:cellSpan ");
        xml = xml.Replace("<hp10:cellSz ", "<hp:cellSz ");
        xml = xml.Replace("<hp10:cellMargin ", "<hp:cellMargin ");
        xml = xml.Replace("<hp10:colSz ", "<hp:colSz ");
        xml = xml.Replace("<hp10:sz ", "<hp:sz ");
        xml = xml.Replace("<hp10:pos ", "<hp:pos ");
        xml = xml.Replace("<hp10:outMargin ", "<hp:outMargin ");
        xml = xml.Replace("<hp10:inMargin ", "<hp:inMargin ");
        xml = xml.Replace("<hp10:secPr ", "<hp:secPr ");
        xml = xml.Replace("</hp10:secPr>", "</hp:secPr>");
        xml = xml.Replace("<hp10:linesegarray>", "<hp:linesegarray>");
        xml = xml.Replace("</hp10:linesegarray>", "</hp:linesegarray>");
        xml = xml.Replace("<hp10:lineseg ", "<hp:lineseg ");
        xml = xml.Replace("<hp10:ctrl ", "<hp:ctrl ");
        xml = xml.Replace("</hp10:ctrl>", "</hp:ctrl>");
        // Catch-all: any remaining hp10: elements → hp:
        xml = System.Text.RegularExpressions.Regex.Replace(xml, @"<hp10:(\w+)", "<hp:$1");
        xml = System.Text.RegularExpressions.Regex.Replace(xml, @"</hp10:(\w+)>", "</hp:$1>");

        // Fix attributes: LINQ to XML may swap hp10: for hp: in namespaced attributes
        // (e.g. hp10:required-namespace) since both resolved to the same URI during
        // normalization. Original files always use hp: for attributes.
        xml = System.Text.RegularExpressions.Regex.Replace(xml, @" hp10:([\w-]+)=""", " hp:$1=\"");

        return xml;
    }

    // ==================== Write ====================

    /// <summary>
    /// Pack entries into a new HWPX ZIP file.
    /// Atomic write: creates temp file → validates → renames to target.
    /// </summary>
    public static void Pack(string targetPath, Dictionary<string, string> entries,
                            string? mimeType = null)
    {
        var tempPath = targetPath + ".tmp";

        try
        {
            using (var fs = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                // 1. mimetype MUST be first entry, stored (no compression)
                var mime = mimeType ?? "application/hwp+zip";
                var mimeEntry = zip.CreateEntry("mimetype", CompressionLevel.NoCompression);
                using (var mimeStream = mimeEntry.Open())
                {
                    var mimeBytes = Encoding.ASCII.GetBytes(mime);
                    mimeStream.Write(mimeBytes, 0, mimeBytes.Length);
                }

                // 2. All other entries: Deflate compression
                foreach (var (name, content) in entries)
                {
                    var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
                    using var entryStream = entry.Open();
                    var bytes = Encoding.UTF8.GetBytes(content);
                    entryStream.Write(bytes, 0, bytes.Length);
                }
            }

            // 3. Validate the temp file before committing
            ValidateZip(tempPath);

            // 4. Atomic rename
            File.Move(tempPath, targetPath, overwrite: true);
        }
        catch
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
            throw;
        }
    }

    private static void ValidateZip(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            using var stream = entry.Open();
            var buffer = new byte[4096];
            while (stream.Read(buffer, 0, buffer.Length) > 0) { }
        }
    }

    // ==================== XML Utilities ====================

    /// <summary>
    /// Remove all &lt;hp:linesegarray&gt; blocks from XML.
    /// These are stale layout cache elements generated by Hancom's renderer.
    /// </summary>
    public static string StripLinesegarray(string xml)
    {
        return Regex.Replace(
            xml,
            @"<hp:linesegarray[^>]*>.*?</hp:linesegarray>",
            "",
            RegexOptions.Singleline);
    }

    // Removed duplicate MinifyXml — the one above (regex-based) handles inter-element whitespace correctly
}
