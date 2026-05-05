// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Xml.Linq;

namespace OfficeCli.Handlers;

internal class HwpxDocument
{
    public ZipArchive Archive { get; init; } = null!;
    public XDocument? Header { get; set; }
    /// <summary>Actual ZIP entry path of header.xml (e.g. "Contents/header.xml").</summary>
    public string? HeaderEntryPath { get; set; }
    public List<HwpxSection> Sections { get; } = new();
    /// <summary>Parsed content.hpf manifest document for section/spine management.</summary>
    public XDocument? ManifestDoc { get; set; }
    /// <summary>ZIP entry path for the manifest (e.g. "Contents/content.hpf").</summary>
    public string? ManifestEntryPath { get; set; }
    /// <summary>Selected rootfile path from container.xml (null if conventional fallback used).</summary>
    public string? RootfilePath { get; set; }
    public HwpxSection PrimarySection => Sections[0];  // convenience

    /// <summary>Read binary data from BinData directory in the ZIP archive.</summary>
    /// <remarks>
    /// Hancom stores BinData at the archive root (BinData/imageN.ext), not under Contents/.
    /// We try root-level first (canonical), then Contents/ prefix (legacy fallback).
    /// </remarks>
    public byte[]? GetBinData(string reference)
    {
        var name = reference.StartsWith("BinData/") ? reference : $"BinData/{reference}";
        // Try root-level first (Hancom canonical path)
        var entry = Archive.GetEntry(name)
            // Fallback: Contents/ prefix (legacy officecli path)
            ?? Archive.GetEntry($"Contents/{name}");
        if (entry == null) return null;
        using var stream = entry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>All paragraphs across all sections. SectionIndex is 0-based LOCAL index within that section.</summary>
    public IEnumerable<(HwpxSection Section, XElement Paragraph, int SectionIndex)> AllParagraphs()
    {
        foreach (var sec in Sections)
        {
            int localIdx = 0;
            foreach (var p in sec.Paragraphs)
                yield return (sec, p, localIdx++);
        }
    }

    /// <summary>All tables across all sections. SectionIndex is 0-based LOCAL index within that section.</summary>
    public IEnumerable<(HwpxSection Section, XElement Table, int SectionIndex)> AllTables()
    {
        foreach (var sec in Sections)
        {
            int localIdx = 0;
            foreach (var tbl in sec.Tables)
                yield return (sec, tbl, localIdx++);
        }
    }

    /// <summary>All content elements (paragraphs + table cells) in document order for text extraction.
    /// Handles both officecli-created tables (direct section children) and
    /// Hancom-created tables (nested inside p &gt; run &gt; tbl).</summary>
    public IEnumerable<(HwpxSection Section, XElement Paragraph, string Path)> AllContentInOrder()
    {
        foreach (var sec in Sections)
        {
            int paraIdx = 0;
            int tblIdx = 0;
            foreach (var child in sec.Root.Elements())
            {
                var localName = child.Name.LocalName;
                if (localName == "p")
                {
                    paraIdx++;
                    yield return (sec, child, $"/section[{sec.Index + 1}]/p[{paraIdx}]");

                    // Hancom nests tables inside p > run > tbl
                    var nestedTables = child.Descendants(HwpxNs.Hp + "tbl");
                    foreach (var ntbl in nestedTables)
                    {
                        tblIdx++;
                        foreach (var item in EnumerateTableCells(sec, ntbl, tblIdx))
                            yield return item;
                    }
                }
                else if (localName == "tbl")
                {
                    tblIdx++;
                    foreach (var item in EnumerateTableCells(sec, child, tblIdx))
                        yield return item;
                }
            }
        }
    }

    private static IEnumerable<(HwpxSection Section, XElement Paragraph, string Path)> EnumerateTableCells(
        HwpxSection sec, XElement tbl, int tblIdx)
    {
        int rowIdx = 0;
        foreach (var tr in tbl.Elements(HwpxNs.Hp + "tr"))
        {
            rowIdx++;
            int cellIdx = 0;
            foreach (var tc in tr.Elements(HwpxNs.Hp + "tc"))
            {
                cellIdx++;
                var subList = tc.Element(HwpxNs.Hp + "subList");
                var paragraphs = subList?.Elements(HwpxNs.Hp + "p")
                    ?? tc.Elements(HwpxNs.Hp + "p");
                int cellParaIdx = 0;
                foreach (var p in paragraphs)
                {
                    cellParaIdx++;
                    var path = $"/section[{sec.Index + 1}]/tbl[{tblIdx}]/tr[{rowIdx}]/tc[{cellIdx}]/p[{cellParaIdx}]";
                    yield return (sec, p, path);
                }
            }
        }
    }
}
