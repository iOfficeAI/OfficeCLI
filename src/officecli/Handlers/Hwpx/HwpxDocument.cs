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
    public HwpxSection PrimarySection => Sections[0];  // convenience

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
}
