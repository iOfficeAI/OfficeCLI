// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Xml.Linq;

namespace OfficeCli.Handlers;

internal class HwpxSection
{
    public int Index { get; set; }
    /// <summary>Actual ZIP entry path discovered from manifest (e.g. "Contents/section0.xml", "Contents/body_section.xml").</summary>
    public string EntryPath { get; init; } = null!;
    public XDocument Document { get; set; } = null!;
    public XElement Root => Document.Root!;
    public List<XElement> Paragraphs => Root.Elements(HwpxNs.Hp + "p").ToList();
    /// <summary>All tables: both direct children and nested inside paragraphs (Hancom format).</summary>
    public List<XElement> Tables => Root.Descendants(HwpxNs.Hp + "tbl").ToList();
}
