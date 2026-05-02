// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeCli.Handlers;

/// <summary>Parse HWPX OPF manifest to discover section order.</summary>
public class HwpxManifest
{
    /// <summary>Section entry paths in spine order.</summary>
    public List<string> SectionPaths { get; }

    /// <summary>Path to header.xml (typically "Contents/header.xml").</summary>
    public string HeaderPath { get; }

    private HwpxManifest(List<string> sectionPaths, string headerPath, string? rootfilePath = null)
    {
        SectionPaths = sectionPaths;
        HeaderPath = headerPath;
        RootfilePath = rootfilePath;
    }

    /// <summary>Rootfile path selected from container.xml (null if fallback used).</summary>
    public string? RootfilePath { get; private set; }

    /// <summary>Parse directly from a ZipArchive (Plan 80).</summary>
    public static HwpxManifest Parse(ZipArchive archive)
    {
        var entries = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries)
        {
            // Only read XML/text entries relevant to manifest discovery
            if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".hpf", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".opf", StringComparison.OrdinalIgnoreCase))
            {
                using var s = entry.Open();
                using var reader = new StreamReader(s, System.Text.Encoding.UTF8);
                entries[entry.FullName] = reader.ReadToEnd();
            }
        }
        return Parse(entries);
    }

    /// <summary>
    /// Parse OPF manifest from pre-read ZIP entries to discover section order.
    /// Tries container.xml → content.opf → content.hpf → fallback to entry scan.
    /// </summary>
    public static HwpxManifest Parse(Dictionary<string, string> entries)
    {
        string? opfXml = null;
        string? opfPath = null;
        string? rootfilePath = null;

        if (entries.TryGetValue("META-INF/container.xml", out var containerXml))
        {
            var containerDoc = XDocument.Parse(containerXml);
            var rootFile = containerDoc.Descendants()
                .FirstOrDefault(e => e.Name.LocalName == "rootfile");
            var fullPath = rootFile?.Attribute("full-path")?.Value;

            if (fullPath != null && entries.TryGetValue(fullPath, out var found))
            {
                opfXml = found;
                opfPath = fullPath;
                rootfilePath = fullPath;
            }
        }

        if (opfXml == null)
        {
            var candidates = new[] { "content.opf", "Contents/content.opf",
                                     "content.hpf", "Contents/content.hpf" };
            foreach (var candidate in candidates)
            {
                if (entries.TryGetValue(candidate, out var found))
                {
                    opfXml = found;
                    opfPath = candidate;
                    break;
                }
            }
        }

        if (opfXml != null)
        {
            var opfDoc = XDocument.Parse(opfXml);
            var manifest = ParseOpfManifest(opfDoc);
            if (manifest != null)
            {
                manifest.RootfilePath = rootfilePath;
                return manifest;
            }
        }

        var fallbackSections = FindSectionsFromEntries(entries);
        var headerPath = entries.ContainsKey("Contents/header.xml")
            ? "Contents/header.xml"
            : "";

        return new HwpxManifest(fallbackSections, headerPath, rootfilePath);
    }

    private static HwpxManifest? ParseOpfManifest(XDocument opfDoc)
    {
        var root = opfDoc.Root;
        if (root == null) return null;

        var manifestEl = root.Element(HwpxNs.Opf + "manifest")
                      ?? root.Elements().FirstOrDefault(e => e.Name.LocalName == "manifest");
        if (manifestEl == null) return null;

        var items = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in manifestEl.Elements()
            .Where(e => e.Name.LocalName == "item"))
        {
            var id = item.Attribute("id")?.Value;
            var href = item.Attribute("href")?.Value;
            if (id != null && href != null)
                items[id] = href;
        }

        var spineEl = root.Element(HwpxNs.Opf + "spine")
                   ?? root.Elements().FirstOrDefault(e => e.Name.LocalName == "spine");

        var sectionPaths = new List<string>();
        var headerPath = "";

        if (spineEl != null)
        {
            foreach (var itemref in spineEl.Elements()
                .Where(e => e.Name.LocalName == "itemref"))
            {
                var idref = itemref.Attribute("idref")?.Value;
                if (idref == null || !items.TryGetValue(idref, out var href))
                    continue;

                var fullPath = href.StartsWith("Contents/")
                    ? href
                    : $"Contents/{href}";

                if (fullPath.Contains("header", StringComparison.OrdinalIgnoreCase))
                {
                    headerPath = fullPath;
                }
                else if (fullPath.Contains("section", StringComparison.OrdinalIgnoreCase))
                {
                    sectionPaths.Add(fullPath);
                }
            }
        }

        if (sectionPaths.Count == 0)
        {
            foreach (var (_, href) in items)
            {
                var fullPath = href.StartsWith("Contents/") ? href : $"Contents/{href}";
                if (fullPath.Contains("section", StringComparison.OrdinalIgnoreCase))
                    sectionPaths.Add(fullPath);
                else if (fullPath.Contains("header", StringComparison.OrdinalIgnoreCase))
                    headerPath = fullPath;
            }
        }

        sectionPaths.Sort((a, b) =>
        {
            var numA = ExtractSectionNumber(a);
            var numB = ExtractSectionNumber(b);
            return numA.CompareTo(numB);
        });

        if (sectionPaths.Count == 0)
            return null;

        return new HwpxManifest(sectionPaths, headerPath);
    }

    // G1: Regex-based section file matching for non-standard paths
    private static readonly Regex SectionFileRegex = new(
        @"(?:^|[/\\])section[_\-]?(\d+)\.xml$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static List<string> FindSectionsFromEntries(Dictionary<string, string> entries)
    {
        var sectionPaths = new List<(string Path, int Number)>();
        foreach (var key in entries.Keys)
        {
            var m = SectionFileRegex.Match(key);
            if (m.Success)
            {
                var num = int.TryParse(m.Groups[1].Value, out var n) ? n : 0;
                sectionPaths.Add((key, num));
            }
        }

        return sectionPaths
            .OrderBy(x => x.Number)
            .Select(x => x.Path)
            .ToList();
    }

    private static int ExtractSectionNumber(string path)
    {
        var m = SectionFileRegex.Match(path);
        return m.Success && int.TryParse(m.Groups[1].Value, out var num) ? num : 0;
    }
}
