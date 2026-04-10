// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Xml.Linq;
using System.Text.Json.Nodes;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler : IDocumentHandler
{
    private readonly HwpxDocument _doc;
    private readonly string _filePath;
    private readonly bool _editable;
    private readonly Stream _stream;
    private bool _dirty;

    public HwpxHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _editable = editable;
        Stream? stream = null;
        ZipArchive? archive = null;
        try
        {
            stream = File.Open(filePath, FileMode.Open,
                editable ? FileAccess.ReadWrite : FileAccess.Read);
            archive = new ZipArchive(stream,
                editable ? ZipArchiveMode.Update : ZipArchiveMode.Read);
            _doc = LoadDocument(archive);
            _stream = stream;
        }
        catch
        {
            archive?.Dispose();
            stream?.Dispose();
            throw;
        }
    }

    private static HwpxDocument LoadDocument(ZipArchive archive)
    {
        var doc = new HwpxDocument { Archive = archive };

        // Discover header and sections from content.hpf manifest
        var hpfEntry = archive.GetEntry("Contents/content.hpf");
        if (hpfEntry != null)
        {
            using var hpfStream = hpfEntry.Open();
            var hpf = LoadAndNormalize(hpfStream);
            var allItems = hpf.Descendants()
                .Where(e => e.Name.LocalName == "item")
                .ToList();

            // Discover header from manifest: item whose href contains "header" or id contains "head"
            var headerItem = allItems.FirstOrDefault(e =>
                (e.Attribute("href")?.Value?.Contains("header", StringComparison.OrdinalIgnoreCase) ?? false)
                || (e.Attribute("id")?.Value?.Contains("head", StringComparison.OrdinalIgnoreCase) ?? false));
            if (headerItem != null)
            {
                var headerHref = headerItem.Attribute("href")?.Value;
                if (headerHref != null)
                {
                    var headerPath = headerHref.StartsWith("Contents/") ? headerHref : $"Contents/{headerHref}";
                    var headerEntry = archive.GetEntry(headerPath);
                    if (headerEntry != null)
                    {
                        doc.HeaderEntryPath = headerEntry.FullName;
                        using var stream = headerEntry.Open();
                        doc.Header = LoadAndNormalize(stream);
                    }
                }
            }

            // Build item id→href lookup for spine resolution
            var itemById = allItems
                .Where(e => e.Attribute("id")?.Value != null && e.Attribute("href")?.Value != null)
                .ToDictionary(
                    e => e.Attribute("id")!.Value,
                    e => e,
                    StringComparer.Ordinal);

            // Discover sections: prefer spine order (OPF reading order), fall back to manifest order
            var spine = hpf.Descendants()
                .Where(e => e.Name.LocalName == "itemref")
                .Select(e => e.Attribute("idref")?.Value)
                .Where(id => id != null)
                .ToList();

            var orderedSectionHrefs = new List<string>();
            if (spine.Count > 0)
            {
                // Spine order: resolve each itemref to its manifest item, keep only sections
                foreach (var idref in spine)
                {
                    if (itemById.TryGetValue(idref!, out var item))
                    {
                        var mt = item.Attribute("media-type")?.Value;
                        if (mt != null && mt.Contains("section"))
                        {
                            var href = item.Attribute("href")?.Value;
                            if (href != null) orderedSectionHrefs.Add(href);
                        }
                    }
                }
            }

            // Fall back to manifest order if spine is absent or yielded nothing
            if (orderedSectionHrefs.Count == 0)
            {
                orderedSectionHrefs.AddRange(
                    allItems
                        .Where(e => e.Attribute("media-type")?.Value?.Contains("section") ?? false)
                        .Select(e => e.Attribute("href")?.Value!)
                        .Where(h => h != null));
            }

            int idx = 0;
            foreach (var href in orderedSectionHrefs)
            {
                var entryPath = href.StartsWith("Contents/") ? href : $"Contents/{href}";
                var entry = archive.GetEntry(entryPath);
                if (entry == null) continue;
                using var s = entry.Open();
                doc.Sections.Add(new HwpxSection
                {
                    Index = idx++,
                    EntryPath = entry.FullName,
                    Document = LoadAndNormalize(s)
                });
            }
        }

        // Fallback: load header from conventional path if not discovered from manifest
        if (doc.Header == null)
        {
            var headerEntry = archive.GetEntry("Contents/header.xml");
            if (headerEntry != null)
            {
                doc.HeaderEntryPath = headerEntry.FullName;
                using var stream = headerEntry.Open();
                doc.Header = LoadAndNormalize(stream);
            }
        }

        // Fallback: try section0.xml, section1.xml, ...
        if (doc.Sections.Count == 0)
        {
            for (int i = 0; i < 100; i++)
            {
                var entry = archive.GetEntry($"Contents/section{i}.xml");
                if (entry == null) break;
                using var s = entry.Open();
                doc.Sections.Add(new HwpxSection
                {
                    Index = i,
                    EntryPath = entry.FullName,
                    Document = LoadAndNormalize(s)
                });
            }
        }

        if (doc.Sections.Count == 0)
            throw new InvalidOperationException("No sections found in HWPX document");

        return doc;
    }

    // --- Helper: read ZIP entry, normalize HWPML 2016→2011 namespaces, then parse ---
    private static XDocument LoadAndNormalize(Stream stream)
    {
        using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
        var raw = reader.ReadToEnd();
        foreach (var (old, canonical) in HwpxNs.LegacyToCanonical)
            raw = raw.Replace(old, canonical, StringComparison.Ordinal);
        return XDocument.Parse(raw);
    }

    public void Dispose()
    {
        _doc.Archive.Dispose();
        _stream.Dispose();
    }
}
