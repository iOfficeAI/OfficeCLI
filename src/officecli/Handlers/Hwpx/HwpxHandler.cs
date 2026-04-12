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
    private readonly HashSet<string> _deletedBinData = new();

    public HwpxHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _editable = editable;
        Stream? stream = null;
        ZipArchive? archive = null;
        try
        {
            stream = new FileStream(filePath, FileMode.Open,
                editable ? FileAccess.ReadWrite : FileAccess.Read,
                FileShare.ReadWrite);
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

        // Plan 80: Rootfile-aware loading via HwpxManifest
        // Tries: container.xml → rootfile → OPF manifest → conventional fallback
        var manifest = HwpxManifest.Parse(archive);
        doc.RootfilePath = manifest.RootfilePath;

        // Load manifest doc (for SaveManifest and validation)
        var manifestPath = manifest.RootfilePath ?? "Contents/content.hpf";
        var hpfEntry = archive.GetEntry(manifestPath);
        if (hpfEntry != null)
        {
            using var hpfStream = hpfEntry.Open();
            doc.ManifestDoc = LoadAndNormalize(hpfStream);
            doc.ManifestEntryPath = hpfEntry.FullName;
        }

        // Load header
        if (!string.IsNullOrEmpty(manifest.HeaderPath))
        {
            var headerEntry = archive.GetEntry(manifest.HeaderPath);
            if (headerEntry != null)
            {
                doc.HeaderEntryPath = headerEntry.FullName;
                using var stream = headerEntry.Open();
                doc.Header = LoadAndNormalize(stream);
            }
        }

        // Fallback: conventional header path
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

        // Load sections from manifest-discovered paths
        int idx = 0;
        foreach (var sectionPath in manifest.SectionPaths)
        {
            var entry = archive.GetEntry(sectionPath);
            if (entry == null) continue;
            using var s = entry.Open();
            doc.Sections.Add(new HwpxSection
            {
                Index = idx++,
                EntryPath = entry.FullName,
                Document = LoadAndNormalize(s)
            });
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

    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        // HWPX binary extraction not yet implemented
        return false;
    }

    public void Dispose()
    {
        _doc.Archive.Dispose();
        _stream.Dispose();
    }
}
