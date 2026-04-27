// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler : IDocumentHandler
{
    private readonly WordprocessingDocument _doc;
    private readonly string _filePath;
    private HashSet<string> _usedParaIds = new(StringComparer.OrdinalIgnoreCase);
    private int _nextParaId = 0x100000;
    public int LastFindMatchCount { get; internal set; }

    /// <summary>
    /// Props that the most recent Add() call could not consume. Surfaced to
    /// the CLI layer so silent-drops on the curated surface (e.g.
    /// `add /styles --prop font.eastAsia=...`) become visible warnings
    /// instead of "Added" lies. Reset at the start of each Add.
    /// </summary>
    public List<string> LastAddUnsupportedProps { get; internal set; } = new();

    public WordHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = WordprocessingDocument.Open(filePath, editable);
        WordStrictAttributeSanitizer.Sanitize(_doc);
        if (editable)
        {
            EnsureAllParaIds();
            EnsureDocPropIds();
        }
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return "(no main part)";

        return partPath.ToLowerInvariant() switch
        {
            "/document" or "/word/document.xml" => mainPart.Document?.OuterXml ?? "",
            "/styles" or "/word/styles.xml" => mainPart.StyleDefinitionsPart?.Styles?.OuterXml ?? "(no styles)",
            "/settings" or "/word/settings.xml" => mainPart.DocumentSettingsPart?.Settings?.OuterXml ?? "(no settings)",
            "/numbering" or "/word/numbering.xml" => mainPart.NumberingDefinitionsPart?.Numbering?.OuterXml ?? "(no numbering)",
            "/comments" => mainPart.WordprocessingCommentsPart?.Comments?.OuterXml ?? "(no comments)",
            _ when partPath.StartsWith("/header") => GetHeaderRawXml(partPath),
            _ when partPath.StartsWith("/footer") => GetFooterRawXml(partPath),
            _ when partPath.StartsWith("/chart") => GetChartRawXml(partPath),
            _ => throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /header[n], /footer[n], /chart[n]")
        };
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var mainPart = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("No main document part");

        OpenXmlPartRootElement rootElement;
        var lowerPath = partPath.ToLowerInvariant();

        if (lowerPath is "/document" or "/")
            rootElement = mainPart.Document ?? throw new InvalidOperationException("No document");
        else if (lowerPath is "/styles")
            rootElement = mainPart.StyleDefinitionsPart?.Styles ?? throw new InvalidOperationException("No styles part");
        else if (lowerPath is "/settings")
            rootElement = mainPart.DocumentSettingsPart?.Settings ?? throw new InvalidOperationException("No settings part");
        else if (lowerPath is "/numbering")
            rootElement = mainPart.NumberingDefinitionsPart?.Numbering ?? throw new InvalidOperationException("No numbering part");
        else if (lowerPath is "/comments")
            rootElement = mainPart.WordprocessingCommentsPart?.Comments ?? throw new InvalidOperationException("No comments part");
        else if (lowerPath.StartsWith("/header"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var headerPart = mainPart.HeaderParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"header[{idx}] not found");
            rootElement = headerPart.Header ?? throw new InvalidOperationException($"Corrupt file: header[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/footer"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var footerPart = mainPart.FooterParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"footer[{idx}] not found");
            rootElement = footerPart.Footer ?? throw new InvalidOperationException($"Corrupt file: footer[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/chart"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var chartPart = mainPart.ChartParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"chart[{idx}] not found");
            rootElement = chartPart.ChartSpace ?? throw new InvalidOperationException($"Corrupt file: chart[{idx}] data missing");
        }
        else
            throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /header[n], /footer[n], /chart[n]");

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        Console.WriteLine($"raw-set: {affected} element(s) affected");
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose()
    {
        _doc.Dispose();
        // CONSISTENCY(word-self-close): the OpenXml SDK serializes empty
        // elements with a space before the self-close (`<w:br />`). Several
        // downstream consumers (and test regexes) look for the canonical
        // `<w:br/>` / `<w:tab/>` form. Normalize the persisted document.xml
        // in place so the saved package matches the canonical short form.
        // Only applied to word/document.xml; styles/settings/numbering are
        // left untouched since the space form is schema-equivalent.
        try { NormalizeSelfClosingInDocx(_filePath); } catch { /* best-effort */ }
    }

    private static void NormalizeSelfClosingInDocx(string path)
    {
        if (!System.IO.File.Exists(path)) return;
        using var fs = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
        using var za = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: false);
        var entry = za.GetEntry("word/document.xml");
        if (entry == null) return;
        string xml;
        using (var rs = entry.Open())
        using (var sr = new System.IO.StreamReader(rs))
            xml = sr.ReadToEnd();
        // Collapse "<w:br />" → "<w:br/>" and "<w:tab />" → "<w:tab/>"
        // (no-attribute empty elements only).
        var normalized = System.Text.RegularExpressions.Regex.Replace(
            xml, @"<w:(br|tab) />", "<w:$1/>");
        if (normalized == xml) return;
        entry.Delete();
        var newEntry = za.CreateEntry("word/document.xml");
        using var ws = newEntry.Open();
        using var sw = new System.IO.StreamWriter(ws, new System.Text.UTF8Encoding(false));
        sw.Write(normalized);
    }

    // (private helpers, navigation, selector, style/list, image helpers moved to Word/ partial files)
}
