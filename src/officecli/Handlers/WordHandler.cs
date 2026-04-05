// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler : IDocumentHandler
{
    private readonly WordprocessingDocument _doc;
    private readonly string _filePath;
    private HashSet<string> _usedParaIds = new(StringComparer.OrdinalIgnoreCase);
    private int _nextParaId = 0x100000;
    public int LastFindMatchCount { get; internal set; }

    public WordHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = WordprocessingDocument.Open(filePath, editable);
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

    /// <summary>
    /// Execute a JSON batch of operations on this document.
    /// Returns one BatchResult per item, with Success=true or Success=false+Error.
    /// </summary>
    public List<Core.BatchResult> Batch(string json)
    {
        var items = System.Text.Json.JsonSerializer.Deserialize(json, Core.BatchJsonContext.Default.ListBatchItem)
            ?? throw new ArgumentException("Invalid batch JSON");
        var results = new List<Core.BatchResult>();
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            try
            {
                var output = CommandBuilder.ExecuteBatchItem(this, item, json: false);
                results.Add(new Core.BatchResult { Index = i, Success = true, Output = output });
            }
            catch (Exception ex)
            {
                results.Add(new Core.BatchResult { Index = i, Success = false, Error = ex.Message, Item = item });
            }
        }
        return results;
    }

    public void Dispose()
    {
        _doc.Dispose();
    }

    // (private helpers, navigation, selector, style/list, image helpers moved to Word/ partial files)
}
