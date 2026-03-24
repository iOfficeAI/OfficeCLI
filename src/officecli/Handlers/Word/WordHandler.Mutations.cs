// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public string? Remove(string path)
    {
        // Handle /watermark removal
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
        {
            RemoveWatermarkHeaders();
            _doc.MainDocumentPart?.Document?.Save();
            return null;
        }

        var parts = ParsePath(path);

        // Handle header/footer removal by deleting the part itself
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() is "header" or "footer")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var idx = (parts[0].Index ?? 1) - 1;
            var isHeader = parts[0].Name.ToLowerInvariant() == "header";

            if (isHeader)
            {
                var headerPart = mainPart.HeaderParts.ElementAtOrDefault(idx)
                    ?? throw new ArgumentException($"Path not found: {path}");
                // Remove header references from section properties
                var partId = mainPart.GetIdOfPart(headerPart);
                foreach (var sectProps in mainPart.Document?.Body?.Descendants<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
                {
                    var refs = sectProps.Elements<HeaderReference>().Where(r => r.Id?.Value == partId).ToList();
                    foreach (var r in refs) r.Remove();
                }
                mainPart.DeletePart(headerPart);
            }
            else
            {
                var footerPart = mainPart.FooterParts.ElementAtOrDefault(idx)
                    ?? throw new ArgumentException($"Path not found: {path}");
                var partId = mainPart.GetIdOfPart(footerPart);
                foreach (var sectProps in mainPart.Document?.Body?.Descendants<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
                {
                    var refs = sectProps.Elements<FooterReference>().Where(r => r.Id?.Value == partId).ToList();
                    foreach (var r in refs) r.Remove();
                }
                mainPart.DeletePart(footerPart);
            }

            mainPart.Document?.Save();
            return null;
        }

        var element = NavigateToElement(parts, out var ctx)
            ?? throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));

        element.Remove();
        _doc.MainDocumentPart?.Document?.Save();
        return null;
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        // Determine target parent
        string effectiveParentPath;
        OpenXmlElement targetParent;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within current parent
            targetParent = element.Parent
                ?? throw new InvalidOperationException("Element has no parent");
            // Compute parent path by removing last segment
            var lastSlash = sourcePath.LastIndexOf('/');
            effectiveParentPath = lastSlash > 0 ? sourcePath[..lastSlash] : "/body";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            if (targetParentPath is "/" or "" or "/body")
                targetParent = _doc.MainDocumentPart!.Document!.Body!;
            else
            {
                var tgtParts = ParsePath(targetParentPath);
                targetParent = NavigateToElement(tgtParts)
                    ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
            }
        }

        element.Remove();

        // Insert at the specified position among same-type siblings (0-based index)
        if (index.HasValue)
        {
            var sameTypeSiblings = targetParent.ChildElements
                .Where(e => e.LocalName == element.LocalName).ToList();
            if (index.Value >= 0 && index.Value < sameTypeSiblings.Count)
                sameTypeSiblings[index.Value].InsertBeforeSelf(element);
            else
                AppendToParent(targetParent, element);
        }
        else
        {
            targetParent.AppendChild(element);
        }

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == element.LocalName).ToList();
        var newIdx = siblings.IndexOf(element) + 1;
        return $"{effectiveParentPath}/{element.LocalName}[{newIdx}]";
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        var parts1 = ParsePath(path1);
        var elem1 = NavigateToElement(parts1)
            ?? throw new ArgumentException($"Element not found: {path1}");
        var parts2 = ParsePath(path2);
        var elem2 = NavigateToElement(parts2)
            ?? throw new ArgumentException($"Element not found: {path2}");

        if (elem1.Parent != elem2.Parent)
            throw new ArgumentException("Cannot swap elements with different parents");

        PowerPointHandler.SwapXmlElements(elem1, elem2);
        _doc.MainDocumentPart?.Document?.Save();

        // Recompute paths
        var parent = elem1.Parent!;
        var lastSlash = path1.LastIndexOf('/');
        var parentPath = lastSlash > 0 ? path1[..lastSlash] : "/body";

        var siblings1 = parent.ChildElements.Where(e => e.LocalName == elem1.LocalName).ToList();
        var newIdx1 = siblings1.IndexOf(elem1) + 1;
        var siblings2 = parent.ChildElements.Where(e => e.LocalName == elem2.LocalName).ToList();
        var newIdx2 = siblings2.IndexOf(elem2) + 1;
        return ($"{parentPath}/{elem1.LocalName}[{newIdx1}]", $"{parentPath}/{elem2.LocalName}[{newIdx2}]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        var clone = element.CloneNode(true);

        OpenXmlElement targetParent;
        if (targetParentPath is "/" or "" or "/body")
            targetParent = _doc.MainDocumentPart!.Document!.Body!;
        else
        {
            var tgtParts = ParsePath(targetParentPath);
            targetParent = NavigateToElement(tgtParts)
                ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
        }

        InsertAtPosition(targetParent, clone, index);

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == clone.LocalName).ToList();
        var newIdx = siblings.IndexOf(clone) + 1;
        return $"{targetParentPath}/{clone.LocalName}[{newIdx}]";
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }
}
