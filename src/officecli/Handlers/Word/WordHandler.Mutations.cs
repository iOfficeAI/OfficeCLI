// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
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
                // Clean up ImageParts referenced only by this header
                CleanupImageParts(mainPart, headerPart.Header?.Descendants<A.Blip>(), headerPart);
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
                // Clean up ImageParts referenced only by this footer
                CleanupImageParts(mainPart, footerPart.Footer?.Descendants<A.Blip>(), footerPart);
                mainPart.DeletePart(footerPart);
            }

            mainPart.Document?.Save();
            return null;
        }

        // Handle TOC removal
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() == "toc")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var tocIdx = parts[0].Index ?? 1;
            var tocParas = FindTocParagraphs();
            if (tocIdx < 1 || tocIdx > tocParas.Count)
                throw new ArgumentException($"TOC {tocIdx} not found (total: {tocParas.Count})");

            var tocPara = tocParas[tocIdx - 1];

            // Also remove preceding TOCHeading title paragraph if present
            var prevSibling = tocPara.PreviousSibling<Paragraph>();
            if (prevSibling != null)
            {
                var styleId = prevSibling.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && styleId.Equals("TOCHeading", StringComparison.OrdinalIgnoreCase))
                    prevSibling.Remove();
            }

            tocPara.Remove();
            mainPart.Document?.Save();
            return null;
        }

        // Handle footnote/endnote removal
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() == "footnote")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var fnId = parts[0].Index ?? 1;
            var fn = mainPart.FootnotesPart?.Footnotes?
                .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId)
                ?? throw new ArgumentException($"Path not found: {path}");
            // Remove footnote reference from body
            foreach (var fnRef in mainPart.Document.Descendants<FootnoteReference>()
                .Where(r => r.Id?.Value == fnId).ToList())
                fnRef.Parent?.Remove();
            fn.Remove();
            mainPart.FootnotesPart?.Footnotes?.Save();
            mainPart.Document?.Save();
            return null;
        }
        if (parts.Count == 1 && parts[0].Name.ToLowerInvariant() == "endnote")
        {
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var enId = parts[0].Index ?? 1;
            var en = mainPart.EndnotesPart?.Endnotes?
                .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId)
                ?? throw new ArgumentException($"Path not found: {path}");
            // Remove endnote reference from body
            foreach (var enRef in mainPart.Document.Descendants<EndnoteReference>()
                .Where(r => r.Id?.Value == enId).ToList())
                enRef.Parent?.Remove();
            en.Remove();
            mainPart.EndnotesPart?.Endnotes?.Save();
            mainPart.Document?.Save();
            return null;
        }

        // Handle /chart[N] removal
        var chartRemoveMatch = Regex.Match(path, @"^/chart\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (chartRemoveMatch.Success)
        {
            var chartIdx = int.Parse(chartRemoveMatch.Groups[1].Value);
            var mainPart = _doc.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart not found");
            var chartParts = mainPart.ChartParts.ToList();
            if (chartIdx < 1 || chartIdx > chartParts.Count)
                throw new ArgumentException($"Chart index {chartIdx} out of range (1..{chartParts.Count})");
            var chartPart = chartParts[chartIdx - 1];
            var relId = mainPart.GetIdOfPart(chartPart);
            // Find and remove the Run containing the ChartReference in the body
            var chartRef = mainPart.Document?.Body?
                .Descendants<C.ChartReference>()
                .FirstOrDefault(cr => cr.Id?.Value == relId);
            if (chartRef != null)
            {
                var run = chartRef.Ancestors<Run>().FirstOrDefault();
                run?.Remove();
            }
            mainPart.DeletePart(chartPart);
            mainPart.Document?.Save();
            return null;
        }

        var element = NavigateToElement(parts, out var ctx)
            ?? throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));

        // Clean up ImageParts referenced by any inline/anchor pictures in the element
        var mainPart2 = _doc.MainDocumentPart;
        if (mainPart2 != null)
        {
            foreach (var blip in element.Descendants<A.Blip>())
            {
                var embedId = blip.Embed?.Value;
                if (!string.IsNullOrEmpty(embedId))
                {
                    // Count how many times this embedId is referenced across body + headers + footers
                    var refCount = mainPart2.Document.Descendants<A.Blip>()
                        .Count(b => b.Embed?.Value == embedId);
                    foreach (var hp in mainPart2.HeaderParts)
                        refCount += hp.Header?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
                    foreach (var fp in mainPart2.FooterParts)
                        refCount += fp.Footer?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
                    if (refCount <= 1)
                    {
                        try { mainPart2.DeletePart(embedId); } catch { }
                    }
                }
            }
        }

        // If removing a Comment, also clean up dangling references in the body
        if (element is Comment comment && comment.Id?.Value is string commentId)
        {
            var body2 = _doc.MainDocumentPart?.Document?.Body;
            if (body2 != null)
            {
                foreach (var rs in body2.Descendants<CommentRangeStart>()
                    .Where(r => r.Id?.Value == commentId).ToList())
                    rs.Remove();
                foreach (var re in body2.Descendants<CommentRangeEnd>()
                    .Where(r => r.Id?.Value == commentId).ToList())
                    re.Remove();
                foreach (var cr in body2.Descendants<CommentReference>()
                    .Where(r => r.Id?.Value == commentId).ToList())
                    cr.Parent?.Remove(); // Remove the containing Run
            }
        }

        element.Remove();
        _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments?.Save();
        _doc.MainDocumentPart?.Document?.Save();
        return null;
    }

    /// <summary>
    /// Clean up ImageParts in a header/footer part that are not referenced elsewhere.
    /// </summary>
    private static void CleanupImageParts(MainDocumentPart mainPart, IEnumerable<A.Blip>? blips, OpenXmlPart ownerPart)
    {
        if (blips == null) return;
        foreach (var blip in blips.ToList())
        {
            var embedId = blip.Embed?.Value;
            if (string.IsNullOrEmpty(embedId)) continue;

            // Count references across body + all headers + all footers (excluding the part being deleted)
            var refCount = mainPart.Document?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
            foreach (var hp in mainPart.HeaderParts.Where(p => p != ownerPart))
                refCount += hp.Header?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;
            foreach (var fp in mainPart.FooterParts.Where(p => p != ownerPart))
                refCount += fp.Footer?.Descendants<A.Blip>().Count(b => b.Embed?.Value == embedId) ?? 0;

            if (refCount == 0)
            {
                try { mainPart.DeletePart(embedId); } catch { }
            }
        }
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

    // ==================== Track Changes ====================

    /// <summary>
    /// Accept all tracked changes in the document.
    /// - w:ins (InsertedRun): unwrap — keep inner content, remove wrapper
    /// - w:del (DeletedRun): remove entire element
    /// - w:rPrChange (RunPropertiesChange): remove change marker, keep current formatting
    /// - w:pPrChange (ParagraphPropertiesChange): remove change marker, keep current formatting
    /// - w:sectPrChange (SectionPropertiesChange): remove change marker
    /// - w:tblPrChange (TablePropertyExceptionChange): remove change marker
    /// - w:trPr/w:ins (table row insertion): keep row, remove marker
    /// </summary>
    private int AcceptAllChanges()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;

        int count = 0;

        // Accept w:ins — unwrap (keep inner content)
        foreach (var ins in body.Descendants<InsertedRun>().ToList())
        {
            var parent = ins.Parent;
            if (parent == null) { ins.Remove(); count++; continue; }
            foreach (var child in ins.ChildElements.ToList())
                parent.InsertBefore(child.CloneNode(true), ins);
            ins.Remove();
            count++;
        }

        // Accept w:del — remove entirely (deletions are discarded)
        foreach (var del in body.Descendants<DeletedRun>().ToList())
        {
            del.Remove();
            count++;
        }

        // Accept w:rPrChange — remove the change element, keep current run properties
        foreach (var rPrChange in body.Descendants<RunPropertiesChange>().ToList())
        {
            rPrChange.Remove();
            count++;
        }

        // Accept w:pPrChange — remove the change element, keep current paragraph properties
        foreach (var pPrChange in body.Descendants<ParagraphPropertiesChange>().ToList())
        {
            pPrChange.Remove();
            count++;
        }

        // Accept w:sectPrChange — remove the change element
        foreach (var sectPrChange in body.Descendants<SectionPropertiesChange>().ToList())
        {
            sectPrChange.Remove();
            count++;
        }

        // Accept table property changes
        foreach (var tblPrChange in body.Descendants<TablePropertiesChange>().ToList())
        {
            tblPrChange.Remove();
            count++;
        }

        // Accept table row property changes (w:trPr containing w:ins)
        foreach (var trPr in body.Descendants<TableRowProperties>().ToList())
        {
            var trIns = trPr.GetFirstChild<InsertedRun>();
            if (trIns != null) { trIns.Remove(); count++; }
        }

        // Accept w:moveTo / w:moveFrom
        foreach (var moveFrom in body.Descendants<MoveFromRun>().ToList())
        {
            moveFrom.Remove();
            count++;
        }
        foreach (var moveTo in body.Descendants<MoveToRun>().ToList())
        {
            var parent = moveTo.Parent;
            if (parent == null) { moveTo.Remove(); count++; continue; }
            foreach (var child in moveTo.ChildElements.ToList())
                parent.InsertBefore(child.CloneNode(true), moveTo);
            moveTo.Remove();
            count++;
        }

        // Remove move range markers
        foreach (var marker in body.Descendants<MoveFromRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveFromRangeEnd>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeEnd>().ToList()) marker.Remove();

        _doc.MainDocumentPart?.Document?.Save();
        return count;
    }

    /// <summary>
    /// Reject all tracked changes in the document.
    /// - w:ins (InsertedRun): remove entire element (discard insertion)
    /// - w:del (DeletedRun): unwrap — restore content, convert w:delText to w:t
    /// - w:rPrChange: restore original formatting from inside the change element
    /// - w:pPrChange: restore original paragraph properties
    /// - w:sectPrChange: restore original section properties
    /// </summary>
    private int RejectAllChanges()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;

        int count = 0;

        // Reject w:ins — remove entirely (discard insertions)
        foreach (var ins in body.Descendants<InsertedRun>().ToList())
        {
            ins.Remove();
            count++;
        }

        // Reject w:del — unwrap, convert w:delText to w:t
        foreach (var del in body.Descendants<DeletedRun>().ToList())
        {
            var parent = del.Parent;
            if (parent == null) { del.Remove(); count++; continue; }
            foreach (var child in del.ChildElements.ToList())
            {
                var clone = child.CloneNode(true);
                // Convert DeletedText elements to Text elements
                foreach (var delText in clone.Descendants<DeletedText>().ToList())
                {
                    var text = new Text(delText.Text);
                    if (delText.Space != null)
                        text.Space = delText.Space;
                    delText.Parent?.ReplaceChild(text, delText);
                }
                parent.InsertBefore(clone, del);
            }
            del.Remove();
            count++;
        }

        // Reject w:rPrChange — restore original run properties
        foreach (var rPrChange in body.Descendants<RunPropertiesChange>().ToList())
        {
            var rPr = rPrChange.Parent as RunProperties;
            if (rPr != null)
            {
                var originalProps = rPrChange.GetFirstChild<PreviousRunProperties>();
                if (originalProps != null)
                {
                    // Replace current run properties with original ones
                    var run = rPr.Parent;
                    if (run != null)
                    {
                        var newRPr = new RunProperties();
                        foreach (var child in originalProps.ChildElements.ToList())
                            newRPr.AppendChild(child.CloneNode(true));
                        run.ReplaceChild(newRPr, rPr);
                    }
                }
                else
                {
                    rPrChange.Remove();
                }
            }
            else
            {
                rPrChange.Remove();
            }
            count++;
        }

        // Reject w:pPrChange — restore original paragraph properties
        foreach (var pPrChange in body.Descendants<ParagraphPropertiesChange>().ToList())
        {
            var pPr = pPrChange.Parent as ParagraphProperties;
            if (pPr != null)
            {
                var originalProps = pPrChange.GetFirstChild<PreviousParagraphProperties>();
                if (originalProps != null)
                {
                    var para = pPr.Parent;
                    if (para != null)
                    {
                        var newPPr = new ParagraphProperties();
                        foreach (var child in originalProps.ChildElements.ToList())
                            newPPr.AppendChild(child.CloneNode(true));
                        para.ReplaceChild(newPPr, pPr);
                    }
                }
                else
                {
                    pPrChange.Remove();
                }
            }
            else
            {
                pPrChange.Remove();
            }
            count++;
        }

        // Reject w:sectPrChange — restore original section properties
        foreach (var sectPrChange in body.Descendants<SectionPropertiesChange>().ToList())
        {
            sectPrChange.Remove();
            count++;
        }

        // Reject table property changes
        foreach (var tblPrChange in body.Descendants<TablePropertiesChange>().ToList())
        {
            tblPrChange.Remove();
            count++;
        }

        // Reject w:moveTo — remove (discard the move target)
        foreach (var moveTo in body.Descendants<MoveToRun>().ToList())
        {
            moveTo.Remove();
            count++;
        }
        // Reject w:moveFrom — unwrap (restore original position)
        foreach (var moveFrom in body.Descendants<MoveFromRun>().ToList())
        {
            var parent = moveFrom.Parent;
            if (parent == null) { moveFrom.Remove(); count++; continue; }
            foreach (var child in moveFrom.ChildElements.ToList())
                parent.InsertBefore(child.CloneNode(true), moveFrom);
            moveFrom.Remove();
            count++;
        }

        // Remove move range markers
        foreach (var marker in body.Descendants<MoveFromRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveFromRangeEnd>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeStart>().ToList()) marker.Remove();
        foreach (var marker in body.Descendants<MoveToRangeEnd>().ToList()) marker.Remove();

        _doc.MainDocumentPart?.Document?.Save();
        return count;
    }
}
