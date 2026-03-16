// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler : IDocumentHandler
{
    private readonly PresentationDocument _doc;
    private readonly string _filePath;

    public PowerPointHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = PresentationDocument.Open(filePath, editable);
    }

    private (SlidePart slidePart, Shape shape) ResolveShape(int slideIdx, int shapeIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var shapes = shapeTree.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > shapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found");

        return (slidePart, shapes[shapeIdx - 1]);
    }

    private (SlidePart slidePart, GraphicFrame gf, ChartPart chartPart) ResolveChart(int slideIdx, int chartIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var chartFrames = shapeTree.Elements<GraphicFrame>()
            .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any())
            .ToList();
        if (chartIdx < 1 || chartIdx > chartFrames.Count)
            throw new ArgumentException($"Chart {chartIdx} not found (total: {chartFrames.Count})");

        var gf = chartFrames[chartIdx - 1];
        var chartRef = gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().First();
        var chartPart = (ChartPart)slidePart.GetPartById(chartRef.Id!.Value!);
        return (slidePart, gf, chartPart);
    }

    private (SlidePart slidePart, Drawing.Table table) ResolveTable(int slideIdx, int tblIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var tables = shapeTree.Elements<GraphicFrame>()
            .Select(gf => gf.Descendants<Drawing.Table>().FirstOrDefault())
            .Where(t => t != null).ToList();
        if (tblIdx < 1 || tblIdx > tables.Count)
            throw new ArgumentException($"Table {tblIdx} not found (total: {tables.Count})");

        return (slidePart, tables[tblIdx - 1]!);
    }

    /// <summary>
    /// Resolve a logical PPT path (e.g. /slide[1]/table[1]/tr[2]) to the actual OpenXML element.
    /// Returns null if the path doesn't contain logical segments that need resolving.
    /// </summary>
    private (SlidePart slidePart, OpenXmlElement element)? ResolveLogicalPath(string path)
    {
        // /slide[N]/table[M]...
        var tblPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\](.*)$");
        if (tblPathMatch.Success)
        {
            var slideIdx = int.Parse(tblPathMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblPathMatch.Groups[2].Value);
            var rest = tblPathMatch.Groups[3].Value; // e.g. /tr[1]/tc[2]/txBody

            var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
            OpenXmlElement current = table;

            if (!string.IsNullOrEmpty(rest))
            {
                var segments = GenericXmlQuery.ParsePathSegments(rest);
                var target = GenericXmlQuery.NavigateByPath(current, segments);
                if (target != null) current = target;
                else throw new ArgumentException($"Element not found: {path}");
            }
            return (slidePart, current);
        }

        // /slide[N]/placeholder[X]...
        var phPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\](.*)$");
        if (phPathMatch.Success)
        {
            var slideIdx = int.Parse(phPathMatch.Groups[1].Value);
            var phId = phPathMatch.Groups[2].Value;
            var rest = phPathMatch.Groups[3].Value;

            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");
            var slidePart = slideParts[slideIdx - 1];
            OpenXmlElement current = ResolvePlaceholderShape(slidePart, phId);

            if (!string.IsNullOrEmpty(rest))
            {
                var segments = GenericXmlQuery.ParsePathSegments(rest);
                var target = GenericXmlQuery.NavigateByPath(current, segments);
                if (target != null) current = target;
                else throw new ArgumentException($"Element not found: {path}");
            }
            return (slidePart, current);
        }

        return null;
    }

    private static PlaceholderValues? ParsePlaceholderType(string name)
    {
        return name.ToLowerInvariant() switch
        {
            "title" => PlaceholderValues.Title,
            "centertitle" or "centeredtitle" or "ctitle" => PlaceholderValues.CenteredTitle,
            "body" or "content" => PlaceholderValues.Body,
            "subtitle" or "sub" => PlaceholderValues.SubTitle,
            "date" or "datetime" or "dt" => PlaceholderValues.DateAndTime,
            "footer" => PlaceholderValues.Footer,
            "slidenum" or "slidenumber" or "sldnum" => PlaceholderValues.SlideNumber,
            "object" or "obj" => PlaceholderValues.Object,
            "chart" => PlaceholderValues.Chart,
            "table" => PlaceholderValues.Table,
            "clipart" => PlaceholderValues.ClipArt,
            "diagram" or "dgm" => PlaceholderValues.Diagram,
            "media" => PlaceholderValues.Media,
            "picture" or "pic" => PlaceholderValues.Picture,
            "header" => PlaceholderValues.Header,
            _ => null
        };
    }

    private Shape ResolvePlaceholderShape(SlidePart slidePart, string phId)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");

        // Try numeric index first
        if (int.TryParse(phId, out var numIdx))
        {
            // Match by placeholder index
            var byIndex = shapeTree.Elements<Shape>()
                .FirstOrDefault(s =>
                {
                    var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    return ph?.Index?.Value == (uint)numIdx;
                });
            if (byIndex != null) return byIndex;

            // Also try as 1-based ordinal of all placeholders
            var allPh = shapeTree.Elements<Shape>()
                .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>() != null).ToList();
            if (numIdx >= 1 && numIdx <= allPh.Count)
                return allPh[numIdx - 1];

            throw new ArgumentException($"Placeholder index {numIdx} not found");
        }

        // Try by type name
        var phType = ParsePlaceholderType(phId)
            ?? throw new ArgumentException($"Unknown placeholder type: '{phId}'. " +
                "Known types: title, body, subtitle, date, footer, slidenum, object, picture, centerTitle");

        var byType = shapeTree.Elements<Shape>()
            .FirstOrDefault(s =>
            {
                var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                return ph?.Type?.Value == phType;
            });

        if (byType != null) return byType;

        // Check layout for inherited placeholders and create one on the slide
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree != null)
        {
            var layoutShape = layoutPart.SlideLayout.CommonSlideData.ShapeTree.Elements<Shape>()
                .FirstOrDefault(s =>
                {
                    var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    return ph?.Type?.Value == phType;
                });

            if (layoutShape != null)
            {
                // Clone from layout and add to slide
                var newShape = (Shape)layoutShape.CloneNode(true);
                // Clear any text content from layout placeholder
                if (newShape.TextBody != null)
                {
                    newShape.TextBody.RemoveAllChildren<Drawing.Paragraph>();
                    newShape.TextBody.Append(new Drawing.Paragraph(
                        new Drawing.EndParagraphRunProperties { Language = "zh-CN" }));
                }
                shapeTree.AppendChild(newShape);
                return newShape;
            }
        }

        throw new ArgumentException($"Placeholder '{phId}' not found on slide or its layout");
    }

    private DocumentNode GetPlaceholderNode(SlidePart slidePart, int slideIdx, int phIdx, int depth)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");

        // Get all placeholders on slide
        var placeholders = shapeTree.Elements<Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>() != null).ToList();

        if (phIdx < 1 || phIdx > placeholders.Count)
            throw new ArgumentException($"Placeholder {phIdx} not found (total: {placeholders.Count})");

        var shape = placeholders[phIdx - 1];
        var ph = shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<PlaceholderShape>()!;

        var node = ShapeToNode(shape, slideIdx, phIdx, depth);
        node.Path = $"/slide[{slideIdx}]/placeholder[{phIdx}]";
        node.Type = "placeholder";
        if (ph.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
        if (ph.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
        return node;
    }
    // ==================== Layout ====================

    /// <summary>
    /// Resolve a SlideLayoutPart by name, type, or index.
    /// If layoutHint is null, returns the first layout.
    /// Matching order: exact name → layout type → numeric index → first layout.
    /// </summary>
    private static SlideLayoutPart? ResolveSlideLayout(PresentationPart presentationPart, string? layoutHint)
    {
        var allLayouts = presentationPart.SlideMasterParts
            .SelectMany(m => m.SlideLayoutParts).ToList();
        if (allLayouts.Count == 0) return null;

        if (string.IsNullOrEmpty(layoutHint))
            return allLayouts.FirstOrDefault();

        // 1. Match by layout name (CommonSlideData.Name or SlideLayout.MatchingName)
        var byName = allLayouts.FirstOrDefault(lp =>
        {
            var sl = lp.SlideLayout;
            var csdName = sl?.CommonSlideData?.Name?.Value;
            var matchName = sl?.MatchingName?.Value;
            return string.Equals(csdName, layoutHint, StringComparison.OrdinalIgnoreCase)
                || string.Equals(matchName, layoutHint, StringComparison.OrdinalIgnoreCase);
        });
        if (byName != null) return byName;

        // 2. Match by layout type keyword
        var layoutType = layoutHint.ToLowerInvariant() switch
        {
            "title"                                     => SlideLayoutValues.Title,
            "titleonly" or "title_only"                  => SlideLayoutValues.TitleOnly,
            "blank"                                     => SlideLayoutValues.Blank,
            "twocontent" or "two_content" or "twocol"   => SlideLayoutValues.TwoColumnText,
            "titlecontent" or "title_content"            => SlideLayoutValues.ObjectText,
            "section" or "sectionheader"                 => SlideLayoutValues.SectionHeader,
            "comparison"                                 => SlideLayoutValues.TwoTextAndTwoObjects,
            "contentwithcaption" or "caption"            => SlideLayoutValues.ObjectAndText,
            "picturewithcaption" or "pictxt"             => SlideLayoutValues.PictureText,
            "custom"                                     => SlideLayoutValues.Custom,
            _ => (SlideLayoutValues?)null
        };
        if (layoutType.HasValue)
        {
            var byType = allLayouts.FirstOrDefault(lp =>
                lp.SlideLayout?.Type?.HasValue == true &&
                lp.SlideLayout.Type.Value == layoutType.Value);
            if (byType != null) return byType;
        }

        // 3. Match by 1-based numeric index
        if (int.TryParse(layoutHint, out var idx) && idx >= 1 && idx <= allLayouts.Count)
            return allLayouts[idx - 1];

        // 4. Fuzzy match: layout name contains the hint (case-insensitive)
        var fuzzy = allLayouts.FirstOrDefault(lp =>
        {
            var csdName = lp.SlideLayout?.CommonSlideData?.Name?.Value;
            return csdName != null && csdName.Contains(layoutHint, StringComparison.OrdinalIgnoreCase);
        });
        if (fuzzy != null) return fuzzy;

        throw new ArgumentException(
            $"Layout '{layoutHint}' not found. Available layouts: " +
            string.Join(", ", allLayouts.Select((lp, i) =>
            {
                var name = lp.SlideLayout?.CommonSlideData?.Name?.Value ?? "(unnamed)";
                var type = lp.SlideLayout?.Type?.HasValue == true ? lp.SlideLayout.Type.InnerText : "?";
                return $"[{i + 1}] {name} ({type})";
            })));
    }

    /// <summary>
    /// Get the layout name for a slide part.
    /// Falls back to type name if no explicit name is set.
    /// </summary>
    private static string? GetSlideLayoutName(SlidePart slidePart)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart?.SlideLayout == null) return null;
        return layoutPart.SlideLayout.CommonSlideData?.Name?.Value
            ?? layoutPart.SlideLayout.MatchingName?.Value
            ?? (layoutPart.SlideLayout.Type?.HasValue == true
                ? layoutPart.SlideLayout.Type.InnerText : null);
    }

    /// <summary>
    /// Get the layout type for a slide part.
    /// </summary>
    private static string? GetSlideLayoutType(SlidePart slidePart)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart?.SlideLayout?.Type?.HasValue != true) return null;
        return layoutPart.SlideLayout.Type.InnerText;
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == "/" || partPath == "/presentation")
            return _doc.PresentationPart?.Presentation?.OuterXml ?? "(empty)";

        var match = Regex.Match(partPath, @"^/slide\[(\d+)\]$");
        if (match.Success)
        {
            var idx = int.Parse(match.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx >= 1 && idx <= slideParts.Count)
                return GetSlide(slideParts[idx - 1]).OuterXml;
            return $"(slide[{idx}] not found)";
        }

        return $"Unknown part: {partPath}. Available: /presentation, /slide[N]";
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        OpenXmlPartRootElement rootElement;

        if (partPath is "/" or "/presentation")
        {
            rootElement = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
        }
        else if (Regex.Match(partPath, @"^/slide\[(\d+)\]$") is { Success: true } slideMatch)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found");
            rootElement = GetSlide(slideParts[idx - 1]);
        }
        else if (Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$") is { Success: true } masterMatch)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"SlideMaster {idx} not found");
            rootElement = masters[idx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }
        else if (Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$") is { Success: true } layoutMatch)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1 || idx > layouts.Count)
                throw new ArgumentException($"SlideLayout {idx} not found");
            rootElement = layouts[idx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }
        else if (Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$") is { Success: true } noteMatch)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            rootElement = notesPart.NotesSlide
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }
        else
        {
            throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N]");
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        Console.WriteLine($"raw-set: {affected} element(s) affected");
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a SlidePart
                var slideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException(
                        "Chart must be added under a slide: add-part <file> '/slide[N]' --type chart");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide index {slideIdx} out of range");

                var slidePart = slideParts[slideIdx - 1];
                var chartPart = slidePart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                var relId = slidePart.GetIdOfPart(chartPart);

                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = slidePart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/slide[{slideIdx}]/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose() => _doc.Dispose();

    // ==================== Private Helpers ====================

    private static Slide GetSlide(SlidePart part) =>
        part.Slide ?? throw new InvalidOperationException("Corrupt file: slide data missing");

    private IEnumerable<SlidePart> GetSlideParts()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var slideIdList = presentation?.GetFirstChild<SlideIdList>();
        if (slideIdList == null) yield break;

        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            var relId = slideId.RelationshipId?.Value;
            if (relId == null) continue;
            yield return (SlidePart)_doc.PresentationPart!.GetPartById(relId);
        }
    }

}
