// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// Return the 1-based positional index of the shape with the given OOXML
    /// id on the slide at <paramref name="slideIdx"/>, or null if none matches.
    /// Counts the same element types ResolveShape() exposes (plain <p:sp>),
    /// matching what the /slide[N]/shape[K] positional path resolves to.
    /// </summary>
    internal int? ResolveShapeOrdinalById(int slideIdx, uint id)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return null;
        var shapeTree = GetSlide(slideParts[slideIdx - 1]).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return null;
        var shapes = shapeTree.Elements<Shape>().ToList();
        for (int i = 0; i < shapes.Count; i++)
        {
            var sid = shapes[i].NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (sid == id) return i + 1;
        }
        return null;
    }

    private (SlidePart slidePart, Shape shape) ResolveShape(int slideIdx, int shapeIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var shapes = shapeTree.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > shapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found");

        return (slidePart, shapes[shapeIdx - 1]);
    }

    private (SlidePart slidePart, GraphicFrame gf, ChartPart? chartPart, ExtendedChartPart? extChartPart) ResolveChart(int slideIdx, int chartIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var chartFrames = shapeTree.Elements<GraphicFrame>()
            .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any()
                || IsExtendedChartFrame(gf))
            .ToList();
        if (chartIdx < 1 || chartIdx > chartFrames.Count)
            throw new ArgumentException($"Chart {chartIdx} not found (total: {chartFrames.Count})");

        var gf = chartFrames[chartIdx - 1];

        // Regular c:chart reference
        var chartRef = gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
        ChartPart? chartPart = null;
        if (chartRef?.Id?.Value != null)
        {
            // Broken c:chart/@r:id (relationship missing from the slide part —
            // happens after a hand-edited zip or a partially-imported deck) makes
            // GetPartById throw the SDK's bare ArgumentOutOfRangeException
            // ("Specified argument was out of the range of valid values.") with
            // no rId context. Surface a CliException with a stable code and the
            // offending rId so callers can route to repair instead of guessing.
            try
            {
                chartPart = (ChartPart)slidePart.GetPartById(chartRef.Id.Value);
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new CliException(
                    $"Chart relationship '{chartRef.Id.Value}' on slide {slideIdx} points to a missing part. The chart's r:id has no matching relationship in the slide's rels file.")
                    { Code = "broken_chart_relationship" };
            }
        }

        // cx:chart (extended) reference — note: the SDK has TWO classes that
        // both serialize with LocalName "chart":
        //   CX.RelId  — the reference stub inside a:graphicData (has r:id)
        //   CX.Chart  — the content element inside cx:chartSpace (has plotArea)
        // Loaded elements may pick the "wrong" CLR type, so Descendants<CX.RelId>()
        // can miss them. Walk graphic → graphicData and grab the first child
        // matching the cx namespace + "chart" local name instead.
        ExtendedChartPart? extChartPart = null;
        var graphicData = gf.Graphic?.GraphicData;
        if (graphicData != null)
        {
            const string cxNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
            var cxChartRef = graphicData.ChildElements
                .FirstOrDefault(e => e.LocalName == "chart" && e.NamespaceUri == cxNs);
            if (cxChartRef != null)
            {
                // The r:id attribute lives in the relationships namespace.
                const string rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var relIdAttr = cxChartRef.GetAttributes()
                    .FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri == rNs);
                if (!string.IsNullOrEmpty(relIdAttr.Value))
                {
                    try
                    {
                        extChartPart = (ExtendedChartPart)slidePart.GetPartById(relIdAttr.Value);
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        throw new CliException(
                            $"Extended chart relationship '{relIdAttr.Value}' on slide {slideIdx} points to a missing part.")
                            { Code = "broken_chart_relationship" };
                    }
                }
            }
        }

        return (slidePart, gf, chartPart, extChartPart);
    }

    private (SlidePart slidePart, Drawing.Table table) ResolveTable(int slideIdx, int tblIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

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
                else throw new ArgumentException($"Element not found: {path}. Resolved table[{tblIdx}] on slide[{slideIdx}] but sub-path '{rest}' does not exist. Available children: {DescribeChildren(current)}");
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
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
            var slidePart = slideParts[slideIdx - 1];
            OpenXmlElement current = ResolvePlaceholderShape(slidePart, phId);

            if (!string.IsNullOrEmpty(rest))
            {
                var segments = GenericXmlQuery.ParsePathSegments(rest);
                var target = GenericXmlQuery.NavigateByPath(current, segments);
                if (target != null) current = target;
                else throw new ArgumentException($"Element not found: {path}. Resolved placeholder[{phId}] on slide[{slideIdx}] but sub-path '{rest}' does not exist. Available children: {DescribeChildren(current)}");
            }
            return (slidePart, current);
        }

        return null;
    }

    /// <summary>Summarize child element types for error messages.</summary>
    private static string DescribeChildren(OpenXmlElement parent)
    {
        var groups = parent.ChildElements
            .GroupBy(e => e.LocalName)
            .Select(g => g.Count() > 1 ? $"{g.Key}[1..{g.Count()}]" : g.Key)
            .Take(10)
            .ToList();
        return groups.Count > 0 ? string.Join(", ", groups) : "(empty)";
    }

    /// <summary>Summarize slide contents for error messages (e.g. "3 shapes, 1 table, 2 pictures").</summary>
    private static string DescribeSlideInventory(ShapeTree? shapeTree)
    {
        if (shapeTree == null) return "(empty slide)";
        var parts = new List<string>();
        var shapes = shapeTree.Elements<Shape>().Count();
        var tables = shapeTree.Elements<GraphicFrame>().Count(gf => gf.Descendants<Drawing.Table>().Any());
        var charts = shapeTree.Elements<GraphicFrame>().Count(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any());
        var pics = shapeTree.Elements<Picture>().Count();
        var connectors = shapeTree.Elements<ConnectionShape>().Count();
        var groups = shapeTree.Elements<GroupShape>().Count();
        if (shapes > 0) parts.Add($"{shapes} shape(s)");
        if (tables > 0) parts.Add($"{tables} table(s)");
        if (charts > 0) parts.Add($"{charts} chart(s)");
        if (pics > 0) parts.Add($"{pics} picture(s)");
        if (connectors > 0) parts.Add($"{connectors} connector(s)");
        if (groups > 0) parts.Add($"{groups} group(s)");
        return parts.Count > 0 ? string.Join(", ", parts) : "(empty slide)";
    }

    // Inverse of ParsePlaceholderType — picks the canonical human-readable
    // alias so Get's Format["phType"] round-trips through Add's phType prop.
    // Null/absent <p:ph type=…> defaults to "body" per ECMA-376 (§19.7.10
    // says omitting the attr is equivalent to type="body"). The Title vs
    // CenteredTitle distinction is preserved.
    private static string? FormatPlaceholderType(PlaceholderValues? value)
    {
        if (value == null) return "body";
        if (value.Value == PlaceholderValues.Title) return "title";
        if (value.Value == PlaceholderValues.CenteredTitle) return "ctrTitle";
        if (value.Value == PlaceholderValues.Body) return "body";
        if (value.Value == PlaceholderValues.SubTitle) return "subtitle";
        if (value.Value == PlaceholderValues.DateAndTime) return "date";
        if (value.Value == PlaceholderValues.Footer) return "footer";
        if (value.Value == PlaceholderValues.SlideNumber) return "slidenum";
        if (value.Value == PlaceholderValues.Header) return "header";
        if (value.Value == PlaceholderValues.Object) return "obj";
        if (value.Value == PlaceholderValues.Chart) return "chart";
        if (value.Value == PlaceholderValues.Table) return "table";
        if (value.Value == PlaceholderValues.ClipArt) return "clipart";
        if (value.Value == PlaceholderValues.Diagram) return "diagram";
        if (value.Value == PlaceholderValues.Media) return "media";
        if (value.Value == PlaceholderValues.Picture) return "picture";
        return value.Value.ToString();
    }

    private static PlaceholderValues? ParsePlaceholderType(string name)
    {
        return name.ToLowerInvariant() switch
        {
            "title" => PlaceholderValues.Title,
            // 'ctrTitle' is the OOXML serialization (ECMA-376 §19.7.10);
            // accept it alongside the human-readable aliases so the
            // type-name returned by query placeholder round-trips.
            "centertitle" or "centeredtitle" or "ctitle" or "ctrtitle" => PlaceholderValues.CenteredTitle,
            "body" or "content" => PlaceholderValues.Body,
            "subtitle" or "sub" or "subtitlepres" => PlaceholderValues.SubTitle,
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
                        new Drawing.EndParagraphRunProperties { Language = "en-US" }));
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

    // ==================== Media Timing Lookup ====================

    /// <summary>
    /// Find the CommonMediaNode in the timing tree for a given shape ID.
    /// </summary>
    private static CommonMediaNode? FindMediaTimingNode(SlidePart slidePart, uint shapeId)
    {
        var timing = GetSlide(slidePart).GetFirstChild<Timing>();
        if (timing == null) return null;

        foreach (var mediaNode in timing.Descendants<CommonMediaNode>())
        {
            var target = mediaNode.TargetElement?.GetFirstChild<ShapeTarget>();
            if (target?.ShapeId?.Value == shapeId.ToString())
                return mediaNode;
        }
        return null;
    }

    // ==================== Cleanup (POI-style reference counting) ====================

    /// <summary>
    /// Remove a Picture element with proper cleanup of relationships and media parts.
    /// Follows Apache POI's pattern: reference-count blipIds, only delete parts when
    /// no other shapes reference the same media.
    /// </summary>
    private static void RemovePictureWithCleanup(SlidePart slidePart, ShapeTree shapeTree, Picture pic)
    {
        // Collect all relationship IDs referenced by this picture
        var relIdsToClean = new HashSet<string>();

        // BlipFill → Blip.Embed (poster/image)
        var blipEmbed = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>()?.Embed?.Value;
        if (blipEmbed != null) relIdsToClean.Add(blipEmbed);

        // VideoFromFile.Link or AudioFromFile.Link
        var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        var videoLink = nvPr?.GetFirstChild<DocumentFormat.OpenXml.Drawing.VideoFromFile>()?.Link?.Value;
        if (videoLink != null) relIdsToClean.Add(videoLink);
        var audioLink = nvPr?.GetFirstChild<DocumentFormat.OpenXml.Drawing.AudioFromFile>()?.Link?.Value;
        if (audioLink != null) relIdsToClean.Add(audioLink);

        // p14:media.Embed (MediaReferenceRelationship)
        var p14Media = nvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
        var mediaEmbed = p14Media?.Embed?.Value;
        if (mediaEmbed != null) relIdsToClean.Add(mediaEmbed);

        // Reference count: check all OTHER pictures on the same slide for shared relIds
        var sharedRelIds = new HashSet<string>();
        foreach (var otherPic in shapeTree.Elements<Picture>())
        {
            if (otherPic == pic) continue; // skip the one being removed

            var otherBlip = otherPic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>()?.Embed?.Value;
            if (otherBlip != null && relIdsToClean.Contains(otherBlip)) sharedRelIds.Add(otherBlip);

            var otherNvPr = otherPic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
            var otherVid = otherNvPr?.GetFirstChild<DocumentFormat.OpenXml.Drawing.VideoFromFile>()?.Link?.Value;
            if (otherVid != null && relIdsToClean.Contains(otherVid)) sharedRelIds.Add(otherVid);
            var otherAud = otherNvPr?.GetFirstChild<DocumentFormat.OpenXml.Drawing.AudioFromFile>()?.Link?.Value;
            if (otherAud != null && relIdsToClean.Contains(otherAud)) sharedRelIds.Add(otherAud);

            var otherMedia = otherNvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault()?.Embed?.Value;
            if (otherMedia != null && relIdsToClean.Contains(otherMedia)) sharedRelIds.Add(otherMedia);
        }

        // Remove the XML element first
        pic.Remove();

        // Clean up relationships that are no longer referenced
        foreach (var relId in relIdsToClean)
        {
            if (sharedRelIds.Contains(relId)) continue; // still referenced by another shape

            try { slidePart.DeletePart(relId); } catch (ArgumentException) { }
            // Also try removing data part relationships (video/audio/media)
            try
            {
                foreach (var dpr in slidePart.DataPartReferenceRelationships.Where(r => r.Id == relId).ToList())
                    slidePart.DeleteReferenceRelationship(dpr);
            }
            catch (ArgumentException) { }
        }
    }

    // ==================== Layout ====================

    /// <summary>
    /// Resolve a SlideLayoutPart by name, type token, or numeric index. Single
    /// entry point for the layout-selection grammar used by both Add (new slide)
    /// and Set (re-layout existing slide). If layoutHint is null/empty, returns
    /// the first layout. Matching order:
    ///   1. exact display name (CommonSlideData.Name) or MatchingName
    ///   2. layout type token — raw OOXML enum InnerText (e.g. "objTx", "blank",
    ///      "title", "twoObj") AND friendly aliases ("titlecontent",
    ///      "twocontent", "section", …)
    ///   3. 1-based numeric index across all masters
    ///   4. case-insensitive substring match on display name
    /// Throws ArgumentException with a unified available-list string on miss.
    /// </summary>
    internal static SlideLayoutPart? ResolveSlideLayout(PresentationPart presentationPart, string? layoutHint)
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

        // 2a. Match by raw OOXML enum InnerText (e.g. "objTx", "blank", "title")
        // — what Get emits as Format["layoutType"], so it round-trips.
        var byRawType = allLayouts.FirstOrDefault(lp =>
            lp.SlideLayout?.Type?.HasValue == true &&
            string.Equals(lp.SlideLayout.Type.InnerText, layoutHint, StringComparison.OrdinalIgnoreCase));
        if (byRawType != null) return byRawType;

        // 2b. Match by friendly layout type alias
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
            FormatAvailableLayouts(allLayouts));
    }

    /// <summary>
    /// Unified available-layouts list used in Add/Set error messages so the
    /// grammar surface (name | type | index) is discoverable from either path.
    /// </summary>
    internal static string FormatAvailableLayouts(IEnumerable<SlideLayoutPart> allLayouts)
        => string.Join(", ", allLayouts.Select((lp, i) =>
        {
            var name = lp.SlideLayout?.CommonSlideData?.Name?.Value ?? "(unnamed)";
            var type = lp.SlideLayout?.Type?.HasValue == true ? lp.SlideLayout.Type.InnerText : "?";
            return $"[{i + 1}] {name} ({type})";
        }));

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
}
