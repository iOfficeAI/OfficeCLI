// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
#pragma warning disable OOXML0001
using DocumentFormat.OpenXml.Experimental;
#pragma warning restore OOXML0001
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int slideNum = 0;
        int totalSlides = GetSlideParts().Count();

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            if (startLine.HasValue && slideNum < startLine.Value) continue;
            if (endLine.HasValue && slideNum > endLine.Value) break;

            if (maxLines.HasValue && slideNum - (startLine ?? 1) >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {maxLines.Value} of {totalSlides} slides, use --start/--end to see more)");
                break;
            }

            sb.AppendLine($"=== /slide[{slideNum}] ===");
            // CONSISTENCY(pptx-group-flatten): Descendants<Shape>() walks into
            // GroupShape children; Elements<Shape>() would drop them.
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            var shapes = shapeTree?.Descendants<Shape>() ?? Enumerable.Empty<Shape>();

            foreach (var shape in shapes)
            {
                var text = GetShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text))
                    sb.AppendLine(text);
            }

            // Table cell text — Descendants<Shape>() does not reach text inside
            // a:tbl cells (tables are GraphicFrame, not Shape). Emit each cell's
            // text on its own line so view text mirrors the slide's visible copy.
            if (shapeTree != null)
            {
                foreach (var table in shapeTree.Descendants<Drawing.Table>())
                {
                    foreach (var cell in table.Descendants<Drawing.TableCell>())
                    {
                        var cellText = GetCellTextWithParagraphBreaks(cell);
                        if (!string.IsNullOrWhiteSpace(cellText))
                            sb.AppendLine(cellText);
                    }
                }
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int slideNum = 0;
        int totalSlides = GetSlideParts().Count();

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            if (startLine.HasValue && slideNum < startLine.Value) continue;
            if (endLine.HasValue && slideNum > endLine.Value) break;

            if (maxLines.HasValue && slideNum - (startLine ?? 1) >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {maxLines.Value} of {totalSlides} slides, use --start/--end to see more)");
                break;
            }

            sb.AppendLine($"[/slide[{slideNum}]]");
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.ChildElements ?? Enumerable.Empty<OpenXmlElement>();

            RenderAnnotatedChildren(sb, shapes, indent: 1);
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private void RenderAnnotatedChildren(StringBuilder sb, IEnumerable<OpenXmlElement> children, int indent)
    {
        var pad = new string(' ', indent * 2);
        foreach (var child in children)
        {
            if (child is Shape shape)
            {
                var mathElements = FindShapeMathElements(shape);
                if (mathElements.Count > 0)
                {
                    var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                    var text = GetShapeText(shape);
                    var hasOtherText = shape.TextBody?.Elements<Drawing.Paragraph>()
                        .SelectMany(p => p.Elements<Drawing.Run>())
                        .Any(r => !string.IsNullOrWhiteSpace(r.Text?.Text)) == true;
                    if (hasOtherText)
                        sb.AppendLine($"{pad}[Text Box] \"{text}\" \u2190 contains equation: \"{latex}\"");
                    else
                        sb.AppendLine($"{pad}[Equation] \"{latex}\"");
                }
                else
                {
                    var text = GetShapeText(shape);
                    var type = IsTitle(shape) ? "Title" : "Text Box";

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
                        var font = firstRun?.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                            ?? firstRun?.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface
                            ?? "(default)";
                        var fontSize = firstRun?.RunProperties?.FontSize?.Value;
                        var sizeStr = fontSize.HasValue ? $"{fontSize.Value / 100}pt" : "";

                        sb.AppendLine($"{pad}[{type}] \"{text}\" \u2190 {font} {sizeStr}");
                    }
                }
            }
            else if (child is GraphicFrame gf && gf.Descendants<Drawing.Table>().Any())
            {
                var table = gf.Descendants<Drawing.Table>().First();
                var tblRows = table.Elements<Drawing.TableRow>().Count();
                var tblCols = table.Elements<Drawing.TableRow>().FirstOrDefault()?.Elements<Drawing.TableCell>().Count() ?? 0;
                var tblName = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";
                sb.AppendLine($"{pad}[Table] \"{tblName}\" \u2190 {tblRows}x{tblCols}");
            }
            else if (child is Picture pic)
            {
                var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
                var altText = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                var altInfo = string.IsNullOrEmpty(altText) ? "\u26a0 no alt text" : $"alt=\"{altText}\"";
                sb.AppendLine($"{pad}[Picture] \"{name}\" \u2190 {altInfo}");
            }
            else if (child is GroupShape group)
            {
                var groupName = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
                var nested = group.ChildElements
                    .Where(e => e is Shape || e is GroupShape || e is Picture || e is GraphicFrame || e is ConnectionShape)
                    .ToList();
                sb.AppendLine($"{pad}[Group] \"{groupName}\" \u2190 {nested.Count} item(s)");
                RenderAnnotatedChildren(sb, nested, indent + 1);
            }
        }
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        sb.AppendLine($"File: {Path.GetFileName(_filePath)} | {slideParts.Count} slides");

        int slideNum = 0;
        foreach (var slidePart in slideParts)
        {
            slideNum++;
            // CONSISTENCY(pptx-group-flatten)
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            var shapes = shapeTree?.Descendants<Shape>() ?? Enumerable.Empty<Shape>();

            var title = shapes.Where(IsTitle).Select(GetShapeText).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t)) ?? "(untitled)";

            int textBoxes = shapes.Count(s => !IsTitle(s) && !string.IsNullOrWhiteSpace(GetShapeText(s)));
            int pictures = shapeTree?.Descendants<Picture>().Count() ?? 0;
            int oleObjects = CountSlideOleObjects(slidePart);

            var details = new List<string>();
            if (textBoxes > 0) details.Add($"{textBoxes} text box(es)");
            if (pictures > 0) details.Add($"{pictures} picture(s)");
            if (oleObjects > 0) details.Add($"{oleObjects} ole object(s)");

            var detailStr = details.Count > 0 ? $" - {string.Join(", ", details)}" : "";
            sb.AppendLine($"\u251c\u2500\u2500 Slide {slideNum}: \"{title}\"{detailStr}");
        }

        return sb.ToString().TrimEnd();
    }

    // CONSISTENCY(ole-stats): per-slide OLE counter shared by outline and
    // outlineJson. Same dedup rule as ViewAsStats — shapeTree oleObject
    // elements count once, orphan embedded/package parts add extras.
    private int CountSlideOleObjects(SlidePart slidePart)
    {
        int count = 0;
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (shapeTree != null)
        {
            foreach (var oleEl in shapeTree.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>())
            {
                count++;
                if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                    referenced.Add(rid);
            }
        }
        count += slidePart.EmbeddedObjectParts.Count(p => !referenced.Contains(slidePart.GetIdOfPart(p)));
        count += slidePart.EmbeddedPackageParts.Count(p => !referenced.Contains(slidePart.GetIdOfPart(p)));
        return count;
    }

    public string ViewAsStats()
    {
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        int totalShapes = 0;
        int totalPictures = 0;
        int totalTextBoxes = 0;
        int totalWords = 0;
        int totalCharts = 0;
        int slidesWithoutTitle = 0;
        int picturesWithoutAlt = 0;
        var fontCounts = new Dictionary<string, int>();

        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            // CONSISTENCY(pptx-group-flatten): include shapes/pictures/charts
            // nested inside GroupShape.
            var shapes = shapeTree.Descendants<Shape>().ToList();
            var pictures = shapeTree.Descendants<Picture>().ToList();
            // CONSISTENCY(stats-chart-count): charts live in GraphicFrame elements
            // alongside tables; surface them as a separate Charts row so the totals
            // visibly account for chart shapes.
            totalCharts += shapeTree.Descendants<GraphicFrame>()
                .Count(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any()
                          || IsExtendedChartFrame(gf));
            // CONSISTENCY(stats-table-count): tables are GraphicFrame too — pre-R5
            // they vanished from totalShapes entirely and cell text was never
            // word-counted, so a deck whose only content was a 5x5 grid reported
            // "0 shapes / 0 words".
            var tables = shapeTree.Descendants<Drawing.Table>().ToList();
            totalShapes += shapes.Count + tables.Count;
            totalPictures += pictures.Count;
            totalTextBoxes += shapes.Count(s => !IsTitle(s));

            if (!shapes.Any(IsTitle))
                slidesWithoutTitle++;

            picturesWithoutAlt += pictures.Count(p =>
                string.IsNullOrEmpty(p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value));

            // Count words from shape text
            foreach (var shape in shapes)
            {
                var text = GetShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text))
                    totalWords += text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            }

            // Count words from table cells (mirror the ViewAsText walk).
            foreach (var table in tables)
            {
                foreach (var cell in table.Descendants<Drawing.TableCell>())
                {
                    var cellText = GetCellTextWithParagraphBreaks(cell);
                    if (!string.IsNullOrWhiteSpace(cellText))
                        totalWords += cellText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
                }
            }

            // Collect font usage
            foreach (var shape in shapes)
            {
                foreach (var run in shape.Descendants<Drawing.Run>())
                {
                    var font = run.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                        ?? run.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
                    if (font != null)
                        fontCounts[font!] = fontCounts.GetValueOrDefault(font!) + 1;
                }
            }
        }

        // OLE count = oleObj elements + any orphan embedded parts not
        // referenced by one. Mirrors how CollectOleNodesForSlide builds
        // its result so summary == visible query rows.
        int totalOleObjects = 0;
        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (shapeTree != null)
            {
                foreach (var oleEl in shapeTree.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>())
                {
                    totalOleObjects++;
                    if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                        referenced.Add(rid);
                }
            }
            totalOleObjects += slidePart.EmbeddedObjectParts.Count(p => !referenced.Contains(slidePart.GetIdOfPart(p)));
            totalOleObjects += slidePart.EmbeddedPackageParts.Count(p => !referenced.Contains(slidePart.GetIdOfPart(p)));
        }

        sb.AppendLine($"Slides: {slideParts.Count}");
        sb.AppendLine($"Total shapes: {totalShapes}");
        sb.AppendLine($"Text boxes: {totalTextBoxes}");
        sb.AppendLine($"Pictures: {totalPictures}");
        if (totalCharts > 0) sb.AppendLine($"Charts: {totalCharts}");
        if (totalOleObjects > 0) sb.AppendLine($"OLE Objects: {totalOleObjects}");
        sb.AppendLine($"Words: {totalWords}");
        sb.AppendLine($"Slides without title: {slidesWithoutTitle}");
        sb.AppendLine($"Pictures without alt text: {picturesWithoutAlt}");

        if (fontCounts.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Font usage:");
            foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
                sb.AppendLine($"  {font}: {count} occurrence(s)");
        }

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsStatsJson()
    {
        var slideParts = GetSlideParts().ToList();

        int totalShapes = 0, totalPictures = 0, totalTextBoxes = 0, totalWords = 0, totalCharts = 0;
        int slidesWithoutTitle = 0, picturesWithoutAlt = 0;
        var fontCounts = new Dictionary<string, int>();

        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            // CONSISTENCY(pptx-group-flatten)
            var shapes = shapeTree.Descendants<Shape>().ToList();
            var pictures = shapeTree.Descendants<Picture>().ToList();
            // CONSISTENCY(stats-chart-count): see ViewAsStats.
            totalCharts += shapeTree.Descendants<GraphicFrame>()
                .Count(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any()
                          || IsExtendedChartFrame(gf));
            // CONSISTENCY(stats-table-count): see ViewAsStats.
            var tables = shapeTree.Descendants<Drawing.Table>().ToList();
            totalShapes += shapes.Count + tables.Count;
            totalPictures += pictures.Count;
            totalTextBoxes += shapes.Count(s => !IsTitle(s));

            if (!shapes.Any(IsTitle)) slidesWithoutTitle++;
            picturesWithoutAlt += pictures.Count(p =>
                string.IsNullOrEmpty(p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value));

            foreach (var shape in shapes)
            {
                var text = GetShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text))
                    totalWords += text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;

                foreach (var run in shape.Descendants<Drawing.Run>())
                {
                    var font = run.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                        ?? run.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
                    if (font != null)
                        fontCounts[font!] = fontCounts.GetValueOrDefault(font!) + 1;
                }
            }

            foreach (var table in tables)
            {
                foreach (var cell in table.Descendants<Drawing.TableCell>())
                {
                    var cellText = GetCellTextWithParagraphBreaks(cell);
                    if (!string.IsNullOrWhiteSpace(cellText))
                        totalWords += cellText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
                }
            }
        }

        // Mirror the same OLE counting logic as ViewAsStats.
        int jsonOleObjects = 0;
        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (shapeTree != null)
            {
                foreach (var oleEl in shapeTree.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>())
                {
                    jsonOleObjects++;
                    if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                        referenced.Add(rid);
                }
            }
            jsonOleObjects += slidePart.EmbeddedObjectParts.Count(p => !referenced.Contains(slidePart.GetIdOfPart(p)));
            jsonOleObjects += slidePart.EmbeddedPackageParts.Count(p => !referenced.Contains(slidePart.GetIdOfPart(p)));
        }

        var result = new JsonObject
        {
            ["slides"] = slideParts.Count,
            ["totalShapes"] = totalShapes,
            ["textBoxes"] = totalTextBoxes,
            ["pictures"] = totalPictures,
            ["charts"] = totalCharts,
            ["oleObjects"] = jsonOleObjects,
            ["words"] = totalWords,
            ["slidesWithoutTitle"] = slidesWithoutTitle,
            ["picturesWithoutAlt"] = picturesWithoutAlt
        };

        if (fontCounts.Count > 0)
        {
            var fonts = new JsonObject();
            foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
                fonts[font] = count;
            result["fontUsage"] = fonts;
        }

        return result;
    }

    public JsonNode ViewAsOutlineJson()
    {
        var slideParts = GetSlideParts().ToList();
        var slidesArray = new JsonArray();

        int slideNum = 0;
        foreach (var slidePart in slideParts)
        {
            slideNum++;
            // CONSISTENCY(pptx-group-flatten)
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            var shapes = shapeTree?.Descendants<Shape>() ?? Enumerable.Empty<Shape>();
            var title = shapes.Where(IsTitle).Select(GetShapeText).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
            int textBoxes = shapes.Count(s => !IsTitle(s) && !string.IsNullOrWhiteSpace(GetShapeText(s)));
            int pictures = shapeTree?.Descendants<Picture>().Count() ?? 0;

            int oleObjects = CountSlideOleObjects(slidePart);
            var slide = new JsonObject
            {
                ["index"] = slideNum,
                ["title"] = title,
                ["textBoxes"] = textBoxes,
                ["pictures"] = pictures,
                ["oleObjects"] = oleObjects
            };
            slidesArray.Add((JsonNode)slide);
        }

        return new JsonObject
        {
            ["fileName"] = Path.GetFileName(_filePath),
            ["totalSlides"] = slideParts.Count,
            ["slides"] = slidesArray
        };
    }

    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var slidesArray = new JsonArray();
        int slideNum = 0;
        int totalSlides = GetSlideParts().Count();

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            if (startLine.HasValue && slideNum < startLine.Value) continue;
            if (endLine.HasValue && slideNum > endLine.Value) break;

            if (maxLines.HasValue && slidesArray.Count >= maxLines.Value)
                break;

            var textsArray = new JsonArray();
            // CONSISTENCY(pptx-group-flatten)
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Descendants<Shape>() ?? Enumerable.Empty<Shape>();
            foreach (var shape in shapes)
            {
                var text = GetShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text))
                    textsArray.Add((JsonNode)text);
            }

            var slide = new JsonObject
            {
                ["index"] = slideNum,
                ["path"] = $"/slide[{slideNum}]",
                ["texts"] = textsArray
            };
            slidesArray.Add((JsonNode)slide);
        }

        return new JsonObject
        {
            ["totalSlides"] = totalSlides,
            ["slides"] = slidesArray
        };
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueNum = 0;
        int slideNum = 0;

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            // Broken part-relationship detection. A slide's .rels file can
            // point r:id="rN" at /ppt/media/NONEXISTENT.png (or similar) when
            // a deck was hand-edited, partial-extracted, or assembled by a
            // buggy authoring tool. The Open XML SDK loads the relationship
            // but throws when any reader (NodeBuilder picture, ResolveChart,
            // HTML preview) calls GetPartById and the package can't resolve
            // the target. Surface this as a stable, filterable issue so the
            // gap is observable BEFORE callers hit a failure path —
            // mirrors xlsx's definedname_broken / chart_series_ref_missing_sheet
            // observability for "rot in the package" rather than deferred
            // evaluation.
            // A slide's .rels file can point r:id="rN" Target="/ppt/media/NONEXISTENT.png"
            // (or similar) when a deck was hand-edited, partial-extracted, or
            // assembled by a buggy authoring tool. The Open XML SDK silently
            // skips such relationships during enumeration of slidePart.Parts
            // (the target part isn't constructed), but downstream readers
            // (NodeBuilder/ResolveChart/HtmlPreview) that walk the slide XML
            // and call slidePart.GetPartById(rId) raise
            // ArgumentOutOfRangeException "Part: <uri> doesn't exist in the
            // package." with no rId / slide context — opaque to agents and
            // unfilterable in `view issues`.
            //
            // Diff slide-XML referenced rIds against the SDK-loaded ones to
            // surface the gap as a stable broken_part_ref subtype BEFORE
            // any reader hits the failure path. Mirrors xlsx's
            // definedname_broken observability for "rot in the package".
            // Build the set of package Uris and the slide's relationship
            // table by reading the rels XML directly. The SDK's
            // slidePart.Parts iteration silently skips a relationship
            // whose Target points to a missing part — and may also drop
            // valid siblings under the same rels file once the broken one
            // is encountered. Reading rels XML + comparing Target to
            // package.GetParts() Uri gives a reader-faithful gap report
            // without relying on the SDK's tolerant enumeration.
            var brokenRelIds = new Dictionary<string, string>(StringComparer.Ordinal);
            // R8-2: also flag rIds the slide XML references that the rels
            // file doesn't declare at all. The original scan covered
            // declared-but-dangling rels (Target points to a missing part),
            // but a slide carrying r:embed="rId99" with no matching
            // <Relationship Id="rId99"> in the rels XML produced a silent
            // ArgumentException at render time instead of a stable issue.
            var declaredRelIds = new HashSet<string>(StringComparer.Ordinal);
            try
            {
#pragma warning disable OOXML0001
                var pkg = slidePart.OpenXmlPackage.GetPackage();
#pragma warning restore OOXML0001
                var packageUris = new HashSet<string>(
                    pkg.GetParts().Select(p => p.Uri.OriginalString.TrimStart('/')),
                    StringComparer.OrdinalIgnoreCase);
                // Read the slide's .rels file directly — IPackage's
                // Relationships collection silently drops relationships
                // whose Target points to a missing part, so a Parts-vs-rels
                // diff via the SDK's API misses the very rot we want to
                // surface. Reading the rels XML preserves every declared
                // relationship, including dangling ones.
                var slideUriPath = slidePart.Uri.OriginalString.TrimStart('/');
                var lastSlash = slideUriPath.LastIndexOf('/');
                var relPartUri = lastSlash >= 0
                    ? "/" + slideUriPath.Substring(0, lastSlash) + "/_rels/" + slideUriPath.Substring(lastSlash + 1) + ".rels"
                    : "/_rels/" + slideUriPath + ".rels";
                var relsUri = new Uri(relPartUri, UriKind.Relative);
                if (pkg.PartExists(relsUri))
                {
                    var relsPart = pkg.GetPart(relsUri);
                    using var rs = relsPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
                    var relsDoc = System.Xml.Linq.XDocument.Load(rs);
                    System.Xml.Linq.XNamespace relsNs2006 = "http://schemas.openxmlformats.org/package/2006/relationships";
                    foreach (var relEl in relsDoc.Descendants(relsNs2006 + "Relationship"))
                    {
                        var targetMode = (string?)relEl.Attribute("TargetMode") ?? "Internal";
                        if (!string.Equals(targetMode, "Internal", StringComparison.OrdinalIgnoreCase)) continue;
                        var id = (string?)relEl.Attribute("Id");
                        var target = (string?)relEl.Attribute("Target");
                        if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target)) continue;
                        declaredRelIds.Add(id);
                        Uri resolved;
                        try
                        {
                            resolved = System.IO.Packaging.PackUriHelper.ResolvePartUri(
                                slidePart.Uri, new Uri(target, UriKind.RelativeOrAbsolute));
                        }
                        catch { continue; }
                        var key = resolved.OriginalString.TrimStart('/');
                        if (!packageUris.Contains(key))
                            brokenRelIds[id] = key;
                    }
                }
            }
            catch { /* probe is best-effort; never break view issues */ }

            // Resident-mode parity: when a chart/picture/etc. is added via the
            // long-lived resident, the new IdPartPair lives in the SDK's
            // in-memory part tree before the rels XML stream gets flushed to
            // the package. The disk-rels XML probe above then doesn't see the
            // newly-declared rId, and view issues falsely reports
            // broken_part_ref. Mirror the in-memory state by also accepting
            // any rId the SDK currently exposes via Parts /
            // HyperlinkRelationships / ExternalRelationships /
            // DataPartReferenceRelationships. Disk-only state still flags
            // genuinely-broken rels because brokenRelIds is populated from
            // the disk probe above.
            try
            {
                foreach (var pair in slidePart.Parts)
                    if (!string.IsNullOrEmpty(pair.RelationshipId))
                    {
                        declaredRelIds.Add(pair.RelationshipId);
                        // The part exists in-memory under this rId — clear
                        // any stale "broken" verdict the disk probe inferred
                        // before the SDK flushed the new IdPartPair.
                        brokenRelIds.Remove(pair.RelationshipId);
                    }
                foreach (var hr in slidePart.HyperlinkRelationships)
                    if (!string.IsNullOrEmpty(hr.Id)) declaredRelIds.Add(hr.Id);
                foreach (var er in slidePart.ExternalRelationships)
                    if (!string.IsNullOrEmpty(er.Id)) declaredRelIds.Add(er.Id);
                foreach (var dr in slidePart.DataPartReferenceRelationships)
                    if (!string.IsNullOrEmpty(dr.Id)) declaredRelIds.Add(dr.Id);
            }
            catch { /* probe is best-effort */ }

            const string relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            // R8-2: walk the raw slide XML, not the SDK's parsed Slide. The
            // SDK silently drops elements whose rels don't resolve (e.g. a
            // <p:pic> with r:embed pointing at an undeclared rId), so the
            // SDK-loaded view shows nothing to flag. Reading the part stream
            // exposes every r:embed/r:link/r:id the file actually carries.
            var rawRefs = new List<(string ElemName, string AttrLocal, string RId)>();
            try
            {
                using var slideStream = slidePart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
                var slideDoc = System.Xml.Linq.XDocument.Load(slideStream);
                System.Xml.Linq.XNamespace rNs = relNs;
                foreach (var el in slideDoc.Descendants())
                {
                    foreach (var name in new[] { "id", "embed", "link" })
                    {
                        var a = el.Attribute(rNs + name);
                        if (a == null || string.IsNullOrEmpty(a.Value)) continue;
                        rawRefs.Add((el.Name.LocalName, name, a.Value));
                    }
                }
            }
            catch { /* probe is best-effort */ }

            var reportedRids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (elemName, attrLocal, rId) in rawRefs)
            {
                if (string.IsNullOrEmpty(rId)) continue;
                string message;
                string suggestion;
                if (brokenRelIds.TryGetValue(rId, out var missingTarget))
                {
                    message = $"Slide rel '{rId}' on <{elemName}> targets missing part '/{missingTarget}'";
                    suggestion = "Restore the missing target part, or remove the dangling relationship and the referencing element.";
                }
                else if (!declaredRelIds.Contains(rId))
                {
                    // R8-2: rId referenced by slide XML but not declared in
                    // the rels file at all. Distinct diagnostic so users can
                    // tell "I need to add the rel" apart from "I need to
                    // restore the part".
                    message = $"Slide rel '{rId}' on <{elemName}> is referenced but not declared in the slide rels";
                    suggestion = $"Add a matching <Relationship Id=\"{rId}\"> to the slide's rels file, or drop the referencing element.";
                }
                else continue;
                if (!reportedRids.Add(rId)) continue;
                issues.Add(new DocumentIssue
                {
                    Id = $"R{++issueNum}",
                    Type = IssueType.Structure,
                    Subtype = Core.IssueSubtypes.BrokenPartRef,
                    Severity = IssueSeverity.Error,
                    Path = $"/slide[{slideNum}]",
                    Message = message,
                    Context = $"<{elemName} r:{attrLocal}=\"{rId}\">",
                    Suggestion = suggestion
                });
            }

            var shapes = shapeTree.Elements<Shape>().ToList();
            // No "slide has no title" warning: PowerPoint does not require a
            // Title placeholder, free-form / chart-only / image-only slides
            // are legitimate, and the check has no actionable subtype or
            // suggestion. A dedicated a11y audit mode is the right home for
            // this kind of structural lint if it returns later.

            // Check for font consistency issues
            int shapeIdx = 0;
            foreach (var shape in shapes)
            {
                shapeIdx++;
                var shapePath = $"/slide[{slideNum}]/{BuildElementPathSegment("shape", shape, shapeIdx)}";

                // CONSISTENCY(text-overflow-check): merged in from former `check` command.
                var overflow = CheckTextOverflow(shape);
                if (overflow != null)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"O{++issueNum}",
                        Type = IssueType.Format,
                        Severity = IssueSeverity.Warning,
                        Path = shapePath,
                        Message = overflow
                    });
                }

                var runs = shape.Descendants<Drawing.Run>().ToList();
                if (runs.Count <= 1) continue;

                var fonts = runs.Select(r =>
                    r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface)
                    .Where(f => f != null).Distinct().ToList();

                if (fonts.Count > 1)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"F{++issueNum}",
                        Type = IssueType.Format,
                        Severity = IssueSeverity.Info,
                        Path = shapePath,
                        Message = $"Inconsistent fonts in text box: {string.Join(", ", fonts)}"
                    });
                }
            }

            // Table row/grid width mismatch. OOXML requires each <a:tr> to have
            // <a:tc> count equal to its parent <a:tblGrid>'s <a:gridCol> count
            // (gridSpan creates wide cells via hMerge placeholders, not by
            // dropping cells). A short row produces undefined render — real
            // PowerPoint stretches it with empty stubs, other viewers may
            // skip it; either way it's malformed. Common cause: an earlier
            // `add row --prop cols=N` with N < grid count, now rejected at
            // write time but old files may still carry the bug.
            int tableIdx = 0;
            foreach (var graphicFrame in shapeTree.Descendants<GraphicFrame>())
            {
                var table = graphicFrame.Descendants<Drawing.Table>().FirstOrDefault();
                if (table == null) continue;
                tableIdx++;
                var gridColCount = table.TableGrid?.Elements<Drawing.GridColumn>().Count() ?? 0;
                if (gridColCount == 0) continue;
                int rowIdx = 0;
                foreach (var tr in table.Elements<Drawing.TableRow>())
                {
                    rowIdx++;
                    var tcCount = tr.Elements<Drawing.TableCell>().Count();
                    if (tcCount != gridColCount)
                    {
                        issues.Add(new DocumentIssue
                        {
                            Id = $"S{++issueNum}",
                            Type = IssueType.Structure,
                            Severity = IssueSeverity.Warning,
                            Path = $"/slide[{slideNum}]/table[{tableIdx}]/tr[{rowIdx}]",
                            Message = $"Row has {tcCount} cell(s) but table grid has {gridColCount} column(s). " +
                                      "Real PowerPoint silently pads/clips; other viewers may render incorrectly. " +
                                      "Add empty cells or use gridSpan to merge."
                        });
                    }
                }
            }

            // CONSISTENCY(pptx-group-flatten): alt-text accessibility check
            // applies to every picture on the slide, including those nested in
            // groups.
            foreach (var pic in shapeTree.Descendants<Picture>())
            {
                var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                if (string.IsNullOrEmpty(alt))
                {
                    var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";
                    issues.Add(new DocumentIssue
                    {
                        Id = $"F{++issueNum}",
                        Type = IssueType.Format,
                        Severity = IssueSeverity.Info,
                        Path = $"/slide[{slideNum}]",
                        Message = $"Picture \"{name}\" is missing alt text (accessibility issue)"
                    });
                }
            }

            // Slide-level a:fld fields (slidenum, datetime1/2/3/..., footer, header)
            // without a cached rendered text — same observability pattern as Word
            // complex fields and xlsx unevaluated formulas. PowerPoint populates
            // <a:t> inside <a:fld> when it renders the slide; a fresh authoring
            // pass with no PPT open-and-save leaves the slot blank, and the
            // slide silently shows nothing where the slide number / date should
            // have been. Issue lets agents detect the gap before the file ships.
            // Walk slide AND its layout AND master — slide-number / date-time
            // placeholders almost always live in layout/master (cSld inherits),
            // so scanning only the slide's ShapeTree misses the most common
            // shape of the bug.
            var allTrees = new List<(string Scope, DocumentFormat.OpenXml.OpenXmlElement Tree)>
            {
                ("slide", shapeTree),
            };
            if (slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree is { } layoutTree)
                allTrees.Add(("layout", layoutTree));
            if (slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree is { } masterTree)
                allTrees.Add(("master", masterTree));
            foreach (var (scope, tree) in allTrees)
            {
                foreach (var fld in tree.Descendants<Drawing.Field>())
                {
                    if (limit.HasValue && issues.Count >= limit.Value) break;
                    var fldType = fld.Type?.Value ?? "";
                    if (!IsDynamicSlideFieldType(fldType)) continue;
                    var cachedText = string.Concat(fld.Elements<Drawing.Text>().Select(t => t.Text));
                    if (!string.IsNullOrEmpty(cachedText)) continue;
                    issues.Add(new DocumentIssue
                    {
                        Id = $"U{++issueNum}",
                        Type = IssueType.Content,
                        Subtype = Core.IssueSubtypes.SlideFieldNotEvaluated,
                        Severity = IssueSeverity.Warning,
                        Path = scope == "slide"
                            ? $"/slide[{slideNum}]"
                            : $"/slide[{slideNum}] ({scope})",
                        Message = "Slide field written but not evaluated (no cached text, PowerPoint has not rendered it)",
                        Context = $"<a:fld type=\"{fldType}\"> in {scope}",
                        Suggestion = "Open in PowerPoint once so <a:t> inside <a:fld> is populated."
                    });
                }
                if (limit.HasValue && issues.Count >= limit.Value) break;
            }

            if (limit.HasValue && issues.Count >= limit.Value) break;
        }

        // Subtype / type filter. pptx previously ignored issueType entirely.
        // Accept both broad bucket (format/content/structure) and specific
        // subtype identifiers.
        if (issueType != null)
        {
            var bucket = issueType.ToLowerInvariant() switch
            {
                "format" or "f" => IssueType.Format,
                "content" or "c" => IssueType.Content,
                "structure" or "s" => IssueType.Structure,
                _ => (IssueType?)null
            };
            if (bucket.HasValue)
                issues = issues.Where(i => i.Type == bucket.Value).ToList();
            else
                issues = issues.Where(i => string.Equals(i.Subtype, issueType, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        return issues;
    }

    // IsDynamicSlideFieldType has been collapsed into Helpers.cs's
    // IsDynamicSlideFieldTypeStatic — single source of truth.
    private static bool IsDynamicSlideFieldType(string fldType) => IsDynamicSlideFieldTypeStatic(fldType);
}
