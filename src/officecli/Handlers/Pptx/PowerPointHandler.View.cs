// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null, string? range = null)
    {
        Core.ViewRangeGuard.RejectTextRange(range, "pptx");
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
                        var sizeStr = fontSize.HasValue ? $"{fontSize.Value / 100.0:0.##}pt" : ""; // hundredths-pt; keep .5 (e.g. 1850 -> 18.5pt)

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

    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null, string? range = null)
    {
        Core.ViewRangeGuard.RejectTextRange(range, "pptx");
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

        // Slide dimensions for the off-slide geometry check (16:9 default if unset).
        var slideSize = _doc.PresentationPart?.Presentation?.SlideSize;
        long slideW = (long)(slideSize?.Cx?.Value ?? 12192000);
        long slideH = (long)(slideSize?.Cy?.Value ?? 6858000);

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

            const long offTol = 180000; // 0.5cm in EMU — off-slide grazing tolerance
            // Pre-pass: collect off-slide TEXT boxes so a co-located graphic
            // container (a card background carrying no text of its own) can be
            // flagged too. Without this, a fix that only moves the flagged text
            // leaves the card behind — `view issues` passes but the card is
            // detached and still clipped (observed: a fix that satisfied the
            // checker yet broke the layout). Each entry: bounds + text snippet.
            var offSlideTextBoxes = new List<(long x, long y, long w, long h, string text)>();
            // allBoxes[i] aligns with shapes[i] (paint order = document order). Used
            // by the occlusion check: a text box covered by a LATER opaque box reads
            // as hidden. w==0 marks a shape with no usable geometry (never overlaps).
            var allBoxes = new List<(long x, long y, long w, long h, bool opaque)>();
            foreach (var s in shapes)
            {
                var sx = s.ShapeProperties?.Transform2D;
                bool hasGeom = sx?.Offset?.X != null && sx.Offset.Y != null && sx.Extents?.Cx != null && sx.Extents.Cy != null;
                long x = 0, y = 0, w = 0, h = 0;
                if (hasGeom) { x = sx!.Offset!.X!.Value; y = sx.Offset.Y!.Value; w = sx.Extents!.Cx!.Value; h = sx.Extents.Cy!.Value; }
                allBoxes.Add((x, y, w, h, hasGeom && IsOpaqueOccluder(s)));

                var st = GetShapeText(s);
                if (hasGeom && !string.IsNullOrWhiteSpace(st))
                {
                    long worst = Math.Max(Math.Max(-x, (x + w) - slideW), Math.Max(-y, (y + h) - slideH));
                    if (worst > offTol) offSlideTextBoxes.Add((x, y, w, h, st));
                }
            }

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

                // Off-slide: a shape whose declared box extends substantially
                // past a slide edge. Box geometry only (x/y/w/h) — a shape placed
                // off the canvas is a layout defect regardless of how text wraps.
                // Two emit cases:
                //   1. the shape carries text → its own content is clipped;
                //   2. the shape carries NO text but hosts an off-slide text box
                //      (a card background) → flag it so a fix moves the whole card.
                // Full-bleed pictures aren't in this Shape loop, and a spanning
                // decorative fill (extent ≈ slide) is skipped as legit bleed.
                var offText = GetShapeText(shape);
                var offXfrm = shape.ShapeProperties?.Transform2D;
                if (offXfrm?.Offset?.X != null && offXfrm.Offset.Y != null
                    && offXfrm.Extents?.Cx != null && offXfrm.Extents.Cy != null)
                {
                    long ox = offXfrm.Offset.X!.Value, oy = offXfrm.Offset.Y!.Value;
                    long ow = offXfrm.Extents.Cx!.Value, oh = offXfrm.Extents.Cy!.Value;
                    long offL = -ox, offT = -oy, offR = (ox + ow) - slideW, offB = (oy + oh) - slideH;
                    long offWorst = Math.Max(Math.Max(offL, offR), Math.Max(offT, offB));
                    if (offWorst > offTol)
                    {
                        string offEdge = offWorst == offR ? "right" : offWorst == offB ? "bottom"
                                       : offWorst == offL ? "left" : "top";
                        if (!string.IsNullOrWhiteSpace(offText))
                        {
                            // Name the clipped text so `view issues` is self-contained —
                            // the reader sees WHICH content runs off-canvas without a
                            // follow-up `get`. Whitespace-collapse and cap the snippet.
                            var offSnippet = System.Text.RegularExpressions.Regex.Replace(offText.Trim(), @"\s+", " ");
                            if (offSnippet.Length > 40) offSnippet = offSnippet[..40] + "…";
                            issues.Add(new DocumentIssue
                            {
                                Id = $"O{++issueNum}",
                                Type = IssueType.Format,
                                Severity = IssueSeverity.Warning,
                                Path = shapePath,
                                Message = $"Text shape \"{offSnippet}\" extends {offWorst / 360000.0:F1}cm past slide {offEdge} edge"
                            });
                        }
                        else
                        {
                            // Container case: a textless shape that runs off-slide
                            // AND hosts an off-slide text box. Skip spanning fills
                            // (full-width/height backdrops = legit bleed); require a
                            // hosted off-slide text box (rectangle intersection) so a
                            // bare decorative bleed never trips it.
                            bool spans = ow >= slideW - offTol || oh >= slideH - offTol;
                            var hosted = spans
                                ? default
                                : offSlideTextBoxes.FirstOrDefault(b =>
                                    b.x < ox + ow && b.x + b.w > ox && b.y < oy + oh && b.y + b.h > oy);
                            if (hosted.text != null)
                            {
                                var hostSnip = System.Text.RegularExpressions.Regex.Replace(hosted.text.Trim(), @"\s+", " ");
                                if (hostSnip.Length > 40) hostSnip = hostSnip[..40] + "…";
                                issues.Add(new DocumentIssue
                                {
                                    Id = $"O{++issueNum}",
                                    Type = IssueType.Format,
                                    Severity = IssueSeverity.Warning,
                                    Path = shapePath,
                                    Message = $"Container shape extends {offWorst / 360000.0:F1}cm past slide {offEdge} edge (clips \"{hostSnip}\")"
                                });
                            }
                        }
                    }
                }

                // Occlusion: a text shape whose box is substantially covered by a
                // LATER (higher z-order) opaque shape — the text is painted under an
                // opaque fill and reads as hidden. The shape's OWN background is
                // earlier in paint order, so it is never the culprit. Requires an
                // explicit opaque fill on the occluder and >20% area overlap, so a
                // text box correctly stacked over its own card, a translucent scrim,
                // or an incidental graze never trips it.
                if (!string.IsNullOrWhiteSpace(offText) && shapeIdx - 1 < allBoxes.Count)
                {
                    var ab = allBoxes[PathIndex.ToArrayIndex(shapeIdx)];
                    long aArea = ab.w * ab.h;
                    if (aArea > 0)
                    {
                        for (int j = shapeIdx; j < allBoxes.Count; j++)
                        {
                            var bb = allBoxes[j];
                            if (!bb.opaque) continue;
                            long ixw = Math.Min(ab.x + ab.w, bb.x + bb.w) - Math.Max(ab.x, bb.x);
                            long ixh = Math.Min(ab.y + ab.h, bb.y + bb.h) - Math.Max(ab.y, bb.y);
                            if (ixw <= 0 || ixh <= 0) continue;
                            if ((ixw * ixh) * 5 >= aArea) // ≥20% of the text box covered
                            {
                                var occPath = $"/slide[{slideNum}]/{BuildElementPathSegment("shape", shapes[j], j + 1)}";
                                var occSnip = System.Text.RegularExpressions.Regex.Replace(offText.Trim(), @"\s+", " ");
                                if (occSnip.Length > 40) occSnip = occSnip[..40] + "…";
                                issues.Add(new DocumentIssue
                                {
                                    Id = $"O{++issueNum}",
                                    Type = IssueType.Format,
                                    Severity = IssueSeverity.Warning,
                                    Path = shapePath,
                                    Message = $"Text \"{occSnip}\" is hidden behind overlapping shape {occPath}"
                                });
                                break;
                            }
                        }
                    }
                }

                // Low contrast: a shape with its OWN opaque dark solid fill that
                // carries opaque dark text reads fine on a laptop and vanishes on
                // projection. Declared-model only — the shape's explicit fill IS
                // the backdrop for its own runs, so there is no z-order guesswork
                // (unlike text over a fill=none box sitting on another shape). Scoped
                // tight to keep false positives near zero: explicit srgbClr on BOTH
                // fill and run (scheme / inherited colors skipped), translucent runs
                // skipped (intentional ghost / watermark text), and any color carrying
                // a +lumMod/+shade transform skipped (those shift brightness). Floor
                // matches the SKILL contrast rule: fill < 30%, text < 80% brightness.
                var fillSolid = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>();
                if (fillSolid != null
                    && TryOpaqueRgbLuminance(ReadColorFromFill(fillSolid), out double fillLum, out _)
                    && fillLum < 0.30 * 255)
                {
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        if (run.Text?.Text is not { Length: > 0 }) continue;
                        var runSolid = run.RunProperties?.GetFirstChild<Drawing.SolidFill>();
                        if (runSolid == null) continue;
                        if (TryOpaqueRgbLuminance(ReadColorFromFill(runSolid), out double runLum, out string runHex)
                            && runLum < 0.80 * 255)
                        {
                            issues.Add(new DocumentIssue
                            {
                                Id = $"C{++issueNum}",
                                Type = IssueType.Format,
                                Subtype = Core.IssueSubtypes.LowContrast,
                                Severity = IssueSeverity.Warning,
                                Path = shapePath,
                                Message = $"Low-contrast text #{runHex} on dark fill — unreadable on projection. "
                                        + "Use FFFFFF or a color brighter than 80%."
                            });
                            break; // one report per shape is enough
                        }
                    }
                }

                // Renderer-approximation audit: a text fill / warp can be stored
                // faithfully in OOXML yet NOT displayable by the HTML/SVG preview,
                // so a deck that "validates clean" still screens to a single color
                // in the agent's preview. This is the silent-acceptance gap P0-A
                // targets: the document is correct (PowerPoint renders it) but the
                // preview only approximates it. Surface each advanced text fill /
                // warp as a warning so the gap is observable in `view issues`
                // instead of shipping a deck whose gradient never shows up.
                string? warpApprox = GetTextWarpApproximation(shape);
                if (warpApprox != null)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"W{++issueNum}",
                        Type = IssueType.Format,
                        Subtype = Core.IssueSubtypes.TextWarpRendererApproximated,
                        Severity = IssueSeverity.Info,
                        Path = shapePath,
                        Message = warpApprox,
                        Suggestion = "PowerPoint renders the true warp; the preview shows only a text-warp marker class."
                    });
                }

                // One report per shape for text-fill approximation (mirrors the
                // LowContrast one-report-per-shape policy).
                string? fillApprox = shape.Descendants<Drawing.Run>()
                    .Select(TextFillApproximationReason)
                    .FirstOrDefault(r => r != null);
                if (fillApprox != null)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"T{++issueNum}",
                        Type = IssueType.Format,
                        Subtype = Core.IssueSubtypes.TextFillRendererApproximated,
                        Severity = IssueSeverity.Warning,
                        Path = shapePath,
                        Message = fillApprox,
                        Suggestion = "The stored OOXML is correct and PowerPoint will render it; run a screenshot to confirm the delivered look."
                    });
                }
                if (limit.HasValue && issues.Count >= limit.Value) break;
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

    /// <summary>
    /// Would this shape's fill hide whatever sits beneath it? True only for an
    /// EXPLICIT opaque fill — opaque solid (via <see cref="TryOpaqueRgbLuminance"/>),
    /// or a picture / pattern / gradient fill. Explicit no-fill, a translucent
    /// solid, and an absent (inherited) fill all return false so the occlusion
    /// check keeps its false-positive rate near zero.
    /// </summary>
    private static bool IsOpaqueOccluder(Shape s)
    {
        var sp = s.ShapeProperties;
        if (sp == null) return false;
        if (sp.GetFirstChild<Drawing.NoFill>() != null) return false;
        var solid = sp.GetFirstChild<Drawing.SolidFill>();
        if (solid != null) return TryOpaqueRgbLuminance(ReadColorFromFill(solid), out _, out _);
        if (sp.GetFirstChild<Drawing.BlipFill>() != null) return true;
        if (sp.GetFirstChild<Drawing.PatternFill>() != null) return true;
        if (sp.GetFirstChild<Drawing.GradientFill>() != null) return true;
        return false; // inherited / no explicit fill → conservatively not an occluder
    }

    /// <summary>
    /// Reason (or null) that a shape's text warp cannot be reproduced by the
    /// HTML/SVG preview. The preview only paints a text-warp marker class, not
    /// the per-glyph path PowerPoint computes, so a warped title previews as
    /// straight text. Returns null when textWarp is absent or 'none'.
    /// </summary>
    private static string? GetTextWarpApproximation(Shape s)
    {
        var warp = s.TextBody?.BodyProperties?.PresetTextWarp;
        if (warp == null) return null;
        // Match the existing readback pattern (HtmlPreview reads
        // PresetTextWarp.Preset.InnerText) instead of referencing the OOXML
        // enum by name — its exact C# type name varies across SDK versions.
        var prst = warp.Preset;
        if (prst?.HasValue != true) return null;
        if (string.Equals(prst.InnerText, "textNoShape", StringComparison.OrdinalIgnoreCase)) return null;
        return $"Text warp '{prst.InnerText}' is only marked in the preview — the HTML preview cannot warp glyphs like PowerPoint.";
    }

    /// <summary>
    /// Reason (or null) that a run's text fill is rendered only as a solid
    /// approximation by the HTML/SVG preview, even though the stored OOXML is
    /// correct. Returns null for the faithfully-rendered cases: no explicit
    /// fill, an un-transformed solid, or a simple two-stop un-transformed
    /// linear gradient. Flags image (blip) text fills, path gradients, and
    /// gradients with 3+ stops — the forms the CSS path reduces to a single
    /// color.
    /// </summary>
    private static string? TextFillApproximationReason(Drawing.Run run)
    {
        var rp = run.RunProperties;
        if (rp == null) return null;

        if (rp.GetFirstChild<Drawing.BlipFill>() != null)
            return "Image (blip) text fill — the preview falls back to a single color unless the embedded image can be resolved.";

        var grad = rp.GetFirstChild<Drawing.GradientFill>();
        if (grad == null) return null;

        if (grad.GetFirstChild<Drawing.PathGradientFill>() != null)
            return "Path gradient text fill — the preview cannot reproduce the path and shows a solid approximation.";

        var stops = grad.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stops != null && stops.Count > 2)
            return $"Gradient text fill with {stops.Count} stops — the preview shows a solid approximation.";
        return null;
    }

    /// <summary>
    /// Parse a <see cref="ReadColorFromFill"/> result into a perceived luminance
    /// (0–255, <c>r*0.299 + g*0.587 + b*0.114</c>) for the low-contrast check.
    /// Returns false — caller skips — for anything that is not a clean opaque
    /// explicit RGB: scheme/preset/system color names (non-hex), colors carrying
    /// a <c>+lumMod/+shade/…</c> transform suffix, and translucent colors
    /// (8-digit <c>RRGGBBAA</c> with alpha &lt; 80%). This keeps the check on
    /// firmly declared ground where the false-positive rate is near zero.
    /// </summary>
    private static bool TryOpaqueRgbLuminance(string? color, out double luminance, out string hex6)
    {
        luminance = 0;
        hex6 = "";
        if (string.IsNullOrEmpty(color)) return false;
        if (color.Contains('+')) return false;          // +lumMod/+shade transform — brightness shifts, skip
        var hex = color.TrimStart('#');
        if (hex.Length != 6 && hex.Length != 8) return false;   // scheme names / partial → skip
        foreach (var c in hex) if (!Uri.IsHexDigit(c)) return false;
        if (hex.Length == 8)
        {
            int a = Convert.ToInt32(hex.Substring(6, 2), 16); // CSS RRGGBBAA — alpha last
            if (a < 0.80 * 255) return false;            // translucent (ghost / watermark) — intentional, skip
            hex = hex.Substring(0, 6);
        }
        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
        luminance = r * 0.299 + g * 0.587 + b * 0.114;
        hex6 = hex.ToUpperInvariant();
        return true;
    }
}
