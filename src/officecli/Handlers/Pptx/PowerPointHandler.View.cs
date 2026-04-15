// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();

            foreach (var shape in shapes)
            {
                var text = GetShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text))
                    sb.AppendLine(text);
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

            int shapeIdx = 0;
            foreach (var child in shapes)
            {
                if (child is Shape shape)
                {
                    // Check if shape contains equations
                    var mathElements = FindShapeMathElements(shape);
                    if (mathElements.Count > 0)
                    {
                        var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                        var text = GetShapeText(shape);
                        // Check for text runs NOT inside mc:Fallback
                        var hasOtherText = shape.TextBody?.Elements<Drawing.Paragraph>()
                            .SelectMany(p => p.Elements<Drawing.Run>())
                            .Any(r => !string.IsNullOrWhiteSpace(r.Text?.Text)) == true;
                        if (hasOtherText)
                            sb.AppendLine($"  [Text Box] \"{text}\" \u2190 contains equation: \"{latex}\"");
                        else
                            sb.AppendLine($"  [Equation] \"{latex}\"");
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

                            sb.AppendLine($"  [{type}] \"{text}\" \u2190 {font} {sizeStr}");
                        }
                    }
                    shapeIdx++;
                }
                else if (child is GraphicFrame gf && gf.Descendants<Drawing.Table>().Any())
                {
                    var table = gf.Descendants<Drawing.Table>().First();
                    var tblRows = table.Elements<Drawing.TableRow>().Count();
                    var tblCols = table.Elements<Drawing.TableRow>().FirstOrDefault()?.Elements<Drawing.TableCell>().Count() ?? 0;
                    var tblName = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";
                    sb.AppendLine($"  [Table] \"{tblName}\" \u2190 {tblRows}x{tblCols}");
                }
                else if (child is Picture pic)
                {
                    var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
                    var altText = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                    var altInfo = string.IsNullOrEmpty(altText) ? "\u26a0 no alt text" : $"alt=\"{altText}\"";
                    sb.AppendLine($"  [Picture] \"{name}\" \u2190 {altInfo}");
                }
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
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
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();

            var title = shapes.Where(IsTitle).Select(GetShapeText).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t)) ?? "(untitled)";

            int textBoxes = shapes.Count(s => !IsTitle(s) && !string.IsNullOrWhiteSpace(GetShapeText(s)));
            int pictures = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Picture>().Count() ?? 0;
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
        int slidesWithoutTitle = 0;
        int picturesWithoutAlt = 0;
        var fontCounts = new Dictionary<string, int>();

        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            var shapes = shapeTree.Elements<Shape>().ToList();
            var pictures = shapeTree.Elements<Picture>().ToList();
            totalShapes += shapes.Count;
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

        int totalShapes = 0, totalPictures = 0, totalTextBoxes = 0, totalWords = 0;
        int slidesWithoutTitle = 0, picturesWithoutAlt = 0;
        var fontCounts = new Dictionary<string, int>();

        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            var shapes = shapeTree.Elements<Shape>().ToList();
            var pictures = shapeTree.Elements<Picture>().ToList();
            totalShapes += shapes.Count;
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
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();
            var title = shapes.Where(IsTitle).Select(GetShapeText).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
            int textBoxes = shapes.Count(s => !IsTitle(s) && !string.IsNullOrWhiteSpace(GetShapeText(s)));
            int pictures = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Picture>().Count() ?? 0;

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
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();
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

            var shapes = shapeTree.Elements<Shape>().ToList();
            if (!shapes.Any(IsTitle))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"S{++issueNum}",
                    Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/slide[{slideNum}]",
                    Message = "Slide has no title"
                });
            }

            // Check for font consistency issues
            int shapeIdx = 0;
            foreach (var shape in shapes)
            {
                var runs = shape.Descendants<Drawing.Run>().ToList();
                if (runs.Count <= 1) { shapeIdx++; continue; }

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
                        Path = $"/slide[{slideNum}]/{BuildElementPathSegment("shape", shape, shapeIdx + 1)}",
                        Message = $"Inconsistent fonts in text box: {string.Join(", ", fonts)}"
                    });
                }
                shapeIdx++;
            }

            foreach (var pic in shapeTree.Elements<Picture>())
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

            if (limit.HasValue && issues.Count >= limit.Value) break;
        }

        return issues;
    }
}
