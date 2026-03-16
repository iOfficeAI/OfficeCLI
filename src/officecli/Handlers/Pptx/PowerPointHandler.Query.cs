// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
        {
            var node = new DocumentNode { Path = "/", Type = "presentation" };
            int slideNum = 0;
            foreach (var slidePart in GetSlideParts())
            {
                slideNum++;
                var title = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)";

                var slideNode = new DocumentNode
                {
                    Path = $"/slide[{slideNum}]",
                    Type = "slide",
                    Preview = title
                };
                var lName = GetSlideLayoutName(slidePart);
                if (lName != null) slideNode.Format["layout"] = lName;
                ReadSlideBackground(GetSlide(slidePart), slideNode);
                ReadSlideTransition(GetSlide(slidePart), slideNode);

                if (depth > 0)
                {
                    slideNode.Children = GetSlideChildNodes(slidePart, slideNum, depth - 1);
                    slideNode.ChildCount = slideNode.Children.Count;
                }
                else
                {
                    var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                    slideNode.ChildCount = (shapeTree?.Elements<Shape>().Count() ?? 0)
                        + (shapeTree?.Elements<Picture>().Count() ?? 0);
                }

                node.Children.Add(slideNode);
            }
            node.ChildCount = node.Children.Count;
            return node;
        }

        // Try notes path: /slide[N]/notes
        var notesGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesGetMatch.Success)
        {
            var notesSlideIdx = int.Parse(notesGetMatch.Groups[1].Value);
            var slidePartsN = GetSlideParts().ToList();
            if (notesSlideIdx < 1 || notesSlideIdx > slidePartsN.Count)
                throw new ArgumentException($"Slide {notesSlideIdx} not found");
            var slidePartN = slidePartsN[notesSlideIdx - 1];
            var notesText = slidePartN.NotesSlidePart != null
                ? GetNotesText(slidePartN.NotesSlidePart) : "";
            return new DocumentNode { Path = path, Type = "notes", Text = notesText };
        }

        // Try paragraph/run paths: /slide[N]/shape[M]/paragraph[P] or .../run[K] or .../paragraph[P]/run[K]
        var runPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/run\[(\d+)\]$");
        if (runPathMatch.Success)
        {
            var sIdx = int.Parse(runPathMatch.Groups[1].Value);
            var shIdx = int.Parse(runPathMatch.Groups[2].Value);
            var rIdx = int.Parse(runPathMatch.Groups[3].Value);
            var (runSlidePart, shape) = ResolveShape(sIdx, shIdx);
            var allRuns = GetAllRuns(shape);
            if (rIdx < 1 || rIdx > allRuns.Count)
                throw new ArgumentException($"Run {rIdx} not found (shape has {allRuns.Count} runs)");
            return RunToNode(allRuns[rIdx - 1], path, runSlidePart);
        }

        var paraPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\](?:/run\[(\d+)\])?$");
        if (paraPathMatch.Success)
        {
            var sIdx = int.Parse(paraPathMatch.Groups[1].Value);
            var shIdx = int.Parse(paraPathMatch.Groups[2].Value);
            var pIdx = int.Parse(paraPathMatch.Groups[3].Value);
            var (paraSlidePart, shape) = ResolveShape(sIdx, shIdx);
            var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                ?? throw new ArgumentException("Shape has no text body");
            if (pIdx < 1 || pIdx > paragraphs.Count)
                throw new ArgumentException($"Paragraph {pIdx} not found (shape has {paragraphs.Count} paragraphs)");

            var para = paragraphs[pIdx - 1];

            if (paraPathMatch.Groups[4].Success)
            {
                // /slide[N]/shape[M]/paragraph[P]/run[K]
                var rIdx = int.Parse(paraPathMatch.Groups[4].Value);
                var paraRuns = para.Elements<Drawing.Run>().ToList();
                if (rIdx < 1 || rIdx > paraRuns.Count)
                    throw new ArgumentException($"Run {rIdx} not found (paragraph has {paraRuns.Count} runs)");
                return RunToNode(paraRuns[rIdx - 1],
                    $"/slide[{sIdx}]/shape[{shIdx}]/paragraph[{pIdx}]/run[{rIdx}]", paraSlidePart);
            }

            // /slide[N]/shape[M]/paragraph[P]
            var paraText = string.Join("", para.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? ""));
            var paraNode = new DocumentNode
            {
                Path = path,
                Type = "paragraph",
                Text = paraText
            };
            var align = para.ParagraphProperties?.Alignment;
            if (align != null && align.HasValue) paraNode.Format["align"] = align.InnerText;

            var runs = para.Elements<Drawing.Run>().ToList();
            paraNode.ChildCount = runs.Count;
            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in runs)
                {
                    paraNode.Children.Add(RunToNode(run,
                        $"/slide[{sIdx}]/shape[{shIdx}]/paragraph[{pIdx}]/run[{runIdx + 1}]", paraSlidePart));
                    runIdx++;
                }
            }
            return paraNode;
        }

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellGetMatch.Success)
        {
            var sIdx = int.Parse(tblCellGetMatch.Groups[1].Value);
            var tIdx = int.Parse(tblCellGetMatch.Groups[2].Value);
            var rIdx = int.Parse(tblCellGetMatch.Groups[3].Value);
            var cIdx = int.Parse(tblCellGetMatch.Groups[4].Value);

            var (slidePart2, table) = ResolveTable(sIdx, tIdx);
            var tableRows = table.Elements<Drawing.TableRow>().ToList();
            if (rIdx < 1 || rIdx > tableRows.Count)
                throw new ArgumentException($"Row {rIdx} not found (table has {tableRows.Count} rows)");
            var cells = tableRows[rIdx - 1].Elements<Drawing.TableCell>().ToList();
            if (cIdx < 1 || cIdx > cells.Count)
                throw new ArgumentException($"Cell {cIdx} not found (row has {cells.Count} cells)");

            var cell = cells[cIdx - 1];
            var cellText = cell.TextBody?.InnerText ?? "";
            var cellNode = new DocumentNode
            {
                Path = path,
                Type = "tc",
                Text = cellText
            };

            // Cell fill
            var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
            var cellFillHex = tcPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (cellFillHex != null) cellNode.Format["fill"] = cellFillHex;

            // Font info from first run
            var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
            if (firstRun?.RunProperties != null)
            {
                var f = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                    ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
                if (f != null) cellNode.Format["font"] = f;
                var fs = firstRun.RunProperties.FontSize?.Value;
                if (fs.HasValue) cellNode.Format["size"] = $"{fs.Value / 100}pt";
                if (firstRun.RunProperties.Bold?.Value == true) cellNode.Format["bold"] = true;
                if (firstRun.RunProperties.Italic?.Value == true) cellNode.Format["italic"] = true;
                var colorHex = firstRun.RunProperties.GetFirstChild<Drawing.SolidFill>()
                    ?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (colorHex != null) cellNode.Format["color"] = colorHex;
            }

            return cellNode;
        }

        // Try placeholder path with type name: /slide[N]/placeholder[title]
        var phGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phGetMatch.Success && !Regex.IsMatch(path, @"^/slide\[\d+\](?:/\w+\[\d+\])?$"))
        {
            var phSlideIdx = int.Parse(phGetMatch.Groups[1].Value);
            var phId = phGetMatch.Groups[2].Value;

            var phSlideParts = GetSlideParts().ToList();
            if (phSlideIdx < 1 || phSlideIdx > phSlideParts.Count)
                throw new ArgumentException($"Slide {phSlideIdx} not found");

            var phSlidePart = phSlideParts[phSlideIdx - 1];

            // If numeric, delegate to GetPlaceholderNode
            if (int.TryParse(phId, out var phNumIdx))
                return GetPlaceholderNode(phSlidePart, phSlideIdx, phNumIdx, depth);

            // By type name: resolve the shape and return its node
            var phShape = ResolvePlaceholderShape(phSlidePart, phId);
            var ph = phShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            var shapeTree = GetSlide(phSlidePart).CommonSlideData?.ShapeTree;
            var shapeIdx = shapeTree?.Elements<Shape>().ToList().IndexOf(phShape) ?? 0;
            var node = ShapeToNode(phShape, phSlideIdx, shapeIdx + 1, depth, phSlidePart);
            node.Path = path;
            node.Type = "placeholder";
            if (ph?.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
            if (ph?.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
            return node;
        }

        // Try resolving logical paths with deeper segments (e.g. /slide[1]/table[1]/tr[1])
        // Only for paths not handled by dedicated handlers above
        if (Regex.IsMatch(path, @"^/slide\[\d+\]/(table\[\d+\]/(tr|tc)|placeholder\[\w+\]/)"))
        {
            var logicalResolved = ResolveLogicalPath(path);
            if (logicalResolved.HasValue)
                return GenericXmlQuery.ElementToNode(logicalResolved.Value.element, path, depth);
        }

        // Parse /slide[N] or /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!match.Success)
        {
            // Generic XML fallback: navigate by element localName
            var allSegments = GenericXmlQuery.ParsePathSegments(path);
            if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                throw new ArgumentException($"Path must start with /slide[N]: {path}");

            var fbSlideIdx = allSegments[0].Index!.Value;
            var fbSlideParts = GetSlideParts().ToList();
            if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

            OpenXmlElement fbCurrent = GetSlide(fbSlideParts[fbSlideIdx - 1]);
            var remaining = allSegments.Skip(1).ToList();
            if (remaining.Count > 0)
            {
                var target = GenericXmlQuery.NavigateByPath(fbCurrent, remaining);
                if (target == null)
                    return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {path}" };
                return GenericXmlQuery.ElementToNode(target, path, depth);
            }
            return GenericXmlQuery.ElementToNode(fbCurrent, path, depth);
        }

        var slideIdx = int.Parse(match.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var targetSlidePart = slideParts[slideIdx - 1];

        if (!match.Groups[2].Success)
        {
            // Return slide node
            var slide = GetSlide(targetSlidePart);
            var slideNode = new DocumentNode
            {
                Path = path,
                Type = "slide",
                Preview = slide.CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)"
            };
            var layoutName = GetSlideLayoutName(targetSlidePart);
            if (layoutName != null) slideNode.Format["layout"] = layoutName;
            var layoutType = GetSlideLayoutType(targetSlidePart);
            if (layoutType != null) slideNode.Format["layoutType"] = layoutType;
            ReadSlideBackground(slide, slideNode);
            ReadSlideTransition(slide, slideNode);
            slideNode.Children = GetSlideChildNodes(targetSlidePart, slideIdx, depth);
            slideNode.ChildCount = slideNode.Children.Count;
            return slideNode;
        }

        // Shape or picture
        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);
        var shapeTreeEl = GetSlide(targetSlidePart).CommonSlideData?.ShapeTree;
        if (shapeTreeEl == null)
            throw new ArgumentException($"Slide {slideIdx} has no shapes");

        if (elementType == "shape")
        {
            var shapes = shapeTreeEl.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found (total: {shapes.Count})");
            return ShapeToNode(shapes[elementIdx - 1], slideIdx, elementIdx, depth, targetSlidePart);
        }
        else if (elementType == "table")
        {
            var tables = shapeTreeEl.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > tables.Count)
                throw new ArgumentException($"Table {elementIdx} not found (total: {tables.Count})");
            return TableToNode(tables[elementIdx - 1], slideIdx, elementIdx, depth);
        }
        else if (elementType == "placeholder")
        {
            return GetPlaceholderNode(targetSlidePart, slideIdx, elementIdx, depth);
        }
        else if (elementType == "chart")
        {
            var charts = shapeTreeEl.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > charts.Count)
                throw new ArgumentException($"Chart {elementIdx} not found (total: {charts.Count})");
            return ChartToNode(charts[elementIdx - 1], targetSlidePart, slideIdx, elementIdx, depth);
        }
        else if (elementType == "picture" || elementType == "pic")
        {
            var pics = shapeTreeEl.Elements<Picture>().ToList();
            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"Picture {elementIdx} not found (total: {pics.Count})");
            return PictureToNode(pics[elementIdx - 1], slideIdx, elementIdx);
        }

        // Generic fallback for unknown element types
        {
            var shapes2 = shapeTreeEl.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase)).ToList();
            if (elementIdx < 1 || elementIdx > shapes2.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {shapes2.Count})");
            return GenericXmlQuery.ElementToNode(shapes2[elementIdx - 1], path, depth);
        }
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();
        var parsed = ParseShapeSelector(selector);
        bool isEquationSelector = parsed.ElementType is "equation" or "math" or "formula";

        // Scheme B: generic XML fallback for unrecognized element types
        // Check if selector has a type that ParseShapeSelector didn't recognize
        // Extract raw element type for generic XML fallback check
        // Strip pseudo-selectors (:contains, :empty, :no-alt) and attribute filters before checking
        var selectorForType = Regex.Replace(selector, @":(contains\([^)]*\)|empty|no-alt)", "");
        var typeMatch = Regex.Match(selectorForType.Contains(']') ? selectorForType.Split(']').Last() : selectorForType, @"^(?:slide\[\d+\]\s*>?\s*)?([\w:]+)");
        var rawType = typeMatch.Success ? typeMatch.Groups[1].Value.ToLowerInvariant() : "";
        bool isKnownType = string.IsNullOrEmpty(rawType)
            || rawType is "shape" or "textbox" or "title" or "picture" or "pic"
                or "equation" or "math" or "formula"
                or "table" or "chart" or "placeholder" or "notes";
        if (!isKnownType)
        {
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var slidePart in GetSlideParts())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSlide(slidePart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        // Notes query (notes live outside the shape tree in NotesSlidePart)
        if (rawType == "notes")
        {
            int notesSlideNum = 0;
            foreach (var sp in GetSlideParts())
            {
                notesSlideNum++;
                if (sp.NotesSlidePart == null) continue;
                var notesText = GetNotesText(sp.NotesSlidePart);
                if (string.IsNullOrEmpty(notesText)) continue;
                if (parsed.TextContains != null && !notesText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                    continue;
                results.Add(new DocumentNode
                {
                    Path = $"/slide[{notesSlideNum}]/notes",
                    Type = "notes",
                    Text = notesText
                });
            }
            return results;
        }

        int slideNum = 0;

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;

            // Slide filter
            if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != slideNum)
                continue;

            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            int shapeIdx = 0;
            foreach (var shape in shapeTree.Elements<Shape>())
            {
                if (isEquationSelector)
                {
                    var mathElements = FindShapeMathElements(shape);
                    foreach (var mathElem in mathElements)
                    {
                        var latex = FormulaParser.ToLatex(mathElem);
                        if (parsed.TextContains == null || latex.Contains(parsed.TextContains))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/slide[{slideNum}]/shape[{shapeIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "display" }
                            });
                        }
                    }
                }
                else if (MatchesShapeSelector(shape, parsed))
                {
                    results.Add(ShapeToNode(shape, slideNum, shapeIdx + 1, 0, slidePart));
                }
                shapeIdx++;
            }

            if (parsed.ElementType == "picture" || parsed.ElementType == "pic" || parsed.ElementType == null)
            {
                int picIdx = 0;
                foreach (var pic in shapeTree.Elements<Picture>())
                {
                    if (MatchesPictureSelector(pic, parsed))
                    {
                        results.Add(PictureToNode(pic, slideNum, picIdx + 1));
                    }
                    picIdx++;
                }
            }

            if (parsed.ElementType == "table" || (parsed.ElementType == null && !isEquationSelector))
            {
                int tblIdx = 0;
                foreach (var gf in shapeTree.Elements<GraphicFrame>())
                {
                    if (!gf.Descendants<Drawing.Table>().Any()) continue;
                    tblIdx++;
                    var tblNode = TableToNode(gf, slideNum, tblIdx, 0);
                    if (parsed.TextContains != null)
                    {
                        // GraphicData children may be opaque when loaded from disk,
                        // so extract text from all <a:t> elements via OuterXml
                        var xml = gf.OuterXml;
                        var textMatches = Regex.Matches(xml, @"<a:t[^>]*>([^<]*)</a:t>");
                        var allText = string.Concat(textMatches.Select(m => m.Groups[1].Value));
                        if (!allText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(tblNode);
                }
            }

            if (parsed.ElementType == "chart" || (parsed.ElementType == null && !isEquationSelector))
            {
                int chartIdx = 0;
                foreach (var gf in shapeTree.Elements<GraphicFrame>())
                {
                    if (!gf.Descendants<C.ChartReference>().Any()) continue;
                    chartIdx++;
                    var chartNode = ChartToNode(gf, slidePart, slideNum, chartIdx, 0);
                    if (parsed.TextContains != null)
                    {
                        var titleVal = chartNode.Format.ContainsKey("title") ? (string)chartNode.Format["title"] : "";
                        if (!titleVal.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(chartNode);
                }
            }

            if (parsed.ElementType == "placeholder")
            {
                int phIdx = 0;
                foreach (var shape in shapeTree.Elements<Shape>())
                {
                    var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    if (ph == null) continue;
                    phIdx++;

                    if (parsed.TextContains != null)
                    {
                        var shapeText = GetShapeText(shape);
                        if (!shapeText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }

                    var node = ShapeToNode(shape, slideNum, phIdx, 0, slidePart);
                    node.Path = $"/slide[{slideNum}]/placeholder[{phIdx}]";
                    node.Type = "placeholder";
                    if (ph.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
                    if (ph.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
                    results.Add(node);
                }
            }
        }

        return results;
    }
}
