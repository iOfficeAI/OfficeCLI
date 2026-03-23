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
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Path cannot be empty.");
        if (path == "/")
        {
            var node = new DocumentNode { Path = "/", Type = "presentation" };

            // Slide size
            var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
            if (sldSz != null)
            {
                if (sldSz.Cx?.HasValue == true) node.Format["slideWidth"] = FormatEmu(sldSz.Cx.Value);
                if (sldSz.Cy?.HasValue == true) node.Format["slideHeight"] = FormatEmu(sldSz.Cy.Value);
                if (sldSz.Type?.HasValue == true) node.Format["slideSize"] = sldSz.Type.InnerText switch
                {
                    "screen16x9" => "widescreen",
                    "screen4x3" => "standard",
                    "screen16x10" => "16:10",
                    "a4" => "a4",
                    "a3" => "a3",
                    "letter" => "letter",
                    "b4ISO" => "b4",
                    "b5ISO" => "b5",
                    "35mm" => "35mm",
                    "overhead" => "overhead",
                    "banner" => "banner",
                    "ledger" => "ledger",
                    "custom" => "custom",
                    var other => other
                };
            }

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
                ReadSlideTransition(slidePart, slideNode);

                if (depth > 0)
                {
                    slideNode.Children = GetSlideChildNodes(slidePart, slideNum, depth - 1);
                    slideNode.ChildCount = slideNode.Children.Count;
                }
                else
                {
                    var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                    slideNode.ChildCount = (shapeTree?.Elements<Shape>().Count() ?? 0)
                        + (shapeTree?.Elements<Picture>().Count() ?? 0)
                        + (shapeTree?.Elements<GraphicFrame>().Count() ?? 0)
                        + (shapeTree?.Elements<ConnectionShape>().Count() ?? 0)
                        + (shapeTree?.Elements<GroupShape>().Count() ?? 0)
                        + (shapeTree != null ? GetZoomElements(shapeTree).Count : 0);
                }

                node.Children.Add(slideNode);
            }
            node.ChildCount = node.Children.Count;
            return node;
        }

        if (path.Equals("/theme", StringComparison.OrdinalIgnoreCase))
            return GetThemeNode();

        if (path.Equals("/morph-check", StringComparison.OrdinalIgnoreCase))
            return GetMorphCheckNode();

        // Try notes path: /slide[N]/notes
        var notesGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesGetMatch.Success)
        {
            var notesSlideIdx = int.Parse(notesGetMatch.Groups[1].Value);
            var slidePartsN = GetSlideParts().ToList();
            if (notesSlideIdx < 1 || notesSlideIdx > slidePartsN.Count)
                throw new ArgumentException($"Slide {notesSlideIdx} not found (total: {slidePartsN.Count})");
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
            var qParaPProps = para.ParagraphProperties;
            if (qParaPProps?.Alignment?.HasValue == true) paraNode.Format["align"] = NormalizeAlignment(qParaPProps.Alignment.InnerText);
            if (qParaPProps?.Indent?.HasValue == true) paraNode.Format["indent"] = FormatEmu(qParaPProps.Indent.Value);
            if (qParaPProps?.LeftMargin?.HasValue == true) paraNode.Format["marginLeft"] = FormatEmu(qParaPProps.LeftMargin.Value);
            if (qParaPProps?.RightMargin?.HasValue == true) paraNode.Format["marginRight"] = FormatEmu(qParaPProps.RightMargin.Value);
            var qLsPct = qParaPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (qLsPct.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(qLsPct.Value);
            var qLsPts = qParaPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (qLsPts.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(qLsPts.Value);
            var qSb = qParaPProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (qSb.HasValue) paraNode.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(qSb.Value);
            var qSa = qParaPProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (qSa.HasValue) paraNode.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(qSa.Value);

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

        // Try zoom path: /slide[N]/zoom[M]
        var zoomGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/zoom\[(\d+)\]$");
        if (zoomGetMatch.Success)
        {
            var sIdx = int.Parse(zoomGetMatch.Groups[1].Value);
            var zmIdx = int.Parse(zoomGetMatch.Groups[2].Value);
            var zmSlideParts = GetSlideParts().ToList();
            if (sIdx < 1 || sIdx > zmSlideParts.Count)
                throw new ArgumentException($"Slide {sIdx} not found (total: {zmSlideParts.Count})");
            var zmSlidePart = zmSlideParts[sIdx - 1];
            var zmShapeTree = GetSlide(zmSlidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException($"Slide {sIdx} has no shapes");
            var zoomElements = GetZoomElements(zmShapeTree);
            if (zmIdx < 1 || zmIdx > zoomElements.Count)
                throw new ArgumentException($"Zoom {zmIdx} not found (total: {zoomElements.Count})");
            return ZoomToNode(zoomElements[zmIdx - 1], sIdx, zmIdx);
        }

        // Try animation path: /slide[N]/shape[M]/animation[A]
        var animPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/animation\[(\d+)\]$");
        if (animPathMatch.Success)
        {
            var sIdx = int.Parse(animPathMatch.Groups[1].Value);
            var shIdx = int.Parse(animPathMatch.Groups[2].Value);
            var aIdx = int.Parse(animPathMatch.Groups[3].Value);
            var (animSlidePart, animShape) = ResolveShape(sIdx, shIdx);

            var animNode = new DocumentNode { Path = path, Type = "animation" };

            // Read animation info from timing tree
            var shapeId = animShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (shapeId != null)
            {
                var timing = GetSlide(animSlidePart).GetFirstChild<Timing>();
                if (timing != null)
                {
                    var shapeIdStr = shapeId.Value.ToString();
                    // Find all effect CTns for this shape
                    var effectCTns = timing.Descendants<CommonTimeNode>()
                        .Where(ctn => ctn.PresetClass != null && ctn.PresetId != null &&
                               ctn.GetAttributes().All(a => a.LocalName != "presetClass" || a.Value != "motion") &&
                               ctn.Descendants<ShapeTarget>().Any(st => st.ShapeId?.Value == shapeIdStr))
                        .ToList();

                    if (aIdx >= 1 && aIdx <= effectCTns.Count)
                    {
                        var effectCTn = effectCTns[aIdx - 1];
                        var presetId = effectCTn.PresetId?.Value ?? 0;
                        var clsVal = effectCTn.PresetClass?.Value;
                        var cls = clsVal == TimeNodePresetClassValues.Exit ? "exit"
                                : clsVal == TimeNodePresetClassValues.Emphasis ? "emphasis"
                                : "entrance";

                        var animEffect = effectCTn.Descendants<AnimateEffect>().FirstOrDefault();
                        var filter = animEffect?.Filter?.Value ?? "";

                        var effectName = filter switch
                        {
                            "fly" => "fly",
                            "fade" => "fade",
                            "zoom" => "zoom",
                            "" when presetId == 1 => "appear",
                            "" when presetId == 24 => "bounce",
                            _ => presetId switch
                            {
                                1 => "appear", 2 => "fly", 10 => "fade",
                                21 => "zoom", 24 => "bounce", _ => "unknown"
                            }
                        };

                        animNode.Format["effect"] = effectName;
                        animNode.Format["class"] = cls;
                        animNode.Format["presetId"] = presetId;

                        var dur = 500;
                        if (int.TryParse(animEffect?.CommonBehavior?.CommonTimeNode?.Duration, out var dd)) dur = dd;
                        animNode.Format["duration"] = dur;

                        // Easing (stored as 0-100000 on effectCTn)
                        if (effectCTn.Acceleration?.HasValue == true && effectCTn.Acceleration.Value > 0)
                            animNode.Format["easein"] = (int)(effectCTn.Acceleration.Value / 1000);
                        if (effectCTn.Deceleration?.HasValue == true && effectCTn.Deceleration.Value > 0)
                            animNode.Format["easeout"] = (int)(effectCTn.Deceleration.Value / 1000);

                        // Delay (stored on midCTn start condition)
                        // Walk up from effectCTn to find the wrapping midCTn that holds the delay.
                        // The nesting depth can vary: effectCTn > ParallelTimeNode > ChildTimeNodeList > midCTn
                        CommonTimeNode? midCTn = null;
                        var cur = effectCTn.Parent;
                        for (int walkDepth = 0; walkDepth < 5 && cur != null; walkDepth++)
                        {
                            if (cur is CommonTimeNode candidate && candidate != effectCTn
                                && candidate.PresetId == null) // midCTn has no presetId
                            {
                                midCTn = candidate;
                                break;
                            }
                            cur = cur.Parent;
                        }
                        var midDelayVal = midCTn?.StartConditionList?.GetFirstChild<Condition>()?.Delay?.Value;
                        if (midDelayVal != null && midDelayVal != "0"
                            && int.TryParse(midDelayVal, out var dMs) && dMs > 0)
                            animNode.Format["delay"] = dMs;
                    }
                }
            }
            return animNode;
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

            // GridSpan / RowSpan
            if (cell.GridSpan?.HasValue == true && cell.GridSpan.Value > 1)
                cellNode.Format["gridSpan"] = cell.GridSpan.Value;
            if (cell.RowSpan?.HasValue == true && cell.RowSpan.Value > 1)
                cellNode.Format["rowSpan"] = cell.RowSpan.Value;
            if (cell.HorizontalMerge?.HasValue == true && cell.HorizontalMerge.Value)
                cellNode.Format["hmerge"] = true;
            if (cell.VerticalMerge?.HasValue == true && cell.VerticalMerge.Value)
                cellNode.Format["vmerge"] = true;

            // Cell fill (blip, gradient, or solid)
            var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
            var cellBlipFill = tcPr?.GetFirstChild<Drawing.BlipFill>();
            if (cellBlipFill != null)
            {
                var blipEmbed = cellBlipFill.GetFirstChild<Drawing.Blip>()?.Embed?.Value;
                cellNode.Format["fill"] = "image";
                if (blipEmbed != null) cellNode.Format["image.relId"] = blipEmbed;
            }
            else if (tcPr?.GetFirstChild<Drawing.GradientFill>() is { } gradFill)
            {
                var stops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
                if (stops != null && stops.Count >= 2)
                {
                    var gc1 = ParseHelpers.FormatHexColor(stops[0].GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "");
                    var gc2 = ParseHelpers.FormatHexColor(stops[^1].GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "");
                    var lin = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
                    int deg = lin?.Angle?.Value != null ? lin.Angle.Value / 60000 : 0;
                    var gradient = $"linear;{gc1};{gc2};{deg}";
                    cellNode.Format["fill"] = deg != 0 ? $"{gc1}-{gc2}-{deg}" : $"{gc1}-{gc2}";
                    cellNode.Format["gradient"] = gradient;
                }
            }
            else
            {
                var cellFillHex = tcPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (cellFillHex != null) cellNode.Format["fill"] = ParseHelpers.FormatHexColor(cellFillHex);
            }

            // Cell borders — following POI's getBorderWidth/getBorderColor pattern
            if (tcPr != null)
                ReadTableCellBorders(tcPr, cellNode);

            // Vertical alignment
            if (tcPr?.Anchor?.HasValue == true)
            {
                cellNode.Format["valign"] = tcPr.Anchor.InnerText switch
                {
                    "ctr" => "center",
                    _ => tcPr.Anchor.InnerText
                };
            }

            // Alignment from first paragraph
            var cellFirstPara = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
            var cellParaAlign = cellFirstPara?.ParagraphProperties?.Alignment;
            if (cellParaAlign?.HasValue == true)
            {
                var align = NormalizeAlignment(cellParaAlign.InnerText);
                cellNode.Format["alignment"] = align;
                cellNode.Format["align"] = align;
            }

            // Font info from first run
            var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
            if (firstRun?.RunProperties != null)
            {
                var f = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                if (f != null) cellNode.Format["font"] = f;
                var fs = firstRun.RunProperties.FontSize?.Value;
                if (fs.HasValue) cellNode.Format["size"] = $"{fs.Value / 100.0:0.##}pt";
                if (firstRun.RunProperties.Bold?.Value == true) cellNode.Format["bold"] = true;
                if (firstRun.RunProperties.Italic?.Value == true) cellNode.Format["italic"] = true;
                var colorHex = firstRun.RunProperties.GetFirstChild<Drawing.SolidFill>()
                    ?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (colorHex != null) cellNode.Format["color"] = ParseHelpers.FormatHexColor(colorHex);
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
                throw new ArgumentException($"Slide {phSlideIdx} not found (total: {phSlideParts.Count})");

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
            ReadSlideTransition(targetSlidePart, slideNode);
            if (targetSlidePart.NotesSlidePart != null)
            {
                var notesText = GetNotesText(targetSlidePart.NotesSlidePart);
                if (!string.IsNullOrEmpty(notesText))
                    slideNode.Format["notes"] = notesText;
            }
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
            return PictureToNode(pics[elementIdx - 1], slideIdx, elementIdx, targetSlidePart);
        }
        else if (elementType == "connector" || elementType == "connection")
        {
            var connectors = shapeTreeEl.Elements<ConnectionShape>().ToList();
            if (elementIdx < 1 || elementIdx > connectors.Count)
                throw new ArgumentException($"Connector {elementIdx} not found (total: {connectors.Count})");
            return ConnectorToNode(connectors[elementIdx - 1], slideIdx, elementIdx);
        }
        else if (elementType == "group")
        {
            var groups = shapeTreeEl.Elements<GroupShape>().ToList();
            if (elementIdx < 1 || elementIdx > groups.Count)
                throw new ArgumentException($"Group {elementIdx} not found (total: {groups.Count})");
            var grp = groups[elementIdx - 1];
            var grpName = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
            var grpNode = new DocumentNode
            {
                Path = $"/slide[{slideIdx}]/group[{elementIdx}]",
                Type = "group",
                Preview = grpName,
                ChildCount = grp.Elements<Shape>().Count() + grp.Elements<Picture>().Count()
                    + grp.Elements<GraphicFrame>().Count() + grp.Elements<ConnectionShape>().Count()
                    + grp.Elements<GroupShape>().Count()
            };
            grpNode.Format["name"] = grpName;
            return grpNode;
        }

        // Generic fallback for unknown element types
        {
            var shapes2 = shapeTreeEl.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase)).ToList();
            if (elementIdx < 1 || elementIdx > shapes2.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {shapes2.Count}). Slide {slideIdx} contains: {DescribeSlideInventory(shapeTreeEl)}");
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
        // Strip pseudo-selectors (:contains, :empty, :no-alt) and shorthand :text before checking
        var selectorForType = Regex.Replace(selector, @":(contains\([^)]*\)|empty|no-alt)", "");
        // Also strip shorthand ":text" syntax so "shape:Find me" → "shape"
        selectorForType = Regex.Replace(selectorForType, @":(?![\[\(]).*$", "");
        var typeMatch = Regex.Match(selectorForType.Contains(']') ? selectorForType.Split(']').Last() : selectorForType, @"^(?:slide\[\d+\]\s*>?\s*)?([\w]+)");
        var rawType = typeMatch.Success ? typeMatch.Groups[1].Value.ToLowerInvariant() : "";
        bool isKnownType = string.IsNullOrEmpty(rawType)
            || rawType is "shape" or "textbox" or "title" or "picture" or "pic"
                or "video" or "audio"
                or "equation" or "math" or "formula"
                or "table" or "chart" or "placeholder" or "notes"
                or "connector" or "connection"
                or "group" or "zoom";
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
                    var node = ShapeToNode(shape, slideNum, shapeIdx + 1, 0, slidePart);
                    if (MatchesGenericAttributes(node, parsed.Attributes))
                        results.Add(node);
                }
                shapeIdx++;
            }

            if (parsed.ElementType is "picture" or "pic" or "video" or "audio" or null)
            {
                int picIdx = 0;
                foreach (var pic in shapeTree.Elements<Picture>())
                {
                    var picNvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                    var picIsVideo = picNvPr?.GetFirstChild<Drawing.VideoFromFile>() != null;
                    var picIsAudio = picNvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;

                    // Filter by media type
                    if (parsed.ElementType == "video" && !picIsVideo) { picIdx++; continue; }
                    if (parsed.ElementType == "audio" && !picIsAudio) { picIdx++; continue; }
                    if (parsed.ElementType is "picture" or "pic" && (picIsVideo || picIsAudio)) { picIdx++; continue; }

                    if (MatchesPictureSelector(pic, parsed))
                    {
                        var picNode = PictureToNode(pic, slideNum, picIdx + 1, slidePart);
                        if (MatchesGenericAttributes(picNode, parsed.Attributes))
                            results.Add(picNode);
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
                        var titleVal = chartNode.Format.ContainsKey("title") ? chartNode.Format["title"]?.ToString() ?? "" : "";
                        if (!titleVal.Contains(parsed.TextContains!, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(chartNode);
                }
            }

            if (parsed.ElementType is "connector" or "connection" || (parsed.ElementType == null && !isEquationSelector))
            {
                int cxnIdx = 0;
                foreach (var cxn in shapeTree.Elements<ConnectionShape>())
                {
                    cxnIdx++;
                    if (parsed.ElementType is "connector" or "connection" || parsed.ElementType == null)
                    {
                        var cxnNode = ConnectorToNode(cxn, slideNum, cxnIdx);
                        if (MatchesGenericAttributes(cxnNode, parsed.Attributes))
                            results.Add(cxnNode);
                    }
                }
            }

            if (parsed.ElementType == "group" || (parsed.ElementType == null && !isEquationSelector))
            {
                int grpIdx = 0;
                foreach (var grp in shapeTree.Elements<GroupShape>())
                {
                    grpIdx++;
                    if (parsed.ElementType == "group" || parsed.ElementType == null)
                    {
                        var grpName = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
                        var grpNode = new DocumentNode
                        {
                            Path = $"/slide[{slideNum}]/group[{grpIdx}]",
                            Type = "group",
                            Preview = grpName,
                            ChildCount = grp.Elements<Shape>().Count() + grp.Elements<Picture>().Count()
                            + grp.Elements<GraphicFrame>().Count() + grp.Elements<ConnectionShape>().Count()
                            + grp.Elements<GroupShape>().Count()
                        };
                        grpNode.Format["name"] = grpName;
                        if (MatchesGenericAttributes(grpNode, parsed.Attributes))
                            results.Add(grpNode);
                    }
                }
            }

            if (parsed.ElementType == "zoom" || (parsed.ElementType == null && !isEquationSelector))
            {
                var zoomElements = GetZoomElements(shapeTree);
                int zmIdx = 0;
                foreach (var zmEl in zoomElements)
                {
                    zmIdx++;
                    var zmNode = ZoomToNode(zmEl, slideNum, zmIdx);
                    if (parsed.TextContains != null)
                    {
                        var zmName = zmNode.Format.ContainsKey("name") ? zmNode.Format["name"]?.ToString() ?? "" : "";
                        if (!zmName.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesGenericAttributes(zmNode, parsed.Attributes))
                        results.Add(zmNode);
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
