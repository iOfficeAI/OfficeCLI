// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        // Try notes path: /slide[N]/notes
        var notesSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesSetMatch.Success)
        {
            var slideIdx = int.Parse(notesSetMatch.Groups[1].Value);
            var slidePartsN = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slidePartsN.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");
            var notesPart = EnsureNotesSlidePart(slidePartsN[slideIdx - 1]);
            var unsupportedN = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (key.Equals("text", StringComparison.OrdinalIgnoreCase))
                    SetNotesText(notesPart, value);
                else
                    unsupportedN.Add(key);
            }
            return unsupportedN;
        }

        // Try run-level path: /slide[N]/shape[M]/run[K]
        var runMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/run\[(\d+)\]$");
        if (runMatch.Success)
        {
            var slideIdx = int.Parse(runMatch.Groups[1].Value);
            var shapeIdx = int.Parse(runMatch.Groups[2].Value);
            var runIdx = int.Parse(runMatch.Groups[3].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var allRuns = GetAllRuns(shape);
            if (runIdx < 1 || runIdx > allRuns.Count)
                throw new ArgumentException($"Run {runIdx} not found (shape has {allRuns.Count} runs)");

            var targetRun = allRuns[runIdx - 1];
            var linkValRun = properties.GetValueOrDefault("link");
            var runOnlyProps = properties
                .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase))
                .ToDictionary(kv => kv.Key, kv => kv.Value);
            var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart);
            if (linkValRun != null) ApplyRunHyperlink(slidePart, targetRun, linkValRun);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try paragraph/run path: /slide[N]/shape[M]/paragraph[P]/run[K]
        var paraRunMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]/run\[(\d+)\]$");
        if (paraRunMatch.Success)
        {
            var slideIdx = int.Parse(paraRunMatch.Groups[1].Value);
            var shapeIdx = int.Parse(paraRunMatch.Groups[2].Value);
            var paraIdx = int.Parse(paraRunMatch.Groups[3].Value);
            var runIdx = int.Parse(paraRunMatch.Groups[4].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                ?? throw new ArgumentException("Shape has no text body");
            if (paraIdx < 1 || paraIdx > paragraphs.Count)
                throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paraIdx - 1];
            var paraRuns = para.Elements<Drawing.Run>().ToList();
            if (runIdx < 1 || runIdx > paraRuns.Count)
                throw new ArgumentException($"Run {runIdx} not found (paragraph has {paraRuns.Count} runs)");

            var targetRun = paraRuns[runIdx - 1];
            var unsupported = SetRunOrShapeProperties(properties, new List<Drawing.Run> { targetRun }, shape, slidePart);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try paragraph-level path: /slide[N]/shape[M]/paragraph[P]
        var paraMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]$");
        if (paraMatch.Success)
        {
            var slideIdx = int.Parse(paraMatch.Groups[1].Value);
            var shapeIdx = int.Parse(paraMatch.Groups[2].Value);
            var paraIdx = int.Parse(paraMatch.Groups[3].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                ?? throw new ArgumentException("Shape has no text body");
            if (paraIdx < 1 || paraIdx > paragraphs.Count)
                throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paraIdx - 1];
            var paraRuns = para.Elements<Drawing.Run>().ToList();
            var unsupported = new List<string>();

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "align":
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = ParseTextAlignment(value);
                        break;
                    default:
                        // Apply run-level properties to all runs in this paragraph
                        var runUnsup = SetRunOrShapeProperties(
                            new Dictionary<string, string> { { key, value } }, paraRuns, shape, slidePart);
                        unsupported.AddRange(runUnsup);
                        break;
                }
            }

            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try chart path: /slide[N]/chart[M]
        var chartSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/chart\[(\d+)\]$");
        if (chartSetMatch.Success)
        {
            var slideIdx = int.Parse(chartSetMatch.Groups[1].Value);
            var chartIdx = int.Parse(chartSetMatch.Groups[2].Value);

            var (slidePart, chartGf, chartPart) = ResolveChart(slideIdx, chartIdx);

            // Handle position/size properties on GraphicFrame
            var chartProps = new Dictionary<string, string>();
            var gfProps = new Dictionary<string, string>();
            foreach (var (key, value) in properties)
            {
                if (key.ToLowerInvariant() is "x" or "y" or "width" or "height" or "name")
                    gfProps[key] = value;
                else
                    chartProps[key] = value;
            }

            // Position/size
            foreach (var (key, value) in gfProps)
            {
                switch (key.ToLowerInvariant())
                {
                    case "x" or "y" or "width" or "height":
                        var xfrm = chartGf.Transform ?? (chartGf.Transform = new Transform());
                        var offset = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                        var extents = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                        var emu = ParseEmu(value);
                        switch (key.ToLowerInvariant())
                        {
                            case "x": offset.X = emu; break;
                            case "y": offset.Y = emu; break;
                            case "width": extents.Cx = emu; break;
                            case "height": extents.Cy = emu; break;
                        }
                        break;
                    case "name":
                        var nvPr = chartGf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                        if (nvPr != null) nvPr.Name = value;
                        break;
                }
            }

            var unsupported = SetChartProperties(chartPart, chartProps);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellMatch.Success)
        {
            var slideIdx = int.Parse(tblCellMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblCellMatch.Groups[2].Value);
            var rowIdx = int.Parse(tblCellMatch.Groups[3].Value);
            var cellIdx = int.Parse(tblCellMatch.Groups[4].Value);

            var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
            var tableRows = table.Elements<Drawing.TableRow>().ToList();
            if (rowIdx < 1 || rowIdx > tableRows.Count)
                throw new ArgumentException($"Row {rowIdx} not found (table has {tableRows.Count} rows)");
            var cells = tableRows[rowIdx - 1].Elements<Drawing.TableCell>().ToList();
            if (cellIdx < 1 || cellIdx > cells.Count)
                throw new ArgumentException($"Cell {cellIdx} not found (row has {cells.Count} cells)");

            var cell = cells[cellIdx - 1];
            var unsupported = SetTableCellProperties(cell, properties);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try table-level path: /slide[N]/table[M]
        var tblMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]$");
        if (tblMatch.Success)
        {
            var slideIdx = int.Parse(tblMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblMatch.Groups[2].Value);

            var slideParts2 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts2.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slidePart = slideParts2[slideIdx - 1];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");
            var graphicFrames = shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (tblIdx < 1 || tblIdx > graphicFrames.Count)
                throw new ArgumentException($"Table {tblIdx} not found (total: {graphicFrames.Count})");

            var gf = graphicFrames[tblIdx - 1];
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "x" or "y" or "width" or "height":
                    {
                        var xfrm = gf.Transform ?? (gf.Transform = new Transform());
                        var offset = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                        var extents = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                        var emu = ParseEmu(value);
                        switch (key.ToLowerInvariant())
                        {
                            case "x": offset.X = emu; break;
                            case "y": offset.Y = emu; break;
                            case "width": extents.Cx = emu; break;
                            case "height": extents.Cy = emu; break;
                        }
                        break;
                    }
                    case "name":
                        var nvPr = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                        if (nvPr != null) nvPr.Name = value;
                        break;
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(gf, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try table row path: /slide[N]/table[M]/tr[R]
        var tblRowMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (tblRowMatch.Success)
        {
            var slideIdx = int.Parse(tblRowMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblRowMatch.Groups[2].Value);
            var rowIdx = int.Parse(tblRowMatch.Groups[3].Value);

            var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
            var tableRows = table.Elements<Drawing.TableRow>().ToList();
            if (rowIdx < 1 || rowIdx > tableRows.Count)
                throw new ArgumentException($"Row {rowIdx} not found (table has {tableRows.Count} rows)");

            var row = tableRows[rowIdx - 1];
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "height":
                        row.Height = ParseEmu(value);
                        break;
                    default:
                        // Apply to all cells in this row
                        var cellUnsup = new HashSet<string>();
                        foreach (var cell in row.Elements<Drawing.TableCell>())
                        {
                            var u = SetTableCellProperties(cell, new Dictionary<string, string> { { key, value } });
                            foreach (var k in u) cellUnsup.Add(k);
                        }
                        unsupported.AddRange(cellUnsup);
                        break;
                }
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try placeholder path: /slide[N]/placeholder[M] or /slide[N]/placeholder[type]
        var phMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phMatch.Success)
        {
            var slideIdx = int.Parse(phMatch.Groups[1].Value);
            var phId = phMatch.Groups[2].Value;

            var slideParts2 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts2.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");
            var slidePart = slideParts2[slideIdx - 1];
            var shape = ResolvePlaceholderShape(slidePart, phId);

            var allRuns = shape.Descendants<Drawing.Run>().ToList();
            var unsupported = SetRunOrShapeProperties(properties, allRuns, shape, slidePart);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try slide-level path: /slide[N]
        var slideOnlyMatch = Regex.Match(path, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success)
        {
            var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
            var slideParts2 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts2.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");
            var slidePart2 = slideParts2[slideIdx - 1];
            var slide2 = GetSlide(slidePart2);

            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "background":
                        ApplySlideBackground(slidePart2, value);
                        break;
                    case "transition":
                        ApplyTransition(slidePart2, value);
                        break;
                    case "advancetime" or "advanceaftertime":
                    {
                        var trans = slide2.GetFirstChild<Transition>() ?? slide2.AppendChild(new Transition());
                        trans.AdvanceAfterTime = value;
                        break;
                    }
                    case "advanceclick" or "advanceonclick":
                    {
                        var trans = slide2.GetFirstChild<Transition>() ?? slide2.AppendChild(new Transition());
                        trans.AdvanceOnClick = bool.Parse(value);
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(slide2, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
            slide2.Save();
            return unsupported;
        }

        // Try shape-level path: /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
        if (match.Success)
        {
            var slideIdx = int.Parse(match.Groups[1].Value);
            var shapeIdx = int.Parse(match.Groups[2].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var allRuns = shape.Descendants<Drawing.Run>().ToList();

            // Separate animation and link (both need slidePart) from other shape properties
            var animValue = properties.GetValueOrDefault("animation")
                ?? properties.GetValueOrDefault("animate");
            var linkValue = properties.GetValueOrDefault("link");
            var shapeProps = properties
                .Where(kv => !kv.Key.Equals("animation", StringComparison.OrdinalIgnoreCase)
                          && !kv.Key.Equals("animate", StringComparison.OrdinalIgnoreCase)
                          && !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase))
                .ToDictionary(kv => kv.Key, kv => kv.Value);

            var unsupported = SetRunOrShapeProperties(shapeProps, allRuns, shape, slidePart);

            if (animValue != null)
                ApplyShapeAnimation(slidePart, shape, animValue);
            if (linkValue != null)
                ApplyShapeHyperlink(slidePart, shape, linkValue);

            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Generic XML fallback: navigate to element and set attributes
        {
            SlidePart fbSlidePart;
            OpenXmlElement target;

            // Try logical path resolution first (table/placeholder paths)
            var logicalResult = ResolveLogicalPath(path);
            if (logicalResult.HasValue)
            {
                fbSlidePart = logicalResult.Value.slidePart;
                target = logicalResult.Value.element;
            }
            else
            {
                var allSegments = GenericXmlQuery.ParsePathSegments(path);
                if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                    throw new ArgumentException($"Path must start with /slide[N]: {path}");

                var fbSlideIdx = allSegments[0].Index!.Value;
                var fbSlideParts = GetSlideParts().ToList();
                if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                    throw new ArgumentException($"Slide {fbSlideIdx} not found");

                fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                var remaining = allSegments.Skip(1).ToList();
                target = GetSlide(fbSlidePart);
                if (remaining.Count > 0)
                {
                    target = GenericXmlQuery.NavigateByPath(target, remaining)
                        ?? throw new ArgumentException($"Element not found: {path}");
                }
            }

            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSlide(fbSlidePart).Save();
            return unsup;
        }
    }
}
