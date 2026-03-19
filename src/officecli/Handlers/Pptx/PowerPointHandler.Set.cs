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
        // Presentation-level properties: / or /presentation
        if (path is "/" or "" or "/presentation")
        {
            var presentation = _doc.PresentationPart?.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "slidewidth" or "width":
                        var sldSz = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        sldSz.Cx = Core.EmuConverter.ParseEmuAsInt(value);
                        sldSz.Type = SlideSizeValues.Custom;
                        break;
                    case "slideheight" or "height":
                        var sldSz2 = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        sldSz2.Cy = Core.EmuConverter.ParseEmuAsInt(value);
                        sldSz2.Type = SlideSizeValues.Custom;
                        break;
                    case "slidesize":
                        var sz = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        switch (value.ToLowerInvariant())
                        {
                            case "16:9" or "widescreen":
                                sz.Cx = 12192000; sz.Cy = 6858000;
                                sz.Type = SlideSizeValues.Screen16x9;
                                break;
                            case "4:3" or "standard":
                                sz.Cx = 9144000; sz.Cy = 6858000;
                                sz.Type = SlideSizeValues.Screen4x3;
                                break;
                            case "16:10":
                                sz.Cx = 12192000; sz.Cy = 7620000;
                                sz.Type = SlideSizeValues.Screen16x10;
                                break;
                            case "a4":
                                sz.Cx = 10692000; sz.Cy = 7560000;
                                sz.Type = SlideSizeValues.A4;
                                break;
                            default:
                                unsupported.Add(key);
                                break;
                        }
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            presentation.Save();
            return unsupported;
        }

        // Try slideMaster/slideLayout editing: /slideMaster[N]/shape[M] or /slideLayout[N]/shape[M]
        var masterShapeMatch = Regex.Match(path, @"^/(slideMaster|slideLayout)\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (masterShapeMatch.Success)
        {
            var partType = masterShapeMatch.Groups[1].Value;
            var partIdx = int.Parse(masterShapeMatch.Groups[2].Value);
            var presentationPart = _doc.PresentationPart!;

            OpenXmlPartRootElement rootEl;
            if (partType == "slideMaster")
            {
                var masters = presentationPart.SlideMasterParts.ToList();
                if (partIdx < 1 || partIdx > masters.Count)
                    throw new ArgumentException($"SlideMaster {partIdx} not found");
                rootEl = masters[partIdx - 1].SlideMaster
                    ?? throw new InvalidOperationException("Corrupt slide master");
            }
            else
            {
                var layouts = presentationPart.SlideMasterParts
                    .SelectMany(m => m.SlideLayoutParts).ToList();
                if (partIdx < 1 || partIdx > layouts.Count)
                    throw new ArgumentException($"SlideLayout {partIdx} not found");
                rootEl = layouts[partIdx - 1].SlideLayout
                    ?? throw new InvalidOperationException("Corrupt slide layout");
            }

            if (!masterShapeMatch.Groups[3].Success)
            {
                // Set properties on the master/layout itself
                var unsupported = new List<string>();
                foreach (var (key, value) in properties)
                {
                    if (key.Equals("name", StringComparison.OrdinalIgnoreCase))
                    {
                        var csd = rootEl.GetFirstChild<CommonSlideData>();
                        if (csd != null) csd.Name = value;
                    }
                    else
                    {
                        unsupported.Add(key);
                    }
                }
                rootEl.Save();
                return unsupported;
            }

            // Set on a specific shape within master/layout
            var elType = masterShapeMatch.Groups[3].Value;
            var elIdx = int.Parse(masterShapeMatch.Groups[4].Value);
            var shapeTree = rootEl.Descendants<ShapeTree>().FirstOrDefault()
                ?? throw new ArgumentException("No shape tree found");

            if (elType == "shape")
            {
                var shapes = shapeTree.Elements<Shape>().ToList();
                if (elIdx < 1 || elIdx > shapes.Count)
                    throw new ArgumentException($"Shape {elIdx} not found");
                var shape = shapes[elIdx - 1];
                var allRuns = shape.Descendants<Drawing.Run>().ToList();
                var unsupported = SetRunOrShapeProperties(properties, allRuns, shape);
                rootEl.Save();
                return unsupported;
            }

            throw new ArgumentException($"Unsupported element type: {elType}");
        }

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
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = ParseTextAlignment(value);
                        break;
                    }
                    case "indent":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Indent = (int)ParseEmu(value);
                        break;
                    }
                    case "marginleft" or "marl":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.LeftMargin = (int)ParseEmu(value);
                        break;
                    }
                    case "marginright" or "marr":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RightMargin = (int)ParseEmu(value);
                        break;
                    }
                    case "linespacing" or "line.spacing":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.LineSpacing>();
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPercent { Val = (int)(ParseHelpers.SafeParseDouble(value, "lineSpacing") * 100000) }));
                        break;
                    }
                    case "spacebefore" or "space.before":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                        pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = (int)(ParseHelpers.SafeParseDouble(value, "spaceBefore") * 100) }));
                        break;
                    }
                    case "spaceafter" or "space.after":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                        pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = (int)(ParseHelpers.SafeParseDouble(value, "spaceAfter") * 100) }));
                        break;
                    }
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

            var unsupported = ChartHelper.SetChartProperties(chartPart, chartProps);
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
            // Clone cell for rollback on failure (atomic: no partial modifications)
            var cellBackup = cell.CloneNode(true);
            try
            {
                var unsupported = SetTableCellProperties(cell, properties);
                GetSlide(slidePart).Save();
                return unsupported;
            }
            catch
            {
                cell.Parent?.ReplaceChild(cellBackup, cell);
                throw;
            }
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
                    case "tablestyle" or "style":
                    {
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table != null)
                        {
                            var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                                ?? table.PrependChild(new Drawing.TableProperties());
                            // Well-known style names → GUIDs
                            var styleId = value.ToLowerInvariant() switch
                            {
                                "medium1" or "mediumstyle1" => "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
                                "medium2" or "mediumstyle2" => "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
                                "medium3" or "mediumstyle3" => "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
                                "medium4" or "mediumstyle4" => "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
                                "light1" or "lightstyle1" => "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
                                "light2" or "lightstyle2" => "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
                                "light3" or "lightstyle3" => "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
                                "dark1" or "darkstyle1" => "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
                                "dark2" or "darkstyle2" => "{125E5076-3810-47DD-B79F-674D7AD40C01}",
                                "none" => "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
                                _ when value.StartsWith("{") => value, // Direct GUID
                                _ => value // Pass through
                            };
                            tblPr.RemoveAllChildren<Drawing.TableStyleId>();
                            tblPr.AppendChild(new Drawing.TableStyleId(styleId));
                        }
                        break;
                    }
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

        // Try video/audio path: /slide[N]/video[M] or /slide[N]/audio[M]
        var mediaSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(video|audio)\[(\d+)\]$");
        if (mediaSetMatch.Success)
        {
            var slideIdx = int.Parse(mediaSetMatch.Groups[1].Value);
            var mediaType = mediaSetMatch.Groups[2].Value;
            var mediaIdx = int.Parse(mediaSetMatch.Groups[3].Value);

            var slideParts4 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts4.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slidePart = slideParts4[slideIdx - 1];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");

            var mediaPics = shapeTree.Elements<Picture>()
                .Where(p =>
                {
                    var nvPr = p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                    return mediaType == "video"
                        ? nvPr?.GetFirstChild<Drawing.VideoFromFile>() != null
                        : nvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;
                }).ToList();
            if (mediaIdx < 1 || mediaIdx > mediaPics.Count)
                throw new ArgumentException($"{mediaType} {mediaIdx} not found (total: {mediaPics.Count})");

            var pic = mediaPics[mediaIdx - 1];
            var shapeId = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
            var unsupported = new List<string>();

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "volume":
                    {
                        if (shapeId == null) { unsupported.Add(key); break; }
                        if (!double.TryParse(value, out var volVal))
                            throw new ArgumentException($"Invalid volume value: '{value}'. Expected a number (0-100).");
                        var vol = (int)(volVal * 1000); // 0-100 → 0-100000
                        var mediaNode = FindMediaTimingNode(slidePart, shapeId.Value);
                        if (mediaNode != null) mediaNode.Volume = vol;
                        break;
                    }
                    case "autoplay":
                    {
                        if (shapeId == null) { unsupported.Add(key); break; }
                        var mediaNode = FindMediaTimingNode(slidePart, shapeId.Value);
                        var cTn = mediaNode?.CommonTimeNode;
                        var startCond = cTn?.StartConditionList?.GetFirstChild<Condition>();
                        if (startCond != null)
                            startCond.Delay = IsTruthy(value) ? "0" : "indefinite";
                        break;
                    }
                    case "trimstart":
                    {
                        var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                        var p14Media = nvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
                        if (p14Media != null)
                        {
                            var trim = p14Media.MediaTrim ?? (p14Media.MediaTrim = new DocumentFormat.OpenXml.Office2010.PowerPoint.MediaTrim());
                            trim.Start = value;
                        }
                        break;
                    }
                    case "trimend":
                    {
                        var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                        var p14Media = nvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
                        if (p14Media != null)
                        {
                            var trim = p14Media.MediaTrim ?? (p14Media.MediaTrim = new DocumentFormat.OpenXml.Office2010.PowerPoint.MediaTrim());
                            trim.End = value;
                        }
                        break;
                    }
                    case "x" or "y" or "width" or "height":
                    {
                        var spPr = pic.ShapeProperties ?? (pic.ShapeProperties = new ShapeProperties());
                        var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
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
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try picture path: /slide[N]/picture[M]
        var picSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/picture\[(\d+)\]$");
        if (picSetMatch.Success)
        {
            var slideIdx = int.Parse(picSetMatch.Groups[1].Value);
            var picIdx = int.Parse(picSetMatch.Groups[2].Value);

            var slideParts3 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts3.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slidePart = slideParts3[slideIdx - 1];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");
            var pics = shapeTree.Elements<Picture>().ToList();
            if (picIdx < 1 || picIdx > pics.Count)
                throw new ArgumentException($"Picture {picIdx} not found (total: {pics.Count})");

            var pic = pics[picIdx - 1];
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "alt":
                        var nvPicPr = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                        if (nvPicPr != null) nvPicPr.Description = value;
                        break;
                    case "x" or "y" or "width" or "height":
                    {
                        var spPr = pic.ShapeProperties ?? (pic.ShapeProperties = new ShapeProperties());
                        var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
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
                    case "path" or "src":
                    {
                        // Replace image source
                        var blipFill = pic.BlipFill;
                        var blip = blipFill?.GetFirstChild<Drawing.Blip>();
                        if (blip == null) { unsupported.Add(key); break; }
                        if (!File.Exists(value))
                            throw new FileNotFoundException($"Image file not found: {value}");

                        var imgExt = Path.GetExtension(value).ToLowerInvariant();
                        var imgType = imgExt switch
                        {
                            ".png" => ImagePartType.Png,
                            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                            ".gif" => ImagePartType.Gif,
                            ".bmp" => ImagePartType.Bmp,
                            ".tif" or ".tiff" => ImagePartType.Tiff,
                            ".emf" => ImagePartType.Emf,
                            ".wmf" => ImagePartType.Wmf,
                            _ => ImagePartType.Png
                        };
                        // Remove old image part to avoid storage bloat
                        var oldEmbedId = blip.Embed?.Value;
                        if (oldEmbedId != null)
                        {
                            try { slidePart.DeletePart(oldEmbedId); } catch { }
                        }
                        var newImgPart = slidePart.AddImagePart(imgType);
                        using (var stream = File.OpenRead(value))
                            newImgPart.FeedData(stream);
                        blip.Embed = slidePart.GetIdOfPart(newImgPart);
                        break;
                    }
                    case "crop" or "cropleft" or "cropright" or "croptop" or "cropbottom":
                    {
                        var blipFill = pic.BlipFill;
                        if (blipFill == null) { unsupported.Add(key); break; }
                        var srcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>()
                            ?? blipFill.AppendChild(new Drawing.SourceRectangle());

                        if (key.Equals("crop", StringComparison.OrdinalIgnoreCase))
                        {
                            // Single value: "left,top,right,bottom" as percentages (0-100)
                            var parts = value.Split(',');
                            if (parts.Length == 4)
                            {
                                var cropVals = new double[4];
                                for (int ci = 0; ci < 4; ci++)
                                    cropVals[ci] = ParseHelpers.SafeParseDouble(parts[ci].Trim(), "crop");
                                srcRect.Left = (int)(cropVals[0] * 1000);
                                srcRect.Top = (int)(cropVals[1] * 1000);
                                srcRect.Right = (int)(cropVals[2] * 1000);
                                srcRect.Bottom = (int)(cropVals[3] * 1000);
                            }
                            else if (parts.Length == 1)
                            {
                                if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cropVal))
                                    throw new ArgumentException($"Invalid 'crop' value: '{value}'. Expected a percentage (e.g. 10 = 10% from each edge).");
                                var cropPct = (int)(cropVal * 1000);
                                srcRect.Left = cropPct; srcRect.Top = cropPct; srcRect.Right = cropPct; srcRect.Bottom = cropPct;
                            }
                        }
                        else
                        {
                            if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cropSingle))
                                throw new ArgumentException($"Invalid '{key}' value: '{value}'. Expected a percentage (0-100).");
                            var pct = (int)(cropSingle * 1000); // percent (0-100) → 1/1000ths
                            switch (key.ToLowerInvariant())
                            {
                                case "cropleft": srcRect.Left = pct; break;
                                case "croptop": srcRect.Top = pct; break;
                                case "cropright": srcRect.Right = pct; break;
                                case "cropbottom": srcRect.Bottom = pct; break;
                            }
                        }
                        break;
                    }
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
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
                        if (value.StartsWith("morph", StringComparison.OrdinalIgnoreCase))
                            AutoPrefixMorphNames(slidePart2);
                        else
                            AutoUnprefixMorphNames(slidePart2);
                        break;
                    case "advancetime" or "advanceaftertime":
                    {
                        var trans = FindOrCreateTransition(slide2);
                        trans.AdvanceAfterTime = value;
                        break;
                    }
                    case "advanceclick" or "advanceonclick":
                    {
                        var trans = FindOrCreateTransition(slide2);
                        trans.AdvanceOnClick = IsTruthy(value);
                        break;
                    }
                    case "notes":
                    {
                        var notesPart = EnsureNotesSlidePart(slidePart2);
                        SetNotesText(notesPart, value);
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

            // Handle z-order first (changes shape position in tree)
            var zOrderValue = properties.GetValueOrDefault("zorder")
                ?? properties.GetValueOrDefault("z-order")
                ?? properties.GetValueOrDefault("order");
            if (zOrderValue != null)
            {
                ApplyZOrder(slidePart, shape, zOrderValue);
            }

            // Clone shape for rollback on failure (atomic: no partial modifications)
            var shapeBackup = shape.CloneNode(true);

            try
            {
                var allRuns = shape.Descendants<Drawing.Run>().ToList();

                // Separate animation, link, and z-order from other shape properties
                var animValue = properties.GetValueOrDefault("animation")
                    ?? properties.GetValueOrDefault("animate");
                var linkValue = properties.GetValueOrDefault("link");
                var excludeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "animation", "animate", "link", "zorder", "z-order", "order" };
                var shapeProps = properties
                    .Where(kv => !excludeKeys.Contains(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);

                var unsupported = SetRunOrShapeProperties(shapeProps, allRuns, shape, slidePart);

                if (animValue != null)
                    ApplyShapeAnimation(slidePart, shape, animValue);
                if (linkValue != null)
                    ApplyShapeHyperlink(slidePart, shape, linkValue);

                GetSlide(slidePart).Save();
                return unsupported;
            }
            catch
            {
                // Rollback: restore shape to pre-modification state
                shape.Parent?.ReplaceChild(shapeBackup, shape);
                throw;
            }
        }

        // Try connector path: /slide[N]/connector[M]
        var cxnMatch = Regex.Match(path, @"^/slide\[(\d+)\]/connector\[(\d+)\]$");
        if (cxnMatch.Success)
        {
            var slideIdx = int.Parse(cxnMatch.Groups[1].Value);
            var cxnIdx = int.Parse(cxnMatch.Groups[2].Value);

            var slideParts5 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts5.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slidePart = slideParts5[slideIdx - 1];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");
            var connectors = shapeTree.Elements<ConnectionShape>().ToList();
            if (cxnIdx < 1 || cxnIdx > connectors.Count)
                throw new ArgumentException($"Connector {cxnIdx} not found (total: {connectors.Count})");

            var cxn = connectors[cxnIdx - 1];
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "name":
                        var nvCxnPr = cxn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties;
                        if (nvCxnPr != null) nvCxnPr.Name = value;
                        break;
                    case "x" or "y" or "width" or "height":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
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
                    case "linewidth" or "line.width":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        outline.Width = (int)ParseEmu(value);
                        break;
                    }
                    case "linecolor" or "line.color":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                        outline.RemoveAllChildren<Drawing.SolidFill>();
                        outline.PrependChild(new Drawing.SolidFill(
                            new Drawing.RgbColorModelHex { Val = rgb }));
                        break;
                    }
                    case "preset" or "prstgeom":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var prstGeom = spPr.GetFirstChild<Drawing.PresetGeometry>()
                            ?? spPr.AppendChild(new Drawing.PresetGeometry());
                        prstGeom.Preset = new Drawing.ShapeTypeValues(value);
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(cxn, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try group path: /slide[N]/group[M]
        var grpMatch = Regex.Match(path, @"^/slide\[(\d+)\]/group\[(\d+)\]$");
        if (grpMatch.Success)
        {
            var slideIdx = int.Parse(grpMatch.Groups[1].Value);
            var grpIdx = int.Parse(grpMatch.Groups[2].Value);

            var slideParts6 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts6.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slidePart = slideParts6[slideIdx - 1];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");
            var groups = shapeTree.Elements<GroupShape>().ToList();
            if (grpIdx < 1 || grpIdx > groups.Count)
                throw new ArgumentException($"Group {grpIdx} not found (total: {groups.Count})");

            var grp = groups[grpIdx - 1];
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "name":
                        var nvGrpPr = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
                        if (nvGrpPr != null) nvGrpPr.Name = value;
                        break;
                    case "x" or "y" or "width" or "height":
                    {
                        var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                        var xfrm = grpSpPr.TransformGroup ?? (grpSpPr.TransformGroup = new Drawing.TransformGroup());
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
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(grp, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
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
