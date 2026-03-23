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
        if (path.Equals("/theme", StringComparison.OrdinalIgnoreCase))
            return SetThemeProperties(properties);

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
                            case "a3":
                                sz.Cx = 15120000; sz.Cy = 10692000;
                                sz.Type = SlideSizeValues.A3;
                                break;
                            case "letter":
                                sz.Cx = 9144000; sz.Cy = 6858000;
                                sz.Type = SlideSizeValues.Letter;
                                break;
                            case "b4":
                                sz.Cx = 11430000; sz.Cy = 8574000;
                                sz.Type = SlideSizeValues.B4ISO;
                                break;
                            case "b5":
                                sz.Cx = 8208000; sz.Cy = 5760000;
                                sz.Type = SlideSizeValues.B5ISO;
                                break;
                            case "35mm":
                                sz.Cx = 10287000; sz.Cy = 6858000;
                                sz.Type = SlideSizeValues.Film35mm;
                                break;
                            case "overhead":
                                sz.Cx = 9144000; sz.Cy = 6858000;
                                sz.Type = SlideSizeValues.Overhead;
                                break;
                            case "banner":
                                sz.Cx = 7315200; sz.Cy = 914400;
                                sz.Type = SlideSizeValues.Banner;
                                break;
                            case "ledger":
                                sz.Cx = 12192000; sz.Cy = 9144000;
                                sz.Type = SlideSizeValues.Ledger;
                                break;
                            default:
                                unsupported.Add(key);
                                break;
                        }
                        break;
                    default:
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid presentation props: slideWidth, slideHeight, slideSize, width, height)");
                        else
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
                    throw new ArgumentException($"SlideMaster {partIdx} not found (total: {masters.Count})");
                rootEl = masters[partIdx - 1].SlideMaster
                    ?? throw new InvalidOperationException("Corrupt slide master");
            }
            else
            {
                var layouts = presentationPart.SlideMasterParts
                    .SelectMany(m => m.SlideLayoutParts).ToList();
                if (partIdx < 1 || partIdx > layouts.Count)
                    throw new ArgumentException($"SlideLayout {partIdx} not found (total: {layouts.Count})");
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
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid master/layout props: name)");
                        else
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

            throw new ArgumentException($"Unsupported element type: '{elType}' for master/layout. Valid types: shape.");
        }

        // Try notes path: /slide[N]/notes
        var notesSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesSetMatch.Success)
        {
            var slideIdx = int.Parse(notesSetMatch.Groups[1].Value);
            var slidePartsN = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slidePartsN.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slidePartsN.Count})");
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
            var linkVal = properties.GetValueOrDefault("link");
            var runOnlyProps = properties
                .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase))
                .ToDictionary(kv => kv.Key, kv => kv.Value);
            var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart);
            if (linkVal != null) ApplyRunHyperlink(slidePart, targetRun, linkVal);
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
                        var (lsVal2, lsIsPercent) = SpacingConverter.ParsePptLineSpacing(value);
                        if (lsIsPercent)
                            pProps.AppendChild(new Drawing.LineSpacing(
                                new Drawing.SpacingPercent { Val = lsVal2 }));
                        else
                            pProps.AppendChild(new Drawing.LineSpacing(
                                new Drawing.SpacingPoints { Val = lsVal2 }));
                        break;
                    }
                    case "spacebefore" or "space.before":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                        pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
                        break;
                    }
                    case "spaceafter" or "space.after":
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                        pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
                        break;
                    }
                    case "link":
                        foreach (var r in paraRuns)
                            ApplyRunHyperlink(slidePart, r, value);
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
                    {
                        var xfrm = chartGf.Transform ?? (chartGf.Transform = new Transform());
                        TryApplyPositionSize(key.ToLowerInvariant(), value,
                            xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                            xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                        break;
                    }
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
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts2.Count})");

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
                        TryApplyPositionSize(key.ToLowerInvariant(), value,
                            xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                            xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
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
                    case var k when k.StartsWith("border") || k is "text" or "bold" or "italic" or "size" or "font" or "color" or "underline" or "strike" or "valign" or "fill" or "baseline" or "charspacing":
                    {
                        // Apply cell-level properties to all cells in the table
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table != null)
                        {
                            foreach (var cell in table.Descendants<Drawing.TableCell>())
                            {
                                var u = SetTableCellProperties(cell, new Dictionary<string, string> { { key, value } });
                                foreach (var uk in u) { if (!unsupported.Contains(uk)) unsupported.Add(uk); }
                            }
                        }
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(gf, key, value))
                        {
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid table props: x, y, width, height, name, style, firstrow, lastrow, bandedrows, bandedcols)");
                            else
                                unsupported.Add(key);
                        }
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
                        // c1, c2, ... shorthand: set text of specific cell by index
                        if (key.Length >= 2 && key[0] == 'c' && int.TryParse(key.AsSpan(1), out var cIdx))
                        {
                            var rowCells = row.Elements<Drawing.TableCell>().ToList();
                            if (cIdx < 1 || cIdx > rowCells.Count)
                                throw new ArgumentException($"Cell c{cIdx} out of range (row has {rowCells.Count} cells)");
                            var targetCell = rowCells[cIdx - 1];
                            // Replace text in first paragraph's first run, or create one
                            var txBody = targetCell.TextBody;
                            if (txBody == null)
                            {
                                txBody = new Drawing.TextBody(
                                    new Drawing.BodyProperties(),
                                    new Drawing.ListStyle(),
                                    new Drawing.Paragraph());
                                targetCell.AppendChild(txBody);
                            }
                            var para = txBody.Elements<Drawing.Paragraph>().FirstOrDefault()
                                ?? txBody.AppendChild(new Drawing.Paragraph());
                            para.RemoveAllChildren<Drawing.Run>();
                            para.RemoveAllChildren<Drawing.Break>();
                            // Remove EndParagraphRunProperties before appending Run,
                            // then re-add after — schema requires Run before EndParagraphRunProperties
                            var savedEndParaRPr = para.Elements<Drawing.EndParagraphRunProperties>().FirstOrDefault();
                            if (savedEndParaRPr != null)
                                savedEndParaRPr.Remove();
                            if (!string.IsNullOrEmpty(value))
                            {
                                var newRun = new Drawing.Run(
                                    new Drawing.RunProperties { Language = "en-US" },
                                    new Drawing.Text(value));
                                para.AppendChild(newRun);
                            }
                            if (savedEndParaRPr != null)
                                para.AppendChild(savedEndParaRPr);
                        }
                        else
                        {
                            // Apply to all cells in this row
                            var cellUnsup = new HashSet<string>();
                            foreach (var cell in row.Elements<Drawing.TableCell>())
                            {
                                var u = SetTableCellProperties(cell, new Dictionary<string, string> { { key, value } });
                                foreach (var k in u) cellUnsup.Add(k);
                            }
                            unsupported.AddRange(cellUnsup);
                        }
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
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts2.Count})");
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
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts4.Count})");

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
                        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var volVal)
                            || double.IsNaN(volVal) || double.IsInfinity(volVal))
                            throw new ArgumentException($"Invalid volume value: '{value}'. Expected a finite number (0-100).");
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
                        TryApplyPositionSize(key.ToLowerInvariant(), value,
                            xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                            xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                        break;
                    }
                    default:
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid media props: volume, autoplay, trimstart, trimend, x, y, width, height)");
                        else
                            unsupported.Add(key);
                        break;
                }
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try picture path: /slide[N]/picture[M] or /slide[N]/pic[M]
        var picSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:picture|pic)\[(\d+)\]$");
        if (picSetMatch.Success)
        {
            var slideIdx = int.Parse(picSetMatch.Groups[1].Value);
            var picIdx = int.Parse(picSetMatch.Groups[2].Value);

            var slideParts3 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts3.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts3.Count})");

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
                        TryApplyPositionSize(key.ToLowerInvariant(), value,
                            xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                            xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                        break;
                    }
                    case "path" or "src":
                    {
                        // Replace image source
                        var blipFill = pic.BlipFill;
                        var blip = blipFill?.GetFirstChild<Drawing.Blip>();
                        if (blip == null) { unsupported.Add(key); break; }
                        var (imgStream, imgType) = OfficeCli.Core.ImageSource.Resolve(value);
                        using var imgStreamDispose2 = imgStream;
                        // Remove old image part to avoid storage bloat
                        var oldEmbedId = blip.Embed?.Value;
                        if (oldEmbedId != null)
                        {
                            try { slidePart.DeletePart(oldEmbedId); } catch { }
                        }
                        var newImgPart = slidePart.AddImagePart(imgType);
                        newImgPart.FeedData(imgStream);
                        blip.Embed = slidePart.GetIdOfPart(newImgPart);
                        break;
                    }
                    case "rotation" or "rotate":
                    {
                        var spPr = pic.ShapeProperties ?? (pic.ShapeProperties = new ShapeProperties());
                        var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                        xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
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
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid picture props: path, src, x, y, width, height, rotation, opacity, name, crop, cropleft, croptop, cropright, cropbottom)");
                        else
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
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts2.Count})");
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
                        SetAdvanceTime(slide2, value);
                        break;
                    case "advanceclick" or "advanceonclick":
                        SetAdvanceClick(slide2, IsTruthy(value));
                        break;
                    case "notes":
                    {
                        var notesPart = EnsureNotesSlidePart(slidePart2);
                        SetNotesText(notesPart, value);
                        break;
                    }
                    case "align":
                    {
                        var targets = properties.GetValueOrDefault("targets");
                        AlignShapes(slidePart2, value, targets);
                        break;
                    }
                    case "distribute":
                    {
                        var targets = properties.GetValueOrDefault("targets");
                        DistributeShapes(slidePart2, value, targets);
                        break;
                    }
                    case "targets":
                        break; // consumed by align/distribute
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(slide2, key, value))
                        {
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid slide props: background, layout, transition, name, align, distribute, targets)");
                            else
                                unsupported.Add(key);
                        }
                        break;
                }
            }
            slide2.Save();
            return unsupported;
        }

        // Try model3d-level path: /slide[N]/model3d[M]
        var model3dSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/model3d\[(\d+)\]$");
        if (model3dSetMatch.Success)
        {
            var slideIdx = int.Parse(model3dSetMatch.Groups[1].Value);
            var m3dIdx = int.Parse(model3dSetMatch.Groups[2].Value);
            var m3dSlideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > m3dSlideParts.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {m3dSlideParts.Count})");
            var m3dSlidePart = m3dSlideParts[slideIdx - 1];
            var m3dShapeTree = GetSlide(m3dSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
            var model3dElements = GetModel3DElements(m3dShapeTree);
            if (m3dIdx < 1 || m3dIdx > model3dElements.Count)
                throw new ArgumentException($"3D model {m3dIdx} not found (total: {model3dElements.Count})");

            var acElement = model3dElements[m3dIdx - 1];
            var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
            var fallback = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
            var sp = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame")
                  ?? choice?.ChildElements.FirstOrDefault(e => e.LocalName == "sp");

            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "x" or "y" or "width" or "height":
                    {
                        var emu = ParseEmu(value);
                        // Update xfrm (graphicFrame level or spPr level)
                        var xfrmEl = sp?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                        if (xfrmEl == null)
                        {
                            var spPr = sp?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
                            xfrmEl = spPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                        }
                        if (xfrmEl != null)
                        {
                            if (key.ToLowerInvariant() is "x" or "y")
                            {
                                var off = xfrmEl.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                                off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                            }
                            else
                            {
                                var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                                var ext = xfrmEl.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                                ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                            }
                        }
                        // Also update fallback pic spPr
                        var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                        var fbSpPr = fbPic?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
                        var fbXfrm = fbSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                        if (fbXfrm != null)
                        {
                            if (key.ToLowerInvariant() is "x" or "y")
                            {
                                var off = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                                off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                            }
                            else
                            {
                                var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                                var ext = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                                ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                            }
                        }
                        break;
                    }
                    case "name":
                    {
                        var nvSpPr = sp?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr")
                                  ?? sp?.ChildElements.FirstOrDefault(e => e.LocalName == "nvSpPr");
                        var cNvPr = nvSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
                        cNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                        // Also update fallback name
                        var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                        var fbCNvPr = fbPic?.Descendants().FirstOrDefault(d => d.LocalName == "cNvPr");
                        fbCNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                        break;
                    }
                    case "rotx" or "roty" or "rotz":
                    {
                        var model3dEl = acElement.Descendants().FirstOrDefault(d => d.LocalName == "model3d");
                        var trans = model3dEl?.ChildElements.FirstOrDefault(e => e.LocalName == "trans");
                        if (trans != null)
                        {
                            var rot = trans.ChildElements.FirstOrDefault(e => e.LocalName == "rot");
                            if (rot == null)
                            {
                                rot = new OpenXmlUnknownElement("am3d", "rot", Am3dNs);
                                trans.AppendChild(rot);
                            }
                            var attrName = key.ToLowerInvariant() switch { "rotx" => "ax", "roty" => "ay", _ => "az" };
                            rot.SetAttribute(new OpenXmlAttribute("", attrName, null!, ParseAngle60k(value).ToString()));
                        }
                        break;
                    }
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            GetSlide(m3dSlidePart).Save();
            return unsupported;
        }

        // Try zoom-level path: /slide[N]/zoom[M]
        var zoomSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/zoom\[(\d+)\]$");
        if (zoomSetMatch.Success)
        {
            var slideIdx = int.Parse(zoomSetMatch.Groups[1].Value);
            var zmIdx = int.Parse(zoomSetMatch.Groups[2].Value);
            var zmSlideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > zmSlideParts.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {zmSlideParts.Count})");
            var zmSlidePart = zmSlideParts[slideIdx - 1];
            var zmShapeTree = GetSlide(zmSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
            var zoomElements = GetZoomElements(zmShapeTree);
            if (zmIdx < 1 || zmIdx > zoomElements.Count)
                throw new ArgumentException($"Zoom {zmIdx} not found (total: {zoomElements.Count})");

            var acElement = zoomElements[zmIdx - 1];
            var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
            var fallback = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
            var gf = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame");
            var sldZmObj = acElement.Descendants().FirstOrDefault(d => d.LocalName == "sldZmObj");
            var zmPr = acElement.Descendants().FirstOrDefault(d => d.LocalName == "zmPr");

            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "target" or "slide":
                    {
                        if (!int.TryParse(value, out var targetNum))
                            throw new ArgumentException($"Invalid target value: '{value}'. Expected a slide number.");
                        if (targetNum < 1 || targetNum > zmSlideParts.Count)
                            throw new ArgumentException($"Target slide {targetNum} not found (total: {zmSlideParts.Count})");
                        var zmPresentation = _doc.PresentationPart?.Presentation
                            ?? throw new InvalidOperationException("No presentation");
                        var zmSlideIds = zmPresentation.GetFirstChild<SlideIdList>()
                            ?.Elements<SlideId>().ToList()
                            ?? throw new InvalidOperationException("No slides");
                        var newSldId = zmSlideIds[targetNum - 1].Id!.Value;
                        sldZmObj?.SetAttribute(new OpenXmlAttribute("", "sldId", null!, newSldId.ToString()));

                        // Update fallback hyperlink relationship
                        var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                        var fbCNvPr = fbPic?.Descendants().FirstOrDefault(d => d.LocalName == "cNvPr");
                        var hlinkClick = fbCNvPr?.ChildElements.FirstOrDefault(e => e.LocalName == "hlinkClick");
                        if (hlinkClick != null)
                        {
                            var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                            var targetSlidePart = zmSlideParts[targetNum - 1];
                            var newRelId = zmSlidePart.CreateRelationshipToPart(targetSlidePart);
                            hlinkClick.SetAttribute(new OpenXmlAttribute("r", "id", rNs, newRelId));
                        }
                        break;
                    }
                    case "returntoparent":
                        zmPr?.SetAttribute(new OpenXmlAttribute("", "returnToParent", null!, IsTruthy(value) ? "1" : "0"));
                        break;
                    case "transitiondur":
                        zmPr?.SetAttribute(new OpenXmlAttribute("", "transitionDur", null!, value));
                        break;
                    case "x" or "y" or "width" or "height":
                    {
                        var emu = ParseEmu(value);
                        // Update graphicFrame xfrm
                        var gfXfrm = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                        if (gfXfrm != null)
                        {
                            if (key.ToLowerInvariant() is "x" or "y")
                            {
                                var off = gfXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                                off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                            }
                            else
                            {
                                var ext = gfXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                                var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                                ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                            }
                        }
                        // Update fallback spPr xfrm
                        var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                        var fbSpPr = fbPic?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
                        var fbXfrm = fbSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                        if (fbXfrm != null)
                        {
                            if (key.ToLowerInvariant() is "x" or "y")
                            {
                                var off = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                                off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                            }
                            else
                            {
                                var ext = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                                var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                                ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                            }
                        }
                        // Update inner zmPr > spPr > xfrm (only for width/height)
                        if (key.ToLowerInvariant() is "width" or "height")
                        {
                            var p166Ns = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
                            var zmSpPr = zmPr?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr" && e.NamespaceUri == p166Ns);
                            var zmSpXfrm = zmSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                            var zmSpExt = zmSpXfrm?.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                            var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                            zmSpExt?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                        }
                        break;
                    }
                    case "name":
                    {
                        // Update cNvPr name in Choice
                        var nvGfPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr");
                        var choiceCNvPr = nvGfPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
                        choiceCNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                        // Update cNvPr name in Fallback
                        var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                        var fbNvPicPr = fbPic?.ChildElements.FirstOrDefault(e => e.LocalName == "nvPicPr");
                        var fbCNvPr = fbNvPicPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
                        fbCNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                        break;
                    }
                    case "image" or "path" or "src" or "cover":
                    {
                        var (zmImgStream, zmImgPartType) = OfficeCli.Core.ImageSource.Resolve(value);
                        using var zmImgDispose = zmImgStream;
                        // Add new image part
                        var newImagePart = zmSlidePart.AddImagePart(zmImgPartType);
                        newImagePart.FeedData(zmImgStream);
                        var newImgRelId = zmSlidePart.GetIdOfPart(newImagePart);
                        var rNs2 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                        // Update blip in zmPr > blipFill
                        var zmBlip = zmPr?.Descendants().FirstOrDefault(d => d.LocalName == "blip");
                        zmBlip?.SetAttribute(new OpenXmlAttribute("r", "embed", rNs2, newImgRelId));
                        // Update blip in fallback > blipFill
                        var fbBlipFill = fallback?.Descendants().FirstOrDefault(d => d.LocalName == "blipFill");
                        var fbBlip = fbBlipFill?.ChildElements.FirstOrDefault(e => e.LocalName == "blip");
                        fbBlip?.SetAttribute(new OpenXmlAttribute("r", "embed", rNs2, newImgRelId));
                        // Set imageType to "cover" so PowerPoint uses our image instead of auto-preview
                        zmPr?.SetAttribute(new OpenXmlAttribute("", "imageType", null!, "cover"));
                        break;
                    }
                    case "imagetype":
                        zmPr?.SetAttribute(new OpenXmlAttribute("", "imageType", null!, value));
                        break;
                    default:
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid zoom props: target, image, src, path, imagetype, x, y, width, height)");
                        else
                            unsupported.Add(key);
                        break;
                }
            }
            GetSlide(zmSlidePart).Save();
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

                // Separate animation, motionPath, link, and z-order from other shape properties
                var animValue = properties.GetValueOrDefault("animation")
                    ?? properties.GetValueOrDefault("animate");
                var motionPathValue = properties.GetValueOrDefault("motionpath")
                    ?? properties.GetValueOrDefault("motionPath");
                var linkValue = properties.GetValueOrDefault("link");
                var excludeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "animation", "animate", "motionpath", "motionPath", "link", "zorder", "z-order", "order" };
                var shapeProps = properties
                    .Where(kv => !excludeKeys.Contains(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);

                var unsupported = SetRunOrShapeProperties(shapeProps, allRuns, shape, slidePart);

                if (animValue != null)
                    ApplyShapeAnimation(slidePart, shape, animValue);
                if (motionPathValue != null)
                    ApplyMotionPathAnimation(slidePart, shape, motionPathValue);
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

        // Try connector path: /slide[N]/connector[M] or /slide[N]/connection[M]
        var cxnMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:connector|connection)\[(\d+)\]$");
        if (cxnMatch.Success)
        {
            var slideIdx = int.Parse(cxnMatch.Groups[1].Value);
            var cxnIdx = int.Parse(cxnMatch.Groups[2].Value);

            var slideParts5 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts5.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts5.Count})");

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
                        TryApplyPositionSize(key.ToLowerInvariant(), value,
                            xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                            xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                        break;
                    }
                    case "linewidth" or "line.width":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        outline.Width = Core.EmuConverter.ParseLineWidth(value);
                        break;
                    }
                    case "linecolor" or "line.color" or "line" or "color":
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
                    case "linedash" or "line.dash":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        outline.RemoveAllChildren<Drawing.PresetDash>();
                        outline.AppendChild(new Drawing.PresetDash { Val = value.ToLowerInvariant() switch
                        {
                            "solid" => Drawing.PresetLineDashValues.Solid,
                            "dot" => Drawing.PresetLineDashValues.Dot,
                            "dash" => Drawing.PresetLineDashValues.Dash,
                            "dashdot" or "dash_dot" => Drawing.PresetLineDashValues.DashDot,
                            "longdash" or "lgdash" or "lg_dash" => Drawing.PresetLineDashValues.LargeDash,
                            "longdashdot" or "lgdashdot" or "lg_dash_dot" => Drawing.PresetLineDashValues.LargeDashDot,
                            _ => throw new ArgumentException($"Invalid 'lineDash' value: '{value}'. Valid values: solid, dot, dash, dashdot, longdash, longdashdot.")
                        }});
                        break;
                    }
                    case "lineopacity" or "line.opacity":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var lnOpacity)
                            || double.IsNaN(lnOpacity) || double.IsInfinity(lnOpacity))
                            throw new ArgumentException($"Invalid 'lineOpacity' value: '{value}'. Expected a finite decimal 0.0-1.0.");
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        var solidFill = outline.GetFirstChild<Drawing.SolidFill>();
                        if (solidFill == null)
                        {
                            // Auto-create a black line fill (matching Apache POI behavior)
                            solidFill = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" });
                            outline.PrependChild(solidFill);
                        }
                        {
                            var colorEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                                ?? solidFill.GetFirstChild<Drawing.SchemeColor>();
                            if (colorEl != null)
                            {
                                colorEl.RemoveAllChildren<Drawing.Alpha>();
                                colorEl.AppendChild(new Drawing.Alpha { Val = (int)(lnOpacity * 100000) });
                            }
                        }
                        break;
                    }
                    case "rotation" or "rotate":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                        xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
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
                    case "headend" or "headEnd":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        outline.RemoveAllChildren<Drawing.HeadEnd>();
                        outline.AppendChild(new Drawing.HeadEnd { Type = ParseLineEndType(value) });
                        break;
                    }
                    case "tailend" or "tailEnd":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        outline.RemoveAllChildren<Drawing.TailEnd>();
                        outline.AppendChild(new Drawing.TailEnd { Type = ParseLineEndType(value) });
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(cxn, key, value))
                        {
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid connector props: line, color, fill, x, y, width, height, rotation, name, headEnd, tailEnd, geometry)");
                            else
                                unsupported.Add(key);
                        }
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
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts6.Count})");

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
                        TryApplyPositionSize(key.ToLowerInvariant(), value,
                            xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                            xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                        break;
                    }
                    case "rotation" or "rotate":
                    {
                        var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                        var xfrm = grpSpPr.TransformGroup ?? (grpSpPr.TransformGroup = new Drawing.TransformGroup());
                        xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
                        break;
                    }
                    case "fill":
                    {
                        var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                        grpSpPr.RemoveAllChildren<Drawing.SolidFill>();
                        grpSpPr.RemoveAllChildren<Drawing.NoFill>();
                        grpSpPr.RemoveAllChildren<Drawing.GradientFill>();
                        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            grpSpPr.AppendChild(new Drawing.NoFill());
                        else
                            grpSpPr.AppendChild(BuildSolidFill(value));
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(grp, key, value))
                        {
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid group props: x, y, width, height, rotation, name, fill)");
                            else
                                unsupported.Add(key);
                        }
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
                    throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

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
