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
        path = NormalizeCellPath(path);
        path = ResolveIdPath(path);

        // Batch Set: if path looks like a selector (not starting with /), Query → Set each
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
        {
            var unsupported = new List<string>();
            var targets = Query(path);
            if (targets.Count == 0)
                throw new ArgumentException($"No elements matched selector: {path}");
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        if (path.Equals("/theme", StringComparison.OrdinalIgnoreCase))
            return SetThemeProperties(properties);

        // Unified find: if 'find' key is present, route to ProcessPptFind
        if (properties.TryGetValue("find", out var findText))
        {
            var replace = properties.TryGetValue("replace", out var r) ? r : null;
            var formatProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            formatProps.Remove("find");
            formatProps.Remove("replace");
            formatProps.Remove("scope");
            formatProps.Remove("regex");

            if (replace == null && formatProps.Count == 0)
                throw new ArgumentException("'find' requires either 'replace' and/or format properties (e.g. bold, color, size).");

            // Support regex=true as an alternative to r"..." prefix.
            // CONSISTENCY(find-regex): mirror of WordHandler.Set.cs:60-61. grep
            // "CONSISTENCY(find-regex)" for every project-wide call site.
            if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
                findText = $"r\"{findText}\"";

            var matchCount = ProcessPptFind(path, findText, replace, formatProps);
            LastFindMatchCount = matchCount;
            return [];
        }

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
                        if (SlideSizeDefaults.Presets.TryGetValue(value, out var preset))
                        {
                            sz.Cx = (int)preset.Cx;
                            sz.Cy = (int)preset.Cy;
                            sz.Type = preset.Type;
                        }
                        else
                        {
                            unsupported.Add(key);
                        }
                        break;
                    // Core document properties
                    case "title":
                        _doc.PackageProperties.Title = value;
                        break;
                    case "author" or "creator":
                        _doc.PackageProperties.Creator = value;
                        break;
                    case "subject":
                        _doc.PackageProperties.Subject = value;
                        break;
                    case "description":
                        _doc.PackageProperties.Description = value;
                        break;
                    case "category":
                        _doc.PackageProperties.Category = value;
                        break;
                    case "keywords":
                        _doc.PackageProperties.Keywords = value;
                        break;
                    case "lastmodifiedby":
                        _doc.PackageProperties.LastModifiedBy = value;
                        break;
                    case "revision":
                        _doc.PackageProperties.Revision = value;
                        break;
                    case "defaultfont" or "font":
                    {
                        var masterPart = _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault();
                        var theme = masterPart?.ThemePart?.Theme;
                        var fontScheme = theme?.ThemeElements?.FontScheme;
                        if (fontScheme != null)
                        {
                            if (fontScheme.MajorFont?.LatinFont != null)
                                fontScheme.MajorFont.LatinFont.Typeface = value;
                            if (fontScheme.MinorFont?.LatinFont != null)
                                fontScheme.MinorFont.LatinFont.Typeface = value;
                            masterPart!.ThemePart!.Theme!.Save();
                        }
                        break;
                    }
                    default:
                        var lowerKey = key.ToLowerInvariant();
                        if (!TrySetPresentationSetting(lowerKey, value)
                            && !Core.ThemeHandler.TrySetTheme(
                                _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault()?.ThemePart, lowerKey, value)
                            && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                                Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                        {
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid presentation props: slideWidth, slideHeight, slideSize, title, author, defaultFont, firstSlideNum, rtl, compatMode, print.*, show.*)");
                            else
                                unsupported.Add(key);
                        }
                        break;
                }
            }
            presentation.Save();
            return unsupported;
        }

        // Try slidemaster/slidelayout bg-aware path first (case-insensitive):
        // /slidemaster[N], /slidemaster[N]/slidelayout[M], /slidelayout[N]
        // Handles background and name props. Falls through for shape-nested paths.
        {
            var masterBgMatch = Regex.Match(path, @"^/slidemaster\[(\d+)\](?:/slidelayout\[(\d+)\])?$", RegexOptions.IgnoreCase);
            var layoutBgMatch = Regex.Match(path, @"^/slidelayout\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (masterBgMatch.Success || layoutBgMatch.Success)
            {
                OpenXmlPart targetPart;
                OpenXmlPartRootElement targetRoot;
                if (masterBgMatch.Success)
                {
                    var masterIdx = int.Parse(masterBgMatch.Groups[1].Value);
                    var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
                    if (masterIdx < 1 || masterIdx > masters.Count)
                        throw new ArgumentException($"Slide master {masterIdx} not found (total: {masters.Count})");
                    var mp = masters[masterIdx - 1];
                    if (masterBgMatch.Groups[2].Success)
                    {
                        var lIdx = int.Parse(masterBgMatch.Groups[2].Value);
                        var layouts = mp.SlideLayoutParts?.ToList() ?? [];
                        if (lIdx < 1 || lIdx > layouts.Count)
                            throw new ArgumentException($"Slide layout {lIdx} not found under master {masterIdx} (total: {layouts.Count})");
                        targetPart = layouts[lIdx - 1];
                        targetRoot = layouts[lIdx - 1].SlideLayout
                            ?? throw new InvalidOperationException("Corrupt slide layout");
                    }
                    else
                    {
                        targetPart = mp;
                        targetRoot = mp.SlideMaster
                            ?? throw new InvalidOperationException("Corrupt slide master");
                    }
                }
                else
                {
                    var lIdx = int.Parse(layoutBgMatch.Groups[1].Value);
                    var allLayouts = (_doc.PresentationPart?.SlideMasterParts ?? Enumerable.Empty<SlideMasterPart>())
                        .SelectMany(m => m.SlideLayoutParts ?? Enumerable.Empty<SlideLayoutPart>()).ToList();
                    if (lIdx < 1 || lIdx > allLayouts.Count)
                        throw new ArgumentException($"Slide layout {lIdx} not found (total: {allLayouts.Count})");
                    targetPart = allLayouts[lIdx - 1];
                    targetRoot = allLayouts[lIdx - 1].SlideLayout
                        ?? throw new InvalidOperationException("Corrupt slide layout");
                }

                var unsupported = new List<string>();
                foreach (var (key, value) in properties)
                {
                    switch (key.ToLowerInvariant())
                    {
                        case "background":
                            ApplyBackground(targetPart, value, ReadBackgroundImageOptions(properties));
                            break;
                        case "background.mode":
                        case "background.alpha":
                        case "background.scale":
                            break;
                        case "name":
                        {
                            var csd = targetRoot.GetFirstChild<CommonSlideData>();
                            if (csd != null) csd.Name = value;
                            break;
                        }
                        default:
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid slidemaster/slidelayout props: background, background.mode, background.alpha, background.scale, name)");
                            else
                                unsupported.Add(key);
                            break;
                    }
                }
                MaybeMutateExistingBackgroundImage(targetPart, properties);
                SaveBackgroundRoot(targetPart);
                return unsupported;
            }
        }

        // Try slideMaster/slideLayout shape editing: /slideMaster[N]/shape[M] or /slideLayout[N]/shape[M]
        var masterShapeMatch = Regex.Match(path, @"^/(slideMaster|slideLayout)\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (masterShapeMatch.Success) return SetMasterShapeByPath(masterShapeMatch, properties);

        // Try notes path: /slide[N]/notes
        var notesSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesSetMatch.Success) return SetNotesByPath(notesSetMatch, properties);

        // Try run-level path: /slide[N]/shape[M]/run[K]
        var runMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/run\[(\d+)\]$");
        if (runMatch.Success) return SetShapeRunByPath(runMatch, properties);

        // Try paragraph/run path: /slide[N]/shape[M]/paragraph[P]/run[K]
        var paraRunMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]/run\[(\d+)\]$");
        if (paraRunMatch.Success) return SetParagraphRunByPath(paraRunMatch, properties);

        // Try paragraph-level path: /slide[N]/shape[M]/paragraph[P]
        var paraMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]$");
        if (paraMatch.Success) return SetParagraphByPath(paraMatch, properties);

        // Try chart axis-by-role sub-path: /slide[N]/chart[M]/axis[@role=ROLE].
        // Routed separately from the chart[]/series[] path because the role capture
        // needs to drive a different forwarder (SetAxisProperties, not series-prefix).
        var chartAxisSetMatch = Regex.Match(path,
            @"^/slide\[(\d+)\]/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success) return SetChartAxisByPath(chartAxisSetMatch, properties);

        // Try chart path: /slide[N]/chart[M] or /slide[N]/chart[M]/series[K]
        var chartSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartSetMatch.Success) return SetChartByPath(chartSetMatch, properties);

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellMatch.Success) return SetTableCellByPath(tblCellMatch, properties);

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
                            var styleId = ResolveTableStyleId(value);
                            tblPr.RemoveAllChildren<Drawing.TableStyleId>();
                            tblPr.AppendChild(new Drawing.TableStyleId(styleId));
                        }
                        break;
                    }
                    case "firstrow":
                    case "lastrow":
                    case "firstcol" or "firstcolumn":
                    case "lastcol" or "lastcolumn":
                    case "bandrow" or "bandedrows" or "bandrows":
                    case "bandcol" or "bandedcols" or "bandcols":
                    {
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table != null)
                        {
                            var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                                ?? table.PrependChild(new Drawing.TableProperties());
                            var bv = IsTruthy(value);
                            switch (key.ToLowerInvariant())
                            {
                                case "firstrow": tblPr.FirstRow = bv; break;
                                case "lastrow": tblPr.LastRow = bv; break;
                                case "firstcol" or "firstcolumn": tblPr.FirstColumn = bv; break;
                                case "lastcol" or "lastcolumn": tblPr.LastColumn = bv; break;
                                case "bandrow" or "bandedrows" or "bandrows": tblPr.BandRow = bv; break;
                                case "bandcol" or "bandedcols" or "bandcols": tblPr.BandColumn = bv; break;
                            }
                        }
                        break;
                    }
                    case "colwidth" or "colwidths":
                    {
                        // Set individual column widths: "3cm,5cm,3cm" or single value for all
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table != null)
                        {
                            var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
                            if (gridCols != null && gridCols.Count > 0)
                            {
                                var widths = value.Split(',').Select(w => ParseEmu(w.Trim())).ToArray();
                                for (int ci = 0; ci < gridCols.Count; ci++)
                                    gridCols[ci].Width = ci < widths.Length ? widths[ci] : widths[^1];
                            }
                        }
                        break;
                    }
                    case "autofit" or "autowidth":
                    {
                        // Heuristic auto column width: measure max text length per column
                        if (!IsTruthy(value)) break;
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table == null) break;
                        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
                        var tableRows = table.Elements<Drawing.TableRow>().ToList();
                        if (gridCols == null || gridCols.Count == 0 || tableRows.Count == 0) break;

                        var totalWidth = gridCols.Sum(gc => gc.Width?.Value ?? 0);
                        var colCount = gridCols.Count;
                        var maxLens = new int[colCount];
                        foreach (var row in tableRows)
                        {
                            var cells = row.Elements<Drawing.TableCell>().ToList();
                            for (int ci = 0; ci < Math.Min(cells.Count, colCount); ci++)
                            {
                                var text = cells[ci].TextBody?.InnerText ?? "";
                                maxLens[ci] = Math.Max(maxLens[ci], text.Length);
                            }
                        }
                        var totalLen = maxLens.Sum();
                        if (totalLen == 0) break;
                        // Minimum 10% per column, distribute rest by text length
                        var minWidth = totalWidth * 0.1 / colCount;
                        var distributable = totalWidth - minWidth * colCount;
                        for (int ci = 0; ci < colCount; ci++)
                            gridCols[ci].Width = (long)(minWidth + distributable * maxLens[ci] / totalLen);
                        break;
                    }
                    case "shadow":
                    {
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table != null)
                        {
                            var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                                ?? table.PrependChild(new Drawing.TableProperties());
                            var effectList = tblPr.GetFirstChild<Drawing.EffectList>();
                            if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            {
                                effectList?.RemoveAllChildren<Drawing.OuterShadow>();
                                if (effectList?.ChildElements.Count == 0) effectList.Remove();
                            }
                            else
                            {
                                if (effectList == null) effectList = tblPr.AppendChild(new Drawing.EffectList());
                                effectList.RemoveAllChildren<Drawing.OuterShadow>();
                                var shadow = OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(value, BuildColorElement);
                                InsertEffectInOrder(effectList, shadow);
                            }
                        }
                        break;
                    }
                    case "glow":
                    {
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table != null)
                        {
                            var tblPr = table.GetFirstChild<Drawing.TableProperties>()
                                ?? table.PrependChild(new Drawing.TableProperties());
                            var effectList = tblPr.GetFirstChild<Drawing.EffectList>();
                            if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            {
                                effectList?.RemoveAllChildren<Drawing.Glow>();
                                if (effectList?.ChildElements.Count == 0) effectList.Remove();
                            }
                            else
                            {
                                if (effectList == null) effectList = tblPr.AppendChild(new Drawing.EffectList());
                                effectList.RemoveAllChildren<Drawing.Glow>();
                                var glow = OfficeCli.Core.DrawingEffectsHelper.BuildGlow(value, BuildColorElement);
                                InsertEffectInOrder(effectList, glow);
                            }
                        }
                        break;
                    }
                    case "bandcolor.odd" or "bandcolor.even":
                    {
                        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                        if (table != null)
                        {
                            var isOdd = key.ToLowerInvariant().EndsWith("odd");
                            var rows = table.Elements<Drawing.TableRow>().ToList();
                            for (int ri = 0; ri < rows.Count; ri++)
                            {
                                bool matchesOddEven = isOdd ? (ri % 2 == 0) : (ri % 2 == 1); // 0-based: odd rows are 0,2,4...
                                if (matchesOddEven)
                                {
                                    foreach (var cell in rows[ri].Elements<Drawing.TableCell>())
                                        SetTableCellProperties(cell, new Dictionary<string, string> { { "fill", value } });
                                }
                            }
                        }
                        break;
                    }
                    case var k when k.StartsWith("border") || k is "text" or "bold" or "italic" or "size" or "font" or "color" or "underline" or "strike" or "valign" or "fill" or "baseline" or "charspacing" or "opacity" or "bevel" or "margin" or "padding" or "textdirection" or "wordwrap" or "linespacing" or "spacebefore" or "spaceafter":
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
                                unsupported.Add($"{key} (valid table props: x, y, width, height, name, style, firstRow, lastRow, firstCol, lastCol, bandedRows, bandedCols, colWidths)");
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
        if (tblRowMatch.Success) return SetTableRowByPath(tblRowMatch, properties);

        // Try placeholder path: /slide[N]/placeholder[M] or /slide[N]/placeholder[type]
        var phMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phMatch.Success) return SetPlaceholderByPath(phMatch, properties);

        // Try video/audio path: /slide[N]/video[M] or /slide[N]/audio[M]
        var mediaSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(video|audio)\[(\d+)\]$");
        if (mediaSetMatch.Success) return SetMediaByPath(mediaSetMatch, properties);

        // Try picture path: /slide[N]/picture[M] or /slide[N]/pic[M]
        // OLE set path: /slide[N]/ole[M]
        // Replace backing embedded part + refresh ProgID automatically
        // when the extension changes. Cleans up the old part to avoid
        // storage bloat (mirrors picture path clean-up).
        var oleSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:ole|object|embed)\[(\d+)\]$");
        if (oleSetMatch.Success) return SetOleByPath(oleSetMatch, properties);

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
                        // Remove old image part(s) to avoid storage bloat,
                        // including the asvg:svgBlip-referenced SVG part
                        // when the previous image was SVG.
                        var oldEmbedId = blip.Embed?.Value;
                        if (oldEmbedId != null)
                        {
                            try { slidePart.DeletePart(oldEmbedId); } catch { }
                        }
                        var oldPicSvgRelId = OfficeCli.Core.SvgImageHelper.GetSvgRelId(blip);
                        if (oldPicSvgRelId != null)
                        {
                            try { slidePart.DeletePart(oldPicSvgRelId); } catch { }
                        }

                        if (imgType == ImagePartType.Svg)
                        {
                            using var newSvgBuf = new MemoryStream();
                            imgStream.CopyTo(newSvgBuf);
                            newSvgBuf.Position = 0;
                            var newSvgPart = slidePart.AddImagePart(ImagePartType.Svg);
                            newSvgPart.FeedData(newSvgBuf);
                            var newPicSvgRelId = slidePart.GetIdOfPart(newSvgPart);

                            var pngFb = slidePart.AddImagePart(ImagePartType.Png);
                            pngFb.FeedData(new MemoryStream(
                                OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                            blip.Embed = slidePart.GetIdOfPart(pngFb);
                            OfficeCli.Core.SvgImageHelper.AppendSvgExtension(blip, newPicSvgRelId);
                        }
                        else
                        {
                            var newImgPart = slidePart.AddImagePart(imgType);
                            newImgPart.FeedData(imgStream);
                            blip.Embed = slidePart.GetIdOfPart(newImgPart);
                            if (oldPicSvgRelId != null)
                            {
                                var extLst = blip.GetFirstChild<Drawing.BlipExtensionList>();
                                if (extLst != null)
                                {
                                    foreach (var ext in extLst.Elements<Drawing.BlipExtension>().ToList())
                                    {
                                        if (string.Equals(ext.Uri?.Value,
                                            OfficeCli.Core.SvgImageHelper.SvgExtensionUri,
                                            StringComparison.OrdinalIgnoreCase))
                                            ext.Remove();
                                    }
                                    if (!extLst.Elements<Drawing.BlipExtension>().Any())
                                        extLst.Remove();
                                }
                            }
                        }
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
                        var srcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>();
                        if (srcRect == null)
                        {
                            srcRect = new Drawing.SourceRectangle();
                            // CONSISTENCY(ooxml-element-order): in CT_BlipFillProperties
                            // srcRect must precede the fill-mode element (stretch/tile).
                            // PowerPoint silently ignores out-of-order srcRect.
                            var fillMode = (OpenXmlElement?)blipFill.GetFirstChild<Drawing.Stretch>()
                                ?? blipFill.GetFirstChild<Drawing.Tile>();
                            if (fillMode != null)
                                blipFill.InsertBefore(srcRect, fillMode);
                            else
                                blipFill.AppendChild(srcRect);
                        }

                        if (key.Equals("crop", StringComparison.OrdinalIgnoreCase))
                        {
                            // Single value: "left,top,right,bottom" as percentages (0-100)
                            var parts = value.Split(',');
                            if (parts.Length == 4)
                            {
                                var cropVals = new double[4];
                                for (int ci = 0; ci < 4; ci++)
                                {
                                    cropVals[ci] = ParseHelpers.SafeParseDouble(parts[ci].Trim(), "crop");
                                    if (cropVals[ci] < 0 || cropVals[ci] > 100)
                                        throw new ArgumentException($"Invalid 'crop' value: '{parts[ci].Trim()}'. Crop percentage must be between 0 and 100.");
                                }
                                srcRect.Left = (int)(cropVals[0] * 1000);
                                srcRect.Top = (int)(cropVals[1] * 1000);
                                srcRect.Right = (int)(cropVals[2] * 1000);
                                srcRect.Bottom = (int)(cropVals[3] * 1000);
                            }
                            else if (parts.Length == 2)
                            {
                                // 2-value: vertical,horizontal (top/bottom, left/right)
                                var vCrop = ParseHelpers.SafeParseDouble(parts[0].Trim(), "crop");
                                var hCrop = ParseHelpers.SafeParseDouble(parts[1].Trim(), "crop");
                                if (vCrop < 0 || vCrop > 100 || hCrop < 0 || hCrop > 100)
                                    throw new ArgumentException($"Invalid 'crop' value: '{value}'. Crop percentages must be between 0 and 100.");
                                srcRect.Top = (int)(vCrop * 1000); srcRect.Bottom = (int)(vCrop * 1000);
                                srcRect.Left = (int)(hCrop * 1000); srcRect.Right = (int)(hCrop * 1000);
                            }
                            else if (parts.Length == 1)
                            {
                                if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cropVal))
                                    throw new ArgumentException($"Invalid 'crop' value: '{value}'. Expected a percentage (e.g. 10 = 10% from each edge).");
                                if (cropVal < 0 || cropVal > 100)
                                    throw new ArgumentException($"Invalid 'crop' value: '{value}'. Crop percentage must be between 0 and 100.");
                                var cropPct = (int)(cropVal * 1000);
                                srcRect.Left = cropPct; srcRect.Top = cropPct; srcRect.Right = cropPct; srcRect.Bottom = cropPct;
                            }
                            else
                            {
                                throw new ArgumentException($"Invalid 'crop' value: '{value}'. Expected 1 value (symmetric), 2 values (vertical,horizontal), or 4 values (left,top,right,bottom).");
                            }
                        }
                        else
                        {
                            if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cropSingle))
                                throw new ArgumentException($"Invalid '{key}' value: '{value}'. Expected a percentage (0-100).");
                            if (cropSingle < 0 || cropSingle > 100)
                                throw new ArgumentException($"Invalid '{key}' value: '{value}'. Crop percentage must be between 0 and 100.");
                            var pct = (int)(cropSingle * 1000); // percent (0-100) → 1/1000ths
                            switch (key.ToLowerInvariant())
                            {
                                case "cropleft": srcRect.Left = pct; break;
                                case "croptop": srcRect.Top = pct; break;
                                case "cropright": srcRect.Right = pct; break;
                                case "cropbottom": srcRect.Bottom = pct; break;
                            }
                        }
                        // Reset semantics: if all four sides are zero (or unset),
                        // drop the srcRect entirely so the XML is clean.
                        int L = srcRect.Left?.Value ?? 0;
                        int T = srcRect.Top?.Value ?? 0;
                        int R = srcRect.Right?.Value ?? 0;
                        int B = srcRect.Bottom?.Value ?? 0;
                        if (L == 0 && T == 0 && R == 0 && B == 0)
                            srcRect.Remove();
                        break;
                    }
                    case "opacity":
                    {
                        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var opacityVal)
                            || double.IsNaN(opacityVal) || double.IsInfinity(opacityVal))
                            throw new ArgumentException($"Invalid 'opacity' value: '{value}'. Expected a finite decimal 0.0-1.0.");
                        if (opacityVal > 1.0) opacityVal /= 100.0;
                        var blip = pic.BlipFill?.GetFirstChild<Drawing.Blip>();
                        if (blip != null)
                        {
                            blip.RemoveAllChildren<Drawing.AlphaModulationFixed>();
                            var alphaVal = (int)(opacityVal * 100000); // 0.0-1.0 → 0-100000
                            blip.AppendChild(new Drawing.AlphaModulationFixed { Amount = alphaVal });
                        }
                        break;
                    }
                    case "name":
                    {
                        var nvPr = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                        if (nvPr != null) nvPr.Name = value;
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
        if (slideOnlyMatch.Success) return SetSlideByPath(slideOnlyMatch, properties);

        // Try model3d-level path: /slide[N]/model3d[M]
        var model3dSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/model3d\[(\d+)\]$");
        if (model3dSetMatch.Success) return SetModel3DByPath(model3dSetMatch, properties);

        // Try zoom-level path: /slide[N]/zoom[M]
        var zoomSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/zoom\[(\d+)\]$");
        if (zoomSetMatch.Success) return SetZoomByPath(zoomSetMatch, properties);

        // Try shape-level path: /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
        if (match.Success) return SetShapeByPath(match, properties);

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
                        var newFill = new Drawing.SolidFill(
                            new Drawing.RgbColorModelHex { Val = rgb });
                        // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                        var prstDash = outline.GetFirstChild<Drawing.PresetDash>();
                        if (prstDash != null)
                            outline.InsertBefore(newFill, prstDash);
                        else
                        {
                            var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                            if (headEnd != null)
                                outline.InsertBefore(newFill, headEnd);
                            else
                            {
                                var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                                if (tailEnd != null)
                                    outline.InsertBefore(newFill, tailEnd);
                                else
                                    outline.AppendChild(newFill);
                            }
                        }
                        break;
                    }
                    case "fill":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        ApplyShapeFill(spPr, value);
                        break;
                    }
                    case "linedash" or "line.dash":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        outline.RemoveAllChildren<Drawing.PresetDash>();
                        var newDash = new Drawing.PresetDash { Val = value.ToLowerInvariant() switch
                        {
                            "solid" => Drawing.PresetLineDashValues.Solid,
                            "dot" => Drawing.PresetLineDashValues.Dot,
                            "dash" => Drawing.PresetLineDashValues.Dash,
                            "dashdot" or "dash_dot" => Drawing.PresetLineDashValues.DashDot,
                            "longdash" or "lgdash" or "lg_dash" => Drawing.PresetLineDashValues.LargeDash,
                            "longdashdot" or "lgdashdot" or "lg_dash_dot" => Drawing.PresetLineDashValues.LargeDashDot,
                            _ => throw new ArgumentException($"Invalid 'lineDash' value: '{value}'. Valid values: solid, dot, dash, dashdot, longdash, longdashdot.")
                        }};
                        // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                        var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                        if (headEnd != null)
                            outline.InsertBefore(newDash, headEnd);
                        else
                        {
                            var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                            if (tailEnd != null)
                                outline.InsertBefore(newDash, tailEnd);
                            else
                                outline.AppendChild(newDash);
                        }
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
                            // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                            solidFill = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" });
                            var prstDashEl = outline.GetFirstChild<Drawing.PresetDash>();
                            if (prstDashEl != null)
                                outline.InsertBefore(solidFill, prstDashEl);
                            else
                            {
                                var headEndEl = outline.GetFirstChild<Drawing.HeadEnd>();
                                if (headEndEl != null)
                                    outline.InsertBefore(solidFill, headEndEl);
                                else
                                {
                                    var tailEndEl = outline.GetFirstChild<Drawing.TailEnd>();
                                    if (tailEndEl != null)
                                        outline.InsertBefore(solidFill, tailEndEl);
                                    else
                                        outline.AppendChild(solidFill);
                                }
                            }
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
                    case "preset" or "prstgeom" or "shape":
                    {
                        // CONSISTENCY(canonical-key): schema canonical is 'shape';
                        // 'preset'/'prstgeom' retained as legacy aliases.
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
                        var newHeadEnd = new Drawing.HeadEnd { Type = ParseLineEndType(value) };
                        // CT_LineProperties: ... → headEnd → tailEnd (headEnd before tailEnd)
                        var existingTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                        if (existingTailEnd != null)
                            outline.InsertBefore(newHeadEnd, existingTailEnd);
                        else
                            outline.AppendChild(newHeadEnd);
                        break;
                    }
                    case "tailend" or "tailEnd":
                    {
                        var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                        var outline = spPr.GetFirstChild<Drawing.Outline>()
                            ?? spPr.AppendChild(new Drawing.Outline());
                        outline.RemoveAllChildren<Drawing.TailEnd>();
                        // CT_LineProperties: tailEnd is last — always append
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
                        var off = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                        var ext = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                        var keyLower = key.ToLowerInvariant();
                        // CONSISTENCY(group-scale-baseline): group scaling needs <a:chOff>/<a:chExt>
                        // as a child-coordinate baseline. Before we mutate ext/off, snapshot the
                        // current ext/off into chExt/chOff if they aren't already present — that
                        // way the first Set of width/height captures the "before" as the logical
                        // child coordinate space, so shrinking ext shrinks the rendered children.
                        if (keyLower is "x" or "y")
                        {
                            if (xfrm.ChildOffset == null)
                                xfrm.ChildOffset = new Drawing.ChildOffset { X = off.X ?? 0, Y = off.Y ?? 0 };
                        }
                        else // width or height
                        {
                            if (xfrm.ChildExtents == null)
                                xfrm.ChildExtents = new Drawing.ChildExtents { Cx = ext.Cx ?? 0, Cy = ext.Cy ?? 0 };
                        }
                        TryApplyPositionSize(keyLower, value, off, ext);
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

    // ==================== Per-element-type Set helpers ====================
    // Mechanical extractions from the original god-method Set(); each helper
    // owns one path-pattern's full handling. Splitting was for AI-readability
    // (each helper now <100 lines, fits in one Read) — no behavior change.

    private List<string> SetNotesByPath(Match notesSetMatch, Dictionary<string, string> properties)
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

    private List<string> SetShapeRunByPath(Match runMatch, Dictionary<string, string> properties)
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
        var tooltipValRun = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart);
        if (linkValRun != null) ApplyRunHyperlink(slidePart, targetRun, linkValRun, tooltipValRun);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetParagraphRunByPath(Match paraRunMatch, Dictionary<string, string> properties)
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
        var tooltipVal = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart);
        if (linkVal != null) ApplyRunHyperlink(slidePart, targetRun, linkVal, tooltipVal);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetChartAxisByPath(Match chartAxisSetMatch, Dictionary<string, string> properties)
    {
        var caSlideIdx = int.Parse(chartAxisSetMatch.Groups[1].Value);
        var caChartIdx = int.Parse(chartAxisSetMatch.Groups[2].Value);
        var caRole = chartAxisSetMatch.Groups[3].Value;
        var (caSlidePart, _, caChartPart, _) = ResolveChart(caSlideIdx, caChartIdx);
        if (caChartPart == null)
            throw new ArgumentException($"Axis Set not supported on extended charts.");
        var axUnsupported = ChartHelper.SetAxisProperties(caChartPart, caRole, properties);
        GetSlide(caSlidePart).Save();
        return axUnsupported;
    }

    private List<string> SetChartByPath(Match chartSetMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(chartSetMatch.Groups[1].Value);
        var chartIdx = int.Parse(chartSetMatch.Groups[2].Value);
        var seriesIdx = chartSetMatch.Groups[3].Success ? int.Parse(chartSetMatch.Groups[3].Value) : 0;

        var (slidePart, chartGf, chartPart, extChartPart) = ResolveChart(slideIdx, chartIdx);

        // If series sub-path, prefix all properties with series{N}. for ChartSetter
        var chartProps = new Dictionary<string, string>();
        var gfProps = new Dictionary<string, string>();
        if (seriesIdx > 0)
        {
            foreach (var (key, value) in properties)
                chartProps[$"series{seriesIdx}.{key}"] = value;
        }
        else
        {
            foreach (var (key, value) in properties)
            {
                if (key.ToLowerInvariant() is "x" or "y" or "width" or "height" or "name")
                    gfProps[key] = value;
                else
                    chartProps[key] = value;
            }
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

        List<string> unsupported;
        if (chartPart != null)
        {
            unsupported = ChartHelper.SetChartProperties(chartPart, chartProps);
        }
        else if (extChartPart != null)
        {
            // cx:chart — delegates to ChartExBuilder.SetChartProperties.
            // Same shared implementation as Excel/Word.
            unsupported = ChartExBuilder.SetChartProperties(extChartPart, chartProps);
        }
        else
        {
            unsupported = chartProps.Keys.ToList();
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetMasterShapeByPath(Match masterShapeMatch, Dictionary<string, string> properties)
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
            var unsupp = SetRunOrShapeProperties(properties, allRuns, shape);
            rootEl.Save();
            return unsupp;
        }

        throw new ArgumentException($"Unsupported element type: '{elType}' for master/layout. Valid types: shape.");
    }

    private List<string> SetParagraphByPath(Match paraMatch, Dictionary<string, string> properties)
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
                case "level":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var lvl) || lvl < 0 || lvl > 8)
                        throw new ArgumentException($"Invalid 'level' value: '{value}'. Expected an integer between 0 and 8 (OOXML a:pPr/@lvl).");
                    pProps.Level = lvl;
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
                {
                    var paraTooltip = properties.GetValueOrDefault("tooltip");
                    foreach (var r in paraRuns)
                        ApplyRunHyperlink(slidePart, r, value, paraTooltip);
                    break;
                }
                case "tooltip":
                    // handled in tandem with "link"; standalone tooltip change is not supported here
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

    private List<string> SetTableCellByPath(Match tblCellMatch, Dictionary<string, string> properties)
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

    private List<string> SetTableRowByPath(Match tblRowMatch, Dictionary<string, string> properties)
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
                case "text":
                {
                    // Two behaviors based on presence of tab:
                    //  - No tab: broadcast the same text to all cells in the row
                    //  - Tab-delimited: distribute tokens across cells by position
                    //    ("X1\tX2\tX3" → tc[1]="X1", tc[2]="X2", tc[3]="X3")
                    // Extra tokens beyond cell count are dropped; cells beyond token
                    // count are left unchanged.
                    var rowCells = row.Elements<Drawing.TableCell>().ToList();
                    if (value.Contains('\t'))
                    {
                        var tokens = value.Split('\t');
                        for (int i = 0; i < rowCells.Count && i < tokens.Length; i++)
                            ReplaceCellText(rowCells[i], tokens[i]);
                    }
                    else
                    {
                        foreach (var c in rowCells)
                            ReplaceCellText(c, value);
                    }
                    break;
                }
                default:
                    // c1, c2, ... shorthand: set text of specific cell by index
                    if (key.Length >= 2 && key[0] == 'c' && int.TryParse(key.AsSpan(1), out var cIdx))
                    {
                        var rowCells = row.Elements<Drawing.TableCell>().ToList();
                        if (cIdx < 1 || cIdx > rowCells.Count)
                            throw new ArgumentException($"Cell c{cIdx} out of range (row has {rowCells.Count} cells)");
                        ReplaceCellText(rowCells[cIdx - 1], value);
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

    private List<string> SetPlaceholderByPath(Match phMatch, Dictionary<string, string> properties)
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

    private List<string> SetZoomByPath(Match zoomSetMatch, Dictionary<string, string> properties)
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

    private List<string> SetModel3DByPath(Match model3dSetMatch, Dictionary<string, string> properties)
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

    private List<string> SetSlideByPath(Match slideOnlyMatch, Dictionary<string, string> properties)
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
                    ApplyBackground(slidePart2, value, ReadBackgroundImageOptions(properties));
                    break;
                case "background.mode":
                case "background.alpha":
                case "background.scale":
                    // If paired with "background=", consumed inside the "background" case
                    // via ReadBackgroundImageOptions. Otherwise mutate the existing image
                    // fill in place — done once for the whole property batch, gated below.
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
                case "showfooter":
                case "showslidenumber":
                case "showdate":
                case "showheader":
                {
                    // Toggle header/footer visibility flags on the slide.
                    // Emits <p:hf ftr="1" sldNum="0" dt="1" hdr="0"/> as a
                    // direct child of <p:sld>. The OpenXml SDK models this
                    // via DocumentFormat.OpenXml.Presentation.HeaderFooter
                    // (local name "hf"). Although CT_Slide's published
                    // schema does not list hf, PowerPoint itself writes it
                    // on slides when the "Insert > Header & Footer" dialog
                    // toggles per-slide overrides — we mirror that.
                    var hf = slide2.GetFirstChild<HeaderFooter>() ?? new HeaderFooter();
                    bool isNew = hf.Parent == null;
                    bool flag = IsTruthy(value);
                    switch (key.ToLowerInvariant())
                    {
                        case "showfooter": hf.Footer = flag; break;
                        case "showslidenumber": hf.SlideNumber = flag; break;
                        case "showdate": hf.DateTime = flag; break;
                        case "showheader": hf.Header = flag; break;
                    }
                    if (isNew) slide2.AppendChild(hf);
                    break;
                }
                case "layout":
                {
                    // Change slide layout
                    var presentationPart = _doc.PresentationPart
                        ?? throw new InvalidOperationException("No presentation part");
                    var allLayouts = presentationPart.SlideMasterParts
                        .SelectMany(m => m.SlideLayoutParts).ToList();
                    var targetLayout = allLayouts.FirstOrDefault(lp =>
                        lp.SlideLayout?.CommonSlideData?.Name?.Value?.Equals(value, StringComparison.OrdinalIgnoreCase) == true);
                    if (targetLayout == null)
                    {
                        var availableNames = allLayouts
                            .Select(lp => lp.SlideLayout?.CommonSlideData?.Name?.Value)
                            .Where(n => n != null)
                            .ToList();
                        throw new ArgumentException($"Layout '{value}' not found. Available layouts: {string.Join(", ", availableNames)}");
                    }
                    // Point the slide's layout relationship to the new layout
                    if (slidePart2.SlideLayoutPart != null)
                        slidePart2.DeletePart(slidePart2.SlideLayoutPart);
                    slidePart2.AddPart(targetLayout);
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(slide2, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid slide props: background, background.mode, background.alpha, background.scale, layout, transition, name, align, distribute, targets, showFooter, showSlideNumber, showDate, showHeader)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        MaybeMutateExistingBackgroundImage(slidePart2, properties);
        slide2.Save();
        return unsupported;
    }

    private List<string> SetOleByPath(Match oleSetMatch, Dictionary<string, string> properties)
    {
        var oleSlideIdx = int.Parse(oleSetMatch.Groups[1].Value);
        var oleEntryIdx = int.Parse(oleSetMatch.Groups[2].Value);
        var oleSlideParts = GetSlideParts().ToList();
        if (oleSlideIdx < 1 || oleSlideIdx > oleSlideParts.Count)
            throw new ArgumentException($"Slide {oleSlideIdx} not found (total: {oleSlideParts.Count})");
        var oleSlidePart = oleSlideParts[oleSlideIdx - 1];
        var oleShapeTree = GetSlide(oleSlidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var oleFrames = oleShapeTree.Elements<GraphicFrame>()
            .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any())
            .ToList();
        if (oleEntryIdx < 1 || oleEntryIdx > oleFrames.Count)
            throw new ArgumentException($"OLE object {oleEntryIdx} not found (total: {oleFrames.Count})");
        var oleFrame = oleFrames[oleEntryIdx - 1];
        var oleEl = oleFrame.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().First();
        var oleUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "path" or "src":
                {
                    // Delete old payload part and attach the new one.
                    if (oleEl.Id?.Value is string oldRel && !string.IsNullOrEmpty(oldRel))
                    {
                        try { oleSlidePart.DeletePart(oldRel); } catch { }
                    }
                    var (newRel, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(oleSlidePart, value, _filePath);
                    oleEl.Id = newRel;
                    // Auto-refresh progId from the new extension unless
                    // the caller explicitly pinned one in the same call.
                    if (!properties.ContainsKey("progId") && !properties.ContainsKey("progid"))
                    {
                        var autoProgId = OfficeCli.Core.OleHelper.DetectProgId(value);
                        OfficeCli.Core.OleHelper.ValidateProgId(autoProgId);
                        oleEl.ProgId = autoProgId;
                    }
                    break;
                }
                case "progid":
                    OfficeCli.Core.OleHelper.ValidateProgId(value);
                    oleEl.ProgId = value;
                    break;
                case "name":
                    oleEl.Name = value;
                    break;
                case "display":
                {
                    // Strict: only "icon" or "content" are accepted —
                    // see OleHelper.NormalizeOleDisplay.
                    var oleDisp = OfficeCli.Core.OleHelper.NormalizeOleDisplay(value);
                    oleEl.ShowAsIcon = oleDisp != "content";
                    break;
                }
                case "x" or "y" or "width" or "height":
                {
                    var xfrm = oleFrame.Transform ?? (oleFrame.Transform = new Transform());
                    var off = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset { X = 0, Y = 0 });
                    var ext = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents { Cx = 0, Cy = 0 });
                    var emu = ParseEmu(value);
                    var k = key.ToLowerInvariant();
                    // CONSISTENCY(ole-nonnegative-size): width/height are
                    // OOXML positive-sized types (ST_PositiveCoordinate).
                    // Silently storing a negative EMU breaks the shape
                    // frame and opens unpredictably in PowerPoint. Reject
                    // it explicitly; x/y may legitimately be negative
                    // (off-slide anchors) so they pass through.
                    if ((k == "width" || k == "height") && emu < 0)
                        throw new ArgumentException($"{k} must be non-negative");
                    switch (k)
                    {
                        case "x": off.X = emu; break;
                        case "y": off.Y = emu; break;
                        case "width": ext.Cx = emu; break;
                        case "height": ext.Cy = emu; break;
                    }
                    break;
                }
                default:
                    oleUnsupported.Add(key);
                    break;
            }
        }
        GetSlide(oleSlidePart).Save();
        return oleUnsupported;
    }

    private List<string> SetMediaByPath(Match mediaSetMatch, Dictionary<string, string> properties)
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

    private List<string> SetShapeByPath(Match match, Dictionary<string, string> properties)
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
            var tooltipValue = properties.GetValueOrDefault("tooltip");
            var excludeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "animation", "animate", "motionpath", "motionPath", "link", "tooltip", "zorder", "z-order", "order" };
            var shapeProps = properties
                .Where(kv => !excludeKeys.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value);

            var unsupported = SetRunOrShapeProperties(shapeProps, allRuns, shape, slidePart);

            if (animValue != null)
            {
                // Remove existing animations before applying new one (replace, not accumulate)
                var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (shapeId.HasValue)
                    RemoveShapeAnimations(slidePart.Slide!, shapeId.Value);
                ApplyShapeAnimation(slidePart, shape, animValue);
            }
            if (motionPathValue != null)
                ApplyMotionPathAnimation(slidePart, shape, motionPathValue);
            if (linkValue != null)
                ApplyShapeHyperlink(slidePart, shape, linkValue, tooltipValue);

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
}
