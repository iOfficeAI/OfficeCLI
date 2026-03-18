// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // Document-level properties
        if (path == "/" || path == "")
        {
            SetDocumentProperties(properties);
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Handle /watermark path
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
        {
            // Find watermark VML shape in headers and modify properties
            foreach (var hp in _doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
            {
                if (hp.Header == null) continue;
                var picts = hp.Header.Descendants<Picture>().ToList();
                foreach (var pict in picts)
                {
                    if (!pict.InnerXml.Contains("WaterMark", StringComparison.OrdinalIgnoreCase)) continue;

                    // Rebuild VML with updated properties — parse existing values as defaults
                    var xml = pict.InnerXml;
                    foreach (var (key, value) in properties)
                    {
                        switch (key.ToLowerInvariant())
                        {
                            case "text":
                                xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                    @"string=""[^""]*""", $@"string=""{System.Security.SecurityElement.Escape(value)}""");
                                break;
                            case "color":
                                var clr = value.TrimStart('#').ToUpperInvariant();
                                xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                    @"fillcolor=""[^""]*""", $@"fillcolor=""{clr}""");
                                break;
                            case "font":
                                xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                    @"font-family:&quot;[^&]*&quot;", $@"font-family:&quot;{System.Security.SecurityElement.Escape(value)}&quot;");
                                break;
                            case "opacity":
                                xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                    @"opacity=""[^""]*""", $@"opacity=""{value}""");
                                break;
                            case "rotation":
                                xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                    @"rotation:\d+", $@"rotation:{value}");
                                break;
                            default:
                                unsupported.Add(key);
                                break;
                        }
                    }
                    pict.InnerXml = xml;
                }
                hp.Header.Save();
            }
            return unsupported;
        }

        // Handle header/footer paths
        var hfParts = ParsePath(path);
        if (hfParts.Count >= 1)
        {
            var firstName = hfParts[0].Name.ToLowerInvariant();
            if ((firstName == "header" || firstName == "footer") && hfParts.Count == 1)
            {
                SetHeaderFooter(firstName, (hfParts[0].Index ?? 1) - 1, properties, unsupported);
                return unsupported;
            }
        }

        // Chart paths: /chart[N]
        var chartMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartIdx = int.Parse(chartMatch.Groups[1].Value);
            var chartParts = _doc.MainDocumentPart?.ChartParts.ToList()
                ?? throw new ArgumentException("No charts in this document");
            if (chartIdx < 1 || chartIdx > chartParts.Count)
                throw new ArgumentException($"Chart {chartIdx} not found (total: {chartParts.Count})");
            unsupported = Core.ChartHelper.SetChartProperties(chartParts[chartIdx - 1], properties);
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // TOC paths: /toc[N]
        var tocMatch = System.Text.RegularExpressions.Regex.Match(path, @"/toc\[(\d+)\]$");
        if (tocMatch.Success)
        {
            var tocIdx = int.Parse(tocMatch.Groups[1].Value);
            var tocParas = FindTocParagraphs();
            if (tocIdx < 1 || tocIdx > tocParas.Count)
                throw new ArgumentException($"TOC {tocIdx} not found (total: {tocParas.Count})");

            var tocPara = tocParas[tocIdx - 1];

            // Rebuild the field code from properties
            var instrRun = tocPara.Descendants<Run>()
                .FirstOrDefault(r => r.GetFirstChild<FieldCode>() != null);
            if (instrRun == null)
                throw new InvalidOperationException("TOC field code not found");

            var fieldCode = instrRun.GetFirstChild<FieldCode>()!;
            var instr = fieldCode.Text ?? "";

            // Update levels
            if (properties.TryGetValue("levels", out var newLevels))
            {
                var levelsRx = System.Text.RegularExpressions.Regex.Match(instr, @"\\o\s+""[^""]+""");
                instr = levelsRx.Success
                    ? instr.Replace(levelsRx.Value, $"\\o \"{newLevels}\"")
                    : instr.TrimEnd() + $" \\o \"{newLevels}\" ";
            }

            // Update hyperlinks switch
            if (properties.TryGetValue("hyperlinks", out var hlSwitch))
            {
                if (IsTruthy(hlSwitch) && !instr.Contains("\\h"))
                    instr = instr.TrimEnd() + " \\h ";
                else if (!IsTruthy(hlSwitch))
                    instr = instr.Replace("\\h", "").Replace("  ", " ");
            }

            // Update page numbers switch (\\z = hide page numbers)
            if (properties.TryGetValue("pagenumbers", out var pnSwitch))
            {
                if (!IsTruthy(pnSwitch) && !instr.Contains("\\z"))
                    instr = instr.TrimEnd() + " \\z ";
                else if (IsTruthy(pnSwitch))
                    instr = instr.Replace("\\z", "").Replace("  ", " ");
            }

            fieldCode.Text = instr;

            // Mark field as dirty so Word updates it on open
            var beginRun = tocPara.Descendants<Run>()
                .FirstOrDefault(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin);
            if (beginRun != null)
            {
                var fldChar = beginRun.GetFirstChild<FieldChar>()!;
                fldChar.Dirty = true;
            }

            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Footnote paths: /footnote[N] or .../footnote[N]
        var fnSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"/footnote\[(\d+)\]$");
        if (fnSetMatch.Success)
        {
            var fnId = int.Parse(fnSetMatch.Groups[1].Value);
            var fn = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
            if (fn == null)
            {
                // Try ordinal lookup (1-based index among user footnotes)
                var userFns = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                    .Elements<Footnote>().Where(f => f.Id?.Value > 0).ToList();
                if (userFns != null && fnId >= 1 && fnId <= userFns.Count)
                    fn = userFns[fnId - 1];
                else
                    throw new ArgumentException($"Footnote {fnId} not found");
            }

            if (properties.TryGetValue("text", out var fnText))
            {
                // Find the content paragraph (skip the reference mark run)
                var contentRuns = fn.Descendants<Run>()
                    .Where(r => r.GetFirstChild<FootnoteReferenceMark>() == null).ToList();
                if (contentRuns.Count > 0)
                {
                    // Update first content run; keep space as separate element
                    var textEl = contentRuns[0].GetFirstChild<Text>();
                    if (textEl != null)
                    {
                        textEl.Text = fnText;
                        textEl.Space = SpaceProcessingModeValues.Preserve;
                    }
                    else
                        contentRuns[0].AppendChild(new Text(fnText) { Space = SpaceProcessingModeValues.Preserve });
                    // Remove extra runs so text is not duplicated
                    for (int i = 1; i < contentRuns.Count; i++)
                        contentRuns[i].Remove();
                }
            }
            _doc.MainDocumentPart?.FootnotesPart?.Footnotes?.Save();
            return unsupported;
        }

        // Endnote paths: /endnote[N] or .../endnote[N]
        var enSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"/endnote\[(\d+)\]$");
        if (enSetMatch.Success)
        {
            var enId = int.Parse(enSetMatch.Groups[1].Value);
            var en = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
            if (en == null)
            {
                // Try ordinal lookup (1-based index among user endnotes)
                var userEns = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                    .Elements<Endnote>().Where(e => e.Id?.Value > 0).ToList();
                if (userEns != null && enId >= 1 && enId <= userEns.Count)
                    en = userEns[enId - 1];
                else
                    throw new ArgumentException($"Endnote {enId} not found");
            }

            if (properties.TryGetValue("text", out var enText))
            {
                var contentRuns = en.Descendants<Run>()
                    .Where(r => r.GetFirstChild<EndnoteReferenceMark>() == null).ToList();
                if (contentRuns.Count > 0)
                {
                    var textEl = contentRuns[0].GetFirstChild<Text>();
                    if (textEl != null)
                    {
                        textEl.Text = enText;
                        textEl.Space = SpaceProcessingModeValues.Preserve;
                    }
                    else
                        contentRuns[0].AppendChild(new Text(enText) { Space = SpaceProcessingModeValues.Preserve });
                    // Remove extra runs so text is not duplicated
                    for (int i = 1; i < contentRuns.Count; i++)
                        contentRuns[i].Remove();
                }
            }
            _doc.MainDocumentPart?.EndnotesPart?.Endnotes?.Save();
            return unsupported;
        }

        // Section paths: /section[N]
        var secSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/section\[(\d+)\]$");
        if (secSetMatch.Success)
        {
            var secIdx = int.Parse(secSetMatch.Groups[1].Value);
            var sectionProps = FindSectionProperties();

            // If no section properties exist and requesting section 1, create one
            if (sectionProps.Count == 0 && secIdx == 1)
            {
                var sBody = _doc.MainDocumentPart?.Document?.Body;
                if (sBody != null)
                {
                    var newSectPr = new SectionProperties();
                    sBody.AppendChild(newSectPr);
                    sectionProps = FindSectionProperties();
                }
            }

            if (secIdx < 1 || secIdx > sectionProps.Count)
                throw new ArgumentException($"Section {secIdx} not found (total: {sectionProps.Count})");

            var sectPr = sectionProps[secIdx - 1];
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "type":
                        var st = sectPr.GetFirstChild<SectionType>() ?? sectPr.PrependChild(new SectionType());
                        st.Val = value.ToLowerInvariant() switch
                        {
                            "continuous" => SectionMarkValues.Continuous,
                            "evenpage" or "even" => SectionMarkValues.EvenPage,
                            "oddpage" or "odd" => SectionMarkValues.OddPage,
                            _ => SectionMarkValues.NextPage
                        };
                        break;
                    case "pagewidth":
                        EnsureSectPrPageSize(sectPr).Width = uint.TryParse(value, out var pgW) ? pgW : throw new FormatException($"Invalid page width: {value}");
                        break;
                    case "pageheight":
                        EnsureSectPrPageSize(sectPr).Height = uint.TryParse(value, out var pgH) ? pgH : throw new FormatException($"Invalid page height: {value}");
                        break;
                    case "orientation":
                        var ps = EnsureSectPrPageSize(sectPr);
                        ps.Orient = value.ToLowerInvariant() == "landscape"
                            ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
                        break;
                    case "margintop":
                        EnsureSectPrPageMargin(sectPr).Top = int.Parse(value);
                        break;
                    case "marginbottom":
                        EnsureSectPrPageMargin(sectPr).Bottom = int.Parse(value);
                        break;
                    case "marginleft":
                        EnsureSectPrPageMargin(sectPr).Left = uint.Parse(value);
                        break;
                    case "marginright":
                        EnsureSectPrPageMargin(sectPr).Right = uint.Parse(value);
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Style paths: /styles/StyleId
        var styleSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/styles/(.+)$");
        if (styleSetMatch.Success)
        {
            var styleId = styleSetMatch.Groups[1].Value;
            var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?? throw new InvalidOperationException("No styles part");
            var style = styles.Elements<Style>().FirstOrDefault(s =>
                s.StyleId?.Value == styleId || s.StyleName?.Val?.Value == styleId)
                ?? throw new ArgumentException($"Style '{styleId}' not found");

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "name":
                        var sn = style.StyleName ?? style.AppendChild(new StyleName());
                        sn.Val = value;
                        break;
                    case "basedon":
                        var bo = style.BasedOn ?? style.AppendChild(new BasedOn());
                        bo.Val = value;
                        break;
                    case "next":
                        var ns = style.NextParagraphStyle ?? style.AppendChild(new NextParagraphStyle());
                        ns.Val = value;
                        break;
                    case "font":
                        var rPr = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        rPr.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                        break;
                    case "size":
                        var rPr2 = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        rPr2.FontSize = new FontSize { Val = ((int)(ParseFontSize(value) * 2)).ToString() };
                        break;
                    case "bold":
                        var rPr3 = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        rPr3.Bold = IsTruthy(value) ? new Bold() : null;
                        break;
                    case "italic":
                        var rPr4 = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        rPr4.Italic = IsTruthy(value) ? new Italic() : null;
                        break;
                    case "color":
                        var rPr5 = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        rPr5.Color = new Color { Val = value.TrimStart('#').ToUpperInvariant() };
                        break;
                    case "alignment":
                        var pPr = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());
                        pPr.Justification = new Justification
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => JustificationValues.Center,
                                "right" => JustificationValues.Right,
                                "justify" => JustificationValues.Both,
                                _ => JustificationValues.Left
                            }
                        };
                        break;
                    case "spacebefore":
                        var pPr2 = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());
                        var sp2 = pPr2.SpacingBetweenLines ?? (pPr2.SpacingBetweenLines = new SpacingBetweenLines());
                        sp2.Before = value;
                        break;
                    case "spaceafter":
                        var pPr3 = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());
                        var sp3 = pPr3.SpacingBetweenLines ?? (pPr3.SpacingBetweenLines = new SpacingBetweenLines());
                        sp3.After = value;
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            styles.Save();
            return unsupported;
        }

        var parts = ParsePath(path);
        var element = NavigateToElement(parts);
        if (element == null)
            throw new ArgumentException($"Path not found: {path}");

        if (element is BookmarkStart bkStart)
        {
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "name":
                        // Check for duplicate bookmark names
                        var existingBk = _doc.MainDocumentPart?.Document?.Body?
                            .Descendants<BookmarkStart>()
                            .FirstOrDefault(b => b.Name?.Value == value && b != bkStart);
                        if (existingBk != null)
                            throw new ArgumentException($"Bookmark name '{value}' already exists");
                        bkStart.Name = value;
                        break;
                    case "text":
                        var bkId = bkStart.Id?.Value;
                        if (bkId != null)
                        {
                            var toRemove = new List<OpenXmlElement>();
                            var sib = bkStart.NextSibling();
                            while (sib != null)
                            {
                                if (sib is BookmarkEnd bkEnd && bkEnd.Id?.Value == bkId)
                                    break;
                                toRemove.Add(sib);
                                sib = sib.NextSibling();
                            }
                            foreach (var el in toRemove) el.Remove();
                            bkStart.InsertAfterSelf(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                        }
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }

            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        if (element is Run run)
        {
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        var textEl = run.GetFirstChild<Text>();
                        if (textEl != null) textEl.Text = value;
                        break;
                    case "bold":
                        EnsureRunProperties(run).Bold = IsTruthy(value) ? new Bold() : null;
                        break;
                    case "italic":
                        EnsureRunProperties(run).Italic = IsTruthy(value) ? new Italic() : null;
                        break;
                    case "caps":
                        EnsureRunProperties(run).Caps = IsTruthy(value) ? new Caps() : null;
                        break;
                    case "smallcaps":
                        EnsureRunProperties(run).SmallCaps = IsTruthy(value) ? new SmallCaps() : null;
                        break;
                    case "dstrike":
                        EnsureRunProperties(run).DoubleStrike = IsTruthy(value) ? new DoubleStrike() : null;
                        break;
                    case "vanish":
                        EnsureRunProperties(run).Vanish = IsTruthy(value) ? new Vanish() : null;
                        break;
                    case "outline":
                        EnsureRunProperties(run).Outline = IsTruthy(value) ? new Outline() : null;
                        break;
                    case "shadow":
                        EnsureRunProperties(run).Shadow = IsTruthy(value) ? new Shadow() : null;
                        break;
                    case "emboss":
                        EnsureRunProperties(run).Emboss = IsTruthy(value) ? new Emboss() : null;
                        break;
                    case "imprint":
                        EnsureRunProperties(run).Imprint = IsTruthy(value) ? new Imprint() : null;
                        break;
                    case "noproof":
                        EnsureRunProperties(run).NoProof = IsTruthy(value) ? new NoProof() : null;
                        break;
                    case "rtl":
                        EnsureRunProperties(run).RightToLeftText = IsTruthy(value) ? new RightToLeftText() : null;
                        break;
                    case "font":
                        var rPrFont = EnsureRunProperties(run);
                        var existingFonts = rPrFont.RunFonts;
                        if (existingFonts != null)
                        {
                            existingFonts.Ascii = value;
                            existingFonts.HighAnsi = value;
                            existingFonts.EastAsia = value;
                        }
                        else
                        {
                            rPrFont.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                        }
                        break;
                    case "size":
                        EnsureRunProperties(run).FontSize = new FontSize
                        {
                            Val = ((int)(ParseFontSize(value) * 2)).ToString() // half-points
                        };
                        break;
                    case "highlight":
                        EnsureRunProperties(run).Highlight = new Highlight
                        {
                            Val = new HighlightColorValues(value)
                        };
                        break;
                    case "color":
                        EnsureRunProperties(run).Color = new Color { Val = value.TrimStart('#').ToUpperInvariant() };
                        break;
                    case "underline":
                        EnsureRunProperties(run).Underline = new Underline
                        {
                            Val = new UnderlineValues(value)
                        };
                        break;
                    case "strike":
                        EnsureRunProperties(run).Strike = IsTruthy(value) ? new Strike() : null;
                        break;
                    case "superscript":
                        EnsureRunProperties(run).VerticalTextAlignment = IsTruthy(value)
                            ? new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }
                            : null;
                        break;
                    case "subscript":
                        EnsureRunProperties(run).VerticalTextAlignment = IsTruthy(value)
                            ? new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }
                            : null;
                        break;
                    case "shading":
                    case "shd":
                        // shd has w:val, w:fill, w:color — value format: "fill" or "val;fill" or "val;fill;color"
                        var shdParts = value.Split(';');
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = shdParts[0].TrimStart('#').ToUpperInvariant();
                        }
                        else if (shdParts.Length >= 2)
                        {
                            shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = shdParts[1].TrimStart('#').ToUpperInvariant();
                            if (shdParts.Length >= 3) shd.Color = shdParts[2].TrimStart('#').ToUpperInvariant();
                        }
                        EnsureRunProperties(run).Shading = shd;
                        break;
                    case "alt":
                        var drawingAlt = run.GetFirstChild<Drawing>();
                        if (drawingAlt != null)
                        {
                            var docPropsAlt = drawingAlt.Descendants<DW.DocProperties>().FirstOrDefault();
                            if (docPropsAlt != null) docPropsAlt.Description = value;
                        }
                        else unsupported.Add(key);
                        break;
                    case "width":
                        var drawingW = run.GetFirstChild<Drawing>();
                        if (drawingW != null)
                        {
                            var extentW = drawingW.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentW != null) extentW.Cx = ParseEmu(value);
                            var extentsW = drawingW.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsW != null) extentsW.Cx = ParseEmu(value);
                        }
                        else unsupported.Add(key);
                        break;
                    case "height":
                        var drawingH = run.GetFirstChild<Drawing>();
                        if (drawingH != null)
                        {
                            var extentH = drawingH.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentH != null) extentH.Cy = ParseEmu(value);
                            var extentsH = drawingH.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsH != null) extentsH.Cy = ParseEmu(value);
                        }
                        else unsupported.Add(key);
                        break;
                    case "link":
                    {
                        var mainPart3 = _doc.MainDocumentPart!;
                        if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            // Remove hyperlink wrapper if present
                            if (run.Parent is Hyperlink existingHlNone)
                            {
                                foreach (var childRun in existingHlNone.Elements<Run>().ToList())
                                    existingHlNone.InsertBeforeSelf(childRun);
                                existingHlNone.Remove();
                            }
                        }
                        else
                        {
                            // Accept both absolute and relative URIs (Open-XML-SDK supports both)
                            var uri = Uri.TryCreate(value, UriKind.Absolute, out var absUri)
                                ? absUri
                                : new Uri(value, UriKind.Relative);
                            var newRelId = mainPart3.AddHyperlinkRelationship(uri, isExternal: true).Id;
                            if (run.Parent is Hyperlink existingHl)
                            {
                                existingHl.Id = newRelId;
                            }
                            else
                            {
                                var newHl = new Hyperlink { Id = newRelId };
                                run.InsertBeforeSelf(newHl);
                                run.Remove();
                                newHl.AppendChild(run);
                            }
                        }
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(EnsureRunProperties(run), key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is Paragraph para)
        {
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            foreach (var (key, value) in properties)
            {
                var k = key.ToLowerInvariant();
                if (ApplyParagraphLevelProperty(pProps, key, value))
                {
                    // handled by paragraph-level helper
                }
                else switch (k)
                {
                    case "liststyle":
                        ApplyListStyle(para, value);
                        break;
                    case "start":
                        SetListStartValue(para, int.Parse(value));
                        break;
                    case "size" or "font" or "bold" or "italic" or "color" or "highlight" or "underline" or "strike":
                        // Apply run-level formatting to all runs in the paragraph
                        var allParaRuns = para.Descendants<Run>().ToList();
                        // Also update paragraph mark run properties (rPr inside pPr)
                        // so new runs inherit the formatting
                        var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                        ApplyRunFormatting(markRPr, key, value);
                        foreach (var pRun in allParaRuns)
                        {
                            var pRunProps = EnsureRunProperties(pRun);
                            ApplyRunFormatting(pRunProps, key, value);
                        }
                        break;
                    case "text":
                        // Set text on paragraph: update first run or create one
                        var existingRuns = para.Elements<Run>().ToList();
                        if (existingRuns.Count > 0)
                        {
                            var firstText = existingRuns[0].GetFirstChild<Text>();
                            if (firstText != null) firstText.Text = value;
                            else existingRuns[0].AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                            // Remove extra runs
                            for (int i = 1; i < existingRuns.Count; i++) existingRuns[i].Remove();
                        }
                        else
                        {
                            // Use paragraph mark run properties as default for new run
                            var newRun = new Run();
                            var markProps = pProps.ParagraphMarkRunProperties;
                            if (markProps != null)
                            {
                                var cloned = new RunProperties();
                                foreach (var child in markProps.ChildElements)
                                    cloned.AppendChild(child.CloneNode(true));
                                newRun.PrependChild(cloned);
                            }
                            newRun.AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                            para.AppendChild(newRun);
                        }
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(pProps, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }

        else if (element is TableCell cell)
        {
            var tcPr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
            string? deferredText = null;
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        // Defer text handling until after formatting is applied
                        deferredText = value;
                        break;
                    case "font":
                    case "size":
                    case "bold":
                    case "italic":
                    case "color":
                    case "highlight":
                    case "underline":
                    case "strike":
                        // Apply to all runs in all paragraphs in the cell
                        bool hasRuns = false;
                        foreach (var cellPara in cell.Elements<Paragraph>())
                        {
                            foreach (var cellRun in cellPara.Elements<Run>())
                            {
                                hasRuns = true;
                                var rPr = EnsureRunProperties(cellRun);
                                switch (key.ToLowerInvariant())
                                {
                                    case "font":
                                        rPr.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                                        break;
                                    case "size":
                                        rPr.FontSize = new FontSize { Val = ((int)(ParseFontSize(value) * 2)).ToString() };
                                        break;
                                    case "bold":
                                        rPr.Bold = IsTruthy(value) ? new Bold() : null;
                                        break;
                                    case "italic":
                                        rPr.Italic = IsTruthy(value) ? new Italic() : null;
                                        break;
                                    case "color":
                                        rPr.Color = new Color { Val = value.TrimStart('#').ToUpperInvariant() };
                                        break;
                                    case "highlight":
                                        rPr.Highlight = new Highlight { Val = new HighlightColorValues(value) };
                                        break;
                                    case "underline":
                                        rPr.Underline = new Underline { Val = new UnderlineValues(value) };
                                        break;
                                    case "strike":
                                        rPr.Strike = IsTruthy(value) ? new Strike() : null;
                                        break;
                                }
                            }
                        }
                        // If no runs exist, store formatting in ParagraphMarkRunProperties on first paragraph
                        if (!hasRuns)
                        {
                            var fp = cell.Elements<Paragraph>().FirstOrDefault();
                            if (fp == null) { fp = new Paragraph(); cell.AppendChild(fp); }
                            var pPr = fp.ParagraphProperties ?? fp.PrependChild(new ParagraphProperties());
                            var pmrp = pPr.ParagraphMarkRunProperties ?? pPr.AppendChild(new ParagraphMarkRunProperties());
                            switch (key.ToLowerInvariant())
                            {
                                case "font":
                                    pmrp.RemoveAllChildren<RunFonts>();
                                    pmrp.AppendChild(new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value });
                                    break;
                                case "size":
                                    pmrp.RemoveAllChildren<FontSize>();
                                    pmrp.AppendChild(new FontSize { Val = ((int)(ParseFontSize(value) * 2)).ToString() });
                                    break;
                                case "bold":
                                    pmrp.RemoveAllChildren<Bold>();
                                    if (IsTruthy(value)) pmrp.AppendChild(new Bold());
                                    break;
                                case "italic":
                                    pmrp.RemoveAllChildren<Italic>();
                                    if (IsTruthy(value)) pmrp.AppendChild(new Italic());
                                    break;
                                case "color":
                                    pmrp.RemoveAllChildren<Color>();
                                    pmrp.AppendChild(new Color { Val = value.TrimStart('#').ToUpperInvariant() });
                                    break;
                                case "highlight":
                                    pmrp.RemoveAllChildren<Highlight>();
                                    pmrp.AppendChild(new Highlight { Val = new HighlightColorValues(value) });
                                    break;
                                case "underline":
                                    pmrp.RemoveAllChildren<Underline>();
                                    pmrp.AppendChild(new Underline { Val = new UnderlineValues(value) });
                                    break;
                                case "strike":
                                    pmrp.RemoveAllChildren<Strike>();
                                    if (IsTruthy(value)) pmrp.AppendChild(new Strike());
                                    break;
                            }
                        }
                        break;
                    case "shd" or "shading":
                        var shdParts = value.Split(';');
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = shdParts[0].TrimStart('#').ToUpperInvariant();
                        }
                        else if (shdParts.Length >= 2)
                        {
                            shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = shdParts[1].TrimStart('#').ToUpperInvariant();
                            if (shdParts.Length >= 3) shd.Color = shdParts[2].TrimStart('#').ToUpperInvariant();
                        }
                        tcPr.Shading = shd;
                        break;
                    case "alignment":
                        var alignVal = value.ToLowerInvariant() switch
                        {
                            "center" => JustificationValues.Center,
                            "right" => JustificationValues.Right,
                            "justify" => JustificationValues.Both,
                            _ => JustificationValues.Left
                        };
                        // Apply alignment to ALL paragraphs in the cell, not just the first
                        foreach (var cellAlignPara in cell.Elements<Paragraph>())
                        {
                            var cpProps = cellAlignPara.ParagraphProperties ?? cellAlignPara.PrependChild(new ParagraphProperties());
                            cpProps.Justification = new Justification
                            {
                                Val = alignVal
                            };
                        }
                        break;
                    case "valign":
                        tcPr.TableCellVerticalAlignment = new TableCellVerticalAlignment
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => TableVerticalAlignmentValues.Center,
                                "bottom" => TableVerticalAlignmentValues.Bottom,
                                _ => TableVerticalAlignmentValues.Top
                            }
                        };
                        break;
                    case "width":
                        tcPr.TableCellWidth = new TableCellWidth { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    case "padding":
                    {
                        var dxa = value;
                        var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                        mar.TopMargin = new TopMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                        mar.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                        mar.LeftMargin = new LeftMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                        mar.RightMargin = new RightMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                        break;
                    }
                    case "padding.top":
                    {
                        var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                        mar.TopMargin = new TopMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    }
                    case "padding.bottom":
                    {
                        var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                        mar.BottomMargin = new BottomMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    }
                    case "padding.left":
                    {
                        var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                        mar.LeftMargin = new LeftMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    }
                    case "padding.right":
                    {
                        var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                        mar.RightMargin = new RightMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    }
                    case "textdirection" or "textdir":
                        tcPr.TextDirection = new TextDirection
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "btlr" or "vertical" => TextDirectionValues.BottomToTopLeftToRight,
                                "tbrl" or "vertical-rl" => TextDirectionValues.TopToBottomRightToLeft,
                                "lrtb" or "horizontal" => TextDirectionValues.LefToRightTopToBottom,
                                "tbrl-r" or "tb-rl-rotated" => TextDirectionValues.TopToBottomRightToLeftRotated,
                                "lrtb-r" or "lr-tb-rotated" => TextDirectionValues.LefttoRightTopToBottomRotated,
                                "tblr-r" or "tb-lr-rotated" => TextDirectionValues.TopToBottomLeftToRightRotated,
                                _ => TextDirectionValues.LefToRightTopToBottom
                            }
                        };
                        break;
                    case "nowrap":
                        tcPr.NoWrap = IsTruthy(value) ? new NoWrap() : null;
                        break;
                    case "vmerge":
                        tcPr.VerticalMerge = new VerticalMerge
                        {
                            Val = value.ToLowerInvariant() == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue
                        };
                        break;
                    case var k when k.StartsWith("border"):
                        ApplyCellBorders(tcPr, key, value);
                        break;
                    case "gridspan":
                        var newSpan = int.Parse(value);
                        tcPr.GridSpan = new GridSpan { Val = newSpan };
                        // Ensure the row has the correct number of tc elements.
                        // Calculate total grid columns occupied by all cells in this row,
                        // then remove/add cells so it matches the table grid.
                        if (element.Parent is TableRow parentRow)
                        {
                            var table = parentRow.Parent as Table;
                            var gridColList = table?.GetFirstChild<TableGrid>()
                                ?.Elements<GridColumn>().ToList();
                            var gridCols = gridColList?.Count ?? 0;
                            if (gridCols > 0)
                            {
                                // Calculate the grid column index where this cell starts
                                int startCol = 0;
                                foreach (var prevTc in parentRow.Elements<TableCell>())
                                {
                                    if (prevTc == element) break;
                                    startCol += prevTc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                                }

                                // Update cell width to sum of spanned grid columns
                                int spanWidth = 0;
                                for (int gi = startCol; gi < startCol + newSpan && gi < gridCols; gi++)
                                {
                                    if (int.TryParse(gridColList![gi].Width?.Value, out var gw))
                                        spanWidth += gw;
                                }
                                if (spanWidth > 0)
                                    tcPr.TableCellWidth = new TableCellWidth { Width = spanWidth.ToString(), Type = TableWidthUnitValues.Dxa };

                                // Calculate total columns occupied by current cells
                                var totalSpan = parentRow.Elements<TableCell>().Sum(tc =>
                                    tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                                // Remove excess cells after the current cell
                                while (totalSpan > gridCols)
                                {
                                    var nextCell = ((TableCell)element).NextSibling<TableCell>();
                                    if (nextCell == null) break;
                                    totalSpan -= nextCell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                                    nextCell.Remove();
                                }
                            }
                        }
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tcPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
            // Process deferred "text" AFTER formatting so font/size/bold are applied to existing runs first
            if (deferredText != null)
            {
                var firstPara = cell.Elements<Paragraph>().FirstOrDefault();
                if (firstPara == null)
                {
                    firstPara = new Paragraph();
                    cell.AppendChild(firstPara);
                }
                // Preserve RunProperties from first run before replacing
                var cellExistingRuns = firstPara.Elements<Run>().ToList();
                var cellRunProps = cellExistingRuns.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
                // Also check ParagraphMarkRunProperties if no run props found
                if (cellRunProps == null)
                {
                    var pmrp = firstPara.ParagraphProperties?.ParagraphMarkRunProperties;
                    if (pmrp != null) cellRunProps = new RunProperties(pmrp.CloneNode(true).ChildElements.Select(c => c.CloneNode(true)));
                }
                foreach (var r in cellExistingRuns) r.Remove();
                var cellNewRun = new Run(new Text(deferredText) { Space = SpaceProcessingModeValues.Preserve });
                if (cellRunProps != null) cellNewRun.PrependChild(cellRunProps);
                firstPara.AppendChild(cellNewRun);
            }
        }
        else if (element is TableRow row)
        {
            var trPr = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "height":
                        trPr.GetFirstChild<TableRowHeight>()?.Remove();
                        trPr.AppendChild(new TableRowHeight { Val = uint.Parse(value), HeightType = HeightRuleValues.AtLeast });
                        break;
                    case "height.exact":
                        trPr.GetFirstChild<TableRowHeight>()?.Remove();
                        trPr.AppendChild(new TableRowHeight { Val = uint.Parse(value), HeightType = HeightRuleValues.Exact });
                        break;
                    case "header":
                        if (IsTruthy(value))
                            trPr.AppendChild(new TableHeader());
                        else
                            trPr.GetFirstChild<TableHeader>()?.Remove();
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(trPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is Table tbl)
        {
            var tblPr = tbl.GetFirstChild<TableProperties>() ?? tbl.PrependChild(new TableProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "style":
                        var tblStyle = tblPr.TableStyle ?? (tblPr.TableStyle = new TableStyle());
                        tblStyle.Val = value;
                        break;
                    case "alignment":
                        tblPr.TableJustification = new TableJustification
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => TableRowAlignmentValues.Center,
                                "right" => TableRowAlignmentValues.Right,
                                _ => TableRowAlignmentValues.Left
                            }
                        };
                        break;
                    case "width":
                        if (value.EndsWith('%'))
                        {
                            var pct = int.Parse(value.TrimEnd('%')) * 50; // OOXML pct = percent * 50
                            tblPr.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                        }
                        else
                        {
                            tblPr.TableWidth = new TableWidth { Width = value, Type = TableWidthUnitValues.Dxa };
                        }
                        break;
                    case "indent":
                        tblPr.TableIndentation = new TableIndentation { Width = int.Parse(value), Type = TableWidthUnitValues.Dxa };
                        break;
                    case "cellspacing":
                        tblPr.TableCellSpacing = new TableCellSpacing { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    case "layout":
                        tblPr.TableLayout = new TableLayout
                        {
                            Type = value.ToLowerInvariant() == "fixed" ? TableLayoutValues.Fixed : TableLayoutValues.Autofit
                        };
                        break;
                    case "padding":
                    {
                        var dxa = value;
                        var cm = tblPr.TableCellMarginDefault ?? tblPr.AppendChild(new TableCellMarginDefault());
                        cm.TopMargin = new TopMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                        cm.TableCellLeftMargin = new TableCellLeftMargin { Width = short.Parse(dxa), Type = TableWidthValues.Dxa };
                        cm.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                        cm.TableCellRightMargin = new TableCellRightMargin { Width = short.Parse(dxa), Type = TableWidthValues.Dxa };
                        break;
                    }
                    case var k when k.StartsWith("border"):
                        ApplyTableBorders(tblPr, key, value);
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tblPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }

        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private void SetHeaderFooter(string kind, int index, Dictionary<string, string> properties, List<string> unsupported)
    {
        var mainPart = _doc.MainDocumentPart!;
        OpenXmlCompositeElement? container;

        if (kind == "header")
        {
            var part = mainPart.HeaderParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Header not found: /header[{index + 1}]");
            container = part.Header;
        }
        else
        {
            var part = mainPart.FooterParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Footer not found: /footer[{index + 1}]");
            container = part.Footer;
        }

        if (container == null)
            throw new ArgumentException($"{kind} content not found at index {index + 1}");

        var firstPara = container.Elements<Paragraph>().FirstOrDefault();
        if (firstPara == null)
        {
            firstPara = new Paragraph();
            container.AppendChild(firstPara);
        }
        var pProps = firstPara.ParagraphProperties ?? firstPara.PrependChild(new ParagraphProperties());

        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            if (ApplyParagraphLevelProperty(pProps, key, value))
            {
                // handled by paragraph-level helper
            }
            else switch (k)
            {
                case "text":
                {
                    RunProperties? existingRProps = null;
                    var existingRun = firstPara.Elements<Run>().FirstOrDefault();
                    if (existingRun?.RunProperties != null)
                        existingRProps = (RunProperties)existingRun.RunProperties.CloneNode(true);
                    foreach (var r in firstPara.Elements<Run>().ToList()) r.Remove();
                    var newRun = new Run();
                    if (existingRProps != null)
                        newRun.AppendChild(existingRProps);
                    newRun.AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                    firstPara.AppendChild(newRun);
                    break;
                }
                case "size" or "font" or "bold" or "italic" or "color" or "highlight" or "underline" or "strike":
                    // Apply run-level formatting to all runs in the container
                    foreach (var run in container.Descendants<Run>())
                        ApplyRunFormatting(EnsureRunProperties(run), key, value);
                    // Also update paragraph mark run properties so new runs inherit formatting
                    var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    ApplyRunFormatting(markRPr, key, value);
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        if (kind == "header")
            mainPart.HeaderParts.ElementAt(index).Header?.Save();
        else
            mainPart.FooterParts.ElementAt(index).Footer?.Save();
    }

    // Border style format: "style" or "style;size" or "style;size;color" or "style;size;color;space"
    // Styles: none, single, thick, double, dotted, dashed, dotDash, dotDotDash, triple,
    //         thinThickSmallGap, thickThinSmallGap, thinThickThinSmallGap,
    //         thinThickMediumGap, thickThinMediumGap, thinThickThinMediumGap,
    //         thinThickLargeGap, thickThinLargeGap, thinThickThinLargeGap, wave, doubleWave, threeDEmboss, threeDEngrave
    private static BorderValues ParseBorderStyle(string style) => style.ToLowerInvariant() switch
    {
        "none" => BorderValues.None,
        "nil" => BorderValues.Nil,
        "single" or "thin" => BorderValues.Single,
        "thick" or "medium" => BorderValues.Thick,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        "dotdash" => BorderValues.DotDash,
        "dotdotdash" => BorderValues.DotDotDash,
        "triple" => BorderValues.Triple,
        "thinthicksmallgap" => BorderValues.ThinThickSmallGap,
        "thickthinsmallgap" => BorderValues.ThickThinSmallGap,
        "thinthickthinsmallgap" => BorderValues.ThinThickThinSmallGap,
        "thinthickmediumgap" => BorderValues.ThinThickMediumGap,
        "thickthinmediumgap" => BorderValues.ThickThinMediumGap,
        "thinthickthinmediumgap" => BorderValues.ThinThickThinMediumGap,
        "thinthicklargegap" => BorderValues.ThinThickLargeGap,
        "thickthinlargegap" => BorderValues.ThickThinLargeGap,
        "thinthickthinlargegap" => BorderValues.ThinThickThinLargeGap,
        "wave" => BorderValues.Wave,
        "doublewave" => BorderValues.DoubleWave,
        "threedembed" or "3demboss" => BorderValues.ThreeDEmboss,
        "threedengrave" or "3dengrave" => BorderValues.ThreeDEngrave,
        _ => BorderValues.Single
    };

    private static (BorderValues style, uint size, string? color, uint space) ParseBorderValue(string value)
    {
        var parts = value.Split(';');
        var style = ParseBorderStyle(parts[0]);
        uint size = parts.Length > 1 && uint.TryParse(parts[1], out var s) ? s
            : style == BorderValues.Nil ? 0u
            : style == BorderValues.Thick ? 12u : 4u;
        string? color = parts.Length > 2 ? parts[2].TrimStart('#').ToUpperInvariant() : null;
        uint space = parts.Length > 3 && uint.TryParse(parts[3], out var sp) ? sp : 0u;
        return (style, size, color, space);
    }

    private static T MakeBorder<T>(BorderValues style, uint size, string? color, uint space) where T : BorderType, new()
    {
        var b = new T { Val = style, Size = size, Space = space };
        if (color != null) b.Color = color;
        return b;
    }

    /// <summary>
    /// Apply a paragraph-level property. Returns true if handled, false if not recognized.
    /// Handles: style, alignment, indent, spacing, keepNext, keepLines, pageBreakBefore, widowControl, shading, pbdr.
    /// </summary>
    private static bool ApplyParagraphLevelProperty(ParagraphProperties pProps, string key, string value)
    {
        switch (key.ToLowerInvariant())
        {
            case "style":
                pProps.ParagraphStyleId = new ParagraphStyleId { Val = value };
                return true;
            case "alignment":
                pProps.Justification = new Justification
                {
                    Val = value.ToLowerInvariant() switch
                    {
                        "center" => JustificationValues.Center,
                        "right" => JustificationValues.Right,
                        "justify" => JustificationValues.Both,
                        _ => JustificationValues.Left
                    }
                };
                return true;
            case "firstlineindent":
                var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indent.FirstLine = ((long)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 480)).ToString();
                indent.Hanging = null;
                return true;
            case "leftindent" or "indentleft":
                var indentL = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentL.Left = value;
                return true;
            case "rightindent" or "indentright":
                var indentR = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentR.Right = value;
                return true;
            case "hangingindent" or "hanging":
                var indentH = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentH.Hanging = value;
                indentH.FirstLine = null;
                return true;
            case "keepnext":
                if (IsTruthy(value)) pProps.KeepNext ??= new KeepNext();
                else pProps.KeepNext = null;
                return true;
            case "keeplines" or "keeptogether":
                if (IsTruthy(value)) pProps.KeepLines ??= new KeepLines();
                else pProps.KeepLines = null;
                return true;
            case "pagebreakbefore":
                if (IsTruthy(value)) pProps.PageBreakBefore ??= new PageBreakBefore();
                else pProps.PageBreakBefore = null;
                return true;
            case "widowcontrol":
                if (IsTruthy(value)) pProps.WidowControl ??= new WidowControl();
                else pProps.WidowControl = null;
                return true;
            case "shading" or "shd":
                var shdParts = value.Split(';');
                var shd = new Shading();
                if (shdParts.Length == 1)
                {
                    shd.Val = ShadingPatternValues.Clear;
                    shd.Fill = shdParts[0].TrimStart('#').ToUpperInvariant();
                }
                else if (shdParts.Length >= 2)
                {
                    shd.Val = new ShadingPatternValues(shdParts[0]);
                    shd.Fill = shdParts[1].TrimStart('#').ToUpperInvariant();
                    if (shdParts.Length >= 3) shd.Color = shdParts[2].TrimStart('#').ToUpperInvariant();
                }
                pProps.Shading = shd;
                return true;
            case "spacebefore":
                var spacingBefore = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingBefore.Before = value;
                return true;
            case "spaceafter":
                var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingAfter.After = value;
                return true;
            case "linespacing":
                var spacingLine = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingLine.Line = value;
                spacingLine.LineRule = LineSpacingRuleValues.Auto;
                return true;
            case "numid":
                var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr.NumberingId = new NumberingId { Val = int.Parse(value) };
                return true;
            case "numlevel" or "ilvl":
                var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr2.NumberingLevelReference = new NumberingLevelReference { Val = int.Parse(value) };
                return true;
            case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
                ApplyParagraphBorders(pProps, key, value);
                return true;
            default:
                return false;
        }
    }

    private static void ApplyParagraphBorders(ParagraphProperties pProps, string key, string value)
    {
        var borders = pProps.ParagraphBorders ?? pProps.AppendChild(new ParagraphBorders());
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "pbdr.all" or "pbdr":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "pbdr.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "pbdr.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "pbdr.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "pbdr.between":
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.bar":
                borders.BarBorder = MakeBorder<BarBorder>(style, size, color, space);
                break;
        }
    }

    private static void ApplyTableBorders(TableProperties tblPr, string key, string value)
    {
        var borders = tblPr.TableBorders ?? tblPr.AppendChild(new TableBorders());
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space);
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.insideh" or "border.horizontal":
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space);
                break;
            case "border.insidev" or "border.vertical":
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space);
                break;
        }
    }

    private static void ApplyCellBorders(TableCellProperties tcPr, string key, string value)
    {
        var borders = tcPr.TableCellBorders ?? tcPr.AppendChild(new TableCellBorders());
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
        }
    }
}
