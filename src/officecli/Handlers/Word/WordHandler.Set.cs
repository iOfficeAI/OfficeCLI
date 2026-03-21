// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // Batch Set: if path looks like a selector (not starting with /), Query → Set each
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
        {
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

        // Document-level properties (including find/replace)
        if (path == "/" || path == "")
        {
            // Find & Replace: special handling before document properties
            if (properties.TryGetValue("find", out var findText) && properties.TryGetValue("replace", out var replaceText))
            {
                var count = FindAndReplace(findText, replaceText, properties.GetValueOrDefault("scope", "all"));
                properties.Remove("find");
                properties.Remove("replace");
                properties.Remove("scope");
                // If there are remaining properties, apply them as document properties
                if (properties.Count > 0)
                    SetDocumentProperties(properties);
                _doc.MainDocumentPart?.Document?.Save();
                return unsupported;
            }

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
                                var clr = "#" + SanitizeHex(value);
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

        // Field paths: /field[N]
        var fieldSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/field\[(\d+)\]$");
        if (fieldSetMatch.Success)
        {
            var fieldIdx = int.Parse(fieldSetMatch.Groups[1].Value);
            var allFields = FindFields();
            if (fieldIdx < 1 || fieldIdx > allFields.Count)
                throw new ArgumentException($"Field {fieldIdx} not found (total: {allFields.Count})");

            var field = allFields[fieldIdx - 1];

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "instruction" or "instr":
                        field.InstrCode.Text = value.StartsWith(" ") ? value : $" {value} ";
                        // Auto-mark dirty when instruction changes
                        var beginCharI = field.BeginRun.GetFirstChild<FieldChar>();
                        if (beginCharI != null) beginCharI.Dirty = true;
                        break;
                    case "text" or "result":
                        // Replace result text (between separate and end)
                        if (field.ResultRuns.Count > 0)
                        {
                            // Set text on first result run, clear the rest
                            var firstResultText = field.ResultRuns[0].GetFirstChild<Text>();
                            if (firstResultText != null)
                                firstResultText.Text = value;
                            else
                                field.ResultRuns[0].AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                            for (int ri = 1; ri < field.ResultRuns.Count; ri++)
                            {
                                var t = field.ResultRuns[ri].GetFirstChild<Text>();
                                if (t != null) t.Text = "";
                            }
                        }
                        break;
                    case "dirty":
                        var beginCharD = field.BeginRun.GetFirstChild<FieldChar>();
                        if (beginCharD != null) beginCharD.Dirty = IsTruthy(value);
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
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
            // Report any keys besides "text" as unsupported
            foreach (var k in properties.Keys)
            {
                if (!k.Equals("text", StringComparison.OrdinalIgnoreCase))
                    unsupported.Add(k);
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
            // Report any keys besides "text" as unsupported
            foreach (var k in properties.Keys)
            {
                if (!k.Equals("text", StringComparison.OrdinalIgnoreCase))
                    unsupported.Add(k);
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
                            "nextpage" or "next" => SectionMarkValues.NextPage,
                            "continuous" => SectionMarkValues.Continuous,
                            "evenpage" or "even" => SectionMarkValues.EvenPage,
                            "oddpage" or "odd" => SectionMarkValues.OddPage,
                            _ => throw new ArgumentException($"Invalid section break type: '{value}'. Valid values: nextPage, continuous, evenPage, oddPage.")
                        };
                        break;
                    case "pagewidth" or "pageWidth":
                        EnsureSectPrPageSize(sectPr).Width = ParseHelpers.SafeParseUint(value, "pagewidth");
                        break;
                    case "pageheight" or "pageHeight":
                        EnsureSectPrPageSize(sectPr).Height = ParseHelpers.SafeParseUint(value, "pageheight");
                        break;
                    case "orientation":
                    {
                        var ps = EnsureSectPrPageSize(sectPr);
                        var isLandscape = value.ToLowerInvariant() == "landscape";
                        ps.Orient = isLandscape
                            ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
                        // Default to A4 if no dimensions set
                        var w = ps.Width?.Value ?? 11906u;
                        var h = ps.Height?.Value ?? 16838u;
                        // Swap width/height if orientation changes and dimensions are misaligned
                        if ((isLandscape && w < h) || (!isLandscape && w > h))
                        {
                            ps.Width = h;
                            ps.Height = w;
                        }
                        break;
                    }
                    case "margintop":
                        EnsureSectPrPageMargin(sectPr).Top = ParseHelpers.SafeParseInt(value, "margintop");
                        break;
                    case "marginbottom":
                        EnsureSectPrPageMargin(sectPr).Bottom = ParseHelpers.SafeParseInt(value, "marginbottom");
                        break;
                    case "marginleft":
                        EnsureSectPrPageMargin(sectPr).Left = ParseHelpers.SafeParseUint(value, "marginleft");
                        break;
                    case "marginright":
                        EnsureSectPrPageMargin(sectPr).Right = ParseHelpers.SafeParseUint(value, "marginright");
                        break;
                    case "columns" or "cols" or "col":
                    {
                        // Equal-width columns: "3" or "3,720" (count,space in twips)
                        var eqCols = EnsureColumns(sectPr);
                        var colParts = value.Split(',');
                        if (!short.TryParse(colParts[0], out var colCount))
                            throw new ArgumentException($"Invalid 'columns' value: '{value}'. Expected an integer or integer,space (e.g. '3' or '3,720').");
                        eqCols.ColumnCount = (Int16Value)colCount;
                        eqCols.EqualWidth = true;
                        if (colParts.Length > 1)
                            eqCols.Space = colParts[1];
                        else
                            eqCols.Space ??= "720"; // default ~1.27cm
                        // Remove any individual column definitions for equal width
                        eqCols.RemoveAllChildren<Column>();
                        break;
                    }
                    case "colwidths":
                    {
                        // Custom column widths: "3000,720,2000,720,3000"
                        // Alternating: width,space,width,space,...,width
                        var cwCols = EnsureColumns(sectPr);
                        cwCols.EqualWidth = false;
                        cwCols.RemoveAllChildren<Column>();
                        var vals = value.Split(',');
                        int colCount = 0;
                        for (int ci = 0; ci < vals.Length; ci += 2)
                        {
                            var col = new Column { Width = vals[ci] };
                            if (ci + 1 < vals.Length)
                                col.Space = vals[ci + 1];
                            cwCols.AppendChild(col);
                            colCount++;
                        }
                        cwCols.ColumnCount = (Int16Value)(short)colCount;
                        break;
                    }
                    case "separator" or "sep":
                    {
                        var sepCols = EnsureColumns(sectPr);
                        sepCols.Separator = IsTruthy(value);
                        break;
                    }
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
                        rPr5.Color = new Color { Val = SanitizeHex(value) };
                        break;
                    case "underline":
                    {
                        var rPrU = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        var ulVal = value.ToLowerInvariant() switch
                        {
                            "true" or "single" => "single",
                            "double" => "double",
                            "false" or "none" => "none",
                            _ => value
                        };
                        rPrU.Underline = new Underline { Val = new UnderlineValues(ulVal) };
                        break;
                    }
                    case "strike" or "strikethrough":
                        var rPrS = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        rPrS.Strike = IsTruthy(value) ? new Strike() : null;
                        break;
                    case "alignment":
                        var pPr = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());
                        pPr.Justification = new Justification { Val = ParseJustification(value) };
                        break;
                    case "spacebefore" or "spaceBefore":
                        var pPr2 = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());
                        var sp2 = pPr2.SpacingBetweenLines ?? (pPr2.SpacingBetweenLines = new SpacingBetweenLines());
                        sp2.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                        break;
                    case "spaceafter" or "spaceAfter":
                        var pPr3 = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());
                        var sp3 = pPr3.SpacingBetweenLines ?? (pPr3.SpacingBetweenLines = new SpacingBetweenLines());
                        sp3.After = SpacingConverter.ParseWordSpacing(value).ToString();
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

        // Clone element for rollback on failure (atomic: no partial modifications)
        var elementBackup = element.CloneNode(true);
        try
        {
        return SetElement(element, properties);
        }
        catch
        {
            // Rollback: restore element to pre-modification state
            element.Parent?.ReplaceChild(elementBackup, element);
            throw;
        }
    }

    private List<string> SetElement(OpenXmlElement element, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

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

        // SDT (Content Control) handling — both SdtBlock and SdtRun
        if (element is SdtBlock sdtBlock || element is SdtRun)
        {
            var sdtProps = element is SdtBlock sb
                ? sb.SdtProperties
                : ((SdtRun)element).SdtProperties;
            sdtProps ??= element.PrependChild(new SdtProperties());

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "alias" or "name":
                        var existingAlias = sdtProps.GetFirstChild<SdtAlias>();
                        if (existingAlias != null) existingAlias.Val = value;
                        else sdtProps.AppendChild(new SdtAlias { Val = value });
                        break;
                    case "tag":
                        var existingTag = sdtProps.GetFirstChild<Tag>();
                        if (existingTag != null) existingTag.Val = value;
                        else sdtProps.AppendChild(new Tag { Val = value });
                        break;
                    case "lock":
                        var existingLock = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
                        var lockEnum = value.ToLowerInvariant() switch
                        {
                            "contentlocked" or "content" => LockingValues.ContentLocked,
                            "sdtlocked" or "sdt" => LockingValues.SdtLocked,
                            "sdtcontentlocked" or "both" => LockingValues.SdtContentLocked,
                            "unlocked" or "none" => LockingValues.Unlocked,
                            _ => throw new ArgumentException($"Invalid lock value: '{value}'. Valid values: unlocked, contentLocked, sdtLocked, sdtContentLocked.")
                        };
                        if (existingLock != null) existingLock.Val = lockEnum;
                        else sdtProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock { Val = lockEnum });
                        break;
                    case "text":
                        // Replace content text
                        if (element is SdtBlock sdtB)
                        {
                            var content = sdtB.SdtContentBlock;
                            if (content != null)
                            {
                                content.RemoveAllChildren();
                                content.AppendChild(new Paragraph(
                                    new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })));
                            }
                        }
                        else if (element is SdtRun sdtR)
                        {
                            var content = sdtR.SdtContentRun;
                            if (content != null)
                            {
                                content.RemoveAllChildren();
                                content.AppendChild(
                                    new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                            }
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
                            Val = ParseHighlightColor(value)
                        };
                        break;
                    case "color":
                        EnsureRunProperties(run).Color = new Color { Val = SanitizeHex(value) };
                        break;
                    case "underline":
                    {
                        var ulVal = value.ToLowerInvariant() switch
                        {
                            "true" => "single",
                            "false" or "none" => "none",
                            _ => value
                        };
                        EnsureRunProperties(run).Underline = new Underline
                        {
                            Val = new UnderlineValues(ulVal)
                        };
                        break;
                    }
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
                            shd.Fill = SanitizeHex(shdParts[0]);
                        }
                        else if (shdParts.Length >= 2)
                        {
                            WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = SanitizeHex(shdParts[1]);
                            if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
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
                    case "path" or "src":
                    {
                        // Replace image source in a run containing a Drawing
                        var drawingSrc = run.GetFirstChild<Drawing>();
                        var blip = drawingSrc?.Descendants<A.Blip>().FirstOrDefault();
                        if (blip == null) { unsupported.Add(key); break; }
                        if (!File.Exists(value))
                            throw new FileNotFoundException($"Image file not found: {value}");

                        var mainPartImg = _doc.MainDocumentPart!;
                        var imgExt = Path.GetExtension(value).ToLowerInvariant();
                        var imgType = imgExt switch
                        {
                            ".png" => DocumentFormat.OpenXml.Packaging.ImagePartType.Png,
                            ".jpg" or ".jpeg" => DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg,
                            ".gif" => DocumentFormat.OpenXml.Packaging.ImagePartType.Gif,
                            ".bmp" => DocumentFormat.OpenXml.Packaging.ImagePartType.Bmp,
                            ".tif" or ".tiff" => DocumentFormat.OpenXml.Packaging.ImagePartType.Tiff,
                            ".emf" => DocumentFormat.OpenXml.Packaging.ImagePartType.Emf,
                            ".wmf" => DocumentFormat.OpenXml.Packaging.ImagePartType.Wmf,
                            ".svg" => DocumentFormat.OpenXml.Packaging.ImagePartType.Svg,
                            _ => throw new ArgumentException($"Unsupported image format: {imgExt}")
                        };

                        // Remove old image part to avoid storage bloat
                        var oldEmbedId = blip.Embed?.Value;
                        if (oldEmbedId != null)
                        {
                            try { mainPartImg.DeletePart(oldEmbedId); } catch { }
                        }

                        var newImgPart = mainPartImg.AddImagePart(imgType);
                        using (var stream = File.OpenRead(value))
                            newImgPart.FeedData(stream);
                        blip.Embed = mainPartImg.GetIdOfPart(newImgPart);
                        break;
                    }
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
                    case "textoutline":
                        ApplyW14TextEffect(run, "textOutline", value, BuildW14TextOutline);
                        break;
                    case "textfill":
                        ApplyW14TextEffect(run, "textFill", value, BuildW14TextFill);
                        break;
                    case "w14shadow":
                        ApplyW14TextEffect(run, "shadow", value, BuildW14Shadow);
                        break;
                    case "w14glow":
                        ApplyW14TextEffect(run, "glow", value, BuildW14Glow);
                        break;
                    case "w14reflection":
                        ApplyW14TextEffect(run, "reflection", value, BuildW14Reflection);
                        break;
                    case "formula":
                    {
                        // Replace this run with an inline oMath in the same position
                        var mathContent = FormulaParser.Parse(value);
                        M.OfficeMath oMath = mathContent is M.OfficeMath dm
                            ? dm : new M.OfficeMath(mathContent.CloneNode(true));
                        run.InsertAfterSelf(oMath);
                        run.Remove();
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
                    case "formula":
                    {
                        // Replace paragraph content with OMML equation in-place
                        foreach (var child in para.ChildElements
                            .Where(c => c is not ParagraphProperties).ToList())
                            child.Remove();
                        var mathContent = FormulaParser.Parse(value);
                        M.OfficeMath oMath = mathContent is M.OfficeMath dm
                            ? dm : new M.OfficeMath(mathContent.CloneNode(true));
                        para.AppendChild(new M.Paragraph(oMath));
                        break;
                    }
                    case "liststyle":
                        ApplyListStyle(para, value);
                        break;
                    case "start":
                        SetListStartValue(para, ParseHelpers.SafeParseInt(value, "start"));
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
                                        rPr.Color = new Color { Val = SanitizeHex(value) };
                                        break;
                                    case "highlight":
                                        rPr.Highlight = new Highlight { Val = ParseHighlightColor(value) };
                                        break;
                                    case "underline":
                                    {
                                        var ulVal = value.ToLowerInvariant() switch { "true" => "single", "false" or "none" => "none", _ => value };
                                        rPr.Underline = new Underline { Val = new UnderlineValues(ulVal) };
                                        break;
                                    }
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
                                    pmrp.AppendChild(new Color { Val = SanitizeHex(value) });
                                    break;
                                case "highlight":
                                    pmrp.RemoveAllChildren<Highlight>();
                                    pmrp.AppendChild(new Highlight { Val = ParseHighlightColor(value) });
                                    break;
                                case "underline":
                                {
                                    var ulVal = value.ToLowerInvariant() switch { "true" => "single", "false" or "none" => "none", _ => value };
                                    pmrp.RemoveAllChildren<Underline>();
                                    pmrp.AppendChild(new Underline { Val = new UnderlineValues(ulVal) });
                                    break;
                                }
                                case "strike":
                                    pmrp.RemoveAllChildren<Strike>();
                                    if (IsTruthy(value)) pmrp.AppendChild(new Strike());
                                    break;
                            }
                        }
                        break;
                    case "shd" or "shading" or "fill":
                        var shdParts = value.Split(';');
                        if (shdParts.Length >= 3 && shdParts[0].Equals("gradient", StringComparison.OrdinalIgnoreCase))
                        {
                            // gradient;startColor;endColor[;angle]  e.g. gradient;FF0000;0000FF;90
                            var startColor = SanitizeHex(shdParts[1]);
                            var endColor = SanitizeHex(shdParts[2]);
                            // Warn if color positions look like numbers (likely swapped with angle)
                            if (int.TryParse(shdParts[1], out _) && shdParts[1].Length <= 3)
                                Console.Error.WriteLine($"Warning: '{shdParts[1]}' looks like an angle, not a color. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                            if (int.TryParse(shdParts[2], out _) && shdParts[2].Length <= 3)
                                Console.Error.WriteLine($"Warning: '{shdParts[2]}' looks like an angle, not a color. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                            int angleDeg = 180;
                            if (shdParts.Length >= 4)
                            {
                                if (!int.TryParse(shdParts[3], out angleDeg))
                                {
                                    Console.Error.WriteLine($"Warning: invalid gradient angle '{shdParts[3]}', expected integer. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                                    angleDeg = 180;
                                }
                            }
                            ApplyCellGradient(tcPr, startColor, endColor, angleDeg);
                        }
                        else
                        {
                            // Remove any existing gradient
                            RemoveCellGradient(tcPr);
                            var shd = new Shading();
                            if (shdParts.Length == 1)
                            {
                                shd.Val = ShadingPatternValues.Clear;
                                shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[0]).Rgb;
                            }
                            else if (shdParts.Length >= 2)
                            {
                                WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                                shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[1]).Rgb;
                                if (shdParts.Length >= 3) shd.Color = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[2]).Rgb;
                            }
                            tcPr.Shading = shd;
                        }
                        break;
                    case "alignment":
                        var alignVal = ParseJustification(value);
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
                                "top" => TableVerticalAlignmentValues.Top,
                                "center" => TableVerticalAlignmentValues.Center,
                                "bottom" => TableVerticalAlignmentValues.Bottom,
                                _ => throw new ArgumentException($"Invalid valign value: '{value}'. Valid values: top, center, bottom.")
                            }
                        };
                        break;
                    case "width":
                        tcPr.TableCellWidth = new TableCellWidth { Width = ParseHelpers.SafeParseUint(value, "width").ToString(), Type = TableWidthUnitValues.Dxa };
                        break;
                    case "padding":
                    {
                        var dxa = ParseHelpers.SafeParseUint(value, "padding").ToString();
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
                                _ => throw new ArgumentException($"Invalid textDirection value: '{value}'. Valid values: lrtb, btlr, tbrl, horizontal, vertical.")
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
                        var newSpan = ParseHelpers.SafeParseInt(value, "gridspan");
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
                        trPr.AppendChild(new TableRowHeight { Val = ParseTwips(value), HeightType = HeightRuleValues.AtLeast });
                        break;
                    case "height.exact":
                        trPr.GetFirstChild<TableRowHeight>()?.Remove();
                        trPr.AppendChild(new TableRowHeight { Val = ParseTwips(value), HeightType = HeightRuleValues.Exact });
                        break;
                    case "header":
                        if (IsTruthy(value))
                        {
                            if (trPr.GetFirstChild<TableHeader>() == null)
                                trPr.AppendChild(new TableHeader());
                        }
                        else
                            trPr.RemoveAllChildren<TableHeader>();
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
                                "left" => TableRowAlignmentValues.Left,
                                "center" => TableRowAlignmentValues.Center,
                                "right" => TableRowAlignmentValues.Right,
                                _ => throw new ArgumentException($"Invalid table alignment value: '{value}'. Valid values: left, center, right.")
                            }
                        };
                        break;
                    case "width":
                        if (value.EndsWith('%'))
                        {
                            var pct = ParseHelpers.SafeParseInt(value.TrimEnd('%'), "width") * 50; // OOXML pct = percent * 50
                            tblPr.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                        }
                        else
                        {
                            tblPr.TableWidth = new TableWidth { Width = ParseHelpers.SafeParseUint(value, "width").ToString(), Type = TableWidthUnitValues.Dxa };
                        }
                        break;
                    case "indent":
                        tblPr.TableIndentation = new TableIndentation { Width = ParseHelpers.SafeParseInt(value, "indent"), Type = TableWidthUnitValues.Dxa };
                        break;
                    case "cellspacing":
                        tblPr.TableCellSpacing = new TableCellSpacing { Width = ParseHelpers.SafeParseUint(value, "cellspacing").ToString(), Type = TableWidthUnitValues.Dxa };
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
                        var paddingVal = ParseHelpers.SafeParseInt(dxa, "padding");
                        cm.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                        cm.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                        cm.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
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
        _ => throw new ArgumentException($"Invalid border style: '{style}'. Valid values: single, thick, double, dotted, dashed, none, triple, wave, etc.")
    };

    private static (BorderValues style, uint size, string? color, uint space) ParseBorderValue(string value)
    {
        var parts = value.Split(';');
        var style = ParseBorderStyle(parts[0]);
        uint size;
        if (parts.Length > 1)
        {
            if (!uint.TryParse(parts[1], out size))
            {
                Console.Error.WriteLine($"Warning: invalid border size '{parts[1]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
                size = style == BorderValues.Thick ? 12u : 4u;
            }
        }
        else
            size = style == BorderValues.Nil ? 0u : style == BorderValues.Thick ? 12u : 4u;
        string? color = parts.Length > 2 ? SanitizeHex(parts[2]) : null;
        uint space = 0u;
        if (parts.Length > 3 && !uint.TryParse(parts[3], out space))
        {
            Console.Error.WriteLine($"Warning: invalid border space '{parts[3]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
            space = 0u;
        }
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
            case "alignment" or "align":
                pProps.Justification = new Justification { Val = ParseJustification(value) };
                return true;
            case "firstlineindent":
                var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indent.FirstLine = value; // raw twips, consistent with Get and other indent properties
                indent.Hanging = null;
                return true;
            case "leftindent" or "indentleft":
                var indentL = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentL.Left = ParseHelpers.SafeParseUint(value, "leftindent").ToString();
                return true;
            case "rightindent" or "indentright":
                var indentR = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentR.Right = ParseHelpers.SafeParseUint(value, "rightindent").ToString();
                return true;
            case "hangingindent" or "hanging":
                var indentH = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentH.Hanging = ParseHelpers.SafeParseUint(value, "hangingindent").ToString();
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
                    shd.Fill = SanitizeHex(shdParts[0]);
                }
                else if (shdParts.Length >= 2)
                {
                    WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                    shd.Fill = SanitizeHex(shdParts[1]);
                    if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
                }
                pProps.Shading = shd;
                return true;
            case "spacebefore":
                var spacingBefore = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingBefore.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "spaceafter":
                var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingAfter.After = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "linespacing":
                var spacingLine = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                var (lsTwips, lsIsMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
                spacingLine.Line = lsTwips.ToString();
                spacingLine.LineRule = lsIsMultiplier ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
                return true;
            case "numid":
                var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr.NumberingId = new NumberingId { Val = ParseHelpers.SafeParseInt(value, "numid") };
                return true;
            case "numlevel" or "ilvl":
                var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr2.NumberingLevelReference = new NumberingLevelReference { Val = ParseHelpers.SafeParseInt(value, "numlevel") };
                return true;
            case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
            case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right":
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
            case "pbdr.all" or "pbdr" or "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.top" or "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "pbdr.bottom" or "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "pbdr.left" or "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "pbdr.right" or "border.right":
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
            case "border.tl2br":
                borders.TopLeftToBottomRightCellBorder = MakeBorder<TopLeftToBottomRightCellBorder>(style, size, color, space);
                break;
            case "border.tr2bl":
                borders.TopRightToBottomLeftCellBorder = MakeBorder<TopRightToBottomLeftCellBorder>(style, size, color, space);
                break;
        }
    }

    /// <summary>
    /// Apply gradient fill to a Word table cell using mc:AlternativeContent with w14:gradFill.
    /// Fallback is a solid shading with the start color.
    /// </summary>
    private static void ApplyCellGradient(TableCellProperties tcPr, string startColor, string endColor, int angleDeg)
    {
        // Sanitize colors: strip 8-char RRGGBBAA to 6-char RGB (w14:srgbClr requires 6 chars)
        var (startRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(startColor);
        var (endRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(endColor);

        // Remove existing shading/gradient
        RemoveCellGradient(tcPr);
        tcPr.Shading?.Remove();

        // Set fallback solid fill
        tcPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = startRgb };

        // Build w14:gradFill XML via raw OpenXml
        var w14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        // Convert angle to OOXML 60000ths of a degree
        var angleOoxml = angleDeg * 60000;

        var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);
        acElement.InnerXml = $@"<mc:Choice xmlns:mc=""{mcNs}"" xmlns:w14=""{w14Ns}"" Requires=""w14"">
    <w14:gradFill>
      <w14:gsLst>
        <w14:gs w14:pos=""0"">
          <w14:srgbClr w14:val=""{startRgb}""/>
        </w14:gs>
        <w14:gs w14:pos=""100000"">
          <w14:srgbClr w14:val=""{endRgb}""/>
        </w14:gs>
      </w14:gsLst>
      <w14:lin w14:ang=""{angleOoxml}"" w14:scaled=""1""/>
    </w14:gradFill>
  </mc:Choice>";

        tcPr.AppendChild(acElement);
    }

    /// <summary>
    /// Remove any existing gradient mc:AlternateContent from table cell properties.
    /// </summary>
    private static void RemoveCellGradient(TableCellProperties tcPr)
    {
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        var existing = tcPr.ChildElements
            .Where(e => e.LocalName == "AlternateContent" && e.NamespaceUri == mcNs)
            .ToList();
        foreach (var e in existing) e.Remove();
    }

    /// <summary>
    /// Parse twips from a string with optional unit suffix: "1.5cm", "0.5in", "36pt", or raw twips.
    /// 1 inch = 1440 twips, 1 cm = 567 twips, 1 pt = 20 twips.
    /// </summary>
    internal static uint ParseTwips(string value)
    {
        value = value.Trim();
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (cm)");
            return (uint)Math.Round(num * 567);
        }
        if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (in)");
            return (uint)Math.Round(num * 1440);
        }
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (pt)");
            return (uint)Math.Round(num * 20);
        }
        return ParseHelpers.SafeParseUint(value, "twips");
    }
}
