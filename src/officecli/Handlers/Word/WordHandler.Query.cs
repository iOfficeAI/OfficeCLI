// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Path cannot be empty.");
        if (path == "/")
            return GetRootNode(depth);

        // Handle /watermark path
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
        {
            var node = new DocumentNode { Path = "/watermark", Type = "watermark" };
            var wmText = FindWatermark();
            if (wmText == null)
            {
                node.Text = "(no watermark)";
                return node;
            }
            node.Text = wmText;

            // Extract properties from VML shape in headers
            foreach (var hp in _doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<DocumentFormat.OpenXml.Packaging.HeaderPart>())
            {
                if (hp.Header == null) continue;
                foreach (var pict in hp.Header.Descendants<Picture>())
                {
                    var xml = pict.InnerXml;
                    if (!xml.Contains("WaterMark", StringComparison.OrdinalIgnoreCase)) continue;

                    node.Format["text"] = wmText;

                    // Extract fillcolor
                    var fillMatch = System.Text.RegularExpressions.Regex.Match(xml, @"fillcolor=""([^""]*)""");
                    if (fillMatch.Success) node.Format["color"] = ParseHelpers.FormatHexColor(fillMatch.Groups[1].Value);

                    // Extract opacity
                    var opacityMatch = System.Text.RegularExpressions.Regex.Match(xml, @"opacity=""([^""]*)""");
                    if (opacityMatch.Success) node.Format["opacity"] = opacityMatch.Groups[1].Value;

                    // Extract font
                    var fontMatch = System.Text.RegularExpressions.Regex.Match(xml, @"font-family:&quot;([^&]*)&quot;");
                    if (fontMatch.Success) node.Format["font"] = fontMatch.Groups[1].Value;

                    // Extract rotation
                    var rotMatch = System.Text.RegularExpressions.Regex.Match(xml, @"rotation:(\d+)");
                    if (rotMatch.Success) node.Format["rotation"] = rotMatch.Groups[1].Value;

                    return node;
                }
            }
            return node;
        }

        // Handle header/footer paths
        var segments = ParsePath(path);
        if (segments.Count >= 1)
        {
            var firstName = segments[0].Name.ToLowerInvariant();
            if (firstName == "header" && segments.Count == 1)
            {
                var hIdx = (segments[0].Index ?? 1) - 1;
                return GetHeaderNode(hIdx, path, depth);
            }
            if (firstName == "footer" && segments.Count == 1)
            {
                var fIdx = (segments[0].Index ?? 1) - 1;
                return GetFooterNode(fIdx, path, depth);
            }
        }

        // Footnote/Endnote paths: /footnote[N], /endnote[N]
        var fnMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/footnote\[(\d+)\]$");
        if (fnMatch.Success)
        {
            var fnId = int.Parse(fnMatch.Groups[1].Value);
            var fn = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
            if (fn == null)
                throw new ArgumentException($"Footnote {fnId} not found");
            var fnNode = new DocumentNode { Path = path, Type = "footnote" };
            fnNode.Text = string.Join("", fn.Descendants<Text>().Select(t => t.Text));
            return fnNode;
        }
        var enMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/endnote\[(\d+)\]$");
        if (enMatch.Success)
        {
            var enId = int.Parse(enMatch.Groups[1].Value);
            var en = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
            if (en == null)
                throw new ArgumentException($"Endnote {enId} not found");
            var enNode = new DocumentNode { Path = path, Type = "endnote" };
            enNode.Text = string.Join("", en.Descendants<Text>().Select(t => t.Text));
            return enNode;
        }

        // TOC paths: /toc[N]
        var tocMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/toc\[(\d+)\]$");
        if (tocMatch.Success)
        {
            var tocIdx = int.Parse(tocMatch.Groups[1].Value);
            var tocParas = FindTocParagraphs();
            if (tocIdx < 1 || tocIdx > tocParas.Count)
                throw new ArgumentException($"TOC {tocIdx} not found (total: {tocParas.Count})");

            var tocPara = tocParas[tocIdx - 1];
            var instrText = string.Join("", tocPara.Descendants<FieldCode>().Select(fc => fc.Text));
            var tocNode = new DocumentNode { Path = path, Type = "toc" };
            tocNode.Text = instrText.Trim();

            // Parse field code switches
            var levelsMatch = System.Text.RegularExpressions.Regex.Match(instrText, @"\\o\s+""([^""]+)""");
            if (levelsMatch.Success) tocNode.Format["levels"] = levelsMatch.Groups[1].Value;
            tocNode.Format["hyperlinks"] = instrText.Contains("\\h");
            tocNode.Format["pageNumbers"] = !instrText.Contains("\\z");
            return tocNode;
        }

        // Field paths: /field[N]
        var fieldMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/field\[(\d+)\]$");
        if (fieldMatch.Success)
        {
            var fieldIdx = int.Parse(fieldMatch.Groups[1].Value);
            var allFields = FindFields();
            if (fieldIdx < 1 || fieldIdx > allFields.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"Field {fieldIdx} not found (total: {allFields.Count})" };
            return FieldToNode(allFields[fieldIdx - 1], path);
        }

        // Chart paths: /chart[N] or /chart[N]/series[K]
        var chartGetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartGetMatch.Success)
        {
            var chartIdx = int.Parse(chartGetMatch.Groups[1].Value);
            var allCharts = GetAllWordCharts();
            if (chartIdx < 1 || chartIdx > allCharts.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"Chart {chartIdx} not found" };

            var chartInfo = allCharts[chartIdx - 1];
            var chartNode = new DocumentNode { Path = $"/chart[{chartIdx}]", Type = "chart" };

            if (chartInfo.IsExtended)
            {
                // Extended chart (funnel, treemap, etc.)
                var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace;
                var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                if (cxType != null) chartNode.Format["chartType"] = cxType;
                // Title
                var cxTitle = cxChartSpace?.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
                var cxTitleText = cxTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
                if (cxTitleText != null) chartNode.Format["title"] = cxTitleText;
                // Count series
                var cxSeries = cxChartSpace!.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Series>().ToList();
                chartNode.Format["seriesCount"] = cxSeries.Count;
            }
            else
            {
                var chart = chartInfo.StandardPart!.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                if (chart != null)
                    Core.ChartHelper.ReadChartProperties(chart, chartNode, chartGetMatch.Groups[2].Success ? 1 : depth);
            }

            // If series sub-path requested, extract the specific series child
            if (chartGetMatch.Groups[2].Success)
            {
                var seriesIdx = int.Parse(chartGetMatch.Groups[2].Value);
                var seriesChildren = chartNode.Children.Where(c => c.Type == "series").ToList();
                if (seriesIdx < 1 || seriesIdx > seriesChildren.Count)
                    throw new ArgumentException($"Series {seriesIdx} not found (total: {seriesChildren.Count})");
                var seriesNode = seriesChildren[seriesIdx - 1];
                seriesNode.Path = path;
                return seriesNode;
            }
            return chartNode;
        }

        // Section paths: /section[N]
        var secMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/section\[(\d+)\]$");
        if (secMatch.Success)
        {
            var secIdx = int.Parse(secMatch.Groups[1].Value);
            var sectionProps = FindSectionProperties();
            if (secIdx < 1 || secIdx > sectionProps.Count)
                throw new ArgumentException($"Section {secIdx} not found (total: {sectionProps.Count})");

            var sectPr = sectionProps[secIdx - 1];
            var secNode = new DocumentNode { Path = path, Type = "section" };

            var sectType = sectPr.GetFirstChild<SectionType>();
            if (sectType?.Val?.Value != null)
                secNode.Format["type"] = sectType.Val.InnerText;
            var pageSize = sectPr.GetFirstChild<PageSize>();
            // Default to A4 size (11906 × 16838 twips) if no explicit page size
            var pgW = pageSize?.Width?.Value ?? 11906u;
            var pgH = pageSize?.Height?.Value ?? 16838u;
            secNode.Format["pageWidth"] = FormatTwipsToCm(pgW);
            secNode.Format["pageHeight"] = FormatTwipsToCm(pgH);
            if (pageSize?.Orient?.Value != null) secNode.Format["orientation"] = pageSize.Orient.InnerText;
            var margin = sectPr.GetFirstChild<PageMargin>();
            if (margin?.Top?.Value != null) secNode.Format["marginTop"] = margin.Top.Value;
            if (margin?.Bottom?.Value != null) secNode.Format["marginBottom"] = margin.Bottom.Value;
            if (margin?.Left?.Value != null) secNode.Format["marginLeft"] = margin.Left.Value;
            if (margin?.Right?.Value != null) secNode.Format["marginRight"] = margin.Right.Value;

            // Column properties
            var cols = sectPr.GetFirstChild<Columns>();
            if (cols != null)
            {
                secNode.Format["columns"] = cols.ColumnCount?.Value ?? 1;
                if (cols.Space?.Value != null) secNode.Format["columnSpace"] = cols.Space.Value;
                if (cols.EqualWidth?.Value != null) secNode.Format["equalWidth"] = cols.EqualWidth.Value;
                if (cols.Separator?.Value == true) secNode.Format["separator"] = true;
                var colDefs = cols.Elements<Column>().ToList();
                if (colDefs.Count > 0)
                {
                    var widths = colDefs.Select(c => c.Width?.Value ?? "0");
                    var spaces = colDefs.Select(c => c.Space?.Value ?? "0");
                    secNode.Format["colWidths"] = string.Join(",", widths);
                    secNode.Format["colSpaces"] = string.Join(",", spaces);
                }
            }
            return secNode;
        }

        // Style paths: /styles/StyleId
        var styleMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/styles/(.+)$");
        if (styleMatch.Success)
        {
            var styleId = styleMatch.Groups[1].Value;
            var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            var style = styles?.Elements<Style>().FirstOrDefault(s =>
                s.StyleId?.Value == styleId || s.StyleName?.Val?.Value == styleId);
            if (style == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"Style '{styleId}' not found" };

            var styleNode = new DocumentNode { Path = path, Type = "style" };
            styleNode.Text = style.StyleName?.Val?.Value ?? styleId;
            styleNode.Format["id"] = style.StyleId?.Value ?? "";
            styleNode.Format["name"] = style.StyleName?.Val?.Value ?? "";
            if (style.Type?.Value != null) styleNode.Format["type"] = style.Type.InnerText;
            if (style.BasedOn?.Val?.Value != null) styleNode.Format["basedOn"] = style.BasedOn.Val.Value;
            if (style.NextParagraphStyle?.Val?.Value != null) styleNode.Format["next"] = style.NextParagraphStyle.Val.Value;

            // Read run properties
            var rPr = style.StyleRunProperties;
            if (rPr != null)
            {
                if (rPr.RunFonts?.Ascii?.Value != null) styleNode.Format["font"] = rPr.RunFonts.Ascii.Value;
                if (rPr.FontSize?.Val?.Value != null) styleNode.Format["size"] = $"{int.Parse(rPr.FontSize.Val.Value) / 2.0:0.##}pt";
                if (rPr.Bold != null) styleNode.Format["bold"] = true;
                if (rPr.Italic != null) styleNode.Format["italic"] = true;
                if (rPr.Color?.Val?.Value != null) styleNode.Format["color"] = ParseHelpers.FormatHexColor(rPr.Color.Val.Value);
                else if (rPr.Color?.ThemeColor?.HasValue == true) styleNode.Format["color"] = rPr.Color.ThemeColor.InnerText;
                if (rPr.Underline?.Val != null) styleNode.Format["underline"] = rPr.Underline.Val.InnerText;
                if (rPr.Strike != null) styleNode.Format["strike"] = true;
            }

            // Read paragraph properties
            var pPr = style.StyleParagraphProperties;
            if (pPr != null)
            {
                if (pPr.Justification?.Val?.Value != null) styleNode.Format["alignment"] = pPr.Justification.Val.InnerText;
                if (pPr.SpacingBetweenLines?.Before?.Value != null) styleNode.Format["spaceBefore"] = SpacingConverter.FormatWordSpacing(pPr.SpacingBetweenLines.Before.Value);
                if (pPr.SpacingBetweenLines?.After?.Value != null) styleNode.Format["spaceAfter"] = SpacingConverter.FormatWordSpacing(pPr.SpacingBetweenLines.After.Value);
            }
            return styleNode;
        }

        var parts = ParsePath(path);
        var element = NavigateToElement(parts, out var ctx);
        if (element == null)
        {
            // Check if the path contains footnote/endnote/toc which are handled differently
            if (path.Contains("footnote") || path.Contains("endnote") || path.Contains("toc"))
                return new DocumentNode { Path = path, Type = "error", Text = $"Path not found: {path}" };
            var msg = $"Path not found: {path}";
            if (ctx != null) msg += $". {ctx}";
            throw new ArgumentException(msg);
        }

        return ElementToNode(element, path, depth);
    }

    /// <summary>Find all SectionProperties in the document (paragraph-level + body-level).</summary>
    private List<SectionProperties> FindSectionProperties()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return new();

        var result = new List<SectionProperties>();
        // Paragraph-level section properties (section breaks)
        foreach (var p in body.Elements<Paragraph>())
        {
            var sectPr = p.ParagraphProperties?.GetFirstChild<SectionProperties>();
            if (sectPr != null) result.Add(sectPr);
        }
        // Body-level section properties (last section)
        var bodySectPr = body.GetFirstChild<SectionProperties>();
        if (bodySectPr != null)
            result.Add(bodySectPr);
        else if (result.Count == 0)
        {
            // Always have at least one implicit section (the document body itself acts as a section)
            var implicitSectPr = new SectionProperties();
            body.AppendChild(implicitSectPr);
            result.Add(implicitSectPr);
        }
        return result;
    }

    /// <summary>
    /// Represents a complex field (fldChar begin → instrText → separate → result → end).
    /// </summary>
    private record FieldInfo(Run BeginRun, FieldCode InstrCode, Run? SeparateRun, List<Run> ResultRuns, Run EndRun, OpenXmlElement Container);

    /// <summary>Find all complex fields in the document body (and optionally headers/footers).</summary>
    private List<FieldInfo> FindFields()
    {
        var fields = new List<FieldInfo>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return fields;

        CollectFieldsFrom(body.Descendants<Run>(), fields, body);

        // Also search headers and footers
        foreach (var hp in _doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<DocumentFormat.OpenXml.Packaging.HeaderPart>())
            if (hp.Header != null) CollectFieldsFrom(hp.Header.Descendants<Run>(), fields, hp.Header);
        foreach (var fp in _doc.MainDocumentPart?.FooterParts ?? Enumerable.Empty<DocumentFormat.OpenXml.Packaging.FooterPart>())
            if (fp.Footer != null) CollectFieldsFrom(fp.Footer.Descendants<Run>(), fields, fp.Footer);

        return fields;
    }

    private static void CollectFieldsFrom(IEnumerable<Run> runs, List<FieldInfo> fields, OpenXmlElement container)
    {
        Run? beginRun = null;
        FieldCode? instrCode = null;
        Run? separateRun = null;
        var resultRuns = new List<Run>();
        bool inResult = false;

        foreach (var run in runs)
        {
            var fldChar = run.GetFirstChild<FieldChar>();
            if (fldChar != null)
            {
                var charType = fldChar.FieldCharType?.Value;
                if (charType == FieldCharValues.Begin)
                {
                    beginRun = run;
                    instrCode = null;
                    separateRun = null;
                    resultRuns.Clear();
                    inResult = false;
                }
                else if (charType == FieldCharValues.Separate)
                {
                    separateRun = run;
                    inResult = true;
                }
                else if (charType == FieldCharValues.End)
                {
                    if (beginRun != null && instrCode != null)
                    {
                        fields.Add(new FieldInfo(beginRun, instrCode, separateRun,
                            new List<Run>(resultRuns), run, container));
                    }
                    beginRun = null;
                    instrCode = null;
                    separateRun = null;
                    resultRuns.Clear();
                    inResult = false;
                }
            }
            else if (beginRun != null && !inResult)
            {
                var fc = run.GetFirstChild<FieldCode>();
                if (fc != null) instrCode = fc;
            }
            else if (inResult)
            {
                resultRuns.Add(run);
            }
        }
    }

    private static DocumentNode FieldToNode(FieldInfo field, string path)
    {
        var instr = field.InstrCode.Text?.Trim() ?? "";
        var resultText = string.Join("", field.ResultRuns.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));

        // Determine field type from instruction
        var fieldType = "field";
        var instrUpper = instr.TrimStart().Split(' ', 2)[0].ToUpperInvariant();
        if (!string.IsNullOrEmpty(instrUpper))
            fieldType = instrUpper.ToLowerInvariant(); // e.g., "page", "numpages", "date", "toc", "author"

        var node = new DocumentNode { Path = path, Type = "field" };
        node.Text = resultText;
        node.Format["instruction"] = instr;
        node.Format["fieldType"] = fieldType;

        // Check dirty flag
        var beginChar = field.BeginRun.GetFirstChild<FieldChar>();
        if (beginChar?.Dirty?.Value == true)
            node.Format["dirty"] = true;

        return node;
    }

    /// <summary>Find all paragraphs containing TOC field codes.</summary>
    private List<Paragraph> FindTocParagraphs()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return new();

        return body.Elements<Paragraph>()
            .Where(p => p.Descendants<FieldCode>().Any(fc =>
                fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)))
            .ToList();
    }

    private DocumentNode GetHeaderNode(int index, string path, int depth)
    {
        var mainPart = _doc.MainDocumentPart;
        var headerPart = mainPart?.HeaderParts.ElementAtOrDefault(index);
        if (headerPart?.Header == null)
            return new DocumentNode { Path = path, Type = "error", Text = $"Path not found: {path}" };

        var header = headerPart.Header;
        var node = new DocumentNode { Path = path, Type = "header" };
        node.Text = string.Concat(header.Descendants<Text>().Select(t => t.Text)).Trim();

        var relId = mainPart!.GetIdOfPart(headerPart);
        var body = mainPart.Document?.Body;
        if (body != null)
        {
            foreach (var sectPr in body.Elements<SectionProperties>())
                foreach (var href in sectPr.Elements<HeaderReference>())
                    if (href.Id?.Value == relId && href.Type?.Value != null)
                    {
                        node.Format["type"] = href.Type.InnerText;
                        break;
                    }
        }

        var firstRun = header.Descendants<Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var rp = firstRun.RunProperties;
            var font = rp.RunFonts?.Ascii?.Value ?? rp.RunFonts?.HighAnsi?.Value;
            if (font != null) node.Format["font"] = font;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = true;
            if (rp.Italic != null) node.Format["italic"] = true;
            if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            else if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
        }

        var firstPara = header.Elements<Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Justification?.Val != null)
            node.Format["alignment"] = firstPara.ParagraphProperties.Justification.Val.InnerText;

        node.ChildCount = header.Elements<Paragraph>().Count();
        if (depth > 0)
        {
            int pIdx = 0;
            foreach (var para in header.Elements<Paragraph>())
            {
                node.Children.Add(ElementToNode(para, $"{path}/p[{pIdx + 1}]", depth - 1));
                pIdx++;
            }
        }

        return node;
    }

    private DocumentNode GetFooterNode(int index, string path, int depth)
    {
        var mainPart = _doc.MainDocumentPart;
        var footerPart = mainPart?.FooterParts.ElementAtOrDefault(index);
        if (footerPart?.Footer == null)
            return new DocumentNode { Path = path, Type = "error", Text = $"Path not found: {path}" };

        var footer = footerPart.Footer;
        var node = new DocumentNode { Path = path, Type = "footer" };
        node.Text = string.Concat(footer.Descendants<Text>().Select(t => t.Text)).Trim();

        var relId = mainPart!.GetIdOfPart(footerPart);
        var body = mainPart.Document?.Body;
        if (body != null)
        {
            foreach (var sectPr in body.Elements<SectionProperties>())
                foreach (var fref in sectPr.Elements<FooterReference>())
                    if (fref.Id?.Value == relId && fref.Type?.Value != null)
                    {
                        node.Format["type"] = fref.Type.InnerText;
                        break;
                    }
        }

        var firstRun = footer.Descendants<Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var rp = firstRun.RunProperties;
            var font = rp.RunFonts?.Ascii?.Value ?? rp.RunFonts?.HighAnsi?.Value;
            if (font != null) node.Format["font"] = font;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = true;
            if (rp.Italic != null) node.Format["italic"] = true;
            if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            else if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
        }

        var firstPara = footer.Elements<Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Justification?.Val != null)
            node.Format["alignment"] = firstPara.ParagraphProperties.Justification.Val.InnerText;

        node.ChildCount = footer.Elements<Paragraph>().Count();
        if (depth > 0)
        {
            int pIdx = 0;
            foreach (var para in footer.Elements<Paragraph>())
            {
                node.Children.Add(ElementToNode(para, $"{path}/p[{pIdx + 1}]", depth - 1));
                pIdx++;
            }
        }

        return node;
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return results;

        // Simple selector parser: element[attr=value]
        var parsed = ParseSelector(selector);

        // Handle header/footer selectors
        if (parsed.Element is "header" or "footer")
        {
            var mainPart = _doc.MainDocumentPart!;
            if (parsed.Element == "header")
            {
                int hIdx = 0;
                foreach (var _ in mainPart.HeaderParts)
                {
                    var node = GetHeaderNode(hIdx, $"/header[{hIdx + 1}]", 0);
                    if (node.Type != "error")
                    {
                        if (parsed.ContainsText == null || (node.Text?.Contains(parsed.ContainsText) == true))
                            results.Add(node);
                    }
                    hIdx++;
                }
            }
            else
            {
                int fIdx = 0;
                foreach (var _ in mainPart.FooterParts)
                {
                    var node = GetFooterNode(fIdx, $"/footer[{fIdx + 1}]", 0);
                    if (node.Type != "error")
                    {
                        if (parsed.ContainsText == null || (node.Text?.Contains(parsed.ContainsText) == true))
                            results.Add(node);
                    }
                    fIdx++;
                }
            }
            return results;
        }

        // Handle field selector
        if (parsed.Element == "field")
        {
            var allFields = FindFields();
            for (int fi = 0; fi < allFields.Count; fi++)
            {
                var fieldNode = FieldToNode(allFields[fi], $"/field[{fi + 1}]");
                // Filter by :contains
                if (parsed.ContainsText != null)
                {
                    var instr = fieldNode.Format.TryGetValue("instruction", out var instrObj) ? instrObj?.ToString() : "";
                    if (instr == null || !instr.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                // Filter by attribute (e.g., field[fieldType=page] or field[fieldType!=page])
                bool matchAttrs = true;
                foreach (var (attrKey, rawVal) in parsed.Attributes)
                {
                    bool negate = rawVal.StartsWith("!");
                    var val = negate ? rawVal[1..] : rawVal;
                    var hasKey = fieldNode.Format.TryGetValue(attrKey, out var fmtVal);
                    bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                    if (negate ? matches : !matches)
                    { matchAttrs = false; break; }
                }
                if (matchAttrs) results.Add(fieldNode);
            }
            return results;
        }

        // Determine if main selector targets runs directly (no > parent)
        bool isRunSelector = parsed.ChildSelector == null &&
            (parsed.Element == "r" || parsed.Element == "run");
        bool isPictureSelector = parsed.ChildSelector == null &&
            (parsed.Element == "picture" || parsed.Element == "image" || parsed.Element == "img");
        bool isEquationSelector = parsed.ChildSelector == null &&
            (parsed.Element == "equation" || parsed.Element == "math" || parsed.Element == "formula");
        bool isBookmarkSelector = parsed.ChildSelector == null &&
            parsed.Element == "bookmark";
        bool isSdtSelector = parsed.ChildSelector == null &&
            (parsed.Element == "sdt" || parsed.Element == "contentcontrol");

        // Scheme B: generic XML fallback for unrecognized element types
        // Use GenericXmlQuery.ParseSelector which properly handles namespace prefixes (e.g., "a:ln")
        var genericParsed = GenericXmlQuery.ParseSelector(selector);
        bool isKnownType = string.IsNullOrEmpty(genericParsed.element)
            || genericParsed.element is "p" or "paragraph" or "r" or "run"
                or "picture" or "image" or "img"
                or "equation" or "math" or "formula"
                or "bookmark"
                or "sdt" or "contentcontrol"
                or "chart"
                or "comment"
                or "footnote" or "endnote"
                or "field"
                or "table" or "tbl"
                or "revision" or "change" or "trackchange";
        if (!isKnownType && parsed.ChildSelector == null)
        {
            var root = _doc.MainDocumentPart?.Document;
            if (root != null)
                return GenericXmlQuery.Query(root, genericParsed.element, genericParsed.attrs, genericParsed.containsText);
            return results;
        }

        // Handle chart query (both standard and extended chart types)
        bool isChartSelector = parsed.ChildSelector == null && parsed.Element == "chart";
        if (isChartSelector)
        {
            var allCharts = GetAllWordCharts();
            for (int i = 0; i < allCharts.Count; i++)
            {
                var chartInfo = allCharts[i];
                var node = new DocumentNode { Path = $"/chart[{i + 1}]", Type = "chart" };

                if (chartInfo.IsExtended)
                {
                    var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace;
                    var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                    if (cxType != null) node.Format["chartType"] = cxType;
                    // Title
                    var cxTitle = cxChartSpace?.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
                    var cxTitleText = cxTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
                    if (cxTitleText != null) node.Format["title"] = cxTitleText;
                    // Count series
                    var cxSeries = cxChartSpace!.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Series>().ToList();
                    node.Format["seriesCount"] = cxSeries.Count;
                }
                else
                {
                    var chart = chartInfo.StandardPart!.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                    if (chart != null)
                        Core.ChartHelper.ReadChartProperties(chart, node, 0);
                }

                if (parsed.ContainsText != null)
                {
                    var title = node.Format.TryGetValue("title", out var t) ? t?.ToString() : null;
                    if (title == null || !title.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                results.Add(node);
            }
            return results;
        }

        // Handle comment query
        bool isCommentSelector = parsed.ChildSelector == null && parsed.Element == "comment";
        if (isCommentSelector)
        {
            var commentsPart = _doc.MainDocumentPart?.WordprocessingCommentsPart;
            if (commentsPart?.Comments != null)
            {
                int cIdx = 0;
                foreach (var comment in commentsPart.Comments.Elements<Comment>())
                {
                    cIdx++;
                    var text = string.Join("", comment.Descendants<Text>().Select(t => t.Text));
                    if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var cNode = new DocumentNode
                    {
                        Path = $"/comments/comment[{cIdx}]",
                        Type = "comment",
                        Text = text
                    };
                    if (comment.Author?.Value != null) cNode.Format["author"] = comment.Author.Value;
                    if (comment.Initials?.Value != null) cNode.Format["initials"] = comment.Initials.Value;
                    if (comment.Id?.Value != null) cNode.Format["id"] = comment.Id.Value;
                    if (comment.Date?.Value != null) cNode.Format["date"] = comment.Date.Value.ToString("o");
                    results.Add(cNode);
                }
            }
            return results;
        }

        // Handle footnote query
        bool isFootnoteSelector = parsed.ChildSelector == null && parsed.Element == "footnote";
        if (isFootnoteSelector)
        {
            var footnotesPart = _doc.MainDocumentPart?.FootnotesPart;
            if (footnotesPart?.Footnotes != null)
            {
                int fnIdx = 0;
                foreach (var fn in footnotesPart.Footnotes.Elements<Footnote>())
                {
                    // Skip separator/continuation footnotes (type != null means special)
                    if (fn.Type?.Value != null) continue;
                    fnIdx++;
                    var text = string.Join("", fn.Descendants<Text>().Select(t => t.Text));
                    if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var fnNode = new DocumentNode
                    {
                        Path = $"/footnote[{fn.Id?.Value ?? fnIdx}]",
                        Type = "footnote",
                        Text = text
                    };
                    if (fn.Id?.Value != null) fnNode.Format["id"] = fn.Id.Value.ToString();
                    results.Add(fnNode);
                }
            }
            return results;
        }

        // Handle endnote query
        bool isEndnoteSelector = parsed.ChildSelector == null && parsed.Element == "endnote";
        if (isEndnoteSelector)
        {
            var endnotesPart = _doc.MainDocumentPart?.EndnotesPart;
            if (endnotesPart?.Endnotes != null)
            {
                int enIdx = 0;
                foreach (var en in endnotesPart.Endnotes.Elements<Endnote>())
                {
                    // Skip separator/continuation endnotes (type != null means special)
                    if (en.Type?.Value != null) continue;
                    enIdx++;
                    var text = string.Join("", en.Descendants<Text>().Select(t => t.Text));
                    if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var enNode = new DocumentNode
                    {
                        Path = $"/endnote[{en.Id?.Value ?? enIdx}]",
                        Type = "endnote",
                        Text = text
                    };
                    if (en.Id?.Value != null) enNode.Format["id"] = en.Id.Value.ToString();
                    results.Add(enNode);
                }
            }
            return results;
        }

        // Handle revision / track changes query
        bool isRevisionSelector = parsed.ChildSelector == null &&
            (parsed.Element is "revision" or "change" or "trackchange");
        if (isRevisionSelector)
        {
            int revIdx = 0;
            // w:ins (InsertedRun)
            foreach (var ins in body.Descendants<InsertedRun>())
            {
                revIdx++;
                var text = string.Join("", ins.Descendants<Text>().Select(t => t.Text));
                if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                { revIdx--; continue; }
                var node = new DocumentNode
                {
                    Path = $"/revision[{revIdx}]",
                    Type = "revision",
                    Text = text
                };
                node.Format["revisionType"] = "insertion";
                if (ins.Author?.Value != null) node.Format["author"] = ins.Author.Value;
                if (ins.Date?.Value != null) node.Format["date"] = ins.Date.Value.ToString("o");
                results.Add(node);
            }
            // w:del (DeletedRun)
            foreach (var del in body.Descendants<DeletedRun>())
            {
                revIdx++;
                var text = string.Join("", del.Descendants<DeletedText>().Select(t => t.Text));
                if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                { revIdx--; continue; }
                var node = new DocumentNode
                {
                    Path = $"/revision[{revIdx}]",
                    Type = "revision",
                    Text = text
                };
                node.Format["revisionType"] = "deletion";
                if (del.Author?.Value != null) node.Format["author"] = del.Author.Value;
                if (del.Date?.Value != null) node.Format["date"] = del.Date.Value.ToString("o");
                results.Add(node);
            }
            // w:rPrChange (RunPropertiesChange)
            foreach (var rPrChange in body.Descendants<RunPropertiesChange>())
            {
                revIdx++;
                // Get text from parent run
                var parentRun = rPrChange.Ancestors<Run>().FirstOrDefault();
                var text = parentRun != null ? string.Join("", parentRun.Descendants<Text>().Select(t => t.Text)) : "";
                if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                { revIdx--; continue; }
                var node = new DocumentNode
                {
                    Path = $"/revision[{revIdx}]",
                    Type = "revision",
                    Text = text
                };
                node.Format["revisionType"] = "formatChange";
                if (rPrChange.Author?.Value != null) node.Format["author"] = rPrChange.Author.Value;
                if (rPrChange.Date?.Value != null) node.Format["date"] = rPrChange.Date.Value.ToString("o");
                results.Add(node);
            }
            // w:pPrChange (ParagraphPropertiesChange)
            foreach (var pPrChange in body.Descendants<ParagraphPropertiesChange>())
            {
                revIdx++;
                var parentPara = pPrChange.Ancestors<Paragraph>().FirstOrDefault();
                var text = parentPara != null ? string.Join("", parentPara.Descendants<Text>().Select(t => t.Text)) : "";
                if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                { revIdx--; continue; }
                var node = new DocumentNode
                {
                    Path = $"/revision[{revIdx}]",
                    Type = "revision",
                    Text = text
                };
                node.Format["revisionType"] = "paragraphChange";
                if (pPrChange.Author?.Value != null) node.Format["author"] = pPrChange.Author.Value;
                if (pPrChange.Date?.Value != null) node.Format["date"] = pPrChange.Date.Value.ToString("o");
                results.Add(node);
            }
            return results;
        }

        // Handle bookmark query
        if (isBookmarkSelector)
        {
            foreach (var bkStart in body.Descendants<BookmarkStart>())
            {
                var bkName = bkStart.Name?.Value ?? "";
                if (bkName.StartsWith("_")) continue;

                if (parsed.ContainsText != null)
                {
                    var bkText = GetBookmarkText(bkStart);
                    if (!bkText.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                results.Add(ElementToNode(bkStart, $"/bookmark[{bkName}]", 0));
            }
            return results;
        }

        if (isSdtSelector)
        {
            int blockSdtIdx = 0;
            foreach (var sdt in body.Descendants().Where(e => e is SdtBlock or SdtRun))
            {
                string path;
                if (sdt is SdtBlock)
                {
                    blockSdtIdx++;
                    path = $"/body/sdt[{blockSdtIdx}]";
                }
                else if (sdt is SdtRun sdtRun)
                {
                    // Inline SDT: compute path via parent paragraph
                    var parentPara = sdtRun.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
                    if (parentPara != null)
                    {
                        int pIdx = 1;
                        foreach (var el in body.ChildElements)
                        {
                            if (el == parentPara) break;
                            if (el is DocumentFormat.OpenXml.Wordprocessing.Paragraph) pIdx++;
                        }
                        int sdtInParaIdx = 1;
                        foreach (var child in parentPara.ChildElements)
                        {
                            if (child == sdtRun) break;
                            if (child is SdtRun) sdtInParaIdx++;
                        }
                        path = $"/body/p[{pIdx}]/sdt[{sdtInParaIdx}]";
                    }
                    else
                    {
                        blockSdtIdx++;
                        path = $"/body/sdt[{blockSdtIdx}]";
                    }
                }
                else
                {
                    blockSdtIdx++;
                    path = $"/body/sdt[{blockSdtIdx}]";
                }
                var node = ElementToNode(sdt, path, 0);
                if (parsed.ContainsText != null && !(node.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                    continue;
                results.Add(node);
            }
            return results;
        }

        int paraIdx = -1;
        int mathParaIdx = -1;
        foreach (var element in body.ChildElements)
        {
            // Display equations (m:oMathPara) at body level
            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                mathParaIdx++;
                if (isEquationSelector)
                {
                    var latex = FormulaParser.ToLatex(element);
                    if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                    {
                        results.Add(new DocumentNode
                        {
                            Path = $"/body/oMathPara[{mathParaIdx + 1}]",
                            Type = "equation",
                            Text = latex,
                            Format = { ["mode"] = "display" }
                        });
                    }
                }
                continue;
            }

            if (element is DocumentFormat.OpenXml.Wordprocessing.Table tbl)
            {
                bool isTableSelector = parsed.ChildSelector == null &&
                    (parsed.Element is "table" or "tbl");
                if (isTableSelector)
                {
                    var tblIdx = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>()
                        .TakeWhile(t => t != tbl).Count();
                    var node = ElementToNode(tbl, $"/body/tbl[{tblIdx + 1}]", 0);
                    if (parsed.ContainsText != null)
                    {
                        var tblText = string.Concat(tbl.Descendants<Text>().Select(t => t.Text));
                        if (!tblText.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(node);
                }
                continue;
            }

            if (element is Paragraph para)
            {
                paraIdx++;

                if (isEquationSelector)
                {
                    // Check for display equation (oMathPara inside w:p)
                    var oMathParaInPara = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                    if (oMathParaInPara != null)
                    {
                        mathParaIdx++;
                        var latex = FormulaParser.ToLatex(oMathParaInPara);
                        if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/body/oMathPara[{mathParaIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "display" }
                            });
                        }
                        continue;
                    }

                    // Find inline math in this paragraph
                    int mathIdx = 0;
                    foreach (var oMath in para.ChildElements.Where(e => e.LocalName == "oMath" || e is M.OfficeMath))
                    {
                        var latex = FormulaParser.ToLatex(oMath);
                        if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/body/p[{paraIdx + 1}]/oMath[{mathIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "inline" }
                            });
                        }
                        mathIdx++;
                    }
                }
                else if (isPictureSelector)
                {
                    int runIdx = 0;
                    foreach (var run in GetAllRuns(para))
                    {
                        var drawing = run.GetFirstChild<Drawing>();
                        if (drawing != null)
                        {
                            bool noAlt = parsed.Attributes.ContainsKey("__no-alt");
                            if (noAlt)
                            {
                                var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
                                if (string.IsNullOrEmpty(docProps?.Description?.Value))
                                    results.Add(CreateImageNode(drawing, run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]"));
                            }
                            else
                            {
                                results.Add(CreateImageNode(drawing, run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]"));
                            }
                        }
                        runIdx++;
                    }
                }
                else if (isRunSelector)
                {
                    // Main selector targets runs: search all runs in all paragraphs
                    int runIdx = 0;
                    foreach (var run in GetAllRuns(para))
                    {
                        if (MatchesRunSelector(run, para, parsed))
                        {
                            results.Add(ElementToNode(run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]", 0));
                        }
                        runIdx++;
                    }
                }
                else
                {
                    if (MatchesSelector(para, parsed, paraIdx))
                    {
                        results.Add(ElementToNode(para, $"/body/p[{paraIdx + 1}]", 0));
                    }

                    if (parsed.ChildSelector != null)
                    {
                        int runIdx = 0;
                        foreach (var run in GetAllRuns(para))
                        {
                            if (MatchesRunSelector(run, para, parsed.ChildSelector))
                            {
                                results.Add(ElementToNode(run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]", 0));
                            }
                            runIdx++;
                        }
                    }
                }
            }
        }

        return results;
    }
}
