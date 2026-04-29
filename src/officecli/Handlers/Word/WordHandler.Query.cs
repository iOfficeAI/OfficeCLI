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
    // ==================== Binary Extraction ====================
    //
    // Support for `officecli get --save <dest>` on nodes that have a
    // backing binary part (picture, ole object, media). We re-call Get()
    // to obtain the node's relId, then look up the part on the right
    // host part (MainDocumentPart for body content, HeaderPart/FooterPart
    // for header/footer content — since rel ids are locally-scoped per
    // OpenXmlPart, OLE relationships for header-embedded objects live on
    // the HeaderPart itself, not on MainDocumentPart).
    //
    // BUG-R11-01: Previously this unconditionally resolved against
    // MainDocumentPart, which caused `get --save` to fail for OLE in
    // /header[N]/... or /footer[N]/..., mirroring the round 5/10
    // CreateOleNode regression. Match round 10's CreateOleNode refactor:
    // iterate candidate hosts (main → headers → footers) and pick the
    // one whose GetPartById(relId) succeeds. Rel ids are locally-scoped,
    // so at most one host matches.
    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        var node = Get(path, 0);
        if (node == null) return false;
        if (!node.Format.TryGetValue("relId", out var relObj) || relObj is not string relId
            || string.IsNullOrEmpty(relId))
            return false;

        var main = _doc.MainDocumentPart;
        if (main == null) return false;

        DocumentFormat.OpenXml.Packaging.OpenXmlPart? part = null;

        // Enumerate candidate host parts in the order they most commonly
        // hold the target: MainDocumentPart first (body pictures/OLEs),
        // then header parts, then footer parts. Stop at the first match.
        var candidates = new List<DocumentFormat.OpenXml.Packaging.OpenXmlPart> { main };
        candidates.AddRange(main.HeaderParts);
        candidates.AddRange(main.FooterParts);

        foreach (var host in candidates)
        {
            try
            {
                var candidate = host.GetPartById(relId);
                if (candidate != null)
                {
                    part = candidate;
                    break;
                }
            }
            catch
            {
                // rel id not in this host — try the next
            }
        }

        if (part == null) return false;

        // BUG-R10-04: create the destination directory if missing so
        // `get --save ./outdir/file.bin` works when outdir doesn't exist.
        var destDir = Path.GetDirectoryName(destPath);
        if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            Directory.CreateDirectory(destDir);

        // CONSISTENCY(ole-cfb-wrap): unwrap CFB Ole10Native payload on read.
        byte[] rawBytes;
        using (var src = part.GetStream())
        using (var ms = new MemoryStream())
        {
            src.CopyTo(ms);
            rawBytes = ms.ToArray();
        }
        var payload = OfficeCli.Core.OleHelper.UnwrapOle10NativeIfCfb(rawBytes);
        File.WriteAllBytes(destPath, payload);
        byteCount = payload.Length;
        contentType = part.ContentType;
        return true;
    }

    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Path cannot be empty.");
        if (path == "/")
            return GetRootNode(depth);

        // Handle /body/ole[N] and friends — Word does not expose OLE as a
        // native child of body (it lives inside a run), so NavigateToElement
        // would bottom out in the generic "No ole found at /body" error.
        // Intercept here and emit the consistent cross-handler message.
        // CONSISTENCY(ole-invalid-index): match PPT/Excel phrasing exactly.
        //
        // BUG-R11-03: root-level `/ole[N]` shorthand is aliased to
        // `/body/ole[N]`. This mirrors the `/` → `/body` aliasing applied
        // by many other Word commands: users already think of the body
        // as the root, so OLE at the root should resolve there instead of
        // producing "Path not found: /ole[99]".
        var wordOleMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"^(?<parent>/body|/header\[\d+\]|/footer\[\d+\])?/(?:ole|oleobject|object|embed)\[(?<idx>\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wordOleMatch.Success)
        {
            var wOleIdx = int.Parse(wordOleMatch.Groups["idx"].Value);
            var wOleParent = wordOleMatch.Groups["parent"].Success && wordOleMatch.Groups["parent"].Value.Length > 0
                ? wordOleMatch.Groups["parent"].Value
                : "/body";
            var allOles = Query("ole").Where(n => n.Path.StartsWith(wOleParent + "/", StringComparison.OrdinalIgnoreCase)).ToList();
            if (wOleIdx < 1 || wOleIdx > allOles.Count)
                throw new ArgumentException(
                    $"OLE object {wOleIdx} not found at {wOleParent} (available: {allOles.Count}).");
            return allOles[wOleIdx - 1];
        }

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

                    // Extract opacity — normalize to canonical decimal (e.g. ".5" → "0.5")
                    var opacityMatch = System.Text.RegularExpressions.Regex.Match(xml, @"opacity=""([^""]*)""");
                    if (opacityMatch.Success)
                    {
                        var rawOpacity = opacityMatch.Groups[1].Value;
                        node.Format["opacity"] = double.TryParse(rawOpacity, System.Globalization.CultureInfo.InvariantCulture, out var opVal)
                            ? opVal.ToString(System.Globalization.CultureInfo.InvariantCulture)
                            : rawOpacity;
                    }

                    // Extract font
                    var fontMatch = System.Text.RegularExpressions.Regex.Match(xml, @"font-family:&quot;([^&]*)&quot;");
                    if (fontMatch.Success) node.Format["font"] = fontMatch.Groups[1].Value;

                    // Extract rotation — allow negative / decimal values, and tolerate
                    // intra-style whitespace ("rotation : 315").
                    var rotMatch = System.Text.RegularExpressions.Regex.Match(xml, @"rotation\s*:\s*(-?\d+(?:\.\d+)?)");
                    if (rotMatch.Success) node.Format["rotation"] = rotMatch.Groups[1].Value;

                    // BUG-R36-B3: surface size/width/height so callers can read them back.
                    var sizeMatch = System.Text.RegularExpressions.Regex.Match(xml, @"font-size\s*:\s*([^;""]+)");
                    if (sizeMatch.Success) node.Format["size"] = sizeMatch.Groups[1].Value.Trim();
                    var widthMatch = System.Text.RegularExpressions.Regex.Match(xml, @"(?<![-\w])width\s*:\s*([^;""]+)");
                    if (widthMatch.Success) node.Format["width"] = widthMatch.Groups[1].Value.Trim();
                    var heightMatch = System.Text.RegularExpressions.Regex.Match(xml, @"(?<![-\w])height\s*:\s*([^;""]+)");
                    if (heightMatch.Success) node.Format["height"] = heightMatch.Groups[1].Value.Trim();

                    return node;
                }
            }
            return node;
        }

        // FormField paths: /formfield[N] or /formfield[name]
        // Routed BEFORE ParsePath because the generic predicate validator
        // only accepts positive-integer / last() / [@attr=v] predicates and
        // would reject the documented /formfield[name] form.
        var ffMatchEarly = System.Text.RegularExpressions.Regex.Match(path, @"^/formfield\[(\w+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (ffMatchEarly.Success)
        {
            var allFormFields = FindFormFields();
            var indexOrName = ffMatchEarly.Groups[1].Value;
            if (int.TryParse(indexOrName, out var ffIdx))
            {
                if (ffIdx < 1 || ffIdx > allFormFields.Count)
                    return new DocumentNode { Path = path, Type = "error", Text = $"FormField {ffIdx} not found (total: {allFormFields.Count})" };
                return FormFieldToNode(allFormFields[ffIdx - 1], path);
            }
            else
            {
                var match = allFormFields.FirstOrDefault(ff =>
                    ff.FfData.GetFirstChild<FormFieldName>()?.Val?.Value == indexOrName);
                if (match.Field == null)
                    return new DocumentNode { Path = path, Type = "error", Text = $"FormField '{indexOrName}' not found" };
                var idx = allFormFields.IndexOf(match) + 1;
                return FormFieldToNode(match, $"/formfield[{idx}]");
            }
        }

        // Numbering paths: /numbering/num[@id=N], /numbering/abstractNum[@id=N],
        // /numbering/abstractNum[@id=N]/level[L]. Routed BEFORE ParsePath because
        // these use [@id=...] / [N starting at 0] predicates ParsePath rejects.
        //
        // Positional aliases /numbering/abstractNum[N] and /numbering/num[N]
        // translate to the canonical [@id=K] form of the Nth element. Without
        // this translation, the positional path falls through to generic
        // ParsePath and emits a node with raw OOXML field names (abstractNumId,
        // multiLevelType, lvl[N]) instead of the canonical keys (id, type,
        // level[L]) returned by [@id=K] — same data, two vocabularies.
        var numPosMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/(abstractNum|num)\[(\d+)\](.*)$");
        if (numPosMatch.Success)
        {
            var kind = numPosMatch.Groups[1].Value;
            var posIdx = int.Parse(numPosMatch.Groups[2].Value); // 1-based
            var rest = numPosMatch.Groups[3].Value;
            var nb = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            int? canonId = null;
            if (kind == "abstractNum")
            {
                var abs = nb?.Elements<AbstractNum>().ElementAtOrDefault(posIdx - 1);
                canonId = abs?.AbstractNumberId?.Value;
            }
            else
            {
                var inst = nb?.Elements<NumberingInstance>().ElementAtOrDefault(posIdx - 1);
                canonId = inst?.NumberID?.Value;
            }
            if (canonId != null)
            {
                // Re-enter Get with the canonical [@id=K] form so the rest of
                // this method's branches (level[L], format keys) all hit.
                return Get($"/numbering/{kind}[@id={canonId}]{rest}", depth);
            }
        }

        var numMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/num\[@id=(\d+)\]$");
        if (numMatch.Success)
        {
            var nid = int.Parse(numMatch.Groups[1].Value);
            var nb = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            var inst = nb?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == nid);
            if (inst == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"num with id={nid} not found" };
            var nNode = new DocumentNode { Path = path, Type = "num" };
            if (inst.AbstractNumId?.Val?.Value != null)
                nNode.Format["abstractNumId"] = inst.AbstractNumId.Val.Value.ToString();
            foreach (var ovr in inst.Elements<LevelOverride>())
            {
                var lvl = ovr.LevelIndex?.Value;
                var startV = ovr.StartOverrideNumberingValue?.Val?.Value;
                if (lvl != null && startV != null)
                    nNode.Format[$"startOverride.{lvl}"] = startV.ToString()!;
            }
            return nNode;
        }

        // Accept three child-path forms for a level:
        //   /level[L]            (positional 1-based, legacy)
        //   /lvl[@ilvl=L]        (canonical OOXML attribute)
        //   /lvl[L]              (positional 1-based on the lvl alias)
        // All translate to the same lvl element (matched by LevelIndex.Value).
        var absMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/abstractNum\[@id=(\d+)\](?:/(?:level|lvl)\[(?:@ilvl=)?(\d+)\])?$");
        if (absMatch.Success)
        {
            var aid = int.Parse(absMatch.Groups[1].Value);
            var nb = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            var abs = nb?.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == aid);
            if (abs == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"abstractNum with id={aid} not found" };
            if (absMatch.Groups[2].Success)
            {
                int lvlIdx = int.Parse(absMatch.Groups[2].Value);
                var lvl = abs.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == lvlIdx);
                if (lvl == null)
                    return new DocumentNode { Path = path, Type = "error", Text = $"level[{lvlIdx}] not found in abstractNum {aid}" };
                var lNode = new DocumentNode { Path = path, Type = "level" };
                lNode.Format["ilvl"] = lvlIdx.ToString();
                if (lvl.StartNumberingValue?.Val?.Value != null) lNode.Format["start"] = lvl.StartNumberingValue.Val.Value.ToString()!;
                if (lvl.NumberingFormat?.Val?.HasValue == true) lNode.Format["format"] = lvl.NumberingFormat.Val.InnerText;
                if (lvl.LevelText?.Val?.Value != null)
                {
                    // CONSISTENCY(canonical-keys): only emit canonical "lvlText";
                    // legacy "text" alias dropped from Get output to honor root
                    // CLAUDE.md "Canonical DocumentNode.Format Rules". Set still
                    // accepts both keys via case "text" or "lvltext".
                    lNode.Format["lvlText"] = lvl.LevelText.Val.Value;
                }
                if (lvl.LevelJustification?.Val?.HasValue == true) lNode.Format["justification"] = lvl.LevelJustification.Val.InnerText;
                if (lvl.LevelSuffix?.Val?.HasValue == true) lNode.Format["suff"] = lvl.LevelSuffix.Val.InnerText;
                var lvlR = lvl.GetFirstChild<LevelRestart>();
                if (lvlR?.Val?.Value != null) lNode.Format["lvlRestart"] = lvlR.Val.Value.ToString()!;
                if (lvl.GetFirstChild<IsLegalNumberingStyle>() != null) lNode.Format["isLgl"] = true;
                var ind = lvl.PreviousParagraphProperties?.Indentation;
                if (ind?.Left?.Value != null) lNode.Format["indent"] = ind.Left.Value;
                if (ind?.Hanging?.Value != null) lNode.Format["hanging"] = ind.Hanging.Value;
                var rpr = lvl.NumberingSymbolRunProperties;
                if (rpr != null)
                {
                    var rfn = rpr.GetFirstChild<RunFonts>();
                    if (rfn?.Ascii?.Value != null) lNode.Format["font"] = rfn.Ascii.Value;
                    var fsz = rpr.GetFirstChild<FontSize>();
                    if (fsz?.Val?.Value != null) lNode.Format["size"] = $"{int.Parse(fsz.Val.Value) / 2.0:0.##}pt";
                    var clr = rpr.GetFirstChild<Color>();
                    if (clr?.ThemeColor?.HasValue == true)
                        lNode.Format["color"] = clr.ThemeColor.InnerText;
                    else if (clr?.Val?.Value != null)
                        lNode.Format["color"] = ParseHelpers.FormatHexColor(clr.Val.Value);
                    if (rpr.GetFirstChild<Bold>() != null) lNode.Format["bold"] = true;
                    if (rpr.GetFirstChild<Italic>() != null) lNode.Format["italic"] = true;
                }
                return lNode;
            }
            else
            {
                var aNode = new DocumentNode { Path = path, Type = "abstractNum" };
                aNode.Format["id"] = aid.ToString();
                var mlt = abs.GetFirstChild<MultiLevelType>();
                if (mlt?.Val?.HasValue == true) aNode.Format["type"] = mlt.Val.InnerText;
                var nm = abs.GetFirstChild<AbstractNumDefinitionName>();
                if (nm?.Val?.Value != null) aNode.Format["name"] = nm.Val.Value;
                var sl = abs.GetFirstChild<StyleLink>();
                if (sl?.Val?.Value != null) aNode.Format["styleLink"] = sl.Val.Value;
                var nsl = abs.GetFirstChild<NumberingStyleLink>();
                if (nsl?.Val?.Value != null) aNode.Format["numStyleLink"] = nsl.Val.Value;
                foreach (var lvl in abs.Elements<Level>())
                {
                    var li = lvl.LevelIndex?.Value;
                    if (li.HasValue)
                        aNode.Children.Add(new DocumentNode { Path = $"{path}/level[{li}]", Type = "level" });
                }
                return aNode;
            }
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

        // Footnote/Endnote paths: /footnote[N], /footnote[@footnoteId=N], /endnote[N], /endnote[@endnoteId=N]
        var fnMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/footnote\[(?:@footnoteId=)?(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (fnMatch.Success)
        {
            var fnId = int.Parse(fnMatch.Groups[1].Value);
            var fn = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
            if (fn == null)
                throw new ArgumentException($"Footnote {fnId} not found");
            var fnNode = new DocumentNode { Path = $"/footnote[@footnoteId={fnId}]", Type = "footnote" };
            fnNode.Text = GetFootnoteText(fn);
            if (fn.Id?.Value != null) fnNode.Format["id"] = fn.Id.Value;
            return fnNode;
        }
        var enMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/endnote\[(?:@endnoteId=)?(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (enMatch.Success)
        {
            var enId = int.Parse(enMatch.Groups[1].Value);
            var en = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
            if (en == null)
                throw new ArgumentException($"Endnote {enId} not found");
            var enNode = new DocumentNode { Path = $"/endnote[@endnoteId={enId}]", Type = "endnote" };
            enNode.Text = GetFootnoteText(en);
            if (en.Id?.Value != null) enNode.Format["id"] = en.Id.Value;
            return enNode;
        }

        // TOC paths: /toc[N], /toc (= first), /tableofcontents (long alias).
        // CONSISTENCY(toc-aliases): the type alias `tableofcontents` is already
        // accepted by Add (WordHandler.Add.cs) and the help text documents
        // both `/toc` and `/tableofcontents` — Get must mirror them.
        var tocMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/(?:toc|tableofcontents)(?:\[(\d+)\])?$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (tocMatch.Success)
        {
            var tocIdx = tocMatch.Groups[1].Success ? int.Parse(tocMatch.Groups[1].Value) : 1;
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

            // BUG-R11-05: recover the `title=` supplied to `add toc` — it is
            // stored as a preceding paragraph styled `TOCHeading`, not on the
            // TOC field itself. Read the previous sibling, and if it carries
            // that style, surface its text as `Format["title"]` so that
            // Add→Get round-trips the title prop.
            var prevPara = tocPara.PreviousSibling<Paragraph>();
            if (prevPara != null)
            {
                var prevStyle = prevPara.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (prevStyle == "TOCHeading")
                {
                    var titleText = string.Concat(prevPara.Descendants<Text>().Select(t => t.Text));
                    if (!string.IsNullOrEmpty(titleText))
                        tocNode.Format["title"] = titleText;
                }
            }
            return tocNode;
        }

        // Field paths: /field[N]
        var fieldMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/field\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (fieldMatch.Success)
        {
            var fieldIdx = int.Parse(fieldMatch.Groups[1].Value);
            var allFields = FindFields();
            if (fieldIdx < 1 || fieldIdx > allFields.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"Field {fieldIdx} not found (total: {allFields.Count})" };
            return FieldToNode(allFields[fieldIdx - 1], path);
        }

        // Chart axis-by-role sub-path: /chart[N]/axis[@role=ROLE].
        // Per schemas/help/pptx/chart-axis.json (shared contract across Pptx/Word/Excel).
        var chartAxisGetMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (chartAxisGetMatch.Success)
        {
            var caChartIdx = int.Parse(chartAxisGetMatch.Groups[1].Value);
            var caRole = chartAxisGetMatch.Groups[2].Value;
            var caAllCharts = GetAllWordCharts();
            if (caChartIdx < 1 || caChartIdx > caAllCharts.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"Chart {caChartIdx} not found" };
            var caChartInfo = caAllCharts[caChartIdx - 1];
            if (caChartInfo.IsExtended || caChartInfo.StandardPart?.ChartSpace == null)
                throw new ArgumentException($"Axis not available on chart {caChartIdx}: extended charts not supported.");
            var axisNode = Core.ChartHelper.BuildAxisNode(caChartInfo.StandardPart.ChartSpace, caRole, path);
            if (axisNode == null)
                throw new ArgumentException($"Axis with role '{caRole}' not found on chart {caChartIdx}.");
            return axisNode;
        }

        // Chart paths: /chart[N] or /chart[N]/series[K]
        var chartGetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/chart\[(\d+)\](?:/series\[(\d+)\])?$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (chartGetMatch.Success)
        {
            var chartIdx = int.Parse(chartGetMatch.Groups[1].Value);
            var allCharts = GetAllWordCharts();
            if (chartIdx < 1 || chartIdx > allCharts.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"Chart {chartIdx} not found" };

            var chartInfo = allCharts[chartIdx - 1];
            var chartNode = new DocumentNode { Path = $"/chart[{chartIdx}]", Type = "chart" };
            if (chartInfo.DocProperties?.Id?.HasValue == true)
                chartNode.Format["id"] = chartInfo.DocProperties.Id.Value;
            if (chartInfo.DocProperties?.Name?.Value != null)
                chartNode.Format["name"] = chartInfo.DocProperties.Name.Value;

            if (chartInfo.IsExtended)
            {
                // Extended chart (funnel, treemap, etc.)
                var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace!;
                var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                if (cxType != null) chartNode.Format["chartType"] = cxType;
                // Title
                var cxTitle = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
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
        // CONSISTENCY(path-element-case-insensitive): top-level element paths like
        // /section[N], /chart[N], /footnote[N], /toc[N] are matched case-insensitively
        // so /Section[1] and /section[1] are equivalent. The returned node's Path is
        // canonicalised to lowercase so callers see a round-trippable form. Style ids
        // (/styles/<id>) remain case-sensitive — they are user-defined identifiers.
        var secMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/section\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (secMatch.Success)
        {
            var secIdx = int.Parse(secMatch.Groups[1].Value);
            var sectionProps = FindSectionProperties();
            if (secIdx < 1 || secIdx > sectionProps.Count)
                throw new ArgumentException($"Section {secIdx} not found (total: {sectionProps.Count})");

            var sectPr = sectionProps[secIdx - 1];
            return BuildSectionNode(sectPr, path.ToLowerInvariant());
        }

        // /docDefaults — root-level access to docDefaults rPr/pPr. Mirrors
        // PopulateDocDefaults output but as a standalone node so the
        // effective.X.src = "/docDefaults" pointer (run/paragraph
        // provenance) is directly Get-able without retrieving the whole
        // document root node.
        if (path == "/docDefaults")
        {
            var ddNode = new DocumentNode { Path = path, Type = "docDefaults" };
            PopulateDocDefaults(ddNode);
            return ddNode;
        }

        // Style paths: /styles/StyleId (read the style itself).
        // Restrict to a single segment so deeper paths like /styles/<id>/tab[N]
        // fall through to generic Navigation.
        var styleMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/styles/([^/]+)$");
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
            if (style.BasedOn?.Val?.Value != null)
            {
                styleNode.Format["basedOn"] = style.BasedOn.Val.Value;
                styleNode.Format["basedOn.path"] = $"/styles/{style.BasedOn.Val.Value}";
            }
            if (style.NextParagraphStyle?.Val?.Value != null) styleNode.Format["next"] = style.NextParagraphStyle.Val.Value;

            // Read run properties
            var rPr = style.StyleRunProperties;
            if (rPr != null)
            {
                if (rPr.RunFonts != null)
                {
                    var rf = rPr.RunFonts;
                    if (rf.Ascii?.Value != null) styleNode.Format["font.ascii"] = rf.Ascii.Value;
                    if (rf.EastAsia?.Value != null) styleNode.Format["font.eastAsia"] = rf.EastAsia.Value;
                    if (rf.HighAnsi?.Value != null) styleNode.Format["font.hAnsi"] = rf.HighAnsi.Value;
                    if (rf.ComplexScript?.Value != null) styleNode.Format["font.cs"] = rf.ComplexScript.Value;
                    // CONSISTENCY(canonical-keys): font.ascii is canonical; do not also emit flat "font" alias.
                }
                if (rPr.FontSize?.Val?.Value != null) styleNode.Format["size"] = $"{int.Parse(rPr.FontSize.Val.Value) / 2.0:0.##}pt";
                // Complex-script size (<w:szCs/>) — half-points like <w:sz/>.
                // Mirrors the run-node readback in WordHandler.Navigation.cs:1287
                // so a get→add round-trip on a style preserves CS sizing.
                if (rPr.GetFirstChild<FontSizeComplexScript>()?.Val?.Value is string szCsVal
                    && int.TryParse(szCsVal, out var szCsHalfPt))
                    styleNode.Format["size.cs"] = $"{szCsHalfPt / 2.0:0.##}pt";
                if (rPr.Bold != null) styleNode.Format["bold"] = true;
                if (rPr.GetFirstChild<BoldComplexScript>() != null) styleNode.Format["bold.cs"] = true;
                if (rPr.Italic != null) styleNode.Format["italic"] = true;
                if (rPr.GetFirstChild<ItalicComplexScript>() != null) styleNode.Format["italic.cs"] = true;
                if (rPr.Color?.Val?.Value != null) styleNode.Format["color"] = ParseHelpers.FormatHexColor(rPr.Color.Val.Value);
                else if (rPr.Color?.ThemeColor?.HasValue == true) styleNode.Format["color"] = rPr.Color.ThemeColor.InnerText;
                if (rPr.Underline?.Val != null) styleNode.Format["underline"] = rPr.Underline.Val.InnerText;
                // CONSISTENCY(underline-color): underline.color not yet exposed by paragraph/run Get; backfill there too.
                if (rPr.Underline?.Color?.Value != null) styleNode.Format["underline.color"] = ParseHelpers.FormatHexColor(rPr.Underline.Color.Value);
                if (rPr.Strike != null) styleNode.Format["strike"] = true;
            }

            // Read paragraph properties
            var pPr = style.StyleParagraphProperties;
            if (pPr != null)
            {
                if (pPr.Justification?.Val?.Value != null) styleNode.Format["alignment"] = pPr.Justification.Val.InnerText;
                if (pPr.SpacingBetweenLines != null)
                {
                    var sp = pPr.SpacingBetweenLines;
                    if (sp.Before?.Value != null) styleNode.Format["spaceBefore"] = SpacingConverter.FormatWordSpacing(sp.Before.Value);
                    if (sp.After?.Value != null) styleNode.Format["spaceAfter"] = SpacingConverter.FormatWordSpacing(sp.After.Value);
                    if (sp.Line?.Value != null) styleNode.Format["lineSpacing"] = SpacingConverter.FormatWordLineSpacing(sp.Line.Value, sp.LineRule?.InnerText);
                    // CONSISTENCY(line-rule): lineRule not yet exposed by paragraph Get; backfill there too.
                    if (sp.LineRule?.HasValue == true) styleNode.Format["lineRule"] = sp.LineRule.InnerText;
                    // CONSISTENCY(spacing-lines): *Lines variants not yet exposed by paragraph Get.
                    if (sp.BeforeLines?.Value != null) styleNode.Format["spaceBeforeLines"] = sp.BeforeLines.Value;
                    if (sp.AfterLines?.Value != null) styleNode.Format["spaceAfterLines"] = sp.AfterLines.Value;
                }

                if (pPr.Indentation != null)
                {
                    var ind = pPr.Indentation;
                    // Left/Right and Start/End are OOXML aliases; modern Word writes Start/End.
                    // CONSISTENCY(unit-qualified-spacing): unit-qualified output via SpacingConverter.
                    if (ind.FirstLine?.Value != null) styleNode.Format["firstLineIndent"] = SpacingConverter.FormatWordSpacing(ind.FirstLine.Value);
                    if (ind.Hanging?.Value != null) styleNode.Format["hangingIndent"] = SpacingConverter.FormatWordSpacing(ind.Hanging.Value);
                    var leftTwips = ind.Left?.Value ?? ind.Start?.Value;
                    if (leftTwips != null) styleNode.Format["leftIndent"] = SpacingConverter.FormatWordSpacing(leftTwips);
                    var rightTwips = ind.Right?.Value ?? ind.End?.Value;
                    if (rightTwips != null) styleNode.Format["rightIndent"] = SpacingConverter.FormatWordSpacing(rightTwips);
                    // CONSISTENCY(ind-chars): *Chars variants not yet exposed by paragraph Get.
                    if (ind.FirstLineChars?.Value != null) styleNode.Format["firstLineChars"] = ind.FirstLineChars.Value;
                    if (ind.HangingChars?.Value != null) styleNode.Format["hangingChars"] = ind.HangingChars.Value;
                    var leftChars = ind.LeftChars?.Value ?? ind.StartCharacters?.Value;
                    if (leftChars != null) styleNode.Format["leftChars"] = leftChars;
                    var rightChars = ind.RightChars?.Value ?? ind.EndCharacters?.Value;
                    if (rightChars != null) styleNode.Format["rightChars"] = rightChars;
                }

                // CONSISTENCY(outline-lvl): outlineLvl not yet exposed by paragraph Get.
                if (pPr.OutlineLevel?.Val?.Value != null) styleNode.Format["outlineLvl"] = (int)pPr.OutlineLevel.Val.Value;

                // Numbering linkage on the style itself (numPr in style/pPr).
                // Mirrors paragraph-level numId/ilvl readback in Navigation.cs.
                if (pPr.NumberingProperties != null)
                {
                    var sNumPr = pPr.NumberingProperties;
                    if (sNumPr.NumberingId?.Val?.Value != null)
                        styleNode.Format["numId"] = sNumPr.NumberingId.Val.Value.ToString();
                    if (sNumPr.NumberingLevelReference?.Val?.Value != null)
                        styleNode.Format["ilvl"] = sNumPr.NumberingLevelReference.Val.Value.ToString();
                }

                // Toggle props: respect explicit val="false" instead of treating presence as true.
                if (pPr.KeepNext != null)
                {
                    var v = pPr.KeepNext.Val;
                    styleNode.Format["keepNext"] = v == null || v.Value;
                }
                if (pPr.KeepLines != null)
                {
                    var v = pPr.KeepLines.Val;
                    styleNode.Format["keepLines"] = v == null || v.Value;
                }
                if (pPr.PageBreakBefore != null)
                {
                    var v = pPr.PageBreakBefore.Val;
                    styleNode.Format["pageBreakBefore"] = v == null || v.Value;
                }
                if (pPr.WidowControl != null)
                {
                    var v = pPr.WidowControl.Val;
                    styleNode.Format["widowControl"] = v == null || v.Value;
                }
                if (pPr.ContextualSpacing != null)
                {
                    var v = pPr.ContextualSpacing.Val;
                    styleNode.Format["contextualSpacing"] = v == null || v.Value;
                }

                // CONSISTENCY(canonical-keys): split shading into shading.val/.fill/.color sub-keys.
                if (pPr.Shading != null)
                {
                    var shdVal = pPr.Shading.Val?.InnerText;
                    var shdFill = pPr.Shading.Fill?.Value;
                    var shdColor = pPr.Shading.Color?.Value;
                    if (!string.IsNullOrEmpty(shdVal)) styleNode.Format["shading.val"] = shdVal;
                    if (!string.IsNullOrEmpty(shdFill)) styleNode.Format["shading.fill"] = ParseHelpers.FormatHexColor(shdFill);
                    if (!string.IsNullOrEmpty(shdColor)) styleNode.Format["shading.color"] = ParseHelpers.FormatHexColor(shdColor);
                }

                var pBdr = pPr.ParagraphBorders;
                if (pBdr != null)
                {
                    ReadBorder(pBdr.TopBorder, "pbdr.top", styleNode);
                    ReadBorder(pBdr.BottomBorder, "pbdr.bottom", styleNode);
                    ReadBorder(pBdr.LeftBorder, "pbdr.left", styleNode);
                    ReadBorder(pBdr.RightBorder, "pbdr.right", styleNode);
                    ReadBorder(pBdr.BetweenBorder, "pbdr.between", styleNode);
                    ReadBorder(pBdr.BarBorder, "pbdr.bar", styleNode);
                }

                var numProps = pPr.NumberingProperties;
                if (numProps?.NumberingId?.Val?.Value != null)
                {
                    styleNode.Format["numId"] = numProps.NumberingId.Val.Value.ToString();
                    if (numProps.NumberingLevelReference?.Val?.Value != null)
                        styleNode.Format["numLevel"] = numProps.NumberingLevelReference.Val.Value.ToString();
                }

                // CONSISTENCY(tabs): tabs[] not yet exposed by paragraph Get.
                if (pPr.Tabs != null)
                {
                    var tabList = new List<Dictionary<string, object?>>();
                    foreach (var tab in pPr.Tabs.Elements<TabStop>())
                    {
                        var t = new Dictionary<string, object?>();
                        if (tab.Position?.Value != null) t["pos"] = tab.Position.Value;
                        if (tab.Val?.HasValue == true) t["val"] = tab.Val.InnerText;
                        if (tab.Leader?.HasValue == true) t["leader"] = tab.Leader.InnerText;
                        if (t.Count > 0) tabList.Add(t);
                    }
                    if (tabList.Count > 0) styleNode.Format["tabs"] = tabList;
                }
            }

            // Long-tail fallback: surface every rPr/pPr child element the
            // curated reader did not consume. Keys are bare OOXML localNames
            // (e.g. "kinsoku", "snapToGrid"), symmetric with the Set side's
            // GenericXmlQuery.TryCreateTypedChild — so values round-trip
            // through `get | set` without any special namespace.
            // CONSISTENCY(generic-fallback): paragraph/run Get should adopt the
            // same pattern in a future sweep so curated drift stops being a P0.
            FillUnknownChildProps(rPr, styleNode);
            FillUnknownChildProps(pPr, styleNode);
            return styleNode;
        }

        var parts = ParsePath(path);
        var element = NavigateToElement(parts, out var ctx, out var resolvedPath);
        if (element == null)
        {
            // Check if the path contains footnote/endnote/toc which are handled differently
            if (path.Contains("footnote") || path.Contains("endnote") || path.Contains("toc"))
                return new DocumentNode { Path = path, Type = "error", Text = $"Path not found: {path}" };
            var msg = $"Path not found: {path}";
            if (ctx != null) msg += $". {ctx}";
            throw new ArgumentException(msg);
        }

        // Use the resolved positional path when available (normalizes @paraId etc.)
        var nodePath = !string.IsNullOrEmpty(resolvedPath) ? resolvedPath : path;
        return ElementToNode(element, nodePath, depth);
    }

    /// <summary>Build a DocumentNode for a section from its SectionProperties element.</summary>
    private DocumentNode BuildSectionNode(SectionProperties sectPr, string path)
    {
        var secNode = new DocumentNode { Path = path, Type = "section" };

        var sectType = sectPr.GetFirstChild<SectionType>();
        if (sectType?.Val?.Value != null)
            secNode.Format["type"] = sectType.Val.InnerText;
        var pageSize = sectPr.GetFirstChild<PageSize>();
        // Default to A4 size if no explicit page size
        var pgW = pageSize?.Width?.Value ?? WordPageDefaults.A4WidthTwips;
        var pgH = pageSize?.Height?.Value ?? WordPageDefaults.A4HeightTwips;
        secNode.Format["pageWidth"] = FormatTwipsToCm(pgW);
        secNode.Format["pageHeight"] = FormatTwipsToCm(pgH);
        if (pageSize?.Orient?.Value != null) secNode.Format["orientation"] = pageSize.Orient.InnerText;
        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin?.Top?.Value != null) secNode.Format["marginTop"] = FormatTwipsToCm((uint)Math.Abs(margin.Top.Value));
        if (margin?.Bottom?.Value != null) secNode.Format["marginBottom"] = FormatTwipsToCm((uint)Math.Abs(margin.Bottom.Value));
        if (margin?.Left?.Value != null) secNode.Format["marginLeft"] = FormatTwipsToCm(margin.Left.Value);
        if (margin?.Right?.Value != null) secNode.Format["marginRight"] = FormatTwipsToCm(margin.Right.Value);

        // Page numbering start (w:pgNumType/@start) and format (w:pgNumType/@fmt)
        var pgNumType = sectPr.GetFirstChild<PageNumberType>();
        if (pgNumType?.Start?.Value != null)
            secNode.Format["pageStart"] = pgNumType.Start.Value;
        if (pgNumType?.Format?.Value != null)
            secNode.Format["pageNumFmt"] = pgNumType.Format.InnerText;

        // Title page flag (w:titlePg) — first-page header/footer differs from rest
        if (sectPr.GetFirstChild<TitlePage>() != null)
            secNode.Format["titlePage"] = true;

        // Section-level RTL (Arabic / Hebrew page direction).
        if (sectPr.GetFirstChild<BiDi>() != null)
            secNode.Format["direction"] = "rtl";

        // Header / footer references — expose so users can debug inheritance
        var mainPart = _doc.MainDocumentPart;
        if (mainPart != null)
        {
            // headerRef = primary (default or first encountered) /header[N] path;
            // headerRef.<type> = per-type entry (default/first/even) for inheritance debugging.
            string? primaryHeader = null;
            foreach (var href in sectPr.Elements<HeaderReference>())
            {
                if (href.Id?.Value == null) continue;
                var refType = href.Type?.InnerText ?? "default";
                try
                {
                    var part = mainPart.GetPartById(href.Id.Value) as DocumentFormat.OpenXml.Packaging.HeaderPart;
                    if (part != null)
                    {
                        var idx = mainPart.HeaderParts.ToList().IndexOf(part);
                        if (idx >= 0)
                        {
                            var pathRef = $"/header[{idx + 1}]";
                            secNode.Format[$"headerRef.{refType}"] = pathRef;
                            if (primaryHeader == null || refType == "default") primaryHeader = pathRef;
                        }
                    }
                }
                catch { /* dangling rel — skip */ }
            }
            if (primaryHeader != null) secNode.Format["headerRef"] = primaryHeader;

            string? primaryFooter = null;
            foreach (var fref in sectPr.Elements<FooterReference>())
            {
                if (fref.Id?.Value == null) continue;
                var refType = fref.Type?.InnerText ?? "default";
                try
                {
                    var part = mainPart.GetPartById(fref.Id.Value) as DocumentFormat.OpenXml.Packaging.FooterPart;
                    if (part != null)
                    {
                        var idx = mainPart.FooterParts.ToList().IndexOf(part);
                        if (idx >= 0)
                        {
                            var pathRef = $"/footer[{idx + 1}]";
                            secNode.Format[$"footerRef.{refType}"] = pathRef;
                            if (primaryFooter == null || refType == "default") primaryFooter = pathRef;
                        }
                    }
                }
                catch { /* dangling rel — skip */ }
            }
            if (primaryFooter != null) secNode.Format["footerRef"] = primaryFooter;
        }

        // Line numbers
        var lnNum = sectPr.GetFirstChild<LineNumberType>();
        if (lnNum != null)
        {
            var countBy = lnNum.CountBy?.Value ?? 1;
            var restartVal = lnNum.Restart?.InnerText ?? "continuous";
            var restart = restartVal switch
            {
                "newPage" => "restartPage",
                "newSection" => "restartSection",
                _ => "continuous"
            };
            secNode.Format["lineNumbers"] = restart;
            if (countBy != 1) secNode.Format["lineNumberCountBy"] = countBy;
        }

        // Column properties — dotted canonical keys mirror Set's input form
        // (columns.count / columns.space / columns.equalWidth / columns.separator)
        // and the sibling DocSettings readback in WordHandler.Navigation.DocSettings.cs.
        var cols = sectPr.GetFirstChild<Columns>();
        if (cols != null)
        {
            secNode.Format["columns.count"] = cols.ColumnCount?.Value ?? 1;
            if (cols.Space?.Value != null && uint.TryParse(cols.Space.Value, out var colSpaceTwips))
                secNode.Format["columns.space"] = FormatTwipsToCm(colSpaceTwips);
            if (cols.EqualWidth?.Value != null) secNode.Format["columns.equalWidth"] = cols.EqualWidth.Value;
            if (cols.Separator?.Value == true) secNode.Format["columns.separator"] = true;
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
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
        }

        var firstPara = header.Elements<Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Justification?.Val != null)
            node.Format["alignment"] = firstPara.ParagraphProperties.Justification.Val.InnerText;

        node.ChildCount = header.Elements<Paragraph>().Count() + header.Elements<Table>().Count();
        // CONSISTENCY(header-footer-get): default depth (=1) returns the
        // single header/footer node, mirroring `query header` / `query footer`.
        // Block children (paragraphs + tables) only expand at explicit depth >= 2.
        if (depth >= 1)
        {
            int pIdx = 0, tblIdx = 0;
            foreach (var child in header.ChildElements)
            {
                if (child is Paragraph para)
                {
                    pIdx++;
                    var paraSegment = BuildParaPathSegment(para, pIdx);
                    node.Children.Add(ElementToNode(para, $"{path}/{paraSegment}", depth - 1));
                }
                else if (child is Table)
                {
                    tblIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/tbl[{tblIdx}]", depth - 1));
                }
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
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
        }

        var firstPara = footer.Elements<Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Justification?.Val != null)
            node.Format["alignment"] = firstPara.ParagraphProperties.Justification.Val.InnerText;

        node.ChildCount = footer.Elements<Paragraph>().Count() + footer.Elements<Table>().Count();
        // CONSISTENCY(header-footer-get): see GetHeaderNode.
        if (depth >= 1)
        {
            int pIdx = 0, tblIdx = 0;
            foreach (var child in footer.ChildElements)
            {
                if (child is Paragraph para)
                {
                    pIdx++;
                    var paraSegment = BuildParaPathSegment(para, pIdx);
                    node.Children.Add(ElementToNode(para, $"{path}/{paraSegment}", depth - 1));
                }
                else if (child is Table)
                {
                    tblIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/tbl[{tblIdx}]", depth - 1));
                }
            }
        }

        return node;
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return results;

        // BUG-R18-01: scoped OLE selector `/body/ole`, `/header[N]/ole`,
        // `/footer[N]/ole` (and `object`/`embed` aliases) was not recognized
        // by ParseSingleSelector — it truncated at the first `[`, so the
        // element became `/header` and never matched the OLE branch.
        // Intercept here and delegate to the general `ole` query, filtering
        // results whose Path starts with the requested parent scope.
        // CONSISTENCY(word-ole-scope): mirrors the scoped `Get` path at
        // WordHandler.Query.cs line ~108 (wordOleMatch).
        var wordOleScopeMatch = System.Text.RegularExpressions.Regex.Match(
            selector,
            // BUG-R38-01: attr filter suffix `[...]` was not captured, so
            // `/body/ole[fileSize>0]` fell through to ParseSelector and matched 0.
            // CONSISTENCY(word-ole-scope): delegate attr filter to Query("ole[...]")
            // exactly as the unscoped branch does.
            @"^(?<parent>/body|/header\[\d+\]|/footer\[\d+\])/(?:ole|oleobject|object|embed)(?<attrs>\[.*\])?$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wordOleScopeMatch.Success)
        {
            var scopePrefix = wordOleScopeMatch.Groups["parent"].Value;
            var attrSuffix = wordOleScopeMatch.Groups["attrs"].Value; // "" when absent
            var oleSelector = "ole" + attrSuffix;
            return Query(oleSelector)
                .Where(n => n.Path.StartsWith(scopePrefix + "/", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        // Simple selector parser: element[attr=value]
        var parsed = ParseSelector(selector);

        // Handle section selector — sections live in paragraph-level sectPr
        // and the body-level sectPr (last section). OOXML tag is "sectPr",
        // so GenericXmlQuery with element "section" never matches; route
        // explicitly here for parity with /section[N] Get.
        if (parsed.Element == "section")
        {
            var sectionProps = FindSectionProperties();
            for (int si = 0; si < sectionProps.Count; si++)
            {
                var node = BuildSectionNode(sectionProps[si], $"/section[{si + 1}]");
                if (parsed.ContainsText == null || (node.Text?.Contains(parsed.ContainsText) == true))
                    results.Add(node);
            }
            return results;
        }

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

        // Handle style selector — styles live in StylesPart, not Body
        if (parsed.Element == "style")
        {
            var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles != null)
            {
                int sIdx = 0;
                foreach (var style in styles.Elements<Style>())
                {
                    sIdx++;
                    var styleId = style.StyleId?.Value ?? "";
                    var styleName = style.StyleName?.Val?.Value ?? styleId;
                    var styleNode = new DocumentNode
                    {
                        Path = $"/styles/{styleId}",
                        Type = "style",
                        Text = styleName
                    };
                    styleNode.Format["id"] = styleId;
                    styleNode.Format["name"] = styleName;
                    if (style.Type?.Value != null) styleNode.Format["type"] = style.Type.InnerText;
                    if (style.BasedOn?.Val?.Value != null)
            {
                styleNode.Format["basedOn"] = style.BasedOn.Val.Value;
                styleNode.Format["basedOn.path"] = $"/styles/{style.BasedOn.Val.Value}";
            }

                    // Filter by :contains
                    if (parsed.ContainsText != null && !(styleName.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) == true))
                        continue;
                    // Filter by attributes
                    bool matchAttrs = true;
                    foreach (var (attrKey, rawVal) in parsed.Attributes)
                    {
                        bool negate = rawVal.StartsWith("!");
                        var val = negate ? rawVal[1..] : rawVal;
                        var hasKey = styleNode.Format.TryGetValue(attrKey, out var fmtVal);
                        bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                        if (negate ? matches : !matches) { matchAttrs = false; break; }
                    }
                    if (matchAttrs) results.Add(styleNode);
                }
            }
            return results;
        }

        // Handle watermark selector — at most one watermark per document.
        // Schema declares query=true; reuse the singleton /watermark Get logic.
        if (parsed.Element == "watermark")
        {
            if (FindWatermark() != null)
            {
                var wmNode = Get("/watermark");
                if (wmNode != null && wmNode.Type == "watermark"
                    && (parsed.ContainsText == null || (wmNode.Text?.Contains(parsed.ContainsText) == true)))
                {
                    results.Add(wmNode);
                }
            }
            return results;
        }

        // Handle /styles container selector — styles container is a singleton.
        // Schema declares query=true on the styles container. Return exactly one
        // node representing the container; individual styles remain queryable
        // via `query style`.
        if (parsed.Element == "styles")
        {
            var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles != null)
            {
                var node = new DocumentNode
                {
                    Path = "/styles",
                    Type = "styles"
                };
                node.Format["count"] = styles.Elements<Style>().Count();
                results.Add(node);
            }
            return results;
        }

        // Handle numbering container selector — singleton, mirrors `query styles`.
        // Schema also exposes `num` and `abstractNum` as queryable element types
        // that live under NumberingDefinitionsPart, not under body. Without
        // these intercepts, the generic XML fallback only walks body and
        // returns 0 results despite Get(/numbering/...) working fine.
        if (parsed.Element == "numbering")
        {
            var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering != null)
            {
                var node = new DocumentNode { Path = "/numbering", Type = "numbering" };
                node.Format["abstractNumCount"] = numbering.Elements<AbstractNum>().Count();
                node.Format["numCount"] = numbering.Elements<NumberingInstance>().Count();
                results.Add(node);
            }
            return results;
        }

        if (parsed.Element == "abstractNum" || parsed.Element == "abstractnum")
        {
            var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering != null)
            {
                foreach (var abs in numbering.Elements<AbstractNum>())
                {
                    var aid = abs.AbstractNumberId?.Value;
                    if (aid == null) continue;
                    var node = Get($"/numbering/abstractNum[@id={aid}]");
                    if (node == null || node.Type == "error") continue;
                    // Filter by attributes (e.g. abstractNum[type=hybridMultilevel])
                    bool matchAttrs = true;
                    foreach (var (attrKey, rawVal) in parsed.Attributes)
                    {
                        bool negate = rawVal.StartsWith("!");
                        var val = negate ? rawVal[1..] : rawVal;
                        var hasKey = node.Format.TryGetValue(attrKey, out var fmtVal);
                        bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                        if (negate ? matches : !matches) { matchAttrs = false; break; }
                    }
                    if (matchAttrs) results.Add(node);
                }
            }
            return results;
        }

        if (parsed.Element == "num")
        {
            var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering != null)
            {
                foreach (var inst in numbering.Elements<NumberingInstance>())
                {
                    var nid = inst.NumberID?.Value;
                    if (nid == null) continue;
                    var node = Get($"/numbering/num[@id={nid}]");
                    if (node == null || node.Type == "error") continue;
                    bool matchAttrs = true;
                    foreach (var (attrKey, rawVal) in parsed.Attributes)
                    {
                        bool negate = rawVal.StartsWith("!");
                        var val = negate ? rawVal[1..] : rawVal;
                        var hasKey = node.Format.TryGetValue(attrKey, out var fmtVal);
                        bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                        if (negate ? matches : !matches) { matchAttrs = false; break; }
                    }
                    if (matchAttrs) results.Add(node);
                }
            }
            return results;
        }

        // Handle toc selector
        if (parsed.Element is "toc" or "tableofcontents")
        {
            var tocParas = FindTocParagraphs();
            for (int ti = 0; ti < tocParas.Count; ti++)
            {
                var tocPara = tocParas[ti];
                var instrText = string.Join("", tocPara.Descendants<FieldCode>().Select(fc => fc.Text));
                var tocNode = new DocumentNode { Path = $"/toc[{ti + 1}]", Type = "toc" };
                tocNode.Text = instrText.Trim();

                var levelsMatch = System.Text.RegularExpressions.Regex.Match(instrText, @"\\o\s+""([^""]+)""");
                if (levelsMatch.Success) tocNode.Format["levels"] = levelsMatch.Groups[1].Value;
                tocNode.Format["hyperlinks"] = instrText.Contains("\\h");
                tocNode.Format["pageNumbers"] = !instrText.Contains("\\z");

                if (parsed.ContainsText != null && !(tocNode.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                    continue;
                results.Add(tocNode);
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

        // Handle formfield selector
        if (parsed.Element == "formfield")
        {
            var allFormFields = FindFormFields();
            for (int fi = 0; fi < allFormFields.Count; fi++)
            {
                var ffNode = FormFieldToNode(allFormFields[fi], $"/formfield[{fi + 1}]");
                // Filter by :contains
                if (parsed.ContainsText != null && !(ffNode.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                    continue;
                // Filter by attributes
                bool matchAttrs = true;
                foreach (var (attrKey, rawVal) in parsed.Attributes)
                {
                    bool negate = rawVal.StartsWith("!");
                    var val = negate ? rawVal[1..] : rawVal;
                    var hasKey = ffNode.Format.TryGetValue(attrKey, out var fmtVal);
                    bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                    if (negate ? matches : !matches) { matchAttrs = false; break; }
                }
                if (matchAttrs) results.Add(ffNode);
            }
            return results;
        }

        // Handle editable selector — aggregates all editable SDTs and form fields, sorted by document position
        if (parsed.Element == "editable")
        {
            // Collect editable SDTs
            int blockSdtIdx = 0;
            foreach (var sdt in body.Descendants().Where(e => e is SdtBlock or SdtRun))
            {
                string sdtPath;
                if (sdt is SdtBlock)
                {
                    blockSdtIdx++;
                    sdtPath = $"/body/{BuildSdtPathSegment(sdt, blockSdtIdx)}";
                }
                else if (sdt is SdtRun sdtRun)
                {
                    var parentPara = sdtRun.Ancestors<Paragraph>().FirstOrDefault();
                    if (parentPara != null)
                    {
                        int pIdx = 1;
                        foreach (var el in body.ChildElements)
                        {
                            if (el == parentPara) break;
                            if (el is Paragraph) pIdx++;
                        }
                        int sdtInParaIdx = 1;
                        foreach (var child in parentPara.ChildElements)
                        {
                            if (child == sdtRun) break;
                            if (child is SdtRun) sdtInParaIdx++;
                        }
                        sdtPath = $"/body/{BuildParaPathSegment(parentPara, pIdx)}/{BuildSdtPathSegment(sdt, sdtInParaIdx)}";
                    }
                    else
                    {
                        blockSdtIdx++;
                        sdtPath = $"/body/{BuildSdtPathSegment(sdt, blockSdtIdx)}";
                    }
                }
                else continue;

                var sdtNode = ElementToNode(sdt, sdtPath, 0);
                if (sdtNode.Format.TryGetValue("editable", out var editableVal) && editableVal is true)
                    results.Add(sdtNode);
            }

            // Collect editable form fields
            var allFormFields = FindFormFields();
            for (int fi = 0; fi < allFormFields.Count; fi++)
            {
                var ffNode = FormFieldToNode(allFormFields[fi], $"/formfield[{fi + 1}]");
                if (ffNode.Format.TryGetValue("editable", out var editableVal) && editableVal is true)
                    results.Add(ffNode);
            }

            return results;
        }

        // Determine if main selector targets runs directly (no > parent).
        // CONSISTENCY(run-special-content): the specialized run-kind types
        // exposed by Get (ptab/fieldChar/instrText/tab/break) all live as
        // <w:r> children of a paragraph — they reuse the run-walk dispatch
        // and let MatchesRunSelector type-filter on the actual inline payload.
        bool isRunSelector = parsed.ChildSelector == null &&
            (parsed.Element == "r" || parsed.Element == "run"
             || parsed.Element == "ptab" || parsed.Element == "positionaltab"
             || parsed.Element == "fieldchar" || parsed.Element == "fldchar"
             || parsed.Element == "instrtext"
             || parsed.Element == "tab"
             || parsed.Element == "break" || parsed.Element == "br");
        bool isPictureSelector = parsed.ChildSelector == null &&
            (parsed.Element == "picture" || parsed.Element == "image" || parsed.Element == "img");
        bool isOleSelector = parsed.ChildSelector == null &&
            // CONSISTENCY(ole-alias): "oleobject" mirrors Add's "ole"/"oleobject"/"object"/"embed" switch
            (parsed.Element is "ole" or "oleobject" or "object" or "embed");
        bool isEquationSelector = parsed.ChildSelector == null &&
            (parsed.Element == "equation" || parsed.Element == "math" || parsed.Element == "formula");
        bool isBookmarkSelector = parsed.ChildSelector == null &&
            parsed.Element == "bookmark";
        bool isSdtSelector = parsed.ChildSelector == null &&
            (parsed.Element == "sdt" || parsed.Element == "contentcontrol");
        // CONSISTENCY(word-table-recurse): paragraph selectors must descend
        // into table cells (B11 — fuzzer-A). Mirrors run/ole/equation
        // table-recurse branches added previously (issue #68).
        bool isParagraphSelector = parsed.ChildSelector == null &&
            (parsed.Element == "p" || parsed.Element == "paragraph");

        // Scheme B: generic XML fallback for unrecognized element types
        // Use GenericXmlQuery.ParseSelector which properly handles namespace prefixes (e.g., "a:ln")
        var genericParsed = GenericXmlQuery.ParseSelector(selector);
        // CONSISTENCY(selector-case): high-level element names are case-insensitive
        // ("OLE" == "ole"). Compare against the lowercase literal list.
        var genericElementLower = (genericParsed.element ?? "").ToLowerInvariant();
        bool isKnownType = string.IsNullOrEmpty(genericElementLower)
            || genericElementLower is "p" or "paragraph" or "r" or "run"
                or "picture" or "image" or "img"
                or "equation" or "math" or "formula"
                or "bookmark"
                or "sdt" or "contentcontrol"
                or "chart"
                or "comment"
                or "footnote" or "endnote"
                or "field" or "formfield" or "editable"
                or "table" or "tbl"
                or "row" or "tr" or "cell" or "tc"
                or "toc" or "tableofcontents"
                or "style" or "styles"
                or "watermark"
                or "revision" or "change" or "trackchange"
                or "media"
                or "hyperlink"
                or "section"
                or "ole" or "oleobject" or "object" or "embed"
                // CONSISTENCY(run-special-content): specialized run-kind
                // types are dispatched via the run-walk above; treat them
                // as known so they don't fall through to GenericXmlQuery,
                // which would emit non-canonical OOXML-element paths
                // (/p[N]/r[N]/br[1] etc.) that don't pipe back to set/get.
                or "ptab" or "positionaltab"
                or "fieldchar" or "fldchar"
                or "instrtext"
                or "tab"
                or "break" or "br";
        if (!isKnownType && parsed.ChildSelector == null)
        {
            var root = _doc.MainDocumentPart?.Document;
            if (root != null)
            {
                var genericResults = GenericXmlQuery.Query(root, genericParsed.element ?? "", genericParsed.attrs, genericParsed.containsText);
                // Canonicalize emitted paths so they resolve via `get` /
                // `add --after`. The generic traversal starts at <w:document>
                // and produces `/document[1]/body[1]/...` but Navigation
                // expects paths rooted at `/body`. Strip the document prefix.
                const string docPrefix = "/document[1]/body[1]";
                foreach (var n in genericResults)
                {
                    if (n.Path != null && n.Path.StartsWith(docPrefix, StringComparison.Ordinal))
                        n.Path = "/body" + n.Path[docPrefix.Length..];
                    else if (n.Path == "/document[1]")
                        n.Path = "/";
                }
                return genericResults;
            }
            return results;
        }

        // Handle media query (same as picture/image but explicitly named "media")
        if (parsed.ChildSelector == null && parsed.Element == "media")
        {
            int mediaPIdx = 0;
            foreach (var para in body.Elements<Paragraph>())
            {
                int mediaRIdx = 0;
                foreach (var run in GetAllRuns(para))
                {
                    var drawing = run.GetFirstChild<Drawing>();
                    if (drawing != null)
                    {
                        var node = CreateImageNode(drawing, run, $"/body/{BuildParaPathSegment(para, mediaPIdx + 1)}/r[{mediaRIdx + 1}]");
                        // Add content type from image part
                        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                        if (blip?.Embed?.Value != null)
                        {
                            var part = _doc.MainDocumentPart?.GetPartById(blip.Embed.Value);
                            if (part != null)
                            {
                                node.Format["contentType"] = part.ContentType;
                                node.Format["fileSize"] = part.GetStream().Length;
                            }
                        }
                        results.Add(node);
                    }
                    mediaRIdx++;
                }
                mediaPIdx++;
            }
            return results;
        }

        // Handle toc query
        if (parsed.ChildSelector == null && (parsed.Element is "toc" or "tableofcontents"))
        {
            var tocParas = FindTocParagraphs();
            for (int ti = 0; ti < tocParas.Count; ti++)
            {
                var tocPara = tocParas[ti];
                var instrText = string.Join("", tocPara.Descendants<FieldCode>().Select(fc => fc.Text));
                var tocNode = new DocumentNode { Path = $"/toc[{ti + 1}]", Type = "toc" };
                tocNode.Text = instrText.Trim();
                var levelsMatch = System.Text.RegularExpressions.Regex.Match(instrText, @"\\o\s+""([^""]+)""");
                if (levelsMatch.Success) tocNode.Format["levels"] = levelsMatch.Groups[1].Value;
                tocNode.Format["hyperlinks"] = instrText.Contains("\\h");
                tocNode.Format["pageNumbers"] = !instrText.Contains("\\z");
                if (parsed.ContainsText != null && !(tocNode.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                    continue;
                results.Add(tocNode);
            }
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
                if (chartInfo.DocProperties?.Id?.HasValue == true)
                    node.Format["id"] = chartInfo.DocProperties.Id.Value;
                if (chartInfo.DocProperties?.Name?.Value != null)
                    node.Format["name"] = chartInfo.DocProperties.Name.Value;

                if (chartInfo.IsExtended)
                {
                    var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace!;
                    var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                    if (cxType != null) node.Format["chartType"] = cxType;
                    // Title
                    var cxTitle = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
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

        // Handle OLE query via descendants walk — covers body paragraphs,
        // top-level tables, nested tables, textboxes, etc. CONSISTENCY(word-ole-query):
        // a single Descendants<EmbeddedObject>() pass replaces the previous
        // hand-rolled body + top-level-table scan which missed nested tables.
        // Also walks HeaderPart/FooterPart documents so that OLEs added via
        // `Add("/header[N]", "ole", ...)` are surfaced after reopen.
        if (isOleSelector)
        {
            // BUG-R15-01: the OLE query block never applied parsed.Attributes filters,
            // so Query("ole[objectType=nonexistent]") returned all OLEs instead of 0.
            // CONSISTENCY(query-attr-filter): apply the same Format-key attribute
            // matching used by style/field/formfield/PPT-OLE selectors in the same file.
            static bool OleMatchesAttrs(DocumentNode node, Dictionary<string, string> attrs)
            {
                foreach (var (attrKey, rawVal) in attrs)
                {
                    bool negate = rawVal.StartsWith("!");
                    var val = negate ? rawVal[1..] : rawVal;
                    var hasKey = node.Format.TryGetValue(attrKey, out var fmtVal);
                    bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                    if (negate ? matches : !matches) return false;
                }
                return true;
            }

            foreach (var oleObject in body.Descendants<EmbeddedObject>())
            {
                var run = oleObject.Ancestors<Run>().FirstOrDefault();
                if (run == null) continue;
                var olePath = BuildOleRunPath(body, "/body", run);
                var oleNode = CreateOleNode(oleObject, run, olePath);
                if (OleMatchesAttrs(oleNode, parsed.Attributes)) results.Add(oleNode);
            }

            var mainPart = _doc.MainDocumentPart;
            if (mainPart != null)
            {
                int hIdx = 0;
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    hIdx++;
                    var header = headerPart.Header;
                    if (header == null) continue;
                    foreach (var oleObject in header.Descendants<EmbeddedObject>())
                    {
                        var run = oleObject.Ancestors<Run>().FirstOrDefault();
                        if (run == null) continue;
                        var olePath = BuildOleRunPath(header, $"/header[{hIdx}]", run);
                        // BUG-R10-02: rel id lives on the HeaderPart, not
                        // MainDocumentPart — pass the headerPart so
                        // CreateOleNode can populate contentType/fileSize.
                        var oleNode = CreateOleNode(oleObject, run, olePath, headerPart);
                        if (OleMatchesAttrs(oleNode, parsed.Attributes)) results.Add(oleNode);
                    }
                }
                int fIdx = 0;
                foreach (var footerPart in mainPart.FooterParts)
                {
                    fIdx++;
                    var footer = footerPart.Footer;
                    if (footer == null) continue;
                    foreach (var oleObject in footer.Descendants<EmbeddedObject>())
                    {
                        var run = oleObject.Ancestors<Run>().FirstOrDefault();
                        if (run == null) continue;
                        var olePath = BuildOleRunPath(footer, $"/footer[{fIdx}]", run);
                        // BUG-R10-02: same fix for footers.
                        var oleNode = CreateOleNode(oleObject, run, olePath, footerPart);
                        if (OleMatchesAttrs(oleNode, parsed.Attributes)) results.Add(oleNode);
                    }
                }
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
                        Path = comment.Id?.Value != null ? $"/comments/comment[@commentId={comment.Id.Value}]" : $"/comments/comment[{cIdx}]",
                        Type = "comment",
                        Text = text
                    };
                    if (comment.Author?.Value != null) cNode.Format["author"] = comment.Author.Value;
                    if (comment.Initials?.Value != null) cNode.Format["initials"] = comment.Initials.Value;
                    if (comment.Id?.Value != null) cNode.Format["id"] = comment.Id.Value;
                    if (comment.Date?.Value != null) cNode.Format["date"] = comment.Date.Value.ToString("o");
                    if (comment.Id?.Value != null)
                    {
                        var anchorPath = FindCommentAnchorPath(comment.Id.Value);
                        if (anchorPath != null) cNode.Format["anchoredTo"] = anchorPath;
                    }
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
                    var text = GetFootnoteText(fn);
                    if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var fnNode = new DocumentNode
                    {
                        Path = fn.Id?.Value != null ? $"/footnote[@footnoteId={fn.Id.Value}]" : $"/footnote[{fnIdx}]",
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
                    var text = GetFootnoteText(en);
                    if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var enNode = new DocumentNode
                    {
                        Path = en.Id?.Value != null ? $"/endnote[@endnoteId={en.Id.Value}]" : $"/endnote[{enIdx}]",
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

        // Handle hyperlink query
        bool isHyperlinkSelector = parsed.ChildSelector == null && parsed.Element == "hyperlink";
        if (isHyperlinkSelector)
        {
            int hlIdx = 0;
            foreach (var hl in body.Descendants<Hyperlink>())
            {
                hlIdx++;
                var text = string.Concat(hl.Descendants<Text>().Select(t => t.Text));
                if (parsed.ContainsText != null && !text.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase))
                { hlIdx--; continue; }

                // Build node via ElementToNode to get full format (link, color, underline, etc.)
                var parentPara = hl.Ancestors<Paragraph>().FirstOrDefault();
                int pIdx = 1;
                int hlInParaIdx = 1;
                if (parentPara != null)
                {
                    foreach (var el in body.ChildElements)
                    {
                        if (el == parentPara) break;
                        if (el is Paragraph) pIdx++;
                    }
                    foreach (var child in parentPara.ChildElements)
                    {
                        if (child == hl) break;
                        if (child is Hyperlink) hlInParaIdx++;
                    }
                }
                var hlPath = parentPara != null ? $"/body/{BuildParaPathSegment(parentPara, pIdx)}/hyperlink[{hlInParaIdx}]" : $"/body/p[{pIdx}]/hyperlink[{hlInParaIdx}]";
                var node = ElementToNode(hl, hlPath, 0);

                // Filter by attributes
                bool matchAttrs = true;
                foreach (var (attrKey, rawVal) in parsed.Attributes)
                {
                    bool negate = rawVal.StartsWith("!");
                    var val = negate ? rawVal[1..] : rawVal;
                    var hasKey = node.Format.TryGetValue(attrKey, out var fmtVal);
                    bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                    if (negate ? matches : !matches) { matchAttrs = false; break; }
                }
                if (!matchAttrs) continue;
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

                results.Add(ElementToNode(bkStart, $"/bookmark[@name={bkName}]", 0));
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
                    path = $"/body/{BuildSdtPathSegment(sdt, blockSdtIdx)}";
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
                        path = $"/body/{BuildParaPathSegment(parentPara, pIdx)}/{BuildSdtPathSegment(sdt, sdtInParaIdx)}";
                    }
                    else
                    {
                        blockSdtIdx++;
                        path = $"/body/{BuildSdtPathSegment(sdt, blockSdtIdx)}";
                    }
                }
                else
                {
                    blockSdtIdx++;
                    path = $"/body/{BuildSdtPathSegment(sdt, blockSdtIdx)}";
                }
                var node = ElementToNode(sdt, path, 0);
                if (parsed.ContainsText != null && !(node.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                    continue;
                // Filter by attributes (e.g., sdt[tag=partyA])
                bool matchAttrs = true;
                foreach (var (attrKey, rawVal) in parsed.Attributes)
                {
                    bool negate = rawVal.StartsWith("!");
                    var val = negate ? rawVal[1..] : rawVal;
                    var hasKey = node.Format.TryGetValue(attrKey, out var fmtVal);
                    bool matches = hasKey && string.Equals(fmtVal?.ToString(), val, StringComparison.OrdinalIgnoreCase);
                    if (negate ? matches : !matches) { matchAttrs = false; break; }
                }
                if (!matchAttrs) continue;
                results.Add(node);
            }
            return results;
        }

        // BUG-R34-02: row / cell queries (canonical names + tr/tc internal aliases).
        // Walks every body-level table emitting one node per row or per cell. Type field
        // is canonical "row" / "cell" (matches ElementToNode + Get readback in
        // WordHandler.Navigation.cs ~line 1300). Path uses internal `tr[]/tc[]` segments
        // for round-trip with Get.
        if (parsed.ChildSelector == null &&
            parsed.Element is "row" or "tr" or "cell" or "tc")
        {
            bool wantRow = parsed.Element is "row" or "tr";
            int tblIdxRC = 0;
            foreach (var tbl in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>())
            {
                tblIdxRC++;
                int rowIdxRC = 0;
                foreach (var row in tbl.Elements<TableRow>())
                {
                    rowIdxRC++;
                    if (wantRow)
                    {
                        var rowPath = $"/body/tbl[{tblIdxRC}]/tr[{rowIdxRC}]";
                        var rowNode = ElementToNode(row, rowPath, 0);
                        if (parsed.ContainsText != null
                            && !(rowNode.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                            continue;
                        bool ok = true;
                        foreach (var (attrKey, rawVal) in parsed.Attributes)
                        {
                            bool negate = rawVal.StartsWith("!");
                            var aval = negate ? rawVal[1..] : rawVal;
                            var has = rowNode.Format.TryGetValue(attrKey, out var fv);
                            bool m = has && string.Equals(fv?.ToString(), aval, StringComparison.OrdinalIgnoreCase);
                            if (negate ? m : !m) { ok = false; break; }
                        }
                        if (ok) results.Add(rowNode);
                    }
                    else
                    {
                        int cellIdxRC = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            cellIdxRC++;
                            var cellPath = $"/body/tbl[{tblIdxRC}]/tr[{rowIdxRC}]/tc[{cellIdxRC}]";
                            var cellNode = ElementToNode(cell, cellPath, 0);
                            if (parsed.ContainsText != null
                                && !(cellNode.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                                continue;
                            bool ok = true;
                            foreach (var (attrKey, rawVal) in parsed.Attributes)
                            {
                                bool negate = rawVal.StartsWith("!");
                                var aval = negate ? rawVal[1..] : rawVal;
                                var has = cellNode.Format.TryGetValue(attrKey, out var fv);
                                bool m = has && string.Equals(fv?.ToString(), aval, StringComparison.OrdinalIgnoreCase);
                                if (negate ? m : !m) { ok = false; break; }
                            }
                            if (ok) results.Add(cellNode);
                        }
                    }
                }
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
                else if (isOleSelector)
                {
                    // Scan inside table cells for OLE objects. CONSISTENCY(word-ole-query):
                    // mirrors the body-level OLE branch (see isOleSelector block below for
                    // free-body paragraphs). Without this branch, `Query("ole")` silently
                    // skips any OLE embedded in a table cell.
                    var tblIdx = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>()
                        .TakeWhile(t => t != tbl).Count();
                    int rowIdx = 0;
                    foreach (var row in tbl.Elements<TableRow>())
                    {
                        rowIdx++;
                        int cellIdx = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            cellIdx++;
                            int cellParaIdx = 0;
                            foreach (var cellPara in cell.Elements<Paragraph>())
                            {
                                cellParaIdx++;
                                int cellRunIdx = 0;
                                foreach (var cellRun in GetAllRuns(cellPara))
                                {
                                    cellRunIdx++;
                                    var oleObject = cellRun.GetFirstChild<EmbeddedObject>();
                                    if (oleObject != null)
                                    {
                                        results.Add(CreateOleNode(oleObject, cellRun,
                                            $"/body/tbl[{tblIdx + 1}]/tr[{rowIdx}]/tc[{cellIdx}]/{BuildParaPathSegment(cellPara, cellParaIdx)}/r[{cellRunIdx}]"));
                                    }
                                }
                            }
                        }
                    }
                }
                else if (isEquationSelector)
                {
                    // Scan inside table cells for equations
                    var tblIdx = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>()
                        .TakeWhile(t => t != tbl).Count();
                    int rowIdx = 0;
                    foreach (var row in tbl.Elements<TableRow>())
                    {
                        rowIdx++;
                        int cellIdx = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            cellIdx++;
                            int cellParaIdx = 0;
                            foreach (var cellPara in cell.Elements<Paragraph>())
                            {
                                cellParaIdx++;
                                // Display equations inside table cell paragraphs
                                var oMathParaInCell = cellPara.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                                if (oMathParaInCell != null)
                                {
                                    mathParaIdx++;
                                    var latex = FormulaParser.ToLatex(oMathParaInCell);
                                    if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                                    {
                                        results.Add(new DocumentNode
                                        {
                                            Path = $"/body/tbl[{tblIdx + 1}]/tr[{rowIdx}]/tc[{cellIdx}]/oMathPara[{mathParaIdx + 1}]",
                                            Type = "equation",
                                            Text = latex,
                                            Format = { ["mode"] = "display" }
                                        });
                                    }
                                    continue;
                                }

                                // Inline equations inside table cell paragraphs
                                int cellMathIdx = 0;
                                foreach (var oMath in cellPara.ChildElements.Where(e => e.LocalName == "oMath" || e is M.OfficeMath))
                                {
                                    var latex = FormulaParser.ToLatex(oMath);
                                    if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                                    {
                                        results.Add(new DocumentNode
                                        {
                                            Path = $"/body/tbl[{tblIdx + 1}]/tr[{rowIdx}]/tc[{cellIdx}]/p[{cellParaIdx}]/oMath[{cellMathIdx + 1}]",
                                            Type = "equation",
                                            Text = latex,
                                            Format = { ["mode"] = "inline" }
                                        });
                                    }
                                    cellMathIdx++;
                                }
                            }
                        }
                    }
                }
                else if (isParagraphSelector)
                {
                    // Scan inside table cells for paragraphs. CONSISTENCY(word-table-recurse):
                    // mirrors the run/ole/equation branches. Without this, `query paragraph`
                    // silently skips any paragraph inside a table cell. (B11)
                    var tblIdx = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>()
                        .TakeWhile(t => t != tbl).Count();
                    int rowIdxP = 0;
                    foreach (var row in tbl.Elements<TableRow>())
                    {
                        rowIdxP++;
                        int cellIdxP = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            cellIdxP++;
                            int cellParaIdx = 0;
                            foreach (var cellPara in cell.Elements<Paragraph>())
                            {
                                cellParaIdx++;
                                var paraPath = $"/body/tbl[{tblIdx + 1}]/tr[{rowIdxP}]/tc[{cellIdxP}]/{BuildParaPathSegment(cellPara, cellParaIdx)}";
                                var paraNode = ElementToNode(cellPara, paraPath, 0);
                                if (parsed.ContainsText != null
                                    && !(paraNode.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                                    continue;
                                bool ok = true;
                                foreach (var (attrKey, rawVal) in parsed.Attributes)
                                {
                                    bool negate = rawVal.StartsWith("!");
                                    var aval = negate ? rawVal[1..] : rawVal;
                                    var has = paraNode.Format.TryGetValue(attrKey, out var fv);
                                    bool m = has && string.Equals(fv?.ToString(), aval, StringComparison.OrdinalIgnoreCase);
                                    if (negate ? m : !m) { ok = false; break; }
                                }
                                if (ok) results.Add(paraNode);
                            }
                        }
                    }
                }
                else if (isRunSelector)
                {
                    // Scan inside table cells for runs. CONSISTENCY(word-ole-query):
                    // mirrors the OLE/equation branches above. Without this, run
                    // selectors like `run[color=#FF0000]` silently skip any run
                    // inside a table cell. (issue #68)
                    var tblIdx = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>()
                        .TakeWhile(t => t != tbl).Count();
                    int rowIdx = 0;
                    foreach (var row in tbl.Elements<TableRow>())
                    {
                        rowIdx++;
                        int cellIdx = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            cellIdx++;
                            int cellParaIdx = 0;
                            foreach (var cellPara in cell.Elements<Paragraph>())
                            {
                                cellParaIdx++;
                                int cellRunIdx = 0;
                                foreach (var cellRun in GetAllRuns(cellPara))
                                {
                                    cellRunIdx++;
                                    if (MatchesRunSelector(cellRun, cellPara, parsed))
                                    {
                                        results.Add(ElementToNode(cellRun,
                                            $"/body/tbl[{tblIdx + 1}]/tr[{rowIdx}]/tc[{cellIdx}]/{BuildParaPathSegment(cellPara, cellParaIdx)}/r[{cellRunIdx}]", 0));
                                    }
                                }
                            }
                        }
                    }
                }
                continue;
            }

            if (element is Paragraph para)
            {
                // #6: a w:p whose sole content is m:oMathPara is addressed
                // via /body/oMathPara[M], not /body/p[N]. Don't bump paraIdx
                // for these wrappers so /body/p[N] indexes only real prose.
                if (IsOMathParaWrapperParagraph(para))
                {
                    mathParaIdx++;
                    if (isEquationSelector)
                    {
                        var oMathParaInPara = para.ChildElements.FirstOrDefault(
                            e => e.LocalName == "oMathPara" || e is M.Paragraph);
                        var latex = FormulaParser.ToLatex(oMathParaInPara!);
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

                paraIdx++;

                if (isEquationSelector)
                {

                    // Find inline math in this paragraph
                    int mathIdx = 0;
                    foreach (var oMath in para.ChildElements.Where(e => e.LocalName == "oMath" || e is M.OfficeMath))
                    {
                        var latex = FormulaParser.ToLatex(oMath);
                        if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/body/{BuildParaPathSegment(para, paraIdx + 1)}/oMath[{mathIdx + 1}]",
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
                                    results.Add(CreateImageNode(drawing, run, $"/body/{BuildParaPathSegment(para, paraIdx + 1)}/r[{runIdx + 1}]"));
                            }
                            else
                            {
                                results.Add(CreateImageNode(drawing, run, $"/body/{BuildParaPathSegment(para, paraIdx + 1)}/r[{runIdx + 1}]"));
                            }
                        }

                        // CONSISTENCY(ole-query-separation): OLE objects have
                        // their own `query ole` selector. Do not surface them
                        // in picture/image results — even though OLE wraps a
                        // v:imagedata for the icon preview, that is not a real
                        // picture from the user's perspective.
                        runIdx++;
                    }
                }
                else if (isOleSelector)
                {
                    int runIdx = 0;
                    foreach (var run in GetAllRuns(para))
                    {
                        var oleObject = run.GetFirstChild<EmbeddedObject>();
                        if (oleObject != null)
                        {
                            results.Add(CreateOleNode(oleObject, run, $"/body/{BuildParaPathSegment(para, paraIdx + 1)}/r[{runIdx + 1}]"));
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
                            results.Add(ElementToNode(run, $"/body/{BuildParaPathSegment(para, paraIdx + 1)}/r[{runIdx + 1}]", 0));
                        }
                        runIdx++;
                    }
                }
                else
                {
                    // When ChildSelector is present (e.g. "paragraph[...] > run[...]"),
                    // the user is asking for child runs whose parent matches, not
                    // mixed parent+child results. Only emit child runs in that case.
                    if (parsed.ChildSelector != null)
                    {
                        // MatchesSelector already gated the paragraph via its
                        // ChildSelector-aware branch; iterate matching runs here.
                        if (MatchesSelector(para, parsed, paraIdx))
                        {
                            int runIdx = 0;
                            foreach (var run in GetAllRuns(para))
                            {
                                if (MatchesRunSelector(run, para, parsed.ChildSelector))
                                {
                                    results.Add(ElementToNode(run, $"/body/{BuildParaPathSegment(para, paraIdx + 1)}/r[{runIdx + 1}]", 0));
                                }
                                runIdx++;
                            }
                        }
                    }
                    else if (MatchesSelector(para, parsed, paraIdx))
                    {
                        results.Add(ElementToNode(para, $"/body/{BuildParaPathSegment(para, paraIdx + 1)}", 0));
                    }
                }
            }
        }

        // CONSISTENCY(word-headerfooter-recurse): paragraph/run selectors must
        // also descend into header/footer parts (B12 — fuzzer-B). Without this,
        // `query paragraph` and `query run` silently skip any paragraph/run
        // that lives in a header or footer. Path prefix is /header[N] or
        // /footer[N], indexed by 1-based encounter order in the rels.
        // CONSISTENCY(query-combinator-headerfooter): combinator selectors
        // (p > ptab / paragraph > fieldChar) also need to descend so the
        // child runs in headers/footers are reachable; the dispatch inside
        // CollectParaRunInHeaderFooter handles all three modes.
        bool isCombinator = parsed.ChildSelector != null
            && (parsed.Element == null || parsed.Element == "p" || parsed.Element == "paragraph");
        if (isParagraphSelector || isRunSelector || isCombinator)
        {
            var mainPart = _doc.MainDocumentPart;
            if (mainPart != null)
            {
                int hIdx = 0;
                foreach (var hp in mainPart.HeaderParts)
                {
                    hIdx++;
                    var hRoot = hp.Header;
                    if (hRoot == null) continue;
                    CollectParaRunInHeaderFooter(hRoot, $"/header[{hIdx}]", parsed, isParagraphSelector, isRunSelector, results);
                }
                int fIdx = 0;
                foreach (var fp in mainPart.FooterParts)
                {
                    fIdx++;
                    var fRoot = fp.Footer;
                    if (fRoot == null) continue;
                    CollectParaRunInHeaderFooter(fRoot, $"/footer[{fIdx}]", parsed, isParagraphSelector, isRunSelector, results);
                }
                // CONSISTENCY(query-aux-parts-recurse): paragraph/run/combinator
                // selectors must descend into footnotes/endnotes/comments too.
                // EnsureAllParaIds (Round 2) already scans these for paraId
                // uniqueness; query was the asymmetric outlier that hid every
                // ptab/fieldChar/instrText living in those parts.
                if (mainPart.FootnotesPart?.Footnotes != null)
                {
                    int fnIdx = 0;
                    foreach (var fn in mainPart.FootnotesPart.Footnotes.Elements<Footnote>())
                    {
                        fnIdx++;
                        CollectParaRunInHeaderFooter(fn, $"/footnote[{fnIdx}]", parsed, isParagraphSelector, isRunSelector, results);
                    }
                }
                if (mainPart.EndnotesPart?.Endnotes != null)
                {
                    int enIdx = 0;
                    foreach (var en in mainPart.EndnotesPart.Endnotes.Elements<Endnote>())
                    {
                        enIdx++;
                        CollectParaRunInHeaderFooter(en, $"/endnote[{enIdx}]", parsed, isParagraphSelector, isRunSelector, results);
                    }
                }
                if (mainPart.WordprocessingCommentsPart?.Comments != null)
                {
                    int cIdx = 0;
                    foreach (var cmt in mainPart.WordprocessingCommentsPart.Comments.Elements<Comment>())
                    {
                        cIdx++;
                        CollectParaRunInHeaderFooter(cmt, $"/comment[{cIdx}]", parsed, isParagraphSelector, isRunSelector, results);
                    }
                }
            }
        }

        return results;
    }

    /// <summary>
    /// Collect paragraphs/runs inside a header/footer root using positional
    /// indexing matching the body convention (no table recursion yet — keep
    /// the recurse minimal; mirrors Selection's known-positional limitation).
    /// </summary>
    private void CollectParaRunInHeaderFooter(
        OpenXmlElement root, string pathPrefix, SelectorPart parsed,
        bool isParagraphSelector, bool isRunSelector, List<DocumentNode> results)
    {
        int paraIdx = 0;
        foreach (var element in root.ChildElements.OfType<Paragraph>())
        {
            paraIdx++;
            if (isParagraphSelector)
            {
                var paraPath = $"{pathPrefix}/{BuildParaPathSegment(element, paraIdx)}";
                var paraNode = ElementToNode(element, paraPath, 0);
                if (parsed.ContainsText != null
                    && !(paraNode.Text?.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase) ?? false))
                    continue;
                bool ok = true;
                foreach (var (attrKey, rawVal) in parsed.Attributes)
                {
                    bool negate = rawVal.StartsWith("!");
                    var aval = negate ? rawVal[1..] : rawVal;
                    var has = paraNode.Format.TryGetValue(attrKey, out var fv);
                    bool m = has && string.Equals(fv?.ToString(), aval, StringComparison.OrdinalIgnoreCase);
                    if (negate ? m : !m) { ok = false; break; }
                }
                if (ok) results.Add(paraNode);
            }
            else if (isRunSelector)
            {
                int runIdx = 0;
                foreach (var run in GetAllRuns(element))
                {
                    runIdx++;
                    if (MatchesRunSelector(run, element, parsed))
                    {
                        results.Add(ElementToNode(run,
                            $"{pathPrefix}/{BuildParaPathSegment(element, paraIdx)}/r[{runIdx}]", 0));
                    }
                }
            }
            else if (parsed.ChildSelector != null
                && (parsed.Element == null || parsed.Element == "p" || parsed.Element == "paragraph"))
            {
                // CONSISTENCY(query-combinator-headerfooter): mirror the body
                // dispatch's combinator branch (`p > X` / `p[...] > X`) so
                // descendant selectors find runs inside header/footer too.
                // Without this, `query "p > ptab"` returned 0 for documents
                // whose ptabs all live in headers/footers (the typical case).
                bool ok = true;
                foreach (var (attrKey, rawVal) in parsed.Attributes)
                {
                    bool negate = rawVal.StartsWith("!");
                    var aval = negate ? rawVal[1..] : rawVal;
                    var paraNode = ElementToNode(element,
                        $"{pathPrefix}/{BuildParaPathSegment(element, paraIdx)}", 0);
                    var has = paraNode.Format.TryGetValue(attrKey, out var fv);
                    bool m = has && string.Equals(fv?.ToString(), aval, StringComparison.OrdinalIgnoreCase);
                    if (negate ? m : !m) { ok = false; break; }
                }
                if (!ok) continue;
                int runIdx = 0;
                foreach (var run in GetAllRuns(element))
                {
                    runIdx++;
                    if (MatchesRunSelector(run, element, parsed.ChildSelector))
                    {
                        results.Add(ElementToNode(run,
                            $"{pathPrefix}/{BuildParaPathSegment(element, paraIdx)}/r[{runIdx}]", 0));
                    }
                }
            }
        }
    }

    /// <summary>
    /// Builds a root-rooted path to a Run by walking its ancestor chain,
    /// emitting a tbl[i]/tr[j]/tc[k] segment for every enclosing table.
    /// Covers top-level runs, runs inside top-level tables, and runs inside
    /// nested tables. Used by OLE Query so that Descendants&lt;EmbeddedObject&gt;()
    /// can surface OLEs at any depth. The root can be a Body, Header, or
    /// Footer; the rootPath prefix is used verbatim (e.g. "/body",
    /// "/header[1]", "/footer[2]").
    /// </summary>
    private static string BuildOleRunPath(OpenXmlElement root, string rootPath, Run run)
    {
        // Walk from root down to the run, collecting path segments.
        // Ancestors() returns innermost first; reverse to outer-to-inner order.
        var ancestors = run.Ancestors().TakeWhile(a => a != root).Reverse().ToList();

        var sb = new System.Text.StringBuilder(rootPath);
        OpenXmlElement cursor = root;
        foreach (var anc in ancestors)
        {
            if (anc is SdtBlock sdtBlockAnc)
            {
                // Count SdtBlocks among the current cursor's direct children
                var sdtIdx = cursor.ChildElements.OfType<SdtBlock>()
                    .TakeWhile(s => s != sdtBlockAnc).Count() + 1;
                sb.Append($"/{BuildSdtPathSegment(sdtBlockAnc, sdtIdx)}");
                cursor = sdtBlockAnc;
            }
            else if (anc is SdtContentBlock sdtContentBlockAnc)
            {
                // SdtContentBlock is implicit in the path format; descend
                // into it without emitting a segment, mirroring Navigation.
                cursor = sdtContentBlockAnc;
            }
            else if (anc is SdtRun sdtRunAnc)
            {
                var sdtIdx = cursor.ChildElements.OfType<SdtRun>()
                    .TakeWhile(s => s != sdtRunAnc).Count() + 1;
                sb.Append($"/{BuildSdtPathSegment(sdtRunAnc, sdtIdx)}");
                cursor = sdtRunAnc;
            }
            else if (anc is SdtContentRun sdtContentRunAnc)
            {
                cursor = sdtContentRunAnc;
            }
            else if (anc is DocumentFormat.OpenXml.Wordprocessing.Table tblAnc)
            {
                // Index among sibling tables within the current cursor
                var tblIdx = cursor.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>()
                    .TakeWhile(t => t != tblAnc).Count() + 1;
                sb.Append($"/tbl[{tblIdx}]");
                cursor = tblAnc;
            }
            else if (anc is TableRow rowAnc)
            {
                var rowIdx = cursor.Elements<TableRow>()
                    .TakeWhile(r => r != rowAnc).Count() + 1;
                sb.Append($"/tr[{rowIdx}]");
                cursor = rowAnc;
            }
            else if (anc is TableCell cellAnc)
            {
                var cellIdx = cursor.Elements<TableCell>()
                    .TakeWhile(c => c != cellAnc).Count() + 1;
                sb.Append($"/tc[{cellIdx}]");
                cursor = cellAnc;
            }
            else if (anc is Paragraph paraAnc)
            {
                var paraIdx = cursor.Elements<Paragraph>()
                    .TakeWhile(p => p != paraAnc).Count() + 1;
                sb.Append($"/{BuildParaPathSegment(paraAnc, paraIdx)}");
                cursor = paraAnc;
            }
        }

        // Run index within its parent paragraph (via GetAllRuns to handle sdt wrappers)
        if (run.Ancestors<Paragraph>().FirstOrDefault() is Paragraph parentPara)
        {
            var runs = GetAllRuns(parentPara);
            var runIdx = runs.TakeWhile(r => r != run).Count() + 1;
            sb.Append($"/r[{runIdx}]");
        }

        return sb.ToString();
    }
}
