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

        // Unified find: if 'find' key is present (at any path level), route to ProcessFind
        if (properties.TryGetValue("find", out var findText))
        {
            var replace = properties.TryGetValue("replace", out var r) ? r : null;
            // Separate run-level format properties from paragraph-level properties
            var formatProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var paraProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (key, value) in properties)
            {
                var k = key.ToLowerInvariant();
                if (k is "find" or "replace" or "scope" or "regex") continue;
                // Paragraph-level properties go to paraProps
                if (k is "style" or "alignment" or "align" or "firstlineindent" or "leftindent" or "indentleft"
                    or "indent" or "rightindent" or "indentright" or "hangingindent" or "spacebefore"
                    or "spaceafter" or "linespacing" or "keepnext" or "keeplines" or "pagebreakbefore"
                    or "widowcontrol" or "liststyle" or "start" or "text" or "formula"
                    or "contextualspacing")
                    paraProps[key] = value;
                else
                    formatProps[key] = value;
            }

            if (replace == null && formatProps.Count == 0 && paraProps.Count == 0)
                throw new ArgumentException("'find' requires either 'replace' and/or format properties (e.g. bold, highlight, color).");

            // CONSISTENCY(find-regex): canonical site for the `regex=true` → `r"..."`
            // raw-string normalization. `mark` and the other handlers' Set paths all
            // copy this pattern verbatim. To change the find/regex protocol,
            // grep "CONSISTENCY(find-regex)" and update every site project-wide;
            // do not diverge in a single handler.
            if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
                findText = $"r\"{findText}\"";

            var effectivePath = (path is "" or "/") ? "/body" : path;
            var matchCount = ProcessFind(effectivePath, findText, replace, formatProps.Count > 0 ? formatProps : new Dictionary<string, string>());
            LastFindMatchCount = matchCount;

            // Apply paragraph-level properties to the matched paragraphs
            if (paraProps.Count > 0)
            {
                var paragraphs = ResolveParagraphsForFind(effectivePath);
                foreach (var para in paragraphs)
                {
                    var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                    foreach (var (key, value) in paraProps)
                        ApplyParagraphLevelProperty(pProps, key, value);
                }
            }

            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Document-level properties
        if (path == "/" || path == "" || path.Equals("/body", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Handle /settings path — route to SetDocumentProperties which calls TrySetDocSetting
        if (path.Equals("/settings", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            EnsureSettings().Save();
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

        // FormField paths: /formfield[N] or /formfield[name]
        // Routed BEFORE ParsePath because the generic predicate validator
        // only accepts positive-integer / last() / [@attr=v] predicates and
        // would reject the documented /formfield[name] form.
        var ffSetMatchEarly = System.Text.RegularExpressions.Regex.Match(path, @"^/formfield\[(\w+)\]$");
        if (ffSetMatchEarly.Success)
        {
            var allFormFields = FindFormFields();
            var indexOrName = ffSetMatchEarly.Groups[1].Value;
            (FieldInfo Field, FormFieldData FfData) target;
            if (int.TryParse(indexOrName, out var ffIdx))
            {
                if (ffIdx < 1 || ffIdx > allFormFields.Count)
                    throw new ArgumentException($"FormField {ffIdx} not found (total: {allFormFields.Count})");
                target = allFormFields[ffIdx - 1];
            }
            else
            {
                target = allFormFields.FirstOrDefault(ff =>
                    ff.FfData.GetFirstChild<FormFieldName>()?.Val?.Value == indexOrName);
                if (target.Field == null)
                    throw new ArgumentException($"FormField '{indexOrName}' not found");
            }
            return SetFormField(target, properties);
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

        // Chart axis-by-role sub-path: /chart[N]/axis[@role=ROLE].
        var chartAxisSetMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success)
        {
            var caChartIdx = int.Parse(chartAxisSetMatch.Groups[1].Value);
            var caRole = chartAxisSetMatch.Groups[2].Value;
            var caAllCharts = GetAllWordCharts();
            if (caAllCharts.Count == 0)
                throw new ArgumentException("No charts in this document");
            if (caChartIdx < 1 || caChartIdx > caAllCharts.Count)
                throw new ArgumentException($"Chart {caChartIdx} not found (total: {caAllCharts.Count})");
            var caChartInfo = caAllCharts[caChartIdx - 1];
            if (caChartInfo.IsExtended || caChartInfo.StandardPart == null)
                throw new ArgumentException($"Axis Set not supported on extended charts.");
            unsupported.AddRange(Core.ChartHelper.SetAxisProperties(
                caChartInfo.StandardPart, caRole, properties));
            return unsupported;
        }

        // Chart paths: /chart[N] or /chart[N]/series[K]
        var chartMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success)
        {
            var chartIdx = int.Parse(chartMatch.Groups[1].Value);
            var allCharts = GetAllWordCharts();
            if (allCharts.Count == 0)
                throw new ArgumentException("No charts in this document");
            if (chartIdx < 1 || chartIdx > allCharts.Count)
                throw new ArgumentException($"Chart {chartIdx} not found (total: {allCharts.Count})");

            var chartInfo = allCharts[chartIdx - 1];

            // If series sub-path, prefix all properties with series{N}. for ChartSetter
            var chartProps = properties;
            var isSeriesPath = chartMatch.Groups[2].Success;
            if (isSeriesPath)
            {
                var seriesIdx = int.Parse(chartMatch.Groups[2].Value);
                chartProps = new Dictionary<string, string>();
                foreach (var (key, value) in properties)
                    chartProps[$"series{seriesIdx}.{key}"] = value;
            }

            // Chart-level position/size Set — mutate the hosting wp:inline's
            // wp:extent. Word inline charts have no positional x/y (they
            // flow in text), so only width/height are meaningful here.
            //
            // CONSISTENCY(chart-position-set): same vocabulary as Excel and
            // PPTX. x/y are silently dropped (flagged as unsupported) since
            // inline mode has no absolute position.
            if (!isSeriesPath && chartInfo.Inline != null)
            {
                ApplyWordChartPositionSet(chartInfo.Inline, chartProps, unsupported);
                // Drop ALL position keys (x/y/width/height) from chartProps
                // after handling — unsupported ones were already reported by
                // ApplyWordChartPositionSet. Forwarding them to ChartHelper
                // would double-report them.
                foreach (var k in new[] { "x", "y", "width", "height" })
                {
                    var matched = chartProps.Keys
                        .FirstOrDefault(key => key.Equals(k, StringComparison.OrdinalIgnoreCase));
                    if (matched != null) chartProps.Remove(matched);
                }
            }

            if (chartInfo.IsExtended)
            {
                // cx:chart — delegates to ChartExBuilder.SetChartProperties.
                // Same shared implementation as Excel/PPTX: title/axis/gridline
                // styling, series fill, histogram binning, etc.
                unsupported.AddRange(Core.ChartExBuilder.SetChartProperties(
                    chartInfo.ExtendedPart!, chartProps));
            }
            else
            {
                unsupported.AddRange(Core.ChartHelper.SetChartProperties(chartInfo.StandardPart!, chartProps));
            }
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

        // Section paths: /section[N] or /body/sectPr[N] (canonical form returned by Get/Query)
        var secSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^(?:/section\[(\d+)\]|/body/sectPr(?:\[(\d+)\])?)$");
        if (secSetMatch.Success)
        {
            var secIdxStr = secSetMatch.Groups[1].Success ? secSetMatch.Groups[1].Value
                : (secSetMatch.Groups[2].Success ? secSetMatch.Groups[2].Value : "1");
            var secIdx = int.Parse(secIdxStr);
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
                        EnsureSectPrPageSize(sectPr).Width = ParseTwips(value);
                        break;
                    case "pageheight" or "pageHeight":
                        EnsureSectPrPageSize(sectPr).Height = ParseTwips(value);
                        break;
                    case "orientation":
                    {
                        var ps = EnsureSectPrPageSize(sectPr);
                        var isLandscape = value.ToLowerInvariant() == "landscape";
                        ps.Orient = isLandscape
                            ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
                        // Default to A4 if no dimensions set
                        var w = ps.Width?.Value ?? WordPageDefaults.A4WidthTwips;
                        var h = ps.Height?.Value ?? WordPageDefaults.A4HeightTwips;
                        // Swap width/height if orientation changes and dimensions are misaligned
                        if ((isLandscape && w < h) || (!isLandscape && w > h))
                        {
                            ps.Width = h;
                            ps.Height = w;
                        }
                        break;
                    }
                    case "margintop":
                        EnsureSectPrPageMargin(sectPr).Top = (int)ParseTwips(value);
                        break;
                    case "marginbottom":
                        EnsureSectPrPageMargin(sectPr).Bottom = (int)ParseTwips(value);
                        break;
                    case "marginleft":
                        EnsureSectPrPageMargin(sectPr).Left = ParseTwips(value);
                        break;
                    case "marginright":
                        EnsureSectPrPageMargin(sectPr).Right = ParseTwips(value);
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
                    case "columnspace" or "columns.space":
                    {
                        // Standalone column-spacing update — preserves existing
                        // column count/widths. Pairs with the canonical 'columnSpace'
                        // key returned by Get/Query (WordHandler.Query.cs:491).
                        var spaceCols = EnsureColumns(sectPr);
                        spaceCols.Space = ParseTwips(value).ToString();
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
                    case "linenumbers" or "linenumbering":
                    {
                        var lower = value.ToLowerInvariant();
                        if (lower == "none" || lower == "off" || lower == "false")
                        {
                            sectPr.RemoveAllChildren<LineNumberType>();
                        }
                        else
                        {
                            var lnNum = sectPr.GetFirstChild<LineNumberType>();
                            if (lnNum == null)
                            {
                                lnNum = new LineNumberType();
                                sectPr.AppendChild(lnNum);
                            }
                            // If value is a number, set CountBy to that number
                            if (int.TryParse(lower, out var countBy))
                            {
                                lnNum.CountBy = (short)countBy;
                                lnNum.Restart = LineNumberRestartValues.Continuous;
                            }
                            else
                            {
                                lnNum.CountBy = 1;
                                lnNum.Restart = lower switch
                                {
                                    "continuous" => LineNumberRestartValues.Continuous,
                                    "restartpage" or "page" => LineNumberRestartValues.NewPage,
                                    "restartsection" or "section" => LineNumberRestartValues.NewSection,
                                    _ => LineNumberRestartValues.Continuous
                                };
                            }
                        }
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
            var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart
                ?? _doc.MainDocumentPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
            if (stylesPart.Styles == null) stylesPart.Styles = new Styles();
            var styles = stylesPart.Styles;
            var style = styles.Elements<Style>().FirstOrDefault(s =>
                s.StyleId?.Value == styleId || s.StyleName?.Val?.Value == styleId);
            if (style == null)
            {
                var isBuiltIn = styleId is "Normal" or "Heading1" or "Heading2" or "Heading3" or "Heading4"
                    or "Heading5" or "Heading6" or "Heading7" or "Heading8" or "Heading9"
                    or "Title" or "Subtitle" or "Quote" or "IntenseQuote" or "ListParagraph"
                    or "NoSpacing" or "TOCHeading";
                style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
                if (!isBuiltIn) style.CustomStyle = true;
                style.AppendChild(new StyleName { Val = styleId });
                styles.AppendChild(style);
            }

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
                        rPr2.FontSize = new FontSize { Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString() };
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
                        var ulVal = NormalizeUnderlineValue(value);
                        rPrU.Underline = new Underline { Val = new UnderlineValues(ulVal) };
                        break;
                    }
                    case "strike" or "strikethrough":
                        var rPrS = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        rPrS.Strike = IsTruthy(value) ? new Strike() : null;
                        break;
                    case "alignment":
                        var pPr = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                        pPr.Justification = new Justification { Val = ParseJustification(value) };
                        break;
                    case "spacebefore" or "spaceBefore":
                        var pPr2 = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                        var sp2 = pPr2.SpacingBetweenLines ?? (pPr2.SpacingBetweenLines = new SpacingBetweenLines());
                        sp2.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                        break;
                    case "spaceafter" or "spaceAfter":
                        var pPr3 = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                        var sp3 = pPr3.SpacingBetweenLines ?? (pPr3.SpacingBetweenLines = new SpacingBetweenLines());
                        sp3.After = SpacingConverter.ParseWordSpacing(value).ToString();
                        break;
                    case "linespacing" or "lineSpacing":
                    {
                        var pPr4 = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                        var sp4 = pPr4.SpacingBetweenLines ?? (pPr4.SpacingBetweenLines = new SpacingBetweenLines());
                        var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
                        sp4.Line = twips.ToString();
                        sp4.LineRule = isMultiplier
                            ? new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Auto)
                            : new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Exact);
                        break;
                    }
                    case "contextualspacing" or "contextualSpacing":
                    {
                        var pPrCs = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                        if (IsTruthy(value))
                            pPrCs.ContextualSpacing ??= new ContextualSpacing();
                        else
                            pPrCs.ContextualSpacing = null;
                        break;
                    }
                    case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
                    case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right":
                    {
                        var pPrB = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                        ApplyStyleParagraphBorders(pPrB, key, value);
                        break;
                    }
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            styles.Save();
            return unsupported;
        }

        // CONSISTENCY(ole-shorthand-set): mirror the /body/ole[N] shorthand
        // already supported in Get (WordHandler.Query.cs) and Remove
        // (WordHandler.Mutations.cs). Without this intercept, Set falls through
        // to NavigateToElement which hits "No ole found at /body" because OLE
        // lives inside a run, not as a direct child of the body.
        var wordOleSetMatch = System.Text.RegularExpressions.Regex.Match(
            path,
            @"^(?<parent>/body|/header\[\d+\]|/footer\[\d+\])?/(?:ole|object|embed)\[(?<idx>\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wordOleSetMatch.Success)
        {
            var wOleIdx = int.Parse(wordOleSetMatch.Groups["idx"].Value);
            var wOleParent = wordOleSetMatch.Groups["parent"].Success && wordOleSetMatch.Groups["parent"].Value.Length > 0
                ? wordOleSetMatch.Groups["parent"].Value
                : "/body";
            var allOles = Query("ole")
                .Where(n => n.Path.StartsWith(wOleParent + "/", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (wOleIdx < 1 || wOleIdx > allOles.Count)
                throw new ArgumentException(
                    $"OLE object {wOleIdx} not found at {wOleParent} (available: {allOles.Count}).");
            return Set(allOles[wOleIdx - 1].Path, properties);
        }

        var parts = ParsePath(path);
        var element = NavigateToElement(parts, out var ctx);
        if (element == null)
            throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));

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
                        // Clear showingPlaceholder flag so Word doesn't display as placeholder style
                        var plcHdr = (element is SdtBlock sb2 ? sb2.SdtProperties : ((SdtRun)element).SdtProperties)
                            ?.GetFirstChild<ShowingPlaceholder>();
                        plcHdr?.Remove();
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
                // CONSISTENCY(run-prop-helper): rPr-only props delegate to
                // ApplyRunLevelProperty so the per-property OOXML write logic
                // lives in one place; non-rPr cases (text content, image
                // swap, OLE resize, etc.) stay in the inline switch below.
                if (ApplyRunLevelProperty(EnsureRunProperties(run), key, value))
                    continue;
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        var textEl = run.GetFirstChild<Text>();
                        if (textEl != null) textEl.Text = value;
                        break;
                    case "alt" or "alttext" or "description":
                        var drawingAlt = run.GetFirstChild<Drawing>();
                        if (drawingAlt != null)
                        {
                            var docPropsAlt = drawingAlt.Descendants<DW.DocProperties>().FirstOrDefault();
                            if (docPropsAlt != null) docPropsAlt.Description = value;
                        }
                        else unsupported.Add(key);
                        break;
                    case "width":
                    {
                        var drawingW = run.GetFirstChild<Drawing>();
                        if (drawingW != null)
                        {
                            var extentW = drawingW.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentW != null) extentW.Cx = ParseEmu(value);
                            var extentsW = drawingW.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsW != null) extentsW.Cx = ParseEmu(value);
                            break;
                        }
                        // OLE run: update VML v:shape style.
                        var oleW = run.GetFirstChild<EmbeddedObject>();
                        var shapeW = oleW?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                        if (shapeW != null)
                        {
                            var styleAttrW = shapeW.GetAttributes().FirstOrDefault(a => a.LocalName == "style");
                            var currentStyleW = styleAttrW.Value ?? "";
                            var ptStrW = (ParseEmu(value) / 12700.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt";
                            var newStyleW = ReplaceVmlStyleDimension(currentStyleW, "width", ptStrW);
                            shapeW.SetAttribute(new OpenXmlAttribute("", "style", "", newStyleW));
                            break;
                        }
                        unsupported.Add(key);
                        break;
                    }
                    case "height":
                    {
                        var drawingH = run.GetFirstChild<Drawing>();
                        if (drawingH != null)
                        {
                            var extentH = drawingH.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentH != null) extentH.Cy = ParseEmu(value);
                            var extentsH = drawingH.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsH != null) extentsH.Cy = ParseEmu(value);
                            break;
                        }
                        // OLE run: update VML v:shape style.
                        var oleH = run.GetFirstChild<EmbeddedObject>();
                        var shapeH = oleH?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                        if (shapeH != null)
                        {
                            var styleAttrH = shapeH.GetAttributes().FirstOrDefault(a => a.LocalName == "style");
                            var currentStyleH = styleAttrH.Value ?? "";
                            var ptStrH = (ParseEmu(value) / 12700.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt";
                            var newStyleH = ReplaceVmlStyleDimension(currentStyleH, "height", ptStrH);
                            shapeH.SetAttribute(new OpenXmlAttribute("", "style", "", newStyleH));
                            break;
                        }
                        unsupported.Add(key);
                        break;
                    }
                    case "path" or "src":
                    {
                        // Replace image source in a run containing a Drawing
                        var drawingSrc = run.GetFirstChild<Drawing>();
                        var blip = drawingSrc?.Descendants<A.Blip>().FirstOrDefault();
                        if (blip != null)
                        {
                            var mainPartImg = _doc.MainDocumentPart!;
                            var (wordImgStream, imgType) = OfficeCli.Core.ImageSource.Resolve(value);
                            using var wordImgDispose = wordImgStream;

                            // Remove old image part(s) to avoid storage bloat —
                            // include the asvg:svgBlip extension part if the
                            // previous image was SVG, otherwise it would be
                            // orphaned in word/media/.
                            var oldEmbedId = blip.Embed?.Value;
                            if (oldEmbedId != null)
                            {
                                try { mainPartImg.DeletePart(oldEmbedId); } catch { }
                            }
                            var oldSvgRelId = OfficeCli.Core.SvgImageHelper.GetSvgRelId(blip);
                            if (oldSvgRelId != null)
                            {
                                try { mainPartImg.DeletePart(oldSvgRelId); } catch { }
                            }

                            if (imgType == ImagePartType.Svg)
                            {
                                // Match AddPicture: SVG part referenced via
                                // extension, raster fallback at r:embed.
                                using var svgBytes = new MemoryStream();
                                wordImgStream.CopyTo(svgBytes);
                                svgBytes.Position = 0;

                                var svgPart = mainPartImg.AddImagePart(ImagePartType.Svg);
                                svgPart.FeedData(svgBytes);
                                var newSvgRelId = mainPartImg.GetIdOfPart(svgPart);

                                var pngPart = mainPartImg.AddImagePart(ImagePartType.Png);
                                pngPart.FeedData(new MemoryStream(
                                    OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                                blip.Embed = mainPartImg.GetIdOfPart(pngPart);
                                OfficeCli.Core.SvgImageHelper.AppendSvgExtension(blip, newSvgRelId);
                            }
                            else
                            {
                                var newImgPart = mainPartImg.AddImagePart(imgType);
                                newImgPart.FeedData(wordImgStream);
                                blip.Embed = mainPartImg.GetIdOfPart(newImgPart);
                                // Drop the SVG extension if we replaced an SVG
                                // with a raster image; otherwise Word would
                                // keep rendering the stale SVG reference.
                                if (oldSvgRelId != null)
                                {
                                    var extLst = blip.GetFirstChild<A.BlipExtensionList>();
                                    if (extLst != null)
                                    {
                                        foreach (var ext in extLst.Elements<A.BlipExtension>().ToList())
                                        {
                                            if (string.Equals(ext.Uri?.Value,
                                                OfficeCli.Core.SvgImageHelper.SvgExtensionUri,
                                                StringComparison.OrdinalIgnoreCase))
                                                ext.Remove();
                                        }
                                        if (!extLst.Elements<A.BlipExtension>().Any())
                                            extLst.Remove();
                                    }
                                }
                            }
                            break;
                        }

                        // OLE case: run contains an EmbeddedObject. Replace
                        // the backing embedded part and (if needed) update
                        // the ProgID automatically from the new extension.
                        // This is the symmetric counterpart to AddOle — the
                        // part-cleanup rule from CLAUDE.md's Known API
                        // Quirks ("always delete old ImagePart to avoid
                        // storage bloat") applies equally to OLE payloads.
                        var ole = run.GetFirstChild<EmbeddedObject>();
                        if (ole != null)
                        {
                            var mainOle = _doc.MainDocumentPart!;
                            var oleEl = ole.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
                            if (oleEl != null)
                            {
                                var relAttr = oleEl.GetAttributes().FirstOrDefault(a => a.LocalName == "id"
                                    && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                                var oldRel = relAttr.Value;
                                if (!string.IsNullOrEmpty(oldRel))
                                {
                                    try { mainOle.DeletePart(oldRel); } catch { }
                                }
                                var (newEmbedRel, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(mainOle, value, _filePath);
                                // Update r:id attribute in place.
                                oleEl.SetAttribute(new OpenXmlAttribute("r", "id",
                                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships", newEmbedRel));
                                // Refresh ProgID if it wasn't explicitly pinned by the caller.
                                var newProgId = OfficeCli.Core.OleHelper.DetectProgId(value);
                                OfficeCli.Core.OleHelper.ValidateProgId(newProgId);
                                oleEl.SetAttribute(new OpenXmlAttribute("", "ProgID", "", newProgId));
                            }
                            break;
                        }
                        unsupported.Add(key);
                        break;
                    }
                    case "progid":
                    {
                        // Standalone ProgID override on an existing OLE run.
                        // Mirrors the ProgID-refresh in the "path"/"src" branch
                        // above, but without touching the backing embedded
                        // part. CONSISTENCY(ole-set-progid): PPT and Excel OLE
                        // Set both accept a bare progId key; Word must too.
                        var oleStandalone = run.GetFirstChild<EmbeddedObject>();
                        var oleElStandalone = oleStandalone?.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
                        if (oleElStandalone != null)
                        {
                            OfficeCli.Core.OleHelper.ValidateProgId(value);
                            oleElStandalone.SetAttribute(new OpenXmlAttribute("", "ProgID", "", value));
                            break;
                        }
                        unsupported.Add(key);
                        break;
                    }
                    case "display":
                    {
                        // Update DrawAspect attribute on o:OLEObject.
                        // Strict: only "icon" or "content" are accepted; any
                        // other value throws (see OleHelper.NormalizeOleDisplay).
                        // CONSISTENCY(ole-set-display): mirrors PPT ShowAsIcon toggle.
                        var normalized = OfficeCli.Core.OleHelper.NormalizeOleDisplay(value);
                        var oleDisplay = run.GetFirstChild<EmbeddedObject>();
                        var oleElDisplay = oleDisplay?.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
                        if (oleElDisplay != null)
                        {
                            var drawAspect = normalized == "content" ? "Content" : "Icon";
                            oleElDisplay.SetAttribute(new OpenXmlAttribute("", "DrawAspect", "", drawAspect));
                            break;
                        }
                        unsupported.Add(key);
                        break;
                    }
                    case "icon":
                    {
                        // Empty/whitespace value: treat as unsupported rather
                        // than feeding it into ImageSource.Resolve (which
                        // throws). Matches the gentler unsupported-key pattern
                        // used elsewhere in the Word Set OLE branch.
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            unsupported.Add(key);
                            break;
                        }
                        // Replace the v:imagedata r:id with a new ImagePart, and
                        // delete the old ImagePart to avoid storage bloat
                        // (mirrors Set src cleanup rule in CLAUDE.md Known
                        // API Quirks for picture/blip replacement).
                        var oleIcon = run.GetFirstChild<EmbeddedObject>();
                        var shapeIcon = oleIcon?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                        var imagedata = shapeIcon?.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
                        if (imagedata == null)
                        {
                            unsupported.Add(key);
                            break;
                        }
                        var mainIcon = _doc.MainDocumentPart!;
                        var oldIconRelAttr = imagedata.GetAttributes().FirstOrDefault(a => a.LocalName == "id"
                            && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        if (oldIconRelAttr.Value is string oldIconRel && !string.IsNullOrEmpty(oldIconRel))
                        {
                            try { mainIcon.DeletePart(oldIconRel); } catch { }
                        }
                        var (iconStream, iconPartType) = OfficeCli.Core.ImageSource.Resolve(value);
                        using var iconDispose = iconStream;
                        var newIconPart = mainIcon.AddImagePart(iconPartType);
                        newIconPart.FeedData(iconStream);
                        var newIconRel = mainIcon.GetIdOfPart(newIconPart);
                        imagedata.SetAttribute(new OpenXmlAttribute("r", "id",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships", newIconRel));
                        break;
                    }
                    case "wrap":
                    {
                        var anchor = ResolveRunAnchor(run);
                        if (anchor == null) { unsupported.Add(key); break; }
                        ReplaceWrapElement(anchor, value);
                        break;
                    }
                    case "hposition":
                    {
                        var anchor = ResolveRunAnchor(run);
                        var hPosEl = anchor?.GetFirstChild<DW.HorizontalPosition>();
                        if (hPosEl == null) { unsupported.Add(key); break; }
                        var emu = ParseEmu(value).ToString();
                        var offset = hPosEl.GetFirstChild<DW.PositionOffset>();
                        if (offset != null) offset.Text = emu;
                        else hPosEl.AppendChild(new DW.PositionOffset(emu));
                        break;
                    }
                    case "vposition":
                    {
                        var anchor = ResolveRunAnchor(run);
                        var vPosEl = anchor?.GetFirstChild<DW.VerticalPosition>();
                        if (vPosEl == null) { unsupported.Add(key); break; }
                        var emu = ParseEmu(value).ToString();
                        var offset = vPosEl.GetFirstChild<DW.PositionOffset>();
                        if (offset != null) offset.Text = emu;
                        else vPosEl.AppendChild(new DW.PositionOffset(emu));
                        break;
                    }
                    case "hrelative":
                    {
                        var anchor = ResolveRunAnchor(run);
                        var hPosEl = anchor?.GetFirstChild<DW.HorizontalPosition>();
                        if (hPosEl == null) { unsupported.Add(key); break; }
                        hPosEl.RelativeFrom = ParseHorizontalRelative(value);
                        break;
                    }
                    case "vrelative":
                    {
                        var anchor = ResolveRunAnchor(run);
                        var vPosEl = anchor?.GetFirstChild<DW.VerticalPosition>();
                        if (vPosEl == null) { unsupported.Add(key); break; }
                        vPosEl.RelativeFrom = ParseVerticalRelative(value);
                        break;
                    }
                    case "behindtext":
                    {
                        var anchor = ResolveRunAnchor(run);
                        if (anchor == null) { unsupported.Add(key); break; }
                        anchor.BehindDoc = value.Equals("true", StringComparison.OrdinalIgnoreCase);
                        break;
                    }
                    // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is `url`
                    // (per schemas/help/docx/hyperlink.json). `link` / `href` are
                    // accepted input aliases.
                    case "url":
                    case "href":
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
                    case "name":
                    {
                        // CONSISTENCY(ole-set-name): PPT OLE Set accepts a
                        // bare `name` key that writes oleObj.Name. Word does
                        // not have an equivalent attribute on o:OLEObject
                        // (the VML CT_OleObject complex type has no Name),
                        // so we store the friendly name on the surrounding
                        // v:shape element's "alt" attribute. AddOle writes
                        // to the same attribute and CreateOleNode reads it
                        // back into Format["name"].
                        var oleName = run.GetFirstChild<EmbeddedObject>();
                        var shapeNameEl = oleName?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                        if (shapeNameEl != null)
                        {
                            shapeNameEl.SetAttribute(new OpenXmlAttribute("", "alt", "", value));
                            break;
                        }
                        unsupported.Add(key);
                        break;
                    }
                    default:
                        // OLE runs use a slim prop vocabulary (src, progId,
                        // width, height, alt) that doesn't overlap the rich
                        // run-formatting hint suffix. Emit bare keys to match
                        // PPT/Excel OLE Set. CONSISTENCY(ole-set-bare-key).
                        if (run.GetFirstChild<EmbeddedObject>() != null)
                        {
                            unsupported.Add(key);
                        }
                        else if (!GenericXmlQuery.TryCreateTypedChild(EnsureRunProperties(run), key, value))
                        {
                            unsupported.Add(unsupported.Count == 0
                                ? $"{key} (valid run props: text, bold, italic, font, size, color, underline, strike, highlight, caps, smallcaps, superscript, subscript, shading, link, formula)"
                                : key);
                        }
                        break;
                }
            }
        }
        else if (element is Hyperlink hl)
        {
            foreach (var (key, value) in properties)
            {
                var k = key.ToLowerInvariant();
                switch (k)
                {
                    case "url":
                    case "link":
                    case "href":
                    {
                        var mainPartHl = _doc.MainDocumentPart!;
                        // Delete old relationship to avoid storage bloat
                        var oldRelId = hl.Id?.Value;
                        if (oldRelId != null)
                        {
                            var oldRel = mainPartHl.HyperlinkRelationships.FirstOrDefault(r => r.Id == oldRelId);
                            if (oldRel != null)
                                mainPartHl.DeleteReferenceRelationship(oldRel);
                        }
                        var uri = Uri.TryCreate(value, UriKind.Absolute, out var absUri)
                            ? absUri
                            : new Uri(value, UriKind.Relative);
                        var newRelId = mainPartHl.AddHyperlinkRelationship(uri, isExternal: true).Id;
                        hl.Id = newRelId;
                        break;
                    }
                    case "text":
                    {
                        // Update text in all runs within the hyperlink
                        var runs = hl.Elements<Run>().ToList();
                        if (runs.Count > 0)
                        {
                            // Set text on the first run, remove the rest
                            var firstRun = runs[0];
                            var t = firstRun.GetFirstChild<Text>()
                                ?? firstRun.AppendChild(new Text());
                            t.Text = value;
                            t.Space = SpaceProcessingModeValues.Preserve;
                            for (int i = 1; i < runs.Count; i++)
                                runs[i].Remove();
                        }
                        else
                        {
                            // No runs yet, create one
                            var newRun = new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                            hl.AppendChild(newRun);
                        }
                        break;
                    }
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is M.Paragraph mPara)
        {
            foreach (var (key, value) in properties)
            {
                var k = key.ToLowerInvariant();
                switch (k)
                {
                    case "formula":
                    {
                        // Clear existing oMath children and rebuild from new formula
                        foreach (var child in mPara.ChildElements.ToList())
                            child.Remove();
                        var mathContent = FormulaParser.Parse(value);
                        M.OfficeMath oMath = mathContent is M.OfficeMath dm
                            ? dm : new M.OfficeMath(mathContent.CloneNode(true));
                        mPara.AppendChild(oMath);
                        break;
                    }
                    default:
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid equation props: formula)"
                            : key);
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
                        // Set text on paragraph: update first run or create one.
                        // CONSISTENCY(text-breaks): route through AppendTextWithBreaks
                        // so \n/\t in value become <w:br/>/<w:tab/>, matching Add behavior.
                        var existingRuns = para.Elements<Run>().ToList();
                        if (existingRuns.Count > 0)
                        {
                            // Preserve RunProperties from first run, drop all prior text/break/tab children.
                            var keepRun = existingRuns[0];
                            var keepRProps = keepRun.RunProperties;
                            keepRun.RemoveAllChildren();
                            if (keepRProps != null)
                                keepRun.AppendChild(keepRProps);
                            AppendTextWithBreaks(keepRun, value);
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
                            AppendTextWithBreaks(newRun, value);
                            para.AppendChild(newRun);
                        }
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(pProps, key, value))
                            unsupported.Add(unsupported.Count == 0
                                ? $"{key} (valid paragraph props: text, style, alignment, bold, italic, font, size, color, spaceBefore, spaceAfter, lineSpacing, indent, liststyle, formula)"
                                : key);
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
                                        rPr.FontSize = new FontSize { Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString() };
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
                                        var ulVal = NormalizeUnderlineValue(value);
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
                                    InsertRunPropInSchemaOrder(pmrp, new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value });
                                    break;
                                case "size":
                                    pmrp.RemoveAllChildren<FontSize>();
                                    InsertRunPropInSchemaOrder(pmrp, new FontSize { Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString() });
                                    break;
                                case "bold":
                                    pmrp.RemoveAllChildren<Bold>();
                                    if (IsTruthy(value)) InsertRunPropInSchemaOrder(pmrp, new Bold());
                                    break;
                                case "italic":
                                    pmrp.RemoveAllChildren<Italic>();
                                    if (IsTruthy(value)) InsertRunPropInSchemaOrder(pmrp, new Italic());
                                    break;
                                case "color":
                                    pmrp.RemoveAllChildren<Color>();
                                    InsertRunPropInSchemaOrder(pmrp, new Color { Val = SanitizeHex(value) });
                                    break;
                                case "highlight":
                                    pmrp.RemoveAllChildren<Highlight>();
                                    InsertRunPropInSchemaOrder(pmrp, new Highlight { Val = ParseHighlightColor(value) });
                                    break;
                                case "underline":
                                {
                                    var ulVal = NormalizeUnderlineValue(value);
                                    pmrp.RemoveAllChildren<Underline>();
                                    InsertRunPropInSchemaOrder(pmrp, new Underline { Val = new UnderlineValues(ulVal) });
                                    break;
                                }
                                case "strike":
                                    pmrp.RemoveAllChildren<Strike>();
                                    if (IsTruthy(value)) InsertRunPropInSchemaOrder(pmrp, new Strike());
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
                            // Validate color positions don't look like numbers (likely swapped with angle)
                            if (int.TryParse(shdParts[1], out _) && shdParts[1].Length <= 3)
                                throw new ArgumentException($"'{shdParts[1]}' looks like an angle, not a color. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                            if (int.TryParse(shdParts[2], out _) && shdParts[2].Length <= 3)
                                throw new ArgumentException($"'{shdParts[2]}' looks like an angle, not a color. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                            int angleDeg = 180;
                            if (shdParts.Length >= 4)
                            {
                                if (!int.TryParse(shdParts[3], out angleDeg))
                                    throw new ArgumentException($"Invalid gradient angle '{shdParts[3]}', expected integer. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
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
                                var cellPat = shdParts[0].TrimStart('#');
                                if (cellPat.Length >= 6 && cellPat.All(char.IsAsciiHexDigit))
                                { shd.Val = ShadingPatternValues.Clear; shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[0]).Rgb; }
                                else
                                {
                                    WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                                    shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[1]).Rgb;
                                    if (shdParts.Length >= 3) shd.Color = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[2]).Rgb;
                                }
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
                    case "gridspan" or "colspan":
                        var newSpan = ParseHelpers.SafeParseInt(value, "gridspan");
                        if (newSpan <= 0)
                            throw new ArgumentException($"Invalid 'gridspan' value: '{value}'. Must be a positive integer (> 0).");
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
                    case "fittext":
                    {
                        // FitText goes on w:rPr (RunProperties), not tcPr
                        var cellWidth = tcPr.TableCellWidth?.Width?.Value;
                        var fitVal = cellWidth != null && uint.TryParse(cellWidth, out var fw) ? fw : 0u;
                        foreach (var cellPara in cell.Elements<Paragraph>())
                        {
                            foreach (var cellRun in cellPara.Elements<Run>())
                            {
                                var rPr = EnsureRunProperties(cellRun);
                                rPr.RemoveAllChildren<FitText>();
                                if (IsTruthy(value))
                                    rPr.AppendChild(new FitText { Val = fitVal });
                            }
                            // Also apply to ParagraphMarkRunProperties
                            var pPr = cellPara.ParagraphProperties;
                            if (pPr?.ParagraphMarkRunProperties != null)
                            {
                                pPr.ParagraphMarkRunProperties.RemoveAllChildren<FitText>();
                                if (IsTruthy(value))
                                    pPr.ParagraphMarkRunProperties.AppendChild(new FitText { Val = fitVal });
                            }
                        }
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tcPr, key, value))
                            unsupported.Add(unsupported.Count == 0
                                ? $"{key} (valid cell props: text, font, size, bold, italic, color, alignment, valign, width, shd, border, colspan, fitText, textDirection, nowrap, padding)"
                                : key);
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
                        // c1, c2, ... shorthand: set text of specific cell by index
                        if (key.Length >= 2 && key[0] == 'c' && int.TryParse(key.AsSpan(1), out var cIdx))
                        {
                            var rowCells = row.Elements<TableCell>().ToList();
                            if (cIdx < 1 || cIdx > rowCells.Count)
                                throw new ArgumentException($"Cell c{cIdx} out of range (row has {rowCells.Count} cells)");
                            var targetPara = rowCells[cIdx - 1].GetFirstChild<Paragraph>()
                                ?? rowCells[cIdx - 1].AppendChild(new Paragraph());
                            targetPara.RemoveAllChildren<Run>();
                            if (!string.IsNullOrEmpty(value))
                                targetPara.AppendChild(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                        }
                        else if (!GenericXmlQuery.TryCreateTypedChild(trPr, key, value))
                            unsupported.Add(unsupported.Count == 0
                                ? $"{key} (valid row props: height, height.exact, header, c1, c2, ...)"
                                : key);
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
                    case "firstrow":
                    case "lastrow":
                    case "firstcol" or "firstcolumn":
                    case "lastcol" or "lastcolumn":
                    case "bandrow" or "bandedrows" or "bandrows":
                    case "bandcol" or "bandedcols" or "bandcols":
                    {
                        var tblLook = tblPr.GetFirstChild<TableLook>();
                        if (tblLook == null) { tblLook = new TableLook { Val = "04A0" }; tblPr.AppendChild(tblLook); }
                        var bv = IsTruthy(value);
                        switch (key.ToLowerInvariant())
                        {
                            case "firstrow": tblLook.FirstRow = bv; break;
                            case "lastrow": tblLook.LastRow = bv; break;
                            case "firstcol" or "firstcolumn": tblLook.FirstColumn = bv; break;
                            case "lastcol" or "lastcolumn": tblLook.LastColumn = bv; break;
                            case "bandrow" or "bandedrows" or "bandrows": tblLook.NoHorizontalBand = !bv; break;
                            case "bandcol" or "bandedcols" or "bandcols": tblLook.NoVerticalBand = !bv; break;
                        }
                        break;
                    }
                    case "position" or "floating":
                    {
                        // Shorthand: "floating" or "none" to toggle floating table
                        if (value.Equals("none", StringComparison.OrdinalIgnoreCase)
                            || value.Equals("false", StringComparison.OrdinalIgnoreCase))
                        {
                            tblPr.RemoveAllChildren<TablePositionProperties>();
                            tblPr.RemoveAllChildren<TableOverlap>();
                        }
                        else
                        {
                            // "floating" enables floating with defaults
                            var tpp = tblPr.GetFirstChild<TablePositionProperties>();
                            if (tpp == null)
                            {
                                tpp = new TablePositionProperties();
                                tblPr.AppendChild(tpp);
                            }
                            if (tpp.VerticalAnchor == null)
                                tpp.VerticalAnchor = VerticalAnchorValues.Page;
                            if (tpp.HorizontalAnchor == null)
                                tpp.HorizontalAnchor = HorizontalAnchorValues.Page;
                        }
                        break;
                    }
                    case "position.x" or "tblpx":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        var v = value.ToLowerInvariant();
                        if (v is "left" or "center" or "right" or "inside" or "outside")
                        {
                            tpp.TablePositionXAlignment = v switch
                            {
                                "left" => HorizontalAlignmentValues.Left,
                                "center" => HorizontalAlignmentValues.Center,
                                "right" => HorizontalAlignmentValues.Right,
                                "inside" => HorizontalAlignmentValues.Inside,
                                "outside" => HorizontalAlignmentValues.Outside,
                                _ => throw new ArgumentException($"Invalid position.x alignment: '{value}'")
                            };
                            tpp.TablePositionX = null;
                        }
                        else
                        {
                            tpp.TablePositionX = (int)ParseTwips(value);
                            tpp.TablePositionXAlignment = null;
                        }
                        break;
                    }
                    case "position.y" or "tblpy":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        var v = value.ToLowerInvariant();
                        if (v is "top" or "center" or "bottom" or "inside" or "outside")
                        {
                            tpp.TablePositionYAlignment = v switch
                            {
                                "top" => VerticalAlignmentValues.Top,
                                "center" => VerticalAlignmentValues.Center,
                                "bottom" => VerticalAlignmentValues.Bottom,
                                "inside" => VerticalAlignmentValues.Inside,
                                "outside" => VerticalAlignmentValues.Outside,
                                _ => throw new ArgumentException($"Invalid position.y alignment: '{value}'")
                            };
                            tpp.TablePositionY = null;
                        }
                        else
                        {
                            tpp.TablePositionY = (int)ParseTwips(value);
                            tpp.TablePositionYAlignment = null;
                        }
                        break;
                    }
                    case "position.hanchor" or "position.horizontalanchor":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        tpp.HorizontalAnchor = value.ToLowerInvariant() switch
                        {
                            "margin" => HorizontalAnchorValues.Margin,
                            "page" => HorizontalAnchorValues.Page,
                            "text" => HorizontalAnchorValues.Text,
                            _ => throw new ArgumentException($"Invalid horizontalAnchor: '{value}'. Valid: margin, page, text.")
                        };
                        break;
                    }
                    case "position.vanchor" or "position.verticalanchor":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        tpp.VerticalAnchor = value.ToLowerInvariant() switch
                        {
                            "margin" => VerticalAnchorValues.Margin,
                            "page" => VerticalAnchorValues.Page,
                            "text" => VerticalAnchorValues.Text,
                            _ => throw new ArgumentException($"Invalid verticalAnchor: '{value}'. Valid: margin, page, text.")
                        };
                        break;
                    }
                    case "position.leftfromtext" or "position.left":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        tpp.LeftFromText = (short)ParseTwips(value);
                        break;
                    }
                    case "position.rightfromtext" or "position.right":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        tpp.RightFromText = (short)ParseTwips(value);
                        break;
                    }
                    case "position.topfromtext" or "position.top":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        tpp.TopFromText = (short)ParseTwips(value);
                        break;
                    }
                    case "position.bottomfromtext" or "position.bottom":
                    {
                        var tpp = EnsureTablePositionProperties(tblPr);
                        tpp.BottomFromText = (short)ParseTwips(value);
                        break;
                    }
                    case "overlap":
                    {
                        tblPr.RemoveAllChildren<TableOverlap>();
                        if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        {
                            var overlapEl = new TableOverlap
                            {
                                Val = value.ToLowerInvariant() switch
                                {
                                    "overlap" or "true" or "always" => TableOverlapValues.Overlap,
                                    "never" or "false" => TableOverlapValues.Never,
                                    _ => throw new ArgumentException($"Invalid overlap: '{value}'. Valid: overlap, never, none.")
                                }
                            };
                            // CT_TblPr schema: tblStyle → tblpPr → tblOverlap → ...
                            var tppRef = tblPr.GetFirstChild<TablePositionProperties>();
                            if (tppRef != null) tppRef.InsertAfterSelf(overlapEl);
                            else
                            {
                                var styleRef = tblPr.GetFirstChild<TableStyle>();
                                if (styleRef != null) styleRef.InsertAfterSelf(overlapEl);
                                else tblPr.PrependChild(overlapEl);
                            }
                        }
                        break;
                    }
                    case "caption":
                        tblPr.RemoveAllChildren<TableCaption>();
                        if (!string.IsNullOrEmpty(value))
                            tblPr.AppendChild(new TableCaption { Val = value });
                        break;
                    case "description":
                        tblPr.RemoveAllChildren<TableDescription>();
                        if (!string.IsNullOrEmpty(value))
                            tblPr.AppendChild(new TableDescription { Val = value });
                        break;
                    case var k when k.StartsWith("border"):
                        ApplyTableBorders(tblPr, key, value);
                        break;
                    case "colwidths" or "colWidths":
                    {
                        var parts = value.Split(',');
                        var tblGrid = tbl.GetFirstChild<TableGrid>();
                        if (tblGrid == null)
                        {
                            tblGrid = new TableGrid();
                            tbl.InsertAfter(tblGrid, tblPr);
                        }
                        var gridCols = tblGrid.Elements<GridColumn>().ToList();
                        for (int ci = 0; ci < parts.Length; ci++)
                        {
                            var twips = ParseTwips(parts[ci].Trim());
                            if (ci < gridCols.Count)
                                gridCols[ci].Width = twips.ToString();
                            else
                                tblGrid.AppendChild(new GridColumn { Width = twips.ToString() });
                            // Also update cell widths in each row for this column
                            foreach (var tblRow in tbl.Elements<TableRow>())
                            {
                                var cells = tblRow.Elements<TableCell>().ToList();
                                if (ci < cells.Count)
                                {
                                    var tcPr = cells[ci].TableCellProperties ?? cells[ci].PrependChild(new TableCellProperties());
                                    tcPr.TableCellWidth = new TableCellWidth { Width = twips.ToString(), Type = TableWidthUnitValues.Dxa };
                                }
                            }
                        }
                        break;
                    }
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tblPr, key, value))
                            unsupported.Add(unsupported.Count == 0
                                ? $"{key} (valid table props: width, alignment, style, indent, cellspacing, layout, padding, border*, colWidths, firstRow, lastRow, firstCol, lastCol, bandedRows, bandedCols, caption, description)"
                                : key);
                        break;
                }
            }
        }

        // Refresh w14:textId on the affected paragraph (content changed)
        var affectedPara = element as Paragraph ?? element.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();

        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private void SetHeaderFooter(string kind, int index, Dictionary<string, string> properties, List<string> unsupported)
    {
        var mainPart = _doc.MainDocumentPart!;
        OpenXmlCompositeElement? container;
        OpenXmlPart partRef;

        if (kind == "header")
        {
            var part = mainPart.HeaderParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Header not found: /header[{index + 1}]");
            container = part.Header;
            partRef = part;
        }
        else
        {
            var part = mainPart.FooterParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Footer not found: /footer[{index + 1}]");
            container = part.Footer;
            partRef = part;
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
                case "type":
                {
                    // Mutate the HeaderReference/FooterReference Type attribute
                    // pointing at this part. Read side (WordHandler.Query.cs:660-666,
                    // 717-723) only inspects body-level SectionProperties, so the
                    // write side stays scoped to the same set for round-trip parity.
                    var newType = value.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        "default" => HeaderFooterValues.Default,
                        _ => throw new ArgumentException(
                            $"Invalid {kind} type: '{value}'. Valid values: default, first, even.")
                    };
                    var partRid = mainPart.GetIdOfPart(partRef);
                    var body = mainPart.Document?.Body
                        ?? throw new InvalidOperationException("Document body not found");
                    bool found = false;
                    foreach (var sp in body.Elements<SectionProperties>())
                    {
                        if (kind == "header")
                        {
                            var ownRef = sp.Elements<HeaderReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<HeaderReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Header of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        else
                        {
                            var ownRef = sp.Elements<FooterReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<FooterReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Footer of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        // Mirrors AddHeader: Title-page header requires <w:titlePg/> on the section.
                        if (newType == HeaderFooterValues.First && sp.GetFirstChild<TitlePage>() == null)
                            sp.AddChild(new TitlePage(), throwOnError: false);
                    }
                    if (!found) unsupported.Add(key);
                    break;
                }
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
    /// <summary>Insert StyleParagraphProperties before StyleRunProperties to maintain OOXML schema order.</summary>
    private static StyleParagraphProperties EnsureStyleParagraphProperties(Style style)
    {
        var pPr = new StyleParagraphProperties();
        var rPr = style.StyleRunProperties;
        if (rPr != null)
            style.InsertBefore(pPr, rPr);
        else
            style.AppendChild(pPr);
        return pPr;
    }

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
                throw new ArgumentException($"Invalid border size '{parts[1]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
        }
        else
            size = style == BorderValues.Nil ? 0u : style == BorderValues.Thick ? 12u : 4u;
        string? color = parts.Length > 2 ? SanitizeHex(parts[2]) : null;
        uint space = 0u;
        if (parts.Length > 3 && !uint.TryParse(parts[3], out space))
            throw new ArgumentException($"Invalid border space '{parts[3]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
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
    private static bool ApplyParagraphLevelProperty(ParagraphProperties pProps, string key, string? value)
    {
        if (value is null) return false;
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
                // Lenient input: accept "2cm", "0.5in", "18pt", or bare twips.
                indent.FirstLine = SpacingConverter.ParseWordSpacing(value).ToString();
                indent.Hanging = null;
                return true;
            case "leftindent" or "indentleft" or "indent":
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
            case "keepnext" or "keepwithnext":
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
            case "widowcontrol" or "widoworphan":
                if (IsTruthy(value)) pProps.WidowControl ??= new WidowControl();
                else pProps.WidowControl = new WidowControl { Val = false };
                return true;
            case "contextualspacing" or "contextualSpacing":
                if (IsTruthy(value)) pProps.ContextualSpacing ??= new ContextualSpacing();
                else pProps.ContextualSpacing = null;
                return true;
            case "shading" or "shd":
                pProps.Shading = ParseShadingValue(value);
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
            case "numId" or "numid":
                var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr.NumberingId = new NumberingId { Val = ParseHelpers.SafeParseInt(value, "numId") };
                return true;
            case "numLevel" or "numlevel" or "ilvl":
                var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                numPr2.NumberingLevelReference = new NumberingLevelReference { Val = ParseHelpers.SafeParseInt(value, "numLevel") };
                return true;
            case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
            case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right":
                ApplyParagraphBorders(pProps, key, value);
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Parse a w:shd value string ("fill", "val;fill", "val;fill;color") into a Shading element.
    /// Shared by paragraph-level and run-level shading handlers.
    /// </summary>
    private static Shading ParseShadingValue(string value)
    {
        var shdParts = value.Split(';');
        var shd = new Shading();
        if (shdParts.Length == 1)
        {
            shd.Val = ShadingPatternValues.Clear;
            shd.Fill = SanitizeHex(shdParts[0]);
        }
        else if (shdParts.Length >= 2)
        {
            var firstAsHex = shdParts[0].TrimStart('#');
            if (firstAsHex.Length >= 6 && firstAsHex.All(char.IsAsciiHexDigit))
            {
                shd.Val = ShadingPatternValues.Clear;
                shd.Fill = SanitizeHex(shdParts[0]);
            }
            else
            {
                WarnIfShadingOrderWrong(shdParts[0]);
                shd.Val = new ShadingPatternValues(shdParts[0]);
                shd.Fill = SanitizeHex(shdParts[1]);
                if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
            }
        }
        return shd;
    }

    /// <summary>
    /// Apply a run-level (rPr) property by key. Returns true if handled, false if the key
    /// isn't a run-property (caller should fall through to its own switch). Mirrors
    /// ApplyParagraphLevelProperty in shape and is reused across the Run / TableCell-run /
    /// Style-run code paths so the per-property OOXML write logic lives in exactly one place.
    /// </summary>
    private static bool ApplyRunLevelProperty(RunProperties rPr, string key, string? value)
    {
        if (value is null) return false;
        switch (key.ToLowerInvariant())
        {
            case "bold":      rPr.Bold = IsTruthy(value) ? new Bold() : null; return true;
            case "italic":    rPr.Italic = IsTruthy(value) ? new Italic() : null; return true;
            case "caps":      rPr.Caps = IsTruthy(value) ? new Caps() : null; return true;
            case "smallcaps": rPr.SmallCaps = IsTruthy(value) ? new SmallCaps() : null; return true;
            case "dstrike":   rPr.DoubleStrike = IsTruthy(value) ? new DoubleStrike() : null; return true;
            case "vanish":    rPr.Vanish = IsTruthy(value) ? new Vanish() : null; return true;
            case "outline":   rPr.Outline = IsTruthy(value) ? new Outline() : null; return true;
            case "shadow":    rPr.Shadow = IsTruthy(value) ? new Shadow() : null; return true;
            case "emboss":    rPr.Emboss = IsTruthy(value) ? new Emboss() : null; return true;
            case "imprint":   rPr.Imprint = IsTruthy(value) ? new Imprint() : null; return true;
            case "noproof":   rPr.NoProof = IsTruthy(value) ? new NoProof() : null; return true;
            case "rtl":       rPr.RightToLeftText = IsTruthy(value) ? new RightToLeftText() : null; return true;
            case "strike":    rPr.Strike = IsTruthy(value) ? new Strike() : null; return true;
            case "font":
            {
                var fonts = rPr.RunFonts;
                if (fonts != null)
                {
                    fonts.Ascii = value;
                    fonts.HighAnsi = value;
                    fonts.EastAsia = value;
                }
                else
                {
                    rPr.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                }
                return true;
            }
            case "size":
                rPr.FontSize = new FontSize
                {
                    Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString()
                };
                return true;
            case "highlight":
                rPr.Highlight = new Highlight { Val = ParseHighlightColor(value) };
                return true;
            case "color":
                rPr.Color = new Color { Val = SanitizeHex(value) };
                return true;
            case "underline":
                rPr.Underline = new Underline { Val = new UnderlineValues(NormalizeUnderlineValue(value)) };
                return true;
            case "superscript":
                rPr.VerticalTextAlignment = IsTruthy(value)
                    ? new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }
                    : null;
                return true;
            case "subscript":
                rPr.VerticalTextAlignment = IsTruthy(value)
                    ? new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }
                    : null;
                return true;
            case "charspacing" or "letterspacing" or "spacing":
            {
                var csVal = value.TrimEnd();
                double csPt = csVal.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? ParseHelpers.SafeParseDouble(csVal[..^2], "charspacing")
                    : ParseHelpers.SafeParseDouble(csVal, "charspacing");
                rPr.Spacing = new Spacing { Val = (int)Math.Round(csPt * 20, MidpointRounding.AwayFromZero) };
                return true;
            }
            case "shading" or "shd":
                rPr.Shading = ParseShadingValue(value);
                return true;
            default:
                return false;
        }
    }

    private static void ApplyParagraphBorders(ParagraphProperties pProps, string key, string value)
    {
        var borders = pProps.ParagraphBorders;
        if (borders == null)
        {
            borders = new ParagraphBorders();
            pProps.ParagraphBorders = borders; // typed setter maintains CT_PPr schema order
        }
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

    private static void ApplyStyleParagraphBorders(StyleParagraphProperties spPr, string key, string value)
    {
        var borders = spPr.GetFirstChild<ParagraphBorders>();
        if (borders == null)
        {
            borders = new ParagraphBorders();
            // StyleParagraphProperties is also OneSequence — use SetElement pattern
            // ParagraphBorders element order index is after Indentation and before Shading
            var afterRef = (OpenXmlElement?)spPr.GetFirstChild<Indentation>()
                ?? (OpenXmlElement?)spPr.GetFirstChild<SpacingBetweenLines>()
                ?? (OpenXmlElement?)spPr.GetFirstChild<Justification>();
            if (afterRef != null)
                spPr.InsertAfter(borders, afterRef);
            else
                spPr.PrependChild(borders);
        }
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
    private static TablePositionProperties EnsureTablePositionProperties(TableProperties tblPr)
    {
        var tpp = tblPr.GetFirstChild<TablePositionProperties>();
        if (tpp == null)
        {
            tpp = new TablePositionProperties
            {
                VerticalAnchor = VerticalAnchorValues.Page,
                HorizontalAnchor = HorizontalAnchorValues.Page
            };
            // CT_TblPr schema order: tblStyle → tblpPr → tblOverlap → ...
            var tblStyle = tblPr.GetFirstChild<TableStyle>();
            if (tblStyle != null)
                tblStyle.InsertAfterSelf(tpp);
            else
                tblPr.PrependChild(tpp);
        }
        return tpp;
    }

    internal static uint ParseTwips(string value)
    {
        value = value.Trim();
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (cm)");
            return (uint)Math.Round(num * 1440.0 / 2.54);
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
