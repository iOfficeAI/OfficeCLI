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

// Per-path-pattern Set helpers extracted from the WordHandler.Set() entry
// method. Each helper owns one path-pattern's full handling. Mechanically
// extracted, no behavior change.
public partial class WordHandler
{
    private List<string> SetWatermarkPath(Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
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
                                @"rotation\s*:\s*-?\d+(?:\.\d+)?", $@"rotation:{value}");
                            break;
                        case "size":
                            // BUG-R36-B3: font-size on the v:textpath. Accept bare or pt-suffixed.
                            var sz = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ? value : value + "pt";
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"font-size:[^;""]+", $@"font-size:{sz}");
                            break;
                        case "width":
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"width:[^;""]+", $@"width:{value}");
                            break;
                        case "height":
                            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                                @"height:[^;""]+", $@"height:{value}");
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

    private List<string> SetChartAxisPath(System.Text.RegularExpressions.Match chartAxisSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
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

    private List<string> SetChartPath(System.Text.RegularExpressions.Match chartMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
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

    private List<string> SetFieldPath(System.Text.RegularExpressions.Match fieldSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var fieldIdx = int.Parse(fieldSetMatch.Groups[1].Value);
        var allFields = FindFields();
        if (fieldIdx < 1 || fieldIdx > allFields.Count)
            throw new ArgumentException($"Field {fieldIdx} not found (total: {allFields.Count})");

        var field = allFields[fieldIdx - 1];

        // CONSISTENCY(field-set-instruction-rewrite): support the same
        // high-level keys Add accepts (fieldType, name, format) by rewriting
        // the field instruction. schemas/help/docx/field.json advertises
        // [add/set/get] for these keys; previously Set rejected them as
        // UNSUPPORTED. We rewrite the instruction code in-place so the field
        // updates on next Word open (Dirty=true also auto-set).
        var rewriteFieldType = properties.GetValueOrDefault("fieldType")
            ?? properties.GetValueOrDefault("fieldtype")
            ?? properties.GetValueOrDefault("type");
        // CONSISTENCY(canonical-keys): mirror AddField's per-fieldType
        // alias chain (field.json declares all of these as set:true).
        // R6 added bookmarkName/styleName/propertyName/etc. on the Add
        // side; Set was rejecting them as unsupported until Round 9.
        var rewriteName = properties.GetValueOrDefault("name")
            ?? properties.GetValueOrDefault("fieldName")
            ?? properties.GetValueOrDefault("fieldname")
            ?? properties.GetValueOrDefault("bookmarkName")
            ?? properties.GetValueOrDefault("bookmarkname")
            ?? properties.GetValueOrDefault("bookmark")
            ?? properties.GetValueOrDefault("styleName")
            ?? properties.GetValueOrDefault("stylename")
            ?? properties.GetValueOrDefault("propertyName")
            ?? properties.GetValueOrDefault("propertyname");
        var hasRewriteFormat = properties.TryGetValue("format", out var rewriteFormat);
        // Accept both bare value (`M/d/yyyy`) and full-switch form (`\@ "M/d/yyyy"`).
        // The case-builder below always wraps effFormat in `\@ "..."`, so a user-supplied
        // \@ prefix would land as `\@ "\@ "M/d/yyyy""`. Strip the prefix + surrounding
        // whitespace + outer quotes so both input shapes produce the same output.
        if (hasRewriteFormat && rewriteFormat != null)
        {
            var fmt = rewriteFormat.Trim();
            if (fmt.StartsWith("\\@", StringComparison.Ordinal))
                fmt = fmt[2..].Trim().Trim('"');
            rewriteFormat = fmt;
        }
        if (!string.IsNullOrEmpty(rewriteFieldType) || !string.IsNullOrEmpty(rewriteName) || hasRewriteFormat)
        {
            var existingInstr = field.InstrCode.Text ?? "";
            // Decide effective field type: prefer explicit fieldType, else
            // sniff first token from existing instruction.
            string effType;
            if (!string.IsNullOrEmpty(rewriteFieldType))
            {
                effType = rewriteFieldType.ToUpperInvariant() switch
                {
                    "PAGENUM" or "PAGENUMBER" => "PAGE",
                    var t => t
                };
            }
            else
            {
                var trimmed = existingInstr.Trim();
                var firstSpace = trimmed.IndexOf(' ');
                effType = (firstSpace > 0 ? trimmed[..firstSpace] : trimmed).ToUpperInvariant();
            }

            // Sniff existing name (token after the field type) when not supplied
            string? effName = rewriteName;
            if (string.IsNullOrEmpty(effName))
            {
                var parts = existingInstr.Trim().Split(' ', 3, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2 && !parts[1].StartsWith("\\"))
                    effName = parts[1].Trim('"');
            }

            // Sniff existing \@ "..." format switch when not supplied
            string? effFormat = hasRewriteFormat ? rewriteFormat : null;
            if (!hasRewriteFormat)
            {
                var fmtMatch = System.Text.RegularExpressions.Regex.Match(existingInstr, "\\\\@\\s+\"([^\"]+)\"");
                if (fmtMatch.Success) effFormat = fmtMatch.Groups[1].Value;
            }

            string newInstr = effType switch
            {
                "PAGE" or "NUMPAGES" or "SECTION" or "SECTIONPAGES"
                or "AUTHOR" or "TITLE" or "SUBJECT" or "FILENAME"
                or "EDITTIME" or "REVNUM" or "TEMPLATE" or "COMMENTS"
                or "KEYWORDS" or "LASTSAVEDBY" or "NUMWORDS" or "NUMCHARS"
                    => $" {effType} ",
                "DATE" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" or "TIME"
                    => string.IsNullOrWhiteSpace(effFormat)
                        ? $" {effType} "
                        : $" {effType} \\@ \"{effFormat}\" ",
                "MERGEFIELD" => string.IsNullOrEmpty(effName)
                    ? throw new ArgumentException("MERGEFIELD requires a 'name' / 'fieldName' property.")
                    : $" MERGEFIELD {effName} ",
                "REF" or "PAGEREF" or "NOTEREF" => string.IsNullOrEmpty(effName)
                    ? throw new ArgumentException($"{effType} requires a 'name' property (target bookmark).")
                    : $" {effType} {effName} ",
                _ => $" {effType}{(string.IsNullOrEmpty(effName) ? "" : " " + effName)}{(string.IsNullOrWhiteSpace(effFormat) ? "" : $" \\@ \"{effFormat}\"")} "
            };
            field.InstrCode.Text = newInstr;
            field.InstrCode.Space = SpaceProcessingModeValues.Preserve;
            var beginCharR = field.BeginRun.GetFirstChild<FieldChar>();
            if (beginCharR != null) beginCharR.Dirty = true;
        }

        foreach (var (key, value) in properties)
        {
            // Handled above by the instruction-rewrite block. Mirror the
            // alias chain in `rewriteName` so type-specific aliases
            // (bookmarkName/styleName/propertyName/...) don't fall
            // through and trigger an unsupported-prop warning even
            // though they were consumed.
            if (key.Equals("fieldType", StringComparison.OrdinalIgnoreCase)
                || key.Equals("fieldtype", StringComparison.OrdinalIgnoreCase)
                || key.Equals("type", StringComparison.OrdinalIgnoreCase)
                || key.Equals("name", StringComparison.OrdinalIgnoreCase)
                || key.Equals("fieldName", StringComparison.OrdinalIgnoreCase)
                || key.Equals("fieldname", StringComparison.OrdinalIgnoreCase)
                || key.Equals("bookmarkName", StringComparison.OrdinalIgnoreCase)
                || key.Equals("bookmarkname", StringComparison.OrdinalIgnoreCase)
                || key.Equals("bookmark", StringComparison.OrdinalIgnoreCase)
                || key.Equals("styleName", StringComparison.OrdinalIgnoreCase)
                || key.Equals("stylename", StringComparison.OrdinalIgnoreCase)
                || key.Equals("propertyName", StringComparison.OrdinalIgnoreCase)
                || key.Equals("propertyname", StringComparison.OrdinalIgnoreCase)
                || key.Equals("format", StringComparison.OrdinalIgnoreCase))
                continue;
            switch (key.ToLowerInvariant())
            {
                // CONSISTENCY(canonical-keys): mirror AddField's
                // GetRawFieldInstruction (instr | instruction | code).
                // Round 6 added the alias trio on Add; the Set side must
                // accept the same set or `--prop code=...` becomes silent
                // unsupported here while it succeeds on Add.
                case "instruction" or "instr" or "code":
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

    private List<string> SetTocPath(System.Text.RegularExpressions.Match tocMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var tocIdx = tocMatch.Groups[1].Success ? int.Parse(tocMatch.Groups[1].Value) : 1;
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

        // Update title — replace text on the immediately-preceding TOCHeading
        // paragraph (mirrors AddToc which inserts one before the TOC field).
        // If no TOCHeading paragraph exists yet, insert one.
        if (properties.TryGetValue("title", out var newTitle))
        {
            var prev = tocPara.PreviousSibling();
            Paragraph? titlePara = null;
            if (prev is Paragraph pp
                && string.Equals(pp.ParagraphProperties?.ParagraphStyleId?.Val?.Value,
                    "TOCHeading", StringComparison.Ordinal))
            {
                titlePara = pp;
            }
            if (titlePara != null)
            {
                titlePara.RemoveAllChildren();
                titlePara.AppendChild(new ParagraphProperties(
                    new ParagraphStyleId { Val = "TOCHeading" }));
                titlePara.AppendChild(new Run(new Text(newTitle)
                    { Space = SpaceProcessingModeValues.Preserve }));
            }
            else if (!string.IsNullOrEmpty(newTitle))
            {
                titlePara = new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                    new Run(new Text(newTitle) { Space = SpaceProcessingModeValues.Preserve }));
                tocPara.InsertBeforeSelf(titlePara);
            }
        }

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

    private List<string> SetFootnotePath(System.Text.RegularExpressions.Match fnSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
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

        // Reject text mutation on separator / continuation-separator footnotes.
        // These are structural placeholders (Type=separator/continuationSeparator,
        // Id=-1/0) that Word renders as a horizontal rule rather than authored
        // text — silently mutating their inner Run text used to be reported as
        // success without any visible effect.
        if (properties.ContainsKey("text") && fn.Type?.Value is FootnoteEndnoteValues fnt
            && (fnt == FootnoteEndnoteValues.Separator
                || fnt == FootnoteEndnoteValues.ContinuationSeparator
                || fnt == FootnoteEndnoteValues.ContinuationNotice))
        {
            throw new ArgumentException(
                $"Cannot set text on footnote separator (id={fn.Id?.Value}, type={fn.Type?.InnerText ?? "?"}). " +
                "Separator footnotes are structural; only user footnotes (id>=1) accept text.");
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
        // i18n: route paragraph-level and run-level format keys through the
        // same helpers SetHeaderFooter uses so direction / font.cs / bold.cs
        // / italic.cs / size.cs etc. work on footnote content. Mirrors the
        // R2-4 footer/header fix.
        ApplyFootnoteEndnoteFormatKeys(fn, properties, unsupported);
        _doc.MainDocumentPart?.FootnotesPart?.Footnotes?.Save();
        return unsupported;
    }

    /// <summary>
    /// Apply paragraph-level and run-level format keys to a footnote /
    /// endnote content body. Skips 'text' (handled separately by the
    /// caller) and silently consumes keys that ApplyParagraphLevelProperty
    /// or ApplyRunFormatting accept. Anything left over is reported as
    /// unsupported.
    /// </summary>
    private void ApplyFootnoteEndnoteFormatKeys(
        OpenXmlElement noteBody,
        Dictionary<string, string> properties,
        List<string> unsupported)
    {
        var firstPara = noteBody.Descendants<Paragraph>().FirstOrDefault();
        if (firstPara == null) return;
        var pProps = firstPara.ParagraphProperties ?? firstPara.PrependChild(new ParagraphProperties());
        // Run targets: skip the reference-mark run so cosmetic styling
        // (bold/italic/font/etc.) doesn't accidentally clobber the
        // footnote/endnote ref mark, which Word renders as a superscript
        // marker outside the authored text.
        var contentRuns = noteBody.Descendants<Run>()
            .Where(r => r.GetFirstChild<FootnoteReferenceMark>() == null
                     && r.GetFirstChild<EndnoteReferenceMark>() == null)
            .ToList();
        var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
        foreach (var (key, value) in properties)
        {
            if (key.Equals("text", StringComparison.OrdinalIgnoreCase)) continue;
            if (ApplyParagraphLevelProperty(pProps, key, value)) continue;
            bool runApplied = false;
            foreach (var run in contentRuns)
            {
                if (ApplyRunFormatting(EnsureRunProperties(run), key, value))
                    runApplied = true;
            }
            if (runApplied)
            {
                // Keep paragraph-mark rPr in sync so later runs inherit.
                ApplyRunFormatting(markRPr, key, value);
                continue;
            }
            unsupported.Add(key);
        }
        if (markRPr.ChildElements.Count == 0) markRPr.Remove();
    }

    /// <summary>
    /// Apply paragraph-level and run-level format keys to a comment body.
    /// Skips text/author/initials/date (handled separately by the caller)
    /// and silently consumes keys that ApplyParagraphLevelProperty or
    /// ApplyRunFormatting accept. Anything left over is reported as
    /// unsupported. Mirrors ApplyFootnoteEndnoteFormatKeys.
    /// </summary>
    private void ApplyCommentFormatKeys(
        Comment comment,
        Dictionary<string, string> properties,
        List<string> unsupported)
    {
        var firstPara = comment.Descendants<Paragraph>().FirstOrDefault();
        if (firstPara == null) return;
        var pProps = firstPara.ParagraphProperties ?? firstPara.PrependChild(new ParagraphProperties());
        var contentRuns = comment.Descendants<Run>().ToList();
        var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
        foreach (var (key, value) in properties)
        {
            var lk = key.ToLowerInvariant();
            if (lk == "text" || lk == "author" || lk == "initials" || lk == "date") continue;
            if (ApplyParagraphLevelProperty(pProps, key, value)) continue;
            bool runApplied = false;
            foreach (var run in contentRuns)
            {
                if (ApplyRunFormatting(EnsureRunProperties(run), key, value))
                    runApplied = true;
            }
            if (runApplied)
            {
                ApplyRunFormatting(markRPr, key, value);
                continue;
            }
            unsupported.Add(key);
        }
        if (markRPr.ChildElements.Count == 0) markRPr.Remove();
    }

    private List<string> SetEndnotePath(System.Text.RegularExpressions.Match enSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
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

        if (properties.ContainsKey("text") && en.Type?.Value is FootnoteEndnoteValues ent
            && (ent == FootnoteEndnoteValues.Separator
                || ent == FootnoteEndnoteValues.ContinuationSeparator
                || ent == FootnoteEndnoteValues.ContinuationNotice))
        {
            throw new ArgumentException(
                $"Cannot set text on endnote separator (id={en.Id?.Value}, type={en.Type?.InnerText ?? "?"}). " +
                "Separator endnotes are structural; only user endnotes (id>=1) accept text.");
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
        // i18n: route paragraph-level and run-level format keys through the
        // same helpers as SetFootnotePath. See ApplyFootnoteEndnoteFormatKeys.
        ApplyFootnoteEndnoteFormatKeys(en, properties, unsupported);
        _doc.MainDocumentPart?.EndnotesPart?.Endnotes?.Save();
        return unsupported;
    }

    private List<string> SetSectionPath(System.Text.RegularExpressions.Match secSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
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
        // CONSISTENCY(set-atomicity): mirror SetDocumentProperties in WordHandler.Add.cs
        // — multi-prop set on /section[N] must be all-or-nothing. Snapshot the whole
        // Document tree on entry; any throw inside the loop restores it before re-throw
        // so partial writes are not visible to the next read in the resident process.
        var atomicSnapshot = _doc.MainDocumentPart!.Document!.OuterXml;
        try
        {
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                // bt-4: 'break' is the natural prop users reach for ("section
                // break = new page"). Treat it as an alias for 'type' and
                // accept the common 'newPage' synonym for nextPage.
                // CONSISTENCY(section-type-alias).
                case "type":
                case "break":
                    var st = sectPr.GetFirstChild<SectionType>() ?? sectPr.PrependChild(new SectionType());
                    st.Val = value.ToLowerInvariant() switch
                    {
                        "nextpage" or "next" or "newpage" or "page" => SectionMarkValues.NextPage,
                        "continuous" => SectionMarkValues.Continuous,
                        "evenpage" or "even" => SectionMarkValues.EvenPage,
                        "oddpage" or "odd" => SectionMarkValues.OddPage,
                        _ => throw new ArgumentException($"Invalid section break type: '{value}'. Valid values: nextPage (alias: newPage/page), continuous, evenPage, oddPage.")
                    };
                    break;
                case "pagewidth" or "pageWidth" or "width":
                {
                    var twW = ParseTwips(value);
                    Core.WordPageDefaults.ValidatePageDim(twW, "pageWidth");
                    EnsureSectPrPageSize(sectPr).Width = twW;
                    break;
                }
                case "pageheight" or "pageHeight" or "height":
                {
                    var twH = ParseTwips(value);
                    Core.WordPageDefaults.ValidatePageDim(twH, "pageHeight");
                    EnsureSectPrPageSize(sectPr).Height = twH;
                    break;
                }
                case "orientation":
                {
                    var ps = EnsureSectPrPageSize(sectPr);
                    var orientLower = value.ToLowerInvariant();
                    if (orientLower != "landscape" && orientLower != "portrait")
                        throw new ArgumentException($"Invalid orientation: '{value}'. Valid: portrait, landscape.");
                    var isLandscape = orientLower == "landscape";
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
                case "pagestart" or "pagenumberstart" or "pagenumstart":
                {
                    var lower = value.ToLowerInvariant();
                    if (lower is "none" or "off" or "false" or "auto")
                    {
                        sectPr.RemoveAllChildren<PageNumberType>();
                    }
                    else
                    {
                        var startN = ParseHelpers.SafeParseInt(value, "pageStart");
                        if (startN < 0)
                            throw new ArgumentException("pageStart must be a non-negative integer.");
                        var pgNum = sectPr.GetFirstChild<PageNumberType>();
                        if (pgNum == null)
                        {
                            pgNum = new PageNumberType();
                            InsertSectPrChildInOrder(sectPr, pgNum);
                        }
                        pgNum.Start = startN;
                    }
                    break;
                }
                case "pagenumfmt" or "pagenumberformat" or "pagenumberfmt":
                {
                    var pgNum = sectPr.GetFirstChild<PageNumberType>();
                    if (pgNum == null)
                    {
                        pgNum = new PageNumberType();
                        InsertSectPrChildInOrder(sectPr, pgNum);
                    }
                    pgNum.Format = ParseNumberFormat(value);
                    break;
                }
                case "direction" or "dir" or "bidi":
                {
                    // Section-level RTL: <w:bidi/> in sectPr flips the page
                    // (margin gutter, header/footer anchors, page-number side).
                    // Required for visually-correct Arabic / Hebrew documents
                    // alongside paragraph-level direction.
                    sectPr.RemoveAllChildren<BiDi>();
                    if (ParseDirectionRtl(value)) InsertSectPrChildInOrder(sectPr, new BiDi());
                    break;
                }
                case "titlepage" or "titlepg":
                {
                    if (IsTruthy(value))
                    {
                        if (sectPr.GetFirstChild<TitlePage>() == null)
                            InsertSectPrChildInOrder(sectPr, new TitlePage());
                    }
                    else
                    {
                        sectPr.RemoveAllChildren<TitlePage>();
                    }
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
                            InsertSectPrChildInOrder(sectPr, lnNum);
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
                                _ => throw new ArgumentException(
                                    $"Invalid lineNumbers value: '{value}'. Valid: continuous, restartPage, restartSection, none, or a positive integer.")
                            };
                        }
                    }
                    break;
                }
                default:
                    // Generic dotted "element.attr=value" fallback (pgSz.orient,
                    // pgMar.top, cols.num, …). Same helper as paragraph/run
                    // and /styles paths.
                    if (key.Contains('.')
                        && Core.TypedAttributeFallback.TrySet(sectPr, key, value))
                        break;
                    unsupported.Add(key);
                    break;
            }
        }
        }
        catch
        {
            _doc.MainDocumentPart!.Document = new Document(atomicSnapshot);
            throw;
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    /// <summary>
    /// Set props on a numbering definition.
    /// Path /numbering/abstractNum[@id=N] targets top-level template props
    /// (name, styleLink, numStyleLink, multiLevelType).
    /// Path /numbering/abstractNum[@id=N]/level[L] targets a specific level
    /// (numFmt, lvlText, start, justification, indent, hanging, suff, font,
    ///  size, color, bold, italic).
    /// CONSISTENCY(set-no-create): never auto-creates the abstractNum or
    /// level — Add owns creation. See SetStylePath for the same rule.
    /// </summary>
    private List<string> SetAbstractNumPath(System.Text.RegularExpressions.Match absNumSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var abstractNumId = int.Parse(absNumSetMatch.Groups[1].Value);
        var levelGroup = absNumSetMatch.Groups[2];
        int? targetLevel = levelGroup.Success ? int.Parse(levelGroup.Value) : (int?)null;

        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering
            ?? throw new ArgumentException("No numbering part. Use `add /numbering --type abstractNum` first.");
        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId)
            ?? throw new ArgumentException(
                $"abstractNum with id={abstractNumId} not found. Use `add /numbering --type abstractNum --prop id={abstractNumId}` first.");

        Level? level = null;
        if (targetLevel.HasValue)
        {
            level = abstractNum.Elements<Level>()
                .FirstOrDefault(l => l.LevelIndex?.Value == targetLevel.Value)
                ?? throw new ArgumentException(
                    $"level[{targetLevel}] not found in abstractNum {abstractNumId}.");
        }

        foreach (var (key, value) in properties)
        {
            if (level != null)
            {
                // Level-scope props
                switch (key.ToLowerInvariant())
                {
                    case "format" or "numfmt":
                        var fmtV = ParseNumberFormat(value);
                        var nf = level.GetFirstChild<NumberingFormat>();
                        if (nf == null) level.AppendChild(new NumberingFormat { Val = fmtV });
                        else nf.Val = fmtV;
                        break;
                    case "text" or "lvltext":
                        var lt = level.GetFirstChild<LevelText>();
                        if (lt == null) level.AppendChild(new LevelText { Val = value });
                        else lt.Val = value;
                        break;
                    case "start":
                        var sn = level.GetFirstChild<StartNumberingValue>();
                        if (sn == null) level.AppendChild(new StartNumberingValue { Val = ParseHelpers.SafeParseInt(value, "start") });
                        else sn.Val = ParseHelpers.SafeParseInt(value, "start");
                        break;
                    case "justification" or "jc" or "lvljc":
                        var jcV = value.ToLowerInvariant() switch
                        {
                            "left" or "start" => LevelJustificationValues.Left,
                            "center" => LevelJustificationValues.Center,
                            "right" or "end" => LevelJustificationValues.Right,
                            _ => throw new ArgumentException($"Invalid justification '{value}'. Valid: left, center, right.")
                        };
                        var jc = level.GetFirstChild<LevelJustification>();
                        if (jc == null) level.AppendChild(new LevelJustification { Val = jcV });
                        else jc.Val = jcV;
                        break;
                    case "suff":
                        var sV = value.ToLowerInvariant() switch
                        {
                            "tab" => LevelSuffixValues.Tab,
                            "space" => LevelSuffixValues.Space,
                            "nothing" or "none" => LevelSuffixValues.Nothing,
                            _ => throw new ArgumentException($"Invalid suff '{value}'. Valid: tab, space, nothing.")
                        };
                        var su = level.GetFirstChild<LevelSuffix>();
                        if (su == null) level.AppendChild(new LevelSuffix { Val = sV });
                        else su.Val = sV;
                        break;
                    case "indent":
                        var ppr = level.PreviousParagraphProperties ?? level.AppendChild(new PreviousParagraphProperties());
                        var indL = ppr.Indentation ?? ppr.AppendChild(new Indentation());
                        indL.Left = ParseHelpers.SafeParseInt(value, "indent").ToString();
                        break;
                    case "hanging":
                        var pprH = level.PreviousParagraphProperties ?? level.AppendChild(new PreviousParagraphProperties());
                        var indH = pprH.Indentation ?? pprH.AppendChild(new Indentation());
                        indH.Hanging = ParseHelpers.SafeParseInt(value, "hanging").ToString();
                        break;
                    case "font":
                        var rpFont = level.NumberingSymbolRunProperties ?? level.AppendChild(new NumberingSymbolRunProperties());
                        var rf = rpFont.GetFirstChild<RunFonts>() ?? rpFont.AppendChild(new RunFonts());
                        rf.Ascii = value;
                        rf.HighAnsi = value;
                        rf.EastAsia = value;
                        break;
                    case "size":
                        var rpSize = level.NumberingSymbolRunProperties ?? level.AppendChild(new NumberingSymbolRunProperties());
                        var halfPt = (int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero);
                        var fs = rpSize.GetFirstChild<FontSize>();
                        if (fs == null) rpSize.AppendChild(new FontSize { Val = halfPt.ToString() });
                        else fs.Val = halfPt.ToString();
                        break;
                    case "color":
                        var rpColor = level.NumberingSymbolRunProperties ?? level.AppendChild(new NumberingSymbolRunProperties());
                        var c = rpColor.GetFirstChild<Color>();
                        if (c == null) rpColor.AppendChild(new Color { Val = SanitizeHex(value) });
                        else c.Val = SanitizeHex(value);
                        break;
                    case "bold":
                        var rpBold = level.NumberingSymbolRunProperties ?? level.AppendChild(new NumberingSymbolRunProperties());
                        if (IsTruthy(value))
                        {
                            if (rpBold.GetFirstChild<Bold>() == null) rpBold.AppendChild(new Bold());
                        }
                        else rpBold.GetFirstChild<Bold>()?.Remove();
                        break;
                    case "italic":
                        var rpItal = level.NumberingSymbolRunProperties ?? level.AppendChild(new NumberingSymbolRunProperties());
                        if (IsTruthy(value))
                        {
                            if (rpItal.GetFirstChild<Italic>() == null) rpItal.AppendChild(new Italic());
                        }
                        else rpItal.GetFirstChild<Italic>()?.Remove();
                        break;
                    case "lvlrestart":
                        // CONSISTENCY(schema-order): CT_Lvl sequence is
                        // start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText,
                        // lvlPicBulletId, legacy, lvlJc, pPr, rPr. Insert before
                        // the first existing sibling that comes later, otherwise
                        // Word silently drops out-of-order children.
                        var lrV = ParseHelpers.SafeParseInt(value, "lvlRestart");
                        var lr = level.GetFirstChild<LevelRestart>();
                        if (lr == null) InsertLevelChildInOrder(level, new LevelRestart { Val = lrV });
                        else lr.Val = lrV;
                        break;
                    case "islgl":
                        var lgl = level.GetFirstChild<IsLegalNumberingStyle>();
                        if (IsTruthy(value))
                        {
                            if (lgl == null) InsertLevelChildInOrder(level, new IsLegalNumberingStyle());
                        }
                        else lgl?.Remove();
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            else
            {
                // abstractNum-scope props (top level)
                // CONSISTENCY(schema-order): CT_AbstractNum sequence is
                // nsid? multiLevelType? tmpl? name? styleLink? numStyleLink? lvl[0..8].
                // When inserting a header element that was absent at Add time, use
                // InsertBefore(firstLevel) rather than AppendChild so the element
                // lands before the level children instead of after them.
                // CONSISTENCY(set-no-create): these only insert; Set never creates levels.
                var firstLvl = abstractNum.GetFirstChild<Level>();
                switch (key.ToLowerInvariant())
                {
                    case "name":
                        var nm = abstractNum.GetFirstChild<AbstractNumDefinitionName>();
                        if (nm == null)
                        {
                            var newNm = new AbstractNumDefinitionName { Val = value };
                            if (firstLvl != null) abstractNum.InsertBefore(newNm, firstLvl);
                            else abstractNum.AppendChild(newNm);
                        }
                        else nm.Val = value;
                        break;
                    case "stylelink":
                        var sl = abstractNum.GetFirstChild<StyleLink>();
                        if (sl == null)
                        {
                            var newSl = new StyleLink { Val = value };
                            if (firstLvl != null) abstractNum.InsertBefore(newSl, firstLvl);
                            else abstractNum.AppendChild(newSl);
                        }
                        else sl.Val = value;
                        break;
                    case "numstylelink":
                        var nsl = abstractNum.GetFirstChild<NumberingStyleLink>();
                        if (nsl == null)
                        {
                            var newNsl = new NumberingStyleLink { Val = value };
                            if (firstLvl != null) abstractNum.InsertBefore(newNsl, firstLvl);
                            else abstractNum.AppendChild(newNsl);
                        }
                        else nsl.Val = value;
                        break;
                    case "type" or "multileveltype":
                        var mltV = value.ToLowerInvariant() switch
                        {
                            "hybridmultilevel" or "hybrid" => MultiLevelValues.HybridMultilevel,
                            "multilevel" or "multi" => MultiLevelValues.Multilevel,
                            "singlelevel" or "single" => MultiLevelValues.SingleLevel,
                            _ => throw new ArgumentException($"Unknown multiLevelType '{value}'. Valid: hybridMultilevel, multilevel, singleLevel.")
                        };
                        var mlt = abstractNum.GetFirstChild<MultiLevelType>();
                        if (mlt == null)
                        {
                            var newMlt = new MultiLevelType { Val = mltV };
                            if (firstLvl != null) abstractNum.InsertBefore(newMlt, firstLvl);
                            else abstractNum.AppendChild(newMlt);
                        }
                        else mlt.Val = mltV;
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
        }

        numbering.Save();
        return unsupported;
    }

    /// <summary>
    /// Insert a new child into a &lt;w:lvl&gt; honoring the CT_Lvl schema order:
    /// start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText, lvlPicBulletId,
    /// legacy, lvlJc, pPr, rPr. Word silently drops out-of-order children, so
    /// AppendChild is only safe when nothing later in the sequence is present.
    /// CONSISTENCY(schema-order): mirrors AbstractNum InsertBefore-firstLevel pattern.
    /// </summary>
    private static int LvlChildOrder(OpenXmlElement e) => e switch
    {
        StartNumberingValue => 0,
        NumberingFormat => 1,
        LevelRestart => 2,
        ParagraphStyleIdInLevel => 3,
        IsLegalNumberingStyle => 4,
        LevelSuffix => 5,
        LevelText => 6,
        LevelPictureBulletId => 7,
        LegacyNumbering => 8,
        LevelJustification => 9,
        PreviousParagraphProperties => 10,
        NumberingSymbolRunProperties => 11,
        _ => int.MaxValue,
    };

    private static void InsertLevelChildInOrder(Level level, OpenXmlElement child)
    {
        var newOrd = LvlChildOrder(child);
        OpenXmlElement? anchor = null;
        foreach (var c in level.ChildElements)
        {
            if (LvlChildOrder(c) > newOrd) { anchor = c; break; }
        }
        if (anchor != null) level.InsertBefore(child, anchor);
        else level.AppendChild(child);
    }

    /// <summary>
    /// Set props on a NumberingInstance (&lt;w:num&gt;).
    /// Path /numbering/num[@id=N] currently supports updating abstractNumId.
    /// CONSISTENCY(set-no-create): never auto-creates the num — Add owns creation.
    /// </summary>
    private List<string> SetNumPath(System.Text.RegularExpressions.Match numSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var numId = int.Parse(numSetMatch.Groups[1].Value);

        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering
            ?? throw new ArgumentException("No numbering part. Use `add /numbering --type num` first.");
        var inst = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId)
            ?? throw new ArgumentException(
                $"num with id={numId} not found. Use `add /numbering --type num --prop abstractNumId=N` first.");

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "abstractnumid":
                    var aidVal = ParseHelpers.SafeParseInt(value, "abstractNumId");
                    var absExists = numbering.Elements<AbstractNum>()
                        .Any(a => a.AbstractNumberId?.Value == aidVal);
                    if (!absExists)
                        throw new ArgumentException(
                            $"abstractNumId={aidVal} not found in /numbering. " +
                            $"Create the abstractNum first, or pick an existing one via 'officecli query <file> abstractNum'.");
                    var aid = inst.AbstractNumId ?? (inst.AbstractNumId = new AbstractNumId());
                    aid.Val = aidVal;
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        numbering.Save();
        return unsupported;
    }

    private List<string> SetStylePath(System.Text.RegularExpressions.Match styleSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var styleId = styleSetMatch.Groups[1].Value;
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        var style = stylesPart?.Styles?.Elements<Style>().FirstOrDefault(s =>
            s.StyleId?.Value == styleId || s.StyleName?.Val?.Value == styleId);
        if (style == null)
        {
            // CONSISTENCY(set-no-create): Set never creates top-level elements,
            // matching every other Set path (/body/p[N], /chart[N], /section[N],
            // /header[N], ...). Auto-creating styles forced an arbitrary
            // type=paragraph default and made `--prop type=` ambiguous (Add
            // owns type; Set has no business inferring it). Force users
            // through Add, where type is an explicit, validated parameter.
            throw new ArgumentException(
                $"Style '{styleId}' not found. Use `add /styles --type style --prop id={styleId} --prop name=... --prop type=paragraph|character` first.");
        }
        var styles = stylesPart!.Styles!;

        foreach (var (key, value) in properties)
        {
            // CONSISTENCY(run-prop-helper): rPr-style props (font/size/bold/
            // italic/color/highlight/underline/strike/caps/smallcaps/...)
            // delegate to ApplyRunFormatting which works on
            // StyleRunProperties via its OpenXmlCompositeElement base. This
            // also extends Style's previously narrow rPr surface (was 7
            // props) to cover the full ~23-prop ApplyRunFormatting set,
            // matching what Word actually accepts in style/rPr.
            // CONSISTENCY(no-empty-container): probe ApplyRunFormatting on a
            // detached rPr first; only attach a real StyleRunProperties to
            // the style if the probe accepts the key. Pre-creating rPr
            // unconditionally pollutes pure-pPr styles with a stray <w:rPr/>.
            // direction lives on style pPr (<w:bidi/>) — must be routed there
            // BEFORE the rPr probe, because ApplyRunFormatting also accepts
            // direction (writes <w:rtl/> on rPr) and would steal the key.
            // Mirror SetParagraphProperties' direction handler (Set.cs).
            var keyLower = key.ToLowerInvariant();
            if (keyLower is "direction" or "dir" or "bidi")
            {
                var dPPr = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                if (ParseDirectionRtl(value)) dPPr.BiDi = new BiDi();
                else dPPr.RemoveAllChildren<BiDi>();
                continue;
            }

            var rPrProbeFmt = new StyleRunProperties();
            if (ApplyRunFormatting(rPrProbeFmt, key, value))
            {
                ApplyRunFormatting(
                    style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties()),
                    key, value);
                continue;
            }

            switch (key.ToLowerInvariant())
            {
                // CONSISTENCY(style-dual-key): mirror AddStyle's alias chain
                // (id/styleId/styleid for the immutable styleId; name /
                // styleName / stylename for the display name). Round 7
                // wired the aliases on Add; Set was the missing half —
                // `set /styles/X --prop styleName=...` was rejected even
                // though Get exposes `styleName` as a canonical readback
                // key. Same alias-trap pattern policy 19b3dd5b banned.
                case "name" or "stylename":
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
                    // Replace, don't ??= — see BUG-LT3 in WordHandler.Set.cs.
                    if (IsTruthy(value))
                        pPrCs.ContextualSpacing = new ContextualSpacing();
                    else
                        pPrCs.ContextualSpacing = null;
                    break;
                }
                // Mirror paragraph Set's curated toggles (BUG-A2). Without
                // explicit cases here the generic TryCreateTypedChild fallback
                // writes the verbose `<w:keepNext w:val="true"/>` form instead
                // of the bare `<w:keepNext/>`. Functionally equivalent in Word
                // but diverges from paragraph Set, breaking automation that
                // diff-compares the two.
                case "keepnext" or "keepwithnext":
                {
                    var pPrKn = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    if (IsTruthy(value)) pPrKn.KeepNext = new KeepNext();
                    else pPrKn.KeepNext = null;
                    break;
                }
                case "keeplines" or "keeptogether":
                {
                    var pPrKl = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    if (IsTruthy(value)) pPrKl.KeepLines = new KeepLines();
                    else pPrKl.KeepLines = null;
                    break;
                }
                case "pagebreakbefore":
                {
                    var pPrPbb = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    if (IsTruthy(value)) pPrPbb.PageBreakBefore = new PageBreakBefore();
                    else pPrPbb.PageBreakBefore = null;
                    break;
                }
                case "widowcontrol" or "widoworphan":
                {
                    var pPrWc = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    if (IsTruthy(value)) pPrWc.WidowControl = new WidowControl();
                    else pPrWc.WidowControl = new WidowControl { Val = false };
                    break;
                }
                // Numbering linkage on the style itself (numPr inside style/pPr).
                // Mirrors paragraph-level numId/ilvl in WordHandler.Set.cs and
                // AddStyle's numPr support — paragraphs inheriting this style
                // (via pStyle) will pick up numbering through ResolveNumPrFromStyle
                // without needing their own numPr.
                case "numId" or "numid":
                {
                    var pPrN = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    var sNumPr = pPrN.NumberingProperties ?? (pPrN.NumberingProperties = new NumberingProperties());
                    var nid = ParseHelpers.SafeParseInt(value, "numId");
                    if (nid < 0) throw new ArgumentException($"numId must be >= 0 (got {nid}).");
                    // CONSISTENCY(numId-ref-check): mirror Add-side validation
                    // in WordHandler.Add.Structure.cs (commit e85dfd3). Without
                    // this, `set /styles/X --prop numId=99` bypasses the Add
                    // check and leaves the style with a dangling reference,
                    // which the HTML preview then renders as a bullet (R4 bt-4).
                    if (nid > 0)
                    {
                        var numberingPart = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                        var numExists = numberingPart?.Elements<NumberingInstance>()
                            .Any(n => n.NumberID?.Value == nid) ?? false;
                        if (!numExists)
                            throw new ArgumentException(
                                $"numId={nid} not found in /numbering. " +
                                "Create the num first (add /numbering --type num), or use numId=0 to remove numbering.");
                    }
                    sNumPr.NumberingId = new NumberingId { Val = nid };
                    break;
                }
                case "ilvl" or "numLevel" or "numlevel" or "listLevel" or "listlevel":
                {
                    var pPrN2 = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    var sNumPr2 = pPrN2.NumberingProperties ?? (pPrN2.NumberingProperties = new NumberingProperties());
                    var ilvl = ParseHelpers.SafeParseInt(value, "ilvl");
                    if (ilvl < 0 || ilvl > 8)
                        throw new ArgumentException($"ilvl must be in range 0..8 (got {ilvl}).");
                    sNumPr2.NumberingLevelReference = new NumberingLevelReference { Val = ilvl };
                    break;
                }
                case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
                case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right" or "border.between" or "border.bar":
                {
                    var pPrB = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                    ApplyStyleParagraphBorders(pPrB, key, value);
                    break;
                }
                // Per-script font split. Each w:rFonts attr is independent and
                // unset attrs fall back through the style chain / docDefaults,
                // so writing only the requested attr is correct — no need to
                // backfill the others. Merge into any existing w:rFonts so a
                // chain of `set font.eastAsia=…` then `set font.ascii=…`
                // produces a single rFonts element with both attrs.
                case "font.ascii" or "font.hansi" or "font.eastasia" or "font.cs":
                {
                    var rPrFonts = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                    rPrFonts.RunFonts ??= new RunFonts();
                    switch (key.ToLowerInvariant())
                    {
                        case "font.ascii":    rPrFonts.RunFonts.Ascii         = value; break;
                        case "font.hansi":    rPrFonts.RunFonts.HighAnsi      = value; break;
                        case "font.eastasia": rPrFonts.RunFonts.EastAsia      = value; break;
                        case "font.cs":       rPrFonts.RunFonts.ComplexScript = value; break;
                    }
                    break;
                }
                default:
                {
                    // Long-tail OOXML fallback — symmetric with the Get-side
                    // FillUnknownChildProps. Probe pPr first (most paragraph-
                    // level toggles like w:kinsoku, w:snapToGrid, w:wordWrap,
                    // w:autoSpaceDE/DN, w:bidi, w:outlineLvl live there), then
                    // rPr (run-level: w:rtl, w:cs, w:specVanish). Schema-
                    // aware AddChild inside TryCreateTypedChild rejects
                    // mismatched containers, so a wrong probe just returns
                    // false. Use detached probes to avoid creating orphan
                    // empty rPr/pPr on misses.

                    // Dotted "element.attr=value" first, so ind.firstLine /
                    // shd.fill / font.eastAsia / spacing.beforeLines etc.
                    // don't get accidentally coerced into a single-val leaf.
                    if (key.Contains('.'))
                    {
                        var pPrAttrProbe = new StyleParagraphProperties();
                        if (Core.TypedAttributeFallback.TrySet(pPrAttrProbe, key, value))
                        {
                            var pPrReal = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                            Core.TypedAttributeFallback.TrySet(pPrReal, key, value);
                            break;
                        }
                        var rPrAttrProbe = new StyleRunProperties();
                        if (Core.TypedAttributeFallback.TrySet(rPrAttrProbe, key, value))
                        {
                            var rPrReal = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                            Core.TypedAttributeFallback.TrySet(rPrReal, key, value);
                            break;
                        }
                    }

                    var pPrProbe = new StyleParagraphProperties();
                    if (Core.GenericXmlQuery.TryCreateTypedChild(pPrProbe, key, value))
                    {
                        var pPrReal = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
                        Core.GenericXmlQuery.TryCreateTypedChild(pPrReal, key, value);
                        break;
                    }
                    var rPrProbe = new StyleRunProperties();
                    if (Core.GenericXmlQuery.TryCreateTypedChild(rPrProbe, key, value))
                    {
                        var rPrReal = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
                        Core.GenericXmlQuery.TryCreateTypedChild(rPrReal, key, value);
                        break;
                    }
                    unsupported.Add(key);
                    break;
                }
            }
        }
        styles.Save();
        return unsupported;
    }

    private List<string> SetWordOlePath(System.Text.RegularExpressions.Match wordOleSetMatch, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
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

}
