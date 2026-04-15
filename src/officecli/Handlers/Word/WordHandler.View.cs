// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== View Helpers ====================

    /// <summary>
    /// CONSISTENCY(ole-stats): OLE objects can live in the body, headers,
    /// or footers. Stats counters previously only walked the body and
    /// undercounted documents that embed OLEs in header/footer regions.
    /// Centralize the cross-part walk so all stats counters stay aligned.
    /// </summary>
    private int CountAllOleObjects()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return 0;
        int total = mainPart.Document?.Body?.Descendants<EmbeddedObject>().Count() ?? 0;
        total += mainPart.HeaderParts.Sum(h => h.Header?.Descendants<EmbeddedObject>().Count() ?? 0);
        total += mainPart.FooterParts.Sum(f => f.Footer?.Descendants<EmbeddedObject>().Count() ?? 0);
        return total;
    }

    /// <summary>
    /// Represents a body element with optional SDT context.
    /// When a paragraph/table is inside an SdtBlock, SdtBlock is set.
    /// </summary>
    private record BodyElementInfo(OpenXmlElement Element, SdtBlock? SdtBlock = null);

    /// <summary>
    /// Enumerate body elements, preserving SDT context.
    /// Elements inside SdtBlock are yielded with a reference to their parent SdtBlock.
    /// </summary>
    private static IEnumerable<BodyElementInfo> GetBodyElementsWithSdtContext(Body body)
    {
        foreach (var element in body.ChildElements)
        {
            if (element is SdtBlock sdt)
            {
                var content = sdt.SdtContentBlock;
                if (content != null)
                {
                    foreach (var child in content.ChildElements)
                        yield return new BodyElementInfo(child, sdt);
                }
            }
            else
            {
                yield return new BodyElementInfo(element);
            }
        }
    }

    /// <summary>
    /// Get SDT label string from an SdtBlock: [sdt:alias] or [sdt:tag] or [sdt].
    /// </summary>
    private static string GetSdtLabel(SdtBlock sdt)
    {
        var props = sdt.SdtProperties;
        var alias = props?.GetFirstChild<SdtAlias>()?.Val?.Value;
        if (!string.IsNullOrEmpty(alias))
            return $"[sdt:{alias}] ";
        var tag = props?.GetFirstChild<Tag>()?.Val?.Value;
        if (!string.IsNullOrEmpty(tag))
            return $"[sdt:{tag}] ";
        return "[sdt] ";
    }

    /// <summary>
    /// Find formfield runs in a paragraph and return (name, type, value) for each.
    /// </summary>
    private static List<(string Name, string FieldType, string Value)> FindFormFieldsInParagraph(Paragraph para)
    {
        var result = new List<(string, string, string)>();
        Run? beginRun = null;
        FormFieldData? ffData = null;
        var resultRuns = new List<Run>();
        bool inResult = false;

        foreach (var run in para.Descendants<Run>())
        {
            var fldChar = run.GetFirstChild<FieldChar>();
            if (fldChar != null)
            {
                var charType = fldChar.FieldCharType?.Value;
                if (charType == FieldCharValues.Begin)
                {
                    beginRun = run;
                    ffData = fldChar.FormFieldData;
                    resultRuns.Clear();
                    inResult = false;
                }
                else if (charType == FieldCharValues.Separate)
                {
                    inResult = true;
                }
                else if (charType == FieldCharValues.End)
                {
                    if (beginRun != null && ffData != null)
                    {
                        var name = ffData.GetFirstChild<FormFieldName>()?.Val?.Value ?? "";
                        var fieldType = "text";
                        if (ffData.GetFirstChild<CheckBox>() != null) fieldType = "checkbox";
                        else if (ffData.GetFirstChild<DropDownListFormField>() != null) fieldType = "dropdown";

                        var value = string.Join("", resultRuns.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
                        if (fieldType == "checkbox")
                        {
                            var cb = ffData.GetFirstChild<CheckBox>();
                            var isChecked = cb?.GetFirstChild<Checked>()?.Val?.Value
                                ?? cb?.GetFirstChild<DefaultCheckBoxFormFieldState>()?.Val?.Value
                                ?? false;
                            value = isChecked ? "true" : "false";
                        }
                        result.Add((name, fieldType, value));
                    }
                    beginRun = null;
                    ffData = null;
                    resultRuns.Clear();
                    inResult = false;
                }
            }
            else if (inResult)
            {
                resultRuns.Add(run);
            }
        }
        return result;
    }

    /// <summary>
    /// Build text line for a paragraph, including formfield markers.
    /// If the paragraph contains formfields, they are annotated as [formfield:name] value.
    /// Otherwise returns null (caller uses normal text extraction).
    /// </summary>
    private static string? GetParagraphTextWithFormFields(Paragraph para)
    {
        var formFields = FindFormFieldsInParagraph(para);
        if (formFields.Count == 0) return null;

        // Build text by walking through paragraph children, replacing field sequences with markers
        var sb = new StringBuilder();
        Run? beginRun = null;
        FormFieldData? currentFfData = null;
        var resultRuns = new List<Run>();
        bool inField = false;
        bool inResult = false;
        int ffIdx = 0;

        foreach (var run in para.Descendants<Run>())
        {
            var fldChar = run.GetFirstChild<FieldChar>();
            if (fldChar != null)
            {
                var charType = fldChar.FieldCharType?.Value;
                if (charType == FieldCharValues.Begin)
                {
                    beginRun = run;
                    currentFfData = fldChar.FormFieldData;
                    resultRuns.Clear();
                    inField = true;
                    inResult = false;
                }
                else if (charType == FieldCharValues.Separate)
                {
                    inResult = true;
                }
                else if (charType == FieldCharValues.End)
                {
                    if (currentFfData != null && ffIdx < formFields.Count)
                    {
                        var ff = formFields[ffIdx++];
                        var label = !string.IsNullOrEmpty(ff.Name) ? $"[formfield:{ff.Name}]" : "[formfield]";
                        sb.Append($"{label} {ff.Value}");
                    }
                    beginRun = null;
                    currentFfData = null;
                    resultRuns.Clear();
                    inField = false;
                    inResult = false;
                }
            }
            else if (inField)
            {
                if (inResult) resultRuns.Add(run);
                // Skip instruction runs
            }
            else
            {
                // Normal run outside any field
                sb.Append(string.Concat(run.Elements<Text>().Select(t => t.Text)));
            }
        }
        return sb.ToString();
    }

    // ==================== Semantic Layer ====================

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        int lineNum = 0;
        int pIdx = 0, tblIdx = 0, eqIdx = 0, sdtIdx = 0;
        int emitted = 0;
        var bodyElements = GetBodyElementsWithSdtContext(body).ToList();
        int totalElements = bodyElements.Count;

        // Track which SdtBlocks we've seen for indexing
        var sdtIndexMap = new Dictionary<SdtBlock, int>();

        foreach (var item in bodyElements)
        {
            var element = item.Element;
            lineNum++;
            string path;
            string sdtLabel = "";

            if (item.SdtBlock != null)
            {
                if (!sdtIndexMap.ContainsKey(item.SdtBlock))
                    sdtIndexMap[item.SdtBlock] = ++sdtIdx;
                sdtLabel = GetSdtLabel(item.SdtBlock);
            }

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                eqIdx++;
                path = $"/body/oMathPara[{eqIdx}]";
            }
            else if (element is Paragraph para1)
            {
                pIdx++;
                var pSeg = BuildParaPathSegment(para1, pIdx);
                path = item.SdtBlock != null ? $"/body/sdt[{sdtIndexMap[item.SdtBlock]}]/{pSeg}" : $"/body/{pSeg}";
            }
            else if (element is Table)
            {
                tblIdx++;
                path = item.SdtBlock != null ? $"/body/sdt[{sdtIndexMap[item.SdtBlock]}]/tbl[{tblIdx}]" : $"/body/tbl[{tblIdx}]";
            }
            else if (IsStructuralElement(element))
            {
                path = $"/body/{element.LocalName}";
            }
            else
            {
                // Skip non-content elements
                continue;
            }

            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {emitted} rows, {totalElements} total, use --start/--end to view more)");
                break;
            }

            if (element is Paragraph para)
            {
                // Check if paragraph contains display equation (oMathPara)
                var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathParaChild != null)
                {
                    var mathText = FormulaParser.ToReadableText(oMathParaChild);
                    sb.AppendLine($"[{path}] {sdtLabel}[Equation] {mathText}");
                }
                else if (para.Descendants<EmbeddedObject>().Any())
                {
                    // CONSISTENCY(word-text-ole): OLE paragraphs emit a
                    // visible placeholder per OLE object so they are
                    // distinguishable from empty paragraphs. Iterate all
                    // EmbeddedObjects in the paragraph — a single paragraph
                    // may contain more than one OLE run. Mirrors
                    // ViewAsAnnotated's word-annotated-ole handling.
                    var listPrefix = GetListPrefix(para);
                    foreach (var embObj in para.Descendants<EmbeddedObject>())
                    {
                        var oleEl = embObj.Descendants()
                            .FirstOrDefault(e => e.LocalName == "OLEObject");
                        var progId = oleEl?.GetAttributes()
                            .FirstOrDefault(a => a.LocalName == "ProgID").Value;
                        if (string.IsNullOrEmpty(progId)) progId = "Object";
                        sb.AppendLine($"[{path}] {sdtLabel}{listPrefix}[OLE: {progId}]");
                    }
                }
                else
                {
                    // Check for formfields first
                    var ffText = GetParagraphTextWithFormFields(para);

                    // Check for inline math
                    var mathElements = FindMathElements(para);
                    if (mathElements.Count > 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
                    {
                        var mathText = string.Concat(mathElements.Select(FormulaParser.ToReadableText));
                        sb.AppendLine($"[{path}] {sdtLabel}[Equation] {mathText}");
                    }
                    else if (ffText != null)
                    {
                        var listPrefix = GetListPrefix(para);
                        sb.AppendLine($"[{path}] {sdtLabel}{listPrefix}{ffText}");
                    }
                    else if (mathElements.Count > 0)
                    {
                        var text = GetParagraphTextWithMath(para);
                        var listPrefix = GetListPrefix(para);
                        sb.AppendLine($"[{path}] {sdtLabel}{listPrefix}{text}");
                    }
                    else
                    {
                        var text = GetParagraphText(para);
                        var listPrefix = GetListPrefix(para);
                        sb.AppendLine($"[{path}] {sdtLabel}{listPrefix}{text}");
                    }
                }
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                var mathText = FormulaParser.ToReadableText(element);
                sb.AppendLine($"[{path}] {sdtLabel}[Equation] {mathText}");
            }
            else if (element is Table table)
            {
                sb.AppendLine($"[{path}] {sdtLabel}[Table: {table.Elements<TableRow>().Count()} rows]");
            }
            else if (IsStructuralElement(element))
            {
                sb.AppendLine($"[{path}] [{element.LocalName}]");
            }
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        int lineNum = 0;
        int pIdx = 0, tblIdx = 0, eqIdx = 0, sdtIdx = 0;
        int emitted = 0;
        var bodyElements = GetBodyElementsWithSdtContext(body).ToList();
        int totalElements = bodyElements.Count;

        // Track which SdtBlocks we've seen for indexing
        var sdtIndexMap = new Dictionary<SdtBlock, int>();

        foreach (var item in bodyElements)
        {
            var element = item.Element;
            lineNum++;
            string path;
            string sdtAnnotation = "";

            if (item.SdtBlock != null)
            {
                if (!sdtIndexMap.ContainsKey(item.SdtBlock))
                    sdtIndexMap[item.SdtBlock] = ++sdtIdx;
                sdtAnnotation = GetSdtLabel(item.SdtBlock).TrimEnd();
            }

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                eqIdx++;
                path = $"/body/oMathPara[{eqIdx}]";
            }
            else if (element is Paragraph para2)
            {
                pIdx++;
                var pSeg = BuildParaPathSegment(para2, pIdx);
                path = item.SdtBlock != null ? $"/body/sdt[{sdtIndexMap[item.SdtBlock]}]/{pSeg}" : $"/body/{pSeg}";
            }
            else if (element is Table)
            {
                tblIdx++;
                path = item.SdtBlock != null ? $"/body/sdt[{sdtIndexMap[item.SdtBlock]}]/tbl[{tblIdx}]" : $"/body/tbl[{tblIdx}]";
            }
            else
            {
                path = $"/body/?[{lineNum}]";
            }

            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {emitted} rows, {totalElements} total, use --start/--end to view more)");
                break;
            }

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                var latex = FormulaParser.ToLatex(element);
                sb.AppendLine($"[{path}] [Equation: \"{latex}\"] ← display");
            }
            else if (element is Paragraph para)
            {
                // Check if paragraph contains display equation (oMathPara)
                var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathParaChild != null)
                {
                    var latex = FormulaParser.ToLatex(oMathParaChild);
                    sb.AppendLine($"[{path}] [Equation: \"{latex}\"] ← display");
                    emitted++;
                    continue;
                }

                var styleName = GetStyleName(para);
                var runs = GetAllRuns(para);

                // Check for inline math
                var inlineMath = FindMathElements(para);
                if (inlineMath.Count > 0 && runs.Count == 0)
                {
                    var latex = string.Concat(inlineMath.Select(FormulaParser.ToLatex));
                    sb.AppendLine($"[{path}] [Equation: \"{latex}\"] ← {styleName} | inline");
                    emitted++;
                    continue;
                }

                if (runs.Count == 0 && inlineMath.Count == 0)
                {
                    var sdtSuffix = !string.IsNullOrEmpty(sdtAnnotation) ? $" | {sdtAnnotation}" : "";
                    sb.AppendLine($"[{path}] [] <- {styleName} | empty paragraph{sdtSuffix}");
                    emitted++;
                    continue;
                }

                var listPrefix = GetListPrefix(para);

                // Build a set of runs that are part of formfield sequences for annotation
                var formFieldRunMap = BuildFormFieldRunMap(para);

                // OLE paragraphs: emit one annotated line per OLE object in the
                // paragraph. A single paragraph may contain multiple OLE runs —
                // iterating all EmbeddedObject descendants ensures none are
                // silently dropped. CONSISTENCY(word-annotated-ole): mirrors
                // the paragraph-level emission fix in ViewAsText above.
                var oleRuns = runs.Where(r => r.GetFirstChild<EmbeddedObject>() != null).ToList();
                if (oleRuns.Count > 0)
                {
                    foreach (var oleRun in oleRuns)
                    {
                        var oleEl = oleRun.GetFirstChild<EmbeddedObject>()!
                            .Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
                        var progId = oleEl?.GetAttributes()
                            .FirstOrDefault(a => a.LocalName == "ProgID").Value ?? "";
                        sb.AppendLine($"[{path}] {listPrefix}[OLE: {progId}] ← {styleName}");
                        emitted++;
                    }
                    continue;
                }

                foreach (var run in runs)
                {
                    // Check if run contains an image
                    var drawing = run.GetFirstChild<Drawing>();
                    if (drawing != null)
                    {
                        var imgInfo = GetDrawingInfo(drawing);
                        sb.AppendLine($"[{path}] {listPrefix}[Image: {imgInfo}] ← {styleName}");
                        continue;
                    }

                    var text = GetRunText(run);
                    var fmt = GetRunFormatDescription(run, para);
                    var extraAnnotations = new List<string>();

                    // Add SDT annotation
                    if (!string.IsNullOrEmpty(sdtAnnotation))
                        extraAnnotations.Add(sdtAnnotation);

                    // Add formfield annotation if this run is part of a formfield
                    if (formFieldRunMap.TryGetValue(run, out var ffInfo))
                        extraAnnotations.Add(ffInfo);

                    var suffix = extraAnnotations.Count > 0 ? " | " + string.Join(" | ", extraAnnotations) : "";

                    sb.AppendLine($"[{path}] {listPrefix}「{text}」 ← {styleName} | {fmt}{suffix}");
                }

                // Show inline math elements
                foreach (var math in inlineMath)
                {
                    var latex = FormulaParser.ToLatex(math);
                    sb.AppendLine($"[{path}] {listPrefix}[Equation: \"{latex}\"] ← {styleName} | inline");
                }
            }
            else if (element is Table table)
            {
                var rows = table.Elements<TableRow>().Count();
                var colCount = table.Elements<TableRow>().FirstOrDefault()
                    ?.Elements<TableCell>().Count() ?? 0;
                sb.AppendLine($"[{path}] [Table: {rows}×{colCount}]");
            }
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    /// <summary>
    /// Build a map from Run to formfield annotation string for runs that are part of formfield sequences.
    /// </summary>
    private static Dictionary<Run, string> BuildFormFieldRunMap(Paragraph para)
    {
        var map = new Dictionary<Run, string>();
        Run? beginRun = null;
        FormFieldData? currentFfData = null;
        var fieldRuns = new List<Run>();
        bool inField = false;

        foreach (var run in para.Descendants<Run>())
        {
            var fldChar = run.GetFirstChild<FieldChar>();
            if (fldChar != null)
            {
                var charType = fldChar.FieldCharType?.Value;
                if (charType == FieldCharValues.Begin)
                {
                    beginRun = run;
                    currentFfData = fldChar.FormFieldData;
                    fieldRuns.Clear();
                    fieldRuns.Add(run);
                    inField = true;
                }
                else if (charType == FieldCharValues.Separate)
                {
                    fieldRuns.Add(run);
                }
                else if (charType == FieldCharValues.End)
                {
                    fieldRuns.Add(run);
                    if (currentFfData != null)
                    {
                        var name = currentFfData.GetFirstChild<FormFieldName>()?.Val?.Value ?? "";
                        var fieldType = "text";
                        if (currentFfData.GetFirstChild<CheckBox>() != null) fieldType = "checkbox";
                        else if (currentFfData.GetFirstChild<DropDownListFormField>() != null) fieldType = "dropdown";

                        var label = !string.IsNullOrEmpty(name) ? $"[formfield:{name} ({fieldType})]" : $"[formfield ({fieldType})]";
                        foreach (var fr in fieldRuns)
                            map[fr] = label;
                    }
                    beginRun = null;
                    currentFfData = null;
                    fieldRuns.Clear();
                    inField = false;
                }
            }
            else if (inField)
            {
                fieldRuns.Add(run);
            }
        }
        return map;
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        // Document info
        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();
        var tables = GetBodyElements(body).OfType<Table>().ToList();
        var imageCount = body.Descendants<Drawing>().Count();
        var oleCount = CountAllOleObjects();
        var equationCount = body.Descendants().Count(e => e.LocalName == "oMathPara" || e is M.Paragraph);
        var formFieldCount = FindFormFields().Count;
        var contentControlCount = body.Descendants<SdtBlock>().Count() + body.Descendants<SdtRun>().Count();
        var statsLine = $"File: {Path.GetFileName(_filePath)} | {paragraphs.Count} paragraphs | {tables.Count} tables | {imageCount} images";
        if (oleCount > 0) statsLine += $" | {oleCount} OLE object{(oleCount == 1 ? "" : "s")}";
        if (equationCount > 0) statsLine += $" | {equationCount} equations";
        if (formFieldCount > 0) statsLine += $" | {formFieldCount} formfields";
        if (contentControlCount > 0) statsLine += $" | {contentControlCount} content controls";
        sb.AppendLine(statsLine);

        // Watermark
        var watermark = FindWatermark();
        if (watermark != null)
            sb.AppendLine($"Watermark: \"{watermark}\"");

        // Headers
        var headers = GetHeaderTexts();
        foreach (var h in headers)
            sb.AppendLine($"Header: \"{h}\"");

        // Footers
        var footers = GetFooterTexts();
        foreach (var f in footers)
            sb.AppendLine($"Footer: \"{f}\"");

        sb.AppendLine();

        // Heading structure
        int lineNum = 0;
        foreach (var para in paragraphs)
        {
            lineNum++;
            var styleName = GetStyleName(para);
            var text = GetParagraphText(para);

            if (styleName.Contains("Heading") || styleName.Contains("标题")
                || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
                || styleName == "Title" || styleName == "Subtitle")
            {
                var level = GetHeadingLevel(styleName);
                var indent = level <= 1 ? "" : new string(' ', (level - 1) * 2);
                var prefix = level == 0 ? "■" : "├──";
                sb.AppendLine($"{indent}{prefix} [{lineNum}] \"{text}\" ({styleName})");
            }
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsStats()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();

        // Style counts
        var styleCounts = new Dictionary<string, int>();
        var fontCounts = new Dictionary<string, int>();
        var sizeCounts = new Dictionary<string, int>();
        int emptyParagraphs = 0;
        int doubleSpaces = 0;
        int totalChars = 0;

        foreach (var para in paragraphs)
        {
            var style = GetStyleName(para);
            styleCounts[style] = styleCounts.GetValueOrDefault(style) + 1;

            var runs = GetAllRuns(para);
            if (runs.Count == 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            {
                emptyParagraphs++;
                continue;
            }

            foreach (var run in runs)
            {
                var text = GetRunText(run);
                totalChars += text.Length;

                if (text.Contains("  "))
                    doubleSpaces++;

                var resolved = ResolveEffectiveRunProperties(run, para);
                var font = GetFontFromProperties(resolved) ?? "(default)";
                fontCounts[font] = fontCounts.GetValueOrDefault(font) + 1;

                var size = GetSizeFromProperties(resolved) ?? "(default)";
                sizeCounts[size] = sizeCounts.GetValueOrDefault(size) + 1;
            }
        }

        int totalWords = 0;
        foreach (var para in paragraphs)
        {
            var paraText = GetParagraphText(para);
            if (!string.IsNullOrWhiteSpace(paraText))
                totalWords += paraText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        }

        sb.AppendLine($"Paragraphs: {paragraphs.Count} | Words: {totalWords} | Total Characters: {totalChars}");
        sb.AppendLine();

        sb.AppendLine("Style Distribution:");
        foreach (var (style, count) in styleCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {style}: {count}");

        sb.AppendLine();
        sb.AppendLine("Font Usage:");
        foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {font}: {count}");

        sb.AppendLine();
        sb.AppendLine("Font Size Usage:");
        foreach (var (size, count) in sizeCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {size}: {count}");

        sb.AppendLine();
        sb.AppendLine($"Empty Paragraphs: {emptyParagraphs}");
        sb.AppendLine($"Consecutive Spaces: {doubleSpaces}");

        // CONSISTENCY(ole-stats): Excel/PPT ViewAsStats report OLE object
        // counts with this exact line format ("OLE Objects: N"). Word must
        // match so users get a uniform cross-handler stats view.
        var oleCount = CountAllOleObjects();
        if (oleCount > 0) sb.AppendLine($"OLE Objects: {oleCount}");

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsStatsJson()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return new JsonObject();

        // CONSISTENCY(ole-stats-json): Excel/PPT ViewAsStatsJson always expose
        // the oleObjects field. Word must too. Count via EmbeddedObject — same
        // source the text-version ViewAsStats() uses.
        var oleObjectsCount = CountAllOleObjects();

        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();
        var styleCounts = new Dictionary<string, int>();
        var fontCounts = new Dictionary<string, int>();
        var sizeCounts = new Dictionary<string, int>();
        int emptyParagraphs = 0, doubleSpaces = 0, totalChars = 0;

        foreach (var para in paragraphs)
        {
            var style = GetStyleName(para);
            styleCounts[style] = styleCounts.GetValueOrDefault(style) + 1;

            var runs = GetAllRuns(para);
            if (runs.Count == 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            {
                emptyParagraphs++;
                continue;
            }

            foreach (var run in runs)
            {
                var text = GetRunText(run);
                totalChars += text.Length;
                if (text.Contains("  ")) doubleSpaces++;

                var resolved = ResolveEffectiveRunProperties(run, para);
                var font = GetFontFromProperties(resolved) ?? "(default)";
                fontCounts[font] = fontCounts.GetValueOrDefault(font) + 1;
                var size = GetSizeFromProperties(resolved) ?? "(default)";
                sizeCounts[size] = sizeCounts.GetValueOrDefault(size) + 1;
            }
        }

        int totalWords = 0;
        foreach (var para in paragraphs)
        {
            var paraText = GetParagraphText(para);
            if (!string.IsNullOrWhiteSpace(paraText))
                totalWords += paraText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        }

        var result = new JsonObject
        {
            ["paragraphs"] = paragraphs.Count,
            ["words"] = totalWords,
            ["totalCharacters"] = totalChars,
            ["emptyParagraphs"] = emptyParagraphs,
            ["consecutiveSpaces"] = doubleSpaces,
            ["oleObjects"] = oleObjectsCount
        };

        var styles = new JsonObject();
        foreach (var (style, count) in styleCounts.OrderByDescending(kv => kv.Value))
            styles[style] = count;
        result["styleDistribution"] = styles;

        var fonts = new JsonObject();
        foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
            fonts[font] = count;
        result["fontUsage"] = fonts;

        var sizes = new JsonObject();
        foreach (var (size, count) in sizeCounts.OrderByDescending(kv => kv.Value))
            sizes[size] = count;
        result["fontSizeUsage"] = sizes;

        return result;
    }

    public JsonNode ViewAsOutlineJson()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return new JsonObject();

        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();
        var tables = GetBodyElements(body).OfType<Table>().ToList();
        var imageCount = body.Descendants<Drawing>().Count();
        var oleCount = CountAllOleObjects();
        var equationCount = body.Descendants().Count(e => e.LocalName == "oMathPara" || e is M.Paragraph);

        var formFieldCount = FindFormFields().Count;
        var contentControlCount = body.Descendants<SdtBlock>().Count() + body.Descendants<SdtRun>().Count();

        var result = new JsonObject
        {
            ["fileName"] = Path.GetFileName(_filePath),
            ["paragraphs"] = paragraphs.Count,
            ["tables"] = tables.Count,
            ["images"] = imageCount,
            ["equations"] = equationCount
        };
        if (oleCount > 0) result["oleObjects"] = oleCount;
        if (formFieldCount > 0) result["formfields"] = formFieldCount;
        if (contentControlCount > 0) result["contentControls"] = contentControlCount;

        var watermark = FindWatermark();
        if (watermark != null) result["watermark"] = watermark;

        var headers = GetHeaderTexts();
        if (headers.Count > 0) result["headers"] = new JsonArray(headers.Select(h => (JsonNode)JsonValue.Create(h)!).ToArray());

        var footers = GetFooterTexts();
        if (footers.Count > 0) result["footers"] = new JsonArray(footers.Select(f => (JsonNode)JsonValue.Create(f)!).ToArray());

        var headingsArray = new JsonArray();
        int lineNum = 0;
        foreach (var para in paragraphs)
        {
            lineNum++;
            var styleName = GetStyleName(para);
            var text = GetParagraphText(para);

            if (styleName.Contains("Heading") || styleName.Contains("标题")
                || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
                || styleName == "Title" || styleName == "Subtitle")
            {
                headingsArray.Add((JsonNode)new JsonObject
                {
                    ["line"] = lineNum,
                    ["text"] = text,
                    ["style"] = styleName,
                    ["level"] = GetHeadingLevel(styleName)
                });
            }
        }
        result["headings"] = headingsArray;

        return result;
    }

    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return new JsonObject { ["elements"] = new JsonArray() };

        var elementsArray = new JsonArray();
        int lineNum = 0;
        int pIdx = 0, tblIdx = 0, eqIdx = 0, sdtIdx = 0;
        int emitted = 0;
        var bodyElements = GetBodyElementsWithSdtContext(body).ToList();
        var sdtIndexMap = new Dictionary<SdtBlock, int>();

        foreach (var item in bodyElements)
        {
            var element = item.Element;
            lineNum++;
            string path;
            string type;
            string? sdtLabel = null;

            if (item.SdtBlock != null)
            {
                if (!sdtIndexMap.ContainsKey(item.SdtBlock))
                    sdtIndexMap[item.SdtBlock] = ++sdtIdx;
                var props = item.SdtBlock.SdtProperties;
                sdtLabel = props?.GetFirstChild<SdtAlias>()?.Val?.Value
                    ?? props?.GetFirstChild<Tag>()?.Val?.Value;
            }

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                eqIdx++;
                path = $"/body/oMathPara[{eqIdx}]";
                type = "equation";
            }
            else if (element is Paragraph para3)
            {
                pIdx++;
                var pSeg = BuildParaPathSegment(para3, pIdx);
                path = item.SdtBlock != null ? $"/body/sdt[{sdtIndexMap[item.SdtBlock]}]/{pSeg}" : $"/body/{pSeg}";
                type = "paragraph";
            }
            else if (element is Table)
            {
                tblIdx++;
                path = item.SdtBlock != null ? $"/body/sdt[{sdtIndexMap[item.SdtBlock]}]/tbl[{tblIdx}]" : $"/body/tbl[{tblIdx}]";
                type = "table";
            }
            else if (IsStructuralElement(element))
            {
                path = $"/body/{element.LocalName}";
                type = element.LocalName;
            }
            else continue;

            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;
            if (maxLines.HasValue && emitted >= maxLines.Value) break;

            string? text = null;
            JsonArray? formFieldsJson = null;
            if (element is Paragraph para)
            {
                var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathParaChild != null)
                {
                    text = FormulaParser.ToReadableText(oMathParaChild);
                    type = "equation";
                }
                else
                {
                    var ffList = FindFormFieldsInParagraph(para);
                    var ffText = ffList.Count > 0 ? GetParagraphTextWithFormFields(para) : null;

                    var mathElements = FindMathElements(para);
                    if (mathElements.Count > 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
                        text = string.Concat(mathElements.Select(FormulaParser.ToReadableText));
                    else if (ffText != null)
                        text = GetListPrefix(para) + ffText;
                    else if (mathElements.Count > 0)
                        text = GetParagraphTextWithMath(para);
                    else
                        text = GetListPrefix(para) + GetParagraphText(para);

                    if (ffList.Count > 0)
                    {
                        formFieldsJson = new JsonArray();
                        foreach (var ff in ffList)
                        {
                            var ffObj = new JsonObject { ["type"] = ff.FieldType, ["value"] = ff.Value };
                            if (!string.IsNullOrEmpty(ff.Name)) ffObj["name"] = ff.Name;
                            formFieldsJson.Add((JsonNode)ffObj);
                        }
                    }
                }
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
                text = FormulaParser.ToReadableText(element);
            else if (element is Table table)
                text = $"[Table: {table.Elements<TableRow>().Count()} rows]";

            var obj = new JsonObject
            {
                ["path"] = path,
                ["type"] = type
            };
            if (text != null) obj["text"] = text;
            if (sdtLabel != null) obj["sdt"] = sdtLabel;
            else if (item.SdtBlock != null) obj["sdt"] = true;
            if (formFieldsJson != null) obj["formFields"] = formFieldsJson;
            elementsArray.Add((JsonNode)obj);
            emitted++;
        }

        return new JsonObject
        {
            ["totalElements"] = bodyElements.Count,
            ["elements"] = elementsArray
        };
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return issues;

        int issueNum = 0;
        int lineNum = -1;

        foreach (var para in GetBodyElements(body).OfType<Paragraph>())
        {
            lineNum++;
            var styleName = GetStyleName(para);
            var runs = GetAllRuns(para);

            // Empty paragraph
            if (runs.Count == 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"S{++issueNum}",
                    Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/body/{BuildParaPathSegment(para, lineNum + 1)}",
                    Message = "Empty paragraph"
                });
            }

            // Paragraph format checks
            var pProps = para.ParagraphProperties;
            if (pProps != null && IsNormalStyle(styleName))
            {
                var indent = pProps.Indentation;
                if (indent?.FirstLine == null || indent.FirstLine.Value == "0")
                {
                    // Skip paragraphs where first-line indent is not expected:
                    // - hanging indent (e.g. bibliography entries)
                    // - centered/right alignment (block-style formatting)
                    // - list items (bullet/numbered)
                    var hasHanging = indent?.Hanging != null && indent.Hanging.Value != "0";
                    var hasHangingChars = indent?.HangingChars != null && indent.HangingChars.Value > 0;
                    var jcVal = pProps.Justification?.Val?.Value;
                    var isCentered = jcVal == JustificationValues.Center || jcVal == JustificationValues.Right
                                  || jcVal == JustificationValues.Distribute;
                    var isList = pProps.NumberingProperties != null;

                    // Only flag if there's actual text and none of the skip conditions apply
                    if (!hasHanging && !hasHangingChars && !isCentered && !isList
                        && runs.Any(r => !string.IsNullOrWhiteSpace(GetRunText(r))))
                    {
                        issues.Add(new DocumentIssue
                        {
                            Id = $"F{++issueNum}",
                            Type = IssueType.Format,
                            Severity = IssueSeverity.Warning,
                            Path = $"/body/{BuildParaPathSegment(para, lineNum + 1)}",
                            Message = "Body paragraph missing first-line indent",
                            Suggestion = "Set first-line indent to 2 characters"
                        });
                    }
                }
            }

            int runIdx = 0;
            foreach (var run in runs)
            {
                var text = GetRunText(run);

                // Double spaces
                if (text.Contains("  "))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Warning,
                        Path = $"/body/{BuildParaPathSegment(para, lineNum + 1)}/r[{runIdx + 1}]",
                        Message = "Consecutive spaces",
                        Context = text,
                        Suggestion = "Merge into a single space"
                    });
                }

                // Duplicate punctuation
                if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[，。！？、；：]{2,}"))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Warning,
                        Path = $"/body/{BuildParaPathSegment(para, lineNum + 1)}/r[{runIdx + 1}]",
                        Message = "Duplicate punctuation",
                        Context = text
                    });
                }

                // Mixed Chinese/English punctuation
                if (HasMixedPunctuation(text))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Info,
                        Path = $"/body/{BuildParaPathSegment(para, lineNum + 1)}/r[{runIdx + 1}]",
                        Message = "Mixed CJK/Latin punctuation",
                        Context = text
                    });
                }

                runIdx++;
            }

            if (limit.HasValue && issues.Count >= limit.Value) break;
        }

        // Filter by type
        if (issueType != null)
        {
            var type = issueType.ToLowerInvariant() switch
            {
                "format" or "f" => IssueType.Format,
                "content" or "c" => IssueType.Content,
                "structure" or "s" => IssueType.Structure,
                _ => (IssueType?)null
            };
            if (type.HasValue)
                issues = issues.Where(i => i.Type == type.Value).ToList();
        }

        return limit.HasValue ? issues.Take(limit.Value).ToList() : issues;
    }

    public string ViewAsForms()
    {
        var sb = new StringBuilder();

        // Document protection
        var (mode, enforced) = GetDocumentProtection();
        var protectionDisplay = mode == "none" || !enforced
            ? "none"
            : $"{mode} (enforced)";
        sb.AppendLine($"Document Protection: {protectionDisplay}");

        // Collect all form fields
        var fields = CollectFormFieldEntries();

        if (fields.Count == 0)
        {
            sb.AppendLine();
            sb.AppendLine("No form fields or content controls found.");
            return sb.ToString().TrimEnd();
        }

        var editable = fields.Where(f => f.Editable).ToList();
        var nonEditable = fields.Where(f => !f.Editable).ToList();

        sb.AppendLine();
        sb.AppendLine($"Editable Fields ({editable.Count}):");
        for (int i = 0; i < editable.Count; i++)
            sb.AppendLine($"  #{i + 1} {FormatFormEntry(editable[i])}");

        if (nonEditable.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"Non-editable Fields ({nonEditable.Count}):");
            for (int i = 0; i < nonEditable.Count; i++)
                sb.AppendLine($"  #{i + 1} {FormatFormEntry(nonEditable[i])}");
        }

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsFormsJson()
    {
        var (mode, enforced) = GetDocumentProtection();
        var fields = CollectFormFieldEntries();

        var result = new JsonObject
        {
            ["protection"] = mode,
            ["protectionEnforced"] = enforced
        };

        var fieldsArray = new JsonArray();
        foreach (var f in fields)
        {
            var obj = new JsonObject
            {
                ["kind"] = f.Kind,
                ["path"] = f.Path,
                ["type"] = f.FieldType,
                ["editable"] = f.Editable
            };
            if (f.Name != null) obj["name"] = f.Name;
            if (f.Alias != null) obj["alias"] = f.Alias;
            if (f.Value != null) obj["value"] = f.Value;
            if (f.Items != null) obj["items"] = f.Items;
            if (f.Lock != null) obj["lock"] = f.Lock;
            if (f.Checked.HasValue) obj["checked"] = f.Checked.Value;
            fieldsArray.Add((JsonNode)obj);
        }
        result["fields"] = fieldsArray;

        return result;
    }

    private record FormFieldEntry(
        string Kind,      // "sdt" or "formfield"
        string Path,
        string FieldType, // "text", "date", "dropdown", "combobox", "checkbox", "richtext"
        bool Editable,
        string? Name = null,
        string? Alias = null,
        string? Value = null,
        string? Items = null,
        string? Lock = null,
        bool? Checked = null);

    private List<FormFieldEntry> CollectFormFieldEntries()
    {
        var entries = new List<FormFieldEntry>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return entries;

        // 1. Collect SDTs
        int blockSdtIdx = 0;
        foreach (var sdt in body.Descendants().Where(e => e is SdtBlock or SdtRun))
        {
            string path;
            SdtProperties? sdtProps;
            string text;

            if (sdt is SdtBlock sdtBlock)
            {
                blockSdtIdx++;
                path = $"/body/sdt[{blockSdtIdx}]";
                sdtProps = sdtBlock.SdtProperties;
                text = string.Concat(sdtBlock.Descendants<Text>().Select(t => t.Text));
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
                    path = $"/body/{BuildParaPathSegment(parentPara, pIdx)}/sdt[{sdtInParaIdx}]";
                }
                else
                {
                    blockSdtIdx++;
                    path = $"/body/sdt[{blockSdtIdx}]";
                }
                sdtProps = sdtRun.SdtProperties;
                text = string.Concat(sdtRun.Descendants<Text>().Select(t => t.Text));
            }
            else continue;

            if (sdtProps == null) continue;

            var alias = sdtProps.GetFirstChild<SdtAlias>()?.Val?.Value;
            var tag = sdtProps.GetFirstChild<Tag>()?.Val?.Value;
            var lockEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            var lockVal = lockEl?.Val?.InnerText;

            // Determine SDT type
            string sdtType;
            if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) sdtType = "dropdown";
            else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) sdtType = "combobox";
            else if (sdtProps.GetFirstChild<SdtContentDate>() != null) sdtType = "date";
            else if (sdtProps.GetFirstChild<SdtContentText>() != null) sdtType = "text";
            else sdtType = "richtext";

            // Items for dropdown/combobox
            string? items = null;
            var ddl = sdtProps.GetFirstChild<SdtContentDropDownList>();
            var combo = sdtProps.GetFirstChild<SdtContentComboBox>();
            var listItems = ddl?.Elements<ListItem>() ?? combo?.Elements<ListItem>();
            if (listItems != null)
            {
                var itemsList = listItems.Select(li => li.DisplayText?.Value ?? li.Value?.Value ?? "").ToList();
                if (itemsList.Count > 0) items = string.Join(",", itemsList);
            }

            var editable = IsSdtEditable(sdtProps);
            var displayValue = string.IsNullOrEmpty(text) ? "(empty)" : text;

            entries.Add(new FormFieldEntry(
                Kind: "sdt",
                Path: path,
                FieldType: sdtType,
                Editable: editable,
                Alias: alias ?? tag,
                Value: displayValue,
                Items: items,
                Lock: lockVal));
        }

        // 2. Collect legacy form fields
        var formFields = FindFormFields();
        for (int i = 0; i < formFields.Count; i++)
        {
            var ff = formFields[i];
            var ffPath = $"/formfield[{i + 1}]";
            var ffNode = FormFieldToNode(ff, ffPath);

            var ffType = ffNode.Format.TryGetValue("formfieldType", out var ftObj) ? ftObj?.ToString() ?? "text" : "text";
            var ffName = ffNode.Format.TryGetValue("name", out var nameObj) ? nameObj?.ToString() : null;
            var ffEditable = ffNode.Format.TryGetValue("editable", out var edObj) && edObj is true;

            string? ffItems = ffNode.Format.TryGetValue("items", out var itemsObj) ? itemsObj?.ToString() : null;
            bool? ffChecked = ffType == "checkbox" && ffNode.Format.TryGetValue("checked", out var chkObj) ? chkObj is true : null;

            var ffValue = ffType == "checkbox" ? null : (string.IsNullOrEmpty(ffNode.Text) ? "(empty)" : ffNode.Text);

            entries.Add(new FormFieldEntry(
                Kind: "formfield",
                Path: ffPath,
                FieldType: ffType,
                Editable: ffEditable,
                Name: ffName,
                Value: ffValue,
                Items: ffItems,
                Checked: ffChecked));
        }

        return entries;
    }

    private static string FormatFormEntry(FormFieldEntry f)
    {
        var sb = new StringBuilder();
        sb.Append($"[{f.Kind}] {f.Path}");

        if (f.Alias != null) sb.Append($"  alias=\"{f.Alias}\"");
        if (f.Name != null) sb.Append($"  name=\"{f.Name}\"");
        sb.Append($"  type={f.FieldType}");

        if (f.Items != null) sb.Append($"  items=\"{f.Items}\"");
        if (f.Checked.HasValue) sb.Append($"  checked={f.Checked.Value.ToString().ToLowerInvariant()}");
        if (f.Value != null) sb.Append($"  value=\"{f.Value}\"");
        if (f.Lock != null) sb.Append($"  lock={f.Lock}");
        sb.Append($"  editable={f.Editable.ToString().ToLowerInvariant()}");

        return sb.ToString();
    }
}
