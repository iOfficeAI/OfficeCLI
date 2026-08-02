// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{

    /// <summary>
    /// Build an &lt;a:t&gt; element whose serialized form preserves leading,
    /// trailing or all-whitespace content. Without xml:space="preserve" the
    /// XML serializer treats the text node as collapsible whitespace and
    /// emits the self-closing &lt;a:t/&gt; form — a run that displayed a
    /// single space in the source comes back empty on round-trip. Apply the
    /// attribute only when the content actually needs it, so byte-equal
    /// round-trips for the common case (non-whitespace runs) are unaffected.
    /// </summary>
    internal static Drawing.Text MakePreservingText(string text)
    {
        var t = new Drawing.Text { Text = text ?? string.Empty };
        if (!string.IsNullOrEmpty(text)
            && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
        {
            // <a:t> is OOXML's drawingml run text element; the SDK exposes
            // xml:space only via SetAttribute (no typed property).
            t.SetAttribute(new OpenXmlAttribute(
                "xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve"));
        }
        return t;
    }

    /// <summary>
    /// Pre-process a verbatim &lt;a:txBody&gt;/run-XML string so the SDK parser
    /// keeps whitespace-only and edge-whitespace &lt;a:t&gt; content. When raw
    /// drawingml XML is reparsed via <c>new Drawing.TextBody(xml)</c>, the SDK
    /// reader treats a text node with no <c>xml:space="preserve"</c> as
    /// collapsible whitespace and drops it — a source run that displayed
    /// spaces (visual spacers, indentation) comes back as an empty
    /// self-closing &lt;a:t/&gt;, silently deleting the text. PowerPoint authors
    /// such whitespace runs WITHOUT the attribute, so the captured OuterXml
    /// lacks it; inject it on the parse boundary for every &lt;a:t&gt; whose
    /// content begins or ends with whitespace. Mirrors MakePreservingText's
    /// per-run rule applied across a raw body string. Non-whitespace runs and
    /// &lt;a:t&gt; that already carry xml:space are left untouched.
    /// </summary>
    internal static string PreserveWhitespaceInRawText(string xml)
    {
        if (string.IsNullOrEmpty(xml) || xml.IndexOf("<a:t", StringComparison.Ordinal) < 0)
            return xml;
        // Match an <a:t ...>content</a:t> element (any attribute prefix; the
        // package writer never realiases the `a` prefix on drawingml runs).
        // The (?<!/) before the closing '>' excludes a self-closing empty run
        // <a:t/> / <a:t />: without it, attrs greedily captures the " /" and the
        // '/>' close is misread as an open tag, so .*? swallows the following
        // siblings up to the next </a:t> and xml:space is injected mid-tag,
        // corrupting the body (e.g. "<a:t / xml:space=\"preserve\">…").
        return Regex.Replace(xml,
            @"<a:t(?<attrs>(?:\s[^>]*?)?)(?<!/)>(?<content>.*?)</a:t>",
            m =>
            {
                var attrs = m.Groups["attrs"].Value;
                var content = m.Groups["content"].Value;
                // Already preserved, or no edge whitespace to protect — leave verbatim.
                if (content.Length == 0
                    || attrs.Contains("xml:space", StringComparison.Ordinal)
                    || (!char.IsWhiteSpace(content[0]) && !char.IsWhiteSpace(content[^1])))
                    return m.Value;
                return $"<a:t{attrs} xml:space=\"preserve\">{content}</a:t>";
            },
            RegexOptions.Singleline);
    }

    /// <summary>
    /// Read a table cell's text content, joining multi-paragraph text with "\n".
    /// CONSISTENCY(cell-text-readback): cell.TextBody?.InnerText concatenates
    /// paragraphs without separators, which silently loses line-break structure
    /// on multi-line cells. Get must return the user's input shape verbatim.
    /// </summary>
    internal static string GetCellTextWithParagraphBreaks(Drawing.TableCell cell)
    {
        var tb = cell.TextBody;
        if (tb == null) return "";
        var paragraphs = tb.Elements<Drawing.Paragraph>().ToList();
        if (paragraphs.Count == 0) return tb.InnerText ?? "";
        // NEWLINE-SEMANTICS-V2: render <a:br/> as '\v' inside each
        // paragraph (InnerText drops it silently).
        return string.Join("\n", paragraphs.Select(ParagraphTextWithBreaks));
    }

    private static string ParagraphTextWithBreaks(Drawing.Paragraph p)
    {
        var sb = new StringBuilder();
        foreach (var child in p.ChildElements)
        {
            if (child is Drawing.Run r) sb.Append(r.Text?.Text ?? "");
            else if (child is Drawing.Break) sb.Append('\v');
        }
        return sb.Length == 0 ? (p.InnerText ?? "") : sb.ToString();
    }

    private static string GetShapeText(Shape shape)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return "";

        var sb = new StringBuilder();
        var first = true;
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            if (!first) sb.Append('\n');
            first = false;
            foreach (var child in para.ChildElements)
            {
                if (child is Drawing.Run run)
                    sb.Append(run.Text?.Text ?? "");
                else if (child is Drawing.Break)
                    // NEWLINE-SEMANTICS-V2: <a:br/> reads back as '\v' (soft
                    // line break), matching docx <w:br/> readback; '\n' is
                    // reserved for the paragraph join below.
                    sb.Append('\v');
                else if (child is OpenXmlUnknownElement unk
                         && unk.LocalName == "tab"
                         && unk.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main")
                {
                    // CONSISTENCY(escape-sequences): <a:tab/> is the OOXML wire
                    // form for a literal TAB between text runs (see
                    // AppendLineWithTabs in Add.Text.cs). Round-trip back to '\t'
                    // so Get text mirrors what the user wrote on Add/Set.
                    sb.Append('\t');
                }
                else if (child is Drawing.Field fld)
                {
                    // a:fld renders its cached <a:t> when present; otherwise
                    // PowerPoint hasn't rendered it yet. Cross-handler
                    // `evaluated` protocol — emit #OCLI_NOTEVAL!{type} to
                    // match Word's complex-field sentinel, so agents see the
                    // unrendered field in view text instead of a silent gap.
                    var cached = string.Concat(fld.Elements<Drawing.Text>().Select(t => t.Text));
                    var fldType = fld.Type?.Value ?? "";
                    if (cached.Length > 0) sb.Append(cached);
                    else if (IsDynamicSlideFieldTypeStatic(fldType))
                        sb.Append("#OCLI_NOTEVAL!{").Append(fldType).Append('}');
                }
                else if (HasMathContent(child))
                    sb.Append(FormulaParser.ToReadableText(GetMathElement(child)));
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find all OMML math elements inside a shape's text body.
    /// </summary>
    private static List<OpenXmlElement> FindShapeMathElements(Shape shape)
    {
        var results = new List<OpenXmlElement>();
        var textBody = shape.TextBody;
        if (textBody == null) return results;

        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var child in para.ChildElements)
            {
                if (HasMathContent(child))
                    results.Add(GetMathElement(child));
            }
        }
        return results;
    }

    /// <summary>
    /// Check if an element contains math content (a14:m or mc:AlternateContent with math).
    /// </summary>
    private static bool HasMathContent(OpenXmlElement element)
    {
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
            return true;
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            if (element.Descendants().Any(e => e.LocalName == "oMath" || e.LocalName == "oMathPara"))
                return true;
            return element.InnerXml.Contains("oMath");
        }
        return false;
    }

    /// <summary>
    /// Extract the OMML math element from an a14:m or mc:AlternateContent wrapper.
    /// </summary>
    private static OpenXmlElement GetMathElement(OpenXmlElement element)
    {
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
        {
            var child = element.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (child != null) return child;

            var desc = element.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (desc != null) return desc;

            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;

            return element;
        }
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            var choice = element.ChildElements.FirstOrDefault(e => e is AlternateContentChoice || e.LocalName == "Choice");
            if (choice != null)
            {
                var a14m = choice.ChildElements.FirstOrDefault(e =>
                    e.LocalName == "m" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main");
                if (a14m != null)
                    return GetMathElement(a14m);

                var mathDesc = choice.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                if (mathDesc != null)
                    return mathDesc;
            }

            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;
        }
        return element;
    }

    /// <summary>
    /// Re-parse OMML XML string into an OpenXmlElement with navigable children.
    /// </summary>
    private static OpenXmlElement? ReparseFromXml(string innerXml)
    {
        try
        {
            var xml = innerXml.Trim();
            if (xml.Contains("oMathPara"))
            {
                var startIdx = xml.IndexOf("<m:oMathPara", StringComparison.Ordinal);
                if (startIdx < 0) startIdx = xml.IndexOf("<oMathPara", StringComparison.Ordinal);
                if (startIdx >= 0)
                {
                    var endTag = xml.Contains("</m:oMathPara>") ? "</m:oMathPara>" : "</oMathPara>";
                    var endIdx = xml.IndexOf(endTag, StringComparison.Ordinal);
                    if (endIdx >= 0)
                    {
                        var oMathParaXml = xml[startIdx..(endIdx + endTag.Length)];
                        if (!oMathParaXml.Contains("xmlns:m="))
                            oMathParaXml = oMathParaXml.Replace("<m:oMathPara", "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"");
                        var wrapper = new OpenXmlUnknownElement("m", "oMathPara", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                        var innerStart = oMathParaXml.IndexOf('>') + 1;
                        var innerEnd = oMathParaXml.LastIndexOf('<');
                        if (innerStart > 0 && innerEnd > innerStart)
                            wrapper.InnerXml = oMathParaXml[innerStart..innerEnd];
                        return wrapper;
                    }
                }
            }
            // Inline math without an oMathPara wrapper — author tools emit just
            // <m:oMath>...</m:oMath> directly inside the a14:m container, and
            // the View / equation readback path used to silently drop those.
            // Mirror the oMathPara branch above so bare oMath round-trips.
            if (xml.Contains("oMath"))
            {
                var bareStart = xml.IndexOf("<m:oMath", StringComparison.Ordinal);
                if (bareStart < 0) bareStart = xml.IndexOf("<oMath", StringComparison.Ordinal);
                if (bareStart >= 0)
                {
                    var bareEndTag = xml.Contains("</m:oMath>") ? "</m:oMath>" : "</oMath>";
                    var bareEnd = xml.IndexOf(bareEndTag, StringComparison.Ordinal);
                    if (bareEnd >= 0)
                    {
                        var oMathXml = xml[bareStart..(bareEnd + bareEndTag.Length)];
                        if (!oMathXml.Contains("xmlns:m="))
                            oMathXml = oMathXml.Replace("<m:oMath", "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"");
                        var bareWrapper = new OpenXmlUnknownElement("m", "oMath", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                        var bareInnerStart = oMathXml.IndexOf('>') + 1;
                        var bareInnerEnd = oMathXml.LastIndexOf('<');
                        if (bareInnerStart > 0 && bareInnerEnd > bareInnerStart)
                            bareWrapper.InnerXml = oMathXml[bareInnerStart..bareInnerEnd];
                        return bareWrapper;
                    }
                }
            }
        }
        catch { }
        return null;
    }

    private static bool IsTitle(Shape shape)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return false;
        var type = ph.Type?.Value;
        return type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle;
    }

    private static string GetShapeName(Shape shape) =>
        shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";
}
