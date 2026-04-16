// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
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
    // CJK line-break hooks — partial methods are eliminated by the compiler when no implementation exists
    partial void OnHtmlParagraphBegin(Paragraph para);
    partial void OnHtmlParagraphEnd(StringBuilder sb);
    partial void OnHtmlRenderText(StringBuilder sb, string text, RunProperties? rProps, string? runStyle, ref bool handled);

    // ==================== Paragraph Content ====================

    private void RenderParagraphHtml(StringBuilder sb, Paragraph para)
    {
        // Use <div> instead of <p> when paragraph contains block-level elements (text boxes, charts, shapes)
        var tag = HasBlockLevelDrawing(para) ? "div" : "p";
        sb.Append($"<{tag}");
        // Add CSS class for TOC paragraphs (suppress hyperlink styling)
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null && styleId.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
            sb.Append(" class=\"toc\"");
        var pStyle = GetParagraphInlineCss(para);
        if (!string.IsNullOrEmpty(pStyle))
            sb.Append($" style=\"{pStyle}\"");
        sb.Append(">");
        RenderParagraphContentHtml(sb, para);
        sb.AppendLine($"</{tag}>");
    }

    private void RenderParagraphContentHtml(StringBuilder sb, Paragraph para)
    {
        OnHtmlParagraphBegin(para);
        _ctx.CurrentParagraphTabIndex = 0;

        // Render bookmark anchors for internal hyperlink targets
        foreach (var bm in para.Elements<BookmarkStart>())
        {
            var bmName = bm.Name?.Value;
            if (!string.IsNullOrEmpty(bmName) && !bmName.StartsWith("_GoBack"))
                sb.Append($"<a id=\"{HtmlEncodeAttr(bmName)}\"></a>");
        }

        // Collect standalone images that precede a text box group (they overlay the group in Word)
        bool hasTextBoxGroup = HasTextBoxContent(para);
        var preGroupImages = hasTextBoxGroup ? new List<Drawing>() : null;
        bool textBoxSeen = false;

        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
            {
                // Find drawing (direct child or inside mc:AlternateContent Choice)
                // SDK's Descendants<Drawing>() naturally skips mc:Fallback (VML w:pict)
                var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();

                if (drawing != null && HasGroupOrShape(drawing))
                {
                    bool hasTextBox = HasTextBox(drawing);
                    if (hasTextBox && preGroupImages != null)
                    {
                        // Render group with preceding images overlaid into text box
                        RenderDrawingWithOverlaidImages(sb, drawing, preGroupImages);
                        preGroupImages.Clear();
                        textBoxSeen = true;
                    }
                    else
                    {
                        RenderDrawingHtml(sb, drawing);
                    }
                    continue;
                }

                // Collect standalone images before text box group for overlay
                if (hasTextBoxGroup && !textBoxSeen && drawing != null)
                {
                    preGroupImages!.Add(drawing);
                    continue;
                }

                RenderRunHtml(sb, run, para);
            }
            else if (child.LocalName is "ins" or "moveTo")
            {
                // Tracked insertions — underline to match Word's default revision mark style
                var author = child.GetAttributes().FirstOrDefault(a => a.LocalName == "author").Value;
                var authorAttr = string.IsNullOrEmpty(author) ? "" : $" title=\"Inserted by {HtmlEncodeAttr(author)}\"";
                sb.Append($"<span class=\"track-ins\" style=\"text-decoration:underline;color:#2E7D32\"{authorAttr}>");
                foreach (var insRun in child.Elements<Run>())
                    RenderRunHtml(sb, insRun, para);
                sb.Append("</span>");
            }
            else if (child.LocalName is "del" or "moveFrom")
            {
                // Tracked deletions — strikethrough with color, preserving the deleted text
                // The delText inside del runs carries the actual deleted content; we render it so
                // a reader of the preview can see what was removed.
                var author = child.GetAttributes().FirstOrDefault(a => a.LocalName == "author").Value;
                var authorAttr = string.IsNullOrEmpty(author) ? "" : $" title=\"Deleted by {HtmlEncodeAttr(author)}\"";
                var delText = string.Concat(child.Descendants()
                    .Where(e => e.LocalName == "delText" || e.LocalName == "t")
                    .Select(e => e.InnerText));
                if (!string.IsNullOrEmpty(delText))
                    sb.Append($"<span class=\"track-del\" style=\"text-decoration:line-through;color:#C62828\"{authorAttr}>{HtmlEncode(delText)}</span>");
            }
            else if (child is Hyperlink hyperlink)
            {
                var relId = hyperlink.Id?.Value;
                string? url = null;
                if (relId != null)
                {
                    try
                    {
                        url = _doc.MainDocumentPart?.HyperlinkRelationships
                            .FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                    }
                    catch { }
                    if (url == null)
                    {
                        try
                        {
                            url = _doc.MainDocumentPart?.ExternalRelationships
                                .FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                        }
                        catch { }
                    }
                }

                // Also check for internal bookmark links (Anchor property)
                if (url == null && hyperlink.Anchor?.Value != null)
                    url = $"#{hyperlink.Anchor.Value}";

                if (url != null)
                    sb.Append($"<a href=\"{HtmlEncodeAttr(url)}\"{(url.StartsWith("#") ? "" : " target=\"_blank\"")}>");

                foreach (var hRun in hyperlink.Elements<Run>())
                    RenderRunHtml(sb, hRun, para);

                if (url != null)
                    sb.Append("</a>");
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                var latex = FormulaParser.ToLatex(child);
                sb.Append($"<span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(latex)}\"></span>");
            }
            else if (child.LocalName is "sdt" or "smartTag" or "customXml")
            {
                // Content controls, smart tags, custom XML — render their child runs
                foreach (var innerRun in child.Descendants<Run>())
                    RenderRunHtml(sb, innerRun, para);
            }
            else if (child.LocalName == "fldSimple")
            {
                // Simple field codes (page numbers, cross-refs) — render cached display text
                foreach (var fldRun in child.Elements<Run>())
                    RenderRunHtml(sb, fldRun, para);
            }
        }

        OnHtmlParagraphEnd(sb);
    }

    // ==================== Run Rendering ====================

    private void RenderRunHtml(StringBuilder sb, Run run, Paragraph para)
    {
        // Check for drawing (direct or inside mc:AlternateContent)
        var drawing = run.GetFirstChild<Drawing>()
            ?? run.Descendants<Drawing>().FirstOrDefault();
        if (drawing != null)
        {
            RenderDrawingHtml(sb, drawing);
            return;
        }

        // OLE embedded objects (Visio, Excel, etc.) carry a v:imagedata
        // preview image that we can render for a read-only snapshot.
        var oleObject = run.GetFirstChild<EmbeddedObject>();
        if (oleObject != null)
        {
            RenderOlePreviewHtml(sb, oleObject);
            return;
        }

        // Form field checkbox: fldChar begin with ffData/ffCheckBox — emit ☑ / ☐ glyph
        var fldChar = run.GetFirstChild<FieldChar>();
        if (fldChar?.FieldCharType?.Value == FieldCharValues.Begin)
        {
            var ffData = fldChar.GetFirstChild<FormFieldData>();
            var checkBox = ffData?.GetFirstChild<CheckBox>();
            if (checkBox != null)
            {
                var defaultChecked = checkBox.GetFirstChild<DefaultCheckBoxFormFieldState>()?.Val?.Value == true;
                var currentChecked = checkBox.GetFirstChild<Checked>()?.Val?.Value == true;
                var isChecked = currentChecked || defaultChecked;
                sb.Append(isChecked ? "☑" : "☐");
                return;
            }
        }

        // Footnote/endnote reference — render superscript number (don't return, run may also have text)
        var fnRef = run.GetFirstChild<FootnoteReference>();
        if (fnRef?.Id?.HasValue == true && fnRef.Id.Value > 0)
        {
            var fnId = (int)fnRef.Id.Value;
            _ctx.FootnoteRefs.Add(fnId);
            var fnNum = _ctx.FootnoteRefs.Count;
            var fnLabel = FormatNoteNumber(fnNum, GetFootnoteNumFmt());
            sb.Append($"<sup class=\"fn-ref\"><a href=\"#fn{fnId}\" id=\"fnref{fnId}\">{fnLabel}</a></sup>");
        }
        var enRef = run.GetFirstChild<EndnoteReference>();
        if (enRef?.Id?.HasValue == true && enRef.Id.Value > 0)
        {
            var enId = (int)enRef.Id.Value;
            _ctx.EndnoteRefs.Add(enId);
            var enNum = _ctx.EndnoteRefs.Count;
            var enLabel = FormatNoteNumber(enNum, GetEndnoteNumFmt());
            sb.Append($"<sup class=\"en-ref\"><a href=\"#en{enId}\" id=\"enref{enId}\">{enLabel}</a></sup>");
        }
        // FootnoteReferenceMark / EndnoteReferenceMark: don't skip the run, just ignore the mark element
        // (the run may also contain text that should be rendered)

        // Ruby (furigana) annotation — emit <ruby>base<rt>annotation</rt></ruby>
        var ruby = run.ChildElements.FirstOrDefault(c => c.LocalName == "ruby");
        if (ruby != null)
        {
            var rubyBase = ruby.ChildElements.FirstOrDefault(c => c.LocalName == "rubyBase");
            var rt = ruby.ChildElements.FirstOrDefault(c => c.LocalName == "rt");
            var baseText = string.Concat(rubyBase?.Descendants<Text>().Select(t => t.Text) ?? []);
            var rtText = string.Concat(rt?.Descendants<Text>().Select(t => t.Text) ?? []);
            if (!string.IsNullOrEmpty(baseText))
            {
                sb.Append($"<ruby>{HtmlEncode(baseText)}<rt>{HtmlEncode(rtText)}</rt></ruby>");
                return;
            }
        }

        var hasContent = run.ChildElements.Any(c =>
            c is Break || c is TabChar || c is SymbolChar || c is CarriageReturn
            || c.LocalName is "noBreakHyphen" or "softHyphen"
            || (c is Text t && !string.IsNullOrEmpty(t.Text)));

        if (!hasContent) return;

        var rProps = ResolveEffectiveRunProperties(run, para);
        // w:vanish / w:specVanish — hidden text should be omitted from the
        // visual preview, matching native Word's default view behavior.
        if (rProps.Vanish != null && (rProps.Vanish.Val == null || rProps.Vanish.Val.Value))
            return;
        if (rProps.SpecVanish != null && (rProps.SpecVanish.Val == null || rProps.SpecVanish.Val.Value))
            return;
        var style = GetRunInlineCss(rProps);
        var needsSpan = !string.IsNullOrEmpty(style);

        // When line-break tracking is active, text is buffered and flushed later
        // with style spans — skip the outer span to avoid double-wrapping
        if (needsSpan && !_ctx.LineBreakEnabled)
            sb.Append($"<span style=\"{style}\">");

        foreach (var child in run.ChildElements)
        {
            if (child is Break brk)
            {
                if (brk.Type?.Value == BreakValues.Page)
                    sb.Append("<!--PAGE_BREAK-->");
                else if (brk.Type?.Value == BreakValues.Column)
                {
                    // Close current span/paragraph, insert block-level column break, reopen
                    if (needsSpan) sb.Append("</span>");
                    sb.Append("</p><p style=\"break-before:column\">");
                    if (needsSpan) sb.Append($"<span style=\"{style}\">");
                }
                else
                    sb.Append("<br>");
            }
            else if (child is TabChar)
            {
                // Resolve tab stops: direct on paragraph, or via its style
                var tabs = para.ParagraphProperties?.Tabs?.Elements<TabStop>();
                if (tabs == null || !tabs.Any())
                {
                    var tsId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    if (tsId != null) tabs = ResolveTabStopsFromStyle(tsId);
                }
                // TOC-style special case: any right-aligned tab with dot leader
                var rightDotTab = tabs?.FirstOrDefault(t =>
                    t.Val?.InnerText == "right" && t.Leader?.InnerText == "dot");
                if (rightDotTab != null)
                {
                    // Close current span, insert dot leader, then page number follows
                    if (needsSpan) { sb.Append("</span>"); needsSpan = false; }
                    sb.Append("<span class=\"dot-leader\"></span>");
                }
                else
                {
                    // General tab: emit inline-block with width = distance to Nth tab stop
                    // (or default 36pt = 0.5in fallback when no custom stops defined)
                    var orderedStops = tabs?
                        .Where(t => t.Val?.InnerText != "clear" && t.Position?.HasValue == true)
                        .OrderBy(t => t.Position!.Value).ToList();
                    double widthPt;
                    int tabIdx = _ctx.CurrentParagraphTabIndex;
                    if (orderedStops != null && tabIdx < orderedStops.Count)
                    {
                        var curPos = orderedStops[tabIdx].Position!.Value / 20.0; // twips → pt
                        var prevPos = tabIdx > 0 ? orderedStops[tabIdx - 1].Position!.Value / 20.0 : 0;
                        widthPt = curPos - prevPos;
                        // Handle tab leader (dot, middleDot, hyphen, underscore) for positional tabs
                        var leader = orderedStops[tabIdx].Leader?.InnerText;
                        var cssLeader = leader switch
                        {
                            "dot" => "border-bottom:1px dotted #000;",
                            // middleDot is centered dot between stops — best CSS equivalent is a
                            // thicker dotted border with larger spacing; browsers render dotted
                            // borders with square dots which read as middle dots at 2px width.
                            "middleDot" => "border-bottom:2px dotted #555;",
                            "hyphen" => "border-bottom:1px dashed #000;",
                            "underscore" => "border-bottom:1px solid #000;",
                            _ => "",
                        };
                        sb.Append($"<span style=\"display:inline-block;width:{widthPt:0.##}pt;{cssLeader}\"></span>");
                    }
                    else
                    {
                        // No explicit tab stop: use document-level defaultTabStop
                        // from settings.xml (twips → pt); fallback to 36pt (0.5in)
                        // when settings are missing.
                        var dts = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<DefaultTabStop>();
                        double defTabPt = 36.0;
                        if (dts?.Val?.HasValue == true && dts.Val.Value > 0)
                            defTabPt = dts.Val.Value / 20.0;
                        sb.Append($"<span style=\"display:inline-block;width:{defTabPt:0.##}pt\"></span>");
                    }
                    _ctx.CurrentParagraphTabIndex++;
                }
            }
            else if (child is CarriageReturn)
                sb.Append("<br>");
            else if (child.LocalName == "noBreakHyphen")
                sb.Append("\u2011"); // non-breaking hyphen
            else if (child.LocalName == "softHyphen")
                sb.Append("&shy;");
            else if (child is Text t && !string.IsNullOrEmpty(t.Text))
            {
                bool handled = false;
                OnHtmlRenderText(sb, t.Text, rProps, style, ref handled);
                if (!handled)
                    sb.Append(HtmlEncode(t.Text));
            }
            else if (child is SymbolChar sym)
            {
                // w:sym — render with correct font family for symbol fonts
                var charCode = sym.Char?.Value;
                var symFont = sym.Font?.Value;
                if (charCode != null && int.TryParse(charCode, System.Globalization.NumberStyles.HexNumber, null, out var code))
                {
                    if (symFont != null)
                        sb.Append($"<span style=\"font-family:'{CssSanitize(symFont)}'\">&#x{code:X};</span>");
                    else
                        sb.Append($"&#x{code:X};");
                }
                else
                    sb.Append("\u25A1"); // fallback: □
            }
        }

        if (needsSpan && !_ctx.LineBreakEnabled)
            sb.Append("</span>");
    }

    // ==================== OLE Object Preview Rendering ====================

    /// <summary>
    /// Render the VML preview image that accompanies an embedded OLE object
    /// (e.g. a Visio diagram). Web-compatible formats (PNG/JPEG/GIF/SVG/WebP/BMP)
    /// render as a data-URI &lt;img&gt;; browser-unrenderable formats (EMF/WMF/TIFF)
    /// fall back to a sized placeholder &lt;div&gt;. Pure OpenXML — no GDI and no
    /// System.Drawing dependency.
    /// </summary>
    private void RenderOlePreviewHtml(StringBuilder sb, OpenXmlElement oleObj)
    {
        var imageData = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
        if (imageData == null) return;

        // The r:id attribute lives in the relationships namespace.
        string? relId = null;
        foreach (var attr in imageData.GetAttributes())
        {
            if (attr.LocalName == "id" && (attr.NamespaceUri?.Contains("relationships") ?? false))
            {
                relId = attr.Value;
                break;
            }
        }
        if (string.IsNullOrEmpty(relId)) return;

        var dataUri = LoadImageAsDataUri(relId);
        if (dataUri == null) return;

        // Display size comes from the companion v:shape style
        // ("width:Xpt;height:Ypt"), falling back to the w:object
        // dxaOrig/dyaOrig twip attributes if the shape style is missing.
        double widthPt = 0, heightPt = 0;
        var shape = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shape != null)
        {
            var styleAttr = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
            if (!string.IsNullOrEmpty(styleAttr))
            {
                var wMatch = Regex.Match(styleAttr, @"width:([\d.]+)pt");
                var hMatch = Regex.Match(styleAttr, @"height:([\d.]+)pt");
                if (wMatch.Success)
                    double.TryParse(wMatch.Groups[1].Value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out widthPt);
                if (hMatch.Success)
                    double.TryParse(hMatch.Groups[1].Value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out heightPt);
            }
        }
        if (widthPt == 0 || heightPt == 0)
        {
            foreach (var attr in oleObj.GetAttributes())
            {
                if (attr.LocalName == "dxaOrig" && int.TryParse(attr.Value, out var dxa))
                    widthPt = dxa / 20.0;
                if (attr.LocalName == "dyaOrig" && int.TryParse(attr.Value, out var dya))
                    heightPt = dya / 20.0;
            }
        }

        var widthPx = widthPt > 0 ? (long)(widthPt * 96 / 72) : 0;
        var heightPx = heightPt > 0 ? (long)(heightPt * 96 / 72) : 0;

        bool isWebCompatible = dataUri.Contains("image/png")
            || dataUri.Contains("image/jpeg")
            || dataUri.Contains("image/gif")
            || dataUri.Contains("image/svg")
            || dataUri.Contains("image/webp")
            || dataUri.Contains("image/bmp");

        if (isWebCompatible)
        {
            var widthAttr = widthPx > 0 ? $" width=\"{widthPx}\"" : "";
            var heightAttr = heightPx > 0 ? $" height=\"{heightPx}\"" : "";
            var sizeStyle = widthPx > 0
                ? $"max-width:100%;width:{widthPx}px;height:auto"
                : "max-width:100%";
            sb.Append($"<img src=\"{dataUri}\" alt=\"Embedded object\"{widthAttr}{heightAttr} style=\"{sizeStyle}\">");
        }
        else
        {
            // EMF / WMF / TIFF — browsers cannot render these natively.
            // Emit a sized placeholder so the layout keeps its footprint.
            var ph = widthPx > 0 && heightPx > 0
                ? $"width:{widthPx}px;height:{heightPx}px;max-width:100%"
                : "min-width:200px;min-height:100px";
            sb.Append($"<div class=\"ole-placeholder\" style=\"{ph};border:1px dashed #bbb;background:#f5f5f5;display:flex;align-items:center;justify-content:center;color:#888;font-size:13px;margin:8px 0\">");
            sb.Append("Embedded Object (preview not supported in browser)");
            sb.Append("</div>");
        }
    }

    // Footnote/endnote reference tracking is in _ctx.FootnoteRefs / _ctx.EndnoteRefs

    private void RenderFootnotesHtml(StringBuilder sb)
    {
        if (_ctx.FootnoteRefs.Count == 0) return;
        var fnPart = _doc.MainDocumentPart?.FootnotesPart;
        if (fnPart?.Footnotes == null) return;

        var fnSize = ResolveStyleFontSize("FootnoteText") ?? "10pt";
        var fnColor = ResolveStyleColor("FootnoteText");
        var fnColorCss = fnColor != null ? $";color:{fnColor}" : "";
        sb.AppendLine($"<div class=\"footnotes\" style=\"font-size:{fnSize}{fnColorCss}\">");
        sb.AppendLine("<hr style=\"margin-top:0;margin-bottom:0.5em;border:none;border-top:1px solid #ccc;width:33%\">");

        var fnFmt = GetFootnoteNumFmt();
        int num = 0;
        foreach (var fnId in _ctx.FootnoteRefs)
        {
            num++;
            var fn = fnPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
            if (fn == null) continue;

            var fnLabel = FormatNoteNumber(num, fnFmt);
            sb.Append($"<div id=\"fn{fnId}\" style=\"margin:0.3em 0\"><sup>{fnLabel}</sup> ");
            var fnParas = fn.Elements<Paragraph>().ToList();
            for (int pi = 0; pi < fnParas.Count; pi++)
            {
                RenderParagraphContentHtml(sb, fnParas[pi]);
                if (pi < fnParas.Count - 1) sb.Append("<br>");
            }
            sb.AppendLine($" <a href=\"#fnref{fnId}\" style=\"text-decoration:none\">\u21A9</a></div>");
        }
        sb.AppendLine("</div>");
    }

    private void RenderEndnotesHtml(StringBuilder sb)
    {
        if (_ctx.EndnoteRefs.Count == 0) return;
        var enPart = _doc.MainDocumentPart?.EndnotesPart;
        if (enPart?.Endnotes == null) return;

        var enSize = ResolveStyleFontSize("EndnoteText") ?? "10pt";
        sb.AppendLine($"<div class=\"endnotes\" style=\"font-size:{enSize}\">");
        sb.AppendLine("<hr style=\"margin-top:2em;margin-bottom:0.5em;border:none;border-top:1px solid #ccc;width:33%\">");

        var enFmt = GetEndnoteNumFmt();
        int num = 0;
        foreach (var enId in _ctx.EndnoteRefs)
        {
            num++;
            var en = enPart.Endnotes.Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
            if (en == null) continue;

            var enLabel = FormatNoteNumber(num, enFmt);
            var enIndent = ResolveStyleIndent("EndnoteText");
            var enIndentCss = enIndent != null ? $"text-indent:{enIndent}" : "";
            sb.Append($"<div id=\"en{enId}\" style=\"margin:0.3em 0;{enIndentCss}\"><sup>{enLabel}</sup> ");
            var enParas = en.Elements<Paragraph>().ToList();
            for (int pi = 0; pi < enParas.Count; pi++)
            {
                RenderParagraphContentHtml(sb, enParas[pi]);
                if (pi < enParas.Count - 1) sb.Append("<br>");
            }
            sb.AppendLine("</div>");
        }
        sb.AppendLine("</div>");
    }

    /// <summary>Get the numbering format for footnotes (default: decimal per OOXML spec §17.11.11).</summary>
    private string GetFootnoteNumFmt()
    {
        // Priority: section properties > document settings > spec default
        var sectProps = _doc.MainDocumentPart?.Document?.Body
            ?.Descendants<SectionProperties>().LastOrDefault();
        var sectFmt = sectProps?.GetFirstChild<FootnoteProperties>()?.NumberingFormat?.Val?.InnerText;
        if (sectFmt != null) return sectFmt;

        var settingsFmt = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings
            ?.GetFirstChild<FootnoteDocumentWideProperties>()?.NumberingFormat?.Val?.InnerText;
        if (settingsFmt != null) return settingsFmt;

        return "decimal";
    }

    /// <summary>Get the numbering format for endnotes (default: lowerRoman per OOXML spec §17.11.4).</summary>
    private string GetEndnoteNumFmt()
    {
        // Priority: section properties > document settings > spec default
        var sectProps = _doc.MainDocumentPart?.Document?.Body
            ?.Descendants<SectionProperties>().LastOrDefault();
        var sectFmt = sectProps?.GetFirstChild<EndnoteProperties>()?.NumberingFormat?.Val?.InnerText;
        if (sectFmt != null) return sectFmt;

        var settingsFmt = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings
            ?.GetFirstChild<EndnoteDocumentWideProperties>()?.NumberingFormat?.Val?.InnerText;
        if (settingsFmt != null) return settingsFmt;

        return "lowerRoman";
    }

    /// <summary>Format a note number according to Word numbering format.</summary>
    private static string FormatNoteNumber(int num, string fmt)
    {
        return fmt switch
        {
            "lowerRoman" => ToLowerRoman(num),
            "upperRoman" => ToLowerRoman(num).ToUpperInvariant(),
            "lowerLetter" => num >= 1 && num <= 26 ? ((char)('a' + num - 1)).ToString() : num.ToString(),
            "upperLetter" => num >= 1 && num <= 26 ? ((char)('A' + num - 1)).ToString() : num.ToString(),
            _ => num.ToString(), // "decimal" and any other format
        };
    }

    private static string ToLowerRoman(int num)
    {
        if (num <= 0 || num > 3999) return num.ToString();
        var sb = new StringBuilder();
        ReadOnlySpan<(int value, string roman)> map =
        [
            (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
            (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
            (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")
        ];
        foreach (var (value, roman) in map)
        {
            while (num >= value)
            {
                sb.Append(roman);
                num -= value;
            }
        }
        return sb.ToString();
    }
}
