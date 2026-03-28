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
    // ==================== Paragraph Content ====================

    private void RenderParagraphHtml(StringBuilder sb, Paragraph para)
    {
        // Use <div> instead of <p> when paragraph contains block-level elements (text boxes, charts, shapes)
        var tag = HasBlockLevelDrawing(para) ? "div" : "p";
        sb.Append($"<{tag}");
        var pStyle = GetParagraphInlineCss(para);
        if (!string.IsNullOrEmpty(pStyle))
            sb.Append($" style=\"{pStyle}\"");
        sb.Append(">");
        RenderParagraphContentHtml(sb, para);
        sb.AppendLine($"</{tag}>");
    }

    private void RenderParagraphContentHtml(StringBuilder sb, Paragraph para)
    {
        // Render bookmark anchors for internal hyperlink targets
        foreach (var bm in para.Elements<BookmarkStart>())
        {
            var bmName = bm.Name?.Value;
            if (!string.IsNullOrEmpty(bmName) && !bmName.StartsWith("_GoBack"))
                sb.Append($"<a id=\"{HtmlEncode(bmName)}\"></a>");
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
                // Tracked insertions — render their child runs
                foreach (var insRun in child.Elements<Run>())
                    RenderRunHtml(sb, insRun, para);
            }
            else if (child.LocalName is "del" or "moveFrom")
            {
                // Tracked deletions — skip (deleted content should not be displayed)
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
                    sb.Append($"<a href=\"{HtmlEncode(url)}\"{(url.StartsWith("#") ? "" : " target=\"_blank\"")}>");

                foreach (var hRun in hyperlink.Elements<Run>())
                    RenderRunHtml(sb, hRun, para);

                if (url != null)
                    sb.Append("</a>");
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                var latex = FormulaParser.ToLatex(child);
                sb.Append($"${HtmlEncode(latex)}$");
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

        // Footnote/endnote reference — render superscript number (don't return, run may also have text)
        var fnRef = run.GetFirstChild<FootnoteReference>();
        if (fnRef?.Id?.HasValue == true && fnRef.Id.Value > 0)
        {
            var fnId = (int)fnRef.Id.Value;
            _ctx.FootnoteRefs.Add(fnId);
            var fnNum = _ctx.FootnoteRefs.Count;
            sb.Append($"<sup class=\"fn-ref\"><a href=\"#fn{fnId}\" id=\"fnref{fnId}\">{fnNum}</a></sup>");
        }
        var enRef = run.GetFirstChild<EndnoteReference>();
        if (enRef?.Id?.HasValue == true && enRef.Id.Value > 0)
        {
            var enId = (int)enRef.Id.Value;
            _ctx.EndnoteRefs.Add(enId);
            var enNum = _ctx.EndnoteRefs.Count;
            sb.Append($"<sup class=\"en-ref\"><a href=\"#en{enId}\" id=\"enref{enId}\">{enNum}</a></sup>");
        }
        // FootnoteReferenceMark / EndnoteReferenceMark: don't skip the run, just ignore the mark element
        // (the run may also contain text that should be rendered)

        var hasContent = run.ChildElements.Any(c =>
            c is Break || c is TabChar || c is SymbolChar || c is CarriageReturn
            || c.LocalName is "noBreakHyphen" or "softHyphen"
            || (c is Text t && !string.IsNullOrEmpty(t.Text)));

        if (!hasContent) return;

        var rProps = ResolveEffectiveRunProperties(run, para);
        var style = GetRunInlineCss(rProps);
        var needsSpan = !string.IsNullOrEmpty(style);
        if (needsSpan)
            sb.Append($"<span style=\"{style}\">");

        foreach (var child in run.ChildElements)
        {
            if (child is Break brk)
            {
                if (brk.Type?.Value == BreakValues.Page)
                    sb.Append("<!--PAGE_BREAK-->");
                else
                    sb.Append("<br>");
            }
            else if (child is TabChar)
                sb.Append("&emsp;");
            else if (child is CarriageReturn)
                sb.Append("<br>");
            else if (child.LocalName == "noBreakHyphen")
                sb.Append("\u2011"); // non-breaking hyphen
            else if (child.LocalName == "softHyphen")
                sb.Append("&shy;");
            else if (child is Text t && !string.IsNullOrEmpty(t.Text))
                sb.Append(HtmlEncode(t.Text));
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

        if (needsSpan)
            sb.Append("</span>");
    }

    // Footnote/endnote reference tracking is in _ctx.FootnoteRefs / _ctx.EndnoteRefs

    private void RenderFootnotesHtml(StringBuilder sb)
    {
        if (_ctx.FootnoteRefs.Count == 0) return;
        var fnPart = _doc.MainDocumentPart?.FootnotesPart;
        if (fnPart?.Footnotes == null) return;

        sb.AppendLine("<hr style=\"margin-top:2em;border:none;border-top:1px solid #ccc;width:33%\">");
        sb.AppendLine("<div class=\"footnotes\" style=\"font-size:9pt;color:#555\">");

        int num = 0;
        foreach (var fnId in _ctx.FootnoteRefs)
        {
            num++;
            var fn = fnPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
            if (fn == null) continue;

            sb.Append($"<div id=\"fn{fnId}\" style=\"margin:0.3em 0\"><sup>{num}</sup> ");
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

        sb.AppendLine("<hr style=\"margin-top:2em;border:none;border-top:1px solid #ccc\">");
        sb.AppendLine("<div class=\"endnotes\" style=\"font-size:9pt;color:#555\">");
        sb.AppendLine("<p style=\"font-weight:bold;margin-bottom:0.5em\">Endnotes</p>");

        int num = 0;
        foreach (var enId in _ctx.EndnoteRefs)
        {
            num++;
            var en = enPart.Endnotes.Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
            if (en == null) continue;

            sb.Append($"<div id=\"en{enId}\" style=\"margin:0.3em 0\"><sup>{num}</sup> ");
            var enParas = en.Elements<Paragraph>().ToList();
            for (int pi = 0; pi < enParas.Count; pi++)
            {
                RenderParagraphContentHtml(sb, enParas[pi]);
                if (pi < enParas.Count - 1) sb.Append("<br>");
            }
            sb.AppendLine($" <a href=\"#enref{enId}\" style=\"text-decoration:none\">\u21A9</a></div>");
        }
        sb.AppendLine("</div>");
    }
}
