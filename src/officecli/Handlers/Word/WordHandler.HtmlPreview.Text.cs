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
    partial void OnHtmlRenderTab(double widthPt);

    // ==================== Paragraph Content ====================

    private void RenderParagraphHtml(StringBuilder sb, Paragraph para)
    {
        // Use <div> instead of <p> when paragraph contains block-level elements (text boxes, charts, shapes)
        var tag = HasBlockLevelDrawing(para) ? "div" : "p";
        sb.Append($"<{tag}");
        // Add CSS class for TOC paragraphs (suppress hyperlink styling)
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var classes = new List<string>();
        if (styleId != null && styleId.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
            classes.Add("toc");
        // CONSISTENCY(run-special-content): paragraphs containing w:ptab
        // (header/footer left/center/right alignment) need a flex container
        // for the .ptab-spacer / .*-leader children to actually push their
        // siblings apart. The has-ptab class enables display:flex without
        // affecting paragraphs that don't need it.
        if (para.Descendants<PositionalTab>().Any())
            classes.Add("has-ptab");
        if (classes.Count > 0)
            sb.Append($" class=\"{string.Join(" ", classes)}\"");
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

        int commentDepth = 0; // tracks nesting level of open comment ranges

        foreach (var child in para.ChildElements)
        {
            // Comment range begin — open a highlight span and embed comment content
            if (child is CommentRangeStart crs && crs.Id?.Value != null)
            {
                commentDepth++;
                var commentContent = GetCommentDisplayHtml(crs.Id!.Value);
                sb.Append($"<span class=\"comment-highlight\">{commentContent}");
                continue;
            }
            // Comment range end — close the highlight span if we're in a comment
            if (child is CommentRangeEnd cre && commentDepth > 0)
            {
                commentDepth--;
                sb.Append("</span>");
                continue;
            }
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
                // Walk all nested runs so a <w:del> or <w:hyperlink> nested
                // inside <w:ins> doesn't drop its content (Descendants<Run>
                // picks up runs at any depth).
                foreach (var insRun in child.Descendants<Run>())
                    RenderRunHtml(sb, insRun, para);
                // Also render nested deletion text (ins-of-del revision) so
                // the reader sees what was removed within the insertion.
                var nestedDelText = string.Concat(child.Descendants()
                    .Where(e => e.LocalName is "del" or "moveFrom")
                    .SelectMany(d => d.Descendants())
                    .Where(e => e.LocalName is "delText" or "t")
                    .Select(e => e.InnerText));
                if (!string.IsNullOrEmpty(nestedDelText))
                    sb.Append($"<span class=\"track-del\" style=\"text-decoration:line-through;color:#C62828\">{HtmlEncode(nestedDelText)}</span>");
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
                RenderHyperlinkHtml(sb, hyperlink, para);
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                var latex = FormulaParser.ToLatex(child);
                sb.Append($"<span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(latex)}\"></span>");
            }
            else if (child.LocalName is "sdt" or "smartTag" or "customXml" or "fldSimple")
            {
                // Content controls, smart tags, custom XML, simple fields —
                // render hyperlinks with href + their own runs (TOC entries
                // are authored as <w:fldSimple> wrapping <w:hyperlink>),
                // then render bare runs. Runs nested inside a hyperlink are
                // emitted by the hyperlink branch so skip them at the
                // outer Run pass.
                var emittedRuns = new HashSet<OpenXmlElement>();
                foreach (var innerHyp in child.Descendants<Hyperlink>())
                {
                    RenderHyperlinkHtml(sb, innerHyp, para);
                    foreach (var r in innerHyp.Descendants<Run>())
                        emittedRuns.Add(r);
                }
                foreach (var innerRun in child.Descendants<Run>())
                {
                    if (emittedRuns.Contains(innerRun)) continue;
                    RenderRunHtml(sb, innerRun, para);
                }
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

        // VML legacy picture (<w:pict>). The full geometry rendering is
        // deferred (see KNOWN_ISSUES #7e); as a safety net, extract any
        // text content so WordArt strings and textbox text don't vanish
        // from the preview entirely.
        var vmlPict = run.ChildElements.FirstOrDefault(c => c.LocalName == "pict");
        if (vmlPict != null)
        {
            // v:textbox → w:txbxContent → w:t
            var txbxTexts = vmlPict.Descendants().Where(e => e.LocalName == "t").Select(e => e.InnerText);
            // v:textpath string="..." (WordArt / classic watermark)
            var textpathStrings = vmlPict.Descendants()
                .Where(e => e.LocalName == "textpath")
                .Select(e => e.GetAttributes().FirstOrDefault(a => a.LocalName == "string").Value ?? "");
            var text = string.Join(" ", txbxTexts.Concat(textpathStrings).Where(s => !string.IsNullOrWhiteSpace(s)));
            if (!string.IsNullOrWhiteSpace(text))
                sb.Append($"<span class=\"vml-fallback\" style=\"color:#666;font-style:italic\">{HtmlEncode(text)}</span>");
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
            // #8a: when the current section has numRestart=eachSect, the
            // displayed number counts from 1 within that section; otherwise
            // it's the document-wide running total.
            int displayNum;
            if (_ctx.FnRestartEachSection)
            {
                _ctx.FnCountInSection++;
                displayNum = _ctx.FnCountInSection;
            }
            else
            {
                displayNum = _ctx.FootnoteRefs.Count;
            }
            var fnLabel = FormatNoteNumber(displayNum, GetFootnoteNumFmt());
            _ctx.FnLabels[fnId] = fnLabel;
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
            // CONSISTENCY(run-special-content): PositionalTab is rendered as
            // a flex spacer (or leader span) by the ptab branch below — must
            // pass the hasContent gate or the run gets silently early-
            // returned, leaving header/footer left/center/right segments
            // collapsed in the html preview.
            || c is PositionalTab
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
                // TOC-style special case: right-aligned tab with any leader.
                // Dot/hyphen/underscore/middleDot all fill the gap between
                // the current inline position and the right edge of the
                // content box via a flex-grow spacer.
                var rightLeaderTab = tabs?.FirstOrDefault(t =>
                    t.Val?.InnerText == "right"
                    && t.Leader?.InnerText is "dot" or "hyphen" or "underscore" or "middleDot" or "dash" or "heavy");
                if (rightLeaderTab != null)
                {
                    if (needsSpan) { sb.Append("</span>"); needsSpan = false; }
                    var leaderClass = rightLeaderTab.Leader?.InnerText switch
                    {
                        "hyphen" or "dash" => "hyphen-leader",
                        "underscore" or "heavy" => "underscore-leader",
                        "middleDot" => "middledot-leader",
                        _ => "dot-leader",
                    };
                    sb.Append($"<span class=\"{leaderClass}\"></span>");
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
                        // Handle tab leader for positional tabs. OOXML values:
                        //   none, dot, hyphen, underscore, heavy, middleDot (spec)
                        //   some authors also emit "dash" as a hyphen alias.
                        var leader = orderedStops[tabIdx].Leader?.InnerText;
                        var cssLeader = leader switch
                        {
                            "dot" => "border-bottom:1px dotted #000;",
                            // middleDot is centered dot between stops — best CSS equivalent is a
                            // thicker dotted border with larger spacing; browsers render dotted
                            // borders with square dots which read as middle dots at 2px width.
                            "middleDot" => "border-bottom:2px dotted #555;",
                            "hyphen" or "dash" => "border-bottom:1px dashed #000;",
                            "underscore" or "heavy" => "border-bottom:1px solid #000;",
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
            else if (child is PositionalTab ptabChild)
            {
                // CONSISTENCY(run-special-content): w:ptab is the OOXML
                // primitive Word emits in headers/footers to anchor
                // left/center/right alignment regions. Without a render
                // branch the html preview silently dropped these and the
                // three header segments collapsed into a single line.
                // Emit a flex-grow spacer (uses existing leader CSS classes
                // when a leader is set, otherwise a plain ptab-spacer with
                // fallback min-width so the gap is still visible inside
                // non-flex paragraphs). For paragraphs hosting ptabs the
                // outer container is already widened to flex via the
                // has-ptab class added in RenderParagraphHtml.
                if (needsSpan) { sb.Append("</span>"); needsSpan = false; }
                var ptabLeader = ptabChild.Leader?.HasValue == true
                    ? ptabChild.Leader.InnerText : null;
                var ptabClass = ptabLeader switch
                {
                    "dot" => "dot-leader",
                    "hyphen" or "dash" => "hyphen-leader",
                    "underscore" or "heavy" => "underscore-leader",
                    "middleDot" => "middledot-leader",
                    _ => "ptab-spacer",
                };
                sb.Append($"<span class=\"{ptabClass}\"></span>");
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

            // #8a: reuse the label that was stored at ref-emit time so the
            // bottom list matches the superscript. Falls back to the flat
            // running number when the ref emitter didn't cache a label
            // (e.g. footnote referenced from header/footer).
            var fnLabel = _ctx.FnLabels.TryGetValue(fnId, out var cached)
                ? cached
                : FormatNoteNumber(num, fnFmt);
            sb.Append($"<div id=\"fn{fnId}\" style=\"margin:0.3em 0\"><sup>{fnLabel}</sup> ");
            RenderFootnoteChildren(sb, fn);
            sb.AppendLine($" <a href=\"#fnref{fnId}\" style=\"text-decoration:none\">\u21A9</a></div>");
        }
        sb.AppendLine("</div>");
    }

    // Render paragraphs AND tables inside a footnote/endnote. The previous
    // implementation only iterated Elements<Paragraph>() so a footnote with
    // a nested table silently dropped the table (and when a footnote
    // contained only a table, the whole footnote rendered empty).
    private IEnumerable<OpenXmlPart> CollectHyperlinkHostParts()
    {
        var main = _doc.MainDocumentPart;
        if (main == null) yield break;
        yield return main;
        foreach (var hp in main.HeaderParts) yield return hp;
        foreach (var fp in main.FooterParts) yield return fp;
        if (main.FootnotesPart != null) yield return main.FootnotesPart;
        if (main.EndnotesPart != null) yield return main.EndnotesPart;
    }

    private void RenderHyperlinkHtml(StringBuilder sb, Hyperlink hyperlink, Paragraph para)
    {
        var relId = hyperlink.Id?.Value;
        string? url = null;
        if (relId != null)
        {
            // Hyperlink rels can live on the enclosing HeaderPart/FooterPart/
            // FootnotesPart/EndnotesPart, not just MainDocumentPart. Falling
            // back to a full-part sweep keeps header/footer links clickable.
            try
            {
                var parts = CollectHyperlinkHostParts();
                foreach (var part in parts)
                {
                    url = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                    if (url != null) break;
                    url = part.ExternalRelationships.FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                    if (url != null) break;
                }
            }
            catch { }
        }
        if (url == null && hyperlink.Anchor?.Value != null)
            url = $"#{hyperlink.Anchor.Value}";
        var urlSafe = url != null && IsSafeLinkUrl(url);
        if (urlSafe)
            sb.Append($"<a href=\"{HtmlEncodeAttr(url!)}\"{(url!.StartsWith("#") ? "" : " target=\"_blank\"")}>");
        foreach (var descendant in hyperlink.Descendants<Run>())
            RenderRunHtml(sb, descendant, para);
        if (urlSafe)
            sb.Append("</a>");
    }

    private void RenderFootnoteChildren(StringBuilder sb, OpenXmlElement note)
    {
        bool first = true;
        foreach (var child in note.ChildElements)
        {
            if (child is Paragraph p)
            {
                if (!first) sb.Append("<br>");
                RenderParagraphContentHtml(sb, p);
                first = false;
            }
            else if (child is Table tbl)
            {
                RenderTableHtml(sb, tbl);
                first = false;
            }
        }
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
            RenderFootnoteChildren(sb, en);
            sb.AppendLine("</div>");
        }
        sb.AppendLine("</div>");
    }

    /// <summary>
    /// Look up a comment by ID and return its content as an inline HTML
    /// snippet styled as an absolute-positioned tooltip anchored to the
    /// comment highlight span.
    /// </summary>
    private string GetCommentDisplayHtml(string commentId)
    {
        var commentsPart = _doc.MainDocumentPart?.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null) return "";

        var comment = commentsPart.Comments
            .Elements<Comment>()
            .FirstOrDefault(c => c.Id?.Value == commentId);
        if (comment == null) return "";

        var author = HtmlEncode(comment.Author?.Value ?? "匿名");
        var dateVal = comment.Date?.Value;
        var dateStr = dateVal.HasValue ? dateVal.Value.ToString("yyyy-MM-dd HH:mm") : "";
        var initials = HtmlEncode(comment.Initials?.Value ?? "");

        // Collect the comment's paragraph text (may span multiple paragraphs)
        var textSb = new StringBuilder();
        foreach (var para in comment.Elements<Paragraph>())
        {
            if (textSb.Length > 0) textSb.Append("<br>");
            textSb.Append(HtmlEncode(para.InnerText));
        }
        var content = textSb.ToString();
        if (string.IsNullOrEmpty(content)) return "";

        // Build the tooltip span — position:absolute so it floats over the
        // highlighted text; comment-highlight (which is position:relative).
        // When hovering, .comment-highlight:hover .comment-content shows it.
        return $"<span class=\"comment-content\" style=\"position:absolute;" +
               $"background:#fffbe6;border:1px solid #f0d0b0;border-radius:4px;" +
               $"padding:6px 10px;font-size:11px;line-height:1.4;max-width:300px;min-width:222px;" +
               $"color:#333;box-shadow:2px 2px 6px rgba(0,0,0,0.12);white-space:normal;z-index:10\">" +
               $"<span style=\"font-weight:600;color:#e65100\">{initials} {author}</span>" +
               (dateStr.Length > 0 ? $" <span style=\"color:#999;font-size:10px\">{dateStr}</span>" : "") +
               $"<br>{content}" +
               $"</span>";
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
