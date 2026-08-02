// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{

    /// <summary>
    /// Get footnote/endnote text, skipping the reference mark run and its trailing space.
    /// </summary>
    private static string GetFootnoteText(OpenXmlElement fnOrEn)
    {
        // No TrimStart: AddFootnote/AddEndnote no longer prepend a synthetic
        // leading space, so the stored first run is the authored text verbatim
        // and readback must reflect it byte-faithfully (a genuinely authored
        // leading space now survives get/view and the dump round-trip).
        return string.Join("", fnOrEn.Descendants<Run>()
            .Where(r => r.GetFirstChild<FootnoteReferenceMark>() == null
                     && r.GetFirstChild<EndnoteReferenceMark>() == null)
            .SelectMany(r => r.Elements<Text>())
            .Select(t => t.Text));
    }

    private static string GetParagraphText(Paragraph para)
    {
        // CONSISTENCY(run-text-tab): use GetRunText so <w:tab/> renders as
        // \t in the paragraph readback (was silently dropped, breaking
        // dump round-trip for tabbed content).
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(GetRunText(run));
            else if (child is Hyperlink hyperlink)
            {
                // BUG-R4B(BUG7): a hyperlink may nest another hyperlink
                // (<w:hyperlink><w:hyperlink><w:r>…). The old loop only looked
                // at the OUTER hyperlink's direct Run children, so the inner
                // hyperlink's anchor text vanished from view text. Recurse so
                // nested hyperlink runs contribute their text. Read-side only.
                AppendHyperlinkText(sb, hyperlink);
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                // BUG-DUMP9-04: inline equations contribute readable text to the
                // paragraph readback so dump round-trip can verify formula
                // survival. Use raw m:t / w:t descendants (not LaTeX) so the
                // glyphs match the source.
                sb.Append(string.Concat(child.Descendants<Text>().Select(t => t.Text))
                    + string.Concat(child.Descendants<M.Text>().Select(t => t.Text)));
            }
            // BUG-DUMP-R35-2: an inline <w:smartTag>/<w:customXml> wrapper nests
            // its own runs (and may nest further smartTag/customXml/hyperlink
            // levels). Like the nested-hyperlink case above, the old loop only
            // saw the paragraph's DIRECT typed children, so a run wrapped in a
            // smartTag/customXml was invisible to readback (get .text / view
            // text returned ""). Recurse into the wrapper so its inner run text
            // surfaces. Read-side only. Mirrors AppendHyperlinkText.
            else if (IsRunContainerWrapper(child))
                AppendWrapperRunText(sb, child);
            // Tracked-change run containers: <w:ins>/<w:moveTo> wrap runs that
            // ARE part of the visible text (an insertion, or a move's
            // destination), so — like a hyperlink/smartTag wrapper — their
            // inner runs must contribute. The old loop only saw the
            // paragraph's direct Run children, so a run wrapped in <w:ins>
            // vanished from readback (get .text / view text returned ""),
            // leaving a numbered insertion showing its marker with no text.
            // <w:del>/<w:moveFrom> wrap REMOVED text and are intentionally not
            // handled — view text reflects how the document reads once pending
            // changes show through.
            else if (child is InsertedRun || child is MoveToRun)
                AppendRevisionRunText(sb, child);
        }
        return sb.ToString();
    }

    // Walk a <w:ins>/<w:moveTo> run container, recursing into nested runs /
    // hyperlinks / smartTag wrappers / further tracked-change containers so the
    // inserted (or moved-in) text surfaces. Mirrors AppendHyperlinkText.
    private static void AppendRevisionRunText(StringBuilder sb, OpenXmlElement revision)
    {
        foreach (var rChild in revision.ChildElements)
        {
            if (rChild is Run rRun) sb.Append(GetRunText(rRun));
            else if (rChild is Hyperlink rHl) AppendHyperlinkText(sb, rHl);
            else if (rChild is InsertedRun || rChild is MoveToRun) AppendRevisionRunText(sb, rChild);
            else if (IsRunContainerWrapper(rChild)) AppendWrapperRunText(sb, rChild);
            else if (rChild.LocalName == "oMath" || rChild is M.OfficeMath)
                sb.Append(string.Concat(rChild.Descendants<Text>().Select(t => t.Text))
                    + string.Concat(rChild.Descendants<M.Text>().Select(t => t.Text)));
        }
    }

    // BUG-DUMP-R35-2: a <w:smartTag>/<w:customXml> inline wrapper. These parse
    // as OpenXmlUnknownElement in the schema set we load (the strongly-typed
    // SmartTagRun/CustomXmlRun classes aren't present in this SDK build — same
    // observation as the BUG-DUMP5-06/07 unknown-element walk in Navigation), so
    // match by namespace + local name.
    private static bool IsRunContainerWrapper(OpenXmlElement el)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        return el.NamespaceUri == wNs
            && (el.LocalName == "smartTag" || el.LocalName == "customXml");
    }

    // BUG-DUMP-R35-2: walk a smartTag/customXml wrapper's children, recursing
    // into nested Run / Hyperlink / smartTag / customXml so the inner run text
    // (including whitespace-only runs between nested wrappers) survives. Mirrors
    // AppendHyperlinkText's per-child handling.
    private static void AppendWrapperRunText(StringBuilder sb, OpenXmlElement wrapper)
    {
        foreach (var wChild in wrapper.ChildElements)
        {
            if (wChild is Run wRun) sb.Append(GetRunText(wRun));
            else if (wChild is Hyperlink wHl) AppendHyperlinkText(sb, wHl);
            else if (wChild is InsertedRun || wChild is MoveToRun) AppendRevisionRunText(sb, wChild);
            else if (IsRunContainerWrapper(wChild)) AppendWrapperRunText(sb, wChild);
            else if (wChild.LocalName == "oMath" || wChild is M.OfficeMath)
                sb.Append(string.Concat(wChild.Descendants<Text>().Select(t => t.Text))
                    + string.Concat(wChild.Descendants<M.Text>().Select(t => t.Text)));
            // BUG-DUMP-R35-2: smartTag/customXml parse as OpenXmlUnknownElement,
            // so the inner <w:r>/<w:t> are unknown too — GetRunText (typed) and
            // the typed-Run case above miss them. Pull <w:t> text directly from
            // the unknown subtree. Use the actual <w:t> elements' .Text (an
            // OpenXmlUnknownElement.InnerText collapses insignificant whitespace,
            // which would drop a bare " " run); reading each w:t child's text
            // node preserves a space-only run between two nested wrappers.
            else if (wChild is DocumentFormat.OpenXml.OpenXmlUnknownElement)
                AppendUnknownRunText(sb, wChild);
        }
    }

    // BUG-DUMP-R35-2: append the visible text of an unknown-element run subtree
    // (a <w:r> that parsed as OpenXmlUnknownElement because its smartTag/
    // customXml wrapper is unknown). Reads each descendant <w:t> element's first
    // text node so a whitespace-only run is preserved (InnerText on the unknown
    // element trims insignificant whitespace).
    private static void AppendUnknownRunText(StringBuilder sb, OpenXmlElement unknownRun)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        // Recurse into nested unknown wrappers (smartTag>smartTag>r) too.
        if (unknownRun.LocalName == "smartTag" || unknownRun.LocalName == "customXml")
        {
            foreach (var inner in unknownRun.ChildElements)
                if (inner is DocumentFormat.OpenXml.OpenXmlUnknownElement)
                    AppendUnknownRunText(sb, inner);
            return;
        }
        foreach (var tEl in unknownRun.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>())
        {
            if (tEl.NamespaceUri == wNs && tEl.LocalName == "t")
                sb.Append(tEl.InnerText);
        }
    }

    // BUG-R4B(BUG7): walk a hyperlink's children, recursing into nested
    // <w:hyperlink> so inner anchor text survives. Mirrors the per-child run /
    // math handling of GetParagraphText.
    private static void AppendHyperlinkText(StringBuilder sb, Hyperlink hyperlink)
    {
        foreach (var hChild in hyperlink.ChildElements)
        {
            if (hChild is Run hRun) sb.Append(GetRunText(hRun));
            else if (hChild is Hyperlink nested) AppendHyperlinkText(sb, nested);
            else if (hChild is InsertedRun || hChild is MoveToRun) AppendRevisionRunText(sb, hChild);
            else if (hChild.LocalName == "oMath" || hChild is M.OfficeMath)
                sb.Append(string.Concat(hChild.Descendants<Text>().Select(t => t.Text))
                    + string.Concat(hChild.Descendants<M.Text>().Select(t => t.Text)));
        }
    }

    /// <summary>
    /// Get paragraph text including inline math rendered as readable Unicode.
    /// </summary>
    private static string GetParagraphTextWithMath(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(GetRunText(run));
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
                sb.Append(FormulaParser.ToReadableText(child));
            else if (child is Hyperlink hyperlink)
                sb.Append(string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text)));
            // <w:ins>/<w:moveTo> inserted/moved-in runs are visible text; render
            // their content (math included) like the direct-child branches above.
            else if (child is InsertedRun || child is MoveToRun)
                foreach (var rc in child.ChildElements)
                {
                    if (rc is Run rr) sb.Append(GetRunText(rr));
                    else if (rc.LocalName == "oMath" || rc is M.OfficeMath)
                        sb.Append(FormulaParser.ToReadableText(rc));
                    else if (rc is Hyperlink rh)
                        sb.Append(string.Concat(rh.Descendants<Text>().Select(t => t.Text)));
                }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find math elements in a paragraph using both type and localName matching.
    /// </summary>
    private static List<OpenXmlElement> FindMathElements(Paragraph para)
    {
        return para.ChildElements
            .Where(e => e.LocalName == "oMath" || e is M.OfficeMath)
            .ToList();
    }

    /// <summary>
    /// Get all body-level elements, flattening SdtContent containers.
    /// This ensures paragraphs and tables inside w:sdt are not missed.
    /// </summary>
    private static IEnumerable<OpenXmlElement> GetBodyElements(Body body)
    {
        foreach (var element in FlattenWrappers(body.ChildElements))
            yield return element;
    }

    // Descend into SDT (structured document tag) and customXml transparent
    // wrappers so their wrapped paragraphs/tables participate in the body
    // element axis. Without this, docs emitted by e.g. Pages/Google Docs
    // that wrap entire sections in <w:customXml> produce an empty preview.
    private static IEnumerable<OpenXmlElement> FlattenWrappers(IEnumerable<OpenXmlElement> elements)
    {
        foreach (var element in elements)
        {
            if (element is SdtBlock sdt)
            {
                var content = sdt.SdtContentBlock;
                if (content != null)
                    foreach (var child in FlattenWrappers(content.ChildElements))
                        yield return child;
            }
            else if (element.LocalName == "customXml"
                && element.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            {
                foreach (var child in FlattenWrappers(element.ChildElements))
                    yield return child;
            }
            else
            {
                yield return element;
            }
        }
    }

    /// <summary>
    /// Checks if an element is a structural document element worth displaying
    /// (not inline markers like bookmarkStart, bookmarkEnd, proofErr, etc.)
    /// </summary>
    private static bool IsStructuralElement(OpenXmlElement element)
    {
        var name = element.LocalName;
        return name == "sectPr" || name == "altChunk" || name == "customXml";
    }

    /// <summary>
    /// Get all Run elements in a paragraph, including those nested inside
    /// Hyperlink and SdtContent containers.
    /// </summary>
    private static List<Run> GetAllRuns(Paragraph para)
    {
        return para.Descendants<Run>()
            .Where(r => r.GetFirstChild<CommentReference>() == null)
            // BUG-DUMP-RUBY: a <w:ruby> (CJK phonetic guide — furigana ABOVE
            // the base text) wraps independent <w:r>s in its <w:rt> (furigana)
            // and <w:rubyBase> (base) — both paragraph Descendant Runs. Drop
            // ONLY those inner runs so they don't flatten かんじ-above-漢字
            // into the sequential plain runs "かんじ" + "漢字". The OUTER
            // ruby-bearing run (a <w:r> directly containing <w:ruby>) is kept:
            // text rendering surfaces its base text via GetRunText, and the
            // dump path routes it to a verbatim raw-set (so <w:rt>/<w:rubyBase>/
            // <w:rubyPr> survive) instead of emitting it as a plain run. A
            // revision-wrapped ruby (<w:ins>/<w:del> ancestor) is left intact
            // for the legacy DUMP7-10 flatten-with-attribution path.
            .Where(r =>
            {
                var rubyAnc = r.Ancestors<Ruby>().FirstOrDefault();
                if (rubyAnc != null
                    && rubyAnc.Ancestors<InsertedRun>().FirstOrDefault() == null
                    && rubyAnc.Ancestors<DeletedRun>().FirstOrDefault() == null)
                    return false; // an inner <w:rt>/<w:rubyBase> run
                return true;
            })
            // BUG-DUMP-R42-9: a <w:bdo> (bidirectional override — forces visual
            // RTL/LTR character ordering of its wrapped runs) is a run-container,
            // not a run. Its inner <w:r>s are paragraph Descendant Runs; drop them
            // here so they don't flatten into sequential plain runs, dropping the
            // <w:bdo> wrapper (which carries the load-bearing w:val direction).
            // The wrapper is surfaced separately as a "bdo" paragraph child whose
            // verbatim XML the emitter raw-sets — mirrors the ruby path above.
            .Where(r => r.Ancestors<BidirectionalOverride>().FirstOrDefault() == null)
            // BUG-DUMP-R43-7: a <w:dir> (BidirectionalEmbedding — bidirectional
            // embedding direction wrapper, distinct from <w:bdo> override and the
            // run-level <w:rtl> toggle) is likewise a run-container. Drop its inner
            // runs here so they don't flatten into plain runs, dropping the <w:dir>
            // wrapper (which carries the load-bearing w:val direction). The wrapper
            // is surfaced separately as a "dir" paragraph child whose verbatim XML
            // the emitter raw-sets — mirrors the bdo/ruby path above.
            .Where(r => r.Ancestors<BidirectionalEmbedding>().FirstOrDefault() == null)
            // BUG-DUMP4-06: skip runs nested inside an inline SdtRun. Those
            // runs are surfaced separately as a typed `sdt` paragraph child so
            // alias/tag/type metadata round-trips. Without this filter the
            // inner run was emitted twice — once unwrapped (losing metadata)
            // and once via the sdt branch.
            .Where(r => r.Ancestors<SdtRun>().FirstOrDefault() == null)
            // BUG-DUMP6-01: skip runs nested inside <w:fldSimple>. Those
            // runs are surfaced separately as a typed `field` paragraph child
            // carrying the SimpleField.Instruction attribute. Without this
            // filter the inner display run was emitted as a plain run and
            // the field instruction was silently dropped on dump round-trip.
            .Where(r => r.Ancestors<SimpleField>().FirstOrDefault() == null)
            // BUG-DUMP-TXBX: skip runs whose nearest TextBoxContent ancestor
            // sits BELOW the current paragraph (i.e. the run lives inside a
            // textbox that is a descendant of `para`). Those runs are
            // surfaced separately under /<host>/textbox[N]/p[M]/r[K] via the
            // textbox navigation branch and the WordBatchEmitter typed
            // `add textbox` recursion. We must NOT skip runs whose para is
            // itself inside TextBoxContent (the inner paragraphs of a
            // textbox) — for those, no TextBoxContent sits between the run
            // and `para`, so they pass through and emit normally.
            .Where(r =>
            {
                // Drop the run iff its nearest TextBoxContent ancestor is a
                // DESCENDANT of `para` (a textbox lives under this para and
                // this run sits inside it). Keep when no TextBoxContent
                // exists, or when the TextBoxContent ancestor sits at-or-
                // above `para` (meaning `para` itself is the textbox-inner
                // paragraph — emitting its runs is the desired behavior).
                var tbc = r.Ancestors<TextBoxContent>().FirstOrDefault();
                if (tbc == null) return true;
                // tbc is a descendant of `para`? walk tbc's ancestors and
                // check whether `para` is among them.
                foreach (var anc in tbc.Ancestors())
                {
                    if (ReferenceEquals(anc, para)) return false;
                }
                return true;
            })
            // BUG-DUMP-ALTCONTENT-DOUBLE: a run inside an <mc:AlternateContent>
            // (a WPS/DrawingML shape with a VML <mc:Fallback>) is NOT caught by
            // the typed TextBoxContent skip above — the SDK parses the
            // AlternateContent subtree as OpenXmlUnknownElement, so its inner
            // <w:txbxContent> is not a typed TextBoxContent and Ancestors<>()
            // misses it. The drawing run itself is round-tripped VERBATIM via a
            // raw-set (textbox/drawing emit), so surfacing its inner runs as
            // plain runs duplicated the text — and BOTH the mc:Choice and the
            // mc:Fallback branch hold the SAME text, so a single shape's text
            // ("Australia – Indonesia…") appeared up to FOUR times in the body.
            // Drop every run whose ancestor chain (below `para`) crosses an
            // AlternateContent wrapper; the verbatim raw-set carries it.
            .Where(r =>
            {
                foreach (var anc in r.Ancestors())
                {
                    if (ReferenceEquals(anc, para)) break; // reached host para
                    if (anc.LocalName == "AlternateContent"
                        || anc.LocalName == "Choice"
                        || anc.LocalName == "Fallback")
                        return false;
                }
                return true;
            })
            .ToList();
    }

    private static string GetRunText(Run run)
    {
        // BUG-DUMP-RUBY: a ruby-bearing run (<w:r><w:ruby>…) carries no direct
        // <w:t>; its visible inline text is the <w:rubyBase> base text (the
        // furigana in <w:rt> renders ABOVE the base, not inline). GetAllRuns
        // keeps the outer ruby run and drops the inner ones, so surface the
        // base text here for view/readback. Mirrors the HtmlPreview ruby path
        // (WordHandler.HtmlPreview.Text.cs), which renders base + <rt>.
        var rubyEl = run.GetFirstChild<Ruby>();
        if (rubyEl != null)
        {
            var rubyBase = rubyEl.ChildElements.FirstOrDefault(c => c.LocalName == "rubyBase");
            return rubyBase == null ? "" :
                string.Concat(rubyBase.Descendants<Text>().Select(t => t.Text));
        }
        // CONSISTENCY(run-text-tab): walk run children in document order so
        // <w:tab/> renders as \t in the readback. Plain Elements<Text>() drops
        // tabs silently, which broke dump round-trip (the tab IS in the XML
        // because AddText splits on \t and emits TabChar — but Get hid it).
        var sb = new System.Text.StringBuilder();
        foreach (var child in run.Elements())
        {
            switch (child)
            {
                case Text t: sb.Append(t.Text); break;
                case TabChar: sb.Append('\t'); break;
                // CONSISTENCY(text-breaks): mirror AppendTextWithBreaks — \n
                // round-trips through <w:br/> (textWrapping, the OOXML default
                // when w:type is absent). Without this case, Set/Add(text=...)
                // with embedded \n loses the break on dump readback. Skip
                // page/column breaks — they have no \n source representation
                // and a paragraph-level `break` property already captures them.
                // NEWLINE-SEMANTICS-V2: soft line breaks read back as '\v'
                // (Word object model Chr(11)); '\n' now denotes a paragraph
                // boundary at block level, so a break must not masquerade as
                // one. AppendTextWithBreaks accepts '\v' (and legacy '\n' in
                // run scope) back into <w:br/> for the dump round-trip.
                case Break br when br.Type == null || br.Type.Value == BreakValues.TextWrapping:
                    sb.Append('\v'); break;
                // BUG-R10A(BUG1): <w:cr/> (CarriageReturn) is a line break too —
                // same visual effect as a textWrapping <w:br/>. Without this case
                // GetRunText dropped it and adjacent text merged ("A"+"B" → "AB")
                // on dump readback. Map to \n so it round-trips through a <w:br/>.
                case CarriageReturn: sb.Append('\v'); break;
                // BUG-DUMP7-01: <w:sym w:font="Wingdings" w:char="F0E0"/> is a
                // glyph substitution — the run carries no <w:t>. Without a case
                // here, GetRunText returned empty and WordBatchEmitter's run-emit
                // dropped the whole run, silently losing the symbol on dump
                // round-trip. Surface the resolved Unicode code point as Text
                // so the run looks non-empty; the canonical `sym` Format key
                // (set in Navigation.cs) carries the font+char metadata that
                // AddRun consumes to rebuild the SymbolChar element verbatim.
                case SymbolChar symChild:
                {
                    var charHex = symChild.Char?.Value;
                    if (!string.IsNullOrEmpty(charHex)
                        && int.TryParse(charHex, System.Globalization.NumberStyles.HexNumber,
                            System.Globalization.CultureInfo.InvariantCulture, out var symCode))
                        sb.Append(char.ConvertFromUtf32(symCode));
                    break;
                }
                // BUG-DUMP4-01: a Run nested inside a w:del wrapper carries its
                // text in <w:delText> (DeletedText), not <w:t>. Without this
                // case the deleted content was silently dropped from Get
                // readback and dump round-trip — the inner Run was reachable
                // via Descendants<Run>() but appeared empty.
                case DeletedText dt: sb.Append(dt.Text); break;
                // BUG-DUMP5-03: inline character elements that carry no <w:t>
                // child but contribute visible glyphs. Map to their Unicode
                // equivalents so dump→batch round-trip preserves the visible
                // text. Without this, every <w:noBreakHyphen/> / <w:softHyphen/>
                // dropped to an empty run and disappeared on replay.
                case NoBreakHyphen: sb.Append('‑'); break; // non-breaking hyphen
                case SoftHyphen: sb.Append('­'); break;   // soft hyphen
                // BUG-DUMP5-04: date / time placeholder elements (dayLong /
                // monthLong / yearShort / dayShort / monthShort / yearLong)
                // are auto-substituted by Word at render time. They carry no
                // text in OOXML — surface a stable placeholder so dump
                // captures their presence (otherwise the runs vanish on
                // round-trip and Word has nothing to substitute against).
                case DayLong: sb.Append("[dayLong]"); break;
                case DayShort: sb.Append("[dayShort]"); break;
                case MonthLong: sb.Append("[monthLong]"); break;
                case MonthShort: sb.Append("[monthShort]"); break;
                case YearLong: sb.Append("[yearLong]"); break;
                case YearShort: sb.Append("[yearShort]"); break;
            }
        }
        return sb.ToString();
    }

    private static bool HasMixedPunctuation(string text)
    {
        var chinesePunct = "\uff0c\u3002\uff01\uff1f\u3001\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011";
        bool hasChinese = text.Any(c => chinesePunct.Contains(c));
        bool hasEnglish = text.Any(c => ",.!?;:\"'()[]".Contains(c));
        bool hasChineseChars = text.Any(c => c >= 0x4E00 && c <= 0x9FFF);
        return hasChinese && hasEnglish && hasChineseChars;
    }

    // BUG-DUMP-BMSPAN: a bookmark "wraps content" when its range
    // (BookmarkStart … matching BookmarkEnd, same w:id) contains at least one
    // content element (Run / equation / field) — possibly spanning paragraphs.
    // Such bookmarks must round-trip as two positioned markers (start before
    // the content, end after), or the range collapses to zero length and every
    // REF/PAGEREF/TOC anchor pointing at it resolves to nothing. A zero-length
    // bookmark (End is the immediate document-order successor of Start) is NOT
    // a span and keeps the single combined `add bookmark` op so it stays empty.
    // Document-order (pre-order) successors of `start` that stay within `root`.
    // Replaces `root.Descendants()` + skip-until-started: that re-walked the whole
    // subtree from the top on every call just to REACH bkStart, so classifying N
    // bookmarks was O(N * position) = O(N²). Walking forward from bkStart is
    // O(span) — bounded by the first content element / matching end, which for a
    // typical span is the very next node.
    private static IEnumerable<OpenXmlElement> ForwardWithin(OpenXmlElement start, OpenXmlElement root)
    {
        var n = start;
        while (true)
        {
            OpenXmlElement? next = n.FirstChild ?? n.NextSibling();
            while (next == null)
            {
                n = n.Parent;
                if (n == null || ReferenceEquals(n, root)) yield break;
                next = n.NextSibling();
            }
            yield return next;
            n = next;
        }
    }

    private static bool IsContentSpanBookmark(BookmarkStart bkStart)
    {
        var id = bkStart.Id?.Value;
        if (string.IsNullOrEmpty(id)) return false;
        var root = bkStart.Ancestors<Body>().FirstOrDefault() as OpenXmlElement
            ?? bkStart.Ancestors<TableCell>().FirstOrDefault() as OpenXmlElement
            ?? bkStart.Ancestors().LastOrDefault();
        if (root == null) return false;
        foreach (var el in ForwardWithin(bkStart, root))
        {
            if (el is BookmarkEnd be && be.Id?.Value == id)
            {
                // BUG-DUMP-BMSDT-DUP: the matching end lives INSIDE an SDT (raw-set
                // verbatim, so it carries the end) while the start is OUTSIDE that
                // SDT. Treat it as a span so the emitter emits an open=true start
                // (reusing the id) that pairs with the SDT's verbatim end — instead
                // of a self-contained `add bookmark` that auto-creates a SECOND,
                // orphan bookmarkEnd (duplicate marker + REB-validate>SRC). A
                // genuinely adjacent end (not separated into an SDT) stays empty.
                var endSdt = be.Ancestors().FirstOrDefault(a => a is SdtBlock || a is SdtRun);
                if (endSdt != null
                    && !bkStart.Ancestors().Any(a => ReferenceEquals(a, endSdt)))
                    return true;
                // BUG-DUMP-BMCELL-DUP: the matching end lives OUTSIDE the start's
                // table cell — a zero-length bookmark straddling a cell/row boundary
                // (start the last child of a cell paragraph, end a <w:tr>-level child
                // between cells, raw-injected verbatim by GetTableStructuralBookmarks
                // which carries the end). Same shape as the SDT case above: treat it
                // as a span so the emitter emits an open=true start (reusing the id)
                // that pairs with the verbatim structural end — NOT a self-contained
                // `add bookmark` that auto-creates a SECOND, orphan bookmarkEnd
                // (duplicate/unbalanced marker, silent: the validator does not flag
                // it). A genuinely adjacent end in the SAME cell stays empty.
                var startCell = bkStart.Ancestors<TableCell>().FirstOrDefault();
                if (startCell != null
                    && !be.Ancestors().Any(a => ReferenceEquals(a, startCell)))
                    return true;
                return false; // adjacent → empty
            }
            // Content between Start and End → this is a wrapping span.
            if (el is Run || el is M.OfficeMath || el is M.Paragraph
                || el is SimpleField || el is Hyperlink)
                return true;
        }
        return false;
    }

    // Cached w:id → BookmarkStart lookup within a scope root (per-Body). Replaces
    // body.Descendants<BookmarkStart>().FirstOrDefault(id), which is O(bookmarks)
    // per call; dump resolves one per bookmarkEnd, so the scan made dump O(N²) on
    // bookmark-dense documents. First-wins on duplicate ids, matching the old
    // FirstOrDefault. Invalidated with the other body caches on any structural
    // mutation (ClearBodyChildIndex → _bookmarkStartByIdCache = null).
    private BookmarkStart? FindBookmarkStartById(OpenXmlElement scopeRoot, string id)
    {
        _bookmarkStartByIdCache ??= new();
        if (!_bookmarkStartByIdCache.TryGetValue(scopeRoot, out var map))
        {
            map = new(StringComparer.Ordinal);
            foreach (var bs in scopeRoot.Descendants<BookmarkStart>())
                if (bs.Id?.Value is { } bid && !map.ContainsKey(bid))
                    map[bid] = bs;
            _bookmarkStartByIdCache[scopeRoot] = map;
        }
        return map.TryGetValue(id, out var found) ? found : null;
    }

    // Overload: classify by the BookmarkEnd half (resolve its paired Start).
    private bool IsContentSpanBookmark(BookmarkEnd bkEnd)
    {
        var id = bkEnd.Id?.Value;
        if (string.IsNullOrEmpty(id)) return false;
        var body = _doc.MainDocumentPart?.Document?.Body;
        var start = body == null ? null : FindBookmarkStartById(body, id);
        return start != null && IsContentSpanBookmark(start);
    }

    // Resolve the w:name of the BookmarkStart paired with this BookmarkEnd
    // (matched by w:id) so a standalone end marker can be addressed by name.
    private string? ResolveBookmarkEndName(BookmarkEnd bkEnd)
    {
        var id = bkEnd.Id?.Value;
        if (string.IsNullOrEmpty(id)) return null;
        var body = _doc.MainDocumentPart?.Document?.Body;
        var start = body == null ? null : FindBookmarkStartById(body, id);
        return start?.Name?.Value;
    }

    private static string GetBookmarkText(BookmarkStart bkStart)
    {
        var bkId = bkStart.Id?.Value;
        if (bkId == null) return "";

        var sb = new System.Text.StringBuilder();
        var sibling = bkStart.NextSibling();
        while (sibling != null)
        {
            if (sibling is BookmarkEnd bkEnd && bkEnd.Id?.Value == bkId)
                break;
            if (sibling is Run run)
                sb.Append(string.Concat(run.Descendants<Text>().Select(t => t.Text)));
            sibling = sibling.NextSibling();
        }
        return sb.ToString();
    }
}
