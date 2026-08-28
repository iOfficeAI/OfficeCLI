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

    // Tolerant BCP-47 shape used to validate run lang.{val,ea,cs} values.
    // RFC 5646 §2.1: language tag is primary (2-3 ALPHA, or 4-8 ALPHA "reserved"
    // for future / "registered"), followed by hyphen-separated subtags each
    // CONSISTENCY(bcp47-validation): shape regex moved to Core/Bcp47LanguageTag.cs
    // so docx and pptx use the same validator. Wrappers kept for readability
    // of the call sites below.
    private const int LangBcp47MaxLength = OfficeCli.Core.Bcp47LanguageTag.MaxLength;
    private static bool LangBcp47IsValid(string value) => OfficeCli.Core.Bcp47LanguageTag.IsValid(value);

    /// <summary>
    /// Resolve the OpenXmlPart that owns a given element. Returns the
    /// HeaderPart/FooterPart/FootnotesPart/EndnotesPart/CommentsPart when the
    /// element lives inside one of those parts, falling back to MainDocumentPart.
    /// Used for part-local relationships like hyperlinks that must be added to
    /// the host part's rels file (e.g. word/_rels/header1.xml.rels) rather than
    /// the document rels.
    /// </summary>
    private OpenXmlPart ResolveHostPart(OpenXmlElement element)
    {
        var main = _doc.MainDocumentPart!;
        // Walk to the part-root element (Header/Footer/Footnotes/Endnotes/Comments/Document).
        // BUG-R14A: use Self-or-ancestor — `element` may itself BE the part-root
        // (e.g. a picture appended directly to /header[1] passes the Header as
        // `element`), in which case Ancestors<Header>() would miss it and we'd
        // wrongly fall through to MainDocumentPart.
        var hdr = element as Header ?? element.Ancestors<Header>().FirstOrDefault();
        if (hdr != null)
        {
            var hp = main.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, hdr));
            if (hp != null) return hp;
        }
        var ftr = element as Footer ?? element.Ancestors<Footer>().FirstOrDefault();
        if (ftr != null)
        {
            var fp = main.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, ftr));
            if (fp != null) return fp;
        }
        // Footnote/Endnote: parts live on MainDocumentPart.FootnotesPart / EndnotesPart
        if ((element is Footnote || element.Ancestors<Footnote>().Any()) && main.FootnotesPart != null)
            return main.FootnotesPart;
        if ((element is Endnote || element.Ancestors<Endnote>().Any()) && main.EndnotesPart != null)
            return main.EndnotesPart;
        // BUG-R13B(BUG2): a hyperlink added into a comment body must register its
        // external relationship on the comments part (word/_rels/comments.xml.rels),
        // not document.xml.rels — otherwise the w:hyperlink r:id living in
        // word/comments.xml dangles and the document fails validation. Mirrors the
        // footnote/endnote host-part resolution above.
        if ((element is Comment || element.Ancestors<Comment>().Any()) && main.WordprocessingCommentsPart != null)
            return main.WordprocessingCommentsPart;
        return main;
    }

    /// <summary>
    /// Resolve a hyperlink relationship by id, searching the element's host
    /// part first, then falling back to MainDocumentPart and other host parts.
    /// </summary>
    private HyperlinkRelationship? ResolveHyperlinkRelationship(OpenXmlElement element, string relId)
    {
        var host = ResolveHostPart(element);
        var rel = host.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
        if (rel != null) return rel;
        // Fallback: scan MainDocumentPart and all header/footer parts (handles
        // documents authored with rels in unexpected places).
        var main = _doc.MainDocumentPart!;
        if (!ReferenceEquals(host, main))
        {
            rel = main.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        foreach (var hp in main.HeaderParts)
        {
            if (ReferenceEquals(hp, host)) continue;
            rel = hp.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        foreach (var fp in main.FooterParts)
        {
            if (ReferenceEquals(fp, host)) continue;
            rel = fp.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        return null;
    }

    // ==================== Private Helpers ====================

    /// <summary>
    /// Format twips as a human-readable cm string (e.g., "21cm").
    /// 1 inch = 1440 twips, 1 inch = 2.54 cm.
    /// </summary>
    private static string FormatTwipsToCm(uint twips)
    {
        var cm = twips * 2.54 / 1440.0;
        return $"{cm:0.##}cm";
    }

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    /// <summary>
    /// BUG #334: true when a run carries non-text content — a drawing
    /// (<c>w:drawing</c>), a VML picture (<c>w:pict</c>), an
    /// <c>mc:AlternateContent</c> shape wrapper, an embedded object
    /// (<c>w:object</c>), or a field (<c>w:fldChar</c> / <c>w:instrText</c>).
    /// Paragraph-level <c>set --prop text=</c> must preserve such runs instead
    /// of wiping them, or it silently deletes every graphic in the paragraph.
    /// </summary>
    private static bool RunCarriesNonText(Run run) =>
        run.ChildElements.Any(c =>
            c is Drawing or Picture or EmbeddedObject or FieldChar or FieldCode
            || c is AlternateContent);

    /// <summary>
    /// BUG-R7-07: a value the user explicitly typed as "false"/"0"/"off" — not
    /// just any non-truthy input (null/empty count as "no override"). Used by
    /// AddParagraph's no-text fallback to decide whether to emit
    /// <c>&lt;w:b w:val="false"/&gt;</c> as an explicit style override vs.
    /// simply removing the element. Set-style call sites continue to use
    /// ApplyRunFormatting's "remove on falsy" semantics so existing tests
    /// (R25/R26 EmptyRpr_NotSurfaced) keep passing.
    /// </summary>
    internal static bool IsExplicitFalseAddOverride(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        var v = value.Trim().ToLowerInvariant();
        return v is "false" or "0" or "off" or "no" or "n";
    }

    /// <summary>
    /// Read a w:val OnOff attribute defensively. Returns null when the
    /// attribute is absent OR when the stored text is not a valid OnOff
    /// token (e.g. <c>&lt;w:bidi w:val="garbage"/&gt;</c>). Default-on
    /// elements (BiDi, Bold, etc.) are conventionally treated as true
    /// when Val is null. R8-fuzz-5: prevents OnOffValue.Parse from
    /// crashing Get/HtmlPreview on a document that loaded fine but
    /// disk-stored a malformed attribute.
    /// </summary>
    internal static bool? TryReadOnOff(DocumentFormat.OpenXml.OnOffValue? val)
    {
        if (val == null) return true; // default-on: <w:bidi/> with no Val
        try { return val.Value; }
        catch (FormatException) { return null; }
    }

    /// <summary>
    /// Find the 1-based run index inside the anchor paragraph where the
    /// CommentRangeStart with <paramref name="commentId"/> sits — i.e. the
    /// number of runs before the range marker plus 1. Returns 0 when the
    /// range marker is not found, or sits before any Run (anchor at paragraph
    /// start).
    /// BUG-DUMP4-03: callers (WordBatchEmitter) need this so dump can preserve
    /// intra-paragraph anchor position; without it replay widens every
    /// comment to the whole paragraph.
    /// </summary>
    public int FindCommentAnchorRunIndex(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;
        foreach (var para in body.Descendants<Paragraph>())
        {
            var rs = para.Descendants<CommentRangeStart>()
                .FirstOrDefault(r => r.Id?.Value == commentId);
            if (rs == null) continue;
            // Count Run elements that appear before the CommentRangeStart in
            // document order within the same paragraph.
            int runCount = 0;
            foreach (var el in para.Descendants())
            {
                if (ReferenceEquals(el, rs)) break;
                if (el is Run r && r.GetFirstChild<CommentReference>() == null) runCount++;
            }
            return runCount; // 0 = before any run; N = after run N (1-based)
        }
        return 0;
    }

    /// <summary>
    /// BUG-DUMP-R26-3: locate the paragraph that holds the
    /// <c>CommentRangeEnd</c> for <paramref name="commentId"/>, returning its
    /// semantic body path plus the 1-based run index after which the end marker
    /// sits (0 = before all runs). Returns null when the comment has no range
    /// end (point comment) or it can't be resolved. Mirrors the
    /// FindCommentAnchorRunIndex / FindCommentAnchorPath pair for the START side
    /// so a multi-paragraph comment range round-trips its full span instead of
    /// collapsing into the start paragraph.
    /// </summary>
    public (string path, int runIndex)? FindCommentRangeEnd(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;
        foreach (var para in body.Descendants<Paragraph>())
        {
            var re = para.Descendants<CommentRangeEnd>()
                .FirstOrDefault(r => r.Id?.Value == commentId);
            if (re == null) continue;
            // Count Run elements before the CommentRangeEnd in document order,
            // excluding the comment-reference run itself (it is inserted after
            // the end marker on replay, so it must not shift the index).
            int runCount = 0;
            foreach (var el in para.Descendants())
            {
                if (ReferenceEquals(el, re)) break;
                if (el is Run r && r.GetFirstChild<CommentReference>() == null) runCount++;
            }
            var p = BuildBodyParagraphFullPath(body, para);
            if (p == null) return null;
            return (p, runCount);
        }
        return null;
    }

    /// <summary>
    /// BUG-DUMP-COMMENT-POINTREF: true when the comment with
    /// <paramref name="commentId"/> has a span — i.e. a CommentRangeStart marker
    /// exists in the body. A reference-only (zero-width / point) comment carries
    /// only a CommentReference run, so this returns false; the dump emitter then
    /// passes range=false so replay doesn't synthesize a spurious range.
    /// </summary>
    public bool CommentHasRange(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return false;
        return body.Descendants<CommentRangeStart>()
            .Any(rs => rs.Id?.Value == commentId);
    }

    /// <summary>
    /// Find the paragraph path where a CommentRangeStart with the given ID is anchored.
    /// Returns "/body/p[N]" for a top-level body paragraph, or the full semantic
    /// table path "/body/tbl[N]/tr[M]/tc[K]/p[J]" when the anchor lives inside a
    /// table cell. Returns null if not found.
    /// BUG-R4 (DBF-R4-01): previously iterated only body.Elements&lt;Paragraph&gt;()
    /// (top-level body paragraphs), so a comment anchored inside a table cell
    /// resolved to null — Query never set Format["anchoredTo"] and EmitComments
    /// fell back to the hard-coded "/body/p[1]", relocating the comment out of
    /// the cell on dump→batch. Resolve through the same full-document paragraph
    /// scope its sibling FindCommentAnchorRunIndex already uses (Descendants),
    /// and build the correct cell path so the comment re-anchors in the cell.
    /// </summary>
    private string? FindCommentAnchorPath(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;

        foreach (var para in body.Descendants<Paragraph>())
        {
            var hasRange = para.Descendants<CommentRangeStart>()
                .Any(rs => rs.Id?.Value == commentId);
            if (hasRange) return BuildBodyParagraphFullPath(body, para);
        }
        // BUG-DUMP-COMMENT-POINTREF: a reference-only (point) comment has no
        // CommentRangeStart, so fall back to the paragraph that hosts its
        // CommentReference run. Without this the point comment loses its anchor
        // and dump relocates it to /body/p[1].
        foreach (var para in body.Descendants<Paragraph>())
        {
            var hasRef = para.Descendants<CommentReference>()
                .Any(cr => cr.Id?.Value == commentId);
            if (hasRef) return BuildBodyParagraphFullPath(body, para);
        }
        return null;
    }

    /// <summary>
    /// BUG-R4 (DBF-R4-02): if the paragraph at <paramref name="paraPath"/> is a
    /// pure display-equation wrapper (a <c>w:p</c> whose only content is
    /// <c>m:oMathPara</c>), return the typed equation DocumentNode (mode/formula/
    /// align) for it; otherwise null. The body walker reaches this via the
    /// dedicated /body/oMathPara[N] addressing, but a cell oMathPara surfaces in
    /// the cell-content walker as a plain paragraph path whose Get returns an
    /// empty paragraph — so EmitTable could not detect it. This lets the cell
    /// walker re-route the wrapper to a typed `add equation`, mirroring the body
    /// walker's behavior, without changing the cell path addressing scheme.
    /// </summary>
    internal DocumentNode? TryGetDisplayEquationAtParagraph(string paraPath)
    {
        var element = NavigateToElement(ParsePath(paraPath));
        if (element is not Paragraph para || !IsOMathParaWrapperParagraph(para))
            return null;
        var inner = para.ChildElements.FirstOrDefault(
            c => c.LocalName == "oMathPara" || c is M.Paragraph);
        if (inner == null) return null;
        // Reuse ElementToNode's oMathPara → typed equation projection so the
        // mode/formula/align extraction stays single-sourced.
        return ElementToNode(inner, paraPath, 0);
    }

    /// <summary>
    /// Build a "/body"-rooted path to a paragraph by walking its ancestor chain,
    /// emitting tbl[i]/tr[j]/tc[k] segments for every enclosing table cell.
    /// Top-level body paragraphs yield "/body/p[N]"; cell paragraphs yield
    /// "/body/tbl[N]/tr[M]/tc[K]/p[J]". Mirrors BuildOleRunPath's table-aware
    /// ancestor walk. Returns null if the paragraph is not under the given body.
    /// </summary>
    private static string? BuildBodyParagraphFullPath(Body body, Paragraph para)
    {
        // Ancestors() returns innermost first; reverse to outer-to-inner order.
        var ancestors = para.Ancestors().TakeWhile(a => a != body).Reverse().ToList();

        var sb = new StringBuilder("/body");
        OpenXmlElement cursor = body;
        foreach (var anc in ancestors)
        {
            if (anc is DocumentFormat.OpenXml.Wordprocessing.Table tblAnc)
            {
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
        }

        var paraIdx = cursor.Elements<Paragraph>().TakeWhile(p => p != para).Count() + 1;
        if (!ReferenceEquals(cursor.Elements<Paragraph>().ElementAtOrDefault(PathIndex.ToArrayIndex(paraIdx)), para))
            return null; // paragraph not a direct child of the resolved cursor
        // Top-level body paragraphs use the paraId form so EmitComments can map
        // the source paraId -> target index via paraIdToTargetIdx. Cell paragraphs
        // are positionally stable across dump→batch (the table is re-created with
        // fresh paraIds), so emit a positional p[J] segment that EmitComments
        // passes through verbatim. CONSISTENCY(word-table-positional).
        bool inCell = cursor is TableCell;
        sb.Append(inCell ? $"/p[{paraIdx}]" : $"/{BuildParaPathSegment(para, paraIdx)}");
        return sb.ToString();
    }

    /// <summary>
    /// Get-or-create a sectPr child in CT_SectPr schema order. Replaces the
    /// `?? sectPr.AppendChild(new T())` idiom which violated schema order when
    /// other higher-ranked elements were already present.
    /// </summary>
    private static T EnsureSectPrChild<T>(SectionProperties sectPr) where T : OpenXmlElement, new()
    {
        var existing = sectPr.GetFirstChild<T>();
        if (existing != null) return existing;
        var created = new T();
        InsertSectPrChildInOrder(sectPr, created);
        return created;
    }

    /// <summary>
    /// Get current document protection mode and enforcement status.
    /// </summary>
    private (string mode, bool enforced) GetDocumentProtection()
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var docProtection = settings?.GetFirstChild<DocumentProtection>();
        if (docProtection == null)
            return ("none", false);

        var mode = docProtection.Edit?.InnerText switch
        {
            "readOnly" => "readOnly",
            "comments" => "comments",
            "trackedChanges" => "trackedChanges",
            "forms" => "forms",
            _ => "none"
        };
        var enforced = docProtection.Enforcement?.Value == true
            || (docProtection.Enforcement?.Value == null && docProtection.Edit != null);
        return (mode, enforced);
    }

    /// <summary>
    /// Check if an SDT element is editable based on its lock attribute and the current document protection.
    /// </summary>
    private bool IsSdtEditable(SdtProperties? sdtProps)
    {
        // A content-locked control is read-only regardless of document
        // protection: <w:lock w:val="contentLocked"|"sdtContentLocked"> both
        // mark the control's content as not editable. (sdtLocked only blocks
        // deletion of the control, leaving content editable.)
        if (IsSdtContentLocked(sdtProps))
            return false;

        var (mode, enforced) = GetDocumentProtection();

        // No protection or not enforced → editable (content not locked above)
        if (!enforced || mode == "none")
            return true;

        // readOnly protection → SDTs are not editable (unless in permRange, P2)
        if (mode == "readOnly")
            return false;

        // forms protection → editable (content-lock already handled above)
        if (mode == "forms")
            return true;

        // comments/trackedChanges → not typically editable
        return false;
    }

    /// <summary>
    /// True when the SDT's content is locked read-only
    /// (<w:lock w:val="contentLocked"> or "sdtContentLocked").
    /// </summary>
    private static bool IsSdtContentLocked(SdtProperties? sdtProps)
    {
        var lockEl = sdtProps?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
        var lockVal = lockEl?.Val?.Value;
        return lockVal == LockingValues.ContentLocked
            || lockVal == LockingValues.SdtContentLocked;
    }
}
