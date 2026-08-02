// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddComment(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        var commentRun = parent as Run;
        var commentPara = commentRun?.Parent as Paragraph ?? parent as Paragraph
            ?? throw new ArgumentException("Comments must be added to a paragraph or run: /body/p[N] or /body/p[N]/r[M]");

        // BUG-DUMP-R26-3: `rangeEnd=true` closes an already-open comment range
        // (a CommentRangeStart with no matching CommentRangeEnd, created by an
        // earlier `add comment` carrying rangeOpen=true) at THIS paragraph,
        // placing the CommentRangeEnd + reference run here. This is the second
        // half of the two-marker round-trip for a multi-paragraph comment range
        // — mirrors the bookmark `open=true`/`end=true` span handling. No new
        // comment is created; we reuse the open range's id. Match the most-
        // recently-opened range (LIFO) so nested ranges close correctly.
        if (IsTruthy(properties.GetValueOrDefault("rangeEnd", "")))
        {
            var openStart = body.Descendants<CommentRangeStart>()
                .Where(rs => rs.Id?.Value != null
                    && !body.Descendants<CommentRangeEnd>().Any(re => re.Id?.Value == rs.Id!.Value))
                .LastOrDefault();
            if (openStart == null)
                throw new ArgumentException(
                    "comment rangeEnd has no matching open comment range " +
                    "(add the comment with rangeOpen=true first)");
            var openId = openStart.Id!.Value!;
            var endMarker = new CommentRangeEnd { Id = openId };
            var endRef = new Run(new CommentReference { Id = openId });
            // Place endMarker after the requested run (runEnd, 1-based; 0 =
            // paragraph start), then the reference run after it — matching the
            // source <w:commentRangeEnd/><w:r><w:commentReference/></w:r> shape.
            int runEndIdx = 0;
            if ((properties.TryGetValue("runEnd", out var reRaw)
                 || properties.TryGetValue("runend", out reRaw))
                && int.TryParse(reRaw, out var reN))
                runEndIdx = reN;
            OpenXmlElement? anchorRunE = null;
            if (runEndIdx >= 1)
            {
                var runs = commentPara.Elements<Run>().ToList();
                if (runEndIdx <= runs.Count)
                    anchorRunE = runs[runEndIdx - 1];
            }
            if (anchorRunE != null)
            {
                anchorRunE.InsertAfterSelf(endMarker);
                endMarker.InsertAfterSelf(endRef);
            }
            else
            {
                commentPara.AppendChild(endMarker);
                commentPara.AppendChild(endRef);
            }
            return $"/comments/comment[@id={openId}]/rangeEnd";
        }

        // BUG-R6B(BUG1): accept an empty/whitespace comment text. An empty
        // comment is valid OOXML; rejecting it broke the dump→batch round-trip
        // for comments whose inline text is empty (or whose only content is an
        // empty table, which flattens to empty text). When `text` is missing
        // entirely we still require it (a typed `add comment` with no text is a
        // user error), but a present-but-empty `text=""` is honoured and builds
        // a comment with an empty paragraph.
        if (!properties.TryGetValue("text", out var commentText))
            throw new ArgumentException("'text' property is required for comment type");

        var author = properties.GetValueOrDefault("author", "officecli");
        var initials = properties.GetValueOrDefault("initials", author[..1]);

        // Pre-validate user-supplied strings for invalid XML 1.0 chars
        // (U+0001..U+001F minus tab/LF/CR). Without this, a C0 control char
        // in author/initials/text would let us append the comment to the
        // comments part, then explode at Save() — producing an orphaned
        // comment with no anchor in the body (torn write).
        static void RejectIllegalXmlChars(string field, string value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c == '\t' || c == '\n' || c == '\r') continue;
                if (c < 0x20)
                    throw new ArgumentException(
                        $"'{field}' contains an illegal XML 1.0 control character (U+{(int)c:X04}); allowed C0 chars are tab/LF/CR only.");
            }
        }
        RejectIllegalXmlChars("text", commentText);
        RejectIllegalXmlChars("author", author);
        RejectIllegalXmlChars("initials", initials);
        var commentsPart = _doc.MainDocumentPart!.WordprocessingCommentsPart
            ?? _doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments ??= new Comments();

        var commentId = (commentsPart.Comments.Elements<Comment>()
            .Select(c => int.TryParse(c.Id?.Value, out var id) ? id : 0)
            .DefaultIfEmpty(0).Max() + 1).ToString();
        // BUG-DUMP-H103: honor an explicit source comment id when it is free.
        // The dump emitter passes the source id on a definition-only comment whose
        // <w:commentRangeStart/End/Reference> markers are carried verbatim by a
        // raw-passed paragraph (a cross-paragraph TOC field span) — reusing that id
        // lets the verbatim markers resolve to this definition. Falls back to the
        // fresh max+1 when the id is absent or already taken (mirrors the bookmark
        // R47-5 raw-set id-reuse). `id` is consumed here so it is not flagged
        // unsupported below.
        if ((properties.TryGetValue("id", out var provId)
                || properties.TryGetValue("commentId", out provId))
            && int.TryParse(provId, out var provIdN) && provIdN >= 0
            && !commentsPart.Comments.Elements<Comment>().Any(c => c.Id?.Value == provId))
            commentId = provId;
        properties.Remove("id");
        properties.Remove("commentId");

        // BUG-R6B(BUG1): empty text -> empty paragraph (no run); non-empty ->
        // a run carrying the text. Both are valid OOXML comment bodies.
        // BUG-DUMP-NOTE-TAB: build the seed run via AppendTextWithBreaks so a tab /
        // newline in the comment text becomes a structural <w:tab/> / <w:br/> rather
        // than a literal U+0009/U+000A glyph (mirrors `add r` and AddFootnote).
        Paragraph commentBody;
        if (string.IsNullOrEmpty(commentText))
            commentBody = new Paragraph();
        else
        {
            var cmtRun = new Run();
            AppendTextWithBreaks(cmtRun, commentText);
            // BUG-DUMP-NOTE-DEL: honor track-change attribution on the comment seed
            // run too (no-op when absent), mirroring the footnote/endnote seed.
            commentBody = new Paragraph(ApplyNoteSeedRevision(cmtRun, properties));
        }
        // BUG-DUMP-R40-2: a Word-authored comment body opens with the comment
        // reference mark run — <w:r><w:rPr><w:rStyle w:val="CommentReference"/>
        // </w:rPr><w:annotationRef/></w:r>. The dump emitter rides this run on
        // `add comment` (annotationRef=true + rStyle). Prepend it so the rebuilt
        // comment keeps its clickable reference glyph and the comment-pane
        // styling. The rStyle rides on THIS run only (consumed here so
        // ApplyCommentFormatKeys does not re-stamp it onto the text run, which
        // the source leaves un-styled).
        if (IsTruthy(properties.GetValueOrDefault("annotationRef", "")))
        {
            var annRefRPr = new RunProperties();
            if ((properties.TryGetValue("rStyle", out var annRStyle)
                 || properties.TryGetValue("rstyle", out annRStyle))
                && !string.IsNullOrEmpty(annRStyle))
                annRefRPr.RunStyle = new RunStyle { Val = annRStyle };
            var annRefRun = new Run(annRefRPr, new AnnotationReferenceMark());
            commentBody.PrependChild(annRefRun);
            // Don't let ApplyCommentFormatKeys re-apply rStyle to every content
            // run (it would wrongly style the trailing text run too).
            properties.Remove("rStyle");
            properties.Remove("rstyle");
            properties.Remove("annotationRef");
            properties.Remove("annotationref");
        }
        // BUG-DUMP-R26-4: preserve the source comment's first-paragraph
        // w14:paraId so commentsExtended.xml reply threading (keyed by paraId)
        // round-trips. EnsureAllParaIds only assigns when paraId is empty, so a
        // pre-stamped id survives. textId is left for the global pass to fill.
        if ((properties.TryGetValue("commentParaId", out var cpId)
             || properties.TryGetValue("commentparaid", out cpId))
            && !string.IsNullOrEmpty(cpId))
        {
            commentBody.ParagraphId = cpId;
        }
        var commentEl = new Comment(commentBody)
        {
            Id = commentId, Author = author, Initials = initials,
            // CONSISTENCY(date-roundtrip): RoundtripKind keeps DateTimeKind.Utc
            // (input ending in Z stays UTC and serializes back with Z) and
            // DateTimeKind.Local with explicit offset (input "...+08:00" keeps
            // the +08:00 form). Default Parse converts everything to Local,
            // poisoning round-trip on docs whose comment dates are UTC.
            Date = properties.TryGetValue("date", out var ds) ? DateTime.Parse(ds, null, System.Globalization.DateTimeStyles.RoundtripKind) : DateTime.UtcNow
        };
        commentsPart.Comments.AppendChild(commentEl);
        // Apply paragraph-level / run-level format keys (direction, font, size, etc.)
        // Mirrors R2-2 footnote/header fix — the same vocabulary should work
        // on comment bodies as on footnote/endnote bodies.
        // Reply threading (w15:paraIdParent) + resolved-state (w15:done) live in
        // word/commentsExtended.xml, keyed by the comment paragraphs' w14:paraId.
        // Consume parentId/done here (translating the parent's w:id -> its paraId)
        // and remove them so the unsupported-forwarding below doesn't flag them.
        if ((properties.TryGetValue("parentId", out var parentIdRaw)
             || properties.TryGetValue("parentid", out parentIdRaw))
            && !string.IsNullOrEmpty(parentIdRaw))
        {
            var parentParaId = GetCommentFirstParaId(parentIdRaw)
                ?? throw new ArgumentException(
                    $"parentId={parentIdRaw}: no comment with that id to reply to.");
            if (string.IsNullOrEmpty(commentBody.ParagraphId?.Value)) AssignParaId(commentBody);
            // Word writes a commentEx for every comment; ensure the parent's
            // thread-root entry exists, then link this reply to it.
            UpsertCommentEx(parentParaId, null, null);
            UpsertCommentEx(commentBody.ParagraphId!.Value!, parentParaId, false);
        }
        properties.Remove("parentId");
        properties.Remove("parentid");
        if ((properties.TryGetValue("done", out var addDoneRaw)
             || properties.TryGetValue("resolved", out addDoneRaw)))
        {
            if (string.IsNullOrEmpty(commentBody.ParagraphId?.Value)) AssignParaId(commentBody);
            UpsertCommentEx(commentBody.ParagraphId!.Value!, null, IsTruthy(addDoneRaw));
        }
        properties.Remove("done");
        properties.Remove("resolved");

        var _commentUnsupported = new List<string>();
        ApplyCommentFormatKeys(commentEl, properties, _commentUnsupported);
        commentsPart.Comments.Save();

        // Surface genuinely-unsupported props through the same channel every
        // other `add` type uses (LastAddUnsupportedProps -> CLI "UNSUPPORTED
        // props:" WARNING). AddComment used to discard _commentUnsupported, so
        // an unknown key (a typo, or a not-yet-supported feature like
        // `parentId` reply-threading / `done` resolution) was swallowed
        // silently — inconsistent with `add paragraph`, where ApplyCommentFormatKeys
        // sees the structural keys AddComment consumes itself (rangeOpen,
        // pointRef, runStart, …) and would otherwise flag them as false
        // positives; exclude that set, forward the rest.
        foreach (var key in _commentUnsupported)
        {
            switch (key.ToLowerInvariant())
            {
                case "text": case "author": case "initials": case "date":
                case "annotationref": case "rstyle":
                case "commentparaid": case "pointref":
                case "range": case "rangeopen": case "rangeend":
                case "runstart": case "runend":
                    continue;
                default:
                    LastAddUnsupportedProps.Add(key);
                    break;
            }
        }

        // BUG-DUMP-H103: definition-only emit (`range=none`). The comment's range
        // start/end + reference run already exist in the document body, carried
        // verbatim by a raw-passed paragraph (a cross-paragraph TOC field span)
        // with this same (source) comment id. Placing typed markers here would
        // duplicate the start under a fresh id and orphan the verbatim markers, so
        // we create the definition only and leave anchoring to the verbatim markers.
        if ((properties.TryGetValue("range", out var rangeNoneRaw)
                || properties.TryGetValue("Range", out rangeNoneRaw))
            && string.Equals(rangeNoneRaw, "none", StringComparison.OrdinalIgnoreCase))
            return $"/comments/comment[@id={commentId}]";

        var rangeStart = new CommentRangeStart { Id = commentId };
        var rangeEnd = new CommentRangeEnd { Id = commentId };
        var refRun = new Run(new CommentReference { Id = commentId });

        // BUG-DUMP-COMMENT-POINTREF: a zero-width / point-anchored comment in the
        // source carries ONLY <w:commentReference> (no commentRangeStart/End).
        // The dump emitter detects this and passes pointRef=true (alias
        // range=false) so we replay just the reference run — without it,
        // AddComment unconditionally synthesized a range and a point comment
        // silently became a ranged comment on round-trip.
        bool pointRef = false;
        if (properties.TryGetValue("pointRef", out var prRaw)
            || properties.TryGetValue("pointref", out prRaw))
            pointRef = IsTruthy(prRaw);
        else if ((properties.TryGetValue("range", out var rngRaw)
                  || properties.TryGetValue("Range", out rngRaw))
                 && IsExplicitFalseAddOverride(rngRaw))
            pointRef = true;

        // BUG-DUMP-R26-3: `rangeOpen=true` marks a multi-paragraph comment range
        // whose CommentRangeEnd + reference run live in a LATER paragraph. Place
        // only the CommentRangeStart here; a follow-up `add comment rangeEnd=true`
        // closes the range at the end paragraph. Without this the rangeEnd/refRun
        // were crammed into the start paragraph and the comment scoped only it.
        bool rangeOpen = !pointRef && IsTruthy(properties.GetValueOrDefault("rangeOpen", ""));

        if (pointRef)
        {
            // Reference-only: place a single <w:commentReference> run at the
            // requested anchor; no range markers. The reference run alone is a
            // valid OOXML point comment.
            if (commentRun != null)
            {
                commentRun.InsertAfterSelf(refRun);
            }
            else if (index.HasValue)
            {
                InsertIntoParagraph(commentPara, new OpenXmlElement[] { refRun }, index);
            }
            else
            {
                int runStartIdx = 0;
                if ((properties.TryGetValue("runstart", out var rsRaw)
                     || properties.TryGetValue("runStart", out rsRaw))
                    && int.TryParse(rsRaw, out var rsN))
                    runStartIdx = rsN;
                OpenXmlElement? anchorRun = null;
                if (runStartIdx >= 1)
                {
                    var runs = commentPara.Elements<Run>().ToList();
                    if (runStartIdx <= runs.Count)
                        anchorRun = runs[runStartIdx - 1];
                }
                if (anchorRun != null) anchorRun.InsertAfterSelf(refRun);
                else commentPara.AppendChild(refRun);
            }
        }
        else if (commentRun != null)
        {
            commentRun.InsertBeforeSelf(rangeStart);
            // BUG-DUMP-R26-3: rangeOpen defers the end/ref to a later paragraph.
            if (!rangeOpen)
            {
                commentRun.InsertAfterSelf(rangeEnd);
                rangeEnd.InsertAfterSelf(refRun);
            }
        }
        else
        {
            // index is a childElement-index (ResolveAnchorPosition counts pPr).
            // Use pPr-aware insert so an index pointing at ParagraphProperties
            // clamps forward (pPr must stay first child).
            if (index.HasValue)
            {
                InsertIntoParagraph(commentPara,
                    rangeOpen
                        ? new OpenXmlElement[] { rangeStart }
                        : new OpenXmlElement[] { rangeStart, rangeEnd, refRun },
                    index);
            }
            else
            {
                // CONSISTENCY(comment-runStart): when caller passes runStart=N (N>=1),
                // place rangeStart immediately AFTER the Nth run in the paragraph
                // so dump round-trip restores the anchor position. N=0 keeps the
                // legacy paragraph-start placement.
                int runStartIdx = 0;
                if ((properties.TryGetValue("runstart", out var rsRaw)
                     || properties.TryGetValue("runStart", out rsRaw))
                    && int.TryParse(rsRaw, out var rsN))
                    runStartIdx = rsN;
                OpenXmlElement? anchorRun = null;
                if (runStartIdx >= 1)
                {
                    var runs = commentPara.Elements<Run>().ToList();
                    if (runStartIdx <= runs.Count)
                        anchorRun = runs[runStartIdx - 1];
                }
                if (anchorRun != null)
                {
                    anchorRun.InsertAfterSelf(rangeStart);
                }
                else
                {
                    var after = commentPara.ParagraphProperties as OpenXmlElement;
                    if (after != null) after.InsertAfterSelf(rangeStart);
                    else commentPara.InsertAt(rangeStart, 0);
                }
                // BUG-DUMP-R26-3: rangeOpen defers the end/ref to a later
                // paragraph (closed by a follow-up `add comment rangeEnd=true`).
                if (!rangeOpen)
                {
                    commentPara.AppendChild(rangeEnd);
                    commentPara.AppendChild(refRun);
                }
            }
        }

        // Return navigable path using /comments/comment[N] (sequential index)
        var commentIndex = commentsPart.Comments.Elements<Comment>().ToList()
            .FindIndex(c => c.Id?.Value == commentId) + 1;
        var resultPath = $"/comments/comment[{commentIndex}]";
        return resultPath;
    }

    // P2 (bookmarkEnd typed support): place a STANDALONE <w:bookmarkEnd w:id=N>
    // at an arbitrary position. `add bookmark --prop end=true` only closes a
    // start that the current batch opened by NAME; a code generator replaying a
    // dump often has just the id-keyed end node (start added by a separate op,
    // or the end sits across a structural boundary the name-match can't reach),
    // and previously fell through to AddDefault (schema-invalid unnamespaced
    // attrs). This is the id-explicit counterpart to AddBookmark's end=true
    // branch. `id` is required (the bookmark id to close); `name` optionally
    // resolves the id from an existing start. EnsureBookmarkIds pairs it to the
    // matching start at flush and never deletes an unmatched end.
    private string AddBookmarkEnd(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Cell redirect mirrors AddBookmark: a bookmarkEnd is inline content, so
        // land it in the cell's first paragraph (cells only accept block-level
        // children), keeping the returned path round-trippable.
        if (parent is TableCell tc)
        {
            var firstPara = tc.Elements<Paragraph>().FirstOrDefault();
            if (firstPara == null)
            {
                firstPara = new Paragraph();
                AssignParaId(firstPara);
                tc.AppendChild(firstPara);
            }
            var paraIdx = PathIndex.FromArrayIndex(tc.Elements<Paragraph>().ToList().IndexOf(firstPara));
            parent = firstPara;
            parentPath = $"{parentPath}/{BuildParaPathSegment(firstPara, paraIdx)}";
            index = null;
        }

        var idVal = properties.GetValueOrDefault("id", "");
        var name = properties.GetValueOrDefault("name", "");
        if (string.IsNullOrEmpty(idVal) && !string.IsNullOrEmpty(name))
        {
            // Resolve the id from the last same-name start that has no end yet
            // (LIFO close for nested same-name bookmarks), falling back to the
            // last named start — mirrors AddBookmark's end=true selection.
            var body = _doc.MainDocumentPart?.Document?.Body;
            var namedStarts = body?.Descendants<BookmarkStart>()
                .Where(bs => string.Equals(bs.Name?.Value, name, StringComparison.Ordinal) && bs.Id?.Value != null)
                .ToList() ?? new List<BookmarkStart>();
            var openStart = namedStarts
                .Where(bs => !(body?.Descendants<BookmarkEnd>().Any(be => be.Id?.Value == bs.Id!.Value) ?? false))
                .LastOrDefault() ?? namedStarts.LastOrDefault();
            if (openStart == null)
                throw new ArgumentException(
                    $"bookmarkEnd by name '{name}' found no matching bookmarkStart. " +
                    "Pass --prop id=N to place the end by explicit id, or add the start first.");
            idVal = openStart.Id!.Value!;
        }
        if (string.IsNullOrEmpty(idVal))
            throw new ArgumentException(
                "bookmarkEnd requires --prop id=N (the bookmark id to close) or --prop name=NAME (to resolve the id from an existing start).");
        if (!int.TryParse(idVal, out _))
            throw new ArgumentException($"bookmarkEnd id must be an integer (got '{idVal}').");

        var endOnly = new BookmarkEnd { Id = idVal };
        if (parent is Paragraph endPara)
            InsertIntoParagraph(endPara, new OpenXmlElement[] { endOnly }, index);
        else
            InsertAtIndexOrAppend(parent, endOnly, index);
        return $"{parentPath}/bookmarkEnd[@id={idVal}]";
    }

    private string AddBookmark(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // BUG-FIX(B2): bookmarks under a table cell are inline content. The cell
        // schema only accepts block-level children (p/tbl/sdt), so redirect to
        // the cell's first paragraph (creating one if the cell is empty) and
        // append the bookmark path segment to the parent path so the returned
        // path is round-trippable via Get.
        if (parent is TableCell tc)
        {
            var firstPara = tc.Elements<Paragraph>().FirstOrDefault();
            if (firstPara == null)
            {
                firstPara = new Paragraph();
                AssignParaId(firstPara);
                tc.AppendChild(firstPara);
            }
            var paraIdx = PathIndex.FromArrayIndex(tc.Elements<Paragraph>().ToList().IndexOf(firstPara));
            parent = firstPara;
            parentPath = $"{parentPath}/{BuildParaPathSegment(firstPara, paraIdx)}";
            // Drop --index — it referred to a position inside the cell, not
            // inside the paragraph; preserving it would silently mis-anchor.
            index = null;
        }

        var bkName = properties.GetValueOrDefault("name", "");
        if (string.IsNullOrEmpty(bkName))
            throw new ArgumentException("'name' property is required for bookmark");
        // OOXML ST_Bookmark caps the name attribute at maxLength=40.
        if (bkName.Length > 40)
            throw new ArgumentException(
                $"bookmark name exceeds OOXML maxLength=40 (got {bkName.Length} chars). Truncate the name.");
        // XML 1.0 §2.2: reject illegal control chars, lone surrogates, and
        // U+FFFE/FFFF noncharacters in the bookmark name. The shared helper
        // raises an ArgumentException whose message contains the prop name.
        OfficeCli.Core.ParseHelpers.ValidateXmlText(bkName, "bookmark name");

        // BUG-DUMP-BMSPAN: `end=true` places ONLY a <w:bookmarkEnd> closing an
        // already-open <w:bookmarkStart> of the same name (the start was added
        // earlier with `open=true`). This is the second half of the two-marker
        // round-trip for a content-wrapping bookmark; replaying it AFTER the
        // wrapped runs keeps them inside the range. Match the most-recent open
        // start (no BookmarkEnd with its id yet) so nested same-name bookmarks
        // close LIFO, and insert the End at the requested position so it lands
        // after the wrapped content in document order.
        if (IsTruthy(properties.GetValueOrDefault("end", "")))
        {
            var namedStarts = body.Descendants<BookmarkStart>()
                .Where(bs => string.Equals(bs.Name?.Value, bkName, StringComparison.Ordinal)
                    && bs.Id?.Value != null)
                .ToList();
            // Prefer an un-closed start (no BookmarkEnd shares its id yet) so
            // nested same-name bookmarks close LIFO. BUG-DUMP-R47-7: fall back to
            // the last named start when that strict filter finds nothing. With
            // span-open id forwarding (BUG-DUMP-R47-5) a start can carry a SOURCE
            // id that transiently collides with another bookmark's end, so the
            // strict "no end with this id" probe wrongly judges the start closed
            // and the end=true op threw — dropping the end and leaving the
            // bookmark range unclosed (every TOC PAGEREF to it then rendered
            // "Error! Bookmark not defined"). Creating the end with a colliding id
            // is safe: EnsureBookmarkIds renumbers the matched start+end pair as a
            // unit at flush, so no duplicate survives.
            var openStart = namedStarts
                .Where(bs => !body.Descendants<BookmarkEnd>().Any(be => be.Id?.Value == bs.Id!.Value))
                .LastOrDefault()
                ?? namedStarts.LastOrDefault();
            if (openStart == null)
                throw new ArgumentException(
                    $"bookmark end for '{bkName}' has no matching open bookmarkStart " +
                    "(add the start with open=true first)");
            var endOnly = new BookmarkEnd { Id = openStart.Id!.Value };
            if (parent is Paragraph endPara2)
                InsertIntoParagraph(endPara2, new OpenXmlElement[] { endOnly }, index);
            else
                InsertAtIndexOrAppend(parent, endOnly, index);
            return $"{parentPath}/bookmarkEnd[@id={openStart.Id!.Value}]";
        }

        bool spanOpen = IsTruthy(properties.GetValueOrDefault("open", ""));

        // BUG-R3 (dump emits a name its own batch rejects): OOXML's w:name
        // permits characters the CLI path/selector grammar treats specially —
        // whitespace, quote/@, AND '/', '[', ']'. Real documents carry such
        // names: legal-doc and HTML-export generators auto-name TOC-anchor
        // bookmarks after heading text, e.g. "Review/Analysis" or
        // "Revenues/Receivables/Unearned_Revenues". A hard reject broke
        // dump→batch round-trip — the dumped `add bookmark name="…"` failed on
        // replay, dropping the bookmark and breaking every REF/TOC field (which
        // reference it by name, not by CLI selector) and PAGEREF page numbers.
        // Preserve the source name verbatim and warn instead, mirroring the
        // duplicate-bookmark-name allow+warn handling below. The only cost is
        // that such a bookmark can't be addressed by a CLI path selector later
        // (REF/TOC resolution in Word is unaffected).
        if (bkName.Any(c => c == '/' || c == '[' || c == ']'))
        {
            LastAddWarnings.Add(
                $"bookmark name '{bkName}' contains path-special characters " +
                "('/', '[', ']') — kept (OOXML allows it, and REF/TOC fields " +
                "reference it by name), but it cannot be addressed via a CLI " +
                "path selector.");
        }
        else if (bkName.Any(char.IsWhiteSpace) || bkName[0] == '@' || bkName[0] == '\'' || bkName.Contains('"'))
        {
            LastAddWarnings.Add(
                $"bookmark name '{bkName}' contains whitespace or quote/@ chars — " +
                "kept (OOXML allows it), but addressing by a bare attribute " +
                "selector (/bookmark[@name=...]) may be ambiguous; quote the value.");
        }

        // BUG-DUMPR2-02: Word permits multiple bookmarks to share a name
        // (legal OOXML; validates clean), so a hard reject broke dump→batch
        // round-trip of any such document — the replay re-adds each bookmark by
        // its source name and the second one threw, dropping a bookmark. The
        // bookmark Id stays unique (allocated below), which is what the format
        // requires; only the display name repeats. Warn instead of failing:
        // /bookmark[@name=X] then resolves to the first match, but the bookmark
        // is preserved. (Mirrors the duplicate form-field-name handling.)
        // Scan every part that can hold a bookmark (body + headers + footers +
        // footnotes + endnotes + comments), not just body — a body-only scan
        // allocated a colliding max+1 when a bookmark already lived in a
        // header/footer/note, and missed cross-part name duplicates. Mirrors
        // EnsureBookmarkIds' scan scope so allocator and dedup agree.
        var mainForBk = _doc.MainDocumentPart;
        var existingStarts = mainForBk != null
            ? EnumerateContentRoots(mainForBk).SelectMany(r => r.Descendants<BookmarkStart>()).ToList()
            : body.Descendants<BookmarkStart>().ToList();
        if (existingStarts.Any(b => string.Equals(b.Name?.Value, bkName, StringComparison.Ordinal)))
        {
            LastAddWarnings.Add(
                $"bookmark name '{bkName}' duplicates an existing bookmark — kept " +
                "(Word allows it), but addressing by this name resolves to the first match.");
        }

        var existingIds = existingStarts
            .Select(b => int.TryParse(b.Id?.Value, out var id) ? id : 0);
        // BUG-DUMP-R47-5: a content-WRAPPING (open=true) bookmark whose matching
        // <w:bookmarkEnd> is preserved verbatim by a raw-set (e.g. the end lives
        // inside a TOC <w:sdt> block raw-set as one unit, while the start is a
        // body-direct marker emitted via `add bookmark open=true`) must reuse
        // the SOURCE id so the add-side start and the raw-set-side end pair up.
        // Allocating a fresh max+1 here left the start unpaired (its end kept the
        // source id) — producing a duplicate/orphan w:id the validator rejects
        // and a broken bookmark range. EnsureBookmarkIds dedupes any residual
        // collision (renumbering the matched start+end pair as a unit), so honoring
        // the source id is safe. Only fires for the open=true path; zero-length
        // bookmarks still allocate fresh (start+end are created together here).
        string bkId;
        if (IsTruthy(properties.GetValueOrDefault("open", ""))
            && properties.TryGetValue("id", out var providedBkId)
            && !string.IsNullOrEmpty(providedBkId))
            bkId = providedBkId;
        else
            bkId = (existingIds.Any() ? existingIds.Max() + 1 : 1).ToString();

        var bookmarkStart = new BookmarkStart { Id = bkId, Name = bkName };
        var bookmarkEnd = new BookmarkEnd { Id = bkId };

        // BUG-DUMP-R32-4: a table-column-range bookmark carries
        // w:colFirst/w:colLast (a rectangular column-span over table columns).
        // BookmarkStartToNode now surfaces them; re-stamp here so dump→batch
        // keeps the bookmark as a column-range bookmark instead of downgrading
        // it to a plain point bookmark. Accept either casing (Get emits the
        // canonical colFirst/colLast).
        if ((properties.TryGetValue("colFirst", out var colFirstStr)
                || properties.TryGetValue("colfirst", out colFirstStr))
            && int.TryParse(colFirstStr, out var colFirstN))
            bookmarkStart.ColumnFirst = colFirstN;
        if ((properties.TryGetValue("colLast", out var colLastStr)
                || properties.TryGetValue("collast", out colLastStr))
            && int.TryParse(colLastStr, out var colLastN))
            bookmarkStart.ColumnLast = colLastN;

        // BUG-DUMP-BMDISPLACED: re-stamp w:displacedByCustomXml ("next"/"prev")
        // on a bookmark adjacent to a custom-XML / SDT boundary. Dropping it
        // (e.g. a TOC heading bookmark before the TOC <w:sdt>) shifted the
        // bookmark across the boundary so PAGEREF/TOC entries to it rendered
        // "Error! Bookmark not defined." BookmarkStartToNode surfaces it.
        if (properties.TryGetValue("displacedByCustomXml", out var dbcxStr)
            && !string.IsNullOrEmpty(dbcxStr))
        {
            if (dbcxStr.Equals("next", StringComparison.OrdinalIgnoreCase))
                bookmarkStart.DisplacedByCustomXml = DisplacedByCustomXmlValues.Next;
            else if (dbcxStr.Equals("prev", StringComparison.OrdinalIgnoreCase))
                bookmarkStart.DisplacedByCustomXml = DisplacedByCustomXmlValues.Previous;
        }

        // BUG-DUMP10-04: optional endPara offset (>0) defers BookmarkEnd
        // placement to a later paragraph in the same body so multi-
        // paragraph bookmark spans round-trip through dump→batch. Default
        // (0 / unset) keeps the End next to the Start as before.
        int crossParaEndOffset = 0;
        if ((properties.TryGetValue("endPara", out var bkEndStr)
                || properties.TryGetValue("endpara", out bkEndStr))
            && int.TryParse(bkEndStr, out var bkEndN) && bkEndN > 0)
        {
            crossParaEndOffset = bkEndN;
        }

        // index is a childElement-index (ResolveAnchorPosition counts pPr).
        // When anchor-based insert is requested, bypass the text-wrapping path
        // (which finds its own position inside existing runs) and do a positional
        // insert — the anchor wins. Route through the pPr-aware helper so an
        // index pointing at ParagraphProperties clamps forward.
        var bkPara = parent as Paragraph;
        var hasAnchor = index.HasValue && bkPara != null
            && index.Value >= 0 && index.Value < bkPara.ChildElements.Count;

        // When the body-wrap branch runs, the bookmark lives inside a newly
        // created <w:p>, not directly under Body. Track that so we can
        // return a path that descends into the wrapping paragraph — otherwise
        // `{parentPath}/bookmarkStart[...]` fails Get (CONSISTENCY(add-get-symmetry)).
        Paragraph? wrappingPara = null;

        if (spanOpen)
        {
            // BUG-DUMP-BMSPAN: content-wrapping bookmark — place ONLY the
            // <w:bookmarkStart> here; the matching <w:bookmarkEnd> arrives via
            // a later `end=true` op positioned after the wrapped runs. Without
            // this branch the fallback below pairs Start+End adjacently and the
            // range collapses to zero length the moment the runs replay after.
            if (parent is Paragraph openPara)
                InsertIntoParagraph(openPara, new OpenXmlElement[] { bookmarkStart }, index);
            else
                InsertAtIndexOrAppend(parent, bookmarkStart, index);
        }
        else if (properties.TryGetValue("text", out var bkText))
        {
            if (hasAnchor && bkPara != null)
            {
                var bkRun = new Run(new Text(bkText) { Space = SpaceProcessingModeValues.Preserve });
                InsertIntoParagraph(bkPara, new OpenXmlElement[] { bookmarkStart, bkRun, bookmarkEnd }, index);
            }
            else if (parent is Body)
            {
                // Runs must live inside a paragraph; wrap Start+Run+End in a new
                // <w:p> before inserting so we don't produce bare <w:r> as a
                // direct body child (schema-invalid).
                var bkRun = new Run(new Text(bkText) { Space = SpaceProcessingModeValues.Preserve });
                var wrapPara = new Paragraph(bookmarkStart, bkRun, bookmarkEnd);
                InsertAtIndexOrAppend(parent, wrapPara, index);
                wrappingPara = wrapPara;
            }
            else
            {
                // Try to find existing runs whose concatenated text contains the bookmark text
                var runs = parent.Elements<Run>().ToList();
                var wrapped = TryWrapExistingRunsWithBookmark(parent, runs, bkText, bookmarkStart, bookmarkEnd);
                if (!wrapped)
                {
                    // No matching text found — create a new run as fallback.
                    // Route through InsertAtIndexOrAppend so body-level inserts
                    // respect the trailing <w:sectPr> invariant (bookmarks
                    // landing after sectPr would be schema-invalid).
                    InsertAtIndexOrAppend(parent, bookmarkStart, index);
                    InsertAtIndexOrAppend(parent, new Run(new Text(bkText) { Space = SpaceProcessingModeValues.Preserve }),
                        index.HasValue ? index + 1 : null);
                    InsertAtIndexOrAppend(parent, bookmarkEnd,
                        index.HasValue ? index + 2 : null);
                }
            }
        }
        else if (hasAnchor && bkPara != null)
        {
            InsertIntoParagraph(bkPara, new OpenXmlElement[] { bookmarkStart, bookmarkEnd }, index);
        }
        else
        {
            // Body/other parents: honor --index/--after/--before and respect
            // Body's trailing <w:sectPr> invariant by routing through
            // InsertAtIndexOrAppend (which falls back to AppendToParent).
            InsertAtIndexOrAppend(parent, bookmarkStart, index);
            InsertAtIndexOrAppend(parent, bookmarkEnd, index.HasValue ? index + 1 : null);
        }

        // BUG-DUMP10-04: relocate the BookmarkEnd to a downstream sibling
        // paragraph when endPara was specified. Done after the initial
        // placement so all the existing schema-aware insertion paths
        // (text wrap, anchor index, body fallback) still run unmodified.
        if (crossParaEndOffset > 0 && bookmarkEnd.Parent != null)
        {
            // Walk up to the start's enclosing paragraph (it may be inside
            // a run if TryWrapExistingRunsWithBookmark wrapped runs).
            var startEnclosingPara = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault()
                ?? bookmarkStart.Parent as Paragraph;
            // Sibling list lives on the paragraph's parent (Body, TableCell, …).
            var siblingHost = startEnclosingPara?.Parent;
            if (startEnclosingPara != null && siblingHost != null)
            {
                var siblings = siblingHost.Elements<Paragraph>().ToList();
                int startIdx = siblings.IndexOf(startEnclosingPara);
                int targetIdx = startIdx + crossParaEndOffset;
                if (startIdx >= 0 && targetIdx < siblings.Count)
                {
                    bookmarkEnd.Remove();
                    siblings[targetIdx].AppendChild(bookmarkEnd);
                }
            }
        }

        // Return a navigable path: /...parent/bookmarkStart[@name=<name>] is
        // a real DOM element Navigation understands (the legacy
        // `/bookmark[<name>]` form addressed a synthetic type that Get/Add
        // could not resolve, breaking --after/--before reuse).
        // ValidateAndNormalizePredicate rejects bare attribute values that
        // contain whitespace, leading '@', or quote chars; double-quote the
        // value when the raw name would otherwise be rejected so the returned
        // path is round-trippable via `get`/`add --after`.
        // BUG-R3 (dump emits a name its own batch rejects): a name with an
        // embedded double-quote (legal in OOXML w:name, e.g. a
        // "Fast_math"_optimization name from some editors) cannot be expressed as an attribute
        // selector value in EITHER the bare or double-quoted form. The
        // bookmark itself is already in the document at this point; only the
        // navigable RETURN path can't carry the name. Fall back to a positional
        // bookmarkStart[N] segment instead of throwing, so dump→batch round-trip
        // preserves the bookmark. Warn that the @name selector is unavailable.
        string BookmarkSelector(OpenXmlElement container)
        {
            if (!bkName.Contains('"'))
                return $"bookmarkStart[@name={QuoteAttrValueIfNeeded(bkName)}]";
            LastAddWarnings.Add(
                $"bookmark name '{bkName}' contains an embedded double-quote — kept " +
                "(OOXML allows it), but it cannot be addressed by /bookmark[@name=...]; " +
                "use a positional bookmarkStart[N] selector instead.");
            var pos = PathIndex.FromArrayIndex(container.Descendants<BookmarkStart>().ToList().IndexOf(bookmarkStart));
            return $"bookmarkStart[{(pos > 0 ? pos : 1)}]";
        }

        string resultPath;
        if (wrappingPara != null)
        {
            var wrapIdx = PathIndex.FromArrayIndex(parent.Elements<Paragraph>().ToList().IndexOf(wrappingPara));
            resultPath = $"{parentPath}/{BuildParaPathSegment(wrappingPara, wrapIdx)}/{BookmarkSelector(wrappingPara)}";
        }
        else
        {
            resultPath = $"{parentPath}/{BookmarkSelector(parent)}";
        }
        return resultPath;
    }

    // BUG-DUMP-PERM: add a ranged editing-permission marker. permStart delimits
    // the start of a region a group (edGrp) or user (ed) may edit inside a
    // protected document; permEnd closes it. Both are positioned paragraph
    // children (like bookmark markers); --index anchors them at the right offset.
    private string AddPerm(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
        var isStart = type.Equals("permStart", StringComparison.OrdinalIgnoreCase);
        var idStr = properties.GetValueOrDefault("id", "");
        if (string.IsNullOrEmpty(idStr) || !int.TryParse(idStr, out var permId))
            throw new ArgumentException($"'id' property (integer) is required for {type}");

        OpenXmlElement marker;
        string segment;
        if (isStart)
        {
            var ps = new PermStart { Id = permId };
            if (properties.TryGetValue("edGrp", out var edGrp) && !string.IsNullOrEmpty(edGrp))
                ps.EditorGroup = new EnumValue<RangePermissionEditingGroupValues>(new RangePermissionEditingGroupValues(edGrp));
            if (properties.TryGetValue("ed", out var ed) && !string.IsNullOrEmpty(ed))
                ps.Ed = ed;
            if (properties.TryGetValue("colFirst", out var cf) && int.TryParse(cf, out var cfn))
                ps.ColumnFirst = cfn;
            if (properties.TryGetValue("colLast", out var cl) && int.TryParse(cl, out var cln))
                ps.ColumnLast = cln;
            marker = ps;
            segment = $"permStart[@id={permId}]";
        }
        else
        {
            marker = new PermEnd { Id = permId };
            segment = $"permEnd[@id={permId}]";
        }

        if (parent is Paragraph permPara)
            InsertIntoParagraph(permPara, new[] { marker }, index);
        else
            InsertAtIndexOrAppend(parent, marker, index);

        return $"{parentPath}/{segment}";
    }

    /// <summary>
    /// Quote an attribute predicate value when the bare form would be rejected
    /// by ValidateAndNormalizePredicate. Bare values must have no whitespace,
    /// no leading '@' or quote. Embedded double quotes cannot be represented
    /// by either form — error up front.
    /// </summary>
    private static string QuoteAttrValueIfNeeded(string value)
    {
        if (value.Contains('"'))
            throw new ArgumentException(
                $"Name '{value}' contains embedded double-quote, which cannot be represented in an attribute selector.");
        bool needsQuote = value.Length == 0
            || value[0] == '@' || value[0] == '\''
            || value.Any(char.IsWhiteSpace);
        return needsQuote ? $"\"{value}\"" : value;
    }

    /// <summary>
    /// Tries to wrap existing runs whose concatenated text contains <paramref name="targetText"/>
    /// with bookmarkStart/bookmarkEnd tags. Returns true if wrapping succeeded.
    /// </summary>
    private static bool TryWrapExistingRunsWithBookmark(
        OpenXmlElement parent, List<Run> runs, string targetText,
        BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd)
    {
        if (runs.Count == 0 || string.IsNullOrEmpty(targetText))
            return false;

        // Build a map: for each run, track the cumulative start offset and its text
        var runTexts = new List<(Run Run, int Start, string Text)>();
        var offset = 0;
        foreach (var run in runs)
        {
            var t = string.Concat(run.Elements<Text>().Select(x => x.Text));
            runTexts.Add((run, offset, t));
            offset += t.Length;
        }
        var fullText = string.Concat(runTexts.Select(r => r.Text));

        var matchIndex = fullText.IndexOf(targetText, StringComparison.Ordinal);
        if (matchIndex < 0)
            return false;

        var matchEnd = matchIndex + targetText.Length;

        // Find runs that overlap with [matchIndex, matchEnd)
        var firstRunIdx = -1;
        var lastRunIdx = -1;
        for (var i = 0; i < runTexts.Count; i++)
        {
            var runStart = runTexts[i].Start;
            var runEnd = runStart + runTexts[i].Text.Length;
            if (runEnd <= matchIndex) continue;
            if (runStart >= matchEnd) break;
            if (firstRunIdx < 0) firstRunIdx = i;
            lastRunIdx = i;
        }

        if (firstRunIdx < 0) return false;

        // Handle partial overlap at the start: split the first run if needed
        var firstRunInfo = runTexts[firstRunIdx];
        if (matchIndex > firstRunInfo.Start)
        {
            var splitPos = matchIndex - firstRunInfo.Start;
            var beforeText = firstRunInfo.Text[..splitPos];
            var afterText = firstRunInfo.Text[splitPos..];

            var beforeRun = (Run)firstRunInfo.Run.CloneNode(true);
            SetRunText(beforeRun, beforeText);
            parent.InsertBefore(beforeRun, firstRunInfo.Run);

            SetRunText(firstRunInfo.Run, afterText);
            // Update info
            runTexts[firstRunIdx] = (firstRunInfo.Run, matchIndex, afterText);
        }

        // Handle partial overlap at the end: split the last run if needed
        var lastRunInfo = runTexts[lastRunIdx];
        var lastRunEnd = lastRunInfo.Start + lastRunInfo.Text.Length;
        if (matchEnd < lastRunEnd)
        {
            var splitPos = matchEnd - lastRunInfo.Start;
            var keepText = lastRunInfo.Text[..splitPos];
            var tailText = lastRunInfo.Text[splitPos..];

            var tailRun = (Run)lastRunInfo.Run.CloneNode(true);
            SetRunText(tailRun, tailText);
            parent.InsertAfter(tailRun, lastRunInfo.Run);

            SetRunText(lastRunInfo.Run, keepText);
            runTexts[lastRunIdx] = (lastRunInfo.Run, lastRunInfo.Start, keepText);
        }

        // Insert bookmarkStart before the first matched run
        parent.InsertBefore(bookmarkStart, runTexts[firstRunIdx].Run);

        // Insert bookmarkEnd after the last matched run
        parent.InsertAfter(bookmarkEnd, runTexts[lastRunIdx].Run);

        return true;
    }

    private static void SetRunText(Run run, string text)
    {
        var existing = run.Elements<Text>().ToList();
        foreach (var t in existing) t.Remove();
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    /// <summary>
    /// Percent-encode characters disallowed in RFC 3986 URIs so the rel
    /// target conforms to OPC requirements. ASCII unreserved + reserved chars
    /// and already-encoded `%xx` sequences pass through; everything else
    /// (non-ASCII, ASCII space, control chars, and the gen-delim outliers
    /// `<>"{}|\^`+backtick) is percent-encoded as UTF-8 bytes. Used for
    /// hyperlink relationship targets on both Add and Set paths.
    /// </summary>
    private static string PercentEncodeUri(string uri)
    {
        if (string.IsNullOrEmpty(uri)) return uri;
        var sb = new System.Text.StringBuilder(uri.Length + 16);
        var bytes = new byte[4]; // max UTF-8 sequence
        for (int i = 0; i < uri.Length; i++)
        {
            char c = uri[i];
            if (c < 0x80 && IsUriSafe(c))
            {
                sb.Append(c);
                continue;
            }
            // Non-safe ASCII (space, controls, "<>`{}|\^"…) → single-byte %xx.
            if (c < 0x80)
            {
                sb.Append('%').Append(((byte)c).ToString("X2"));
                continue;
            }
            // Surrogate pair → 4-byte UTF-8 sequence.
            int codePoint;
            if (char.IsHighSurrogate(c) && i + 1 < uri.Length && char.IsLowSurrogate(uri[i + 1]))
            {
                codePoint = char.ConvertToUtf32(c, uri[i + 1]);
                i++;
            }
            else
            {
                codePoint = c;
            }
            int byteCount = System.Text.Encoding.UTF8.GetBytes(char.ConvertFromUtf32(codePoint), bytes);
            for (int j = 0; j < byteCount; j++)
                sb.Append('%').Append(bytes[j].ToString("X2"));
        }
        return sb.ToString();
    }

    // RFC 3986: unreserved + reserved chars + '%' (for already-encoded
    // sequences). Anything not on this list inside the ASCII range needs
    // percent-encoding. Tightening here also encodes 0x7F (DEL) and the
    // "exclude" set "<>`{}|\^"\""
    private static bool IsUriSafe(char c) =>
        (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') ||
        c == '-' || c == '.' || c == '_' || c == '~' ||      // unreserved
        c == ':' || c == '/' || c == '?' || c == '#' ||      // gen-delims
        c == '[' || c == ']' || c == '@' ||
        c == '!' || c == '$' || c == '&' || c == '\'' ||     // sub-delims
        c == '(' || c == ')' || c == '*' || c == '+' ||
        c == ',' || c == ';' || c == '=' ||
        c == '%';                                            // already-encoded passthrough

    private string AddHyperlink(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is `url`
        // (per schemas/help/docx/hyperlink.json). `href` and `link` are legacy
        // input aliases; Get normalizes readback to `url`.
        // Require a non-whitespace value: an empty/blank url is not a usable
        // target (and since relative URIs are accepted, an empty string would
        // otherwise slip through as a valid relative URI). Treating blank as
        // "no url" means url="" with no anchor hits the "url or anchor required"
        // error below, while url="" alongside an anchor proceeds as anchor-only.
        var hasUrl = (properties.TryGetValue("url", out var hlUrl)
            || properties.TryGetValue("href", out hlUrl)
            || properties.TryGetValue("link", out hlUrl))
            && !string.IsNullOrWhiteSpace(hlUrl);
        var hasAnchor = properties.TryGetValue("anchor", out var hlAnchor) || properties.TryGetValue("bookmark", out hlAnchor);
        // BUG-DUMP10-05: a w:hyperlink element with neither r:id nor anchor
        // is still a valid Word construct (tooltip-only / target-frame-only
        // hover popups). Only reject when none of the four destination /
        // metadata attributes are present so the wrapper can survive
        // dump→batch round-trip.
        var hasTooltip = properties.ContainsKey("tooltip");
        var hasTgtFrame = properties.ContainsKey("tgtFrame") || properties.ContainsKey("tgtframe");
        var hasHistory = properties.ContainsKey("history");
        if (!hasUrl && !hasAnchor && !hasTooltip && !hasTgtFrame && !hasHistory)
            throw new ArgumentException("'url' or 'anchor' property is required for hyperlink type");

        if (parent is not Paragraph hlPara)
            throw new ArgumentException("Hyperlinks can only be added to paragraphs: /body/p[N]");

        string? hlRelId = null;
        if (hasUrl)
        {
            // BUG-FIX(B1): hyperlinks inside header/footer/footnote/endnote
            // must add the rel to the enclosing host part (e.g. header1.xml.rels),
            // not document.xml.rels. Otherwise Word can't resolve the rId.
            var hostPart = ResolveHostPart(hlPara);
            // BUG-DUMP27: accept fragment-only URIs (e.g. "#_ftn1") in addition
            // to absolute URIs, to support dump→batch round-trip of internal-anchor
            // hyperlinks stored as r:id relationships with Target="#anchor".
            // Word's .rels accepts these per RFC 3986; mark them isExternal=false
            // so the .rels TargetMode is omitted (consistent with native Word output).
            var hlIsFragment = !string.IsNullOrEmpty(hlUrl) && hlUrl.StartsWith('#');
            Uri? hlUri;
            if (hlIsFragment)
            {
                // Internal bookmark targets must travel as w:anchor on the
                // <w:hyperlink> element, not as a Target="#name" relationship
                // — real Word rejects the latter as a corrupt file. Promote
                // `url=#bookmark` to the anchor= path and skip relationship
                // creation entirely.
                if (!hasAnchor)
                {
                    hlAnchor = hlUrl!.Substring(1);
                    hasAnchor = true;
                }
                hlUri = null;
            }
            else if (Uri.TryCreate(hlUrl, UriKind.Absolute, out hlUri))
            {
                // CONSISTENCY(hyperlink-scheme-allowlist): gate absolute URIs only.
                Core.HyperlinkUriValidator.RequireSafeScheme(hlUrl!, "url");
                // BUG-DUMP-FILEURI-BACKSLASH: a Windows local-path target
                // (file:///C:\Users\…\file.docx) is a valid Word hyperlink — Word
                // writes the rel Target verbatim with backslashes. System.Uri's
                // file-scheme parser, however, rejects a backslash DOS path
                // ("A Dos path must be rooted, for example, 'c:\'") when re-parsed
                // through `new Uri(...)`, which THREW and — because the throw
                // aborted the `add hyperlink` op — dropped the hyperlink's anchor
                // text from its paragraph (silent content loss). Normalize
                // backslashes to forward slashes for file: URIs before
                // percent-encoding (Word/OPC accept file:///C:/Users/…), so the
                // target round-trips and the op succeeds.
                var encoded = PercentEncodeUri(hlUrl!);
                if (hlUri.IsFile && encoded.Contains('\\'))
                    encoded = encoded.Replace('\\', '/');
                // Defensive: never let a single malformed absolute URI throw and
                // drop the run. If it still won't parse, fall back to the
                // already-parsed hlUri from TryCreate above (a valid Uri), so the
                // hyperlink survives rather than aborting the op.
                if (!Uri.TryCreate(encoded, UriKind.Absolute, out var reparsed))
                    reparsed = hlUri;
                hlUri = reparsed;
            }
            else if (Uri.TryCreate(hlUrl, UriKind.Relative, out hlUri))
            {
                // BUG-DUMPR2: a relative external target (e.g. "court-exif.jpg",
                // "../report.docx") is a valid Word hyperlink relationship with
                // TargetMode="External" and a relative Target — Word writes these
                // for links to sibling files. The dump emits exactly this value,
                // so rejecting it broke the round-trip. No scheme check applies
                // (a relative path carries none, so it cannot smuggle a
                // javascript:/data: scheme — those are absolute).
            }
            else
            {
                throw new ArgumentException($"Invalid hyperlink URL '{hlUrl}'. Expected an absolute URI (e.g. 'https://example.com'), a relative target (e.g. 'file.docx'), or a fragment-only anchor (e.g. '#bookmark').");
            }
            // Absolute and relative both round-trip as External relationships;
            // fragments are handled inline above and skip relationship creation.
            if (hlUri != null)
                hlRelId = hostPart.AddHyperlinkRelationship(hlUri, isExternal: true).Id;
        }

        var hlRProps = new RunProperties();
        // "inherit" sentinel (dump-emitted): the SOURCE hyperlink run carries
        // no <w:color>/<w:u> at all — its appearance comes from styles, or it
        // deliberately looks like plain text (TOC leader rows). Skip the
        // interactive defaults entirely so the rebuilt run stays element-free.
        bool hlColorInherit = properties.TryGetValue("color", out var hlColorProbe)
            && string.Equals(hlColorProbe, "inherit", StringComparison.OrdinalIgnoreCase);
        bool hlUnderlineInherit = (properties.TryGetValue("underline", out var hlUlProbe)
                || properties.TryGetValue("font.underline", out hlUlProbe))
            && string.Equals(hlUlProbe, "inherit", StringComparison.OrdinalIgnoreCase);
        if (hlColorInherit)
        {
            // no color element
        }
        else if (properties.TryGetValue("color", out var hlColor))
        {
            // BUG-R4B(BUG4): accept theme/scheme color names (text1, accent1, …)
            // on hyperlink Add the same way the run color path does — the old
            // bare SanitizeHex turned "text1" into a literal (invalid) hex. Route
            // through ApplyRunFormatting's color case so scheme colors write the
            // ThemeColor attribute and the dump→replay round-trip survives.
            hlRProps.RemoveAllChildren<Color>();
            ApplyRunFormatting(hlRProps, "color", hlColor);
        }
        else
        {
            // Read hyperlink color from document theme, fallback to Word default
            var themeHlink = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements
                ?.ColorScheme?.Hyperlink?.RgbColorModelHex?.Val?.Value;
            hlRProps.Color = new Color { Val = themeHlink ?? "0563C1", ThemeColor = ThemeColorValues.Hyperlink };
        }
        // CONSISTENCY(run-underline-enum): accept an explicit underline style on
        // the hyperlink run via the shared NormalizeUnderlineValue helper (same
        // enum/aliases as plain run `underline`). Default stays `single` for the
        // common case so existing behavior is preserved when no underline is
        // given. Without this, a source hyperlink whose run is dotted/wave/etc.
        // round-trips to single (the dump captures it, AddHyperlink dropped it).
        if (hlUnderlineInherit)
        {
            // no underline element
        }
        else if (properties.TryGetValue("underline", out var hlUnderline)
            || properties.TryGetValue("font.underline", out hlUnderline))
        {
            var hlUlVal = NormalizeUnderlineValue(hlUnderline);
            if (hlUlVal == "none")
                hlRProps.Underline = new Underline { Val = UnderlineValues.None };
            else
                hlRProps.Underline = new Underline { Val = new UnderlineValues(hlUlVal) };
        }
        else
        {
            hlRProps.Underline = new Underline { Val = UnderlineValues.Single };
        }
        // Explicit underline color (<w:u w:color="…">): dump emits it as
        // underline.color next to underline; route through the shared run
        // case so the attribute lands on the Underline element written above.
        if (properties.TryGetValue("underline.color", out var hlUlColor))
            ApplyRunFormatting(hlRProps, "underline.color", hlUlColor);
        if (properties.TryGetValue("font", out var hlFont))
            hlRProps.RunFonts = new RunFonts { Ascii = hlFont, HighAnsi = hlFont };
        // Ascii-only slot (<w:rFonts w:ascii="…"/> with no hAnsi): dump emits
        // font.ascii; the bare `font` case above would wrongly stamp hAnsi too.
        if (properties.TryGetValue("font.ascii", out var hlFontAscii))
        {
            hlRProps.RunFonts ??= new RunFonts();
            hlRProps.RunFonts.Ascii = hlFontAscii;
        }
        // Dump emits font.latin alongside bare font for hyperlink runs; mirror
        // the bare-font behavior so batch replay doesn't silently drop it.
        if (properties.TryGetValue("font.latin", out var hlFontLatin))
        {
            hlRProps.RunFonts ??= new RunFonts();
            hlRProps.RunFonts.Ascii = hlFontLatin;
            hlRProps.RunFonts.HighAnsi = hlFontLatin;
        }
        // BUG-DUMP17-07: mirror per-script font slot from Add.Text. Without this
        // branch, dump emits font.cs on hyperlink runs but batch replay silently
        // drops it.
        if (properties.TryGetValue("font.cs", out var hlFontCs)
            || properties.TryGetValue("font.complexscript", out hlFontCs)
            || properties.TryGetValue("font.complex", out hlFontCs))
        {
            hlRProps.RunFonts ??= new RunFonts();
            hlRProps.RunFonts.ComplexScript = hlFontCs;
        }
        // East-Asian font slot + character spacing: the dump emits these on
        // TOC-row links (Calibri eastAsia + spacing -1); without the cases the
        // wrapper run silently lost them while its sibling runs kept theirs.
        if (properties.TryGetValue("font.ea", out var hlFontEa)
            || properties.TryGetValue("font.eastasia", out hlFontEa))
        {
            hlRProps.RunFonts ??= new RunFonts();
            hlRProps.RunFonts.EastAsia = hlFontEa;
        }
        if (properties.TryGetValue("charSpacing", out var hlCharSp)
            || properties.TryGetValue("charspacing", out hlCharSp))
        {
            ApplyRunFormatting(hlRProps, "charSpacing", hlCharSp);
        }
        // Theme font slots (<w:rFonts w:asciiTheme="…" …/>), run shading and
        // character scale (<w:w/>): dump emits all of these on hyperlink runs
        // (minorEastAsia-themed Korean links, shaded URLs); route through the
        // shared run cases so batch replay keeps the source typography instead
        // of falling back to the document default (serif) font.
        // BUG-DUMP-HLINK-RPR: a run inside a <w:hyperlink> carries the full run
        // rPr vocabulary, but AddHyperlink only hand-rolled color/underline/font/
        // size/bold/italic. Any other character property the source set on the
        // link run — most consequentially <w:vanish/> (a hidden boilerplate /
        // template-guidance link), plus the per-script bold.cs/italic.cs/size.cs,
        // language, caps, strike, vertAlign, … — was silently dropped, so a
        // vanished hyperlink rendered as visible text. Route every remaining
        // character key the dump emits through the shared ApplyRunFormatting (the
        // same applier the plain-run path uses); none of these collide with the
        // color/underline/size/bold/italic/font slots handled specially above.
        foreach (var hlPassKey in new[]
                 {
                     "font.asciiTheme", "font.hAnsiTheme", "font.eaTheme",
                     "font.csTheme", "shading", "w",
                     "vanish", "specVanish", "webHidden",
                     "bold.cs", "italic.cs", "size.cs",
                     // The latin/default language slot is dumped as lang.latin
                     // (canonical), not bare `lang`; without it a hyperlink run's
                     // <w:lang w:val="…"/> was reported UNSUPPORTED and dropped on
                     // replay even though `add run` accepts it. lang.val is the
                     // other latin-slot alias ApplyRunFormatting recognizes.
                     "lang", "lang.latin", "lang.val", "lang.ea", "lang.cs",
                     "caps", "smallCaps", "strike", "dstrike",
                     "outline", "shadow", "emboss", "imprint",
                     // BUG-DUMP-HLINK-SUPERSCRIPT: the dump emits vertAlign in its
                     // canonical bool form `superscript`/`subscript` (not the raw
                     // `vertAlign` key), so a superscript/subscript hyperlink run —
                     // e.g. a footnote URL set superscript to render small — had
                     // its <w:vertAlign> dropped, the link re-rendered at full size
                     // and a long URL wrapped to an extra line, inflating the
                     // footnote and shifting every later page break. Route both
                     // through ApplyRunFormatting (handles superscript/subscript).
                     "superscript", "subscript",
                     "highlight", "vertAlign", "position", "kern",
                     // BUG-DUMP-HLINK-SNAPGRID: snapToGrid on a hyperlink run —
                     // on a docGrid doc, snapToGrid="0" keeps the link line off
                     // the grid (sets its height). Missing from the pass-through
                     // list, so it was dropped and the line re-snapped → reflow.
                     // ApplyRunFormatting gained a snapToGrid case (BUG-15).
                     "snapToGrid",
                     // BUG-DUMP-RPR-CONTAINER: a hyperlink run can be RTL
                     // (Arabic/Hebrew link text) and/or carry a CJK emphasis mark
                     // (<w:em>); both were absent from this list and dropped on
                     // round-trip. ApplyRunFormatting handles direction/rtl, and now
                     // em (added alongside this fix).
                     "direction", "rtl", "em", "emphasisMark",
                 })
        {
            if (properties.TryGetValue(hlPassKey, out var hlPassVal))
                ApplyRunFormatting(hlRProps, hlPassKey, hlPassVal);
        }
        if (properties.TryGetValue("size", out var hlSize))
            hlRProps.FontSize = new FontSize { Val = ((int)Math.Round(ParseFontSize(hlSize) * 2, MidpointRounding.AwayFromZero)).ToString() };
        if (properties.TryGetValue("bold", out var hlBold) && IsTruthy(hlBold))
            hlRProps.Bold = new Bold();
        if (properties.TryGetValue("italic", out var hlItalic) && IsTruthy(hlItalic))
            hlRProps.Italic = new Italic();
        // CONSISTENCY(add-set-symmetry): hyperlink runs commonly bind to the
        // built-in `Hyperlink` character style (rStyle=Hyperlink) so they
        // pick up the document's hyperlink theme color/underline. Run Add
        // and paragraph dump emit echo rStyle back; AddHyperlink must
        // accept it on the wrapped run or batch replay strips it with an
        // UNSUPPORTED warning. BUG-R4-BT5.
        if (properties.TryGetValue("rStyle", out var hlRStyle) || properties.TryGetValue("rstyle", out hlRStyle))
        {
            if (!string.IsNullOrEmpty(hlRStyle))
                hlRProps.RunStyle = new RunStyle { Val = hlRStyle };
        }
        // CONSISTENCY(rtl-cascade): inherit pPr/bidi from the enclosing
        // paragraph onto the hyperlink's run rPr. Mirrors the cascade in
        // SetElementParagraph / Add.Text run insertion (R16-bt-3). Without
        // this, a hyperlink inserted into an RTL paragraph renders LTR
        // because the run's RightToLeftText is missing — and effective.rtl
        // never resolves on the run NodeBuilder side either.
        if (hlPara.ParagraphProperties?.BiDi != null)
            ApplyRunFormatting(hlRProps, "rtl", "true");

        var hlRun = new Run(hlRProps);
        var hlText = properties.GetValueOrDefault("text", hlUrl ?? hlAnchor ?? "link");
        // BUG-DUMP-H102: build the display text via AppendTextWithBreaks (same as
        // AddRun) so a `\t` / `\n` in the hyperlink's text is rebuilt as a real
        // <w:tab/> / <w:br/>, not a literal control char in <w:t>. A hyperlink
        // whose leading content is a <w:tab/> (a TOC entry whose tab + page number
        // sit inside the link) dumps as `add hyperlink text="\t…"`; the old literal
        // `new Text` persisted the tab as a U+0009 character, which does NOT invoke
        // the paragraph's right-tab-stop + dot leader — degrading the TOC layout.
        // Empty text still yields an empty <w:t> (preserves the empty-wrapper case).
        AppendTextWithBreaks(hlRun, hlText);

        var hyperlink = new Hyperlink(hlRun);
        if (hlRelId != null)
            hyperlink.Id = hlRelId;
        if (hasAnchor)
            hyperlink.Anchor = hlAnchor;
        // BUG-DUMP24-02: w:docLocation is a separate "location in target
        // document" attribute, distinct from w:anchor. Round-trip it so
        // dump→batch preserves the wrapping hyperlink fully.
        if (properties.TryGetValue("docLocation", out var hlDocLoc)
            || properties.TryGetValue("doclocation", out hlDocLoc))
            hyperlink.DocLocation = hlDocLoc;
        // BUG-DUMP10-02: round-trip the optional metadata attrs.
        if (hasTooltip && properties.TryGetValue("tooltip", out var hlTooltip))
            hyperlink.Tooltip = hlTooltip;
        if (hasTgtFrame &&
            (properties.TryGetValue("tgtFrame", out var hlTgt)
             || properties.TryGetValue("tgtframe", out hlTgt)))
            hyperlink.TargetFrame = hlTgt;
        // BUG-DUMP-HISTFALSE: write the explicit boolean (true OR false).
        // OOXML default for w:history is true, so dropping an explicit
        // history=false on add silently flips the link to history-on.
        if (hasHistory && properties.TryGetValue("history", out var hlHist))
            hyperlink.History = OnOffValue.FromBoolean(IsTruthy(hlHist));

        // index is a childElement-index (ResolveAnchorPosition counts pPr).
        // Route through pPr-aware helper so index 0 clamps forward past
        // ParagraphProperties (pPr must stay first child of <w:p>).
        InsertIntoParagraph(hlPara, hyperlink, index);

        var hls = hlPara.Elements<Hyperlink>().ToList();
        var idx = hls.FindIndex(h => ReferenceEquals(h, hyperlink));
        var resultPath = $"{parentPath}/hyperlink[{(idx >= 0 ? idx + 1 : hls.Count)}]";
        return resultPath;
    }

    private string AddField(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string>? properties, string type)
    {
        properties ??= new Dictionary<string, string>();
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // Insert a field code (PAGE, NUMPAGES, DATE, etc.) as a run
        // Determines field instruction from type or "field" property
        // When type is "field", check fieldType/type property for dispatch
        var effectiveType = type.ToLowerInvariant();
        if (effectiveType == "field")
        {
            var ft = properties.GetValueOrDefault("fieldType")
                  ?? properties.GetValueOrDefault("fieldtype")
                  ?? properties.GetValueOrDefault("type");
            if (ft != null) effectiveType = ft.ToLowerInvariant();
        }
        // Extract named parameters for field types that require them
        string? mergeFieldName = null;
        string? refBookmarkName = null;
        string? seqIdentifier = null;

        if (effectiveType == "mergefield")
        {
            mergeFieldName = properties.GetValueOrDefault("fieldName")
                          ?? properties.GetValueOrDefault("fieldname")
                          ?? properties.GetValueOrDefault("name");
            if (string.IsNullOrWhiteSpace(mergeFieldName))
                throw new ArgumentException("MERGEFIELD requires a 'fieldName' property (e.g. --prop fieldName=CustomerName).");
        }
        else if (effectiveType is "ref" or "pageref" or "noteref")
        {
            refBookmarkName = properties.GetValueOrDefault("bookmarkName")
                           ?? properties.GetValueOrDefault("bookmarkname")
                           ?? properties.GetValueOrDefault("bookmark")
                           ?? properties.GetValueOrDefault("name");
            if (string.IsNullOrWhiteSpace(refBookmarkName))
                throw new ArgumentException($"{effectiveType.ToUpperInvariant()} requires a 'bookmarkName' property (e.g. --prop bookmarkName=MyBookmark).");
        }
        else if (effectiveType == "seq")
        {
            seqIdentifier = properties.GetValueOrDefault("identifier")
                         ?? properties.GetValueOrDefault("name")
                         ?? properties.GetValueOrDefault("id");
            if (string.IsNullOrWhiteSpace(seqIdentifier))
                throw new ArgumentException("SEQ requires an 'identifier' property (e.g. --prop identifier=Figure).");
        }

        // For STYLEREF and DOCPROPERTY, extract the required name parameter
        string? styleRefName = null;
        if (effectiveType == "styleref")
        {
            styleRefName = properties.GetValueOrDefault("styleName")
                        ?? properties.GetValueOrDefault("stylename")
                        ?? properties.GetValueOrDefault("name");
            if (string.IsNullOrWhiteSpace(styleRefName))
                throw new ArgumentException("STYLEREF requires a 'styleName' property (e.g. --prop styleName=\"Heading 1\").");
        }
        string? docPropertyName = null;
        if (effectiveType == "docproperty")
        {
            docPropertyName = properties.GetValueOrDefault("propertyName")
                           ?? properties.GetValueOrDefault("propertyname")
                           ?? properties.GetValueOrDefault("name");
            if (string.IsNullOrWhiteSpace(docPropertyName))
                throw new ArgumentException("DOCPROPERTY requires a 'propertyName' property (e.g. --prop propertyName=Department).");
        }

        // DATE/TIME `\@` format switch is opt-in: only emit when the user
        // supplied --prop format=… so a vanilla `add field --prop fieldType=date`
        // produces a bare `DATE` field that Word renders with the user's
        // locale default rather than a hardcoded ISO format.
        var dateFmtSwitch = properties.TryGetValue("format", out var dateFmtVal)
            && !string.IsNullOrWhiteSpace(dateFmtVal)
            ? $"\\@ \"{dateFmtVal}\" " : "";
        // BUG-R7A: dump→batch round-trips of typed fields lost every general
        // field switch (`\* roman`, `\* MERGEFORMAT`, `\p`, `\* Upper`, …)
        // because these bare-instruction arms appended nothing. The emitter
        // (WordBatchEmitter.BuildFieldAddProps) now carries the residual
        // switches via the `switches` prop; splice them back here exactly as
        // SEQ/MERGEFIELD already do via AppendFieldSwitches.
        var sw = AppendFieldSwitches(properties);
        var fieldInstr = effectiveType switch
        {
            "pagenum" or "pagenumber" or "page" => $" PAGE{sw} ",
            "numpages" => $" NUMPAGES{sw} ",
            "sectionpages" => $" SECTIONPAGES{sw} ",
            "section" => $" SECTION{sw} ",
            "date" => $" DATE {dateFmtSwitch}{sw.TrimStart()}".TrimEnd() + " ",
            "createdate" => $" CREATEDATE {dateFmtSwitch}{sw.TrimStart()}".TrimEnd() + " ",
            "savedate" => $" SAVEDATE {dateFmtSwitch}{sw.TrimStart()}".TrimEnd() + " ",
            "printdate" => $" PRINTDATE {dateFmtSwitch}{sw.TrimStart()}".TrimEnd() + " ",
            "edittime" => $" EDITTIME{sw} ",
            "author" => $" AUTHOR{sw} ",
            "lastsavedby" => $" LASTSAVEDBY{sw} ",
            "title" => $" TITLE{sw} ",
            "subject" => $" SUBJECT{sw} ",
            "filename" => $" FILENAME{sw} ",
            "time" => $" TIME {dateFmtSwitch}{sw.TrimStart()}".TrimEnd() + " ",
            "numwords" => $" NUMWORDS{sw} ",
            "numchars" => $" NUMCHARS{sw} ",
            "revnum" => $" REVNUM{sw} ",
            "template" => $" TEMPLATE{sw} ",
            "comments" or "doccomments" => $" COMMENTS{sw} ",
            "keywords" => $" KEYWORDS{sw} ",
            // BUG-DUMP9-09: quote MERGEFIELD names containing whitespace so
            // Word parses the full name as one token. " MERGEFIELD First Name "
            // would otherwise be parsed as field "First" with arg "Name".
            "mergefield" => $" MERGEFIELD {QuoteFieldNameIfNeeded(mergeFieldName!)}{sw} ",
            // BUG-R7A: REF/PAGEREF/NOTEREF now carry `\h`/`\p`/`\*` via the
            // `switches` prop. The legacy `hyperlink` prop still emits `\h`
            // for hand-authored adds, but the emitter routes `\h` through
            // `switches`, so guard against double-emitting it.
            "ref" => $" REF {refBookmarkName}{RefHyperlinkSwitch(properties, sw)}{sw} ",
            "pageref" => $" PAGEREF {refBookmarkName}{RefHyperlinkSwitch(properties, sw)}{sw} ",
            "noteref" => $" NOTEREF {refBookmarkName}{RefHyperlinkSwitch(properties, sw)}{sw} ",
            "seq" => $" SEQ {seqIdentifier}{sw} ",
            "styleref" => $" STYLEREF \"{styleRefName}\" ",
            "docproperty" => $" DOCPROPERTY \"{docPropertyName}\" ",
            "if" => BuildIfFieldInstruction(properties),
            // CONSISTENCY(field-add-symmetry): WordBatchEmitter.BuildFieldAddProps
            // emits legacy form fields with fieldType=FORMTEXT / FORMCHECKBOX
            // / FORMDROPDOWN. Without these arms the default arm threw
            // `Unknown field type 'formtext'`, breaking dump→batch round-trips
            // of any document containing a legacy form field. Delegate to
            // AddFormField (the canonical /formfield handler) which builds
            // the full FieldChar/FormFieldData/Bookmark chain.
            "formtext" => "__FORMFIELD_DELEGATE__",
            "formcheckbox" => "__FORMFIELD_DELEGATE__",
            "formdropdown" => "__FORMFIELD_DELEGATE__",
            // CONSISTENCY(field-add-symmetry): WordBatchEmitter.BuildFieldAddProps
            // emits HYPERLINK fields as fieldType=HYPERLINK + url/anchor (+ text),
            // never as a raw `instr`. Without a hyperlink case the default arm
            // throws `Unknown field type 'hyperlink'` and (under the new
            // continue-on-error default) the link is silently dropped on
            // dump→batch round-trips of complex-field HYPERLINK chains.
            "hyperlink" => BuildHyperlinkFieldInstruction(properties),
            // CONSISTENCY(canonical-keys): field.json declares `instr` as
            // the canonical raw-instruction key with `instruction` and
            // `code` as aliases. Help docs and AI prompts use `instr=`
            // (matching the readback key Get surfaces); accept all three.
            _ => GetRawFieldInstruction(properties)
                ?? throw new ArgumentException($"Unknown field type '{effectiveType}'. Provide a known type or an 'instr' / 'instruction' / 'code' property.")
        };
        // Form-field delegation: dump emits legacy form fields with
        // fieldType=FORMTEXT/FORMCHECKBOX/FORMDROPDOWN. Route to AddFormField
        // (the canonical /formfield handler) which builds the FieldChar +
        // FormFieldData + Bookmark chain. Map fieldType → formfieldtype.
        if (fieldInstr == "__FORMFIELD_DELEGATE__")
        {
            var ffProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            ffProps["formfieldtype"] = effectiveType switch
            {
                "formcheckbox" => "checkbox",
                "formdropdown" => "dropdown",
                _ => "text",
            };
            return AddFormField(parent, parentPath, index, ffProps);
        }

        // Allow override via property — same alias set as the no-fieldType path.
        var rawInstr = GetRawFieldInstruction(properties);
        if (rawInstr != null)
            fieldInstr = rawInstr.StartsWith(" ") ? rawInstr : $" {rawInstr} ";

        // CONSISTENCY(field-prop-applicability): the schema in field.json
        // declares per-fieldType-specific props (expression/trueText/
        // falseText for IF, identifier for SEQ, hyperlink for REF, etc.)
        // as universal field-level keys for ergonomic CLI completion.
        // Warn on stderr when a prop that only matters for one fieldType
        // is supplied alongside a different fieldType — Add was silently
        // dropping these per-type props without feedback (Round 5 audit).
        WarnInapplicableFieldProps(properties, effectiveType);

        // Expression fields (raw instruction starting with `=`, e.g. `= 5 * 8 + 2`)
        // arrive via code=/instr= with no named fieldType, so they fall to the `_`
        // default and previously emitted "1" \u2014 Word displays "1" until the user
        // presses F9 to recalc. Use an empty placeholder so Word treats the
        // cached result as "needs recomputation" and shows the real value when
        // the document opens. PAGE/NUMPAGES/SECTION etc. still get "1" (the
        // OOXML convention \u2014 Word auto-updates these on open regardless).
        bool isExpressionField = fieldInstr.TrimStart().StartsWith("=");
        // BUG-R8A(BUG3): the "1" cached placeholder is only correct for numeric
        // page fields (PAGE/PAGEREF/NUMPAGES/SECTION/SECTIONPAGES), where Word
        // auto-updates on open and "1" is a sensible seed. A raw `instruction=`
        // field arrives with effectiveType "field" (no named type), so detect a
        // numeric-page instruction head from the built instruction text too —
        // otherwise `add field --prop instruction="PAGEREF X"` would be denied
        // the seed that `add field --prop fieldType=pageref` gets.
        var instrHead = fieldInstr.TrimStart().Split(' ', 2)[0].ToUpperInvariant();
        bool isNumericPageField =
            effectiveType is "pagenum" or "pagenumber" or "page" or "numpages"
                or "section" or "sectionpages" or "pageref"
            || instrHead is "PAGE" or "PAGEREF" or "NUMPAGES" or "SECTION" or "SECTIONPAGES";
        var fieldPlaceholder = properties.ContainsKey("text")
            ? properties["text"]
            : effectiveType switch
            {
                "mergefield" => $"\u00AB{mergeFieldName}\u00BB",
                "ref" or "noteref" => $"\u00AB{refBookmarkName}\u00BB",
                "styleref" => $"\u00AB{styleRefName}\u00BB",
                "docproperty" => $"\u00AB{docPropertyName}\u00BB",
                "if" => EvaluateIfFieldPlaceholder(properties),
                // SEQ: Word caches the sequence number on insert. Count prior
                // SEQ fields with the same identifier so the appended field
                // seeds with its actual position (Figure 1, Figure 2, …).
                // Previously the cached run stayed empty and get/view showed
                // blank where the number belongs.
                "seq" => FormatSeqCachedValue(
                    CountExistingSeqFields(seqIdentifier),
                    properties.GetValueOrDefault("switches", "")),
                // DATE/TIME family: seed with DateTime.Now formatted via the
                // user's `\@` format switch (if any), otherwise Word-like
                // defaults. The "1" fallback for unrecognized fields is
                // correct for PAGE / NUMPAGES / SECTION etc. but produced
                // a meaningless "1" placeholder for date/time before Word
                // recalculated on open (R11 minor).
                "date" or "createdate" or "savedate" or "printdate"
                    => FormatDateForField(dateFmtVal, "M/d/yyyy"),
                "time" => FormatDateForField(dateFmtVal, "h:mm tt"),
                _ when isExpressionField => "",
                // NOTE(R54-bt-1, by design): a raw `instruction=SEQ …` field is
                // verbatim authoring and stays UNEVALUATED until an explicit
                // `set / recalcFields=seq` (or Word F9) — the SeqEval engine
                // owns \r/\c/\s semantics a naive insert-time count would get
                // wrong. Only the structured fieldType=seq convenience path
                // seeds at insert.
                // BUG-R8A(BUG3): keep the intentional "1" only for numeric page
                // fields. The old `_ => "1"` arm leaked to every unrecognized raw
                // `instruction=` field (SYMBOL/EQ/ADVANCE/TC/...), fabricating a
                // bogus cached "1" reported as evaluated=true. Per the `evaluated`
                // protocol, a field with no genuine cached result must be
                // evaluated=false — omit the cached run (empty placeholder, same
                // path as the expression field) so `view text` shows the sentinel
                // and `view issues` emits field_not_evaluated.
                _ when isNumericPageField => "1",
                _ => ""
            };

        // Build complex field. Canonical shape:
        //   fldChar(begin) + instrText + fldChar(separate) + result + fldChar(end)
        // When the caller passes `noSeparator=true` (typically dump→batch
        // replay of a source whose original field had no separator+result
        // runs), drop fldChar(separate) and the result run — Word treats
        // separator-less fields as "field will be recomputed on open" and
        // renders identically while preserving the source's structural
        // shape on round-trip.
        bool fieldNoSeparator = (properties.TryGetValue("noseparator", out var nsv)
                              || properties.TryGetValue("noSeparator", out nsv))
                              && IsTruthy(nsv);
        // BUG-DUMP-R37-4: <w:fldChar w:fldLock="true"> prevents Word from
        // updating the field on F9/recalc. The lock lives on the BEGIN fldChar
        // (CT_FldChar @w:fldLock); losing it makes a locked field updatable
        // again. The dump emits `fldLock=true` on the synthetic field node;
        // apply it to the begin fldChar (a separator-less complex field still
        // has a begin run, so this covers both shapes).
        bool fieldLocked = (properties.TryGetValue("fldLock", out var flv)
                         || properties.TryGetValue("fldlock", out flv))
                         && IsTruthy(flv);
        var fieldCharBegin = new FieldChar { FieldCharType = FieldCharValues.Begin };
        if (fieldLocked) fieldCharBegin.FieldLock = true;
        var fieldRunBegin = new Run(fieldCharBegin);
        // Reject XML-illegal control chars in the field instruction before
        // they reach the serializer (otherwise the close-time save crashes
        // with "data may be lost").
        OfficeCli.Core.ParseHelpers.ValidateXmlText(fieldInstr, "instr");
        var fieldRunInstr = new Run(new FieldCode(fieldInstr) { Space = SpaceProcessingModeValues.Preserve });
        var fieldRunSep = fieldNoSeparator
            ? null
            : new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
        Run? fieldRunResult = null;
        if (!fieldNoSeparator)
        {
            fieldRunResult = new Run();
            // BUG-DUMP-FIELDTAB: a field's cached display text can carry tabs/line
            // breaks (e.g. a numbered caption "4.1.\tImages" where the tab aligns
            // the title past the list number). Tokenize \t→<w:tab/> and \n→<w:br/>
            // instead of dumping the literal control char into one <w:t> — a raw
            // U+0009 in <w:t> does not advance to the tab stop, so the number/title
            // alignment was lost on every such field round-trip.
            AppendTextWithBreaks(fieldRunResult, fieldPlaceholder);
        }
        var fieldRunEnd = new Run(new FieldChar { FieldCharType = FieldCharValues.End });

        // Apply optional run formatting to all runs
        // BUG-DUMP-FIELDVALIGN: vertAlign (superscript/subscript) on the field
        // runs — a field whose every run (begin/instr/sep/result/end) shares the
        // same <w:vertAlign> (e.g. a superscript cross-reference citation mark).
        // Resolve it up-front so it gates fieldRProps creation alongside
        // font/size/bold/color and is applied uniformly to all field runs below.
        VerticalPositionValues? fieldVertAlign = null;
        if (properties.TryGetValue("vertAlign", out var fVa) || properties.TryGetValue("vertalign", out fVa))
        {
            fieldVertAlign = fVa.ToLowerInvariant() switch
            {
                "superscript" or "super" => VerticalPositionValues.Superscript,
                "subscript" or "sub" => VerticalPositionValues.Subscript,
                "baseline" or "" => VerticalPositionValues.Baseline,
                _ => (VerticalPositionValues?)null
            };
        }
        if (fieldVertAlign == null && properties.TryGetValue("superscript", out var fSup) && IsTruthy(fSup))
            fieldVertAlign = VerticalPositionValues.Superscript;
        if (fieldVertAlign == null && properties.TryGetValue("subscript", out var fSub) && IsTruthy(fSub))
            fieldVertAlign = VerticalPositionValues.Subscript;

        // Per-slot font keys (literal and theme-bound) shared with the run
        // path; a footer PAGE field bound to minorHAnsi keeps its theme face.
        var fieldFontSlotKeys = new[]
        {
            "font.latin", "font.ascii", "font.hAnsi",
            "font.asciiTheme", "font.hAnsiTheme", "font.eaTheme", "font.csTheme",
            // BUG-DUMP-FIELDHINT: rFonts w:hint on a cached result run. Applied
            // via ApplyRunFormatting (writes RunFonts.Hint in CT_RPr order).
            "font.hint",
        };
        bool hasFieldFontSlot = fieldFontSlotKeys.Any(k => properties.ContainsKey(k));
        // BUG-DUMP-RPR-CONTAINER: the fuller rPr vocabulary a field result run can
        // carry (caps/kern/em/highlight/rtl/…) — applied via ApplyRunFormatting below.
        bool hasFieldExtraRpr = WordBatchEmitter.FieldResultExtraRPrKeys.Any(k => properties.ContainsKey(k));
        RunProperties? fieldRProps = null;
        if (properties.TryGetValue("font", out var fFont) || properties.TryGetValue("size", out _) ||
            properties.TryGetValue("bold", out _) || properties.TryGetValue("color", out _) ||
            properties.TryGetValue("italic", out _) || properties.TryGetValue("underline", out _) ||
            properties.TryGetValue("strike", out _) ||
            hasFieldFontSlot || fieldVertAlign != null || hasFieldExtraRpr)
        {
            fieldRProps = new RunProperties();
            // CT_RPr schema order: rFonts → b → ... → color → sz
            if (properties.TryGetValue("font", out var ff))
                fieldRProps.AppendChild(new RunFonts { Ascii = ff, HighAnsi = ff, EastAsia = ff });
            foreach (var slotKey in fieldFontSlotKeys)
            {
                if (properties.TryGetValue(slotKey, out var slotVal))
                    ApplyRunFormatting(fieldRProps, slotKey, slotVal);
            }
            // BUG-DUMP-FIELDBOLD-FALSE: route bold through ApplyRunFormatting so an
            // explicit OFF (<w:b w:val="0"/>) round-trips — a caption field
            // (Table/Figure SEQ) under a bold Caption style turns bold off; the
            // on-only path dropped that override and the caption re-inherited the
            // style's bold, growing every caption line and reflowing the page.
            if (properties.TryGetValue("bold", out var fb))
                ApplyRunFormatting(fieldRProps, "bold", fb);
            // BUG-DUMP-R52-FIELDITALIC: italic/underline/strike on a field's
            // cached result (a title-page <SUBJECT> placeholder rendered italic)
            // were dropped — they were absent from AddField's vocabulary AND from
            // FieldAddSupportedFormatKeys, so the typed `add field` path shed them
            // and (being single-run) the field never took the rich raw-set route.
            // ApplyRunFormatting writes each in CT_RPr schema order.
            if (properties.TryGetValue("italic", out var fi))
                ApplyRunFormatting(fieldRProps, "italic", fi);
            if (properties.TryGetValue("underline", out var fu) && !string.IsNullOrEmpty(fu))
                ApplyRunFormatting(fieldRProps, "underline", fu);
            if (properties.TryGetValue("strike", out var fst) && IsTruthy(fst))
                ApplyRunFormatting(fieldRProps, "strike", fst);
            // BUG-R13B(BUG1): route field color through ApplyRunFormatting (same
            // resolver the run/hyperlink Add paths use) so scheme/theme color
            // names (accent1, dark1, …) write the w:themeColor attribute instead
            // of being passed to SanitizeHex, which rejects non-hex names. Plain
            // hex still works; InsertRunPropInSchemaOrder keeps Color in CT_RPr
            // schema order regardless of the AppendChild sequence above.
            if (properties.TryGetValue("color", out var fc))
                ApplyRunFormatting(fieldRProps, "color", fc);
            if (properties.TryGetValue("size", out var fs))
                fieldRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(fs) * 2, MidpointRounding.AwayFromZero)).ToString() });
            // BUG-DUMP-FIELDVALIGN: vertAlign sits late in CT_RPr (after sz);
            // InsertRunPropInSchemaOrder isn't used here (this block builds the
            // rPr in declaration order), and vertAlign is the last element we
            // append, so AppendChild keeps it in valid schema position.
            if (fieldVertAlign != null)
                fieldRProps.AppendChild(new VerticalTextAlignment { Val = fieldVertAlign.Value });
            // BUG-DUMP-RPR-CONTAINER: apply the fuller rPr vocabulary a field result
            // run can carry (caps/dstrike/outline/shadow/emboss/vanish/spacing/w/kern/
            // position/szCs/highlight/em/lang/rtl/snapToGrid) through the same applier
            // the run/hyperlink paths use, so a single-run formatted field result
            // round-trips losslessly via the typed path. InsertRunPropInSchemaOrder
            // (inside ApplyRunFormatting) keeps CT_RPr order regardless of the
            // AppendChild sequence above.
            foreach (var extraKey in WordBatchEmitter.FieldResultExtraRPrKeys)
                if (properties.TryGetValue(extraKey, out var extraVal))
                    ApplyRunFormatting(fieldRProps, extraKey, extraVal);
        }

        // Final emitted-run ordering: begin → instr → [separate → result] → end
        // (the bracketed pair is skipped when noSeparator=true). Collect in
        // a list so the insertion-site code below doesn't have to repeat the
        // separator-aware conditional.
        var fieldRuns = new List<Run> { fieldRunBegin, fieldRunInstr };
        if (fieldRunSep != null) fieldRuns.Add(fieldRunSep);
        if (fieldRunResult != null) fieldRuns.Add(fieldRunResult);
        fieldRuns.Add(fieldRunEnd);
        // pathRun is what `resultPath` will index to. With separator: the
        // result run (carrying the cached text). Without: end run (no result
        // node exists; pointing at end is the closest stable anchor).
        var pathRun = fieldRunResult ?? fieldRunEnd;

        if (fieldRProps != null)
        {
            foreach (var fr in fieldRuns)
                fr.PrependChild(fieldRProps.CloneNode(true));
        }

        string resultPath;
        if (parent is Paragraph fieldPara)
        {
            // CONSISTENCY(para-path-canonical): canonicalize parentPath to
            // paraId-form so the returned path mirrors what Get later
            // surfaces (paraId is globally unique, works in body / header /
            // footer / cell alike).
            var fieldParaPath = ReplaceTrailingParaSegment(parentPath, fieldPara);
            // CONSISTENCY(paraid-textid-refresh): mirror AddRun — bump
            // textId because the paragraph's content sequence is changing.
            fieldPara.TextId = GenerateParaId();
            // index is a childElement-index (ResolveAnchorPosition counts pPr too).
            // Route the 5 field runs through the pPr-aware multi-insert helper
            // so index 0 clamps forward past ParagraphProperties and they stay
            // in the correct consecutive order.
            if (index.HasValue)
            {
                InsertIntoParagraph(
                    fieldPara,
                    fieldRuns.Cast<OpenXmlElement>().ToArray(),
                    index);
                var runIdxAfterInsert = GetAllRuns(fieldPara).IndexOf(pathRun);
                resultPath = $"{fieldParaPath}/r[{runIdxAfterInsert + 1}]";
            }
            else
            {
                foreach (var fr in fieldRuns) fieldPara.AppendChild(fr);
                var runs = GetAllRuns(fieldPara);
                var runIdx = PathIndex.FromArrayIndex(runs.IndexOf(pathRun));
                resultPath = $"{fieldParaPath}/r[{runIdx}]";
            }
        }
        else if (parent is Hyperlink fieldHl && fieldHl.Parent is Paragraph fieldHlPara)
        {
            // BUG-DUMP18-02: field added with parent=w:hyperlink. The 5 field
            // runs become direct children of the hyperlink so they render
            // INSIDE the hyperlink scope (mirrors AddEquation's Hyperlink
            // branch added in BUG-DUMP15-04).
            fieldHlPara.TextId = GenerateParaId();
            if (index.HasValue)
            {
                var children = fieldHl.ChildElements.ToList();
                if (index.Value < children.Count)
                {
                    var anchor = children[index.Value];
                    foreach (var r in fieldRuns) anchor.InsertBeforeSelf(r);
                }
                else
                {
                    foreach (var r in fieldRuns) fieldHl.AppendChild(r);
                }
            }
            else
            {
                foreach (var r in fieldRuns) fieldHl.AppendChild(r);
            }
            var fieldHlParaPath = ReplaceTrailingParaSegment(parentPath, fieldHlPara);
            var slashIdxHl = fieldHlParaPath.LastIndexOf("/hyperlink[", StringComparison.Ordinal);
            var paraPathOnly = slashIdxHl > 0 ? fieldHlParaPath.Substring(0, slashIdxHl) : fieldHlParaPath;
            var hlIdxF = fieldHlPara.Elements<Hyperlink>().TakeWhile(h => !ReferenceEquals(h, fieldHl)).Count() + 1;
            var runIdxAfter = GetAllRuns(fieldHlPara).IndexOf(pathRun);
            resultPath = $"{paraPathOnly}/hyperlink[{hlIdxF}]/r[{runIdxAfter + 1}]";
        }
        else if (parent is Run hostRun && hostRun.Parent is Paragraph hostRunPara)
        {
            hostRunPara.TextId = GenerateParaId();
            OpenXmlElement cursor = hostRun;
            foreach (var fr in fieldRuns)
            {
                cursor.InsertAfterSelf(fr);
                cursor = fr;
            }
            var hostParaPath = ReplaceTrailingParaSegment(parentPath, hostRunPara);
            var slashIdx = hostParaPath.LastIndexOf("/r[", StringComparison.Ordinal);
            if (slashIdx > 0) hostParaPath = hostParaPath.Substring(0, slashIdx);
            var runIdxAfter = GetAllRuns(hostRunPara).IndexOf(pathRun);
            resultPath = $"{hostParaPath}/r[{runIdxAfter + 1}]";
        }
        else
        {
            // Create a new paragraph containing the field
            var fNewPara = new Paragraph();
            var fPProps = new ParagraphProperties();
            if (properties.TryGetValue("align", out var fAlign) || properties.TryGetValue("alignment", out fAlign))
                fPProps.Justification = new Justification { Val = ParseJustification(fAlign) };
            fNewPara.AppendChild(fPProps);
            foreach (var fr in fieldRuns) fNewPara.AppendChild(fr);
            // CONSISTENCY(paraid-global-uniqueness): newly-created paragraphs
            // get a paraId from the global counter so they remain addressable
            // by paraId regardless of which container they land in.
            AssignParaId(fNewPara);
            InsertAtIndexOrAppend(parent, fNewPara, index);
            // CONSISTENCY(para-path-canonical): paraId-form path works in
            // every container (body / header / footer / cell). Same shape
            // as AddBreak's new-paragraph branch.
            if (parent is Body)
            {
                var fIdx2 = body.Elements<Paragraph>().TakeWhile(p => p != fNewPara).Count();
                resultPath = $"/body/{BuildParaPathSegment(fNewPara, fIdx2 + 1)}";
            }
            else
            {
                var fIdx2 = parent.Elements<Paragraph>().TakeWhile(p => p != fNewPara).Count();
                resultPath = $"{parentPath}/{BuildParaPathSegment(fNewPara, fIdx2 + 1)}";
            }
        }
        // BUG-DUMP-DELFIELD: a field that was tracked-inserted/deleted in the
        // source (its runs wrapped in <w:del>/<w:ins>) must round-trip the
        // revision. CollapseFieldChains propagated revision.type/.author/.date/
        // .id from the wrapper onto the field synth, and the emitter forwarded
        // them here. Wrap each rebuilt field run in the matching marker —
        // reusing the same WrapRunAs* helpers the run path uses, so a deleted
        // HYPERLINK keeps its <w:del> + <w:delInstrText>/<w:delText> instead of
        // resurrecting as live text. del converts FieldCode→DeletedFieldCode
        // (delInstrText is the only valid field-code form inside <w:del>).
        WrapRunsInRevision(fieldRuns, properties);
        return resultPath;
    }

    // Wrap freshly-built runs in a tracked-change marker when the caller supplied
    // revision.type (= ins/del/moveFrom/moveTo) + attribution. Used by the field
    // rebuild (deleted/inserted hyperlinks & fields) AND by AddBreak (a tracked
    // page/column break — BUG-DUMP-DELBREAK). Mirrors the per-run revision wrap
    // used by Set/Add on plain runs; the only field-specific twist is that a
    // <w:del>-wrapped instruction run must carry its code as <w:delInstrText>
    // (DeletedFieldCode), not <w:instrText> (FieldCode) — ECMA-376 §17.16.23 —
    // which is a guarded no-op for runs (e.g. break runs) that carry no FieldCode.
    // format/paraMark* kinds don't apply to a run-level wrap and are ignored here.
    private void WrapRunsInRevision(List<Run> runs, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("revision.type", out var revType)
            || string.IsNullOrWhiteSpace(revType))
            return;
        revType = revType.Trim();
        if (revType is not ("ins" or "del" or "moveFrom" or "moveTo"))
            return;

        var author = properties.GetValueOrDefault("revision.author") ?? "";
        DateTime date = DateTime.UtcNow;
        if (properties.TryGetValue("revision.date", out var dateStr)
            && DateTime.TryParse(dateStr, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind, out var parsedDate))
            date = parsedDate;
        // moveFrom/moveTo: every run in the (contiguous) move shares ONE id so
        // the moveFrom half pairs with its moveTo half — WrapRunAsMove* also
        // brackets the range with the shared Name="Move_{id}". An explicit id is
        // required to pair the two halves across separate AddField calls; fall
        // back to a generated one (single-call moves can't pair anyway).
        // ins/del: a single source <w:del>/<w:ins> is split into one wrapper per
        // field run on this rebuild — each MUST get a UNIQUE w:id (ECMA-376
        // requires w:id uniqueness; Word tolerates duplicates but strict
        // validation rejects them). So we generate a fresh id per run rather
        // than reusing revision.id (which addresses the source's single marker).
        var explicitId = properties.GetValueOrDefault("revision.id");
        var moveId = !string.IsNullOrEmpty(explicitId) ? explicitId : GenerateRevisionId();

        foreach (var fr in runs)
        {
            if (fr.Parent == null) continue;
            if (revType == "del")
            {
                // FieldCode (w:instrText) → DeletedFieldCode (w:delInstrText)
                // before the w:t→w:delText conversion inside WrapRunAsDeleted.
                foreach (var fc in fr.Elements<FieldCode>().ToList())
                {
                    var dfc = new DeletedFieldCode(fc.Text ?? "") { Space = fc.Space };
                    fc.Parent?.ReplaceChild(dfc, fc);
                }
                WrapRunAsDeleted(fr, author, date, explicitId: null);
            }
            else if (revType == "ins")
            {
                WrapRunAsInserted(fr, author, date, explicitId: null);
            }
            else if (revType == "moveFrom")
            {
                WrapRunAsMoveFrom(fr, author, date, moveId);
            }
            else // moveTo
            {
                WrapRunAsMoveTo(fr, author, date, moveId);
            }
        }
    }

    // CONSISTENCY(canonical-keys): the raw field instruction can be passed
    // under `instr` (canonical, mirrors Get readback), `instruction`
    // (legacy, predates the schema rename), or `code` (alias documented in
    // field.json). All three resolve to the same string. Wrapping spaces
    // are reserved by the caller — the wrapping logic at the call site
    // adds them when missing.
    private static string? GetRawFieldInstruction(Dictionary<string, string> properties)
    {
        // Treat empty / whitespace-only as absent so a placeholder
        // `instr=""` doesn't short-circuit the alias chain and emit a
        // degenerate empty <w:instrText> while a non-empty `instruction=`
        // or `code=` is also supplied. Found via Round 7 fuzz BUG-R7-3.
        static string? NotBlank(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;
        return NotBlank(properties.GetValueOrDefault("instr"))
            ?? NotBlank(properties.GetValueOrDefault("instruction"))
            ?? NotBlank(properties.GetValueOrDefault("code"));
    }

    // CONSISTENCY(field-prop-applicability): map each fieldType to the
    // per-type props the Add path actually reads. Anything outside the
    // universal set + this map's value is unused for that fieldType and
    // should surface as a warning so the user notices the typo / wrong
    // assumption (e.g. supplying bookmarkName=... with fieldType=if).
    private static readonly Dictionary<string, string[]> FieldTypeProps =
        new(StringComparer.OrdinalIgnoreCase)
    {
        ["mergefield"] = new[] { "name", "fieldname", "switches" },
        ["ref"] = new[] { "name", "fieldname", "bookmarkname", "bookmark", "hyperlink" },
        ["pageref"] = new[] { "name", "fieldname", "bookmarkname", "bookmark", "hyperlink" },
        ["noteref"] = new[] { "name", "fieldname", "bookmarkname", "bookmark", "hyperlink" },
        ["seq"] = new[] { "identifier", "id", "name", "switches" },
        ["styleref"] = new[] { "stylename", "name" },
        ["docproperty"] = new[] { "propertyname", "name" },
        ["if"] = new[] { "expression", "condition", "truetext", "falsetext" },
        ["date"] = new[] { "format" },
        ["time"] = new[] { "format" },
        ["createdate"] = new[] { "format" },
        ["savedate"] = new[] { "format" },
        ["printdate"] = new[] { "format" },
        ["hyperlink"] = new[] { "url", "anchor" },
    };

    // Universal props every fieldType accepts: routing keys, run rPr,
    // raw-instruction override, anchor placement, cached display text.
    private static readonly HashSet<string> FieldUniversalProps =
        new(StringComparer.OrdinalIgnoreCase)
    {
        "fieldtype", "type", "instr", "instruction", "code",
        "text", "font", "size", "bold", "color",
        // BUG-DUMP-FIELDVALIGN: field-wide vertical alignment (superscript /
        // subscript) applies to ANY field type — the common case is a
        // superscript cross-reference citation mark. Consumed by the
        // fieldRProps builder, applied uniformly to all field runs; same
        // universal-run-format class as font/size/bold/color above.
        "vertalign", "vertAlign", "superscript", "subscript",
        "index", "after", "before",
        // BUG-R7A: `switches` carries residual general field switches
        // (`\* roman`, `\* MERGEFORMAT`, `\p`, `\h`, …) for every typed
        // field on dump→batch round-trips, not just SEQ/MERGEFIELD. The
        // bare-instruction arms (PAGE/NUMPAGES/AUTHOR/TITLE/…) and
        // DATE/TIME and REF/PAGEREF/NOTEREF all now splice it, so it is
        // genuinely universal — promote it out of the per-type lists.
        "switches",
        // BUG-DUMP-DELFIELD: a field can be tracked-inserted/deleted; the
        // revision wrap applies to ANY fieldType. Consumed by
        // WrapFieldRunsInRevision, not the per-type instruction builders.
        "revision.type", "revision.author", "revision.date", "revision.id",
        "noseparator", "noSeparator",
        // BUG-DUMP-R37-4: <w:fldChar w:fldLock="true"> — F9/recalc lock applies
        // to ANY field type. Consumed by the begin-fldChar builder, not a
        // per-type instruction builder, so it is universal.
        "fldlock", "fldLock",
    };

    // Render today's DateTime for the result-run placeholder of a DATE/TIME
    // field. `userFormat` is the value of --prop format=… (the same string
    // Word writes after \@ in the field instruction); empty/missing falls
    // back to a Word-like default. Invalid format strings degrade silently
    // to the default rather than throwing — the seeded value is cosmetic
    // (Word recalculates on open), so a malformed format string would only
    // be visible briefly and shouldn't fail the Add.
    private static string FormatDateForField(string? userFormat, string defaultFormat)
    {
        var fmt = string.IsNullOrWhiteSpace(userFormat) ? defaultFormat : userFormat;
        try
        {
            return DateTime.Now.ToString(fmt, System.Globalization.CultureInfo.InvariantCulture);
        }
        catch (FormatException)
        {
            return DateTime.Now.ToString(defaultFormat, System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private static void WarnInapplicableFieldProps(
        Dictionary<string, string> properties, string effectiveType)
    {
        var typeProps = FieldTypeProps.GetValueOrDefault(effectiveType)
            ?? Array.Empty<string>();
        var typeSet = new HashSet<string>(typeProps, StringComparer.OrdinalIgnoreCase);
        foreach (var key in properties.Keys)
        {
            if (FieldUniversalProps.Contains(key)) continue;
            if (typeSet.Contains(key)) continue;
            // Any other prop is known to no fieldType-specific consumer —
            // the BuildXxxFieldInstruction path won't read it. Surface a
            // warning so silent-ignore (Round 5 R5-T1 / R5-F2) becomes
            // visible. Use stderr, exit code stays 0 (consistent with
            // other Add warning paths via Console.Error.WriteLine).
            Console.Error.WriteLine(
                $"Warning: prop '{key}' is not applicable to field type '{effectiveType}' — silently ignored. " +
                $"Applicable to '{effectiveType}': {(typeProps.Length > 0 ? string.Join(", ", typeProps) : "none beyond universal")}.");
        }
    }

    // BUG-DUMP15-02: HYPERLINK fields may carry any combination of base URL,
    // `\l "anchor"`, and `\o "tooltip"`. Reconstruct the full instruction
    // from whichever props are present so dump→batch round-trips do not
    // silently drop URL or tooltip.
    private static string BuildHyperlinkFieldInstruction(Dictionary<string, string> properties)
    {
        properties.TryGetValue("url", out var hUrl);
        properties.TryGetValue("anchor", out var hAnchor);
        properties.TryGetValue("tooltip", out var hTooltip);
        if (string.IsNullOrEmpty(hUrl) && string.IsNullOrEmpty(hAnchor))
            throw new ArgumentException(
                "HYPERLINK field requires either 'url' or 'anchor' property.");
        var sb = new System.Text.StringBuilder(" HYPERLINK");
        if (!string.IsNullOrEmpty(hUrl)) sb.Append($" \"{hUrl}\"");
        if (!string.IsNullOrEmpty(hAnchor)) sb.Append($" \\l \"{hAnchor}\"");
        if (!string.IsNullOrEmpty(hTooltip)) sb.Append($" \\o \"{hTooltip}\"");
        sb.Append(' ');
        return sb.ToString();
    }

    private static string BuildIfFieldInstruction(Dictionary<string, string> properties)
    {
        var expression = properties.GetValueOrDefault("expression")
                      ?? properties.GetValueOrDefault("condition");
        if (string.IsNullOrWhiteSpace(expression))
            throw new ArgumentException("IF requires an 'expression' property (e.g. --prop expression=\"MERGEFIELD Gender = \\\"Male\\\"\").");
        var trueText = properties.GetValueOrDefault("trueText", properties.GetValueOrDefault("truetext", ""));
        var falseText = properties.GetValueOrDefault("falseText", properties.GetValueOrDefault("falsetext", ""));
        return $" IF {expression} \"{trueText}\" \"{falseText}\" ";
    }

    private string AddBreak(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // Insert an explicit page break, column break, or line break
        var breakType = type.ToLowerInvariant() switch
        {
            "columnbreak" => BreakValues.Column,
            _ => BreakValues.Page
        };
        // CONSISTENCY(canonical-keys): accept both `type=` (legacy alias)
        // and `breakType=` (Set/Get canonical key) on Add — silent-ignore
        // of breakType= violates project red line (commit 19b3dd5b);
        // forcing users to know that Add wants `type` while Set/Get want
        // `breakType` is precisely the alias trap that policy bans.
        if (properties.TryGetValue("type", out var brType)
            || properties.TryGetValue("breakType", out brType)
            || properties.TryGetValue("breaktype", out brType))
        {
            breakType = brType.ToLowerInvariant() switch
            {
                "page" => BreakValues.Page,
                "column" => BreakValues.Column,
                "textwrapping" or "line" => BreakValues.TextWrapping,
                _ => throw new ArgumentException($"Invalid break type: '{brType}'. Valid values: page, column, line, textwrapping.")
            };
        }

        var brk = new Break { Type = breakType };
        // <w:br w:clear> — float-clearing for text-wrapping breaks (Word's
        // "clear all/left/right"). Round-tripped via the breakClear key.
        if (properties.TryGetValue("breakClear", out var brkClear)
            || properties.TryGetValue("breakclear", out brkClear)
            || properties.TryGetValue("clear", out brkClear))
        {
            var clearCanon = brkClear.ToLowerInvariant() switch
            {
                "all" => "all",
                "left" => "left",
                "right" => "right",
                "none" => "none",
                _ => throw new ArgumentException($"Invalid break clear: '{brkClear}'. Valid values: all, left, right, none.")
            };
            brk.Clear = new EnumValue<BreakTextRestartLocationValues>(new BreakTextRestartLocationValues(clearCanon));
        }
        var brkRun = new Run(brk);
        // BUG-DUMP-BREAKRPR: a break-only run (<w:r><w:rPr>…</w:rPr><w:br/></w:r>)
        // carries an rPr whose font/size sets the height of the line the break
        // starts. The verbatim raw-set fallback in TryEmitBreakRun only fires for
        // /body hosts, so a break inside a table cell rebuilt as a bare
        // <w:r><w:br/></w:r> and the broken line collapsed to the default font
        // size — inflating cell/row height and drifting the table. Re-apply the
        // forwarded rPr here so it round-trips in every container.
        if (properties.TryGetValue("breakRunRpr", out var brkRpr)
            && !string.IsNullOrWhiteSpace(brkRpr)
            && brkRpr.Contains("rPr", StringComparison.Ordinal))
        {
            try { brkRun.PrependChild(new RunProperties(brkRpr)); } catch { /* malformed: skip */ }
        }

        string resultPath;
        if (parent is Paragraph brkPara)
        {
            // CONSISTENCY(paraid-textid-refresh): mirror AddRun — bump
            // textId so revision/diff tooling sees the paragraph as
            // modified. Done before we possibly take an early return on
            // the index-resolved path to make sure both branches stamp it.
            brkPara.TextId = GenerateParaId();
            // index is a childElement-index (ResolveAnchorPosition counts pPr).
            // pPr-aware insert keeps pPr as the first child of <w:p>.
            InsertIntoParagraph(brkPara, brkRun, index);
            var brkRunIdx = PathIndex.FromArrayIndex(GetAllRuns(brkPara).IndexOf(brkRun));
            // CONSISTENCY(para-path-canonical): parentPath already targets
            // the paragraph; replacing its trailing /p[...] segment with
            // paraId-form yields a path that mirrors what Get later
            // surfaces and works regardless of which container the
            // paragraph lives in (body / header / footer / cell). The
            // previous /body/-hardcoded path produced wrong prefixes for
            // breaks added inside header/footer paragraphs.
            var canonicalParaPath = ReplaceTrailingParaSegment(parentPath, brkPara);
            resultPath = $"{canonicalParaPath}/r[{brkRunIdx}]";
            // BUG-DUMP-DELBREAK: a tracked-DELETED/inserted break must keep its
            // <w:del>/<w:ins> wrapper, else a deleted (invisible) page break
            // resurrects as a live break and inflates the page count.
            WrapRunsInRevision(new List<Run> { brkRun }, properties);
        }
        else
        {
            // Create a new empty paragraph with the break and insert into the
            // ACTUAL parent (not hard-coded body) so /header[N], /footer[N],
            // table cells, etc. receive the new paragraph. /styles is blocked
            // earlier by ValidateParentChild.
            var brkNewPara = new Paragraph(brkRun);
            // CONSISTENCY(paraid-global-uniqueness): every newly-created
            // paragraph gets a paraId so it remains addressable by paraId
            // across containers (body / headers / footers / cells); the
            // global counter guarantees uniqueness so the same path form
            // works everywhere.
            AssignParaId(brkNewPara);
            InsertAtIndexOrAppend(parent, brkNewPara, index);
            // BUG-DUMP-DELBREAK: see the in-paragraph branch above — preserve the
            // tracked-change wrapper on the rebuilt break run.
            WrapRunsInRevision(new List<Run> { brkRun }, properties);
            // CONSISTENCY(para-path-canonical): paraId-form is valid in
            // every container (the paraId is globally unique and Navigation
            // resolves it inside header/footer/cell parts as well as body).
            // Use the same BuildParaPathSegment helper everywhere instead
            // of a body-only specialization.
            if (parent is Body)
            {
                var brkIdx = body.Elements<Paragraph>().TakeWhile(p => p != brkNewPara).Count();
                resultPath = $"/body/{BuildParaPathSegment(brkNewPara, brkIdx + 1)}";
            }
            else
            {
                var brkIdx = parent.Elements<Paragraph>().TakeWhile(p => p != brkNewPara).Count();
                resultPath = $"{parentPath}/{BuildParaPathSegment(brkNewPara, brkIdx + 1)}";
            }
        }
        return resultPath;
    }

    private string AddSdt(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // Reject SDT nested inside a plain-text SDT (sdtPr/<w:text/>): the
        // outer marks its content as plain-text-only, so adding any SDT
        // descendant produces OOXML that Word rejects (error 0x422). Walk
        // ancestors AND the parent's own SdtBlock chain (when parent is the
        // SDT's content paragraph) so both block- and inline-nest paths are
        // caught before any mutation. Mirrors the R22 nested-textbox guard.
        for (var cur = parent; cur != null; cur = cur.Parent)
        {
            var sdtPr = (cur as SdtBlock)?.SdtProperties
                ?? (cur as SdtRun)?.SdtProperties
                ?? cur.GetFirstChild<SdtProperties>();
            if (sdtPr?.GetFirstChild<SdtContentText>() != null)
                throw new ArgumentException(
                    "Cannot nest an SDT inside a plain-text SDT (sdtPr/<w:text/>). The outer control marks its content as plain-text-only; nested SDTs produce OOXML that Word rejects (error 0x422).");
        }

        // Verbatim carrier (dump-emitted): a rich BLOCK content control whose
        // content references parts/relationships (a cover page with anchored
        // textboxes and a logo image). Same shape as the activex/diagram/
        // vmlshape carriers — sdtXml is the whole <w:sdt> element verbatim,
        // part{N}/ext{N} ship the referenced parts; rel ids are rewritten to
        // the freshly assigned ones before injection.
        if (properties.TryGetValue("sdtXml", out var sdtCarrierXml)
            && !string.IsNullOrEmpty(sdtCarrierXml))
        {
            var carrierHost = ResolveImageHostPart(parent);
            var rewrite = MaterializeInlinedParts(carrierHost, properties, "sdt");
            // A paragraph host means a run-level (inline) control — the same
            // <w:sdt> XML, but the typed wrapper must be SdtRun so the SDK
            // object model matches CT_P's particle (an inline picture control
            // shipped by the dump carrier lands here with parent = p[last()]).
            if (parent is Paragraph)
            {
                var sdtRun = new SdtRun(rewrite(sdtCarrierXml));
                AppendToParent(parent, sdtRun);
                var sdtRunIdx = PathIndex.FromArrayIndex(parent.Elements<SdtRun>().ToList().IndexOf(sdtRun));
                return $"{parentPath}/sdt[{sdtRunIdx}]";
            }
            var sdtBlock = new SdtBlock(rewrite(sdtCarrierXml));
            AppendToParent(parent, sdtBlock);
            var sdtIdx2 = PathIndex.FromArrayIndex(parent.Elements<SdtBlock>().ToList().IndexOf(sdtBlock));
            return $"{parentPath}/sdt[{sdtIdx2}]";
        }

        // Case-insensitive lookup to support camelCase keys like "sdtType", "controlType", etc.
        // CONSISTENCY(tracking-preservation): mirror WordHandler.Add.cs:32-40 — never copy a
        // TrackingPropertyDictionary into a plain Dictionary, otherwise every TryGetValue
        // here stops being recorded and unknown props silently slip through instead of
        // surfacing as UNSUPPORTED warnings (handler-as-truth model). The dispatcher
        // already gives us an OrdinalIgnoreCase Dictionary or a TrackingPropertyDictionary,
        // so the only honest move is to pass it through unchanged when possible.
        var ciProps = properties is OfficeCli.Core.TrackingPropertyDictionary
            ? properties
            : (properties.Comparer == StringComparer.OrdinalIgnoreCase
                ? properties
                : new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase));

        // Add a Structured Document Tag (Content Control)
        // Canonical key is "type" (per schemas/help/docx/sdt.json); "sdttype" / "controltype"
        // retained as legacy aliases for backward-compat.
        var sdtType = ciProps.GetValueOrDefault("type",
            ciProps.GetValueOrDefault("sdttype",
                ciProps.GetValueOrDefault("controltype", "text"))).ToLowerInvariant();
        // Schema-honesty: reject values the SDT builder does not emit the
        // correct child elements for. Keeps the schema and runtime in sync
        // instead of silently falling back to plain-text SDT.
        // BUG-DUMP-R42-7/8: `group` (grouping content control) and `picture`
        // (picture content control) carry a single empty <w:group/> / <w:picture/>
        // marker in sdtPr. Previously unsupported, so a dump round-trip dropped
        // the marker and degraded the control to a generic rich-text SDT.
        var supportedSdtTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "text", "plaintext", "richtext", "rich",
            "dropdown", "dropdownlist", "combobox", "combo",
            "date", "datepicker",
            "group", "picture", "checkbox"
        };
        if (!supportedSdtTypes.Contains(sdtType))
            throw new NotSupportedException(
                $"SDT type '{sdtType}' is not implemented. Supported: text, richtext, dropdown, combobox, date, group, picture, checkbox. " +
                "Create the content control in Word, then edit via CLI.");
        var alias = ciProps.GetValueOrDefault("alias", ciProps.GetValueOrDefault("name", ""));
        var tag = ciProps.GetValueOrDefault("tag", "");
        var lockVal = ciProps.GetValueOrDefault("lock", "");
        var sdtText = ciProps.GetValueOrDefault("text", "");

        // Determine block-level vs inline
        bool isInline = parent is Paragraph;

        string resultPath;
        if (isInline)
        {
            // Inline SDT (SdtRun) inside a paragraph
            var sdtRun = new SdtRun();
            var sdtProps = new SdtProperties();

            // ID
            var inlineSdtIdVal = NextSdtId();
            sdtProps.AppendChild(new SdtId { Val = inlineSdtIdVal });

            if (!string.IsNullOrEmpty(alias))
                sdtProps.AppendChild(new SdtAlias { Val = alias });
            if (!string.IsNullOrEmpty(tag))
                sdtProps.AppendChild(new Tag { Val = tag });
            if (!string.IsNullOrEmpty(lockVal))
            {
                sdtProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock
                {
                    Val = lockVal.ToLowerInvariant() switch
                    {
                        "contentlocked" or "content" => LockingValues.ContentLocked,
                        "sdtlocked" or "sdt" => LockingValues.SdtLocked,
                        "sdtcontentlocked" or "both" => LockingValues.SdtContentLocked,
                        "unlocked" or "none" => LockingValues.Unlocked,
                        _ => throw new ArgumentException($"Invalid lock value: '{lockVal}'. Valid values: unlocked, contentLocked, sdtLocked, sdtContentLocked.")
                    }
                });
            }

            // Content type definition
            switch (sdtType)
            {
                case "dropdown" or "dropdownlist":
                {
                    var ddl = new SdtContentDropDownList();
                    // CONSISTENCY(sdt-items-alias): accept "choices" as alias for "items"
                    // — matches the natural CLI vocabulary users reach for first.
                    if (ciProps.TryGetValue("items", out var items)
                        || ciProps.TryGetValue("choices", out items))
                    {
                        foreach (var li in ParseSdtItems(items))
                            ddl.AppendChild(li);
                    }
                    sdtProps.AppendChild(ddl);
                    break;
                }
                case "combobox" or "combo":
                {
                    var cb = new SdtContentComboBox();
                    if (ciProps.TryGetValue("items", out var items)
                        || ciProps.TryGetValue("choices", out items))
                    {
                        foreach (var li in ParseSdtItems(items))
                            cb.AppendChild(li);
                    }
                    sdtProps.AppendChild(cb);
                    break;
                }
                case "date" or "datepicker":
                    var datePr = new SdtContentDate();
                    if (ciProps.TryGetValue("format", out var dateFmt))
                        datePr.DateFormat = new DateFormat { Val = dateFmt };
                    else
                        datePr.DateFormat = new DateFormat { Val = "yyyy-MM-dd" };
                    sdtProps.AppendChild(datePr);
                    break;
                case "group":
                    // BUG-DUMP-R42-7: grouping content control — empty <w:group/>.
                    sdtProps.AppendChild(new SdtContentGroup());
                    break;
                case "picture":
                    // BUG-DUMP-R42-8: picture content control — empty <w:picture/>.
                    sdtProps.AppendChild(new SdtContentPicture());
                    break;
                case "checkbox":
                    // Word checkbox content control: a <w14:checkbox> marker in sdtPr
                    // (checked flag + checked/unchecked box glyphs). The SDK auto-
                    // declares the w14 namespace on serialization.
                    sdtProps.AppendChild(BuildSdtCheckBox(IsTruthy(ciProps.GetValueOrDefault("checked", "false"))));
                    break;
                case "richtext" or "rich":
                    // Rich text has no specific type element (absence of w:text means rich text)
                    break;
                default: // "text" or "plaintext"
                    sdtProps.AppendChild(new SdtContentText());
                    break;
            }

            ApplySdtExtraProps(sdtProps, ciProps);

            sdtRun.AppendChild(sdtProps);
            var sdtContent = new SdtContentRun();
            // Checkbox controls carry the box glyph (☒ checked / ☐ unchecked) as
            // their content run when no explicit text is supplied.
            var inlineSeedText = sdtType == "checkbox" && string.IsNullOrEmpty(sdtText)
                ? (IsTruthy(ciProps.GetValueOrDefault("checked", "false")) ? "☒" : "☐")
                : sdtText;
            var contentRun = new Run(new Text(inlineSeedText) { Space = SpaceProcessingModeValues.Preserve });

            // CONSISTENCY(rtl-cascade): mirror AddRun (Add.Text.cs:373-376).
            // When the host paragraph is direction=rtl (pPr/bidi or mark
            // rPr/rtl), the new contentRun must carry rPr/rtl — paragraph
            // mark rPr does not cascade to inner runs in OOXML; only style
            // does. Without this, SDT body in an RTL paragraph renders LTR.
            if (parent is Paragraph hostPara && hostPara.ParagraphProperties is { } hostPPr)
            {
                var hostBidi = hostPPr.GetFirstChild<BiDi>();
                var hostMarkRtl = hostPPr.ParagraphMarkRunProperties?
                    .GetFirstChild<RightToLeftText>();
                if (hostBidi != null || hostMarkRtl != null)
                {
                    var crProps = contentRun.RunProperties ??= new RunProperties();
                    if (crProps.GetFirstChild<RightToLeftText>() == null)
                        crProps.AppendChild(new RightToLeftText());
                }
            }
            sdtContent.AppendChild(contentRun);
            sdtRun.AppendChild(sdtContent);

            // index is a childElement-index (ResolveAnchorPosition counts pPr).
            // pPr-aware insert so an index at pPr clamps forward to keep pPr first.
            var sdtPara = (Paragraph)parent;
            InsertIntoParagraph(sdtPara, sdtRun, index);
            // Build stable @paraId= and @sdtId= based path. Determine the
            // root segment (body / header[N] / footer[N]) from the caller's
            // parentPath so returned paths actually resolve when the parent
            // paragraph lives in a header or footer part.
            var inlineRoot = ExtractRootSegment(parentPath);
            var inlineParaId = ((Paragraph)parent).ParagraphId?.Value;
            string inlineParaSegment;
            if (!string.IsNullOrEmpty(inlineParaId))
            {
                inlineParaSegment = $"p[@paraId={inlineParaId}]";
            }
            else
            {
                var parentContainer = parent.Parent;
                var paraIdxIn = parentContainer?.Elements<Paragraph>().TakeWhile(p => p != parent).Count() ?? 0;
                inlineParaSegment = $"p[{paraIdxIn + 1}]";
            }
            resultPath = $"{inlineRoot}/{inlineParaSegment}/sdt[@sdtId={inlineSdtIdVal}]";
        }
        else
        {
            // Block-level SDT (SdtBlock)
            var sdtBlock = new SdtBlock();
            var sdtProps = new SdtProperties();

            sdtProps.AppendChild(new SdtId { Val = NextSdtId() });

            if (!string.IsNullOrEmpty(alias))
                sdtProps.AppendChild(new SdtAlias { Val = alias });
            if (!string.IsNullOrEmpty(tag))
                sdtProps.AppendChild(new Tag { Val = tag });
            if (!string.IsNullOrEmpty(lockVal))
            {
                sdtProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock
                {
                    Val = lockVal.ToLowerInvariant() switch
                    {
                        "contentlocked" or "content" => LockingValues.ContentLocked,
                        "sdtlocked" or "sdt" => LockingValues.SdtLocked,
                        "sdtcontentlocked" or "both" => LockingValues.SdtContentLocked,
                        "unlocked" or "none" => LockingValues.Unlocked,
                        _ => throw new ArgumentException($"Invalid lock value: '{lockVal}'. Valid values: unlocked, contentLocked, sdtLocked, sdtContentLocked.")
                    }
                });
            }

            switch (sdtType)
            {
                case "dropdown" or "dropdownlist":
                {
                    var ddl = new SdtContentDropDownList();
                    // CONSISTENCY(sdt-items-alias): accept "choices" as alias for "items"
                    // — matches the natural CLI vocabulary users reach for first.
                    if (ciProps.TryGetValue("items", out var items)
                        || ciProps.TryGetValue("choices", out items))
                    {
                        foreach (var li in ParseSdtItems(items))
                            ddl.AppendChild(li);
                    }
                    sdtProps.AppendChild(ddl);
                    break;
                }
                case "combobox" or "combo":
                {
                    var cb = new SdtContentComboBox();
                    if (ciProps.TryGetValue("items", out var items)
                        || ciProps.TryGetValue("choices", out items))
                    {
                        foreach (var li in ParseSdtItems(items))
                            cb.AppendChild(li);
                    }
                    sdtProps.AppendChild(cb);
                    break;
                }
                case "date" or "datepicker":
                    var datePr = new SdtContentDate();
                    if (ciProps.TryGetValue("format", out var dateFmt))
                        datePr.DateFormat = new DateFormat { Val = dateFmt };
                    else
                        datePr.DateFormat = new DateFormat { Val = "yyyy-MM-dd" };
                    sdtProps.AppendChild(datePr);
                    break;
                case "group":
                    // BUG-DUMP-R42-7: grouping content control — empty <w:group/>.
                    sdtProps.AppendChild(new SdtContentGroup());
                    break;
                case "picture":
                    // BUG-DUMP-R42-8: picture content control — empty <w:picture/>.
                    sdtProps.AppendChild(new SdtContentPicture());
                    break;
                case "checkbox":
                    // Word checkbox content control — see inline branch above.
                    sdtProps.AppendChild(BuildSdtCheckBox(IsTruthy(ciProps.GetValueOrDefault("checked", "false"))));
                    break;
                case "richtext" or "rich":
                    break;
                default:
                    sdtProps.AppendChild(new SdtContentText());
                    break;
            }

            ApplySdtExtraProps(sdtProps, ciProps);

            sdtBlock.AppendChild(sdtProps);
            var sdtContent = new SdtContentBlock();
            var blockSeedText = sdtType == "checkbox" && string.IsNullOrEmpty(sdtText)
                ? (IsTruthy(ciProps.GetValueOrDefault("checked", "false")) ? "☒" : "☐")
                : sdtText;
            var contentPara = new Paragraph(new Run(new Text(blockSeedText) { Space = SpaceProcessingModeValues.Preserve }));
            sdtContent.AppendChild(contentPara);
            sdtBlock.AppendChild(sdtContent);

            InsertAtIndexOrAppend(parent, sdtBlock, index);
            // BUG-DUMP-R47-8: a table cell whose sole source content is a block
            // SDT — <w:tc><w:tcPr/><w:sdt>…</w:sdt></w:tc>, common in form
            // templates that wrap each cell value in a content control — has NO
            // standalone paragraph. AddTable seeds every cell with one empty
            // paragraph; appending the SDT leaves that seed as a spurious leading
            // empty paragraph, so the cell renders two lines (a blank line above
            // the value) and the row grows — drifting table and page layout. When
            // the SDT becomes the cell's content, drop a leading empty auto-seed
            // paragraph so the rebuilt cell matches the source's SDT-only shape.
            // Gated tightly: only an empty paragraph (no runs / no text) that sits
            // before the just-added SDT is removed; a cell that already had real
            // paragraph content keeps it.
            if (parent is TableCell sdtCell)
            {
                var seed = sdtCell.Elements<Paragraph>().FirstOrDefault();
                if (seed != null
                    && sdtCell.Elements().TakeWhile(e => e != sdtBlock).Contains(seed)
                    && !seed.Elements<Run>().Any()
                    && !seed.Elements<Hyperlink>().Any()
                    && !seed.Descendants<Text>().Any())
                {
                    seed.Remove();
                }
            }
            // Root-aware path: the sdtBlock may have been inserted into a
            // header/footer; count SdtBlock siblings under its actual parent
            // and prefix with the correct root segment.
            var blockRoot = ExtractRootSegment(parentPath);
            var blockSiblingCount = parent.Elements<SdtBlock>().TakeWhile(s => s != sdtBlock).Count() + 1;
            resultPath = parent is Body
                ? $"{blockRoot}/sdt[{blockSiblingCount}]"
                : $"{parentPath}/sdt[{blockSiblingCount}]";
        }
        return resultPath;
    }

    // Build a Word checkbox content-control marker (<w14:checkbox>) for the sdtPr.
    // checked flag + the standard MS-Gothic 2612/2610 box glyph states, matching
    // what Word writes for its "Check Box Content Control". The SDK declares the
    // w14 namespace automatically on serialization.
    private static DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox BuildSdtCheckBox(bool isChecked)
    {
        return new DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox
        {
            Checked = new DocumentFormat.OpenXml.Office2010.Word.Checked
            {
                Val = isChecked
                    ? DocumentFormat.OpenXml.Office2010.Word.OnOffValues.One
                    : DocumentFormat.OpenXml.Office2010.Word.OnOffValues.Zero
            },
            CheckedState = new DocumentFormat.OpenXml.Office2010.Word.CheckedState { Val = "2612", Font = "MS Gothic" },
            UncheckedState = new DocumentFormat.OpenXml.Office2010.Word.UncheckedState { Val = "2610", Font = "MS Gothic" }
        };
    }

    // BUG-DUMP-SDTPROPS: apply the form-control sdtPr children that the typed
    // dump→batch path previously dropped — placeholder docPart + showingPlcHdr,
    // date-picker selected value/locale/calendar/store-as, and combo/dropdown
    // current selection (lastValue). Mirrors the Get-side ReadSdtExtraProps so
    // dump→batch round-trips. Respects CT_SdtPr schema order: placeholder /
    // showingPlcHdr sit before the type element (date/comboBox/dropDownList);
    // the date/lastValue attrs land on the already-appended type element.
    private static void ApplySdtExtraProps(SdtProperties sdtProps, Dictionary<string, string> ciProps)
    {
        // The type-content element (date/comboBox/dropDownList/text) was appended
        // last; placeholder + showingPlcHdr must precede it per schema.
        var typeElement = sdtProps.LastChild as OpenXmlElement;
        // BUG-DUMP-R42-7/8: group / picture markers are also type-content
        // elements that placeholder + showingPlcHdr must precede per CT_SdtPr.
        bool typeIsContent = typeElement is SdtContentDate or SdtContentComboBox
            or SdtContentDropDownList or SdtContentText
            or SdtContentGroup or SdtContentPicture
            or DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox;
        OpenXmlElement? insertBefore = typeIsContent ? typeElement : null;

        void InsertSchemaOrdered(OpenXmlElement el)
        {
            if (insertBefore != null) sdtProps.InsertBefore(el, insertBefore);
            else sdtProps.AppendChild(el);
        }

        // Placeholder: docPart reference (<w:placeholder>) and the showingPlcHdr
        // flag are INDEPENDENT in OOXML. A control can declare a placeholder
        // gallery while displaying real content. Emit each only when its own key
        // is present so the corpus shape (placeholderText, no showingPlcHdr)
        // round-trips byte-for-byte.
        var placeholderText = ciProps.GetValueOrDefault("placeholderText", "");
        if (!string.IsNullOrEmpty(placeholderText))
            InsertSchemaOrdered(new SdtPlaceholder
            {
                DocPartReference = new DocPartReference { Val = placeholderText }
            });
        if (IsTruthy(ciProps.GetValueOrDefault("placeholder", "")))
            InsertSchemaOrdered(new ShowingPlaceholder());

        // BUG-DUMP-R25-5: rebuild <w:dataBinding> (customXml store link). In
        // CT_SdtPr dataBinding ranks before the type-content element, so the
        // schema-ordered insert (before the type element) is correct. xpath +
        // storeItemID are the load-bearing attrs; prefixMappings is optional
        // (present only when the xpath uses namespace prefixes).
        var dbXPath = ciProps.GetValueOrDefault("dataBinding.xpath", "");
        var dbStoreId = ciProps.GetValueOrDefault("dataBinding.storeItemID", "");
        if (!string.IsNullOrEmpty(dbXPath) || !string.IsNullOrEmpty(dbStoreId))
        {
            var db = new DataBinding();
            if (!string.IsNullOrEmpty(dbXPath)) db.XPath = dbXPath;
            if (!string.IsNullOrEmpty(dbStoreId)) db.StoreItemId = dbStoreId;
            var dbPrefix = ciProps.GetValueOrDefault("dataBinding.prefixMappings", "");
            if (!string.IsNullOrEmpty(dbPrefix)) db.PrefixMappings = dbPrefix;
            InsertSchemaOrdered(db);
        }

        // Date-picker selected value + locale/calendar/store-as.
        if (typeElement is SdtContentDate date)
        {
            if (ciProps.TryGetValue("date.fullDate", out var fd)
                && DateTime.TryParse(fd, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.AdjustToUniversal | System.Globalization.DateTimeStyles.AssumeUniversal,
                    out var fdVal))
                date.FullDate = fdVal;
            if (ciProps.TryGetValue("date.lid", out var lid) && !string.IsNullOrEmpty(lid))
                date.LanguageId = new LanguageId { Val = lid };
            if (ciProps.TryGetValue("date.storeMappedDataAs", out var sma) && !string.IsNullOrEmpty(sma))
                date.SdtDateMappingType = new SdtDateMappingType { Val = new EnumValue<DateFormatValues>(new DateFormatValues(sma)) };
            if (ciProps.TryGetValue("date.calendar", out var cal) && !string.IsNullOrEmpty(cal))
                date.Calendar = new Calendar { Val = new EnumValue<CalendarValues>(new CalendarValues(cal)) };
        }

        // Combo / dropdown current selection.
        if (typeElement is SdtContentComboBox combo
            && ciProps.TryGetValue("comboBox.lastValue", out var cbLast) && !string.IsNullOrEmpty(cbLast))
            combo.LastValue = cbLast;
        if (typeElement is SdtContentDropDownList ddl
            && ciProps.TryGetValue("dropDown.lastValue", out var ddLast) && !string.IsNullOrEmpty(ddLast))
            ddl.LastValue = ddLast;
    }

    private string AddWatermark(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var wmText = properties.GetValueOrDefault("text", "DRAFT");
        // BUG-R5A(BUG1): route the Add color through the same sanitizer the Set
        // path uses (SanitizeHex → ParseHelpers.SanitizeColorForOoxml) so CSS
        // RRGGBBAA (#FF000040) and bare AARRGGBB inputs are normalized to 6-digit
        // RGB before hitting VML fillcolor. The previous TrimStart('#') passed
        // 8-digit hex straight through, producing a wrong color (#FF000040 →
        // #000040). Named colors (silver, red…) survive SanitizeHex unchanged.
        var wmColor = properties.TryGetValue("color", out var wmcVal)
            ? SanitizeHex(wmcVal) : "silver";
        var wmFont = properties.GetValueOrDefault("font", OfficeDefaultFonts.MinorLatin);
        var wmSize = properties.GetValueOrDefault("size", "1pt");
        if (!wmSize.EndsWith("pt")) wmSize += "pt";
        var wmRotation = properties.GetValueOrDefault("rotation", "315");
        // Normalize into 0-360 as the schema help promises (-45 → 315); a
        // non-numeric value is rejected instead of landing verbatim in the
        // VML style string.
        {
            if (!double.TryParse(wmRotation,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var wmRotDeg))
                throw new ArgumentException(
                    $"Invalid 'rotation' value: '{wmRotation}'. Expected a number in degrees (e.g. 315 or -45).");
            wmRotDeg %= 360;
            if (wmRotDeg < 0) wmRotDeg += 360;
            wmRotation = wmRotDeg.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }
        var wmOpacity = properties.TryGetValue("opacity", out var wmoVal) ? wmoVal : ".5";
        var wmWidth = properties.GetValueOrDefault("width", "415pt");
        var wmHeight = properties.GetValueOrDefault("height", "207.5pt");

        var mainPartWM = _doc.MainDocumentPart!;

        // Remove existing watermarks first
        RemoveWatermarkHeaders();

        // Create 3 headers (default, first, even) — same as POI's createWatermark()
        var headerTypes = new[] {
            HeaderFooterValues.Default,
            HeaderFooterValues.First,
            HeaderFooterValues.Even
        };

        for (int wi = 0; wi < 3; wi++)
        {
            var wmHeaderPart = mainPartWM.AddNewPart<HeaderPart>();
            var wmIdx = wi + 1;

            // Build VML watermark XML (follows POI's getWatermarkParagraph template)
            var vmlXml = $@"<v:shapetype id=""_x0000_t136"" coordsize=""1600,21600"" o:spt=""136"" adj=""10800"" path=""m@7,0l@8,0m@5,21600l@6,21600e"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <v:formulas>
    <v:f eqn=""sum #0 0 10800""/><v:f eqn=""prod #0 2 1""/><v:f eqn=""sum 21600 0 @1""/>
    <v:f eqn=""sum 0 0 @2""/><v:f eqn=""sum 21600 0 @3""/><v:f eqn=""if @0 @3 0""/>
    <v:f eqn=""if @0 21600 @1""/><v:f eqn=""if @0 0 @2""/><v:f eqn=""if @0 @4 21600""/>
    <v:f eqn=""mid @5 @6""/><v:f eqn=""mid @8 @5""/><v:f eqn=""mid @7 @8""/>
    <v:f eqn=""mid @6 @7""/><v:f eqn=""sum @6 0 @5""/>
  </v:formulas>
  <v:path textpathok=""t"" o:connecttype=""custom"" o:connectlocs=""@9,0;@10,10800;@11,21600;@12,10800"" o:connectangles=""270,180,90,0""/>
  <v:textpath on=""t"" fitshape=""t""/>
  <v:handles><v:h position=""#0,bottomRight"" xrange=""6629,14971""/></v:handles>
  <o:lock v:ext=""edit"" text=""t"" shapetype=""t""/>
</v:shapetype>
<v:shape id=""PowerPlusWaterMarkObject{wmIdx}"" o:spid=""_x0000_s102{4 + wmIdx}"" type=""#_x0000_t136"" style=""position:absolute;margin-left:0;margin-top:0;width:{wmWidth};height:{wmHeight};rotation:{wmRotation};z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"" o:allowincell=""f"" fillcolor=""{wmColor}"" stroked=""f"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <v:fill opacity=""{wmOpacity}""/>
  <v:textpath style=""font-family:&quot;{System.Security.SecurityElement.Escape(wmFont)}&quot;;font-size:{wmSize}"" string=""{System.Security.SecurityElement.Escape(wmText)}""/>
</v:shape>";

            // Build header XML with SDT wrapper (docPartGallery=Watermarks)
            var headerXml = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
       xmlns:w10=""urn:schemas-microsoft-com:office:word"">
  <w:sdt>
    <w:sdtPr>
      <w:id w:val=""{-1000 - wmIdx}""/>
      <w:docPartObj>
        <w:docPartGallery w:val=""Watermarks""/>
        <w:docPartUnique/>
      </w:docPartObj>
    </w:sdtPr>
    <w:sdtContent>
      <w:p>
        <w:pPr><w:pStyle w:val=""Header""/></w:pPr>
        <w:r>
          <w:rPr><w:noProof/></w:rPr>
          <w:pict>{vmlXml}</w:pict>
        </w:r>
      </w:p>
    </w:sdtContent>
  </w:sdt>
</w:hdr>";

            using (var stream = wmHeaderPart.GetStream(System.IO.FileMode.Create))
            using (var writer = new System.IO.StreamWriter(stream, System.Text.Encoding.UTF8))
                writer.Write(headerXml);

            // Link header to section properties
            var wmBody = mainPartWM.Document!.Body!;
            var wmSectPr = wmBody.Elements<SectionProperties>().LastOrDefault()
                ?? wmBody.AppendChild(new SectionProperties());

            // Remove existing header reference of same type
            var existingRef = wmSectPr.Elements<HeaderReference>()
                .FirstOrDefault(r => r.Type?.Value == headerTypes[wi]);
            existingRef?.Remove();

            wmSectPr.PrependChild(new HeaderReference
            {
                Id = mainPartWM.GetIdOfPart(wmHeaderPart),
                Type = headerTypes[wi]
            });
        }

        // Enable even/odd page headers and title page
        var wmSettingsPart = mainPartWM.DocumentSettingsPart
            ?? mainPartWM.AddNewPart<DocumentSettingsPart>();
        wmSettingsPart.Settings ??= new Settings();
        if (wmSettingsPart.Settings.GetFirstChild<EvenAndOddHeaders>() == null)
            wmSettingsPart.Settings.AddChild(new EvenAndOddHeaders(), throwOnError: false);
        var wmSectPrForTitle = mainPartWM.Document!.Body!.Elements<SectionProperties>().LastOrDefault()
            ?? mainPartWM.Document!.Body!.AppendChild(new SectionProperties());
        if (wmSectPrForTitle.GetFirstChild<TitlePage>() == null)
            wmSectPrForTitle.AddChild(new TitlePage(), throwOnError: false);

        return "/watermark";
    }

    private string AddDefault(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
        // Generic fallback: create typed element via SDK schema validation
        // TryCreateTypedElement consumes EVERY prop (as an XML attribute or
        // via the SetGenericAttribute fallback) but reads them through plain
        // foreach, which doesn't fire TrackingPropertyDictionary's accessed-key
        // recording — so a raw-element add (w:spacing w:after=...) warned
        // unsupported_property for props that were in fact applied. Probe each
        // key through ContainsKey (which does fire the tracking comparer).
        foreach (var trackedKey in properties.Keys.ToList())
            properties.ContainsKey(trackedKey);
        var created = GenericXmlQuery.TryCreateTypedElement(parent, type, properties, index);
        if (created == null)
            throw new ArgumentException($"Unknown element type '{type}' for {parentPath}. " +
                "Valid types: paragraph (p), run (r), table (tbl), row, cell, picture, chart, ole (object, embed), equation, comment, section, footnote, endnote, toc, style, watermark, bookmark, hyperlink, field, break, sdt, header, footer. " +
                "Use 'officecli docx add' for details.");

        var siblings = parent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
        var createdIdx = PathIndex.FromArrayIndex(siblings.IndexOf(created));
        var resultPath = $"{parentPath}/{created.LocalName}[{createdIdx}]";
        return resultPath;
    }

    /// <summary>
    /// Parse the SDT --prop items= argument into ListItem children.
    /// BUG-R5-07: previously the comma-split tokens were used as both
    /// displayText and value, which is fine for "Draft,Review,Final" but
    /// erases the distinct value attribute that real Word documents use
    /// ("Draft|DRAFT,Review|REVIEW,Final|FINAL"). dump emits this
    /// pipe-separated form when DisplayText differs from Value; accept it
    /// here so add round-trips correctly. A bare token (no `|`) keeps the
    /// old behavior — display == value.
    /// </summary>
    // BUG-DUMP9-09: MERGEFIELD field names with whitespace must be quoted in
    // the instruction so Word parses them as one token. Already-quoted input
    // is left as-is so the instruction is idempotent under dump round-trip.
    // Append the trailing-switches blob produced by WordBatchEmitter for SEQ /
    // MERGEFIELD round-trips (e.g. `\* ARABIC \r 1`, `\* MERGEFORMAT`).
    // Returns either an empty string or a single space + verbatim switches,
    // so the caller can splice it directly between the identifier and the
    // closing space. BUG-DUMP17-01 / BUG-DUMP17-02.
    private static string AppendFieldSwitches(Dictionary<string, string>? properties)
    {
        if (properties == null) return "";
        if (!properties.TryGetValue("switches", out var sw) || string.IsNullOrWhiteSpace(sw)) return "";
        return " " + sw.Trim();
    }

    // BUG-R7A: REF/PAGEREF/NOTEREF accept `\h` either via the legacy
    // `hyperlink` prop (hand-authored adds) or via the `switches` blob
    // (dump→batch round-trips, where the emitter routes ALL residual
    // switches through `switches`). Emit ` \h` from the legacy prop only
    // when the switches blob doesn't already carry one, so a round-tripped
    // ` REF bm1 \h ` doesn't become ` REF bm1 \h \h `.
    private static string RefHyperlinkSwitch(Dictionary<string, string>? properties, string switchesBlob)
    {
        if (!IsTruthy(properties?.GetValueOrDefault("hyperlink"))) return "";
        if (System.Text.RegularExpressions.Regex.IsMatch(switchesBlob, @"\\h\b")) return "";
        return " \\h";
    }

    // R52-bt-1: the IF field's cached run was seeded with trueText
    // unconditionally, so `IF 2 = 1 "T" "F"` displayed T until a real Word
    // re-evaluation. Statically evaluate literal comparisons (numbers or
    // quoted strings, operators = <> < > <= >=); expressions referencing
    // other fields aren't decidable here — seed empty so the `evaluated`
    // protocol reports field_not_evaluated instead of fabricating a result.
    private static string EvaluateIfFieldPlaceholder(Dictionary<string, string> properties)
    {
        var trueText = properties.GetValueOrDefault("trueText", properties.GetValueOrDefault("truetext", ""));
        var falseText = properties.GetValueOrDefault("falseText", properties.GetValueOrDefault("falsetext", ""));
        var expr = properties.GetValueOrDefault("expression", properties.GetValueOrDefault("condition", ""));
        var m = System.Text.RegularExpressions.Regex.Match(expr.Trim(),
            @"^(?:""(?<ls>[^""]*)""|(?<ln>-?\d+(?:\.\d+)?))\s*(?<op><>|<=|>=|=|<|>)\s*(?:""(?<rs>[^""]*)""|(?<rn>-?\d+(?:\.\d+)?))$");
        if (!m.Success) return "";
        bool result;
        var op = m.Groups["op"].Value;
        if (m.Groups["ln"].Success && m.Groups["rn"].Success)
        {
            var l = double.Parse(m.Groups["ln"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var r = double.Parse(m.Groups["rn"].Value, System.Globalization.CultureInfo.InvariantCulture);
            result = op switch { "=" => l == r, "<>" => l != r, "<" => l < r, ">" => l > r, "<=" => l <= r, _ => l >= r };
        }
        else
        {
            var l = m.Groups["ls"].Success ? m.Groups["ls"].Value : m.Groups["ln"].Value;
            var r = m.Groups["rs"].Success ? m.Groups["rs"].Value : m.Groups["rn"].Value;
            var cmp = string.Compare(l, r, StringComparison.OrdinalIgnoreCase);
            result = op switch { "=" => cmp == 0, "<>" => cmp != 0, "<" => cmp < 0, ">" => cmp > 0, "<=" => cmp <= 0, _ => cmp >= 0 };
        }
        return result ? trueText : falseText;
    }

    // R53-fuzz-1: honour the \r restart and \* numbering-format switches in
    // the cached value — they were written into instrText but the cache
    // always showed continuing arabic numerals.
    private static string FormatSeqCachedValue(int number, string switches)
    {
        var rm = System.Text.RegularExpressions.Regex.Match(switches, @"\\r\s+(\d+)");
        if (rm.Success && int.TryParse(rm.Groups[1].Value, out var restartAt))
            number = restartAt;
        var fm = System.Text.RegularExpressions.Regex.Match(switches, @"\\\*\s+(\S+)");
        var fmt = fm.Success ? fm.Groups[1].Value : "";
        // Delegate to SeqEval's formatter (the RecalcSeqFields engine) so the
        // insert-time seed and the recalc pass agree on \* semantics.
        return FormatSeqValue(number, fmt)
            ?? number.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    // Count SEQ fields already carrying this identifier so a newly appended
    // one caches its 1-based position in the sequence.
    private int CountExistingSeqFields(string? identifier)
    {
        if (string.IsNullOrEmpty(identifier)) return 1;
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 1;
        int count = 0;
        foreach (var instr in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldCode>())
        {
            var text = instr.Text ?? "";
            var im = System.Text.RegularExpressions.Regex.Match(text, @"^\s*SEQ\s+(\S+)");
            if (im.Success && im.Groups[1].Value.Equals(identifier, StringComparison.OrdinalIgnoreCase))
                count++;
        }
        return count + 1;
    }

    private static string QuoteFieldNameIfNeeded(string name)
    {
        if (string.IsNullOrEmpty(name)) return name;
        if (name.Length >= 2 && name[0] == '"' && name[^1] == '"') return name;
        bool needs = false;
        foreach (var ch in name)
        {
            if (char.IsWhiteSpace(ch) || ch == '"' || ch == '\\') { needs = true; break; }
        }
        if (!needs) return name;
        var escaped = name.Replace("\\", "\\\\").Replace("\"", "\\\"");
        return $"\"{escaped}\"";
    }

    private static IEnumerable<ListItem> ParseSdtItems(string items)
    {
        foreach (var raw in items.Split(','))
        {
            var trimmed = raw.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;
            string display, value;
            var pipeIdx = trimmed.IndexOf('|');
            if (pipeIdx > 0)
            {
                display = trimmed[..pipeIdx].Trim();
                value = trimmed[(pipeIdx + 1)..].Trim();
            }
            else
            {
                display = value = trimmed;
            }
            yield return new ListItem { DisplayText = display, Value = value };
        }
    }

    // =====================================================================
    // v5.7-cont: add type=textbox / add type=shape
    // =====================================================================

    // Valid <wp:positionV>/<wp:positionH> relativeFrom values (ST_RelFromV /
    // ST_RelFromH). Stored in schema spelling; matched case-insensitively so
    // SanitizeRelativeFrom returns the canonical casing the XML needs.
    private static readonly HashSet<string> VerticalRelativeFroms = new(StringComparer.OrdinalIgnoreCase)
        { "margin", "page", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin" };
    private static readonly HashSet<string> HorizontalRelativeFroms = new(StringComparer.OrdinalIgnoreCase)
        { "margin", "page", "column", "character", "leftMargin", "rightMargin", "insideMargin", "outsideMargin" };

    /// <summary>
    /// Map a caller-supplied anchor reference frame to a valid schema value,
    /// returning the canonical casing. Unknown/empty falls back to the legacy
    /// default so existing callers (who never set it) are unchanged.
    /// </summary>
    private static string SanitizeRelativeFrom(string? value, HashSet<string> valid, string fallback)
    {
        if (string.IsNullOrWhiteSpace(value)) return fallback;
        return valid.TryGetValue(value.Trim(), out var canonical) ? canonical : fallback;
    }

    // Valid <wp:align> values (ST_AlignH / ST_AlignV). A floating drawing
    // positions each axis EITHER by an absolute posOffset OR by one of these
    // relative alignments — used so a center/right-aligned textbox round-trips.
    private static readonly HashSet<string> HorizontalAligns = new(StringComparer.OrdinalIgnoreCase)
        { "left", "right", "center", "inside", "outside" };
    private static readonly HashSet<string> VerticalAligns = new(StringComparer.OrdinalIgnoreCase)
        { "top", "bottom", "center", "inside", "outside" };

    /// <summary>
    /// Map a caller-supplied relative alignment to its canonical lowercase
    /// schema value, or null when absent/invalid (caller then falls back to a
    /// posOffset). Distinct from <see cref="SanitizeRelativeFrom"/> in that an
    /// empty value yields null rather than a default — alignment is optional.
    /// </summary>
    private static string? SanitizeDrawingAlign(string? value, HashSet<string> valid)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        return valid.TryGetValue(value.Trim(), out var canonical)
            ? canonical.ToLowerInvariant()
            : null;
    }

    /// <summary>
    /// Parse a textbox bodyPr inset (EMU). Empty/invalid falls back to the
    /// Word default so callers that omit it keep the standard padding.
    /// Negative values are clamped to 0 (insets cannot be negative).
    /// </summary>
    private static long ParseInsetEmu(string? value, long fallback)
    {
        if (string.IsNullOrWhiteSpace(value)) return fallback;
        return long.TryParse(value.Trim(), out var emu) ? Math.Max(0, emu) : fallback;
    }

    private string AddTextbox(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // R47: `/chart[N]` resolves to the paragraph that hosts the chart,
        // not a chart-internal container. Adding a textbox there silently
        // landed the drawing in body and returned an unresolvable
        // `/chart[N]/textbox[M]` path. Reject up-front.
        if (parentPath.StartsWith("/chart[", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Cannot add a textbox to a chart path ('{parentPath}'). " +
                "Charts don't host textbox children — use /body or a table cell as parent.");
        // BUG-D1-MULTIDRAWING-HOST: a paragraph parent means "attach the
        // textbox drawing to this existing paragraph" instead of creating a
        // fresh host. Used by the dump emitter when N textboxes share one
        // source paragraph (side-by-side card layout) — we want them all
        // anchored on the same paragraph, not stacked across N paragraphs.
        Paragraph? attachToPara = parent as Paragraph;
        // Resolve target container: body is the canonical anchor; cell/header/
        // footer are also legal (they all hold block-flow paragraphs).
        // When attachToPara is set, the indexing host is the paragraph's
        // ancestor body/cell/header/footer — textbox index stays continuous.
        var (host, hostRoot) = attachToPara != null
            ? ResolveDrawingHostFromParagraph(attachToPara, parentPath)
            : ResolveDrawingHost(parent, parentPath);
        long cxEmu = ParseDrawingSize(properties.GetValueOrDefault("width"), defaultEmu: 2_286_000);  // ~6cm
        long cyEmu = ParseDrawingSize(properties.GetValueOrDefault("height"), defaultEmu: 914_400);   // ~2.4cm
        string wrap = properties.GetValueOrDefault("wrap", "square").ToLowerInvariant();
        long hPos = ParseDrawingPos(properties, "anchor.x", "hposition", defaultEmu: 0);
        long vPos = ParseDrawingPos(properties, "anchor.y", "vposition", defaultEmu: 0);
        // Anchor reference frames. Default to the legacy hardcoded column/
        // paragraph so existing callers are unchanged; the dump emitter forwards
        // the source frames (hRelative/vRelative) so a relativeFrom="page"
        // textbox round-trips faithfully instead of floating off-position.
        string hRel = SanitizeRelativeFrom(
            properties.GetValueOrDefault("hRelative") ?? properties.GetValueOrDefault("hrelative"),
            HorizontalRelativeFroms, "column");
        string vRel = SanitizeRelativeFrom(
            properties.GetValueOrDefault("vRelative") ?? properties.GetValueOrDefault("vrelative"),
            VerticalRelativeFroms, "paragraph");
        // Optional relative alignment (<wp:align>). When present it replaces the
        // posOffset on that axis so a center/right-aligned textbox round-trips
        // instead of collapsing to posOffset=0 (hard left/top).
        string? hAlign = SanitizeDrawingAlign(
            properties.GetValueOrDefault("hAlign") ?? properties.GetValueOrDefault("halign"), HorizontalAligns);
        string? vAlign = SanitizeDrawingAlign(
            properties.GetValueOrDefault("vAlign") ?? properties.GetValueOrDefault("valign"), VerticalAligns);
        string posHInner = hAlign != null ? $"<wp:align>{hAlign}</wp:align>" : $"<wp:posOffset>{hPos}</wp:posOffset>";
        string posVInner = vAlign != null ? $"<wp:align>{vAlign}</wp:align>" : $"<wp:posOffset>{vPos}</wp:posOffset>";
        // bodyPr text insets (EMU). Default to Word's standard insets so callers
        // that don't set them are unchanged; the dump emitter forwards the
        // source insets so a zero-inset letterhead box keeps its text width.
        long lIns = ParseInsetEmu(properties.GetValueOrDefault("inset.left"), 91440);
        long tIns = ParseInsetEmu(properties.GetValueOrDefault("inset.top"), 45720);
        long rIns = ParseInsetEmu(properties.GetValueOrDefault("inset.right"), 91440);
        long bIns = ParseInsetEmu(properties.GetValueOrDefault("inset.bottom"), 45720);
        string? fillColor = properties.GetValueOrDefault("fill") ?? properties.GetValueOrDefault("fillcolor");
        string? lineColor = properties.GetValueOrDefault("line.color") ?? properties.GetValueOrDefault("linecolor");
        string? lineStyle = properties.GetValueOrDefault("line.style") ?? properties.GetValueOrDefault("linestyle");
        string? lineWidth = properties.GetValueOrDefault("line.width") ?? properties.GetValueOrDefault("linewidth");
        string? altText   = properties.GetValueOrDefault("alt") ?? properties.GetValueOrDefault("name") ?? "Text Box";
        string? initialText = properties.GetValueOrDefault("text");

        var siblingShapes = host.Elements<Paragraph>()
            .SelectMany(p => p.Descendants<Drawing>())
            .Count();
        uint docPropId = NextDocPropId();
        // Build the textbox via InnerXml. wps:wsp ships in OOXML 2010+; the
        // namespace declarations are the canonical Word ones.
        // Advanced shape attributes (all optional; the dump emitter forwards
        // the source values so a stress-test textbox round-trips faithfully):
        //   geometry      → <a:prstGeom prst="...">      (rect / roundRect / …)
        //   rotation      → <a:xfrm rot="...">           (degrees or raw 60000ths)
        //   textDirection → <wps:bodyPr vert="...">      (horz / eaVert / vert / …)
        //   textAnchor    → <wps:bodyPr anchor="...">    (t / ctr / b)
        //   fill.gradient → <a:gradFill>                  ("c1@pos;c2@pos" or "c1,c2")
        //   fill.opacity  → <a:alpha> inside solidFill    (0-100000 or "80%")
        //   shadow        → <a:effectLst><a:outerShdw>    ("true" or "blur;dist;dir;color;alpha")
        string geom = SanitizeGeometry(
            properties.GetValueOrDefault("geometry") ?? properties.GetValueOrDefault("shape") ?? "rect");
        // Adjust handle for shapes with a single "adj" guide (most notably
        // roundRect's corner radius). Empty when unset → bare <a:avLst/>.
        string avLstXml = BuildAdjXml(
            properties.GetValueOrDefault("cornerRadius") ?? properties.GetValueOrDefault("adj"));
        string rotAttr = BuildRotAttr(
            properties.GetValueOrDefault("rotation") ?? properties.GetValueOrDefault("rot"));
        string vert = properties.GetValueOrDefault("textDirection") ?? properties.GetValueOrDefault("vert") ?? "";
        string vertAttr = !string.IsNullOrEmpty(vert) ? $" vert=\"{SanitizeBodyVert(vert)}\"" : "";
        string anchorVal = SanitizeBodyAnchor(
            properties.GetValueOrDefault("textAnchor") ?? properties.GetValueOrDefault("vAlign"));
        // BUG-DUMP-R25-6: in-shape text-wrap mode (wps:bodyPr/@wrap) and
        // <a:spAutoFit/> control the textbox's own sizing — distinct from the
        // around-shape `wrap` (wp:wrapNone/Square) handled by wrapInnerXml.
        // Three-way contract so every ST_TextWrappingType value round-trips and
        // a source that had NO @wrap stays attribute-less:
        //   key absent          → legacy interactive caller never set it →
        //                          wrap="square" (preserves prior behaviour).
        //   key present & empty  → dump sentinel: the source bodyPr had no
        //                          @wrap → omit the attribute (preserve absence).
        //   key present & value  → pass the value through verbatim (none /
        //                          square / tight / through), no longer
        //                          clobbered to square by a 2-way map.
        string bodyWrapAttr = BuildBodyWrapAttr(properties);
        string spAutoFitXml = IsTruthy(properties.GetValueOrDefault("autoFit", "")) ? "<a:spAutoFit/>" : "";
        string? gradient = properties.GetValueOrDefault("fill.gradient") ?? properties.GetValueOrDefault("gradient");
        string? fillOpacity = properties.GetValueOrDefault("fill.opacity") ?? properties.GetValueOrDefault("opacity");

        string fillXml;
        if (!string.IsNullOrEmpty(gradient))
            fillXml = BuildGradientXml(gradient);
        else if (!string.IsNullOrEmpty(fillColor) && !IsNoFillColor(fillColor))
            fillXml = BuildSolidFillXml(fillColor, fillOpacity);
        else
            // Unset, or the documented `fill=none` / `fill=transparent` sentinel
            // (see `help docx add textbox`) → explicit no-fill. Previously a
            // literal "none" fell through to BuildSolidFillXml and the color
            // parser rejected it, so a borderless/transparent box was impossible.
            fillXml = "<a:noFill/>";
        string lnXml = BuildLineXml(lineStyle, lineWidth, lineColor);
        // effectLst follows fill+ln in CT_ShapeProperties schema order.
        string effectXml = BuildShadowXml(properties.GetValueOrDefault("shadow"));
        string txbxBodyXml = !string.IsNullOrEmpty(initialText)
            ? $"<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t xml:space=\"preserve\">{System.Security.SecurityElement.Escape(initialText)}</w:t></w:r></w:p>"
            : "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>";

        string wrapInnerXml = WrapXmlFragment(wrap);

        // Z-order: behindDoc pushes the box behind body text; relativeHeight
        // (alias zorder) is the stacking order (higher = front). Defaults keep the
        // legacy in-front, auto-incrementing behaviour.
        string behindDocVal = IsTruthy(properties.GetValueOrDefault("behindDoc", "")) ? "1" : "0";
        string? relHeightRaw = properties.GetValueOrDefault("relativeHeight") ?? properties.GetValueOrDefault("zorder");
        string relHeightVal = !string.IsNullOrWhiteSpace(relHeightRaw)
            && long.TryParse(relHeightRaw.Trim(), out var rh) && rh >= 0
            ? rh.ToString()
            : $"251{siblingShapes:D3}";

        // Drawing scaffolding. EffectExtent + DocProperties + a:graphic with
        // a:graphicData uri = wordprocessingShape; inner wps:wsp carries
        // spPr (preset rect geometry + fill + line) + txbx (body paragraphs) + bodyPr.
        string drawingXml = $@"<w:drawing xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""><wp:anchor distT=""0"" distB=""0"" distL=""114300"" distR=""114300"" simplePos=""0"" relativeHeight=""{relHeightVal}"" behindDoc=""{behindDocVal}"" locked=""0"" layoutInCell=""1"" allowOverlap=""1""><wp:simplePos x=""0"" y=""0""/><wp:positionH relativeFrom=""{hRel}"">{posHInner}</wp:positionH><wp:positionV relativeFrom=""{vRel}"">{posVInner}</wp:positionV><wp:extent cx=""{cxEmu}"" cy=""{cyEmu}""/><wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>{wrapInnerXml}<wp:docPr id=""{docPropId}"" name=""{System.Security.SecurityElement.Escape(altText)}""/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""><wps:wsp><wps:cNvSpPr txBox=""1""/><wps:spPr><a:xfrm{rotAttr}><a:off x=""0"" y=""0""/><a:ext cx=""{cxEmu}"" cy=""{cyEmu}""/></a:xfrm><a:prstGeom prst=""{geom}""><a:avLst>{avLstXml}</a:avLst></a:prstGeom>{fillXml}{lnXml}{effectXml}</wps:spPr><wps:txbx><w:txbxContent>{txbxBodyXml}</w:txbxContent></wps:txbx><wps:bodyPr rot=""0""{vertAttr}{bodyWrapAttr} lIns=""{lIns}"" tIns=""{tIns}"" rIns=""{rIns}"" bIns=""{bIns}"" anchor=""{anchorVal}"" anchorCtr=""0"">{spAutoFitXml}</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>";

        var drawing = ParseDrawingFromXml(drawingXml);
        var run = new Run(drawing);
        Paragraph anchorPara;
        if (attachToPara != null)
        {
            attachToPara.AppendChild(run);
            anchorPara = attachToPara;
        }
        else
        {
            anchorPara = new Paragraph(run);
            AssignParaId(anchorPara);
            InsertAtIndexOrAppend(host, anchorPara, index);
        }

        // Compute the 1-based textbox index across the host. Walk all
        // paragraphs in the host and count those that carry at least one
        // wp:anchor with wsp content — same selector as Get.
        int txbxIdx = CountTextboxesInHost(host, anchorPara);
        return $"{hostRoot}/textbox[{txbxIdx}]";
    }

    private static (OpenXmlElement host, string hostRoot) ResolveDrawingHostFromParagraph(
        Paragraph para, string parentPath)
    {
        // CONSISTENCY(d1-multi-drawing): a paragraph parent (e.g.
        // /body/p[last()]) means "attach to this existing paragraph". Walk
        // up to the nearest Body / TableCell / Header / Footer ancestor —
        // that's the textbox-index host. hostRoot is parentPath stripped of
        // the trailing "/p[..]" segment so /<hostRoot>/textbox[N] keeps its
        // continuous numbering across the entire body/cell/header/footer.
        var idx = parentPath.LastIndexOf("/p[", StringComparison.Ordinal);
        var hostRoot = idx >= 0 ? parentPath.Substring(0, idx) : parentPath;
        if (string.IsNullOrEmpty(hostRoot)) hostRoot = "/";
        // Reject nesting: a textbox/shape inside an existing textbox's
        // txbxContent produces OOXML the spec prohibits — Word refuses to open
        // the file (0x800706BE). Fail fast before any XML mutation occurs.
        for (var walk = para.Parent; walk != null; walk = walk.Parent)
        {
            if (walk.LocalName == "txbxContent")
                throw new ArgumentException(
                    $"Cannot add textbox/shape under {parentPath}: nested textboxes are not permitted by the OOXML spec (w:drawing inside w:txbxContent corrupts the file).");
        }
        OpenXmlElement? anc = para.Parent;
        while (anc != null && anc is not (Body or TableCell or Header or Footer))
            anc = anc.Parent;
        if (anc == null)
            throw new ArgumentException(
                $"Cannot attach textbox to {parentPath}: no body/cell/header/footer ancestor.");
        return (anc, hostRoot);
    }

    private string AddShape(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var (host, hostRoot) = ResolveDrawingHost(parent, parentPath);
        string preset = properties.GetValueOrDefault("geometry")
                     ?? properties.GetValueOrDefault("preset")
                     ?? "rect";
        long cxEmu = ParseDrawingSize(properties.GetValueOrDefault("width"), defaultEmu: 914_400);
        long cyEmu = ParseDrawingSize(properties.GetValueOrDefault("height"), defaultEmu: 914_400);
        string wrap = properties.GetValueOrDefault("wrap", "none").ToLowerInvariant();
        long hPos = ParseDrawingPos(properties, "anchor.x", "hposition", defaultEmu: 0);
        long vPos = ParseDrawingPos(properties, "anchor.y", "vposition", defaultEmu: 0);
        string hRel = SanitizeRelativeFrom(
            properties.GetValueOrDefault("hRelative") ?? properties.GetValueOrDefault("hrelative"),
            HorizontalRelativeFroms, "column");
        string vRel = SanitizeRelativeFrom(
            properties.GetValueOrDefault("vRelative") ?? properties.GetValueOrDefault("vrelative"),
            VerticalRelativeFroms, "paragraph");
        // Optional relative alignment (<wp:align>) — replaces posOffset on that
        // axis so a center/right-aligned shape round-trips. See AddTextbox.
        string? hAlign = SanitizeDrawingAlign(
            properties.GetValueOrDefault("hAlign") ?? properties.GetValueOrDefault("halign"), HorizontalAligns);
        string? vAlign = SanitizeDrawingAlign(
            properties.GetValueOrDefault("vAlign") ?? properties.GetValueOrDefault("valign"), VerticalAligns);
        string posHInner = hAlign != null ? $"<wp:align>{hAlign}</wp:align>" : $"<wp:posOffset>{hPos}</wp:posOffset>";
        string posVInner = vAlign != null ? $"<wp:align>{vAlign}</wp:align>" : $"<wp:posOffset>{vPos}</wp:posOffset>";
        // fill: bare color, or "none"; "line=STYLE;SIZE;COLOR" composite.
        string? fillRaw = properties.GetValueOrDefault("fill");
        string fillXml;
        if (string.IsNullOrEmpty(fillRaw) || string.Equals(fillRaw, "none", StringComparison.OrdinalIgnoreCase))
            fillXml = "<a:noFill/>";
        else
            fillXml = $"<a:solidFill><a:srgbClr val=\"{SanitizeHex(fillRaw)}\"/></a:solidFill>";
        // line: either "line=STYLE;SIZE;COLOR" or split keys.
        string? lineCompact = properties.GetValueOrDefault("line");
        string? lineStyle = null, lineWidth = null, lineColor = null;
        if (!string.IsNullOrEmpty(lineCompact)
            && !string.Equals(lineCompact, "none", StringComparison.OrdinalIgnoreCase))
        {
            var parts = lineCompact.Split(';');
            if (parts.Length >= 1 && !string.IsNullOrEmpty(parts[0])) lineStyle = parts[0];
            if (parts.Length >= 2 && !string.IsNullOrEmpty(parts[1])) lineWidth = parts[1];
            if (parts.Length >= 3 && !string.IsNullOrEmpty(parts[2])) lineColor = parts[2];
        }
        lineStyle ??= properties.GetValueOrDefault("line.style") ?? properties.GetValueOrDefault("linestyle");
        lineWidth ??= properties.GetValueOrDefault("line.width") ?? properties.GetValueOrDefault("linewidth");
        lineColor ??= properties.GetValueOrDefault("line.color") ?? properties.GetValueOrDefault("linecolor");
        string lnXml = BuildLineXml(lineStyle, lineWidth, lineColor);
        string altText = properties.GetValueOrDefault("alt") ?? properties.GetValueOrDefault("name") ?? "Shape";

        var siblingShapes = host.Elements<Paragraph>()
            .SelectMany(p => p.Descendants<Drawing>())
            .Count();
        uint docPropId = NextDocPropId();
        string wrapInnerXml = WrapXmlFragment(wrap);

        string drawingXml = $@"<w:drawing xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""><wp:anchor distT=""0"" distB=""0"" distL=""114300"" distR=""114300"" simplePos=""0"" relativeHeight=""251{siblingShapes:D3}"" behindDoc=""0"" locked=""0"" layoutInCell=""1"" allowOverlap=""1""><wp:simplePos x=""0"" y=""0""/><wp:positionH relativeFrom=""{hRel}"">{posHInner}</wp:positionH><wp:positionV relativeFrom=""{vRel}"">{posVInner}</wp:positionV><wp:extent cx=""{cxEmu}"" cy=""{cyEmu}""/><wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>{wrapInnerXml}<wp:docPr id=""{docPropId}"" name=""{System.Security.SecurityElement.Escape(altText)}""/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{cxEmu}"" cy=""{cyEmu}""/></a:xfrm><a:prstGeom prst=""{SanitizeGeometry(preset)}""><a:avLst/></a:prstGeom>{fillXml}{lnXml}</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>";

        var drawing = ParseDrawingFromXml(drawingXml);
        var run = new Run(drawing);
        var newPara = new Paragraph(run);
        AssignParaId(newPara);
        InsertAtIndexOrAppend(host, newPara, index);

        int shapeIdx = CountShapesInHost(host, newPara);
        return $"{hostRoot}/shape[{shapeIdx}]";
    }

    // ----- helpers shared by AddTextbox / AddShape -----------------------

    private static (OpenXmlElement host, string hostRoot) ResolveDrawingHost(OpenXmlElement parent, string parentPath)
    {
        // Accept body / cell / header / footer roots. Path's first segment
        // ("/body", "/header[N]", "/footer[N]", or "/body/.../tc[N]") is what
        // we re-use for the returned /<root>/textbox[N] path.
        // Reject nesting under a txbxContent ancestor (e.g. a cell inside a
        // textbox-nested table): drawings under txbxContent corrupt the file
        // (Word 0x800706BE). Mirror ResolveDrawingHostFromParagraph.
        for (var walk = parent; walk != null; walk = walk.Parent)
        {
            if (walk.LocalName == "txbxContent")
                throw new ArgumentException(
                    $"Cannot add textbox/shape under {parentPath}: nested textboxes are not permitted by the OOXML spec (w:drawing inside w:txbxContent corrupts the file).");
        }
        if (parent is Body) return (parent, parentPath.TrimEnd('/'));
        if (parent is TableCell) return (parent, parentPath);
        // OpenXmlPartRootElement (Header/Footer): use itself.
        if (parent is Header || parent is Footer) return (parent, parentPath);
        throw new ArgumentException($"Cannot add textbox/shape under {parentPath}: only /body, /body/tbl/tr/tc[N], /header[N], /footer[N] are supported.");
    }

    private static long ParseDrawingSize(string? raw, long defaultEmu)
    {
        if (string.IsNullOrWhiteSpace(raw)) return defaultEmu;
        try { return ParseEmu(raw); }
        catch { return defaultEmu; }
    }

    private static long ParseDrawingPos(Dictionary<string,string> props, string camelKey, string altKey, long defaultEmu)
    {
        if (props.TryGetValue(camelKey, out var v) && !string.IsNullOrWhiteSpace(v))
        { try { return ParseEmu(v); } catch { } }
        if (props.TryGetValue(altKey, out var v2) && !string.IsNullOrWhiteSpace(v2))
        { try { return ParseEmu(v2); } catch { } }
        return defaultEmu;
    }

    /// <summary>v5.7-cont: convert wrap token to its wp:wrap* fragment.</summary>
    private static string WrapXmlFragment(string wrap) => wrap.ToLowerInvariant() switch
    {
        "square"      => "<wp:wrapSquare wrapText=\"bothSides\"/>",
        "tight"       => "<wp:wrapTight wrapText=\"bothSides\"><wp:wrapPolygon edited=\"0\"><wp:start x=\"0\" y=\"0\"/><wp:lineTo x=\"21600\" y=\"0\"/><wp:lineTo x=\"21600\" y=\"21600\"/><wp:lineTo x=\"0\" y=\"21600\"/><wp:lineTo x=\"0\" y=\"0\"/></wp:wrapPolygon></wp:wrapTight>",
        "topbottom" or "topandbottom" => "<wp:wrapTopAndBottom/>",
        "behind"      => "<wp:wrapNone/>",
        "infront"     => "<wp:wrapNone/>",
        "none" or ""  => "<wp:wrapNone/>",
        _             => "<wp:wrapSquare wrapText=\"bothSides\"/>",
    };

    /// <summary>Build the <c>a:ln</c> child for spPr. Returns the empty
    /// string when none of style/width/color was specified — Word then
    /// uses the theme default.</summary>
    private static string BuildLineXml(string? style, string? width, string? color)
    {
        if (string.IsNullOrEmpty(style) && string.IsNullOrEmpty(width) && string.IsNullOrEmpty(color))
            return "";
        // a:ln@w is in EMU (1pt = 12700 EMU). Accept bare integer pt or "Npt"/"Ncm".
        long lnWidthEmu = 0;
        if (!string.IsNullOrEmpty(width))
        {
            try { lnWidthEmu = ParseEmu(width); } catch { lnWidthEmu = 0; }
            if (lnWidthEmu == 0 && double.TryParse(width, out var pts)) lnWidthEmu = (long)Math.Round(pts * EmuConverter.EmuPerPoint);
        }
        string widthAttr = lnWidthEmu > 0 ? $" w=\"{lnWidthEmu}\"" : "";
        // Style "none" — or the color sentinels "none"/"transparent" — emit
        // a:noFill (a borderless box), anything else emits a:solidFill + optional
        // a:prstDash for non-solid line types. Accepting the color sentinel keeps
        // `line.color=none` symmetric with `fill=none`.
        bool isNone = string.Equals(style, "none", StringComparison.OrdinalIgnoreCase)
            || IsNoFillColor(color);
        if (isNone) return $"<a:ln{widthAttr}><a:noFill/></a:ln>";
        string fill = !string.IsNullOrEmpty(color)
            ? $"<a:solidFill><a:srgbClr val=\"{SanitizeHex(color)}\"/></a:solidFill>"
            : "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>";
        string dash = "";
        if (!string.IsNullOrEmpty(style)
            && !string.Equals(style, "solid", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(style, "single", StringComparison.OrdinalIgnoreCase))
        {
            dash = $"<a:prstDash val=\"{MapDashStyle(style)}\"/>";
        }
        return $"<a:ln{widthAttr}>{fill}{dash}</a:ln>";
    }

    private static string MapDashStyle(string style) => style.ToLowerInvariant() switch
    {
        "dot" or "dotted"             => "dot",
        "dash" or "dashed"            => "dash",
        "dashdot" or "dotdash"        => "dashDot",
        "lgdash" or "longdash"        => "lgDash",
        "sysdash"                     => "sysDash",
        "sysdot"                      => "sysDot",
        _                             => "solid",
    };

    /// <summary>Build the <c>&lt;a:avLst&gt;</c> inner XML for a shape's single
    /// adjust handle (roundRect corner radius, etc.). Accepts a percentage 0-100
    /// (the friendly form; ×1000 into the OOXML guide value where 100000 = 100%)
    /// or a raw guide value > 100. Returns "" when unset → callers emit an empty
    /// <c>&lt;a:avLst/&gt;</c> and the shape keeps its preset default.</summary>
    private static string BuildAdjXml(string? adj)
    {
        if (string.IsNullOrWhiteSpace(adj)) return "";
        var v = adj.Trim().TrimEnd('%');
        if (!double.TryParse(v, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var n)) return "";
        long val = n <= 100 ? (long)Math.Round(n * 1000) : (long)Math.Round(n);
        val = Math.Clamp(val, 0, 100000);
        return $"<a:gd name=\"adj\" fmla=\"val {val}\"/>";
    }

    /// <summary>Build the <c>rot</c> attribute (with leading space) for
    /// <c>&lt;a:xfrm&gt;</c>. Accepts raw OOXML 60000ths-of-a-degree (the dump
    /// form, e.g. 2700000) or plain degrees (e.g. 45). A value ≤ 360 is read as
    /// degrees; anything larger is the raw unit. Returns "" when unset.</summary>
    private static string BuildRotAttr(string? rotation)
    {
        if (string.IsNullOrWhiteSpace(rotation)) return "";
        var v = rotation.Trim().TrimEnd('°'); // tolerate a trailing degree sign
        if (!double.TryParse(v, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var num)) return "";
        long units = Math.Abs(num) <= 360 ? (long)Math.Round(num * 60000) : (long)Math.Round(num);
        units %= 21_600_000; if (units < 0) units += 21_600_000;
        return units == 0 ? "" : $" rot=\"{units}\"";
    }

    /// <summary>bodyPr text-flow direction. Pass-through of OOXML ST_TextVerticalType
    /// values (the dump form); a couple of friendly aliases map in too.</summary>
    private static string SanitizeBodyVert(string vert) => vert.Trim().ToLowerInvariant() switch
    {
        "horz" or "horizontal"        => "horz",
        "vert" or "vertical"          => "vert",
        "vert270"                     => "vert270",
        "wordartvert"                 => "wordArtVert",
        "eavert" or "eastasianvert"   => "eaVert",
        "mongolianvert"               => "mongolianVert",
        "wordartvertrtl"              => "wordArtVertRtl",
        _                             => "horz",
    };

    /// <summary>bodyPr vertical text anchor (t / ctr / b). Defaults to top.</summary>
    private static string SanitizeBodyAnchor(string? anchor) => (anchor ?? "").Trim().ToLowerInvariant() switch
    {
        "ctr" or "center" or "middle" => "ctr",
        "b" or "bottom"               => "b",
        "just" or "justified"          => "just",
        "dist" or "distributed"        => "dist",
        _                             => "t",
    };

    // BUG-DUMP-R25-6: build the wps:bodyPr/@wrap attribute (with a leading
    // space) honouring the three-way contract documented at the call site.
    // ST_TextWrappingType is the closed set {none, square}; Word also writes
    // the legacy values {tight, through} on some shapes. Pass ANY of those
    // through verbatim (the earlier 2-way none/square map clobbered tight/
    // through → square). Absent key → legacy default wrap="square". Present
    // but empty (dump sentinel for a source bodyPr with no @wrap) → emit no
    // attribute at all, preserving the source's attribute-less shape.
    private static string BuildBodyWrapAttr(Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("bodyWrap", out var raw))
            return " wrap=\"square\"";                 // key absent → legacy default
        var v = (raw ?? "").Trim().ToLowerInvariant();
        if (v.Length == 0)
            return "";                                  // explicit-absence sentinel → omit
        var sane = v switch
        {
            "none" or "square" or "tight" or "through" => v,
            _                                          => "square",  // unknown → safe default
        };
        return $" wrap=\"{sane}\"";
    }

    /// <summary><c>none</c>/<c>transparent</c> are the documented sentinels for
    /// an explicit no-fill / no-line (<c>a:noFill</c>), distinct from an unset
    /// value (which leaves the theme default in place).</summary>
    private static bool IsNoFillColor(string? color) =>
        string.Equals(color, "none", StringComparison.OrdinalIgnoreCase)
        || string.Equals(color, "transparent", StringComparison.OrdinalIgnoreCase);

    /// <summary>solidFill, optionally translucent. Opacity accepts 0-100000
    /// (OOXML alpha, the dump form), a "NN%" string, or a 0-1 / 0-100 number.</summary>
    private static string BuildSolidFillXml(string color, string? opacity)
    {
        var alpha = ParseAlpha(opacity);
        var clr = alpha == null
            ? $"<a:srgbClr val=\"{SanitizeHex(color)}\"/>"
            : $"<a:srgbClr val=\"{SanitizeHex(color)}\"><a:alpha val=\"{alpha}\"/></a:srgbClr>";
        return $"<a:solidFill>{clr}</a:solidFill>";
    }

    // Normalize an opacity input to OOXML alpha (0-100000), or null when unset.
    private static int? ParseAlpha(string? opacity)
    {
        if (string.IsNullOrWhiteSpace(opacity)) return null;
        var s = opacity.Trim();
        bool pct = s.EndsWith("%");
        if (pct) s = s[..^1].Trim();
        if (!double.TryParse(s, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var n)) return null;
        // Heuristic: "80%"/"0.8" → 80000; "80" → 80000; raw "80000" → 80000.
        double alpha = pct ? n * 1000
            : n <= 1 ? n * 100000
            : n <= 100 ? n * 1000
            : n;
        return Math.Clamp((int)Math.Round(alpha), 0, 100000);
    }

    /// <summary>Build <c>&lt;a:gradFill&gt;</c> from a gradient spec. Two forms are
    /// accepted so a spec is portable between charts and textboxes:
    /// <list type="bullet">
    /// <item>stop list — <c>color[@pos]</c> separated by <c>;</c>/<c>,</c>, positions
    /// (0-100000) spread evenly when omitted. e.g. "FF6B6B@0;FFE66D@100000".</item>
    /// <item>chart form — <c>C1-C2[-C3…][:angleDeg]</c>, the syntax cChart
    /// <c>chartFill</c> uses. e.g. "FF6B6B-FFE66D:90". The angle emits <c>a:lin</c>.</item>
    /// </list></summary>
    private static string BuildGradientXml(string gradient)
    {
        gradient = gradient.Trim();
        // Chart-style "C1-C2[-…][:angle]": distinguished by the '-' color separator
        // and the absence of any stop-list punctuation (','/';'/'@'). Hex/named
        // colors never contain '-', so this is unambiguous.
        if (gradient.IndexOfAny(new[] { ',', ';', '@' }) < 0 && gradient.Contains('-'))
        {
            string linXml = "";
            var colonIdx = gradient.IndexOf(':');
            if (colonIdx >= 0)
            {
                var angleStr = gradient[(colonIdx + 1)..].Trim();
                gradient = gradient[..colonIdx];
                if (double.TryParse(angleStr, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var deg))
                {
                    long ang = ((long)Math.Round(deg * 60000) % 21_600_000 + 21_600_000) % 21_600_000;
                    linXml = $"<a:lin ang=\"{ang}\" scaled=\"1\"/>";
                }
            }
            var cols = gradient.Split('-', StringSplitOptions.RemoveEmptyEntries);
            if (cols.Length == 0) return "<a:noFill/>";
            var csb = new System.Text.StringBuilder("<a:gradFill><a:gsLst>");
            for (int i = 0; i < cols.Length; i++)
            {
                int pos = cols.Length == 1 ? 0 : (int)Math.Round(i * 100000.0 / (cols.Length - 1));
                csb.Append($"<a:gs pos=\"{pos}\"><a:srgbClr val=\"{SanitizeHex(cols[i].Trim())}\"/></a:gs>");
            }
            csb.Append("</a:gsLst>").Append(linXml).Append("</a:gradFill>");
            return csb.ToString();
        }

        var stops = gradient.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (stops.Length == 0) return "<a:noFill/>";
        var sb = new System.Text.StringBuilder("<a:gradFill><a:gsLst>");
        for (int i = 0; i < stops.Length; i++)
        {
            var part = stops[i].Trim();
            string color; int pos;
            var at = part.IndexOf('@');
            if (at >= 0)
            {
                color = part[..at].Trim();
                if (!int.TryParse(part[(at + 1)..].Trim(), out pos))
                    pos = stops.Length == 1 ? 0 : (int)Math.Round(i * 100000.0 / (stops.Length - 1));
            }
            else
            {
                color = part;
                pos = stops.Length == 1 ? 0 : (int)Math.Round(i * 100000.0 / (stops.Length - 1));
            }
            sb.Append($"<a:gs pos=\"{Math.Clamp(pos, 0, 100000)}\"><a:srgbClr val=\"{SanitizeHex(color)}\"/></a:gs>");
        }
        sb.Append("</a:gsLst></a:gradFill>");
        return sb.ToString();
    }

    /// <summary>Build <c>&lt;a:effectLst&gt;</c> with an outer shadow. Accepts
    /// "true" (a standard offset drop shadow) or a compact
    /// "blurRad;dist;dir;color;alpha" tuple (the dump form). Returns "" when
    /// unset / "false".</summary>
    private static string BuildShadowXml(string? shadow)
    {
        if (string.IsNullOrWhiteSpace(shadow)) return "";
        var s = shadow.Trim();
        if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)
            || string.Equals(s, "none", StringComparison.OrdinalIgnoreCase)) return "";
        // Defaults match Word's standard preset drop shadow.
        long blur = 50800, dist = 38100, dir = 5400000; string color = "000000"; int alpha = 40000;
        if (!string.Equals(s, "true", StringComparison.OrdinalIgnoreCase))
        {
            var p = s.Split(';');
            if (p.Length > 0 && long.TryParse(p[0], out var b)) blur = b;
            if (p.Length > 1 && long.TryParse(p[1], out var d)) dist = d;
            if (p.Length > 2 && long.TryParse(p[2], out var dr)) dir = dr;
            if (p.Length > 3 && !string.IsNullOrWhiteSpace(p[3])) color = SanitizeHex(p[3]);
            if (p.Length > 4 && int.TryParse(p[4], out var a)) alpha = Math.Clamp(a, 0, 100000);
        }
        return $"<a:effectLst><a:outerShdw blurRad=\"{blur}\" dist=\"{dist}\" dir=\"{dir}\" algn=\"t\" rotWithShape=\"0\"><a:srgbClr val=\"{color}\"><a:alpha val=\"{alpha}\"/></a:srgbClr></a:outerShdw></a:effectLst>";
    }

    /// <summary>Whitelist of common preset geometry names. Anything else
    /// falls back to rect rather than emitting schema-invalid XML.</summary>
    private static string SanitizeGeometry(string preset) => preset.ToLowerInvariant() switch
    {
        "rect" or "rectangle"   => "rect",
        "ellipse" or "circle"    => "ellipse",
        "line" or "straightline" => "line",
        "roundrect"              => "roundRect",
        "triangle"               => "triangle",
        "diamond"                => "diamond",
        "pentagon"               => "pentagon",
        "hexagon"                => "hexagon",
        "octagon"                => "octagon",
        "rightarrow"             => "rightArrow",
        "leftarrow"              => "leftArrow",
        "uparrow"                => "upArrow",
        "downarrow"              => "downArrow",
        "star5"                  => "star5",
        "wedgerectcallout"       => "wedgeRectCallout",
        _ => throw new ArgumentException($"Unknown geometry '{preset}'. Valid: rect, ellipse, line, roundRect, triangle, diamond, pentagon, hexagon, octagon, rightArrow, leftArrow, upArrow, downArrow, star5, wedgeRectCallout."),
    };

    /// <summary>Parse a w:drawing element from XML with full namespace
    /// declarations and return the typed <see cref="Drawing"/>. The naive
    /// <c>new Drawing { InnerXml = ... }</c> path drops the outer
    /// namespace context the inner elements need (wp:, a:, wps: prefixes
    /// land as undeclared), so we route through an XmlReader to keep the
    /// nsmgr alive for the parse.</summary>
    private static Drawing ParseDrawingFromXml(string xml)
    {
        // Wrap inside w:p > w:r > drawing so the outer namespace context
        // (declared on the root) is visible to inner wp:/a:/wps: prefixes.
        // <w:drawing> belongs inside <w:r>, so the Run wrapper is the
        // minimal schema-legal host.
        var wrapXml = $@"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:r>{xml}</w:r></w:p>";
        var p = new Paragraph(wrapXml);
        var d = p.Descendants<Drawing>().FirstOrDefault();
        if (d == null)
            throw new InvalidOperationException("Drawing parse failed");
        d.Remove();
        return d;
    }

    private static int CountTextboxesInHost(OpenXmlElement host, Paragraph anchor)
    {
        int count = 0;
        foreach (var p in host.Elements<Paragraph>())
        {
            // A textbox is recognized by a wp:anchor containing a wps:wsp
            // that has a txBox=1 cNvSpPr OR a wps:txbx child.
            bool isTextbox = p.Descendants<Drawing>().Any(d =>
                d.InnerXml.Contains("txBox=\"1\"")
                || d.InnerXml.Contains("<wps:txbx"));
            if (isTextbox) count++;
            if (ReferenceEquals(p, anchor)) return count;
        }
        return count;
    }

    private static int CountShapesInHost(OpenXmlElement host, Paragraph anchor)
    {
        // Stay in lockstep with the Navigation "shape" resolver, which
        // excludes textbox-bearing Drawings (a textbox is a <wps:wsp>
        // wrapping a <wps:txbx>, so the unfiltered `<wps:wsp` test counts
        // textboxes as shapes and the Add-side index drifts ahead of the
        // Get-side index by one per textbox.
        int count = 0;
        foreach (var p in host.Elements<Paragraph>())
        {
            bool isShape = p.Descendants<Drawing>().Any(d =>
            {
                var xml = d.InnerXml;
                if (!xml.Contains("<wps:wsp")) return false;
                if (xml.Contains("<wps:txbx") || xml.Contains("txBox=\"1\"")) return false;
                return true;
            });
            if (isShape) count++;
            if (ReferenceEquals(p, anchor)) return count;
        }
        return count;
    }

    // 1-based index of the anchor paragraph's wpg:wgp group among all groups in
    // the host — matches the Navigation "group" resolver (a `diagram` emits one
    // group). Kept separate from CountShapesInHost: a group carries child
    // textboxes, so the shape/textbox heuristics would misclassify it.
    private static int CountGroupsInHost(OpenXmlElement host, Paragraph anchor)
    {
        int count = 0;
        foreach (var p in host.Elements<Paragraph>())
        {
            if (p.Descendants<Drawing>().Any(d => d.InnerXml.Contains("<wpg:wgp")))
                count++;
            if (ReferenceEquals(p, anchor)) return count;
        }
        return count;
    }
}
