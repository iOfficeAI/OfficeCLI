// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// Modern p188 (Office 2018/8) threaded slide comments — distinct OOXML element
// from the legacy p:cm. Top-level threads live in a per-slide
// PowerPointCommentPart (/ppt/comments/modernComment*.xml); authors live in a
// presentation-level PowerPointAuthorsPart (/ppt/authors.xml). Replies are
// nested p188:CommentReply children of the top-level p188:Comment via a
// p188:replyLst container.
//
// Path syntax (lowercased by NormalizePptxPathSegmentCasing):
//   /slide[N]/moderncomment[K]                       — top-level (1-based)
//   /slide[N]/moderncomment[K]/reply[R]              — nested reply (1-based)
//
// Properties: author, initials, text, resolved (top-level only), created
// (auto-set on Add), parent (Add only — empty/missing = top-level).
public partial class PowerPointHandler
{
    private const string ModernCommentNs = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";

    // ----- Authors part (presentation-level) -----

    /// <summary>
    /// Resolve or create the presentation-level PowerPointAuthorsPart and
    /// return the Author with the requested name+initials, creating one if
    /// it doesn't yet exist. Author ids are GUIDs (curly-brace form).
    /// Mirrors legacy GetOrCreateCommentAuthor; the two author lists are
    /// kept separate (modern uses GUID ids, legacy uses uint ids).
    /// </summary>
    private Author GetOrCreateModernCommentAuthor(string name, string initials)
    {
        var pres = _doc.PresentationPart!;
        var authorsPart = pres.GetPartsOfType<PowerPointAuthorsPart>().FirstOrDefault();
        if (authorsPart == null)
        {
            authorsPart = pres.AddNewPart<PowerPointAuthorsPart>();
            authorsPart.AuthorList = new AuthorList();
        }
        authorsPart.AuthorList ??= new AuthorList();

        var existing = authorsPart.AuthorList.Elements<Author>()
            .FirstOrDefault(a => string.Equals(a.Name?.Value, name, StringComparison.Ordinal)
                              && string.Equals(a.Initials?.Value, initials, StringComparison.Ordinal));
        if (existing != null) return existing;

        var author = new Author
        {
            Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}",
            Name = name,
            Initials = initials,
        };
        authorsPart.AuthorList.AppendChild(author);
        authorsPart.AuthorList.Save();
        return author;
    }

    /// <summary>Look up an author by its GUID id.</summary>
    private Author? FindModernCommentAuthor(string? authorId)
    {
        if (string.IsNullOrEmpty(authorId)) return null;
        var authorsPart = _doc.PresentationPart?.GetPartsOfType<PowerPointAuthorsPart>().FirstOrDefault();
        return authorsPart?.AuthorList?.Elements<Author>()
            .FirstOrDefault(a => string.Equals(a.Id?.Value, authorId, StringComparison.Ordinal));
    }

    /// <summary>
    /// Count references to the given authorId across the whole deck's modern
    /// comment threads (top-level + replies), excluding the supplied element.
    /// Used by author/initials Set to decide between in-place rename vs fork.
    /// </summary>
    private int CountModernAuthorReferences(string authorId, OpenXmlElement exclude)
    {
        int count = 0;
        foreach (var sp in GetSlideParts())
        {
            foreach (var cp in sp.GetPartsOfType<PowerPointCommentPart>())
            {
                if (cp.CommentList == null) continue;
                foreach (var top in cp.CommentList.Elements<Comment>())
                {
                    if (!ReferenceEquals(top, exclude) && string.Equals(top.AuthorId?.Value, authorId, StringComparison.Ordinal))
                        count++;
                    var replyLst = top.Elements<CommentReplyList>().FirstOrDefault();
                    if (replyLst == null) continue;
                    foreach (var r in replyLst.Elements<CommentReply>())
                    {
                        if (!ReferenceEquals(r, exclude) && string.Equals(r.AuthorId?.Value, authorId, StringComparison.Ordinal))
                            count++;
                    }
                }
            }
        }
        return count;
    }

    // ----- Comment part (per-slide) -----

    private PowerPointCommentPart GetOrCreateModernCommentPart(SlidePart slidePart)
    {
        var existing = slidePart.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault();
        if (existing != null)
        {
            existing.CommentList ??= new CommentList();
            return existing;
        }
        var part = slidePart.AddNewPart<PowerPointCommentPart>();
        part.CommentList = new CommentList();
        return part;
    }

    /// <summary>
    /// Enumerate top-level modern comments for a slide in document order.
    /// (A slide may carry multiple PowerPointCommentParts in principle —
    /// flatten them in part-iteration order; in practice we always create
    /// exactly one.)
    /// </summary>
    private List<(PowerPointCommentPart part, Comment cm)> EnumerateSlideModernComments(SlidePart slidePart)
    {
        var result = new List<(PowerPointCommentPart, Comment)>();
        foreach (var cp in slidePart.GetPartsOfType<PowerPointCommentPart>())
        {
            if (cp.CommentList == null) continue;
            foreach (var cm in cp.CommentList.Elements<Comment>())
                result.Add((cp, cm));
        }
        return result;
    }

    // ----- Path resolution -----

    /// <summary>
    /// Resolve `/slide[N]/moderncomment[K]` to (slidePart, slideIdx, part, comment, commentIdx).
    /// Returns null on no match.
    /// </summary>
    internal (SlidePart slide, int slideIdx, PowerPointCommentPart part, Comment comment, int commentIdx)?
        ResolveModernComment(string path)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            path, @"^/slide\[(\d+)\]/moderncomment\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return null;
        if (!int.TryParse(m.Groups[1].Value, out var slideIdx)) return null;
        if (!int.TryParse(m.Groups[2].Value, out var cmIdx)) return null;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return null;
        var slidePart = slideParts[slideIdx - 1];
        var all = EnumerateSlideModernComments(slidePart);
        if (cmIdx < 1 || cmIdx > all.Count) return null;
        var (cp, cm) = all[cmIdx - 1];
        return (slidePart, slideIdx, cp, cm, cmIdx);
    }

    /// <summary>
    /// Resolve `/slide[N]/moderncomment[K]/reply[R]` to (slidePart, slideIdx,
    /// part, parentComment, parentIdx, reply, replyIdx). Returns null on no match.
    /// </summary>
    internal (SlidePart slide, int slideIdx, PowerPointCommentPart part, Comment parent, int parentIdx, CommentReply reply, int replyIdx)?
        ResolveModernCommentReply(string path)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            path, @"^/slide\[(\d+)\]/moderncomment\[(\d+)\]/reply\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return null;
        var parent = ResolveModernComment($"/slide[{m.Groups[1].Value}]/moderncomment[{m.Groups[2].Value}]");
        if (parent == null) return null;
        if (!int.TryParse(m.Groups[3].Value, out var rIdx)) return null;
        var replyLst = parent.Value.comment.Elements<CommentReplyList>().FirstOrDefault();
        var replies = replyLst?.Elements<CommentReply>().ToList() ?? new();
        if (rIdx < 1 || rIdx > replies.Count) return null;
        return (parent.Value.slide, parent.Value.slideIdx, parent.Value.part,
                parent.Value.comment, parent.Value.commentIdx, replies[rIdx - 1], rIdx);
    }

    // ----- Body builders -----

    private static TextBodyType BuildModernTextBody(string text)
    {
        var tb = new TextBodyType();
        tb.AppendChild(new Drawing.BodyProperties());
        var para = new Drawing.Paragraph();
        if (!string.IsNullOrEmpty(text))
        {
            var run = new Drawing.Run();
            run.AppendChild(new Drawing.RunProperties() { Language = "en-US" });
            run.AppendChild(new Drawing.Text(text));
            para.AppendChild(run);
        }
        else
        {
            // empty paragraph still needs an EndParagraphRunProperties for a:p schema.
            para.AppendChild(new Drawing.EndParagraphRunProperties() { Language = "en-US" });
        }
        tb.AppendChild(para);
        return tb;
    }

    private static string ReadModernTextBody(OpenXmlElement? owner)
    {
        if (owner == null) return "";
        var tb = owner.Elements<TextBodyType>().FirstOrDefault();
        if (tb == null) return "";
        var sb = new System.Text.StringBuilder();
        foreach (var para in tb.Elements<Drawing.Paragraph>())
        {
            if (sb.Length > 0) sb.Append('\n');
            foreach (var r in para.Elements<Drawing.Run>())
            {
                var t = r.GetFirstChild<Drawing.Text>();
                if (t != null) sb.Append(t.Text);
            }
        }
        return sb.ToString();
    }

    private static void ReplaceTextBody(OpenXmlElement owner, string text)
    {
        var existing = owner.Elements<TextBodyType>().ToList();
        foreach (var e in existing) e.Remove();
        owner.AppendChild(BuildModernTextBody(text));
    }

    // ----- Add -----

    private string AddModernComment(string parentPath, int? index, Dictionary<string, string> properties)
    {
        var slideMatch = System.Text.RegularExpressions.Regex.Match(
            parentPath, @"^/slide\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!slideMatch.Success)
            throw new ArgumentException(
                $"modernComment must be added to a slide path like /slide[N], got '{parentPath}'.");
        if (!int.TryParse(slideMatch.Groups[1].Value, out var slideIdx))
            throw new ArgumentException($"Invalid slide index '{slideMatch.Groups[1].Value}'.");
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[slideIdx - 1];

        var text = properties.GetValueOrDefault("text") ?? "";
        XmlTextValidator.ValidateOrThrow(text, "text");
        var author = properties.GetValueOrDefault("author", "OfficeCli");
        XmlTextValidator.ValidateOrThrow(author, "author");
        var initials = properties.GetValueOrDefault("initials", DeriveInitials(author));
        XmlTextValidator.ValidateOrThrow(initials, "initials");

        var created = properties.TryGetValue("created", out var cv) && DateTime.TryParse(cv, out var parsedDt)
            ? NormalizeToUtc(parsedDt)
            : DateTime.UtcNow;

        var au = GetOrCreateModernCommentAuthor(author, initials);
        var part = GetOrCreateModernCommentPart(slidePart);

        // Reply branch — parent= refers to an existing top-level moderncomment.
        // Per OOXML, p188:CommentReply lives inside the parent <p188:cm>'s
        // <p188:replyLst>.
        if (properties.TryGetValue("parent", out var parentRef) && !string.IsNullOrEmpty(parentRef))
        {
            var resolvedParent = ResolveModernComment(parentRef)
                ?? throw new ArgumentException($"Parent modernComment not found: {parentRef}");
            if (resolvedParent.slideIdx != slideIdx)
                throw new ArgumentException(
                    $"Reply slide ({slideIdx}) must match parent comment slide ({resolvedParent.slideIdx}).");

            var reply = new CommentReply
            {
                Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}",
                AuthorId = au.Id?.Value ?? "",
                Created = created,
            };
            reply.AppendChild(BuildModernTextBody(text));

            var replyLst = resolvedParent.comment.Elements<CommentReplyList>().FirstOrDefault();
            if (replyLst == null)
            {
                replyLst = new CommentReplyList();
                resolvedParent.comment.AppendChild(replyLst);
            }
            replyLst.AppendChild(reply);
            part.CommentList!.Save();

            var newRIdx = replyLst.Elements<CommentReply>().ToList().IndexOf(reply) + 1;
            return $"/slide[{slideIdx}]/modernComment[{resolvedParent.commentIdx}]/reply[{newRIdx}]";
        }

        // Top-level comment.
        var cm = new Comment
        {
            Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}",
            AuthorId = au.Id?.Value ?? "",
            Created = created,
        };
        // resolved=true at top-level → status=Resolved (thread-level state).
        if (properties.TryGetValue("resolved", out var rv) && ParseHelpers.IsTruthySafe(rv))
            cm.Status = new EnumValue<CommentStatus>(CommentStatus.Resolved);

        // p188:pos is required — anchor at (0,0) for slide-level. Shape-anchored
        // is signalled by a p188:unknownAnchor extension; v1 stores the shape
        // ref in the Title attribute as a soft anchor (round-trips via Get)
        // and leaves pos=(0,0). (Full extLst anchoring is a future task.)
        cm.AppendChild(new Point2DType { X = 0L, Y = 0L });
        cm.AppendChild(BuildModernTextBody(text));

        part.CommentList!.AppendChild(cm);
        part.CommentList.Save();

        var addedIdx = EnumerateSlideModernComments(slidePart)
            .Select(t => t.cm).ToList().IndexOf(cm) + 1;
        return $"/slide[{slideIdx}]/modernComment[{addedIdx}]";
    }

    // ----- Get / Node builder -----

    /// <summary>
    /// Build a DocumentNode for a top-level modernComment, including its
    /// replies as Children (path /slide[N]/modernComment[K]/reply[R]).
    /// </summary>
    internal DocumentNode ModernCommentToNode(int slideIdx, Comment cm, int cmIdx)
    {
        var node = new DocumentNode
        {
            Path = $"/slide[{slideIdx}]/modernComment[{cmIdx}]",
            Type = "modernComment",
            Text = ReadModernTextBody(cm),
        };
        node.Format["text"] = node.Text;
        var au = FindModernCommentAuthor(cm.AuthorId?.Value);
        if (au != null)
        {
            node.Format["author"] = au.Name?.Value ?? "";
            node.Format["initials"] = au.Initials?.Value ?? "";
        }
        if (cm.Created?.Value != null)
            node.Format["created"] = NormalizeToUtc(cm.Created.Value).ToString("o");
        var resolved = cm.Status?.Value == CommentStatus.Resolved;
        node.Format["resolved"] = resolved;
        node.Format["parent"] = null!;
        if (!string.IsNullOrEmpty(cm.Id?.Value))
            node.Format["id"] = cm.Id!.Value!;

        var replyLst = cm.Elements<CommentReplyList>().FirstOrDefault();
        if (replyLst != null)
        {
            int rIdx = 0;
            foreach (var r in replyLst.Elements<CommentReply>())
            {
                rIdx++;
                node.Children.Add(ModernCommentReplyToNode(slideIdx, cmIdx, r, rIdx));
            }
            node.ChildCount = node.Children.Count;
        }
        return node;
    }

    internal DocumentNode ModernCommentReplyToNode(int slideIdx, int parentIdx, CommentReply r, int rIdx)
    {
        var node = new DocumentNode
        {
            Path = $"/slide[{slideIdx}]/modernComment[{parentIdx}]/reply[{rIdx}]",
            Type = "modernComment",
            Text = ReadModernTextBody(r),
        };
        node.Format["text"] = node.Text;
        var au = FindModernCommentAuthor(r.AuthorId?.Value);
        if (au != null)
        {
            node.Format["author"] = au.Name?.Value ?? "";
            node.Format["initials"] = au.Initials?.Value ?? "";
        }
        if (r.Created?.Value != null)
            node.Format["created"] = NormalizeToUtc(r.Created.Value).ToString("o");
        node.Format["resolved"] = false;
        node.Format["parent"] = $"/slide[{slideIdx}]/modernComment[{parentIdx}]";
        if (!string.IsNullOrEmpty(r.Id?.Value))
            node.Format["id"] = r.Id!.Value!;
        return node;
    }

    /// <summary>
    /// Enumerate all modern comments (top-level only — replies surface as
    /// children of each top-level node).
    /// </summary>
    internal List<DocumentNode> EnumerateModernComments(int? slideIdxFilter = null)
    {
        var slideParts = GetSlideParts().ToList();
        var results = new List<DocumentNode>();
        for (int i = 0; i < slideParts.Count; i++)
        {
            if (slideIdxFilter.HasValue && (i + 1) != slideIdxFilter.Value) continue;
            int idx = 0;
            foreach (var (_, cm) in EnumerateSlideModernComments(slideParts[i]))
            {
                idx++;
                results.Add(ModernCommentToNode(i + 1, cm, idx));
            }
        }
        return results;
    }

    // ----- Set -----

    internal List<string> SetModernCommentProperties(
        PowerPointCommentPart part, OpenXmlElement target, Dictionary<string, string> properties)
    {
        // target is either Comment or CommentReply. Treat their shared
        // attributes (AuthorId, Created, txBody) generically; resolved
        // applies only to the top-level Comment.
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                {
                    XmlTextValidator.ValidateOrThrow(value, key);
                    ReplaceTextBody(target, value);
                    break;
                }
                case "author":
                case "initials":
                {
                    XmlTextValidator.ValidateOrThrow(value, key);
                    string? curAuthId = target switch
                    {
                        Comment c => c.AuthorId?.Value,
                        CommentReply r => r.AuthorId?.Value,
                        _ => null,
                    };
                    var auth = FindModernCommentAuthor(curAuthId);
                    if (auth == null) { unsupported.Add(key); break; }

                    var newName = key.Equals("author", StringComparison.OrdinalIgnoreCase)
                        ? value : (auth.Name?.Value ?? "");
                    var newInitials = key.Equals("initials", StringComparison.OrdinalIgnoreCase)
                        ? value : (auth.Initials?.Value ?? DeriveInitials(newName));

                    var otherRefs = CountModernAuthorReferences(curAuthId ?? "", target);
                    if (otherRefs == 0)
                    {
                        auth.Name = newName;
                        auth.Initials = newInitials;
                    }
                    else
                    {
                        var forked = GetOrCreateModernCommentAuthor(newName, newInitials);
                        if (target is Comment cTop) cTop.AuthorId = forked.Id?.Value ?? "";
                        else if (target is CommentReply cRep) cRep.AuthorId = forked.Id?.Value ?? "";
                    }
                    break;
                }
                case "resolved":
                {
                    if (target is Comment c)
                    {
                        if (ParseHelpers.IsTruthySafe(value))
                            c.Status = new EnumValue<CommentStatus>(CommentStatus.Resolved);
                        else
                            c.Status = null;
                    }
                    else
                    {
                        // resolved is a thread-level state — replies don't carry it.
                        unsupported.Add(key);
                    }
                    break;
                }
                case "created":
                {
                    if (!DateTime.TryParse(value, out var dt))
                        throw new ArgumentException($"Invalid created '{value}' (expected ISO 8601).");
                    var utc = NormalizeToUtc(dt);
                    if (target is Comment c) c.Created = utc;
                    else if (target is CommentReply r) r.Created = utc;
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        return unsupported;
    }

    // ----- Remove -----

    /// <summary>
    /// Remove a modern comment (top-level removes whole thread, mirror of
    /// PowerPoint UI). A reply path removes only that reply.
    /// </summary>
    internal bool RemoveModernComment(string path)
    {
        var topResolved = ResolveModernComment(path);
        if (topResolved.HasValue)
        {
            var (slidePart, _, partTop, cm, _) = topResolved.Value;
            cm.Remove();
            partTop.CommentList!.Save();
            // Drop empty per-slide part to avoid bloat.
            if (!partTop.CommentList.Elements<Comment>().Any())
                slidePart.DeletePart(partTop);
            return true;
        }
        var replyResolved = ResolveModernCommentReply(path);
        if (replyResolved.HasValue)
        {
            var (_, _, partRep, parent, _, reply, _) = replyResolved.Value;
            reply.Remove();
            // If the replyLst is now empty, drop it for tidiness.
            var lst = parent.Elements<CommentReplyList>().FirstOrDefault();
            if (lst != null && !lst.Elements<CommentReply>().Any()) lst.Remove();
            partRep.CommentList!.Save();
            return true;
        }
        return false;
    }
}
