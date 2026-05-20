// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

// BUG-R36-B11: PPTX legacy slide comments — full add/get/query/set/remove
// lifecycle. Comments live in two parts:
//   - presentation-level CommentAuthorsPart  (commentAuthors.xml)
//   - per-slide SlideCommentsPart           (comments/commentN.xml)
// Path form: /slide[N]/comment[M] (1-based, document order on the slide).
// Properties: text, author, initials, x, y, date.
public partial class PowerPointHandler
{
    /// <summary>
    /// Resolve or create the workbook-level CommentAuthorsPart and return the
    /// CommentAuthor with the requested name+initials, creating one if it
    /// doesn't yet exist. Author ids are assigned monotonically starting at 0.
    /// </summary>
    private CommentAuthor GetOrCreateCommentAuthor(string name, string initials)
    {
        var pres = _doc.PresentationPart!;
        var authorsPart = pres.CommentAuthorsPart;
        if (authorsPart == null)
        {
            authorsPart = pres.AddNewPart<CommentAuthorsPart>();
            authorsPart.CommentAuthorList = new CommentAuthorList();
        }
        authorsPart.CommentAuthorList ??= new CommentAuthorList();

        var existing = authorsPart.CommentAuthorList.Elements<CommentAuthor>()
            .FirstOrDefault(a => string.Equals(a.Name?.Value, name, StringComparison.Ordinal)
                              && string.Equals(a.Initials?.Value, initials, StringComparison.Ordinal));
        if (existing != null) return existing;

        var nextId = (uint)(authorsPart.CommentAuthorList.Elements<CommentAuthor>()
            .Select(a => (int)(a.Id?.Value ?? 0)).DefaultIfEmpty(-1).Max() + 1);
        var author = new CommentAuthor
        {
            Id = nextId,
            Name = name,
            Initials = initials,
            LastIndex = 0,
            ColorIndex = 0,
        };
        authorsPart.CommentAuthorList.AppendChild(author);
        authorsPart.CommentAuthorList.Save();
        return author;
    }

    private SlideCommentsPart GetOrCreateSlideCommentsPart(SlidePart slidePart)
    {
        var commentsPart = slidePart.SlideCommentsPart;
        if (commentsPart == null)
        {
            commentsPart = slidePart.AddNewPart<SlideCommentsPart>();
            commentsPart.CommentList = new CommentList();
        }
        commentsPart.CommentList ??= new CommentList();
        return commentsPart;
    }

    private string AddSlideComment(string parentPath, int? index, Dictionary<string, string> properties)
    {
        // parentPath: /slide[N]
        var slideMatch = System.Text.RegularExpressions.Regex.Match(
            parentPath, @"^/slide\[(\d+)\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!slideMatch.Success)
            throw new ArgumentException(
                $"comment must be added to a slide path like /slide[N], got '{parentPath}'.");
        if (!int.TryParse(slideMatch.Groups[1].Value, out var slideIdx))
            throw new ArgumentException($"Invalid slide index '{slideMatch.Groups[1].Value}'.");
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[slideIdx - 1];

        var text = properties.GetValueOrDefault("text") ?? properties.GetValueOrDefault("comment") ?? "";
        XmlTextValidator.ValidateOrThrow(text, "text");
        var author = properties.GetValueOrDefault("author", "OfficeCli");
        var initials = properties.GetValueOrDefault("initials", DeriveInitials(author));

        // R7-bt-5: PPT comment direction surface. p:cm has no rtl attribute
        // and the body is a plain p:text (no rPr/pPr). Mirror the
        // pure-text RTL convention: prepend U+200F (RIGHT-TO-LEFT MARK) on
        // direction=rtl so PowerPoint and viewers render the comment with
        // Arabic / Hebrew bidi context. ltr / unset leaves the text alone.
        // No UNSUPPORTED — the key is consumed.
        if ((properties.TryGetValue("direction", out var pcmDir)
             || properties.TryGetValue("dir", out pcmDir)
             || properties.TryGetValue("rtl", out pcmDir))
            && ParsePptDirectionRtl(pcmDir)
            && !string.IsNullOrEmpty(text)
            && text[0] != '‏')
        {
            text = "‏" + text;
        }

        var ca = GetOrCreateCommentAuthor(author, initials);

        // x/y positions are stored in EMUs internally; OOXML p:cm uses a CT_Point
        // with 1/100th of EMU? actually p:pos is CT_Point2D (Int64Value, EMU).
        // Default to top-left if omitted.
        var x = properties.TryGetValue("x", out var xv) ? EmuConverter.ParseEmu(xv) : 0L;
        var y = properties.TryGetValue("y", out var yv) ? EmuConverter.ParseEmu(yv) : 0L;
        // CONSISTENCY(comment-date-utc): pin OOXML p:cm/@dt to UTC `Z` form so
        // round-trip readback is timezone-agnostic. DateTime.TryParse with a
        // local-TZ or `Z` input produces Kind=Local/Utc; OpenXml then serializes
        // local-kind values with the host offset (e.g. +08:00), and re-read
        // gives a different readback string on differently-tz'd machines.
        var dt = properties.TryGetValue("date", out var dv) && DateTime.TryParse(dv, out var parsedDt)
            ? NormalizeToUtc(parsedDt)
            : DateTime.UtcNow;

        var commentsPart = GetOrCreateSlideCommentsPart(slidePart);

        // Per-author monotonic comment index; PowerPoint expects ca:lastIdx to
        // track the last issued idx so authoring is unambiguous.
        var lastIdx = (uint)(ca.LastIndex?.Value ?? 0);
        var newIdx = lastIdx + 1;
        ca.LastIndex = newIdx;

        var comment = new Comment
        {
            AuthorId = ca.Id?.Value ?? 0,
            DateTime = dt,
            Index = newIdx,
        };
        comment.AppendChild(new Position { X = (int)x, Y = (int)y });
        comment.AppendChild(new DocumentFormat.OpenXml.Presentation.Text { InnerXml = "" });
        var textEl = comment.GetFirstChild<DocumentFormat.OpenXml.Presentation.Text>()!;
        textEl.Text = text;

        if (index.HasValue)
        {
            var existing = commentsPart.CommentList!.Elements<Comment>().ToList();
            if (index.Value < 0 || index.Value > existing.Count)
                commentsPart.CommentList.AppendChild(comment);
            else if (index.Value == 0)
                commentsPart.CommentList.PrependChild(comment);
            else
                existing[index.Value - 1].InsertAfterSelf(comment);
        }
        else
        {
            commentsPart.CommentList!.AppendChild(comment);
        }
        commentsPart.CommentList.Save();

        var addedIdx = commentsPart.CommentList.Elements<Comment>().ToList().IndexOf(comment) + 1;
        return $"/slide[{slideIdx}]/comment[{addedIdx}]";
    }

    // CONSISTENCY(comment-date-utc): normalize any DateTime to Kind=Utc so
    // OOXML serialization writes the `Z` form (not the host-local offset) and
    // readback is identical on every machine. Unspecified-kind values are
    // treated as already-UTC rather than local — the caller's input is almost
    // always an ISO string and a missing offset is more often "no info" than
    // "machine local".
    private static DateTime NormalizeToUtc(DateTime dt) => dt.Kind switch
    {
        DateTimeKind.Utc => dt,
        DateTimeKind.Local => dt.ToUniversalTime(),
        _ => DateTime.SpecifyKind(dt, DateTimeKind.Utc),
    };

    private static string DeriveInitials(string name)
    {
        // Pull the first letter/digit from each whitespace-separated token,
        // skipping leading punctuation. Authors commonly embed email or
        // handle suffixes ("Author 1 <test@example.com>", "Jane (Acme)"),
        // and the prior implementation picked up the opening '<' / '(' as
        // an initial — producing "A1<" or "J(" instead of a useful tag.
        if (string.IsNullOrWhiteSpace(name)) return "?";
        var parts = name.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) return "?";

        static char? FirstWordChar(string token)
        {
            foreach (var ch in token)
                if (char.IsLetterOrDigit(ch)) return char.ToUpperInvariant(ch);
            return null;
        }

        if (parts.Length == 1)
        {
            var letters = parts[0].Where(char.IsLetterOrDigit)
                .Select(char.ToUpperInvariant).Take(2).ToArray();
            return letters.Length == 0 ? "?" : new string(letters);
        }

        var picks = parts.Select(FirstWordChar)
            .Where(c => c.HasValue)
            .Select(c => c!.Value)
            .Take(3)
            .ToArray();
        return picks.Length == 0 ? "?" : new string(picks);
    }

    /// <summary>Resolve a /slide[N]/comment[M] path to (slidePart, comment).</summary>
    internal (SlidePart slide, int slideIdx, Comment comment, int commentIdx)? ResolveSlideComment(string path)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            path, @"^/slide\[(\d+)\]/comment\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return null;
        if (!int.TryParse(m.Groups[1].Value, out var slideIdx)) return null;
        if (!int.TryParse(m.Groups[2].Value, out var commentIdx)) return null;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return null;
        var slidePart = slideParts[slideIdx - 1];
        var commentsPart = slidePart.SlideCommentsPart;
        if (commentsPart?.CommentList == null) return null;
        var comments = commentsPart.CommentList.Elements<Comment>().ToList();
        if (commentIdx < 1 || commentIdx > comments.Count) return null;
        return (slidePart, slideIdx, comments[commentIdx - 1], commentIdx);
    }

    /// <summary>Build a DocumentNode for a single comment.</summary>
    internal DocumentNode CommentToNode(SlidePart slidePart, int slideIdx, Comment comment, int commentIdx)
    {
        var node = new DocumentNode
        {
            Path = $"/slide[{slideIdx}]/comment[{commentIdx}]",
            Type = "comment",
            Text = comment.GetFirstChild<DocumentFormat.OpenXml.Presentation.Text>()?.Text ?? "",
        };
        node.Format["text"] = node.Text;
        var authId = comment.AuthorId?.Value;
        var authors = _doc.PresentationPart?.CommentAuthorsPart?.CommentAuthorList?
            .Elements<CommentAuthor>().ToList();
        if (authId.HasValue && authors != null)
        {
            var auth = authors.FirstOrDefault(a => a.Id?.Value == authId.Value);
            if (auth != null)
            {
                node.Format["author"] = auth.Name?.Value ?? "";
                node.Format["initials"] = auth.Initials?.Value ?? "";
            }
        }
        node.Format["index"] = (int)(comment.Index?.Value ?? 0);
        // CONSISTENCY(comment-date-utc): always emit UTC `Z` regardless of the
        // on-disk @dt's stored offset, so Get and query give identical readback
        // across machines with different local time zones.
        if (comment.DateTime?.Value != null)
            node.Format["date"] = NormalizeToUtc(comment.DateTime.Value).ToString("o");
        var pos = comment.GetFirstChild<Position>();
        if (pos != null)
        {
            node.Format["x"] = EmuConverter.FormatEmu(pos.X?.Value ?? 0);
            node.Format["y"] = EmuConverter.FormatEmu(pos.Y?.Value ?? 0);
        }
        return node;
    }

    /// <summary>List comments for /slide[N] (slideIdx 1-based) or whole deck.</summary>
    internal List<DocumentNode> EnumerateComments(int? slideIdxFilter = null)
    {
        var slideParts = GetSlideParts().ToList();
        var results = new List<DocumentNode>();
        for (int i = 0; i < slideParts.Count; i++)
        {
            if (slideIdxFilter.HasValue && (i + 1) != slideIdxFilter.Value) continue;
            var commentsPart = slideParts[i].SlideCommentsPart;
            if (commentsPart?.CommentList == null) continue;
            var cmts = commentsPart.CommentList.Elements<Comment>().ToList();
            for (int j = 0; j < cmts.Count; j++)
                results.Add(CommentToNode(slideParts[i], i + 1, cmts[j], j + 1));
        }
        return results;
    }

    internal List<string> SetSlideCommentProperties(Comment comment, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                case "comment":
                {
                    XmlTextValidator.ValidateOrThrow(value, key);
                    var t = comment.GetFirstChild<DocumentFormat.OpenXml.Presentation.Text>();
                    if (t == null)
                    {
                        t = new DocumentFormat.OpenXml.Presentation.Text();
                        comment.AppendChild(t);
                    }
                    t.Text = value;
                    break;
                }
                case "author":
                case "initials":
                {
                    var authId = comment.AuthorId?.Value ?? 0;
                    var authorsPart = _doc.PresentationPart?.CommentAuthorsPart;
                    var auth = authorsPart?.CommentAuthorList?.Elements<CommentAuthor>()
                        .FirstOrDefault(a => a.Id?.Value == authId);
                    if (auth == null) { unsupported.Add(key); break; }
                    XmlTextValidator.ValidateOrThrow(value, key);

                    // CommentAuthor is a shared record keyed by id. Mutating
                    // Name/Initials in place rewrites attribution for every
                    // comment that references the same authorId — silently
                    // poisoning siblings. If any other comment (on any slide)
                    // still references this author, fork: route this comment
                    // to a new CommentAuthor with the requested name+initials
                    // (reusing an existing match if one happens to exist).
                    // If this is the sole reference, mutate in place.
                    var newName = key.Equals("author", StringComparison.OrdinalIgnoreCase)
                        ? value : (auth.Name?.Value ?? "");
                    var newInitials = key.Equals("initials", StringComparison.OrdinalIgnoreCase)
                        ? value : (auth.Initials?.Value ?? DeriveInitials(newName));

                    var otherRefs = CountAuthorReferences(authId, comment);
                    if (otherRefs == 0)
                    {
                        // Sole user — safe to mutate the shared record.
                        auth.Name = newName;
                        auth.Initials = newInitials;
                    }
                    else
                    {
                        // Fork: assign this comment to a (possibly new) author.
                        var forked = GetOrCreateCommentAuthor(newName, newInitials);
                        comment.AuthorId = forked.Id?.Value ?? 0;
                    }
                    break;
                }
                case "x":
                case "y":
                {
                    var pos = comment.GetFirstChild<Position>() ?? comment.AppendChild(new Position { X = 0, Y = 0 });
                    var emu = (int)EmuConverter.ParseEmu(value);
                    if (key.Equals("x", StringComparison.OrdinalIgnoreCase)) pos.X = emu;
                    else pos.Y = emu;
                    break;
                }
                case "date":
                {
                    if (DateTime.TryParse(value, out var dt))
                        comment.DateTime = NormalizeToUtc(dt);
                    else
                        throw new ArgumentException($"Invalid date '{value}' (expected ISO 8601).");
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        return unsupported;
    }

    /// <summary>
    /// Count how many comments across the whole deck reference the given
    /// authorId, excluding the supplied comment. Used by author/initials
    /// Set to decide whether mutating the shared CommentAuthor record is
    /// safe or whether we must fork into a new author record.
    /// </summary>
    private int CountAuthorReferences(uint authId, Comment exclude)
    {
        int count = 0;
        foreach (var sp in GetSlideParts())
        {
            var cp = sp.SlideCommentsPart;
            if (cp?.CommentList == null) continue;
            foreach (var c in cp.CommentList.Elements<Comment>())
            {
                if (ReferenceEquals(c, exclude)) continue;
                if ((c.AuthorId?.Value ?? 0) == authId) count++;
            }
        }
        return count;
    }

    internal bool RemoveSlideComment(string path)
    {
        var resolved = ResolveSlideComment(path);
        if (resolved == null) return false;
        var (slidePart, _, comment, _) = resolved.Value;
        comment.Remove();
        slidePart.SlideCommentsPart!.CommentList!.Save();
        // If this was the last comment on the slide, drop the SlideCommentsPart
        // entirely so empty XML files don't bloat the package.
        if (!slidePart.SlideCommentsPart.CommentList.Elements<Comment>().Any())
            slidePart.DeletePart(slidePart.SlideCommentsPart);
        return true;
    }
}
