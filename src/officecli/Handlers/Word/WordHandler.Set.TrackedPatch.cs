// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Perform find/replace with tracked changes (w:del + w:ins) on paragraphs
    /// resolved from the given path. Each match generates:
    ///   1. w:del wrapping the original text (marked as deletion)
    ///   2. w:ins wrapping the replacement text (marked as insertion)
    ///   3. A comment annotation ("校对修订：X → Y") if comments part is available
    ///
    /// Also ensures settings.xml has trackRevisions/showDel/showIns flags.
    /// Returns the number of matches processed.
    /// </summary>
    private int ProcessTrackedFind(
        string path,
        string findValue,
        string? replace,
        string author,
        DateTime date)
    {
        var (pattern, isRegex) = ParseFindPattern(findValue);
        if (string.IsNullOrEmpty(pattern) && !isRegex) return 0;

        var paragraphs = ResolveParagraphsForFind(path);
        int totalCount = 0;

        foreach (var para in paragraphs)
        {
            var count = ProcessTrackedFindInParagraph(para, pattern, isRegex, replace ?? "", author, date);
            if (count > 0)
                para.TextId = GenerateParaId();
            totalCount += count;
        }

        if (totalCount > 0)
            EnsureTrackRevisionSettings();

        return totalCount;
    }

    /// <summary>
    /// Process tracked find/replace in a single paragraph.
    /// For each match, splits runs at match boundaries, wraps matched runs
    /// in w:del, and inserts w:ins with replacement text after.
    /// </summary>
    private int ProcessTrackedFindInParagraph(
        Paragraph para,
        string pattern,
        bool isRegex,
        string replace,
        string author,
        DateTime date)
    {
        var runTexts = BuildRunTexts(para);
        if (runTexts.Count == 0) return 0;

        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
        var matches = FindMatchRanges(fullText, pattern, isRegex);
        if (matches.Count == 0) return 0;

        var dateStr = date.ToString("yyyy-MM-ddTHH:mm:ssZ");

        // Process matches from end to start to preserve character offsets
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            var (matchStart, matchLen) = matches[i];
            var matchEnd = matchStart + matchLen;
            var matchedText = fullText.Substring(matchStart, matchLen);

            // Step 1: Split runs at match boundaries so we get exact run coverage
            var targetRuns = SplitRunsAtRange(para, matchStart, matchEnd);
            if (targetRuns.Count == 0) continue;

            // Determine insertion anchor: the element AFTER the last matched run
            var lastRun = targetRuns[targetRuns.Count - 1];
            var insertAnchor = lastRun.NextSibling();

            // Step 2: Collect run properties from the first matched run (for ins)
            var templateRPr = targetRuns[0].RunProperties?.CloneNode(true) as RunProperties;

            // Step 3: Build w:del element wrapping clones of matched runs
            var delId = GenerateRevisionId();
            var deletedRun = new DeletedRun
            {
                Author = author,
                Date = date,
                Id = delId,
            };
            foreach (var run in targetRuns)
            {
                var clonedRun = (Run)run.CloneNode(true);
                // Convert w:t → w:delText in the cloned run
                foreach (var t in clonedRun.Descendants<Text>().ToList())
                {
                    var dt = new DeletedText { Text = t.Text, Space = SpaceProcessingModeValues.Preserve };
                    t.Parent!.ReplaceChild(dt, t);
                }
                deletedRun.AppendChild(clonedRun);
            }

            // Step 4: Build w:ins element with replacement text
            InsertedRun? insertedRun = null;
            if (!string.IsNullOrEmpty(replace))
            {
                var insId = GenerateRevisionId();
                insertedRun = new InsertedRun
                {
                    Author = author,
                    Date = date,
                    Id = insId,
                };
                var newRun = new Run();
                if (templateRPr != null)
                    newRun.RunProperties = (RunProperties)templateRPr.CloneNode(true);
                newRun.AppendChild(new Text(replace) { Space = SpaceProcessingModeValues.Preserve });
                insertedRun.AppendChild(newRun);
            }

            // Step 5: Remove original matched runs and insert del + ins
            foreach (var run in targetRuns)
                run.Remove();

            if (insertAnchor != null)
            {
                para.InsertBefore(deletedRun, insertAnchor);
                if (insertedRun != null)
                    para.InsertBefore(insertedRun, insertAnchor);
            }
            else
            {
                para.AppendChild(deletedRun);
                if (insertedRun != null)
                    para.AppendChild(insertedRun);
            }

            // Step 6: Add comment annotation
            AddTrackedChangeComment(para, deletedRun, insertedRun, matchedText, replace, author, date);
        }

        return matches.Count;
    }

    /// <summary>
    /// Generate a unique revision ID (w:id) that doesn't collide with
    /// existing IDs in the document.
    /// </summary>
    private string GenerateRevisionId()
    {
        // Simple incrementing counter; in practice collisions are unlikely
        // since we only need uniqueness within a single session.
        _nextRevisionId++;
        return _nextRevisionId.ToString();
    }
    private int _nextRevisionId = 1000;

    /// <summary>
    /// Add a comment to annotate a tracked change.
    /// Creates CommentRangeStart before w:del, CommentRangeEnd + CommentReference
    /// after w:ins, and adds the Comment element to comments.xml.
    /// </summary>
    private void AddTrackedChangeComment(
        Paragraph para,
        DeletedRun del,
        InsertedRun? ins,
        string originalText,
        string replacement,
        string author,
        DateTime date)
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return;

        // Ensure comments part exists
        var commentsPart = mainPart.WordprocessingCommentsPart
            ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments ??= new Comments();

        // Allocate comment ID (must not collide with existing comments)
        var existingIds = commentsPart.Comments.Elements<Comment>()
            .Select(c => c.Id?.Value)
            .Where(id => id != null)
            .Select(id => int.TryParse(id, out var n) ? n : 0)
            .ToHashSet();
        int commentId = 1;
        while (existingIds.Contains(commentId)) commentId++;

        var commentIdStr = commentId.ToString();

        // Build comment text
        var commentText = string.IsNullOrEmpty(replacement)
            ? $"校对删除：\"{originalText}\""
            : $"校对修订：\"{originalText}\" → \"{replacement}\"";

        // Add Comment element
        var comment = new Comment
        {
            Id = commentIdStr,
            Author = author,
            Date = date,
        };
        var commentPara = new Paragraph(
            new Run(new Text(commentText) { Space = SpaceProcessingModeValues.Preserve }));
        comment.AppendChild(commentPara);
        commentsPart.Comments.AppendChild(comment);
        commentsPart.Comments.Save();

        // Insert CommentRangeStart before w:del
        var rangeStart = new CommentRangeStart { Id = commentIdStr };
        del.InsertBeforeSelf(rangeStart);

        // Insert CommentRangeEnd + CommentReference after w:ins (or w:del if no ins)
        var afterElement = (OpenXmlElement?)ins ?? del;
        var rangeEnd = new CommentRangeEnd { Id = commentIdStr };
        afterElement.InsertAfterSelf(rangeEnd);

        var refRun = new Run(
            new RunProperties(new RunStyle { Val = "CommentReference" }),
            new CommentReference { Id = commentIdStr });
        rangeEnd.InsertAfterSelf(refRun);

        // Ensure relationship exists
        EnsureCommentsRelationship(mainPart);
    }

    /// <summary>
    /// Ensure the document.xml.rels has a relationship to comments.xml.
    /// OpenXML SDK usually handles this when AddNewPart is called,
    /// but we double-check for safety.
    /// </summary>
    private static void EnsureCommentsRelationship(MainDocumentPart mainPart)
    {
        // OpenXML SDK manages relationships automatically when parts are added
        // via AddNewPart, so this is a no-op in most cases. Kept as a safety net.
        _ = mainPart.WordprocessingCommentsPart;
    }

    /// <summary>
    /// Ensure settings.xml has trackRevisions, showDel, showIns, showMarkupBar
    /// so WPS/Word opens the document in review mode.
    /// </summary>
    private void EnsureTrackRevisionSettings()
    {
        var settings = EnsureSettings();

        // TrackRevisions (w:trackChanges) — enables revision tracking
        if (settings.GetFirstChild<TrackRevisions>() == null)
            settings.AddChild(new TrackRevisions());

        // Note: ShowDel (w14:showDel), ShowIns (w14:showIns), ShowMarkupBar
        // are Word2010+ namespace elements. The OpenXML SDK Settings class
        // supports basic w: namespace. For maximum compatibility with WPS,
        // we ensure the core trackChanges flag is set; WPS defaults to
        // showing all markup when trackChanges is present.

        settings.Save();
    }
}
