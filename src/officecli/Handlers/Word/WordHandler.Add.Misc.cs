// Copyright 2025 OfficeCli (officecli.ai)
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

        if (!properties.TryGetValue("text", out var commentText))
            throw new ArgumentException("'text' property is required for comment type");

        var commentRun = parent as Run;
        var commentPara = commentRun?.Parent as Paragraph ?? parent as Paragraph
            ?? throw new ArgumentException("Comments must be added to a paragraph or run: /body/p[N] or /body/p[N]/r[M]");

        var author = properties.GetValueOrDefault("author", "officecli");
        var initials = properties.GetValueOrDefault("initials", author[..1]);
        var commentsPart = _doc.MainDocumentPart!.WordprocessingCommentsPart
            ?? _doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments ??= new Comments();

        var commentId = (commentsPart.Comments.Elements<Comment>()
            .Select(c => int.TryParse(c.Id?.Value, out var id) ? id : 0)
            .DefaultIfEmpty(0).Max() + 1).ToString();

        commentsPart.Comments.AppendChild(new Comment(
            new Paragraph(new Run(new Text(commentText) { Space = SpaceProcessingModeValues.Preserve })))
        {
            Id = commentId, Author = author, Initials = initials,
            Date = properties.TryGetValue("date", out var ds) ? DateTime.Parse(ds) : DateTime.UtcNow
        });
        commentsPart.Comments.Save();

        var rangeStart = new CommentRangeStart { Id = commentId };
        var rangeEnd = new CommentRangeEnd { Id = commentId };
        var refRun = new Run(new CommentReference { Id = commentId });

        if (commentRun != null)
        {
            commentRun.InsertBeforeSelf(rangeStart);
            commentRun.InsertAfterSelf(rangeEnd);
            rangeEnd.InsertAfterSelf(refRun);
        }
        else
        {
            // index is a childElement-index (ResolveAnchorPosition counts pPr).
            // Use pPr-aware insert so an index pointing at ParagraphProperties
            // clamps forward (pPr must stay first child).
            if (index.HasValue)
            {
                InsertIntoParagraph(commentPara, new OpenXmlElement[] { rangeStart, rangeEnd, refRun }, index);
            }
            else
            {
                var after = commentPara.ParagraphProperties as OpenXmlElement;
                if (after != null) after.InsertAfterSelf(rangeStart);
                else commentPara.InsertAt(rangeStart, 0);
                commentPara.AppendChild(rangeEnd);
                commentPara.AppendChild(refRun);
            }
        }

        // Return navigable path using /comments/comment[N] (sequential index)
        var commentIndex = commentsPart.Comments.Elements<Comment>().ToList()
            .FindIndex(c => c.Id?.Value == commentId) + 1;
        var resultPath = $"/comments/comment[{commentIndex}]";
        return resultPath;
    }

    private string AddBookmark(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        var bkName = properties.GetValueOrDefault("name", "");
        if (string.IsNullOrEmpty(bkName))
            throw new ArgumentException("'name' property is required for bookmark");

        if (bkName.Any(c => c == '/' || c == '[' || c == ']'))
            throw new ArgumentException(
                $"Bookmark name '{bkName}' contains path-special characters " +
                "('/', '[', ']'). These characters prevent later addressing via " +
                "selectors. Use only letters, digits, '.', '_', '-' in bookmark names.");
        if (bkName.Any(char.IsWhiteSpace) || bkName[0] == '@' || bkName[0] == '\'' || bkName.Contains('"'))
            throw new ArgumentException(
                $"Bookmark name '{bkName}' contains whitespace or quote/@ chars " +
                "that prevent later addressing via bare attribute selectors. " +
                "Use only letters, digits, '.', '_', '-' in bookmark names.");

        // Reject duplicate bookmark names. OOXML bookmark names are expected
        // to be unique per document; tolerating duplicates makes
        // /bookmark[@name=X] ambiguous (it picks the first), so the path
        // returned by `add` may not identify the bookmark just inserted.
        var existingStarts = body.Descendants<BookmarkStart>().ToList();
        if (existingStarts.Any(b => string.Equals(b.Name?.Value, bkName, StringComparison.Ordinal)))
        {
            throw new ArgumentException(
                $"bookmark name '{bkName}' already exists; pick a unique name.");
        }

        var existingIds = existingStarts
            .Select(b => int.TryParse(b.Id?.Value, out var id) ? id : 0);
        var bkId = (existingIds.Any() ? existingIds.Max() + 1 : 1).ToString();

        var bookmarkStart = new BookmarkStart { Id = bkId, Name = bkName };
        var bookmarkEnd = new BookmarkEnd { Id = bkId };

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

        if (properties.TryGetValue("text", out var bkText))
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

        // Return a navigable path: /...parent/bookmarkStart[@name=<name>] is
        // a real DOM element Navigation understands (the legacy
        // `/bookmark[<name>]` form addressed a synthetic type that Get/Add
        // could not resolve, breaking --after/--before reuse).
        // ValidateAndNormalizePredicate rejects bare attribute values that
        // contain whitespace, leading '@', or quote chars; double-quote the
        // value when the raw name would otherwise be rejected so the returned
        // path is round-trippable via `get`/`add --after`.
        string resultPath;
        if (wrappingPara != null)
        {
            var wrapIdx = parent.Elements<Paragraph>().ToList().IndexOf(wrappingPara) + 1;
            resultPath = $"{parentPath}/{BuildParaPathSegment(wrappingPara, wrapIdx)}/bookmarkStart[@name={QuoteAttrValueIfNeeded(bkName)}]";
        }
        else
        {
            resultPath = $"{parentPath}/bookmarkStart[@name={QuoteAttrValueIfNeeded(bkName)}]";
        }
        return resultPath;
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

    private string AddHyperlink(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is `url`
        // (per schemas/help/docx/hyperlink.json). `href` and `link` are legacy
        // input aliases; Get normalizes readback to `url`.
        var hasUrl = properties.TryGetValue("url", out var hlUrl)
            || properties.TryGetValue("href", out hlUrl)
            || properties.TryGetValue("link", out hlUrl);
        var hasAnchor = properties.TryGetValue("anchor", out var hlAnchor) || properties.TryGetValue("bookmark", out hlAnchor);
        if (!hasUrl && !hasAnchor)
            throw new ArgumentException("'url' or 'anchor' property is required for hyperlink type");

        if (parent is not Paragraph hlPara)
            throw new ArgumentException("Hyperlinks can only be added to paragraphs: /body/p[N]");

        string? hlRelId = null;
        if (hasUrl)
        {
            var mainDocPart = _doc.MainDocumentPart!;
            if (!Uri.TryCreate(hlUrl, UriKind.Absolute, out var hlUri))
                throw new ArgumentException($"Invalid hyperlink URL '{hlUrl}'. Expected a valid absolute URI (e.g. 'https://example.com').");
            hlRelId = mainDocPart.AddHyperlinkRelationship(hlUri, isExternal: true).Id;
        }

        var hlRProps = new RunProperties();
        if (properties.TryGetValue("color", out var hlColor))
            hlRProps.Color = new Color { Val = SanitizeHex(hlColor) };
        else
        {
            // Read hyperlink color from document theme, fallback to Word default
            var themeHlink = _doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements
                ?.ColorScheme?.Hyperlink?.RgbColorModelHex?.Val?.Value;
            hlRProps.Color = new Color { Val = themeHlink ?? "0563C1", ThemeColor = ThemeColorValues.Hyperlink };
        }
        hlRProps.Underline = new Underline { Val = UnderlineValues.Single };
        if (properties.TryGetValue("font", out var hlFont))
            hlRProps.RunFonts = new RunFonts { Ascii = hlFont, HighAnsi = hlFont };
        if (properties.TryGetValue("size", out var hlSize))
            hlRProps.FontSize = new FontSize { Val = ((int)Math.Round(ParseFontSize(hlSize) * 2, MidpointRounding.AwayFromZero)).ToString() };
        if (properties.TryGetValue("bold", out var hlBold) && IsTruthy(hlBold))
            hlRProps.Bold = new Bold();
        if (properties.TryGetValue("italic", out var hlItalic) && IsTruthy(hlItalic))
            hlRProps.Italic = new Italic();

        var hlRun = new Run(hlRProps);
        var hlText = properties.GetValueOrDefault("text", hlUrl ?? hlAnchor ?? "link");
        hlRun.AppendChild(new Text(hlText) { Space = SpaceProcessingModeValues.Preserve });

        var hyperlink = new Hyperlink(hlRun);
        if (hlRelId != null)
            hyperlink.Id = hlRelId;
        if (hasAnchor)
            hyperlink.Anchor = hlAnchor;

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

        var fieldInstr = effectiveType switch
        {
            "pagenum" or "pagenumber" or "page" => " PAGE ",
            "numpages" => " NUMPAGES ",
            "sectionpages" => " SECTIONPAGES ",
            "section" => " SECTION ",
            "date" => " DATE \\@ \"yyyy-MM-dd\" ",
            "createdate" => " CREATEDATE \\@ \"yyyy-MM-dd\" ",
            "savedate" => " SAVEDATE \\@ \"yyyy-MM-dd\" ",
            "printdate" => " PRINTDATE \\@ \"yyyy-MM-dd\" ",
            "edittime" => " EDITTIME ",
            "author" => " AUTHOR ",
            "lastsavedby" => " LASTSAVEDBY ",
            "title" => " TITLE ",
            "subject" => " SUBJECT ",
            "filename" => " FILENAME ",
            "time" => " TIME ",
            "numwords" => " NUMWORDS ",
            "numchars" => " NUMCHARS ",
            "revnum" => " REVNUM ",
            "template" => " TEMPLATE ",
            "comments" or "doccomments" => " COMMENTS ",
            "keywords" => " KEYWORDS ",
            "mergefield" => $" MERGEFIELD {mergeFieldName} ",
            "ref" => $" REF {refBookmarkName}{(IsTruthy(properties.GetValueOrDefault("hyperlink")) ? " \\h" : "")} ",
            "pageref" => $" PAGEREF {refBookmarkName}{(IsTruthy(properties.GetValueOrDefault("hyperlink")) ? " \\h" : "")} ",
            "noteref" => $" NOTEREF {refBookmarkName}{(IsTruthy(properties.GetValueOrDefault("hyperlink")) ? " \\h" : "")} ",
            "seq" => $" SEQ {seqIdentifier} ",
            "styleref" => $" STYLEREF \"{styleRefName}\" ",
            "docproperty" => $" DOCPROPERTY \"{docPropertyName}\" ",
            "if" => BuildIfFieldInstruction(properties),
            _ => properties.ContainsKey("instruction")
                ? properties["instruction"]
                : throw new ArgumentException($"Unknown field type '{effectiveType}'. Provide a known type or an 'instruction' property.")
        };
        // Allow override via property
        if (properties.TryGetValue("instruction", out var instr))
            fieldInstr = instr.StartsWith(" ") ? instr : $" {instr} ";

        var fieldPlaceholder = properties.ContainsKey("text")
            ? properties["text"]
            : effectiveType switch
            {
                "mergefield" => $"\u00AB{mergeFieldName}\u00BB",
                "ref" or "noteref" => $"\u00AB{refBookmarkName}\u00BB",
                "styleref" => $"\u00AB{styleRefName}\u00BB",
                "docproperty" => $"\u00AB{docPropertyName}\u00BB",
                "if" => properties.GetValueOrDefault("trueText", ""),
                _ => "1"
            };

        // Build complex field: fldChar(begin) + instrText + fldChar(separate) + result + fldChar(end)
        var fieldRunBegin = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
        var fieldRunInstr = new Run(new FieldCode(fieldInstr) { Space = SpaceProcessingModeValues.Preserve });
        var fieldRunSep = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
        var fieldRunResult = new Run(new Text(fieldPlaceholder) { Space = SpaceProcessingModeValues.Preserve });
        var fieldRunEnd = new Run(new FieldChar { FieldCharType = FieldCharValues.End });

        // Apply optional run formatting to all runs
        RunProperties? fieldRProps = null;
        if (properties.TryGetValue("font", out var fFont) || properties.TryGetValue("size", out _) ||
            properties.TryGetValue("bold", out _) || properties.TryGetValue("color", out _))
        {
            fieldRProps = new RunProperties();
            // CT_RPr schema order: rFonts → b → ... → color → sz
            if (properties.TryGetValue("font", out var ff))
                fieldRProps.AppendChild(new RunFonts { Ascii = ff, HighAnsi = ff, EastAsia = ff });
            if (properties.TryGetValue("bold", out var fb) && IsTruthy(fb))
                fieldRProps.AppendChild(new Bold());
            if (properties.TryGetValue("color", out var fc))
                fieldRProps.AppendChild(new Color { Val = SanitizeHex(fc) });
            if (properties.TryGetValue("size", out var fs))
                fieldRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(fs) * 2, MidpointRounding.AwayFromZero)).ToString() });
        }

        if (fieldRProps != null)
        {
            fieldRunBegin.PrependChild(fieldRProps.CloneNode(true));
            fieldRunInstr.PrependChild(fieldRProps.CloneNode(true));
            fieldRunSep.PrependChild(fieldRProps.CloneNode(true));
            fieldRunResult.PrependChild(fieldRProps.CloneNode(true));
            fieldRunEnd.PrependChild(fieldRProps.CloneNode(true));
        }

        string resultPath;
        if (parent is Paragraph fieldPara)
        {
            // index is a childElement-index (ResolveAnchorPosition counts pPr too).
            // Route the 5 field runs through the pPr-aware multi-insert helper
            // so index 0 clamps forward past ParagraphProperties and they stay
            // in the correct consecutive order.
            if (index.HasValue)
            {
                InsertIntoParagraph(
                    fieldPara,
                    new OpenXmlElement[] { fieldRunBegin, fieldRunInstr, fieldRunSep, fieldRunResult, fieldRunEnd },
                    index);
                var runIdxAfterInsert = fieldPara.Elements<Run>().TakeWhile(r => r != fieldRunResult).Count();
                resultPath = $"{parentPath}/r[{runIdxAfterInsert + 1}]";
            }
            else
            {
                fieldPara.AppendChild(fieldRunBegin);
                fieldPara.AppendChild(fieldRunInstr);
                fieldPara.AppendChild(fieldRunSep);
                fieldPara.AppendChild(fieldRunResult);
                fieldPara.AppendChild(fieldRunEnd);
                var runIdx = GetAllRuns(fieldPara).Count - 4;
                resultPath = $"{parentPath}/r[{runIdx}]";
            }
        }
        else
        {
            // Create a new paragraph containing the field
            var fNewPara = new Paragraph();
            var fPProps = new ParagraphProperties();
            if (properties.TryGetValue("alignment", out var fAlign))
                fPProps.Justification = new Justification { Val = ParseJustification(fAlign) };
            fNewPara.AppendChild(fPProps);
            fNewPara.AppendChild(fieldRunBegin);
            fNewPara.AppendChild(fieldRunInstr);
            fNewPara.AppendChild(fieldRunSep);
            fNewPara.AppendChild(fieldRunResult);
            fNewPara.AppendChild(fieldRunEnd);
            InsertAtIndexOrAppend(parent, fNewPara, index);
            var fIdx2 = body.Elements<Paragraph>().TakeWhile(p => p != fNewPara).Count();
            resultPath = $"/body/{BuildParaPathSegment(fNewPara, fIdx2 + 1)}";
        }
        return resultPath;
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
        // Allow override via property
        if (properties.TryGetValue("type", out var brType))
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
        var brkRun = new Run(brk);

        string resultPath;
        if (parent is Paragraph brkPara)
        {
            // index is a childElement-index (ResolveAnchorPosition counts pPr).
            // pPr-aware insert keeps pPr as the first child of <w:p>.
            InsertIntoParagraph(brkPara, brkRun, index);
            var brkParaIdx = body.Elements<Paragraph>().TakeWhile(p => p != brkPara).Count();
            var brkRunIdx = brkPara.Elements<Run>().TakeWhile(r => r != brkRun).Count() + 1;
            resultPath = $"/body/{BuildParaPathSegment(brkPara, brkParaIdx + 1)}/r[{brkRunIdx}]";
        }
        else
        {
            // Create a new empty paragraph with the break and insert into the
            // ACTUAL parent (not hard-coded body) so /header[N], /footer[N],
            // table cells, etc. receive the new paragraph. /styles is blocked
            // earlier by ValidateParentChild.
            var brkNewPara = new Paragraph(brkRun);
            InsertAtIndexOrAppend(parent, brkNewPara, index);
            // Build a navigable path based on the actual container. Only /body
            // has the @paraId indexer; other containers (headers/footers/cells)
            // use positional paths rooted at parentPath.
            if (parent is Body)
            {
                var brkIdx = body.Elements<Paragraph>().TakeWhile(p => p != brkNewPara).Count();
                resultPath = $"/body/{BuildParaPathSegment(brkNewPara, brkIdx + 1)}";
            }
            else
            {
                var brkIdx = parent.Elements<Paragraph>().TakeWhile(p => p != brkNewPara).Count();
                resultPath = $"{parentPath}/p[{brkIdx + 1}]";
            }
        }
        return resultPath;
    }

    private string AddSdt(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // Case-insensitive lookup to support camelCase keys like "sdtType", "controlType", etc.
        var ciProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);

        // Add a Structured Document Tag (Content Control)
        // Canonical key is "type" (per schemas/help/docx/sdt.json); "sdttype" / "controltype"
        // retained as legacy aliases for backward-compat.
        var sdtType = ciProps.GetValueOrDefault("type",
            ciProps.GetValueOrDefault("sdttype",
                ciProps.GetValueOrDefault("controltype", "text"))).ToLowerInvariant();
        // Schema-honesty: reject values the SDT builder does not emit the
        // correct child elements for. Keeps the schema and runtime in sync
        // instead of silently falling back to plain-text SDT.
        var supportedSdtTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "text", "plaintext", "richtext", "rich",
            "dropdown", "dropdownlist", "combobox", "combo",
            "date", "datepicker"
        };
        if (!supportedSdtTypes.Contains(sdtType))
            throw new NotSupportedException(
                $"SDT type '{sdtType}' is not implemented. Supported: text, richtext, dropdown, combobox, date. " +
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
                    if (ciProps.TryGetValue("items", out var items))
                    {
                        foreach (var item in items.Split(','))
                        {
                            var trimmed = item.Trim();
                            ddl.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                        }
                    }
                    sdtProps.AppendChild(ddl);
                    break;
                }
                case "combobox" or "combo":
                {
                    var cb = new SdtContentComboBox();
                    if (ciProps.TryGetValue("items", out var items))
                    {
                        foreach (var item in items.Split(','))
                        {
                            var trimmed = item.Trim();
                            cb.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                        }
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
                case "richtext" or "rich":
                    // Rich text has no specific type element (absence of w:text means rich text)
                    break;
                default: // "text" or "plaintext"
                    sdtProps.AppendChild(new SdtContentText());
                    break;
            }

            sdtRun.AppendChild(sdtProps);
            var sdtContent = new SdtContentRun();
            var contentRun = new Run(new Text(sdtText) { Space = SpaceProcessingModeValues.Preserve });
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
                    if (ciProps.TryGetValue("items", out var items))
                    {
                        foreach (var item in items.Split(','))
                        {
                            var trimmed = item.Trim();
                            ddl.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                        }
                    }
                    sdtProps.AppendChild(ddl);
                    break;
                }
                case "combobox" or "combo":
                {
                    var cb = new SdtContentComboBox();
                    if (ciProps.TryGetValue("items", out var items))
                    {
                        foreach (var item in items.Split(','))
                        {
                            var trimmed = item.Trim();
                            cb.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                        }
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
                case "richtext" or "rich":
                    break;
                default:
                    sdtProps.AppendChild(new SdtContentText());
                    break;
            }

            sdtBlock.AppendChild(sdtProps);
            var sdtContent = new SdtContentBlock();
            var contentPara = new Paragraph(new Run(new Text(sdtText) { Space = SpaceProcessingModeValues.Preserve }));
            sdtContent.AppendChild(contentPara);
            sdtBlock.AppendChild(sdtContent);

            InsertAtIndexOrAppend(parent, sdtBlock, index);
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

    private string AddWatermark(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var wmText = properties.GetValueOrDefault("text", "DRAFT");
        // VML watermarks accept named colors (silver, red, etc.) or hex — don't sanitize
        var wmColor = properties.TryGetValue("color", out var wmcVal)
            ? wmcVal.TrimStart('#') : "silver";
        var wmFont = properties.GetValueOrDefault("font", OfficeDefaultFonts.MinorLatin);
        var wmSize = properties.GetValueOrDefault("size", "1pt");
        if (!wmSize.EndsWith("pt")) wmSize += "pt";
        var wmRotation = properties.GetValueOrDefault("rotation", "315");
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
        var created = GenericXmlQuery.TryCreateTypedElement(parent, type, properties, index);
        if (created == null)
            throw new ArgumentException($"Unknown element type '{type}' for {parentPath}. " +
                "Valid types: paragraph (p), run (r), table (tbl), row, cell, picture, chart, ole (object, embed), equation, comment, section, footnote, endnote, toc, style, watermark, bookmark, hyperlink, field, break, sdt, header, footer. " +
                "Use 'officecli docx add' for details.");

        var siblings = parent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
        var createdIdx = siblings.IndexOf(created) + 1;
        var resultPath = $"{parentPath}/{created.LocalName}[{createdIdx}]";
        return resultPath;
    }
}
