// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

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
            var after = commentPara.ParagraphProperties as OpenXmlElement;
            if (after != null) after.InsertAfterSelf(rangeStart);
            else commentPara.InsertAt(rangeStart, 0);
            commentPara.AppendChild(rangeEnd);
            commentPara.AppendChild(refRun);
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

        var existingIds = body.Descendants<BookmarkStart>()
            .Select(b => int.TryParse(b.Id?.Value, out var id) ? id : 0);
        var bkId = (existingIds.Any() ? existingIds.Max() + 1 : 1).ToString();

        var bookmarkStart = new BookmarkStart { Id = bkId, Name = bkName };
        var bookmarkEnd = new BookmarkEnd { Id = bkId };

        if (properties.TryGetValue("text", out var bkText))
        {
            // Try to find existing runs whose concatenated text contains the bookmark text
            var runs = parent.Elements<Run>().ToList();
            var wrapped = TryWrapExistingRunsWithBookmark(parent, runs, bkText, bookmarkStart, bookmarkEnd);
            if (!wrapped)
            {
                // No matching text found — create a new run as fallback
                parent.AppendChild(bookmarkStart);
                parent.AppendChild(new Run(new Text(bkText) { Space = SpaceProcessingModeValues.Preserve }));
                parent.AppendChild(bookmarkEnd);
            }
        }
        else
        {
            parent.AppendChild(bookmarkStart);
            parent.AppendChild(bookmarkEnd);
        }

        var resultPath = $"{parentPath}/bookmark[{bkName}]";
        return resultPath;
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
        var hasUrl = properties.TryGetValue("url", out var hlUrl) || properties.TryGetValue("href", out hlUrl);
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

        if (index.HasValue)
            hlPara.InsertAt(hyperlink, index.Value);
        else
            hlPara.AppendChild(hyperlink);

        var hlCount = hlPara.Elements<Hyperlink>().Count();
        var resultPath = $"{parentPath}/hyperlink[{hlCount}]";
        return resultPath;
    }

    private string AddField(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
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
        var fieldInstr = effectiveType switch
        {
            "pagenum" or "pagenumber" or "page" => " PAGE ",
            "numpages" => " NUMPAGES ",
            "date" => " DATE \\@ \"yyyy-MM-dd\" ",
            "author" => " AUTHOR ",
            "title" => " TITLE ",
            "subject" => " SUBJECT ",
            "filename" => " FILENAME ",
            "time" => " TIME ",
            _ => properties.ContainsKey("instruction")
                ? properties["instruction"]
                : throw new ArgumentException($"Unknown field type '{effectiveType}'. Provide a known type or an 'instruction' property.")
        };
        // Allow override via property
        if (properties.TryGetValue("instruction", out var instr))
            fieldInstr = instr.StartsWith(" ") ? instr : $" {instr} ";

        var fieldPlaceholder = properties.GetValueOrDefault("text", "1");

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
            fieldPara.AppendChild(fieldRunBegin);
            fieldPara.AppendChild(fieldRunInstr);
            fieldPara.AppendChild(fieldRunSep);
            fieldPara.AppendChild(fieldRunResult);
            fieldPara.AppendChild(fieldRunEnd);
            var runIdx = GetAllRuns(fieldPara).Count - 4;
            resultPath = $"{parentPath}/r[{runIdx}]";
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
            AppendToParent(parent, fNewPara);
            var fIdx2 = body.Elements<Paragraph>().TakeWhile(p => p != fNewPara).Count();
            resultPath = $"/body/{BuildParaPathSegment(fNewPara, fIdx2 + 1)}";
        }
        return resultPath;
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
            brkPara.AppendChild(brkRun);
            var brkParaIdx = body.Elements<Paragraph>().TakeWhile(p => p != brkPara).Count();
            resultPath = $"/body/{BuildParaPathSegment(brkPara, brkParaIdx + 1)}/r[{GetAllRuns(brkPara).Count}]";
        }
        else
        {
            // Create a new empty paragraph with the break
            var brkNewPara = new Paragraph(brkRun);
            AppendToParent(parent, brkNewPara);
            var brkIdx = body.Elements<Paragraph>().TakeWhile(p => p != brkNewPara).Count();
            resultPath = $"/body/{BuildParaPathSegment(brkNewPara, brkIdx + 1)}";
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
        var sdtType = ciProps.GetValueOrDefault("sdttype", ciProps.GetValueOrDefault("controltype", "text")).ToLowerInvariant();
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

            ((Paragraph)parent).AppendChild(sdtRun);
            // Build stable @paraId= and @sdtId= based path
            var inlineParaId = ((Paragraph)parent).ParagraphId?.Value;
            var inlineParaSegment = !string.IsNullOrEmpty(inlineParaId)
                ? $"p[@paraId={inlineParaId}]"
                : $"p[{body.Elements<Paragraph>().TakeWhile(p => p != parent).Count() + 1}]";
            resultPath = $"/body/{inlineParaSegment}/sdt[@sdtId={inlineSdtIdVal}]";
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

            AppendToParent(parent, sdtBlock);
            var sdtCount = body.Elements<SdtBlock>().Count();
            resultPath = $"/body/sdt[{sdtCount}]";
        }
        return resultPath;
    }

    private string AddWatermark(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var wmText = properties.GetValueOrDefault("text", "DRAFT");
        // VML watermarks accept named colors (silver, red, etc.) or hex — don't sanitize
        var wmColor = properties.TryGetValue("color", out var wmcVal)
            ? wmcVal.TrimStart('#') : "silver";
        var wmFont = properties.GetValueOrDefault("font", "Calibri");
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
                "Valid types: paragraph (p), run (r), table (tbl), row, cell, picture, chart, equation, comment, section, footnote, endnote, toc, style, watermark, bookmark, hyperlink, field, break, sdt, header, footer. " +
                "Use 'officecli docx add' for details.");

        var siblings = parent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
        var createdIdx = siblings.IndexOf(created) + 1;
        var resultPath = $"{parentPath}/{created.LocalName}[{createdIdx}]";
        return resultPath;
    }
}
