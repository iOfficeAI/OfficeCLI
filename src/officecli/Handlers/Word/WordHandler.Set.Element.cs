// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers extracted from WordHandler.SetElement().
// Each helper owns one element-type's full handling; the entry SetElement
// becomes a thin dispatcher. Mechanically extracted, no behavior change.
public partial class WordHandler
{
    // OOXML CT_TblPPr @w:*FromText is ST_TwipsMeasure (unsigned int) but the
    // SDK exposes the property as `short` so we must reject values that
    // overflow that range before the (short) cast. Pre-fix code silently
    // wrapped 32768 → -32768. Used by all four *FromText Set cases.
    private static short ParseTblpFromTextShort(string value, string key)
    {
        var twips = ParseTwips(value);
        if (twips > 32767)
            throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must be 0..32767 twips (OOXML ST_TwipsMeasure short range).");
        return (short)twips;
    }

    private List<string> SetElementBookmark(BookmarkStart bkStart, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    // Check for duplicate bookmark names
                    var existingBk = _doc.MainDocumentPart?.Document?.Body?
                        .Descendants<BookmarkStart>()
                        .FirstOrDefault(b => b.Name?.Value == value && b != bkStart);
                    if (existingBk != null)
                        throw new ArgumentException($"Bookmark name '{value}' already exists");
                    bkStart.Name = value;
                    break;
                case "text":
                    var bkId = bkStart.Id?.Value;
                    if (bkId != null)
                    {
                        var toRemove = new List<OpenXmlElement>();
                        var sib = bkStart.NextSibling();
                        while (sib != null)
                        {
                            if (sib is BookmarkEnd bkEnd && bkEnd.Id?.Value == bkId)
                                break;
                            toRemove.Add(sib);
                            sib = sib.NextSibling();
                        }
                        foreach (var el in toRemove) el.Remove();
                        bkStart.InsertAfterSelf(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                    }
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        SaveDoc();
        return unsupported;
    }

    private List<string> SetElementComment(Comment comment, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        // Handle text/author/initials/date inline; everything else routes
        // through ApplyCommentFormatKeys (mirrors footnote/endnote fix).
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                {
                    // Replace comment body with a single paragraph/run carrying
                    // the new text. Mirrors AddComment's element shape.
                    comment.RemoveAllChildren();
                    comment.AppendChild(new Paragraph(
                        new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })));
                    break;
                }
                case "author":
                    comment.Author = value;
                    break;
                case "initials":
                    comment.Initials = value;
                    break;
                case "date":
                    comment.Date = DateTime.Parse(value);
                    break;
                case "done":
                case "resolved":
                {
                    // Resolved-state lives in word/commentsExtended.xml (w15:done),
                    // keyed by the comment's first-paragraph w14:paraId.
                    var fp = comment.Descendants<Paragraph>().FirstOrDefault();
                    if (fp != null)
                    {
                        if (string.IsNullOrEmpty(fp.ParagraphId?.Value)) AssignParaId(fp);
                        UpsertCommentEx(fp.ParagraphId!.Value!, null, IsTruthy(value));
                    }
                    break;
                }
            }
        }
        ApplyCommentFormatKeys(comment, properties, unsupported);
        _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments?.Save();
        return unsupported;
    }

    private List<string> SetElementSdt(OpenXmlElement element, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var sdtProps = element is SdtBlock sb
            ? sb.SdtProperties
            : ((SdtRun)element).SdtProperties;
        sdtProps ??= element.PrependChild(new SdtProperties());

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "alias" or "name":
                    var existingAlias = sdtProps.GetFirstChild<SdtAlias>();
                    if (existingAlias != null) existingAlias.Val = value;
                    else sdtProps.AppendChild(new SdtAlias { Val = value });
                    break;
                case "tag":
                    var existingTag = sdtProps.GetFirstChild<Tag>();
                    if (existingTag != null) existingTag.Val = value;
                    else sdtProps.AppendChild(new Tag { Val = value });
                    break;
                case "lock":
                    var existingLock = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
                    var lockEnum = value.ToLowerInvariant() switch
                    {
                        "contentlocked" or "content" => LockingValues.ContentLocked,
                        "sdtlocked" or "sdt" => LockingValues.SdtLocked,
                        "sdtcontentlocked" or "both" => LockingValues.SdtContentLocked,
                        "unlocked" or "none" => LockingValues.Unlocked,
                        _ => throw new ArgumentException($"Invalid lock value: '{value}'. Valid values: unlocked, contentLocked, sdtLocked, sdtContentLocked.")
                    };
                    if (existingLock != null) existingLock.Val = lockEnum;
                    else sdtProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock { Val = lockEnum });
                    break;
                case "text":
                    // Replace content text
                    if (element is SdtBlock sdtB)
                    {
                        var content = sdtB.SdtContentBlock;
                        if (content != null)
                        {
                            content.RemoveAllChildren();
                            content.AppendChild(new Paragraph(
                                new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })));
                        }
                    }
                    else if (element is SdtRun sdtR)
                    {
                        var content = sdtR.SdtContentRun;
                        if (content != null)
                        {
                            content.RemoveAllChildren();
                            content.AppendChild(
                                new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                        }
                    }
                    // Clear showingPlaceholder flag so Word doesn't display as placeholder style
                    var plcHdr = (element is SdtBlock sb2 ? sb2.SdtProperties : ((SdtRun)element).SdtProperties)
                        ?.GetFirstChild<ShowingPlaceholder>();
                    plcHdr?.Remove();
                    break;
                // P1 (sdt post-creation mutation): checkbox state, list choices,
                // combo/dropdown/date current values, and placeholder markers are
                // all settable after the control exists. `type` stays immutable
                // (schema set:false) — changing the content-type element would
                // corrupt the control; recreate it instead. Wrong-control-type
                // targets throw a clear error rather than silently no-op.
                case "checked":
                    SetSdtChecked(element, sdtProps, IsTruthy(value));
                    break;
                case "items" or "choices":
                    SetSdtItems(sdtProps, value);
                    break;
                case "dropdown.lastvalue":
                    RequireSdtType<SdtContentDropDownList>(sdtProps, "dropDown.lastValue", "dropdown").LastValue = value;
                    break;
                case "combobox.lastvalue":
                    RequireSdtType<SdtContentComboBox>(sdtProps, "comboBox.lastValue", "combobox").LastValue = value;
                    break;
                case "format":
                    RequireSdtType<SdtContentDate>(sdtProps, "format", "date").DateFormat = new DateFormat { Val = value };
                    break;
                case "date.fulldate":
                    if (!DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.AdjustToUniversal | System.Globalization.DateTimeStyles.AssumeUniversal,
                            out var fdVal))
                        throw new ArgumentException($"Invalid date.fullDate '{value}'. Expected ISO-8601 (e.g. 2026-01-01T00:00:00Z).");
                    RequireSdtType<SdtContentDate>(sdtProps, "date.fullDate", "date").FullDate = fdVal;
                    break;
                case "date.calendar":
                    RequireSdtType<SdtContentDate>(sdtProps, "date.calendar", "date").Calendar =
                        new Calendar { Val = new EnumValue<CalendarValues>(new CalendarValues(value)) };
                    break;
                case "date.lid":
                    RequireSdtType<SdtContentDate>(sdtProps, "date.lid", "date").LanguageId = new LanguageId { Val = value };
                    break;
                case "date.storemappeddataas":
                    RequireSdtType<SdtContentDate>(sdtProps, "date.storeMappedDataAs", "date").SdtDateMappingType =
                        new SdtDateMappingType { Val = new EnumValue<DateFormatValues>(new DateFormatValues(value)) };
                    break;
                case "placeholder":
                    var existingPlc = sdtProps.GetFirstChild<ShowingPlaceholder>();
                    if (IsTruthy(value))
                    {
                        if (existingPlc == null) InsertSdtPropSchemaOrdered(sdtProps, new ShowingPlaceholder());
                    }
                    else existingPlc?.Remove();
                    break;
                case "placeholdertext":
                    var existingDocPart = sdtProps.GetFirstChild<SdtPlaceholder>();
                    if (string.IsNullOrEmpty(value)) existingDocPart?.Remove();
                    else if (existingDocPart != null)
                        existingDocPart.DocPartReference = new DocPartReference { Val = value };
                    else
                        InsertSdtPropSchemaOrdered(sdtProps,
                            new SdtPlaceholder { DocPartReference = new DocPartReference { Val = value } });
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        SaveDoc();
        return unsupported;
    }

    // P1 sdt helpers. Mirror the Add-side builders (ApplySdtExtraProps /
    // BuildSdtCheckBox / ParseSdtItems in WordHandler.Add.Misc.cs) so Set and
    // Add produce the identical OOXML shape; the difference is mutate-in-place
    // vs build-fresh.

    /// <summary>Fetch the content-type child that a typed prop requires; throw a
    /// clear error when the control is a different variant (type is immutable).</summary>
    private static T RequireSdtType<T>(SdtProperties sdtProps, string prop, string typeName)
        where T : OpenXmlElement
    {
        return sdtProps.GetFirstChild<T>()
            ?? throw new ArgumentException(
                $"'{prop}' applies only to a {typeName} content control; this control is a different type " +
                "(type cannot be changed after creation — recreate the control instead).");
    }

    /// <summary>Insert an sdtPr child before the type-content element per CT_SdtPr
    /// schema order (placeholder / showingPlcHdr / dataBinding precede the type).</summary>
    private static void InsertSdtPropSchemaOrdered(SdtProperties sdtProps, OpenXmlElement el)
    {
        var typeElement = sdtProps.LastChild;
        bool typeIsContent = typeElement is SdtContentDate or SdtContentComboBox
            or SdtContentDropDownList or SdtContentText
            or SdtContentGroup or SdtContentPicture
            or DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox;
        if (typeIsContent) sdtProps.InsertBefore(el, typeElement);
        else sdtProps.AppendChild(el);
    }

    /// <summary>Flip a checkbox control's checked flag and repaint its glyph run
    /// (☒ checked / ☐ unchecked) so `view text` and Word both reflect the state.</summary>
    private void SetSdtChecked(OpenXmlElement element, SdtProperties sdtProps, bool isChecked)
    {
        var checkBox = RequireSdtType<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>(
            sdtProps, "checked", "checkbox");
        checkBox.Checked ??= new DocumentFormat.OpenXml.Office2010.Word.Checked();
        checkBox.Checked.Val = isChecked
            ? DocumentFormat.OpenXml.Office2010.Word.OnOffValues.One
            : DocumentFormat.OpenXml.Office2010.Word.OnOffValues.Zero;

        // Repaint the box glyph only when the content run currently holds a box
        // glyph (☒/☐) — never clobber user-typed content beside the checkbox.
        var glyph = isChecked ? "☒" : "☐";
        var content = (OpenXmlElement?)(element as SdtBlock)?.SdtContentBlock
            ?? (element as SdtRun)?.SdtContentRun;
        foreach (var t in content?.Descendants<Text>() ?? Enumerable.Empty<Text>())
            if (t.Text is "☒" or "☐") t.Text = glyph;
    }

    /// <summary>Replace the list choices on a dropdown/combobox control.</summary>
    private void SetSdtItems(SdtProperties sdtProps, string items)
    {
        OpenXmlElement list = (OpenXmlElement?)sdtProps.GetFirstChild<SdtContentDropDownList>()
            ?? sdtProps.GetFirstChild<SdtContentComboBox>()
            ?? throw new ArgumentException(
                "'items' applies only to a dropdown or combobox content control; this control is a different type " +
                "(type cannot be changed after creation — recreate the control instead).");
        // LastValue (a combo/dropdown attribute) lives on the list element, not a
        // child; RemoveAllChildren clears only the ListItem set, preserving it.
        foreach (var li in list.Elements<ListItem>().ToList()) li.Remove();
        foreach (var li in ParseSdtItems(items)) list.AppendChild(li);
    }

    private List<string> SetElementRun(Run run, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // CONSISTENCY(run-special-content): mirror Get's per-kind type
        // upgrade in WordHandler.Navigation.cs. When a run carries inline
        // structure (ptab/fldChar/instrText/tab/break) instead of <w:t>,
        // expose its settable surface — alignment / fieldCharType / instr
        // / breakType — so audit→fix workflows can correct PAGE→DATE
        // field codes, flip header alignment regions, etc., without
        // dropping to raw-set XML.
        var ptabEl = run.GetFirstChild<PositionalTab>();
        var fldCharEl = run.GetFirstChild<FieldChar>();
        var instrEl = run.GetFirstChild<FieldCode>();
        var breakElInline = run.GetFirstChild<Break>();
        var tabElInline = run.GetFirstChild<TabChar>();
        var hasText = run.GetFirstChild<Text>() != null;
        // CONSISTENCY(run-special-content): mirror the type upgrade in
        // Navigation.cs. `isSpecialRun` (incl. tab / ptab) gates structural
        // rejections like text= injection, which would corrupt a tab/marker
        // run's OOXML. `isMarkerRun` is the narrower set of runs whose rPr
        // paints nothing (fieldChar / instrText / break) — only those reject
        // typography. BUG-DUMP-TABRPR: a tab / positional tab IS a sized
        // inline element (Word stamps font/size/szCs on them to drive line
        // height + leader glyphs), so typography on them is accepted —
        // keeping Set symmetric with Add (which applies it) and Get (which
        // now reads it back). The earlier rejection of tab/ptab typography
        // matched a Get-side strip that was itself wrong for them.
        bool isSpecialRun = ptabEl != null || fldCharEl != null || instrEl != null
                            || (breakElInline != null && !hasText)
                            || (tabElInline != null && !hasText);
        bool isMarkerRun = fldCharEl != null || instrEl != null
                            || (breakElInline != null && !hasText);

        foreach (var (key, value) in properties)
        {
            // CONSISTENCY(run-special-content): typography props (font.* /
            // size / bold / color / underline …) are noise on the MARKER
            // runs (ptab / fieldChar / instrText / break) — their rPr paints
            // nothing, Get strips it, and accepting it on Set would write
            // invisible cruft to <w:rPr>. Reject so the caller sees a clean
            // unsupported notice. A pure-tab run is NOT a marker (it is a
            // sized inline element), so it accepts typography here, matching
            // Add/Get. See BUG-DUMP-TABRPR note above.
            if (isMarkerRun && IsTypographyOnlyKey(key))
            {
                unsupported.Add(key);
                continue;
            }
            // CONSISTENCY(run-prop-helper): rPr-only props delegate to
            // ApplyRunFormatting so the per-property OOXML write logic
            // lives in one place (also used by pmrp / style-run paths);
            // non-rPr cases (text content, image swap, OLE resize, etc.)
            // stay in the inline switch below.
            if (ApplyRunFormatting(EnsureRunProperties(run), key, value))
                continue;
            switch (key.ToLowerInvariant())
            {
                // === run-special-content writes ===
                case "align" or "alignment" when ptabEl != null:
                    ptabEl.Alignment = ParsePtabAlignment(value);
                    break;
                case "relativeto" when ptabEl != null:
                    ptabEl.RelativeTo = ParsePtabRelativeTo(value);
                    break;
                case "leader" when ptabEl != null:
                    ptabEl.Leader = ParsePtabLeader(value);
                    break;
                case "fieldchartype" when fldCharEl != null:
                    fldCharEl.FieldCharType = ParseFieldCharType(value);
                    break;
                case "instr" when instrEl != null:
                    instrEl.Text = value;
                    // CONSISTENCY(field-cache-stale): rewriting a field
                    // instruction (e.g. PAGE → DATE) without invalidating
                    // the cached result run leaves Word displaying the
                    // stale value until the user manually presses F9.
                    // Walk to the owning field's begin <w:fldChar> and set
                    // dirty="true" so Word recomputes the field on next
                    // open. Mirrors Word's own behavior when the user edits
                    // a field code via toggle-codes.
                    MarkOwningFieldDirty(run);
                    break;
                case "breaktype" when breakElInline != null:
                    breakElInline.Type = ParseBreakType(value);
                    break;
                case "text":
                    // Special-content runs have no <w:t> payload — silently
                    // injecting text would corrupt the OOXML structure
                    // (e.g. <w:t> next to <w:instrText> breaks PAGE field
                    // rendering). Reject so the caller sees `unsupported`.
                    if (isSpecialRun)
                    {
                        unsupported.Add(key);
                        break;
                    }
                    OfficeCli.Core.ParseHelpers.ValidateXmlText(value, "text");
                    var textEl = run.GetFirstChild<Text>();
                    if (textEl != null) textEl.Text = value;
                    // CONSISTENCY(field-cache-stale): if this run sits between
                    // a field's `separate` and `end` fldChars, it is the
                    // cached result of the field — Word will recompute it
                    // (overwriting the user's edit) on the next field
                    // refresh. Mark the owning field dirty so Word recomputes
                    // proactively on next open, surfacing the divergence
                    // instead of silently dropping the user's value.
                    if (textEl != null && IsFieldCachedRun(run))
                        MarkOwningFieldDirty(run);
                    break;
                case "alt" or "alttext" or "description":
                    var drawingAlt = run.GetFirstChild<Drawing>();
                    if (drawingAlt != null)
                    {
                        var docPropsAlt = drawingAlt.Descendants<DW.DocProperties>().FirstOrDefault();
                        if (docPropsAlt != null) docPropsAlt.Description = value;
                    }
                    else unsupported.Add(key);
                    break;
                case "decorative":
                    // Accessibility flag stored as an adec:decorative docPr extension.
                    // Symmetric with `alt` (both operate on <wp:docPr>).
                    var drawingDec = run.GetFirstChild<Drawing>();
                    if (drawingDec != null)
                    {
                        var docPropsDec = drawingDec.Descendants<DW.DocProperties>().FirstOrDefault();
                        if (docPropsDec != null)
                        {
                            if (IsTruthy(value)) SetPictureDecorative(docPropsDec);
                            else docPropsDec.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>()?
                                .Elements<A.Extension>()
                                .Where(e => string.Equals(e.Uri?.Value, DecorativeExtUri, StringComparison.OrdinalIgnoreCase))
                                .ToList().ForEach(e => e.Remove());
                        }
                    }
                    else unsupported.Add(key);
                    break;
                case "width":
                {
                    var drawingW = run.GetFirstChild<Drawing>();
                    if (drawingW != null)
                    {
                        var extentW = drawingW.Descendants<DW.Extent>().FirstOrDefault();
                        if (extentW != null) extentW.Cx = ParseEmu(value);
                        var extentsW = drawingW.Descendants<A.Extents>().FirstOrDefault();
                        if (extentsW != null) extentsW.Cx = ParseEmu(value);
                        break;
                    }
                    // OLE run: update VML v:shape style.
                    var oleW = run.GetFirstChild<EmbeddedObject>();
                    var shapeW = oleW?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                    if (shapeW != null)
                    {
                        var styleAttrW = shapeW.GetAttributes().FirstOrDefault(a => a.LocalName == "style");
                        var currentStyleW = styleAttrW.Value ?? "";
                        var ptStrW = (ParseEmu(value) / EmuConverter.EmuPerPointF).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt";
                        var newStyleW = ReplaceVmlStyleDimension(currentStyleW, "width", ptStrW);
                        shapeW.SetAttribute(new OpenXmlAttribute("", "style", "", newStyleW));
                        break;
                    }
                    unsupported.Add(key);
                    break;
                }
                case "height":
                {
                    var drawingH = run.GetFirstChild<Drawing>();
                    if (drawingH != null)
                    {
                        var extentH = drawingH.Descendants<DW.Extent>().FirstOrDefault();
                        if (extentH != null) extentH.Cy = ParseEmu(value);
                        var extentsH = drawingH.Descendants<A.Extents>().FirstOrDefault();
                        if (extentsH != null) extentsH.Cy = ParseEmu(value);
                        break;
                    }
                    // OLE run: update VML v:shape style.
                    var oleH = run.GetFirstChild<EmbeddedObject>();
                    var shapeH = oleH?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                    if (shapeH != null)
                    {
                        var styleAttrH = shapeH.GetAttributes().FirstOrDefault(a => a.LocalName == "style");
                        var currentStyleH = styleAttrH.Value ?? "";
                        var ptStrH = (ParseEmu(value) / EmuConverter.EmuPerPointF).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt";
                        var newStyleH = ReplaceVmlStyleDimension(currentStyleH, "height", ptStrH);
                        shapeH.SetAttribute(new OpenXmlAttribute("", "style", "", newStyleH));
                        break;
                    }
                    unsupported.Add(key);
                    break;
                }
                case "path" or "src":
                {
                    // Replace image source in a run containing a Drawing
                    var drawingSrc = run.GetFirstChild<Drawing>();
                    var blip = drawingSrc?.Descendants<A.Blip>().FirstOrDefault();
                    if (blip != null)
                    {
                        var mainPartImg = _doc.MainDocumentPart!;
                        var (wordImgStream, imgType) = OfficeCli.Core.ImageSource.Resolve(value);
                        using var wordImgDispose = wordImgStream;

                        // Remove old image part(s) to avoid storage bloat —
                        // include the asvg:svgBlip extension part if the
                        // previous image was SVG, otherwise it would be
                        // orphaned in word/media/.
                        var oldEmbedId = blip.Embed?.Value;
                        if (oldEmbedId != null)
                        {
                            try { mainPartImg.DeletePart(oldEmbedId); } catch { }
                        }
                        var oldSvgRelId = OfficeCli.Core.SvgImageHelper.GetSvgRelId(blip);
                        if (oldSvgRelId != null)
                        {
                            try { mainPartImg.DeletePart(oldSvgRelId); } catch { }
                        }

                        if (imgType == ImagePartType.Svg)
                        {
                            // Match AddPicture: SVG part referenced via
                            // extension, raster fallback at r:embed.
                            using var svgBytes = new MemoryStream();
                            wordImgStream.CopyTo(svgBytes);
                            svgBytes.Position = 0;

                            var svgPart = mainPartImg.AddImagePart(ImagePartType.Svg);
                            svgPart.FeedData(svgBytes);
                            var newSvgRelId = mainPartImg.GetIdOfPart(svgPart);

                            var pngPart = mainPartImg.AddImagePart(ImagePartType.Png);
                            pngPart.FeedData(new MemoryStream(
                                OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                            blip.Embed = mainPartImg.GetIdOfPart(pngPart);
                            OfficeCli.Core.SvgImageHelper.AppendSvgExtension(blip, newSvgRelId);
                        }
                        else
                        {
                            var newImgPart = mainPartImg.AddImagePart(imgType);
                            newImgPart.FeedData(wordImgStream);
                            blip.Embed = mainPartImg.GetIdOfPart(newImgPart);
                            // Drop the SVG extension if we replaced an SVG
                            // with a raster image; otherwise Word would
                            // keep rendering the stale SVG reference.
                            if (oldSvgRelId != null)
                            {
                                var extLst = blip.GetFirstChild<A.BlipExtensionList>();
                                if (extLst != null)
                                {
                                    foreach (var ext in extLst.Elements<A.BlipExtension>().ToList())
                                    {
                                        if (string.Equals(ext.Uri?.Value,
                                            OfficeCli.Core.SvgImageHelper.SvgExtensionUri,
                                            StringComparison.OrdinalIgnoreCase))
                                            ext.Remove();
                                    }
                                    if (!extLst.Elements<A.BlipExtension>().Any())
                                        extLst.Remove();
                                }
                            }
                        }
                        break;
                    }

                    // OLE case: run contains an EmbeddedObject. Replace
                    // the backing embedded part and (if needed) update
                    // the ProgID automatically from the new extension.
                    // This is the symmetric counterpart to AddOle — the
                    // part-cleanup rule from the project conventions's Known API
                    // Quirks ("always delete old ImagePart to avoid
                    // storage bloat") applies equally to OLE payloads.
                    var ole = run.GetFirstChild<EmbeddedObject>();
                    if (ole != null)
                    {
                        var mainOle = _doc.MainDocumentPart!;
                        var oleEl = ole.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
                        if (oleEl != null)
                        {
                            var relAttr = oleEl.GetAttributes().FirstOrDefault(a => a.LocalName == "id"
                                && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            var oldRel = relAttr.Value;
                            if (!string.IsNullOrEmpty(oldRel))
                            {
                                try { mainOle.DeletePart(oldRel); } catch { }
                            }
                            var (newEmbedRel, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(mainOle, value, _filePath);
                            // Update r:id attribute in place.
                            oleEl.SetAttribute(new OpenXmlAttribute("r", "id",
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships", newEmbedRel));
                            // Refresh ProgID if it wasn't explicitly pinned by the caller.
                            var newProgId = OfficeCli.Core.OleHelper.DetectProgId(value);
                            OfficeCli.Core.OleHelper.ValidateProgId(newProgId);
                            oleEl.SetAttribute(new OpenXmlAttribute("", "ProgID", "", newProgId));
                        }
                        break;
                    }
                    unsupported.Add(key);
                    break;
                }
                case "progid":
                {
                    // Standalone ProgID override on an existing OLE run.
                    // Mirrors the ProgID-refresh in the "path"/"src" branch
                    // above, but without touching the backing embedded
                    // part. CONSISTENCY(ole-set-progid): PPT and Excel OLE
                    // Set both accept a bare progId key; Word must too.
                    var oleStandalone = run.GetFirstChild<EmbeddedObject>();
                    var oleElStandalone = oleStandalone?.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
                    if (oleElStandalone != null)
                    {
                        OfficeCli.Core.OleHelper.ValidateProgId(value);
                        oleElStandalone.SetAttribute(new OpenXmlAttribute("", "ProgID", "", value));
                        break;
                    }
                    unsupported.Add(key);
                    break;
                }
                case "display":
                {
                    // Update DrawAspect attribute on o:OLEObject.
                    // Strict: only "icon" or "content" are accepted; any
                    // other value throws (see OleHelper.NormalizeOleDisplay).
                    // CONSISTENCY(ole-set-display): mirrors PPT ShowAsIcon toggle.
                    var normalized = OfficeCli.Core.OleHelper.NormalizeOleDisplay(value);
                    var oleDisplay = run.GetFirstChild<EmbeddedObject>();
                    var oleElDisplay = oleDisplay?.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
                    if (oleElDisplay != null)
                    {
                        var drawAspect = normalized == "content" ? "Content" : "Icon";
                        oleElDisplay.SetAttribute(new OpenXmlAttribute("", "DrawAspect", "", drawAspect));
                        break;
                    }
                    unsupported.Add(key);
                    break;
                }
                case "icon":
                {
                    // Empty/whitespace value: treat as unsupported rather
                    // than feeding it into ImageSource.Resolve (which
                    // throws). Matches the gentler unsupported-key pattern
                    // used elsewhere in the Word Set OLE branch.
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        unsupported.Add(key);
                        break;
                    }
                    // Replace the v:imagedata r:id with a new ImagePart, and
                    // delete the old ImagePart to avoid storage bloat
                    // (mirrors Set src cleanup rule in the project conventions Known
                    // API Quirks for picture/blip replacement).
                    var oleIcon = run.GetFirstChild<EmbeddedObject>();
                    var shapeIcon = oleIcon?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                    var imagedata = shapeIcon?.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
                    if (imagedata == null)
                    {
                        unsupported.Add(key);
                        break;
                    }
                    var mainIcon = _doc.MainDocumentPart!;
                    var oldIconRelAttr = imagedata.GetAttributes().FirstOrDefault(a => a.LocalName == "id"
                        && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    if (oldIconRelAttr.Value is string oldIconRel && !string.IsNullOrEmpty(oldIconRel))
                    {
                        try { mainIcon.DeletePart(oldIconRel); } catch { }
                    }
                    var (iconStream, iconPartType) = OfficeCli.Core.ImageSource.Resolve(value);
                    using var iconDispose = iconStream;
                    var newIconPart = mainIcon.AddImagePart(iconPartType);
                    newIconPart.FeedData(iconStream);
                    var newIconRel = mainIcon.GetIdOfPart(newIconPart);
                    imagedata.SetAttribute(new OpenXmlAttribute("r", "id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", newIconRel));
                    break;
                }
                case "wrap":
                {
                    var anchor = ResolveRunAnchor(run);
                    if (anchor == null) { unsupported.Add(key); break; }
                    ReplaceWrapElement(anchor, value);
                    break;
                }
                // CONSISTENCY(pos-alias): accept posH/posV as short aliases for
                // hposition/vposition. posH/posV were
                // silently swallowed (false "Updated" with no XML write); routing
                // them here makes them actually take effect and be read back.
                case "hposition" or "posh":
                {
                    var anchor = ResolveRunAnchor(run);
                    var hPosEl = anchor?.GetFirstChild<DW.HorizontalPosition>();
                    if (hPosEl == null) { unsupported.Add(key); break; }
                    var emu = ParseEmu(value).ToString();
                    var offset = hPosEl.GetFirstChild<DW.PositionOffset>();
                    if (offset != null) offset.Text = emu;
                    else hPosEl.AppendChild(new DW.PositionOffset(emu));
                    break;
                }
                case "vposition" or "posv":
                {
                    var anchor = ResolveRunAnchor(run);
                    var vPosEl = anchor?.GetFirstChild<DW.VerticalPosition>();
                    if (vPosEl == null) { unsupported.Add(key); break; }
                    var emu = ParseEmu(value).ToString();
                    var offset = vPosEl.GetFirstChild<DW.PositionOffset>();
                    if (offset != null) offset.Text = emu;
                    else vPosEl.AppendChild(new DW.PositionOffset(emu));
                    break;
                }
                case "hrelative":
                {
                    var anchor = ResolveRunAnchor(run);
                    var hPosEl = anchor?.GetFirstChild<DW.HorizontalPosition>();
                    if (hPosEl == null) { unsupported.Add(key); break; }
                    hPosEl.RelativeFrom = ParseHorizontalRelative(value);
                    break;
                }
                case "vrelative":
                {
                    var anchor = ResolveRunAnchor(run);
                    var vPosEl = anchor?.GetFirstChild<DW.VerticalPosition>();
                    if (vPosEl == null) { unsupported.Add(key); break; }
                    vPosEl.RelativeFrom = ParseVerticalRelative(value);
                    break;
                }
                case "behindtext":
                {
                    var anchor = ResolveRunAnchor(run);
                    if (anchor == null) { unsupported.Add(key); break; }
                    anchor.BehindDoc = value.Equals("true", StringComparison.OrdinalIgnoreCase);
                    break;
                }
                // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is `url`
                // (per schemas/help/docx/hyperlink.json). `link` / `href` are
                // accepted input aliases.
                case "url":
                case "href":
                case "link":
                {
                    // BUG-FIX(B1): add rel to enclosing host part (header/footer/etc.)
                    var hostPart3 = ResolveHostPart(run);
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove hyperlink wrapper if present
                        if (run.Parent is Hyperlink existingHlNone)
                        {
                            foreach (var childRun in existingHlNone.Elements<Run>().ToList())
                                existingHlNone.InsertBeforeSelf(childRun);
                            existingHlNone.Remove();
                        }
                    }
                    else
                    {
                        // Accept both absolute and relative URIs (Open-XML-SDK supports both).
                        // BUG-DUMP27: fragment-only URIs (e.g. "#_ftn1") are internal-anchor
                        // hyperlinks; mark isExternal=false so .rels TargetMode is omitted.
                        var isAbs = Uri.TryCreate(value, UriKind.Absolute, out var absUri);
                        // CONSISTENCY(hyperlink-scheme-allowlist): only gate
                        // absolute URIs; fragment-only (#_ftn1) and other
                        // relative refs are intra-document and stay open.
                        if (isAbs)
                            Core.HyperlinkUriValidator.RequireSafeScheme(value, key);
                        // OPC requires RFC 3986 encoded URIs in .rels — percent-
                        // encode non-ASCII chars in absolute Uris before write.
                        var uri = isAbs
                            ? new Uri(PercentEncodeUri(value), UriKind.Absolute)
                            : new Uri(value, UriKind.Relative);
                        var isFragment = !string.IsNullOrEmpty(value) && value.StartsWith('#');
                        var newRelId = hostPart3.AddHyperlinkRelationship(uri, isExternal: !isFragment).Id;
                        if (run.Parent is Hyperlink existingHl)
                        {
                            existingHl.Id = newRelId;
                        }
                        else
                        {
                            var newHl = new Hyperlink { Id = newRelId };
                            run.InsertBeforeSelf(newHl);
                            run.Remove();
                            newHl.AppendChild(run);
                        }
                    }
                    break;
                }
                case "textoutline":
                    ApplyW14TextEffect(run, "textOutline", value, BuildW14TextOutline);
                    break;
                case "textfill":
                    ApplyW14TextEffect(run, "textFill", value, BuildW14TextFill);
                    break;
                case "w14shadow":
                    ApplyW14TextEffect(run, "shadow", value, BuildW14Shadow);
                    break;
                case "w14glow":
                    ApplyW14TextEffect(run, "glow", value, BuildW14Glow);
                    break;
                case "w14reflection":
                    ApplyW14TextEffect(run, "reflection", value, BuildW14Reflection);
                    break;
                case "formula":
                {
                    // Replace this run with an inline oMath in the same position
                    var mathContent = FormulaParser.ParseLenient(value, LastUnrecognizedLatex);
                    M.OfficeMath oMath = mathContent is M.OfficeMath dm
                        ? dm : new M.OfficeMath(mathContent.CloneNode(true));
                    run.InsertAfterSelf(oMath);
                    run.Remove();
                    break;
                }
                case "name":
                {
                    // CONSISTENCY(ole-set-name): PPT OLE Set accepts a
                    // bare `name` key that writes oleObj.Name. Word does
                    // not have an equivalent attribute on o:OLEObject
                    // (the VML CT_OleObject complex type has no Name),
                    // so we store the friendly name on the surrounding
                    // v:shape element's "alt" attribute. AddOle writes
                    // to the same attribute and CreateOleNode reads it
                    // back into Format["name"].
                    var oleName = run.GetFirstChild<EmbeddedObject>();
                    var shapeNameEl = oleName?.Descendants().FirstOrDefault(e => e.LocalName == "shape");
                    if (shapeNameEl != null)
                    {
                        shapeNameEl.SetAttribute(new OpenXmlAttribute("", "alt", "", value));
                        break;
                    }
                    unsupported.Add(key);
                    break;
                }
                case "rotation" or "rotate":
                {
                    // Picture rotation: write to a:xfrm/@rot under the inline drawing's pic:spPr.
                    // CONSISTENCY(picture-set-props): mirrors PPTX picture set vocabulary
                    // (PowerPointHandler.Set.Media.cs).
                    var drawingRot = run.GetFirstChild<Drawing>();
                    var spPrPicRot = drawingRot?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties>().FirstOrDefault();
                    if (spPrPicRot != null)
                    {
                        var xfrmRot = spPrPicRot.Transform2D ?? spPrPicRot.AppendChild(new A.Transform2D());
                        xfrmRot.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
                    }
                    else unsupported.Add(key);
                    break;
                }
                case "crop" or "cropleft" or "cropright" or "croptop" or "cropbottom":
                {
                    // BUG-DUMP-CROP: delegate to the shared srcRect writer so Set
                    // and AddPicture stay byte-identical (dump→batch round-trip).
                    var drawingCrop = run.GetFirstChild<Drawing>();
                    var blipFillCrop = drawingCrop?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.BlipFill>().FirstOrDefault();
                    if (blipFillCrop == null) { unsupported.Add(key); break; }
                    ApplyCropToBlipFill(blipFillCrop, key, value);
                    break;
                }
                case "brightness" or "contrast":
                {
                    // Brightness/contrast live in a:lumMod/a:lumOff (luminance modulation
                    // / offset) or a:contrast (effects) on the picture's blip — applied
                    // via a:blip/a:lumMod and a:blip/a:lumOff. Brightness ∈ [-100, 100]
                    // maps to lumOff (positive lightens, negative darkens). Contrast
                    // ∈ [-100, 100] maps to lumMod (>100% boosts contrast, <100% reduces).
                    // For maximum compatibility we encode both via the standard
                    // a:lumMod / a:lumOff pair, matching how PPTX renders these values.
                    var drawingBC = run.GetFirstChild<Drawing>();
                    var blipBC = drawingBC?.Descendants<A.Blip>().FirstOrDefault();
                    if (blipBC == null) { unsupported.Add(key); break; }
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var bcVal)
                        || bcVal < -100 || bcVal > 100)
                        throw new ArgumentException($"Invalid '{key}' value: '{value}'. Expected number in [-100, 100].");

                    // Read existing lumMod/lumOff so brightness and contrast compose.
                    var existingLumMod = blipBC.Elements<A.LuminanceModulation>().FirstOrDefault();
                    var existingLumOff = blipBC.Elements<A.LuminanceOffset>().FirstOrDefault();
                    int curLumModPct = existingLumMod?.Val?.Value is int vm ? vm : 100000;
                    int curLumOffPct = existingLumOff?.Val?.Value is int vo ? vo : 0;

                    if (key.Equals("brightness", StringComparison.OrdinalIgnoreCase))
                        curLumOffPct = (int)(bcVal * 1000); // -100..100 → -100000..100000
                    else
                        curLumModPct = 100000 + (int)(bcVal * 1000); // 0..200 → 0..200000

                    existingLumMod?.Remove();
                    existingLumOff?.Remove();
                    // Schema order: lumMod precedes lumOff inside a:blip.
                    blipBC.AppendChild(new A.LuminanceModulation { Val = curLumModPct });
                    blipBC.AppendChild(new A.LuminanceOffset { Val = curLumOffPct });
                    break;
                }
                default:
                    // OLE runs use a slim prop vocabulary (src, progId,
                    // width, height, alt) that doesn't overlap the rich
                    // run-formatting hint suffix. Emit bare keys to match
                    // PPT/Excel OLE Set. CONSISTENCY(ole-set-bare-key).
                    if (run.GetFirstChild<EmbeddedObject>() != null)
                    {
                        unsupported.Add(key);
                    }
                    else if (key.Contains('.')
                        && Core.TypedAttributeFallback.TrySet(EnsureRunProperties(run), key, value))
                    {
                        // Generic dotted "element.attr=value" fallback
                        // (font.eastAsia, u.color, shd.fill, …). Same helper
                        // as /styles paths — see TypedAttributeFallback for
                        // validation rules and what's intentionally not
                        // covered (composites/lists).
                    }
                    else if (!GenericXmlQuery.TryCreateTypedChild(EnsureRunProperties(run), key, value))
                    {
                        // Help users who reach for paragraph vocabulary on a
                        // run — `align`/`halign`/`alignment` is paragraph- or
                        // ptab-only OOXML, never a run prop. Name the
                        // canonical concept explicitly so the unsupported
                        // notice points the way to the right path.
                        var lk = key.ToLowerInvariant();
                        var isAlignmentKey = lk is "align" or "halign" or "alignment";
                        var detail = isAlignmentKey
                            ? $"{key} (alignment is a paragraph/ptab property, not a run prop)"
                            : (unsupported.Count == 0
                                ? $"{key} (valid run props: text, bold, italic, font, size, color, underline, strike, highlight, caps, smallcaps, superscript, subscript, shading, link, formula)"
                                : key);
                        unsupported.Add(detail);
                    }
                    break;
            }
        }

        var affectedPara = run.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    private List<string> SetElementHyperlink(Hyperlink hl, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            switch (k)
            {
                case "url":
                case "link":
                case "href":
                {
                    // BUG-FIX(B1): add rel to enclosing host part (header/footer/etc.)
                    var hostPartHl = ResolveHostPart(hl);
                    // Delete old relationship to avoid storage bloat. Old rel may
                    // live on a different part (e.g. legacy doc-rooted rel).
                    var oldRelId = hl.Id?.Value;
                    if (oldRelId != null)
                    {
                        var oldRel = ResolveHyperlinkRelationship(hl, oldRelId);
                        if (oldRel?.Container != null)
                            oldRel.Container.DeleteReferenceRelationship(oldRel);
                    }
                    // BUG-DUMP27: fragment-only URIs (e.g. "#_ftn1") are internal-anchor
                    // hyperlinks; mark isExternal=false so .rels TargetMode is omitted.
                    var isAbs = Uri.TryCreate(value, UriKind.Absolute, out var absUri);
                    // CONSISTENCY(hyperlink-scheme-allowlist): only absolute
                    // URIs are scheme-gated; fragment/relative stay open.
                    if (isAbs)
                        Core.HyperlinkUriValidator.RequireSafeScheme(value, k);
                    // OPC: percent-encode non-ASCII in absolute Uris before
                    // writing the .rels target.
                    var uri = isAbs
                        ? new Uri(PercentEncodeUri(value), UriKind.Absolute)
                        : new Uri(value, UriKind.Relative);
                    var isFragment = !string.IsNullOrEmpty(value) && value.StartsWith('#');
                    var newRelId = hostPartHl.AddHyperlinkRelationship(uri, isExternal: !isFragment).Id;
                    hl.Id = newRelId;
                    break;
                }
                case "text":
                {
                    // Update text in all runs within the hyperlink
                    var runs = hl.Elements<Run>().ToList();
                    if (runs.Count > 0)
                    {
                        // Set text on the first run, remove the rest
                        var firstRun = runs[0];
                        var t = firstRun.GetFirstChild<Text>()
                            ?? firstRun.AppendChild(new Text());
                        t.Text = value;
                        t.Space = SpaceProcessingModeValues.Preserve;
                        for (int i = 1; i < runs.Count; i++)
                            runs[i].Remove();
                    }
                    else
                    {
                        // No runs yet, create one
                        var newRun = new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                        hl.AppendChild(newRun);
                    }
                    break;
                }
                default:
                {
                    // BUG-R4B(BUG4): run-level formatting keys (color, bold,
                    // italic, font, size, underline, …) were rejected on a
                    // hyperlink even though the identical key works on a run —
                    // the dump emitter happily emits `color=text1` on a
                    // hyperlink, so replay failed. Route these through the SAME
                    // ApplyRunFormatting path the run handler uses, applied to
                    // every child run, so theme/scheme colors and every other
                    // run property resolve identically. ApplyRunFormatting
                    // returns false for keys it doesn't own → genuinely
                    // unsupported keys still surface.
                    var hlRuns = hl.Descendants<Run>().ToList();
                    if (hlRuns.Count == 0)
                    {
                        // No runs yet — nothing to format. Treat as unsupported
                        // so the user gets feedback (mirrors run-on-empty).
                        unsupported.Add(key);
                        break;
                    }
                    bool appliedAny = false;
                    foreach (var hlRun in hlRuns)
                    {
                        if (ApplyRunFormatting(EnsureRunProperties(hlRun), key, value))
                            appliedAny = true;
                    }
                    if (!appliedAny)
                        unsupported.Add(key);
                    break;
                }
            }
        }

        var affectedPara = hl.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    private List<string> SetElementMPara(M.Paragraph mPara, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            switch (k)
            {
                case "formula":
                {
                    // Clear existing oMath children and rebuild from new formula
                    foreach (var child in mPara.ChildElements.ToList())
                        child.Remove();
                    var mathContent = FormulaParser.ParseLenient(value, LastUnrecognizedLatex);
                    M.OfficeMath oMath = mathContent is M.OfficeMath dm
                        ? dm : new M.OfficeMath(mathContent.CloneNode(true));
                    mPara.AppendChild(oMath);
                    break;
                }
                case "mode":
                {
                    var modeNorm = value.ToLowerInvariant();
                    if (modeNorm == "inline")
                    {
                        // Unwrap m:oMathPara → bare m:oMath inside the host w:p so
                        // the equation renders inline-with-text rather than as a
                        // centered display block.
                        var hostPara = mPara.Ancestors<Paragraph>().FirstOrDefault();
                        var inner = mPara.Elements<M.OfficeMath>().FirstOrDefault();
                        if (hostPara != null && inner != null)
                        {
                            var clone = (M.OfficeMath)inner.CloneNode(true);
                            // Insert into mPara's ACTUAL parent — for a hyperlink-
                            // nested equation that parent is the w:hyperlink, not
                            // hostPara, so hostPara.InsertBefore threw "not a child
                            // of this element". Non-hyperlink: Parent == hostPara.
                            mPara.Parent!.InsertBefore(clone, mPara);
                            mPara.Remove();
                            // R4-bt-1: the equation MOVED (oMathPara → bare
                            // oMath). Report the new resolvable path so the CLI
                            // "Updated …" line points at a path that resolves.
                            LastSetNewPath = ComputeMathElementPath(clone);
                        }
                    }
                    else if (modeNorm == "display")
                    {
                        // Already display — no-op (mPara is m:oMathPara wrapping m:oMath).
                    }
                    else
                    {
                        unsupported.Add($"mode (valid: inline, display)");
                    }
                    break;
                }
                default:
                    unsupported.Add(unsupported.Count == 0
                        ? $"{key} (valid equation props: formula, mode)"
                        : key);
                    break;
            }
        }

        var affectedPara = mPara.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    // An INLINE equation resolves to a bare m:oMath sitting directly inside a
    // w:p (no m:oMathPara wrapper). The display form is m:oMathPara wrapping
    // m:oMath; SetElementMPara owns the m:oMathPara (display) case, so this
    // helper owns the inline case. BUG-BT3: a `set mode=display` on an inline
    // equation previously fell through SetElement with no matching arm,
    // returning an empty unsupported list — the CLI printed "Updated … exit 0"
    // while writing nothing (a silent no-op). Implement inline→display by
    // wrapping the m:oMath in an m:oMathPara in place.
    private List<string> SetElementOMath(M.OfficeMath oMath, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            switch (k)
            {
                case "formula":
                {
                    // Rebuild the equation contents from the new formula, keeping
                    // it inline (mirrors the run-level formula set path).
                    foreach (var child in oMath.ChildElements.ToList())
                        child.Remove();
                    var mathContent = FormulaParser.ParseLenient(value, LastUnrecognizedLatex);
                    if (mathContent is M.OfficeMath parsed)
                        foreach (var c in parsed.ChildElements.ToList())
                            oMath.AppendChild(c.CloneNode(true));
                    else
                        oMath.AppendChild(mathContent.CloneNode(true));
                    break;
                }
                case "mode":
                {
                    var modeNorm = value.ToLowerInvariant();
                    if (modeNorm == "display")
                    {
                        // Wrap the inline m:oMath in an m:oMathPara so it renders
                        // as a centered display block. m:oMathPara is the only
                        // legal parent that gives display placement.
                        var mPara = new M.Paragraph();
                        oMath.InsertBeforeSelf(mPara);
                        oMath.Remove();
                        mPara.AppendChild(oMath);
                        // R4-bt-1: the equation MOVED (bare oMath → oMathPara).
                        // Report the new resolvable path (the wrapping
                        // m:oMathPara is the new target).
                        LastSetNewPath = ComputeMathElementPath(mPara);
                    }
                    else if (modeNorm == "inline")
                    {
                        // Already inline — no-op.
                    }
                    else
                    {
                        unsupported.Add($"mode (valid: inline, display)");
                    }
                    break;
                }
                default:
                    unsupported.Add(unsupported.Count == 0
                        ? $"{key} (valid equation props: formula, mode)"
                        : key);
                    break;
            }
        }

        var affectedPara = oMath.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    // Apply the paragraph-MARK tracked-change revisions the dump carries as the
    // .author/.date/.id namespaces — paraMarkIns (<w:pPr><w:rPr><w:ins/>),
    // paraMarkDel (<w:del/>) and numPrIns (<w:numPr><w:ins/>) — onto a paragraph's
    // pPr, returning the keys it consumed. CONSISTENCY(add-set-symmetry): mirrors
    // the AddParagraph blocks (BUG-DUMP-R44-6 / R49-1). The dump emits these as
    // `set` ops on table-cell paragraphs the table build already created, so the
    // modify path must accept them too or the ¶-mark insertion/deletion (and the
    // tracked numbering assignment) round-trip back as UNSUPPORTED and are dropped.
    private HashSet<string> ApplyParagraphMarkRevisionNamespaces(
        ParagraphProperties pProps, Dictionary<string, string> properties)
    {
        var consumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        static DateTime ParseRev(string? s) =>
            !string.IsNullOrEmpty(s)
            && DateTime.TryParse(s, null, System.Globalization.DateTimeStyles.RoundtripKind, out var d)
                ? d : DateTime.UtcNow;

        if (properties.TryGetValue("paraMarkIns.author", out var iA)
            | properties.TryGetValue("paraMarkIns.date", out var iD)
            | properties.TryGetValue("paraMarkIns.id", out var iI))
        {
            consumed.UnionWith(new[] { "paraMarkIns.author", "paraMarkIns.date", "paraMarkIns.id" });
            var rpr = pProps.ParagraphMarkRunProperties
                      ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            // BUG-DUMP-R71-PARAMARK-INSDEL-ORDER: place via the schema-order
            // helper, not PrependChild. When a paragraph mark carries BOTH ins
            // and del (mark inserted by one reviewer, deleted by another), two
            // blind prepends leave them in execution order (del ends up first);
            // CT_ParaRPr requires ins before del. InsertRunPropInSchemaOrder
            // seats each at its CT_ParaRPr slot regardless of application order.
            if (rpr.GetFirstChild<Inserted>() == null)
                InsertRunPropInSchemaOrder(rpr, new Inserted
                {
                    Author = string.IsNullOrEmpty(iA) ? "OfficeCLI" : iA!,
                    Date = ParseRev(iD),
                    Id = !string.IsNullOrEmpty(iI) ? iI : GenerateRevisionId(),
                });
        }
        if (properties.TryGetValue("paraMarkDel.author", out var dA)
            | properties.TryGetValue("paraMarkDel.date", out var dD)
            | properties.TryGetValue("paraMarkDel.id", out var dI))
        {
            consumed.UnionWith(new[] { "paraMarkDel.author", "paraMarkDel.date", "paraMarkDel.id" });
            var rpr = pProps.ParagraphMarkRunProperties
                      ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            if (rpr.GetFirstChild<Deleted>() == null)
                InsertRunPropInSchemaOrder(rpr, new Deleted
                {
                    Author = string.IsNullOrEmpty(dA) ? "OfficeCLI" : dA!,
                    Date = ParseRev(dD),
                    Id = !string.IsNullOrEmpty(dI) ? dI : GenerateRevisionId(),
                });
        }
        // numPrIns lands inside an existing <w:numPr> (CT_NumPr order: ilvl?, numId?,
        // ins?), so it only applies when the paragraph already carries numbering.
        if (properties.TryGetValue("numPrIns.author", out var nA)
            | properties.TryGetValue("numPrIns.date", out var nD)
            | properties.TryGetValue("numPrIns.id", out var nI))
        {
            consumed.UnionWith(new[] { "numPrIns.author", "numPrIns.date", "numPrIns.id" });
            var numPr = pProps.NumberingProperties;
            if (numPr != null && numPr.GetFirstChild<Inserted>() == null)
                numPr.AppendChild(new Inserted
                {
                    Author = string.IsNullOrEmpty(nA) ? "OfficeCLI" : nA!,
                    Date = ParseRev(nD),
                    Id = !string.IsNullOrEmpty(nI) ? nI : GenerateRevisionId(),
                });
        }
        return consumed;
    }

    private List<string> SetElementParagraph(Paragraph para, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
        // Consume paragraph-mark tracked-change namespaces up front so the per-key
        // loop below doesn't reject them as unsupported (add-set-symmetry).
        var markRevConsumed = ApplyParagraphMarkRevisionNamespaces(pProps, properties);
        // BUG-DUMP-MARKRPR-VERBATIM (Set parity / add-set-symmetry): when the dump
        // emits the WHOLE ¶-mark <w:rPr> verbatim (markRPr.xml), apply it as the
        // authoritative mark rPr ONCE here and skip the redundant per-property
        // markRPr.* dotted keys below. AddParagraph already did this, but the dump
        // round-trips a table-cell paragraph via `set` (the table build creates
        // the paragraph, then sets its props), and Set previously had no
        // markRPr.xml handler — so the key fell into the dotted markrpr.* case as
        // sub="xml" and was dropped, losing every mark-rPr child that no dotted key
        // covers. Most consequentially a paragraph-mark <w:vanish/> (a hidden ¶
        // that merges the paragraph with the next) was lost, so the hidden break
        // re-appeared as a visible blank line and pushed content down.
        bool markRPrVerbatimApplied = false;
        if ((properties.TryGetValue("markRPr.xml", out var setMarkRPrXml)
                || properties.TryGetValue("markrpr.xml", out setMarkRPrXml))
            && !string.IsNullOrEmpty(setMarkRPrXml) && setMarkRPrXml.StartsWith("<"))
        {
            try
            {
                var pmRprVerbatim = new ParagraphMarkRunProperties(setMarkRPrXml);
                pProps.RemoveAllChildren<ParagraphMarkRunProperties>();
                // CT_PPr schema order: ParagraphMarkRunProperties precedes
                // sectPr / pPrChange — insert before the first of those, else append.
                OpenXmlElement? pmSuccessor = pProps.ChildElements
                    .FirstOrDefault(c => c is SectionProperties || c is ParagraphPropertiesChange);
                if (pmSuccessor != null) pmSuccessor.InsertBeforeSelf(pmRprVerbatim);
                else pProps.AppendChild(pmRprVerbatim);
                markRPrVerbatimApplied = true;
            }
            catch { /* malformed fragment — fall back to the dotted keys below */ }
        }
        // CONSISTENCY(markRPr-pre-existed-snapshot): captured ONCE before
        // the property iteration starts. The per-iteration pmrpExisting
        // check inside the bare-key case below otherwise flipped to non-
        // null as soon as the *same* Set call processed an explicit
        // `markRPr.<key>=…` — making every subsequent bare-key iteration
        // mirror to markRPr too, fabricating mark keys the source never
        // had (BUG-DUMP-MARKRPR-LEAK regression form).
        bool markRPrPreExisted = pProps.ParagraphMarkRunProperties != null;
        foreach (var (key, value) in properties)
        {
            if (markRevConsumed.Contains(key)) continue;
            var k = key.ToLowerInvariant();
            if (ApplyParagraphLevelProperty(pProps, key, value, LastSetWarnings))
            {
                // CONSISTENCY(rtl-cascade): direction toggle stamps the full
                // bidi+markRPr+runs cascade. See WordHandler.I18n.cs.
                if (k is "direction" or "dir" or "bidi")
                    ApplyDirectionCascade(para, ParseDirectionRtl(value));
                // handled by paragraph-level helper
            }
            else switch (k)
            {
                case "tabs" or "tabstops":
                {
                    // `tabs=POS:ALIGN[:LEADER],...` shorthand. Replaces the
                    // entire <w:tabs> strip — see ApplyTabsShorthand in
                    // Helpers.cs for the syntax. CONSISTENCY(add-set-symmetry).
                    var paraPpr = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                    ApplyTabsShorthand(paraPpr, value);
                    break;
                }
                case "formula":
                {
                    // Replace paragraph content with OMML equation in-place
                    foreach (var child in para.ChildElements
                        .Where(c => c is not ParagraphProperties).ToList())
                        child.Remove();
                    var mathContent = FormulaParser.ParseLenient(value, LastUnrecognizedLatex);
                    M.OfficeMath oMath = mathContent is M.OfficeMath dm
                        ? dm : new M.OfficeMath(mathContent.CloneNode(true));
                    para.AppendChild(new M.Paragraph(oMath));
                    break;
                }
                case "liststyle":
                    ApplyListStyle(para, value);
                    break;
                case "start":
                    SetListStartValue(para, ParseHelpers.SafeParseInt(value, "start"));
                    break;
                case "rstyle":
                {
                    // BUG-R6-04 / F-4: Set on paragraph rStyle previously
                    // returned UNSUPPORTED, breaking dump→batch round-trip
                    // for table cell paragraphs that carry character
                    // styles (Set is the natural emit since the cell
                    // paragraph already exists). Mirror AddParagraph:
                    // store on the paragraph mark rPr AND propagate to
                    // all existing runs so visible text picks up the
                    // character style.
                    // BUG-DUMP-R37-1: schema-order-safe creation (rPr before
                    // sectPr) so re-applying a section-break paragraph's mark
                    // formatting validates.
                    var pmrp = EnsureMarkRunProperties(pProps);
                    pmrp.RemoveAllChildren<RunStyle>();
                    pmrp.PrependChild(new RunStyle { Val = value });
                    // BUG-DUMP-R28-RSTYLE-SEED: mirror the bare run-formatting
                    // seed path. The dump collapses a single-run cell paragraph
                    // into ONE `set` whose `rStyle` may be iterated BEFORE the
                    // `text` key creates the run. On a still-empty paragraph the
                    // Descendants<Run>() loop below then finds no run, so the
                    // character-style binding reached only the ¶ mark and the
                    // visible text rendered without the style (an italic table
                    // placeholder bound via rStyle="…-Italics" lost its slant).
                    // Seed an empty run when this Set will create one; the later
                    // `text` case preserves the seeded run's rPr (RunStyle).
                    var rStyleRuns = para.Descendants<Run>().ToList();
                    if (rStyleRuns.Count == 0
                        && (properties.ContainsKey("text") || properties.ContainsKey("formula")))
                    {
                        var seedRun = new Run(new RunProperties());
                        para.AppendChild(seedRun);
                        rStyleRuns.Add(seedRun);
                    }
                    foreach (var pRun in rStyleRuns)
                    {
                        var pRP = EnsureRunProperties(pRun);
                        pRP.RemoveAllChildren<RunStyle>();
                        pRP.PrependChild(new RunStyle { Val = value });
                    }
                    break;
                }
                // BUG-DUMP9-02: paragraph-mark-only run formatting. The bare
                // `bold`/`color`/`size`/... keys above propagate to every run
                // in the paragraph; `markRPr.*` writes only to the
                // ParagraphMarkRunProperties so the ¶ glyph carries different
                // formatting than its visible runs (matches OOXML pPr/rPr
                // semantics). ApplyRunFormatting consumes the dotted-suffix
                // form by stripping the prefix.
                case var mk when mk.StartsWith("markrpr.", StringComparison.OrdinalIgnoreCase):
                {
                    // Verbatim ¶-mark subtree already applied (markRPr.xml) — the
                    // dotted keys (and the markRPr.xml key itself) are redundant
                    // with it; skip to avoid double-apply / clobbering.
                    if (markRPrVerbatimApplied) break;
                    var sub = key.Substring("markRPr.".Length);
                    var markOnlyRPr = EnsureMarkRunProperties(pProps);
                    // CONSISTENCY(markRPr-explicit-false): the dotted markRPr.*
                    // form is dump-specific (Navigation emits markRPr.bold=false
                    // only when source <w:rPr><w:b w:val="false"/></w:rPr> sits
                    // on the paragraph mark — explicit style override). Preserve
                    // val=false here so round-trip survives. The bare-key path
                    // keeps ApplyRunFormatting's "remove on falsy" contract
                    // intact for interactive `set bold=false`.
                    var subLower = sub.ToLowerInvariant();
                    if (IsExplicitFalseAddOverride(value))
                    {
                        switch (subLower)
                        {
                            case "bold" or "font.bold":
                                markOnlyRPr.RemoveAllChildren<Bold>();
                                InsertRunPropInSchemaOrder(markOnlyRPr, new Bold { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) });
                                break;
                            case "italic" or "font.italic":
                                markOnlyRPr.RemoveAllChildren<Italic>();
                                InsertRunPropInSchemaOrder(markOnlyRPr, new Italic { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) });
                                break;
                            case "bold.cs" or "font.bold.cs" or "boldcs":
                                markOnlyRPr.RemoveAllChildren<BoldComplexScript>();
                                InsertRunPropInSchemaOrder(markOnlyRPr, new BoldComplexScript { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) });
                                break;
                            case "italic.cs" or "font.italic.cs" or "italiccs":
                                markOnlyRPr.RemoveAllChildren<ItalicComplexScript>();
                                InsertRunPropInSchemaOrder(markOnlyRPr, new ItalicComplexScript { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) });
                                break;
                            default:
                                ApplyRunFormatting(markOnlyRPr, sub, value);
                                break;
                        }
                    }
                    else
                    {
                        ApplyRunFormatting(markOnlyRPr, sub, value);
                    }
                    break;
                }
                case "size" or "fontsize" or "font" or "bold" or "italic" or "color" or "highlight" or "underline" or "strike"
                  or "underline.color" or "underlinecolor" or "underlineColor" or "font.underline.color"
                  or "font.latin" or "font.ascii" or "font.hansi" or "font.hAnsi"
                  or "font.ea" or "font.eastasia" or "font.eastasian"
                  or "font.cs" or "font.complexscript" or "font.complex"
                  or "bold.cs" or "italic.cs" or "size.cs"
                  or "font.bold.cs" or "font.italic.cs" or "font.size.cs"
                  or "font.asciitheme" or "font.asciiTheme"
                  or "font.hansitheme" or "font.hAnsiTheme"
                  or "font.eatheme" or "font.eaTheme" or "font.eastasiatheme"
                  or "font.cstheme" or "font.csTheme"
                  // CONSISTENCY(set-para-run-keys): rPr-bound keys that also
                  // belong on the paragraph mark when no runs exist yet.
                  // ApplyRunFormatting handles each individually.
                  or "kern" or "bdr" or "lang" or "lang.latin" or "lang.val"
                  or "lang.ea" or "lang.eastasia" or "lang.cs" or "lang.bidi"
                  // <w:rFonts w:hint> is run-bound: falling through to the
                  // dotted pPr fallback wrote it on the paragraph MARK only,
                  // so a CJK run lost its eastAsia hint (different font
                  // metrics, re-wrapped table rows) while the mark gained a
                  // phantom one. charspacing/charscale are rPr-bound too.
                  or "font.hint" or "charspacing" or "charscale" or "w"
                  // BUG-DUMP-R46-SCAPS: run on/off typography toggles. The
                  // single-run-collapse dump folds these into `set <paragraph>`,
                  // but they were absent from this run-key case and fell through
                  // to the dotted-pPr fallback (applied to the ¶ mark only, never
                  // the visible runs) — so a small-caps / caps / vanish table
                  // header lost its effect on round-trip. ApplyRunFormatting
                  // handles each; route them to the runs like bold/italic. (rtl /
                  // shading omitted — those carry paragraph-level meaning and are
                  // handled by ApplyParagraphLevelProperty / the direction cascade.)
                  or "caps" or "smallcaps" or "vanish" or "dstrike"
                  or "outline" or "shadow" or "emboss" or "imprint"
                  or "noproof" or "superscript" or "subscript":
                    // Apply run-level formatting to all runs in the paragraph.
                    var allParaRuns = para.Descendants<Run>().ToList();
                    // Paragraph-mark run properties (<w:rPr> inside <w:pPr>)
                    // governs the ¶ glyph and is what cursor-at-end inherits
                    // when the user types more text. Four cases:
                    //   * Already has runs: apply to those runs. Skip markRPr
                    //     unless it already exists (avoid fabricating a
                    //     <w:pPr><w:rPr> the source never had — BUG-DUMP-
                    //     MARKRPR-LEAK leaked 4-5 phantom `markRPr.*` keys
                    //     per paragraph on every dump round-trip).
                    //   * Empty paragraph but this Set has `text`/`formula`:
                    //     the run hasn't been created yet (case order may put
                    //     formatting keys before text). Seed an empty Run
                    //     with an rPr and apply formatting there; the later
                    //     `text` case sees a non-empty run list and takes
                    //     its preserve-rPr branch, so the visible text picks
                    //     up the formatting WITHOUT poisoning markRPr.
                    //   * Truly empty paragraph (no runs now, no text/formula
                    //     in this Set): write to markRPr so the formatting is
                    //     visible on the lone ¶ glyph.
                    //   * Already had markRPr: keep mirroring (preserves
                    //     existing-document semantics; we never strip what
                    //     was there).
                    var pmrpExisting = pProps.ParagraphMarkRunProperties;
                    bool willCreateRun = properties.ContainsKey("text")
                                      || properties.ContainsKey("formula");
                    if (allParaRuns.Count == 0 && willCreateRun)
                    {
                        // Seed a placeholder run if not already seeded by an
                        // earlier iteration of this loop on the same Set call.
                        var seedRun = new Run(new RunProperties());
                        para.AppendChild(seedRun);
                        allParaRuns.Add(seedRun);
                    }
                    // CONSISTENCY(markRPr-bare-vs-dotted): when the same Set
                    // carries an explicit markRPr.<key>, the dotted form is
                    // the authoritative mark-only value — the bare key must
                    // propagate to visible runs but NOT overwrite markRPr.
                    // Without this, source's `markRPr.size=12pt + size=15pt`
                    // (¶ glyph at 12pt, runs at 15pt) collapses to 15pt on
                    // both after replay because iteration order isn't stable.
                    bool explicitMarkOverride = properties.ContainsKey($"markRPr.{key}")
                                             || properties.ContainsKey($"markrpr.{key}");
                    // Use the pre-iteration snapshot: only mirror to markRPr
                    // when the source's markRPr existed before this Set started.
                    // A markRPr created mid-loop by an earlier explicit
                    // markRPr.* setter does NOT enable bare-key mirroring.
                    if ((allParaRuns.Count == 0 || markRPrPreExisted) && !explicitMarkOverride)
                    {
                        var markRPr = EnsureMarkRunProperties(pProps);
                        ApplyRunFormatting(markRPr, key, value);
                    }
                    foreach (var pRun in allParaRuns)
                    {
                        var pRunProps = EnsureRunProperties(pRun);
                        ApplyRunFormatting(pRunProps, key, value);
                    }
                    break;
                case "text":
                    // Set text on paragraph: update first run or create one.
                    // CONSISTENCY(text-breaks): route through AppendTextWithBreaks
                    // so \n/\t in value become <w:br/>/<w:tab/>, matching Add behavior.
                    var existingRuns = para.Elements<Run>().ToList();
                    if (existingRuns.Count > 0)
                    {
                        // Preserve RunProperties from first run, drop all prior text/break/tab children.
                        var keepRun = existingRuns[0];
                        var keepRProps = keepRun.RunProperties;
                        keepRun.RemoveAllChildren();
                        if (keepRProps != null)
                            keepRun.AppendChild(keepRProps);
                        AppendTextWithBreaks(keepRun, value);
                        for (int i = 1; i < existingRuns.Count; i++) existingRuns[i].Remove();
                    }
                    else
                    {
                        // Use paragraph mark run properties as default for new run.
                        // EXCEPT when this same Set call carries explicit
                        // markRPr.* keys: those are mark-ONLY by definition
                        // (the dump emits them for a ¶ mark whose formatting
                        // the visible runs deliberately don't share), and the
                        // mark may have been created by an earlier iteration
                        // of this very call — cloning it would leak mark bold
                        // onto the rebuilt text run.
                        bool setHasMarkKeys = properties.Keys.Any(pk =>
                            pk.StartsWith("markRPr.", StringComparison.OrdinalIgnoreCase));
                        var newRun = new Run();
                        var markProps = setHasMarkKeys ? null : pProps.ParagraphMarkRunProperties;
                        if (markProps != null)
                        {
                            var cloned = new RunProperties();
                            foreach (var child in markProps.ChildElements)
                                cloned.AppendChild(child.CloneNode(true));
                            newRun.PrependChild(cloned);
                        }
                        AppendTextWithBreaks(newRun, value);
                        para.AppendChild(newRun);
                    }
                    break;
                case "framepr.w":
                case "framepr.h":
                {
                    // OOXML w:framePr/@w:w and @w:h are ST_TwipsMeasure (unsigned,
                    // MaxInclusive=31680). CONSISTENCY(length-units): accept pt/cm/in
                    // (bare = twips) via SpacingConverter, then enforce the bound on
                    // the resolved twips. Mirrors Add.Text.cs FrameTwips.
                    uint fpUInt;
                    try { fpUInt = OfficeCli.Core.SpacingConverter.ParseWordSpacing(value); }
                    catch (ArgumentException) { throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must resolve to a twips length (bare number, or a pt/cm/in value)."); }
                    if (fpUInt > 31680)
                        throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must resolve to 0..31680 twips (bare number, or a pt/cm/in length).");
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    if (k == "framepr.w") fp.Width = fpUInt.ToString();
                    else fp.Height = fpUInt;
                    break;
                }
                case "framepr.x":
                case "framepr.y":
                {
                    // ST_SignedTwipsMeasure: -31680 <= v <= 31680.
                    int fpSigned;
                    try { fpSigned = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(value); }
                    catch (ArgumentException) { throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must resolve to a twips length (bare number, or a pt/cm/in value)."); }
                    if (fpSigned < -31680 || fpSigned > 31680)
                        throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must resolve to -31680..31680 twips (bare number, or a pt/cm/in length).");
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    if (k == "framepr.x") fp.X = fpSigned.ToString();
                    else fp.Y = fpSigned.ToString();
                    break;
                }
                case "framepr.hspace":
                case "framepr.vspace":
                {
                    // ST_TwipsMeasure unsigned, MaxInclusive=31680.
                    uint fpSp;
                    try { fpSp = OfficeCli.Core.SpacingConverter.ParseWordSpacing(value); }
                    catch (ArgumentException) { throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must resolve to a twips length (bare number, or a pt/cm/in value)."); }
                    if (fpSp > 31680)
                        throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must resolve to 0..31680 twips (bare number, or a pt/cm/in length).");
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    if (k == "framepr.hspace") fp.HorizontalSpace = fpSp.ToString();
                    else fp.VerticalSpace = fpSp.ToString();
                    break;
                }
                case "framepr.wrap":
                {
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    fp.Wrap = value.ToLowerInvariant() switch
                    {
                        "auto"      => TextWrappingValues.Auto,
                        "around"    => TextWrappingValues.Around,
                        "none"      => TextWrappingValues.None,
                        "notbeside" => TextWrappingValues.NotBeside,
                        "through"   => TextWrappingValues.Through,
                        _ => throw new ArgumentException($"Invalid 'framePr.wrap' value: '{value}'. Valid values: auto, around, none, notBeside, through."),
                    };
                    break;
                }
                case "framepr.hanchor":
                {
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    fp.HorizontalPosition = value.ToLowerInvariant() switch
                    {
                        "page"   => HorizontalAnchorValues.Page,
                        "margin" => HorizontalAnchorValues.Margin,
                        "text"   => HorizontalAnchorValues.Text,
                        _ => throw new ArgumentException($"Invalid 'framePr.hAnchor' value: '{value}'. Valid values: page, margin, text."),
                    };
                    break;
                }
                case "framepr.vanchor":
                {
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    fp.VerticalPosition = value.ToLowerInvariant() switch
                    {
                        "page"   => VerticalAnchorValues.Page,
                        "margin" => VerticalAnchorValues.Margin,
                        "text"   => VerticalAnchorValues.Text,
                        _ => throw new ArgumentException($"Invalid 'framePr.vAnchor' value: '{value}'. Valid values: page, margin, text."),
                    };
                    break;
                }
                case "framepr.xalign":
                {
                    // OOXML ST_XAlign enum; mirror Add.Text.cs.
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    fp.XAlign = value.ToLowerInvariant() switch
                    {
                        "left"    => HorizontalAlignmentValues.Left,
                        "center"  => HorizontalAlignmentValues.Center,
                        "right"   => HorizontalAlignmentValues.Right,
                        "inside"  => HorizontalAlignmentValues.Inside,
                        "outside" => HorizontalAlignmentValues.Outside,
                        _ => throw new ArgumentException($"Invalid 'framePr.xAlign' value: '{value}'. Valid values: left, center, right, inside, outside."),
                    };
                    break;
                }
                case "framepr.yalign":
                {
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    fp.YAlign = value.ToLowerInvariant() switch
                    {
                        "inline"  => VerticalAlignmentValues.Inline,
                        "top"     => VerticalAlignmentValues.Top,
                        "center"  => VerticalAlignmentValues.Center,
                        "bottom"  => VerticalAlignmentValues.Bottom,
                        "inside"  => VerticalAlignmentValues.Inside,
                        "outside" => VerticalAlignmentValues.Outside,
                        _ => throw new ArgumentException($"Invalid 'framePr.yAlign' value: '{value}'. Valid values: inline, top, center, bottom, inside, outside."),
                    };
                    break;
                }
                case "framepr.dropcap":
                {
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    fp.DropCap = value.ToLowerInvariant() switch
                    {
                        "none"   => DropCapLocationValues.None,
                        "drop"   => DropCapLocationValues.Drop,
                        "margin" => DropCapLocationValues.Margin,
                        _ => throw new ArgumentException($"Invalid 'framePr.dropCap' value: '{value}'. Valid values: none, drop, margin."),
                    };
                    break;
                }
                case "framepr.lines":
                {
                    // Drop-cap line span: Word UI exposes 1..10; mirror
                    // Add.Text.cs validation so Set rejects the same range.
                    if (!int.TryParse(value, System.Globalization.NumberStyles.Integer,
                            System.Globalization.CultureInfo.InvariantCulture, out var fpLinesInt)
                        || fpLinesInt < 1 || fpLinesInt > 10)
                        throw new ArgumentException($"Invalid 'framePr.lines' value: '{value}'. Must be an integer 1..10 (drop-cap line span).");
                    var fp = pProps.FrameProperties ?? (pProps.FrameProperties = new FrameProperties());
                    fp.Lines = fpLinesInt;
                    break;
                }
                case "ind.firstline":
                {
                    // OOXML w:ind firstLine and hanging are mutually exclusive.
                    // The canonical `firstlineindent` key already cross-clears
                    // hanging; the ind.* form previously fell through to
                    // TypedAttributeFallback which left a stale hanging behind.
                    // Route through ApplyParagraphLevelProperty (its
                    // "firstlineindent" case nulls hanging and uses
                    // SpacingConverter for lenient cm/in/pt/twips input).
                    ApplyParagraphLevelProperty(pProps, "firstlineindent", value, LastSetWarnings);
                    break;
                }
                case "ind.hanging":
                {
                    ApplyParagraphLevelProperty(pProps, "hangingindent", value, LastSetWarnings);
                    break;
                }
                default:
                    // Generic dotted "element.attr=value" fallback first.
                    // Probe pPr (where most paragraph attrs live: ind.*,
                    // shd.*, spacing.*) then pPr→rPr (run-level attrs at
                    // paragraph mark like rFonts.eastAsia).
                    if (key.Contains('.')
                        && Core.TypedAttributeFallback.TrySet(pProps, key, value))
                    {
                        break;
                    }
                    if (key.Contains('.'))
                    {
                        var paraRPr = EnsureMarkRunProperties(pProps);
                        if (Core.TypedAttributeFallback.TrySet(paraRPr, key, value))
                            break;
                        if (paraRPr.ChildElements.Count == 0)
                            paraRPr.Remove();
                    }
                    // Try as a paragraph-level single-val element first; if pPr
                    // rejects it (schema-aware), it may be a RUN-level element
                    // (<w:w> char scale, <w:em> emphasis, …) that landed on a
                    // paragraph Set because the dump collapsed a single-run
                    // paragraph (ShouldCollapseSingleRun). Apply it to the runs
                    // via the same schema-validated run fallback a direct run Set
                    // uses, so such props round-trip without per-property
                    // curation. Genuinely unknown keys fail both → unsupported.
                    if (!GenericXmlQuery.TryCreateTypedChild(pProps, key, value)
                        && !TryApplyBareKeyToParagraphRuns(para, pProps, key, value,
                                markRPrPreExisted,
                                properties.ContainsKey("text") || properties.ContainsKey("formula")))
                    {
                        var hint = key.ToLowerInvariant() is "sectionbreak" or "section_break" or "sectiontype"
                            ? $"{key} (section breaks are added via 'add --type section --prop type=nextPage/evenPage/oddPage/continuous', not as a paragraph property)"
                            : unsupported.Count == 0
                                ? $"{key} (valid paragraph props: text, style, alignment, bold, italic, font, size, color, spaceBefore, spaceAfter, spaceBeforeLines, spaceAfterLines, lineSpacing, indent, liststyle, formula, direction, bidi)"
                                : key;
                        unsupported.Add(hint);
                    }
                    break;
            }
        }

        var affectedPara = para;
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    /// <summary>
    /// Apply a bare single-val element key (no dot) to a paragraph's runs and ¶
    /// mark when it is a RUN-level element that pPr rejected. The schema-aware
    /// <see cref="GenericXmlQuery.TryCreateTypedChild"/> validates each target,
    /// so a genuine paragraph-only or unknown key lands nowhere and the caller
    /// reports it unsupported. Mirrors the run/markRPr targeting of the curated
    /// run-key branch so the dump's single-run collapse round-trips run props
    /// (e.g. &lt;w:w&gt; character scale, &lt;w:em&gt; emphasis) without a
    /// per-property case. Returns true if the element landed on ≥1 target.
    /// </summary>
    private bool TryApplyBareKeyToParagraphRuns(
        Paragraph para, ParagraphProperties pProps, string key, string value,
        bool markRPrPreExisted, bool willCreateRun)
    {
        // Dotted keys are handled by the TypedAttributeFallback path above.
        if (key.Contains('.')) return false;

        var runs = para.Descendants<Run>().ToList();
        if (runs.Count == 0 && willCreateRun)
        {
            var seedRun = new Run(new RunProperties());
            para.AppendChild(seedRun);
            runs.Add(seedRun);
        }

        bool applied = false;
        foreach (var run in runs)
            if (GenericXmlQuery.TryCreateTypedChild(EnsureRunProperties(run), key, value))
                applied = true;

        // Mirror the curated branch: write to the ¶ mark when there are no runs
        // or the source already carried a markRPr.
        if (runs.Count == 0 || markRPrPreExisted)
        {
            var markRPr = EnsureMarkRunProperties(pProps);
            if (GenericXmlQuery.TryCreateTypedChild(markRPr, key, value))
                applied = true;
            else if (markRPr.ChildElements.Count == 0)
                markRPr.Remove();
        }

        return applied;
    }

    // Modify a single TabStop (paragraph tab stop). Supports pos (twips or any
    // SpacingConverter unit), val (TabStopValues enum), leader (TabStopLeader-
    // CharValues enum). Symmetric with AddTab's writer in Add.Text.cs.
    private List<string> SetElementTabStop(TabStop tab, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pos":
                case "position":
                    // Tab positions may be negative (OOXML allows w:pos < 0).
                    // Match the Add path (AddTabStop) which uses signed parsing;
                    // SpacingConverter.ParseWordSpacing enforces a non-negative
                    // guard that is wrong for tab positions (Add/Set asymmetry).
                    tab.Position = ParseSignedTwips(value);
                    break;
                case "val":
                case "type":
                    if (string.IsNullOrEmpty(value))
                    {
                        tab.Val = null;
                    }
                    else
                    {
                        var tabValNorm = value.ToLowerInvariant();
                        var knownTabVals = new[] { "left", "center", "right", "decimal", "bar", "clear", "num", "start", "end" };
                        if (!knownTabVals.Contains(tabValNorm))
                            throw new ArgumentException($"Invalid tab val '{value}'. Valid: {string.Join(", ", knownTabVals)}.");
                        tab.Val = new EnumValue<TabStopValues>(new TabStopValues(tabValNorm));
                    }
                    break;
                case "leader":
                    if (string.IsNullOrEmpty(value))
                    {
                        tab.Leader = null;
                    }
                    else
                    {
                        var leaderNorm = value.ToLowerInvariant();
                        var knownLeaders = new[] { "none", "dot", "heavy", "hyphen", "middledot", "underscore" };
                        if (!knownLeaders.Contains(leaderNorm))
                            throw new ArgumentException($"Invalid tab leader '{value}'. Valid: {string.Join(", ", knownLeaders)}.");
                        tab.Leader = new EnumValue<TabStopLeaderCharValues>(new TabStopLeaderCharValues(leaderNorm));
                    }
                    break;
                default:
                    unsupported.Add(unsupported.Count == 0
                        ? $"{key} (valid tab props: pos, val, leader)"
                        : key);
                    break;
            }
        }
        SaveDoc();
        return unsupported;
    }

    private List<string> SetElementTableCell(TableCell cell, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var tcPr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        string? deferredText = null;

        // BUG-R2-P0-3: gridSpan/colspan must be processed before width because
        // the width case reads tcPr.GridSpan to know how to distribute the new
        // width across the spanned grid cols. If the dict iteration order put
        // width first, gridSpan was still 1 and the merged width was stamped
        // into a single gridCol — corrupting the tblGrid. Pre-sort so gridspan
        // and aliases ("colspan") run before width.
        // CONSISTENCY(set-prop-order): width depends on gridspan; pre-sort.
        var orderedProperties = properties
            .OrderBy(kv =>
            {
                var k = kv.Key.ToLowerInvariant();
                if (k is "gridspan" or "colspan") return 0;
                if (k is "hmerge") return 0; // hmerge also resolves to gridSpan
                if (k is "width") return 1;
                return 2;
            })
            .ToList();

        foreach (var (key, value) in orderedProperties)
        {
            switch (key.ToLowerInvariant())
            {
                case "skipgridsync":
                    // CONSISTENCY(tblgrid-preserve): consumed inline by the
                    // width branch as a side-effect modifier. Recognized
                    // here so dump→batch replay doesn't flag the emitter-
                    // injected skipGridSync=true as UNSUPPORTED.
                    break;
                case "cellrevision.type":
                {
                    // BUG-DUMP-R51-2: re-stamp a cell-level tracked insert/delete
                    // marker (<w:tcPr><w:cellIns>/<w:cellDel>). The author/date/id
                    // arrive as sibling cellRevision.* keys (consumed below).
                    // Mirrors SetElementTableRow's row-level ins/del wrap.
                    string crAuthor = properties.TryGetValue("cellRevision.author", out var cra) && !string.IsNullOrEmpty(cra)
                        ? cra : "OfficeCLI";
                    string? crDateRaw = properties.TryGetValue("cellRevision.date", out var crd) ? crd : null;
                    DateTime crDate = !string.IsNullOrEmpty(crDateRaw)
                        && DateTime.TryParse(crDateRaw, null, System.Globalization.DateTimeStyles.RoundtripKind, out var crdt)
                        ? crdt : DateTime.UtcNow;
                    string crId = properties.TryGetValue("cellRevision.id", out var cri) && !string.IsNullOrEmpty(cri)
                        ? cri : GenerateRevisionId();
                    tcPr.RemoveAllChildren<CellInsertion>();
                    tcPr.RemoveAllChildren<CellDeletion>();
                    if (string.Equals(value, "del", StringComparison.OrdinalIgnoreCase))
                        tcPr.AppendChild(new CellDeletion { Author = crAuthor, Date = crDate, Id = crId });
                    else
                        tcPr.AppendChild(new CellInsertion { Author = crAuthor, Date = crDate, Id = crId });
                    break;
                }
                case "cellrevision.author":
                case "cellrevision.date":
                case "cellrevision.id":
                    // Consumed by the cellrevision.type case (sibling lookups).
                    break;
                case "text":
                    // Defer text handling until after formatting is applied.
                    // Validate up-front though — the deferred run constructor
                    // below uses `new Text(deferredText)` without a guard so
                    // an XML-illegal char would otherwise leak to save time.
                    OfficeCli.Core.ParseHelpers.ValidateXmlText(value, "text");
                    deferredText = value;
                    break;
                case "font":
                case "size":
                case "fontsize":
                case "bold":
                case "italic":
                case "color":
                case "highlight":
                case "underline":
                case "underline.color":
                case "underlinecolor":
                case "underlineColor":
                case "strike":
                    // Apply to all runs in all paragraphs in the cell
                    // CONSISTENCY(run-prop-helper): per-prop OOXML write
                    // logic lives in ApplyRunFormatting; this branch
                    // just fans out across the cell's runs.
                    bool hasRuns = false;
                    foreach (var cellPara in cell.Elements<Paragraph>())
                    {
                        foreach (var cellRun in cellPara.Elements<Run>())
                        {
                            hasRuns = true;
                            ApplyRunFormatting(EnsureRunProperties(cellRun), key, value);
                        }
                    }
                    // If no runs exist, store formatting in
                    // ParagraphMarkRunProperties on first paragraph so a
                    // future inserted run inherits the formatting.
                    // CONSISTENCY(run-prop-helper): same ApplyRunFormatting
                    // helper as the runs branch above — pmrp extends
                    // OpenXmlCompositeElement so it just works.
                    if (!hasRuns)
                    {
                        var fp = cell.Elements<Paragraph>().FirstOrDefault();
                        if (fp == null) { fp = new Paragraph(); cell.AppendChild(fp); }
                        var pPr = fp.ParagraphProperties ?? fp.PrependChild(new ParagraphProperties());
                        var pmrp = EnsureMarkRunProperties(pPr);
                        ApplyRunFormatting(pmrp, key, value);
                    }
                    break;
                case "direction" or "dir" or "bidi":
                {
                    // CONSISTENCY(rtl-cascade): each cell paragraph runs the
                    // full bidi+markRPr+runs cascade. See WordHandler.I18n.cs.
                    bool cellRtl = ParseDirectionRtl(value);
                    foreach (var cellPara in cell.Elements<Paragraph>())
                        ApplyDirectionCascade(cellPara, cellRtl);
                    break;
                }
                case "shd" or "shading" or "fill" or "cellshading":
                    var shdParts = value.Split(';');
                    if (shdParts.Length >= 3 && shdParts[0].Equals("gradient", StringComparison.OrdinalIgnoreCase))
                    {
                        // gradient;startColor;endColor[;angle]  e.g. gradient;FF0000;0000FF;90
                        var startColor = SanitizeHex(shdParts[1]);
                        var endColor = SanitizeHex(shdParts[2]);
                        // Validate color positions don't look like numbers (likely swapped with angle)
                        if (int.TryParse(shdParts[1], out _) && shdParts[1].Length <= 3)
                            throw new ArgumentException($"'{shdParts[1]}' looks like an angle, not a color. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                        if (int.TryParse(shdParts[2], out _) && shdParts[2].Length <= 3)
                            throw new ArgumentException($"'{shdParts[2]}' looks like an angle, not a color. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                        int angleDeg = 180;
                        if (shdParts.Length >= 4)
                        {
                            if (!int.TryParse(shdParts[3], out angleDeg))
                                throw new ArgumentException($"Invalid gradient angle '{shdParts[3]}', expected integer. Format: gradient;STARTCOLOR;ENDCOLOR[;ANGLE]");
                        }
                        ApplyCellGradient(tcPr, startColor, endColor, angleDeg);
                    }
                    else
                    {
                        // Remove any existing gradient
                        RemoveCellGradient(tcPr);
                        // Route through the shared ParseShadingValue so a cell's
                        // themeFill=/themeFillTint=/themeFillShade= theme-linkage
                        // tail round-trips (the dump applies cell shading via this
                        // Set path); the old hand-rolled split dropped it.
                        tcPr.Shading = ParseShadingValue(value);
                    }
                    break;
                case "align" or "alignment" or "halign":
                    var alignVal = ParseJustification(value);
                    // Apply alignment to ALL paragraphs in the cell, not just the first
                    foreach (var cellAlignPara in cell.Elements<Paragraph>())
                    {
                        var cpProps = cellAlignPara.ParagraphProperties ?? cellAlignPara.PrependChild(new ParagraphProperties());
                        cpProps.Justification = new Justification
                        {
                            Val = alignVal
                        };
                    }
                    break;
                case "valign":
                    tcPr.TableCellVerticalAlignment = new TableCellVerticalAlignment
                    {
                        Val = value.ToLowerInvariant() switch
                        {
                            "top" => TableVerticalAlignmentValues.Top,
                            "center" => TableVerticalAlignmentValues.Center,
                            "bottom" => TableVerticalAlignmentValues.Bottom,
                            _ => throw new ArgumentException($"Invalid valign value: '{value}'. Valid values: top, center, bottom.")
                        }
                    };
                    break;
                case "width":
                    // BUG-DUMP6-04: accept "N%" alongside bare twips so dump→batch
                    // round-trips pct cell widths. OOXML stores pct as fifths-of-percent.
                    if (value.EndsWith('%') &&
                        double.TryParse(value.AsSpan(0, value.Length - 1),
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var pctW))
                    {
                        tcPr.TableCellWidth = new TableCellWidth
                        {
                            Width = ((int)Math.Round(pctW * 50)).ToString(),
                            Type = TableWidthUnitValues.Pct
                        };
                    }
                    else if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        tcPr.TableCellWidth = new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto };
                    }
                    // BUG-DUMP-R42-6: type=nil ("no preferred width") is a distinct
                    // CT_TblWidth value the reader now surfaces as "nil". Apply it
                    // faithfully as <w:tcW w:w="0" w:type="nil"/> instead of letting
                    // the dxa branch below reject the zero width (or coerce to 1).
                    else if (string.Equals(value, "nil", StringComparison.OrdinalIgnoreCase))
                    {
                        tcPr.TableCellWidth = new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Nil };
                    }
                    else
                    {
                        // BUG-R4-05: accept unit-qualified widths (cm/in/pt/dxa) in
                        // addition to bare twips. Mirrors the cross-handler width
                        // contract (the project conventions). Strip a trailing "dxa" suffix
                        // (the form Get now emits) so the bare-twips path still works.
                        var rawWidth = value;
                        long? parsedTwips = null;
                        if (rawWidth.EndsWith("dxa", StringComparison.OrdinalIgnoreCase))
                            rawWidth = rawWidth[..^3];
                        else if (rawWidth.EndsWith("cm", StringComparison.OrdinalIgnoreCase)
                                 || rawWidth.EndsWith("in", StringComparison.OrdinalIgnoreCase)
                                 || rawWidth.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                        {
                            // Reuse SpacingConverter — for Word it returns twips.
                            try { parsedTwips = OfficeCli.Core.SpacingConverter.ParseWordSpacing(value); }
                            catch { parsedTwips = null; }
                        }
                        long widthVal;
                        if (parsedTwips.HasValue)
                            widthVal = parsedTwips.Value;
                        else
                            widthVal = ParseHelpers.SafeParseUint(rawWidth, "width");
                        if (widthVal == 0)
                            throw new ArgumentException($"Invalid 'width' value: '{value}'. Must be a positive integer (> 0); zero-width cells are invalid OOXML.");
                        tcPr.TableCellWidth = new TableCellWidth { Width = widthVal.ToString(), Type = TableWidthUnitValues.Dxa };

                        // BUG-R1-P0-1: keep tblGrid in sync — without this, setting a
                        // cell width drifts cell tcW out of agreement with the
                        // gridCol slot it occupies and Word's column-boundary
                        // inference breaks across all other rows. Mirrors the
                        // startCol calculation used by the gridspan branch.
                        // CONSISTENCY(tblgrid-preserve): dump→batch can disable
                        // the tblGrid sync via skipGridSync=true on the tc set
                        // because AddTable wrote authoritative colWidths and
                        // sources are allowed to carry per-cell tcW values that
                        // disagree with the gridCol widths (Word renders cells
                        // by their own tcW; tblGrid is just a layout hint).
                        bool skipGridSync = properties.TryGetValue("skipgridsync", out var sgs)
                                         || properties.TryGetValue("skipGridSync", out sgs);
                        if (!(skipGridSync && IsTruthy(sgs))
                            && cell.Parent is TableRow widthRow
                            && widthRow.Parent is Table widthTbl)
                        {
                            var widthGrid = widthTbl.GetFirstChild<TableGrid>();
                            var widthGridCols = widthGrid?.Elements<GridColumn>().ToList();
                            if (widthGridCols != null && widthGridCols.Count > 0)
                            {
                                int startCol = 0;
                                foreach (var prevTc in widthRow.Elements<TableCell>())
                                {
                                    if (prevTc == cell) break;
                                    startCol += prevTc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                                }
                                var thisSpan = tcPr.GridSpan?.Val?.Value ?? 1;
                                // Distribute the new width across spanned grid cols.
                                // For span=1 just stamp the column directly. For
                                // span>1 spread evenly so the sum still matches.
                                if (startCol < widthGridCols.Count)
                                {
                                    if (thisSpan <= 1)
                                    {
                                        widthGridCols[startCol].Width = widthVal.ToString();
                                    }
                                    else
                                    {
                                        // CONSISTENCY(tblgrid-preserve): when the
                                        // spanned gridCols already sum to widthVal
                                        // (the common dump-replay case — AddTable
                                        // wrote authoritative colWidths and a later
                                        // tc/width=Σ on a span>1 cell is just a
                                        // restatement) leave them alone. Only
                                        // redistribute evenly when the sum disagrees.
                                        long existingSum = 0;
                                        bool allParsed = true;
                                        for (int gi = 0; gi < thisSpan && startCol + gi < widthGridCols.Count; gi++)
                                        {
                                            if (long.TryParse(widthGridCols[startCol + gi].Width?.Value, out var gw))
                                                existingSum += gw;
                                            else { allParsed = false; break; }
                                        }
                                        if (!(allParsed && existingSum == widthVal))
                                        {
                                            int per = (int)(widthVal / (uint)thisSpan);
                                            int remainder = (int)(widthVal - (uint)(per * thisSpan));
                                            for (int gi = 0; gi < thisSpan && startCol + gi < widthGridCols.Count; gi++)
                                            {
                                                var slice = per + (gi == thisSpan - 1 ? remainder : 0);
                                                widthGridCols[startCol + gi].Width = slice.ToString();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "padding":
                {
                    var dxa = ((int)OfficeCli.Core.SpacingConverter.ParseWordSpacing(value)).ToString();
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.TopMargin = new TopMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    mar.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    mar.LeftMargin = new LeftMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    mar.RightMargin = new RightMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.top":
                {
                    // BUG-R1-07: negative w:tcMar values are invalid OOXML.
                    var ptv = (int)OfficeCli.Core.SpacingConverter.ParseWordSpacing(value);
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.TopMargin = new TopMargin { Width = ptv.ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.bottom":
                {
                    var pbv = (int)OfficeCli.Core.SpacingConverter.ParseWordSpacing(value);
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.BottomMargin = new BottomMargin { Width = pbv.ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.left":
                {
                    var plv = (int)OfficeCli.Core.SpacingConverter.ParseWordSpacing(value);
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.LeftMargin = new LeftMargin { Width = plv.ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.right":
                {
                    var prv = (int)OfficeCli.Core.SpacingConverter.ParseWordSpacing(value);
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.RightMargin = new RightMargin { Width = prv.ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "textdirection" or "textdir":
                    // Reuse the shared section parser so the cell path also accepts
                    // the canonical OOXML InnerText forms the dump emits (tbLrV /
                    // tbRlV / lrTbV — the rotated vertical variants). The previous
                    // local switch only accepted the *-r/*-rotated aliases, so a
                    // dump→batch replay of a vertical-text header cell hit
                    // "Invalid textDirection value: 'tbLrV'" and dropped the
                    // orientation, reflowing the whole table.
                    tcPr.TextDirection = new TextDirection { Val = ParseSectionTextDirection(value) };
                    break;
                case "nowrap":
                    tcPr.NoWrap = IsTruthy(value) ? new NoWrap() : null;
                    break;
                case "cnfstyle":
                {
                    // ST_Cnf @val is a binary bitmask string (see ValidateCnfStyleBitmask).
                    if (string.IsNullOrEmpty(value))
                    {
                        tcPr.RemoveAllChildren<ConditionalFormatStyle>();
                        break;
                    }
                    var cnfVal = ValidateCnfStyleBitmask(value);
                    var cnf = tcPr.GetFirstChild<ConditionalFormatStyle>();
                    if (cnf == null)
                    {
                        cnf = new ConditionalFormatStyle { Val = cnfVal };
                        // cnfStyle is rank 0 in CT_TcPr (FIRST child)
                        tcPr.PrependChild(cnf);
                    }
                    else
                    {
                        cnf.Val = cnfVal;
                    }
                    break;
                }
                case "vmerge":
                {
                    // ST_Merge schema only defines "restart" — continuation is bare <w:vMerge/>.
                    // Removal values (none / clear / remove / false / "") strip
                    // the element entirely so the cell stands alone.
                    var vmLower = value.ToLowerInvariant();
                    bool isRestart = vmLower == "restart";
                    // Accept truthy synonyms (true/yes/on/1) for continue —
                    // pre-fix code silently coerced anything non-restart/non-
                    // remove to the bare <w:vMerge/> continuation glyph and
                    // existing tests / dump→batch use vMerge=true to mean it.
                    bool isContinue = vmLower is "continue" or "true" or "yes" or "on" or "1";
                    bool isRemove = vmLower is "none" or "clear" or "remove" or "false" or "0" or "no" or "off" or "";

                    // BUG-R5-table-merge BUG-9: continuation vMerge in the first
                    // row has no restart anchor above it — Word renders the cell
                    // as invisible / repairs the file.
                    // BUG-R4F-01: a first-row <w:vMerge w:val="continue"/> is
                    // nonetheless valid OOXML and Word-tolerated, so a dump→batch
                    // round-trip of such a source must preserve it. Hard-rejecting
                    // dropped the element on replay (round-trip contract violation).
                    // Accept it but surface a warning about the missing restart
                    // anchor instead of throwing.
                    if (isContinue
                        && cell.Parent is TableRow vmRow0
                        && vmRow0.Parent is Table vmTbl0
                        && vmTbl0.Elements<TableRow>().FirstOrDefault() == vmRow0)
                    {
                        LastSetWarnings.Add(
                            "vmerge=continue on a cell in the first row has no restart anchor above it; Word may render the cell as invisible. Use vmerge=restart if a merge was intended.");
                    }

                    if (!isRestart && !isContinue && !isRemove)
                        throw new ArgumentException($"Invalid 'vMerge' value: '{value}'. Valid: restart, continue, none.");
                    if (isRemove)
                        tcPr.VerticalMerge = null;
                    else if (isRestart)
                        tcPr.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart };
                    else // continue (bare <w:vMerge/>)
                        tcPr.VerticalMerge = new VerticalMerge();
                    break;
                }
                case "hmerge":
                    // BUG-R1-P1-8: <w:hMerge> is a legacy DOC binary-compat
                    // attribute that Word *ignores* in DOCX. The OOXML way to
                    // express horizontal merge is <w:gridSpan>. Redirect
                    // hmerge=restart to gridSpan semantics: merge this cell
                    // with the next physical cell (gridSpan = current + next).
                    // hmerge=continue is a no-op (continuation is implicit
                    // when the previous cell carries gridSpan>1).
                    {
                        // Strip any stale legacy hMerge so we never coexist
                        // with the new gridSpan path.
                        tcPr.HorizontalMerge = null;
                        if (value.ToLowerInvariant() == "restart"
                            && cell.Parent is TableRow hmergeRow)
                        {
                            var nextCell = cell.NextSibling<TableCell>();
                            int currentSpan = tcPr.GridSpan?.Val?.Value ?? 1;
                            int nextSpan = nextCell?.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                            int merged = currentSpan + (nextCell != null ? nextSpan : 1);

                            // Cap to row's grid budget so we don't exceed gridCol count.
                            var hmergeTbl = hmergeRow.Parent as Table;
                            var hmergeGridCount = hmergeTbl?.GetFirstChild<TableGrid>()
                                ?.Elements<GridColumn>().Count() ?? merged;
                            int startCol = 0;
                            foreach (var prevTc in hmergeRow.Elements<TableCell>())
                            {
                                if (prevTc == cell) break;
                                startCol += prevTc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                            }
                            int budget = Math.Max(1, hmergeGridCount - startCol);
                            merged = Math.Min(merged, budget);

                            tcPr.GridSpan = new GridSpan { Val = merged };
                            if (nextCell != null && merged > currentSpan)
                                nextCell.Remove();
                        }
                    }
                    break;
                case var k when k.StartsWith("border"):
                    if (!ApplyCellBorders(tcPr, key, value))
                        unsupported.Add(key);
                    break;
                case "gridspan" or "colspan":
                    var newSpan = ParseHelpers.SafeParseInt(value, "gridspan");
                    if (newSpan <= 0)
                        throw new ArgumentException($"Invalid 'gridspan' value: '{value}'. Must be a positive integer (> 0).");
                    // BUG-R1-03 / BUG-R1-P2-11: reject when gridspan would
                    // exceed the table's grid column count — produces
                    // schema-invalid OOXML and Word repairs the file on open.
                    if (cell.Parent is TableRow gsRow && gsRow.Parent is Table gsTbl)
                    {
                        var gsGridCount = gsTbl.GetFirstChild<TableGrid>()
                            ?.Elements<GridColumn>().Count() ?? 0;
                        if (gsGridCount > 0 && newSpan > gsGridCount)
                            throw new ArgumentException($"Invalid '{key}' value: '{value}'. gridSpan cannot exceed the table's grid column count ({gsGridCount}).");
                        // BUG-R4-table-merge BUG-7: single-cell guard above
                        // misses cumulative overflow — e.g. tc[1] colspan=2 +
                        // tc[2] colspan=2 in a 3-col grid totals 4 slots.
                        // Sum spans of all preceding siblings, then check
                        // startCol + newSpan against gridCount.
                        if (gsGridCount > 0)
                        {
                            int gsStartCol = 0;
                            foreach (var prevTc in gsRow.Elements<TableCell>())
                            {
                                if (ReferenceEquals(prevTc, cell)) break;
                                gsStartCol += prevTc.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
                            }
                            if (gsStartCol + newSpan > gsGridCount)
                                throw new ArgumentException($"Invalid '{key}' value: '{value}'. The row's total gridSpan ({gsStartCol + newSpan}) would exceed the table's grid column count ({gsGridCount}).");
                        }
                    }
                    tcPr.GridSpan = new GridSpan { Val = newSpan };
                    // Ensure the row has the correct number of tc elements.
                    // Calculate total grid columns occupied by all cells in this row,
                    // then remove/add cells so it matches the table grid.
                    if (cell.Parent is TableRow parentRow)
                    {
                        var table = parentRow.Parent as Table;
                        var gridColList = table?.GetFirstChild<TableGrid>()
                            ?.Elements<GridColumn>().ToList();
                        var gridCols = gridColList?.Count ?? 0;
                        if (gridCols > 0)
                        {
                            // Calculate the grid column index where this cell starts
                            int startCol = 0;
                            foreach (var prevTc in parentRow.Elements<TableCell>())
                            {
                                if (prevTc == cell) break;
                                startCol += prevTc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                            }

                            // Update cell width to sum of spanned grid columns
                            int spanWidth = 0;
                            for (int gi = startCol; gi < startCol + newSpan && gi < gridCols; gi++)
                            {
                                if (int.TryParse(gridColList![gi].Width?.Value, out var gw))
                                    spanWidth += gw;
                            }
                            if (spanWidth > 0)
                                tcPr.TableCellWidth = new TableCellWidth { Width = spanWidth.ToString(), Type = TableWidthUnitValues.Dxa };

                            // Calculate total columns occupied by current cells
                            var totalSpan = parentRow.Elements<TableCell>().Sum(tc =>
                                tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                            // Remove excess cells after the current cell
                            while (totalSpan > gridCols)
                            {
                                var nextCell = cell.NextSibling<TableCell>();
                                if (nextCell == null) break;
                                totalSpan -= nextCell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                                nextCell.Remove();
                            }
                            // BUG-R1-table-merge: un-merge (typically newSpan=1
                            // shrinking from a prior larger gridSpan) leaves
                            // the row short of the table's grid column count.
                            // Insert empty placeholder cells immediately after
                            // the anchor so the row matches the grid again.
                            // CONSISTENCY(table-grid-pad): mirrors AddRow grid-
                            // expansion padding in WordHandler.Add.Table.cs.
                            while (totalSpan < gridCols)
                            {
                                var padPara = new Paragraph();
                                AssignParaId(padPara);
                                var padCell = new TableCell(padPara);
                                cell.InsertAfterSelf(padCell);
                                totalSpan += 1;
                            }
                        }
                    }
                    break;
                case "fittext":
                {
                    // FitText goes on w:rPr (RunProperties), not tcPr
                    var cellWidth = tcPr.TableCellWidth?.Width?.Value;
                    var fitVal = cellWidth != null && uint.TryParse(cellWidth, out var fw) ? fw : 0u;
                    foreach (var cellPara in cell.Elements<Paragraph>())
                    {
                        foreach (var cellRun in cellPara.Elements<Run>())
                        {
                            var rPr = EnsureRunProperties(cellRun);
                            rPr.RemoveAllChildren<FitText>();
                            if (IsTruthy(value))
                                rPr.AppendChild(new FitText { Val = fitVal });
                        }
                        // Also apply to ParagraphMarkRunProperties
                        var pPr = cellPara.ParagraphProperties;
                        if (pPr?.ParagraphMarkRunProperties != null)
                        {
                            pPr.ParagraphMarkRunProperties.RemoveAllChildren<FitText>();
                            if (IsTruthy(value))
                                pPr.ParagraphMarkRunProperties.AppendChild(new FitText { Val = fitVal });
                        }
                    }
                    break;
                }
                case "hidemark":
                    // BUG-DUMP-CELLTAIL: <w:hideMark/> toggle (cell marked to
                    // measure its mark for autofit even when empty). Bare
                    // boolean child the generic TryCreateTypedChild can't
                    // synthesise (no @val). Insert in CT_TcPr schema order.
                    tcPr.RemoveAllChildren<HideMark>();
                    if (IsTruthy(value))
                        InsertTcPrChildInOrder(tcPr, new HideMark());
                    break;
                case "tcfittext":
                    // BUG-DUMP-CELLTAIL: <w:tcFitText/> toggle (fit text to
                    // cell width). Distinct from the run-level "fittext" case
                    // above which writes <w:fitText> on rPr — this is the
                    // tcPr-level toggle. Same bare-boolean handling as hideMark.
                    tcPr.RemoveAllChildren<TableCellFitText>();
                    if (IsTruthy(value))
                        InsertTcPrChildInOrder(tcPr, new TableCellFitText());
                    break;
                default:
                    // Generic dotted "element.attr=value" fallback (shd.fill,
                    // tcMar.left, tcBorders.top, …). Same helper as /styles
                    // and paragraph/run paths.
                    if (key.Contains('.')
                        && Core.TypedAttributeFallback.TrySet(tcPr, key, value))
                        break;
                    if (!GenericXmlQuery.TryCreateTypedChild(tcPr, key, value))
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid cell props: text, font, size, bold, italic, color, alignment, valign, width, shd, border, colspan, fitText, textDirection, nowrap, padding)"
                            : key);
                    break;
            }
        }
        // Process deferred "text" AFTER formatting so font/size/bold are applied to existing runs first
        if (deferredText != null)
        {
            var firstPara = cell.Elements<Paragraph>().FirstOrDefault();
            if (firstPara == null)
            {
                firstPara = new Paragraph();
                cell.AppendChild(firstPara);
            }
            // Preserve RunProperties from first run before replacing
            var cellExistingRuns = firstPara.Elements<Run>().ToList();
            var cellRunProps = cellExistingRuns.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
            // Also check ParagraphMarkRunProperties if no run props found
            if (cellRunProps == null)
            {
                var pmrp = firstPara.ParagraphProperties?.ParagraphMarkRunProperties;
                if (pmrp != null) cellRunProps = new RunProperties(pmrp.CloneNode(true).ChildElements.Select(c => c.CloneNode(true)));
            }
            foreach (var r in cellExistingRuns) r.Remove();
            var cellNewRun = new Run(new Text(deferredText) { Space = SpaceProcessingModeValues.Preserve });
            if (cellRunProps != null) cellNewRun.PrependChild(cellRunProps);
            firstPara.AppendChild(cellNewRun);
        }

        var affectedPara = cell.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    // ST_Cnf @val is a conditional-formatting BITMASK string — each character is
    // a binary flag ('0'/'1'), positionally mapped to firstRow, lastRow,
    // firstColumn, lastColumn, oddVBand, evenVBand, oddHBand, evenHBand,
    // firstRowFirstColumn, firstRowLastColumn, lastRowFirstColumn,
    // lastRowLastColumn. It is NOT a hex number: Word writes e.g.
    // "100000000000" for firstRow. Accept 1..12 binary digits (shorter input is
    // right-padded to the canonical 12-bit width so trailing flags default off);
    // reject anything non-binary. Shared by cell (tcPr) and row (trPr) setters.
    private static string ValidateCnfStyleBitmask(string value)
    {
        if (!System.Text.RegularExpressions.Regex.IsMatch(value, "^[01]{1,12}$"))
            throw new ArgumentException(
                $"Invalid cnfStyle '{value}': must be a binary bitmask string of 1..12 '0'/'1' digits " +
                "(ST_Cnf), e.g. '100000000000' for firstRow conditional formatting.");
        return value.Length == 12 ? value : value.PadRight(12, '0');
    }

    // BUG-DUMP-R42-2: canonical CT_TrPr child order. Used to insert
    // gridBefore/wBefore/gridAfter/wAfter at the right rank without disturbing
    // the cnfStyle/cantSplit/trHeight/header/jc children the row setter also
    // manages. Lower rank = earlier child.
    private static int TrPrChildRank(OpenXmlElement e) => e switch
    {
        ConditionalFormatStyle => 0,
        GridBefore => 1,
        WidthBeforeTableRow => 2,
        GridAfter => 3,
        WidthAfterTableRow => 4,
        _ => 100, // cantSplit / trHeight / tblHeader / jc / ins / del / trPrChange — after the skip group
    };

    private static void InsertTrPrChildInOrder(TableRowProperties trPr, OpenXmlElement child)
    {
        int rank = TrPrChildRank(child);
        // Insert before the first existing child whose rank is strictly greater.
        OpenXmlElement? anchor = null;
        foreach (var existing in trPr.ChildElements)
        {
            if (TrPrChildRank(existing) > rank) { anchor = existing; break; }
        }
        if (anchor != null) trPr.InsertBefore(child, anchor);
        else trPr.AppendChild(child);
    }

    // BUG-DUMP-R42-2: parse a row-skip preferred width (wBefore/wAfter) from the
    // same unit-qualified form the reader emits (FormatTableWidth): "nil", "auto",
    // "N%", or "{twips}dxa"/bare twips. Mirrors the cell tcW apply contract.
    private static T BuildRowWidth<T>(string value, string keyName) where T : TableWidthType, new()
    {
        var w = new T();
        if (string.Equals(value, "nil", StringComparison.OrdinalIgnoreCase))
        { w.Width = "0"; w.Type = TableWidthUnitValues.Nil; }
        else if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
        { w.Width = "0"; w.Type = TableWidthUnitValues.Auto; }
        else if (value.EndsWith('%') &&
                 double.TryParse(value.AsSpan(0, value.Length - 1),
                     System.Globalization.NumberStyles.Float,
                     System.Globalization.CultureInfo.InvariantCulture, out var pct))
        { w.Width = ((int)Math.Round(pct * 50)).ToString(); w.Type = TableWidthUnitValues.Pct; }
        else
        {
            var raw = value.EndsWith("dxa", StringComparison.OrdinalIgnoreCase) ? value[..^3] : value;
            w.Width = ParseTwips(raw).ToString();
            w.Type = TableWidthUnitValues.Dxa;
        }
        return w;
    }

    private List<string> SetElementTableRow(TableRow row, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        // BUG-DUMP-R24-4: per-row <w:tblPrEx> (table property exceptions). The
        // value is the verbatim <w:tblPrEx>…</w:tblPrEx> element captured by
        // Navigation.ReadRowProps — parse and re-insert it as the row's FIRST
        // child (CT_Row order: tblPrEx?, trPr?, cells). Handled before trPr
        // creation so it lands ahead of any trPr. Round-trips every CT_TblPrEx
        // child (borders / jc / indent / layout / cellMar / …) losslessly.
        if (properties.TryGetValue("tblPrEx", out var tblPrExXml)
            && !string.IsNullOrWhiteSpace(tblPrExXml))
        {
            row.GetFirstChild<TablePropertyExceptions>()?.Remove();
            var newEx = new TablePropertyExceptions(tblPrExXml);
            row.PrependChild(newEx);
        }
        // CT_Row order is tblPrEx?, trPr?, cells — a new trPr must go AFTER any
        // existing tblPrEx (added by a table-level format-revision cascade), not
        // at the front. PrependChild put trPr before tblPrEx, producing
        // schema-invalid OOXML ("unexpected child element tblPrEx") Word rejects.
        var trPr = row.TableRowProperties;
        if (trPr == null)
        {
            trPr = new TableRowProperties();
            var existingEx = row.GetFirstChild<TablePropertyExceptions>();
            if (existingEx != null) row.InsertAfter(trPr, existingEx);
            else row.PrependChild(trPr);
        }
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "tblprex":
                    // Already applied above as the row's first child; skip the
                    // generic-fallback default so it isn't flagged unsupported.
                    break;
                case "height":
                    // BUG-DUMP-R25-1: bare `height` is AUTO row-sizing — no
                    // @w:hRule (CT_Height default). Previously injected AtLeast,
                    // which the dump round-trip then spuriously re-emitted on
                    // auto-height source rows. Explicit rules use height.atleast
                    // / height.exact.
                    trPr.GetFirstChild<TableRowHeight>()?.Remove();
                    trPr.AppendChild(new TableRowHeight { Val = ParseTwips(value) });
                    break;
                case "height.atleast":
                    trPr.GetFirstChild<TableRowHeight>()?.Remove();
                    trPr.AppendChild(new TableRowHeight { Val = ParseTwips(value), HeightType = HeightRuleValues.AtLeast });
                    break;
                case "height.exact":
                    trPr.GetFirstChild<TableRowHeight>()?.Remove();
                    trPr.AppendChild(new TableRowHeight { Val = ParseTwips(value), HeightType = HeightRuleValues.Exact });
                    break;
                case "header":
                    if (IsTruthy(value))
                    {
                        if (trPr.GetFirstChild<TableHeader>() == null)
                            trPr.AppendChild(new TableHeader());
                    }
                    else
                        trPr.RemoveAllChildren<TableHeader>();
                    break;
                case "cantsplit":
                    if (IsTruthy(value))
                    {
                        if (trPr.GetFirstChild<CantSplit>() == null)
                            trPr.AppendChild(new CantSplit());
                    }
                    else
                        trPr.RemoveAllChildren<CantSplit>();
                    break;
                // BUG-DUMP-R37-3: <w:hidden/> — row not displayed/printed.
                // Mirror the header/cantSplit toggle pattern.
                case "hidden":
                    if (IsTruthy(value))
                    {
                        if (trPr.GetFirstChild<Hidden>() == null)
                            trPr.AppendChild(new Hidden());
                    }
                    else
                        trPr.RemoveAllChildren<Hidden>();
                    break;
                // BUG-DUMP-R62-ROWCELLSPACING: row-level <w:tblCellSpacing> (the
                // inter-cell gap for this row, CT_TrPr). Mirrors the table-level
                // cellspacing case (SetElementTable) but on the row's trPr; same
                // dxa width parse. CT_TblWidth shape — Type must be Dxa for a
                // twips width to take effect.
                case "cellspacing":
                    trPr.RemoveAllChildren<TableCellSpacing>();
                    if (!string.IsNullOrEmpty(value))
                        trPr.AppendChild(new TableCellSpacing
                        {
                            Width = OfficeCli.Core.SpacingConverter.ParseWordSpacing(value).ToString(),
                            Type = TableWidthUnitValues.Dxa
                        });
                    break;
                case "cnfstyle":
                {
                    // ST_Cnf @val bitmask (see ValidateCnfStyleBitmask). cnfStyle
                    // is rank 0 in CT_TrPr (FIRST child), same as CT_TcPr.
                    if (string.IsNullOrEmpty(value))
                    {
                        trPr.RemoveAllChildren<ConditionalFormatStyle>();
                        break;
                    }
                    var cnfVal = ValidateCnfStyleBitmask(value);
                    var cnf = trPr.GetFirstChild<ConditionalFormatStyle>();
                    if (cnf == null)
                        trPr.PrependChild(new ConditionalFormatStyle { Val = cnfVal });
                    else
                        cnf.Val = cnfVal;
                    break;
                }
                // BUG-DUMP-R42-2: leading/trailing grid-column skips. CT_TrPr order
                // is cnfStyle, gridBefore, wBefore, gridAfter, wAfter, cantSplit,
                // trHeight, tblHeader, jc — insert each at its correct rank so the
                // emitted trPr validates regardless of which other children exist.
                case "gridbefore":
                    trPr.RemoveAllChildren<GridBefore>();
                    if (!string.IsNullOrEmpty(value))
                        InsertTrPrChildInOrder(trPr, new GridBefore { Val = ParseHelpers.SafeParseInt(value, "gridBefore") });
                    break;
                case "gridafter":
                    trPr.RemoveAllChildren<GridAfter>();
                    if (!string.IsNullOrEmpty(value))
                        InsertTrPrChildInOrder(trPr, new GridAfter { Val = ParseHelpers.SafeParseInt(value, "gridAfter") });
                    break;
                case "wbefore":
                    trPr.RemoveAllChildren<WidthBeforeTableRow>();
                    if (!string.IsNullOrEmpty(value))
                        InsertTrPrChildInOrder(trPr, BuildRowWidth<WidthBeforeTableRow>(value, "wBefore"));
                    break;
                case "wafter":
                    trPr.RemoveAllChildren<WidthAfterTableRow>();
                    if (!string.IsNullOrEmpty(value))
                        InsertTrPrChildInOrder(trPr, BuildRowWidth<WidthAfterTableRow>(value, "wAfter"));
                    break;
                case "rowalign":
                {
                    // BUG-DUMP-R24-1: row-level <w:jc> (whole-row alignment) in
                    // CT_TrPr. jc ranks after cantSplit/trHeight/tblHeader, so
                    // AppendChild keeps schema order against the children this
                    // setter creates (cnfStyle prepended, height/header appended).
                    trPr.RemoveAllChildren<TableJustification>();
                    if (!string.IsNullOrEmpty(value))
                    {
                        var rowAlignVal = value.ToLowerInvariant() switch
                        {
                            "left" => TableRowAlignmentValues.Left,
                            "center" => TableRowAlignmentValues.Center,
                            "right" => TableRowAlignmentValues.Right,
                            _ => throw new ArgumentException(
                                $"Invalid rowAlign '{value}': must be left, center, or right")
                        };
                        trPr.AppendChild(new TableJustification { Val = rowAlignVal });
                    }
                    break;
                }
                default:
                    // c1, c2, ... shorthand: set text of specific cell by index
                    if (key.Length >= 2 && key[0] == 'c' && int.TryParse(key.AsSpan(1), out var cIdx))
                    {
                        var rowCells = row.Elements<TableCell>().ToList();
                        if (cIdx < 1 || cIdx > rowCells.Count)
                            throw new ArgumentException($"Cell c{cIdx} out of range (row has {rowCells.Count} cells)");
                        var targetPara = rowCells[cIdx - 1].GetFirstChild<Paragraph>()
                            ?? rowCells[cIdx - 1].AppendChild(new Paragraph());
                        targetPara.RemoveAllChildren<Run>();
                        if (!string.IsNullOrEmpty(value))
                        {
                            // CONSISTENCY(escape-sequences): route cell text through
                            // AppendTextWithBreaks so `\n`→<w:br/> and `\t`→<w:tab/>
                            // exactly like --prop text=, instead of storing a literal
                            // backslash-n. The two text-input paths must not diverge.
                            var cellRun = new Run();
                            AppendTextWithBreaks(cellRun, value);
                            targetPara.AppendChild(cellRun);
                        }
                    }
                    else if (key.Contains('.')
                        && Core.TypedAttributeFallback.TrySet(trPr, key, value))
                    {
                        // Generic dotted fallback (e.g. trHeight.* attrs).
                    }
                    else if (!GenericXmlQuery.TryCreateTypedChild(trPr, key, value))
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid row props: height, height.exact, header, cantSplit, hidden, cnfStyle, rowAlign, c1, c2, ...)"
                            : key);
                    break;
            }
        }

        var affectedPara = row.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    private List<string> SetElementTable(Table tbl, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var tblPr = tbl.GetFirstChild<TableProperties>() ?? tbl.PrependChild(new TableProperties());
        foreach (var (rawKey, value) in properties)
        {
            // BUG-R9 (tbllook.* compound key): strip the "tbllook." namespace
            // prefix so callers can write tblLook.firstRow=true alongside the
            // bare `firstRow=true` form. Unknown sub-keys raise instead of
            // being silently dropped (and falsely reporting "Updated"). The
            // bare lookup happens via the lowercased `key` below; we rewrite
            // it here so downstream cases match unchanged.
            var key = rawKey;
            var rkl = rawKey.ToLowerInvariant();
            if (rkl.StartsWith("tbllook."))
            {
                var sub = rkl.Substring("tbllook.".Length);
                if (sub is "firstrow" or "lastrow"
                        or "firstcol" or "firstcolumn"
                        or "lastcol" or "lastcolumn"
                        or "bandrow" or "bandedrows" or "bandrows"
                        or "bandcol" or "bandedcols" or "bandcols"
                        or "nohband" or "nohorizontalband"
                        or "novband" or "noverticalband")
                    key = sub;
                else
                    throw new ArgumentException(
                        $"Unknown tblLook sub-key: '{rawKey}'. Valid sub-keys: " +
                        $"firstRow, lastRow, firstCol, lastCol, bandRow, bandCol, " +
                        $"noHBand, noVBand. Or use the bare hex form tblLook=04A0.");
            }
            switch (key.ToLowerInvariant())
            {
                case "style":
                case "tablestyle":
                case "tablestyleid":
                    // BUG-R3-05: empty/none clears the style — remove element rather
                    // than leave it with an empty Val (which Get would have to filter).
                    if (string.IsNullOrEmpty(value)
                        || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        if (tblPr.TableStyle != null) tblPr.TableStyle.Remove();
                    }
                    else
                    {
                        var tblStyle = tblPr.TableStyle ?? (tblPr.TableStyle = new TableStyle());
                        tblStyle.Val = value;
                    }
                    break;
                case "align" or "alignment":
                    tblPr.TableJustification = new TableJustification
                    {
                        Val = value.ToLowerInvariant() switch
                        {
                            "left" => TableRowAlignmentValues.Left,
                            "center" => TableRowAlignmentValues.Center,
                            "right" => TableRowAlignmentValues.Right,
                            _ => throw new ArgumentException($"Invalid table alignment value: '{value}'. Valid values: left, center, right.")
                        }
                    };
                    break;
                case "width":
                    if (value.EndsWith('%'))
                    {
                        // OOXML pct = percent * 50; double parse keeps fractional
                        // percentages (14.4%) exact, mirroring the cell-width branch.
                        var pct = (int)Math.Round(ParseHelpers.SafeParseDouble(value.TrimEnd('%'), "width") * 50);
                        tblPr.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                    }
                    else
                    {
                        // CONSISTENCY(spacing-units): accept unit-qualified lengths
                        // ('10cm', '5in', '12pt') alongside bare twips, matching
                        // Add and the cross-handler convention from
                        // the project conventions "Spacing input is lenient". Previous
                        // SafeParseUint-only path rejected '10cm'.
                        var twips = OfficeCli.Core.SpacingConverter.ParseWordSpacing(value);
                        tblPr.TableWidth = new TableWidth { Width = twips.ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "indent":
                    // BUG-DUMP-R34-TBLIND: honour a pct-typed indent ("2%") so it
                    // round-trips as <w:tblInd w:type="pct"> rather than collapsing
                    // to dxa twips (which shifts the whole table horizontally).
                    // CONSISTENCY(length-units): non-pct indent accepts pt/cm/in
                    // (bare = twips) via SpacingConverter, matching Add + tc padding.
                    tblPr.TableIndentation = value.TrimEnd().EndsWith("%", StringComparison.Ordinal)
                        ? new TableIndentation { Width = (int)Math.Round(ParseHelpers.SafeParseDouble(value.TrimEnd().TrimEnd('%'), "indent") * 50), Type = TableWidthUnitValues.Pct }
                        : new TableIndentation { Width = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(value), Type = TableWidthUnitValues.Dxa };
                    break;
                case "cellspacing":
                    tblPr.TableCellSpacing = new TableCellSpacing { Width = OfficeCli.Core.SpacingConverter.ParseWordSpacing(value).ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                case "layout":
                    tblPr.TableLayout = new TableLayout
                    {
                        Type = value.ToLowerInvariant() switch
                        {
                            "fixed"   => TableLayoutValues.Fixed,
                            "autofit" or "auto" => TableLayoutValues.Autofit,
                            _ => throw new ArgumentException($"Invalid 'layout' value: '{value}'. Valid values: fixed, autofit."),
                        }
                    };
                    break;
                case "padding":
                {
                    // BUG-R1-07: negative w:tblCellMar values are invalid OOXML.
                    // Lenient input: accepts "5pt"/"0.2cm"/"0.1in" as well as bare
                    // twips. SpacingConverter throws on negatives.
                    var paddingVal = (int)OfficeCli.Core.SpacingConverter.ParseWordSpacing(value);
                    var dxa = paddingVal.ToString();
                    var cm = EnsureTableCellMarginDefault(tblPr);
                    cm.TopMargin = new TopMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    cm.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    break;
                }
                case "shd" or "shading" or "fill" or "cellshading":
                {
                    // BUG-R2-P3-10: table-level shd was falling through to
                    // GenericXmlQuery.TryCreateTypedChild which stamped the
                    // raw color into w:val instead of w:fill. Route through the
                    // shared ParseShadingValue so VAL;FILL;COLOR plus the
                    // themeFill=/themeFillTint=/themeFillShade= theme tail all
                    // round-trip. CONSISTENCY(set-shd-parser).
                    tblPr.Shading = ParseShadingValue(value);
                    break;
                }
                case "tbllook":
                {
                    // BUG-DUMP-R40-5: bare hex tblLook bitmask (w:val). Mirrors
                    // AddTable's `tblLook=04A0` passthrough so the dump→batch
                    // round-trip can carry the exact source bits on a `set table`
                    // step too. The w:val hex is authoritative; we set it
                    // verbatim and leave the decomposed boolean attrs untouched.
                    var tblLook = tblPr.GetFirstChild<TableLook>();
                    if (tblLook == null)
                    {
                        tblLook = new TableLook();
                        InsertTblPrChildInOrder(tblPr, tblLook);
                    }
                    tblLook.Val = value;
                    break;
                }
                case "firstrow":
                case "lastrow":
                case "firstcol" or "firstcolumn":
                case "lastcol" or "lastcolumn":
                case "bandrow" or "bandedrows" or "bandrows":
                case "bandcol" or "bandedcols" or "bandcols":
                case "nohband" or "nohorizontalband":
                case "novband" or "noverticalband":
                {
                    var tblLook = tblPr.GetFirstChild<TableLook>();
                    if (tblLook == null)
                    {
                        // BUG-R3-08: insert tblLook (rank 14) in schema order;
                        // AppendChild placed it AFTER tblCaption (rank 15) /
                        // tblDescription (rank 16) when those existed first.
                        tblLook = new TableLook { Val = "04A0" };
                        InsertTblPrChildInOrder(tblPr, tblLook);
                    }
                    var bv = IsTruthy(value);
                    // "no*Band" carries the inverted sense of band*.
                    var (bit, bitOn) = key.ToLowerInvariant() switch
                    {
                        "firstrow" => (0x0020, bv),
                        "lastrow" => (0x0040, bv),
                        "firstcol" or "firstcolumn" => (0x0080, bv),
                        "lastcol" or "lastcolumn" => (0x0100, bv),
                        "bandrow" or "bandedrows" or "bandrows" => (0x0200, !bv),
                        "bandcol" or "bandedcols" or "bandcols" => (0x0400, !bv),
                        "nohband" or "nohorizontalband" => (0x0200, bv),
                        _ => (0x0400, bv), // novband / noverticalband
                    };
                    switch (bit)
                    {
                        case 0x0020: tblLook.FirstRow = bitOn; break;
                        case 0x0040: tblLook.LastRow = bitOn; break;
                        case 0x0080: tblLook.FirstColumn = bitOn; break;
                        case 0x0100: tblLook.LastColumn = bitOn; break;
                        case 0x0200: tblLook.NoHorizontalBand = bitOn; break;
                        case 0x0400: tblLook.NoVerticalBand = bitOn; break;
                    }
                    // R53-fuzz-3: w:val is the authoritative bitmask — Word
                    // ignores the decomposed boolean attrs when both are
                    // present, so leaving the old hex made the facet Set
                    // silently ineffective in real Word. Recompute the bit.
                    if (int.TryParse(tblLook.Val?.Value ?? "04A0",
                            System.Globalization.NumberStyles.HexNumber,
                            System.Globalization.CultureInfo.InvariantCulture, out var lookBits))
                    {
                        lookBits = bitOn ? lookBits | bit : lookBits & ~bit;
                        tblLook.Val = lookBits.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
                    }
                    break;
                }
                case "position" or "floating":
                {
                    // Shorthand: "floating" or "none" to toggle floating table
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        tblPr.RemoveAllChildren<TablePositionProperties>();
                        tblPr.RemoveAllChildren<TableOverlap>();
                    }
                    else
                    {
                        // "floating" enables floating with defaults
                        var tpp = tblPr.GetFirstChild<TablePositionProperties>();
                        if (tpp == null)
                        {
                            tpp = new TablePositionProperties();
                            // CONSISTENCY(tblpr-schema-order): tblpPr is rank 1.
                            InsertTblPrChildInOrder(tblPr, tpp);
                        }
                        if (tpp.VerticalAnchor == null)
                            tpp.VerticalAnchor = VerticalAnchorValues.Page;
                        if (tpp.HorizontalAnchor == null)
                            tpp.HorizontalAnchor = HorizontalAnchorValues.Page;
                    }
                    break;
                }
                case "position.x" or "tblpx":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    var v = value.ToLowerInvariant();
                    if (v is "left" or "center" or "right" or "inside" or "outside")
                    {
                        tpp.TablePositionXAlignment = v switch
                        {
                            "left" => HorizontalAlignmentValues.Left,
                            "center" => HorizontalAlignmentValues.Center,
                            "right" => HorizontalAlignmentValues.Right,
                            "inside" => HorizontalAlignmentValues.Inside,
                            "outside" => HorizontalAlignmentValues.Outside,
                            _ => throw new ArgumentException($"Invalid position.x alignment: '{value}'")
                        };
                        tpp.TablePositionX = null;
                    }
                    else
                    {
                        tpp.TablePositionX = (int)ParseTwips(value);
                        tpp.TablePositionXAlignment = null;
                    }
                    break;
                }
                case "position.y" or "tblpy":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    var v = value.ToLowerInvariant();
                    if (v is "top" or "center" or "bottom" or "inside" or "outside")
                    {
                        tpp.TablePositionYAlignment = v switch
                        {
                            "top" => VerticalAlignmentValues.Top,
                            "center" => VerticalAlignmentValues.Center,
                            "bottom" => VerticalAlignmentValues.Bottom,
                            "inside" => VerticalAlignmentValues.Inside,
                            "outside" => VerticalAlignmentValues.Outside,
                            _ => throw new ArgumentException($"Invalid position.y alignment: '{value}'")
                        };
                        tpp.TablePositionY = null;
                    }
                    else
                    {
                        tpp.TablePositionY = (int)ParseTwips(value);
                        tpp.TablePositionYAlignment = null;
                    }
                    break;
                }
                case "position.hanchor" or "position.horizontalanchor":
                case "tblp.horzanchor" or "tblp.horizontalanchor":
                case "tblppr.horzanchor" or "tblppr.horizontalanchor":
                {
                    // tblp.horzAnchor is the canonical Get-readback key
                    // (Navigation.cs:2957). Mirror its vocabulary on Set so
                    // dump→batch round-trip works without going through the
                    // generic TypedAttributeFallback (which silently accepts
                    // any string into the StringValue-typed SDK attribute).
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.HorizontalAnchor = value.ToLowerInvariant() switch
                    {
                        "margin" => HorizontalAnchorValues.Margin,
                        "page" => HorizontalAnchorValues.Page,
                        "text" => HorizontalAnchorValues.Text,
                        _ => throw new ArgumentException($"Invalid 'tblp.horzAnchor' value: '{value}'. Valid: margin, page, text.")
                    };
                    break;
                }
                case "position.vanchor" or "position.verticalanchor":
                case "tblp.vertanchor" or "tblp.verticalanchor":
                case "tblppr.vertanchor" or "tblppr.verticalanchor":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.VerticalAnchor = value.ToLowerInvariant() switch
                    {
                        "margin" => VerticalAnchorValues.Margin,
                        "page" => VerticalAnchorValues.Page,
                        "text" => VerticalAnchorValues.Text,
                        _ => throw new ArgumentException($"Invalid 'tblp.vertAnchor' value: '{value}'. Valid: margin, page, text.")
                    };
                    break;
                }
                case "position.leftfromtext" or "position.left":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.LeftFromText = ParseTblpFromTextShort(value, key);
                    break;
                }
                case "position.rightfromtext" or "position.right":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.RightFromText = ParseTblpFromTextShort(value, key);
                    break;
                }
                case "position.topfromtext" or "position.top":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.TopFromText = ParseTblpFromTextShort(value, key);
                    break;
                }
                case "position.bottomfromtext" or "position.bottom":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.BottomFromText = ParseTblpFromTextShort(value, key);
                    break;
                }
                // BUG-DUMP-H89: `tbloverlap.val` is the Get readback key; accept it
                // (and bare `tbloverlap`) as aliases for `overlap` so dump→batch
                // round-trips instead of silently ignoring the emitted key.
                case "overlap":
                case "tbloverlap.val":
                case "tbloverlap":
                {
                    tblPr.RemoveAllChildren<TableOverlap>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var overlapEl = new TableOverlap
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "overlap" or "true" or "always" => TableOverlapValues.Overlap,
                                "never" or "false" => TableOverlapValues.Never,
                                _ => throw new ArgumentException($"Invalid overlap: '{value}'. Valid: overlap, never, none.")
                            }
                        };
                        // CT_TblPr schema: tblStyle → tblpPr → tblOverlap → ...
                        var tppRef = tblPr.GetFirstChild<TablePositionProperties>();
                        if (tppRef != null) tppRef.InsertAfterSelf(overlapEl);
                        else
                        {
                            var styleRef = tblPr.GetFirstChild<TableStyle>();
                            if (styleRef != null) styleRef.InsertAfterSelf(overlapEl);
                            else tblPr.PrependChild(overlapEl);
                        }
                    }
                    break;
                }
                case "caption":
                    tblPr.RemoveAllChildren<TableCaption>();
                    if (!string.IsNullOrEmpty(value))
                        // CONSISTENCY(tblpr-schema-order): tblCaption is rank 15.
                        InsertTblPrChildInOrder(tblPr, new TableCaption { Val = value });
                    break;
                case "description":
                    tblPr.RemoveAllChildren<TableDescription>();
                    if (!string.IsNullOrEmpty(value))
                        // CONSISTENCY(tblpr-schema-order): tblDescription is rank 16.
                        InsertTblPrChildInOrder(tblPr, new TableDescription { Val = value });
                    break;
                case "direction" or "dir" or "bidi":
                {
                    // Table-level bidi: <w:bidiVisual/> on tblPr. CT_TblPrBase
                    // schema: tblStyle → tblpPr → tblOverlap → bidiVisual → ...
                    // Mirrors paragraph/cell direction=rtl vocabulary.
                    // CONSISTENCY(rtl-cascade).
                    tblPr.RemoveAllChildren<BiDiVisual>();
                    if (ParseDirectionRtl(value))
                    {
                        InsertTblPrChildInOrder(tblPr, new BiDiVisual());
                    }
                    break;
                }
                case "bidivisual":
                case "bidivisual.val":
                {
                    // Dotted-form fallback: bidiVisual.val=true. Re-insert in
                    // schema order (must precede tblBorders).
                    tblPr.RemoveAllChildren<BiDiVisual>();
                    var bv = new BiDiVisual();
                    if (key.Equals("bidivisual.val", StringComparison.OrdinalIgnoreCase))
                        bv.Val = IsTruthy(value) ? OnOffOnlyValues.On : OnOffOnlyValues.Off;
                    InsertTblPrChildInOrder(tblPr, bv);
                    break;
                }
                case var k when k.StartsWith("border"):
                    ApplyTableBorders(tblPr, key, value);
                    break;
                case "colwidths" or "colWidths":
                {
                    var parts = value.Split(',');
                    // BUG-R1-01 / BUG-R1-P2-9: reject negative/zero widths
                    // up front. Mirrors Add path validation.
                    foreach (var p in parts)
                    {
                        var trimmed = p.Trim();
                        if (long.TryParse(trimmed, out var pv) && pv <= 0)
                            throw new ArgumentException($"Invalid 'colwidths' value: '{trimmed}'. Each column width must be a positive integer (in twips).");
                    }
                    var tblGrid = tbl.GetFirstChild<TableGrid>();
                    if (tblGrid == null)
                    {
                        tblGrid = new TableGrid();
                        tbl.InsertAfter(tblGrid, tblPr);
                    }
                    var gridCols = tblGrid.Elements<GridColumn>().ToList();
                    // BUG-R1-P1-5 / BUG-R1-04: when fewer values than cols are
                    // supplied, leave the gridCol slots beyond `parts.Length`
                    // untouched. We then re-stamp tcW for ALL cells from the
                    // (possibly-partially-updated) gridCol widths so partial
                    // updates do not leave cells 3,4,… orphaned without tcW.
                    var newGridSnapshot = new List<long>();
                    for (int ci = 0; ci < parts.Length; ci++)
                    {
                        var twips = ParseTwips(parts[ci].Trim());
                        if (ci < gridCols.Count)
                            gridCols[ci].Width = twips.ToString();
                        else
                            tblGrid.AppendChild(new GridColumn { Width = twips.ToString() });
                        // BUG-R1-P1-7: walk cells by GRID column index (accounting
                        // for gridSpan), not by physical cell list index. A
                        // merged cell at the start of a row occupies grid slots
                        // 0..span-1, so the second physical cell maps to grid
                        // index `span`, not `1`. Otherwise rows with merges get
                        // the wrong colWidth stamped.
                        foreach (var tblRow in tbl.Elements<TableRow>())
                        {
                            int gridIdx = 0;
                            foreach (var rc in tblRow.Elements<TableCell>())
                            {
                                var rcSpan = rc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                                if (gridIdx == ci)
                                {
                                    // Only stamp tcW when the cell starts at this
                                    // grid column AND occupies exactly one slot
                                    // (single-span). Multi-span cells should
                                    // sum the spanned widths, not adopt a single
                                    // column's value — leave them untouched here.
                                    if (rcSpan == 1)
                                    {
                                        var rcTcPr = rc.TableCellProperties ?? rc.PrependChild(new TableCellProperties());
                                        rcTcPr.TableCellWidth = new TableCellWidth { Width = twips.ToString(), Type = TableWidthUnitValues.Dxa };
                                    }
                                    break;
                                }
                                if (gridIdx + rcSpan > ci)
                                    break; // cell spans past ci but doesn't start at it; skip
                                gridIdx += rcSpan;
                            }
                        }
                    }
                    // BUG-R1-P1-5 / BUG-R1-04: ensure every single-span cell has
                    // a tcW after the update. Cells touched by the loop above
                    // were stamped from `parts`. Cells beyond parts.Length need
                    // their tcW back-filled from the (untouched) gridCol value
                    // so a partial colWidths update does NOT leave cells 3,4,…
                    // orphaned without a width definition. Multi-span cells
                    // remain untouched — their tcW (if any) is preserved.
                    var gridColsAfter = tblGrid.Elements<GridColumn>().ToList();
                    foreach (var tblRow in tbl.Elements<TableRow>())
                    {
                        int gIdx = 0;
                        foreach (var rc in tblRow.Elements<TableCell>())
                        {
                            var rcSpan = rc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                            if (rcSpan == 1
                                && gIdx >= parts.Length
                                && gIdx < gridColsAfter.Count
                                && rc.TableCellProperties?.TableCellWidth == null)
                            {
                                var rcTcPr = rc.TableCellProperties ?? rc.PrependChild(new TableCellProperties());
                                var gw = gridColsAfter[gIdx].Width?.Value ?? "0";
                                rcTcPr.TableCellWidth = new TableCellWidth { Width = gw, Type = TableWidthUnitValues.Dxa };
                            }
                            gIdx += rcSpan;
                        }
                    }
                    break;
                }
                default:
                    // Generic dotted "element.attr=value" fallback (tblBorders.*,
                    // tblCellMar.*, etc.).
                    if (key.Contains('.')
                        && Core.TypedAttributeFallback.TrySet(tblPr, key, value))
                        break;
                    // Row/cell-scoped props at table level (issue #178): name the
                    // scoped alternative instead of the generic valid-props list,
                    // so an agent can self-correct. Mirrors the Add-side hints in
                    // StyleUnsupportedHints.
                    if (Core.StyleUnsupportedHints.TryGetHint(key, out var scopedHint))
                    {
                        unsupported.Add($"{key} ({scopedHint})");
                        break;
                    }
                    if (!GenericXmlQuery.TryCreateTypedChild(tblPr, key, value))
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid table props: width, alignment, style, indent, cellspacing, layout, padding, border*, colWidths, firstRow, lastRow, firstCol, lastCol, bandedRows, bandedCols, caption, description)"
                            : key);
                    break;
            }
        }

        var affectedPara = tbl.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        SaveDoc();
        return unsupported;
    }

    // /body/textbox[N] resolves to TextBoxContent (the <w:txbxContent> inside
    // <wps:txbx>). Generic SetElement had no case for TextBoxContent, so it
    // returned an empty unsupported list and the CLI reported a silent
    // "Updated" with no XML write — classic false success. We reject the
    // known-confusing position keys explicitly (posH/posV/hposition/vposition)
    // pointing the caller at the path that actually accepts them; everything
    // else still falls through to a proper UNSUPPORTED warning so the silent
    // pass-through is closed for good.
    private List<string> SetElementTextBoxContent(TextBoxContent txbx, Dictionary<string, string> properties)
    {
        // Position keys live on the enclosing <wp:anchor>, not the txbxContent —
        // reject them up-front (unchanged behaviour). Everything else routes to
        // the shared spPr mutator on the wsp that wraps this txbxContent, so the
        // textbox supports the same Set surface (fill/line/width/height/geometry)
        // as a bare shape.
        if (properties.Keys.Any(k => k.ToLowerInvariant() is "hposition" or "posh" or "vposition" or "posv"))
            throw new ArgumentException(
                $"Setting position keys (posH/posV/hposition/vposition) on " +
                $"/body/textbox[N] is not supported (the position lives on the " +
                $"enclosing <wp:anchor>, not on <w:txbxContent>). " +
                $"Add the textbox at the desired position via " +
                $"`add /body --type textbox --prop posH=… posV=…` instead.");

        var wsp = txbx.Ancestors()
            .FirstOrDefault(e => e.LocalName == "wsp"
                && e.NamespaceUri == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        if (wsp == null)
            return properties.Keys.ToList();   // no shape parent → nothing applies
        return SetShapeProps(wsp, properties);
    }

    /// <summary>
    /// Set the curated spPr surface (fill / line / width / height / geometry) on
    /// a <c>wps:wsp</c> shape. Shared by <c>/body/shape[N]</c> (the wsp itself)
    /// and <c>/body/textbox[N]</c> (the wsp wrapping the txbxContent). Reuses the
    /// exact Add-path builders (BuildLineXml, BuildSolidFillXml, SanitizeGeometry,
    /// SanitizeHex, ParseDrawingSize) so Add and Set produce byte-identical XML.
    /// Edits the spPr children in place, preserving CT_ShapeProperties element
    /// order (xfrm, prstGeom, fill, ln). Width/height update BOTH the shape
    /// <c>&lt;a:ext&gt;</c> and the layout <c>&lt;wp:extent&gt;</c>, mirroring Add.
    /// Returns the unsupported keys (out-of-scope props: position/rotation/inset/
    /// text/shadow/gradient/wrap), which the caller forwards as UNSUPPORTED.
    /// </summary>
    private List<string> SetShapeProps(OpenXmlElement wsp, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string WpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

        var spPr = wsp.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
        if (spPr == null)
            return properties.Keys.ToList();   // malformed shape — apply nothing

        OpenXmlElement? AChild(OpenXmlElement parent, string local) =>
            parent.ChildElements.FirstOrDefault(e => e.LocalName == local && e.NamespaceUri == ANs);

        var xfrm = AChild(spPr, "xfrm");
        var prstGeom = AChild(spPr, "prstGeom");

        // Collect the curated keys (split-by-key parsing mirrors AddShape).
        string? fillRaw = properties.GetValueOrDefault("fill") ?? properties.GetValueOrDefault("fillcolor");
        string? geomRaw = properties.GetValueOrDefault("geometry") ?? properties.GetValueOrDefault("preset");
        string? widthRaw = properties.GetValueOrDefault("width");
        string? heightRaw = properties.GetValueOrDefault("height");

        // line: composite "STYLE;SIZE;COLOR" OR split line.style/.width/.color.
        string? lineCompact = properties.GetValueOrDefault("line");
        bool hasLine = lineCompact != null
            || properties.ContainsKey("line.style") || properties.ContainsKey("linestyle")
            || properties.ContainsKey("line.width") || properties.ContainsKey("linewidth")
            || properties.ContainsKey("line.color") || properties.ContainsKey("linecolor");
        string? lineStyle = null, lineWidth = null, lineColor = null;
        if (!string.IsNullOrEmpty(lineCompact)
            && !string.Equals(lineCompact, "none", StringComparison.OrdinalIgnoreCase))
        {
            var parts = lineCompact.Split(';');
            if (parts.Length >= 1 && !string.IsNullOrEmpty(parts[0])) lineStyle = parts[0];
            if (parts.Length >= 2 && !string.IsNullOrEmpty(parts[1])) lineWidth = parts[1];
            if (parts.Length >= 3 && !string.IsNullOrEmpty(parts[2])) lineColor = parts[2];
        }
        else if (string.Equals(lineCompact, "none", StringComparison.OrdinalIgnoreCase))
        {
            lineStyle = "none";
        }
        lineStyle ??= properties.GetValueOrDefault("line.style") ?? properties.GetValueOrDefault("linestyle");
        lineWidth ??= properties.GetValueOrDefault("line.width") ?? properties.GetValueOrDefault("linewidth");
        lineColor ??= properties.GetValueOrDefault("line.color") ?? properties.GetValueOrDefault("linecolor");

        // --- geometry → replace prstGeom@prst ---
        if (geomRaw != null && prstGeom != null)
            prstGeom.SetAttribute(new OpenXmlAttribute("prst", "", SanitizeGeometry(geomRaw)));

        // --- width / height → update <a:ext cx cy> and <wp:extent cx cy> ---
        if ((widthRaw != null || heightRaw != null) && xfrm != null)
        {
            var ext = AChild(xfrm, "ext");
            // Anchor wrapper carries the matching <wp:extent>.
            var wpExtent = wsp.Ancestors()
                .FirstOrDefault(e => e.LocalName == "anchor" || e.LocalName == "inline")
                ?.Descendants().FirstOrDefault(e => e.LocalName == "extent" && e.NamespaceUri == WpNs);
            if (widthRaw != null)
            {
                long cx = ParseDrawingSize(widthRaw, ReadUnqualifiedLong(ext, "cx") ?? 914_400);
                ext?.SetAttribute(new OpenXmlAttribute("cx", "", cx.ToString()));
                wpExtent?.SetAttribute(new OpenXmlAttribute("cx", "", cx.ToString()));
            }
            if (heightRaw != null)
            {
                long cy = ParseDrawingSize(heightRaw, ReadUnqualifiedLong(ext, "cy") ?? 914_400);
                ext?.SetAttribute(new OpenXmlAttribute("cy", "", cy.ToString()));
                wpExtent?.SetAttribute(new OpenXmlAttribute("cy", "", cy.ToString()));
            }
        }

        // --- fill → replace solidFill / noFill / gradFill ---
        if (fillRaw != null)
        {
            string fillXml =
                string.IsNullOrEmpty(fillRaw) || string.Equals(fillRaw, "none", StringComparison.OrdinalIgnoreCase)
                    ? "<a:noFill/>"
                    : $"<a:solidFill><a:srgbClr val=\"{SanitizeHex(fillRaw)}\"/></a:solidFill>";
            // Remove any existing fill element.
            foreach (var f in spPr.ChildElements
                .Where(e => e.NamespaceUri == ANs
                    && e.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill")
                .ToList())
                f.Remove();
            var newFill = ParseShapeFragment(fillXml);
            // Schema order: fill follows prstGeom (or xfrm) and precedes ln.
            InsertSpPrChildInOrder(spPr, newFill, ANs);
        }

        // --- line → replace a:ln ---
        if (hasLine)
        {
            var existingLn = spPr.ChildElements
                .FirstOrDefault(e => e.LocalName == "ln" && e.NamespaceUri == ANs);

            // Merge with the existing outline so a `set line.width` alone keeps
            // the current color/dash and vice versa — BuildLineXml rebuilds the
            // whole <a:ln> from only the keys it's handed, so any sub-prop NOT
            // in this call must be back-filled from the live <a:ln>. Skip the
            // merge for the explicit `line=none` clear (lineStyle=="none"),
            // which is meant to wipe the outline.
            bool isExplicitClear = string.Equals(lineStyle, "none", StringComparison.OrdinalIgnoreCase);
            if (!isExplicitClear && existingLn != null)
                ReadExistingLine(existingLn, ANs, ref lineStyle, ref lineWidth, ref lineColor);

            string lnXml = BuildLineXml(lineStyle, lineWidth, lineColor);
            if (existingLn != null) existingLn.Remove();
            if (!string.IsNullOrEmpty(lnXml))
            {
                var newLn = ParseShapeFragment(lnXml);
                InsertSpPrChildInOrder(spPr, newLn, ANs);
            }
        }

        // Forward genuinely out-of-scope keys as unsupported (position/rotation/
        // inset/text/shadow/gradient/wrap/alt/etc — see Add path).
        foreach (var key in properties.Keys)
        {
            switch (key.ToLowerInvariant())
            {
                case "fill": case "fillcolor":
                case "line": case "line.style": case "linestyle":
                case "line.width": case "linewidth":
                case "line.color": case "linecolor":
                case "width": case "height":
                case "geometry": case "preset":
                    continue;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        SaveDoc();
        return unsupported;
    }

    /// <summary>
    /// Resize a wpg:wgp group (e.g. a `diagram`). Mirrors the pptx group: set the
    /// group <c>&lt;a:ext&gt;</c> and the layout <c>&lt;wp:extent&gt;</c> while
    /// leaving <c>&lt;a:chExt&gt;</c> as the child-coordinate baseline (so Word
    /// compresses the children), then re-bake child font sizes by the net resize —
    /// Word scales group geometry on open but does NOT re-bake font (only an
    /// interactive drag does). fontRatio = min(width-ratio, height-ratio) so a
    /// width-only shrink still scales font down and a one-dimension grow leaves it.
    /// </summary>
    private List<string> SetGroupProps(OpenXmlElement wgp, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string WpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        OpenXmlElement? AChild(OpenXmlElement parent, string local) =>
            parent.ChildElements.FirstOrDefault(e => e.LocalName == local && e.NamespaceUri == ANs);

        // On the SET path an unparseable size is a caller error and must be
        // rejected (mirrors the pptx group). ParseDrawingSize swallows the parse
        // exception and returns the fallback — fine for ADD (fill in a default),
        // but on SET that silently reports success while changing nothing.
        long ParseSizeStrict(string raw, string label)
        {
            try { return ParseEmu(raw); }
            catch
            {
                throw new ArgumentException(
                    $"Invalid {label} '{raw}': expected a length like '6cm', '2in', '72pt' or a raw EMU integer.");
            }
        }

        var grpSpPr = wgp.ChildElements.FirstOrDefault(e => e.LocalName == "grpSpPr");
        var xfrm = grpSpPr != null ? AChild(grpSpPr, "xfrm") : null;
        var ext = xfrm != null ? AChild(xfrm, "ext") : null;

        string? widthRaw = properties.GetValueOrDefault("width");
        string? heightRaw = properties.GetValueOrDefault("height");
        if ((widthRaw != null || heightRaw != null) && ext != null)
        {
            long preCx = ReadUnqualifiedLong(ext, "cx") ?? 0;
            long preCy = ReadUnqualifiedLong(ext, "cy") ?? 0;

            long? newCx = widthRaw  != null ? ParseSizeStrict(widthRaw,  "width")  : null;
            long? newCy = heightRaw != null ? ParseSizeStrict(heightRaw, "height") : null;
            // Reject non-positive: a negative <a:ext> is dropped by the SDK and
            // corrupts the file (schema MinInclusive); zero makes an invisible
            // group. Mirrors the pptx group's rejection.
            if (newCx is <= 0) throw new ArgumentException($"Invalid width '{widthRaw}': width must be a positive size.");
            if (newCy is <= 0) throw new ArgumentException($"Invalid height '{heightRaw}': height must be a positive size.");
            // keepAspect=true + a single dimension → scale the OTHER
            // proportionally so a diagram group stays aspect-correct (a lone
            // width/height stretch squashes a flowchart). Default honors exactly
            // the axes the caller passed, matching `set width` on a plain shape
            // (GitHub #237 — the lock must be opt-in; mirrors the pptx group).
            bool keepAspect = properties.Any(kv =>
                kv.Key.Equals("keepAspect", StringComparison.OrdinalIgnoreCase) && IsTruthy(kv.Value));
            if (keepAspect && newCx != null && newCy == null && preCx > 0)
                newCy = (long)Math.Round(preCy * (newCx.Value / (double)preCx));
            else if (keepAspect && newCy != null && newCx == null && preCy > 0)
                newCx = (long)Math.Round(preCx * (newCy.Value / (double)preCy));

            // The anchor wrapper carries the matching <wp:extent> for the whole group.
            var wpExtent = wgp.Ancestors()
                .FirstOrDefault(e => e.LocalName == "anchor" || e.LocalName == "inline")
                ?.Descendants().FirstOrDefault(e => e.LocalName == "extent" && e.NamespaceUri == WpNs);
            if (newCx != null)
            {
                ext.SetAttribute(new OpenXmlAttribute("cx", "", newCx.Value.ToString()));
                wpExtent?.SetAttribute(new OpenXmlAttribute("cx", "", newCx.Value.ToString()));
            }
            if (newCy != null)
            {
                ext.SetAttribute(new OpenXmlAttribute("cy", "", newCy.Value.ToString()));
                wpExtent?.SetAttribute(new OpenXmlAttribute("cy", "", newCy.Value.ToString()));
            }
            if (preCx > 0 && preCy > 0)
            {
                double ratio = Math.Min((newCx ?? preCx) / (double)preCx, (newCy ?? preCy) / (double)preCy);
                if (Math.Abs(ratio - 1.0) > 1e-6) ScaleGroupFontHalfPts(wgp, ratio);
            }
        }

        // x / y reposition the whole floating group by updating the anchor's
        // <wp:positionH>/<wp:positionV> <wp:posOffset> (matches the pptx group's
        // set x/y, so an agent can move the diagram as a unit in either format).
        string? xRaw = properties.GetValueOrDefault("x");
        string? yRaw = properties.GetValueOrDefault("y");
        if (xRaw != null || yRaw != null)
        {
            var anchor = wgp.Ancestors()
                .FirstOrDefault(e => e.LocalName == "anchor" && e.NamespaceUri == WpNs)
                as DW.Anchor;
            if (anchor != null)
            {
                if (xRaw != null)
                    SetAnchorAxisOffset(anchor.GetFirstChild<DW.HorizontalPosition>(),
                        ParseSizeStrict(xRaw, "x"));
                if (yRaw != null)
                    SetAnchorAxisOffset(anchor.GetFirstChild<DW.VerticalPosition>(),
                        ParseSizeStrict(yRaw, "y"));
            }
        }

        foreach (var k in properties.Keys)
            if (k.ToLowerInvariant() is not ("width" or "height" or "x" or "y" or "keepaspect"))
                unsupported.Add(k);
        SaveDoc();
        return unsupported;
    }

    // Set a floating anchor axis (positionH/positionV) to an absolute EMU
    // <wp:posOffset>, replacing any <wp:align> (align and posOffset are a schema
    // choice — mutually exclusive). No-op if the axis element is absent.
    private static void SetAnchorAxisOffset(OpenXmlElement? posAxis, long emu)
    {
        if (posAxis == null) return;
        posAxis.ChildElements
            .Where(e => e.LocalName is "align" or "posOffset")
            .ToList().ForEach(e => e.Remove());
        posAxis.AppendChild(new DW.PositionOffset(emu.ToString()));
    }

    // Multiply every child run's font size (w:sz / w:szCs, in half-points) by
    // <paramref name="ratio"/>, floor 1pt (2 half-points).
    private static void ScaleGroupFontHalfPts(OpenXmlElement wgp, double ratio)
    {
        foreach (var sz in wgp.Descendants<FontSize>())
            if (int.TryParse(sz.Val?.Value, out var hp))
                sz.Val = Math.Max(2, (int)Math.Round(hp * ratio)).ToString();
        foreach (var sz in wgp.Descendants<FontSizeComplexScript>())
            if (int.TryParse(sz.Val?.Value, out var hp))
                sz.Val = Math.Max(2, (int)Math.Round(hp * ratio)).ToString();
    }

    /// <summary>
    /// Insert a freshly-built spPr child (solidFill/noFill/ln) at the correct
    /// CT_ShapeProperties position. Order is: xfrm, prstGeom (or custGeom), fill
    /// (noFill/solidFill/gradFill/blipFill/pattFill/grpFill), ln, effectLst, …
    /// We place fill before any existing ln; ln after any existing fill. Falls
    /// back to append when no successor anchor is present.
    /// </summary>
    private static void InsertSpPrChildInOrder(OpenXmlElement spPr, OpenXmlElement child, string aNs)
    {
        bool isLn = child.LocalName == "ln";
        if (isLn)
        {
            // ln goes after fill/prstGeom/xfrm but before effectLst/scene3d/…
            var after = spPr.ChildElements.LastOrDefault(e => e.NamespaceUri == aNs
                && e.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill"
                    or "prstGeom" or "custGeom" or "xfrm");
            if (after != null) { after.InsertAfterSelf(child); return; }
            spPr.AppendChild(child);
            return;
        }
        // fill: after prstGeom/xfrm, before ln (and anything later).
        var lnEl = spPr.ChildElements.FirstOrDefault(e => e.LocalName == "ln" && e.NamespaceUri == aNs);
        if (lnEl != null) { lnEl.InsertBeforeSelf(child); return; }
        var geomOrXfrm = spPr.ChildElements.LastOrDefault(e => e.NamespaceUri == aNs
            && e.LocalName is "prstGeom" or "custGeom" or "xfrm");
        if (geomOrXfrm != null) { geomOrXfrm.InsertAfterSelf(child); return; }
        spPr.AppendChild(child);
    }

    /// <summary>
    /// Read the current outline (<c>a:ln</c>) sub-properties — width (as an
    /// explicit "<emu>emu" string), dash style and color — into the
    /// <paramref name="style"/>/<paramref name="width"/>/<paramref name="color"/>
    /// slots, but ONLY where the caller left them null. Lets a partial
    /// <c>set line.X</c> preserve the unspecified sub-props instead of letting
    /// BuildLineXml reset them to its defaults (black solid, theme width).
    /// </summary>
    private static void ReadExistingLine(OpenXmlElement ln, string aNs,
        ref string? style, ref string? width, ref string? color)
    {
        if (width == null)
        {
            var w = ReadUnqualifiedLong(ln, "w");
            // FormatEmu-style explicit suffix so ParseEmu reads it back as raw
            // EMU (a bare integer would be mis-read as points).
            if (w is > 0) width = $"{w}emu";
        }
        if (style == null)
        {
            var dash = ln.ChildElements
                .FirstOrDefault(e => e.LocalName == "prstDash" && e.NamespaceUri == aNs);
            var dashVal = dash?.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "val" && string.IsNullOrEmpty(a.NamespaceUri)).Value;
            if (!string.IsNullOrEmpty(dashVal)) style = dashVal;
        }
        if (color == null)
        {
            var solidFill = ln.ChildElements
                .FirstOrDefault(e => e.LocalName == "solidFill" && e.NamespaceUri == aNs);
            var srgb = solidFill?.ChildElements
                .FirstOrDefault(e => e.LocalName == "srgbClr" && e.NamespaceUri == aNs);
            var clrVal = srgb?.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "val" && string.IsNullOrEmpty(a.NamespaceUri)).Value;
            if (!string.IsNullOrEmpty(clrVal)) color = clrVal;
        }
    }

    /// <summary>Read an unqualified (no-namespace) integer attribute off an
    /// element, returning null when absent/unparsable.</summary>
    private static long? ReadUnqualifiedLong(OpenXmlElement? el, string name)
    {
        if (el == null) return null;
        var attr = el.GetAttributes().FirstOrDefault(a => a.LocalName == name && string.IsNullOrEmpty(a.NamespaceUri));
        return long.TryParse(attr.Value, out var v) ? v : (long?)null;
    }

    /// <summary>Parse a DrawingML spPr fragment (e.g. "&lt;a:solidFill&gt;…")
    /// into an OpenXmlElement, keeping the a:/wps: namespace context alive.
    /// Mirrors ParseDrawingFromXml's XmlReader-via-wrapper approach.</summary>
    private static OpenXmlElement ParseShapeFragment(string fragment)
    {
        // Route through a w:p > w:r > w:drawing wrapper (same approach as
        // ParseDrawingFromXml) so the a:/wps: prefixes resolve, then lift the
        // built fragment out of the parsed spPr.
        var wrapXml =
            $@"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""><wps:wsp><wps:spPr>{fragment}</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        var p = new Paragraph(wrapXml);
        var spPr = p.Descendants().FirstOrDefault(e => e.LocalName == "spPr"
            && e.NamespaceUri == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        var first = spPr?.FirstChild
            ?? throw new InvalidOperationException("Shape fragment parse failed");
        first.Remove();
        return first;
    }

}
