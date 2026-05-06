// Copyright 2025 OfficeCli (officecli.ai)
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

        _doc.MainDocumentPart?.Document?.Save();
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
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
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
        // CONSISTENCY(run-special-content): mirror the 5-way type upgrade
        // in Navigation.cs — ptab / fieldChar / instrText / tab / break.
        // Round 11 caught that `tab` was missing from this judgment:
        // Get strips typography from tab runs, but Set silently accepted
        // bold/color/font writes onto them, breaking read/write symmetry.
        bool isSpecialRun = ptabEl != null || fldCharEl != null || instrEl != null
                            || (breakElInline != null && !hasText)
                            || (tabElInline != null && !hasText);

        foreach (var (key, value) in properties)
        {
            // CONSISTENCY(run-special-content): typography props (font.* /
            // size / bold / color / underline …) are noise on ptab /
            // fieldChar / instrText / tab / break runs because there is no
            // glyph to apply them to. Get strips them on readback (Round 2);
            // accepting them on Set would write to <w:rPr> anyway and
            // diverge between the read and write surfaces. Reject so the
            // caller sees a clean unsupported notice and the OOXML stays
            // free of cosmetic-but-invisible noise.
            if (isSpecialRun && IsTypographyOnlyKey(key))
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
                        var ptStrW = (ParseEmu(value) / 12700.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt";
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
                        var ptStrH = (ParseEmu(value) / 12700.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt";
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
                    // part-cleanup rule from CLAUDE.md's Known API
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
                    // (mirrors Set src cleanup rule in CLAUDE.md Known
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
                case "hposition":
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
                case "vposition":
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
                        var uri = isAbs ? absUri! : new Uri(value, UriKind.Relative);
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
                    var mathContent = FormulaParser.Parse(value);
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
                    static string StripPct(string s)
                    {
                        var t = s.Trim();
                        return t.EndsWith("%", StringComparison.Ordinal) ? t[..^1].Trim() : t;
                    }
                    var drawingCrop = run.GetFirstChild<Drawing>();
                    var blipFillCrop = drawingCrop?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.BlipFill>().FirstOrDefault();
                    if (blipFillCrop == null) { unsupported.Add(key); break; }
                    var srcRectCrop = blipFillCrop.GetFirstChild<A.SourceRectangle>();
                    if (srcRectCrop == null)
                    {
                        srcRectCrop = new A.SourceRectangle();
                        // CONSISTENCY(ooxml-element-order): srcRect precedes the fill-mode element.
                        var fillModeCrop = (OpenXmlElement?)blipFillCrop.GetFirstChild<A.Stretch>()
                            ?? blipFillCrop.GetFirstChild<A.Tile>();
                        if (fillModeCrop != null)
                            blipFillCrop.InsertBefore(srcRectCrop, fillModeCrop);
                        else
                            blipFillCrop.AppendChild(srcRectCrop);
                    }
                    if (key.Equals("crop", StringComparison.OrdinalIgnoreCase))
                    {
                        var partsCrop = value.Split(',');
                        if (partsCrop.Length == 4)
                        {
                            var cv = new double[4];
                            for (int ci = 0; ci < 4; ci++)
                            {
                                cv[ci] = ParseHelpers.SafeParseDouble(StripPct(partsCrop[ci]), "crop");
                                if (cv[ci] < 0 || cv[ci] > 100)
                                    throw new ArgumentException($"Invalid 'crop' value: '{partsCrop[ci].Trim()}'. Crop percentage must be between 0 and 100.");
                            }
                            srcRectCrop.Left = (int)(cv[0] * 1000);
                            srcRectCrop.Top = (int)(cv[1] * 1000);
                            srcRectCrop.Right = (int)(cv[2] * 1000);
                            srcRectCrop.Bottom = (int)(cv[3] * 1000);
                        }
                        else if (partsCrop.Length == 1)
                        {
                            if (!double.TryParse(StripPct(value), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cv1)
                                || cv1 < 0 || cv1 > 100)
                                throw new ArgumentException($"Invalid 'crop' value: '{value}'. Expected percentage 0-100.");
                            var pctAll = (int)(cv1 * 1000);
                            srcRectCrop.Left = pctAll; srcRectCrop.Top = pctAll;
                            srcRectCrop.Right = pctAll; srcRectCrop.Bottom = pctAll;
                        }
                        else
                        {
                            throw new ArgumentException($"Invalid 'crop' value: '{value}'. Expected 1 or 4 comma-separated percentages.");
                        }
                    }
                    else
                    {
                        if (!double.TryParse(StripPct(value), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cs1)
                            || cs1 < 0 || cs1 > 100)
                            throw new ArgumentException($"Invalid '{key}' value: '{value}'. Expected percentage 0-100.");
                        var pctSide = (int)(cs1 * 1000);
                        switch (key.ToLowerInvariant())
                        {
                            case "cropleft": srcRectCrop.Left = pctSide; break;
                            case "croptop": srcRectCrop.Top = pctSide; break;
                            case "cropright": srcRectCrop.Right = pctSide; break;
                            case "cropbottom": srcRectCrop.Bottom = pctSide; break;
                        }
                    }
                    int Lc = srcRectCrop.Left?.Value ?? 0;
                    int Tc = srcRectCrop.Top?.Value ?? 0;
                    int Rc = srcRectCrop.Right?.Value ?? 0;
                    int Bc = srcRectCrop.Bottom?.Value ?? 0;
                    if (Lc == 0 && Tc == 0 && Rc == 0 && Bc == 0)
                        srcRectCrop.Remove();
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
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid run props: text, bold, italic, font, size, color, underline, strike, highlight, caps, smallcaps, superscript, subscript, shading, link, formula)"
                            : key);
                    }
                    break;
            }
        }

        var affectedPara = run.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        _doc.MainDocumentPart?.Document?.Save();
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
                    var uri = isAbs ? absUri! : new Uri(value, UriKind.Relative);
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
                    unsupported.Add(key);
                    break;
            }
        }

        var affectedPara = hl.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        _doc.MainDocumentPart?.Document?.Save();
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
                    var mathContent = FormulaParser.Parse(value);
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
                            hostPara.InsertBefore(clone, mPara);
                            mPara.Remove();
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
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetElementParagraph(Paragraph para, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            if (ApplyParagraphLevelProperty(pProps, key, value))
            {
                // CONSISTENCY(rtl-cascade): direction toggle stamps the full
                // bidi+markRPr+runs cascade. See WordHandler.I18n.cs.
                if (k is "direction" or "dir" or "bidi")
                    ApplyDirectionCascade(para, ParseDirectionRtl(value));
                // handled by paragraph-level helper
            }
            else switch (k)
            {
                case "formula":
                {
                    // Replace paragraph content with OMML equation in-place
                    foreach (var child in para.ChildElements
                        .Where(c => c is not ParagraphProperties).ToList())
                        child.Remove();
                    var mathContent = FormulaParser.Parse(value);
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
                    var pmrp = pProps.ParagraphMarkRunProperties
                        ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    pmrp.RemoveAllChildren<RunStyle>();
                    pmrp.PrependChild(new RunStyle { Val = value });
                    foreach (var pRun in para.Descendants<Run>())
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
                    var sub = key.Substring("markRPr.".Length);
                    var markOnlyRPr = pProps.ParagraphMarkRunProperties
                        ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    ApplyRunFormatting(markOnlyRPr, sub, value);
                    break;
                }
                case "size" or "font" or "bold" or "italic" or "color" or "highlight" or "underline" or "strike"
                  or "font.latin" or "font.ea" or "font.eastasia" or "font.eastasian"
                  or "font.cs" or "font.complexscript" or "font.complex"
                  or "bold.cs" or "italic.cs" or "size.cs"
                  or "font.bold.cs" or "font.italic.cs" or "font.size.cs":
                    // Apply run-level formatting to all runs in the paragraph
                    var allParaRuns = para.Descendants<Run>().ToList();
                    // Also update paragraph mark run properties (rPr inside pPr)
                    // so new runs inherit the formatting
                    var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    ApplyRunFormatting(markRPr, key, value);
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
                        // Use paragraph mark run properties as default for new run
                        var newRun = new Run();
                        var markProps = pProps.ParagraphMarkRunProperties;
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
                        var paraRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                            ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                        if (Core.TypedAttributeFallback.TrySet(paraRPr, key, value))
                            break;
                        if (paraRPr.ChildElements.Count == 0)
                            paraRPr.Remove();
                    }
                    if (!GenericXmlQuery.TryCreateTypedChild(pProps, key, value))
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid paragraph props: text, style, alignment, bold, italic, font, size, color, spaceBefore, spaceAfter, lineSpacing, indent, liststyle, formula, direction, bidi)"
                            : key);
                    break;
            }
        }

        var affectedPara = para;
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
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
                    tab.Position = (int)SpacingConverter.ParseWordSpacing(value);
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
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetElementTableCell(TableCell cell, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var tcPr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        string? deferredText = null;
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                    // Defer text handling until after formatting is applied
                    deferredText = value;
                    break;
                case "font":
                case "size":
                case "bold":
                case "italic":
                case "color":
                case "highlight":
                case "underline":
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
                        var pmrp = pPr.ParagraphMarkRunProperties ?? pPr.AppendChild(new ParagraphMarkRunProperties());
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
                case "shd" or "shading" or "fill":
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
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[0]).Rgb;
                        }
                        else if (shdParts.Length >= 2)
                        {
                            var cellPat = shdParts[0].TrimStart('#');
                            if (cellPat.Length >= 6 && cellPat.All(char.IsAsciiHexDigit))
                            { shd.Val = ShadingPatternValues.Clear; shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[0]).Rgb; }
                            else
                            {
                                WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                                shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[1]).Rgb;
                                if (shdParts.Length >= 3) shd.Color = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[2]).Rgb;
                            }
                        }
                        tcPr.Shading = shd;
                    }
                    break;
                case "align" or "alignment":
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
                    else
                    {
                        tcPr.TableCellWidth = new TableCellWidth { Width = ParseHelpers.SafeParseUint(value, "width").ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "padding":
                {
                    var dxa = ParseHelpers.SafeParseUint(value, "padding").ToString();
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.TopMargin = new TopMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    mar.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    mar.LeftMargin = new LeftMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    mar.RightMargin = new RightMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.top":
                {
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.TopMargin = new TopMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.bottom":
                {
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.BottomMargin = new BottomMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.left":
                {
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.LeftMargin = new LeftMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "padding.right":
                {
                    var mar = tcPr.TableCellMargin ?? (tcPr.TableCellMargin = new TableCellMargin());
                    mar.RightMargin = new RightMargin { Width = value, Type = TableWidthUnitValues.Dxa };
                    break;
                }
                case "textdirection" or "textdir":
                    tcPr.TextDirection = new TextDirection
                    {
                        Val = value.ToLowerInvariant() switch
                        {
                            "btlr" or "vertical" => TextDirectionValues.BottomToTopLeftToRight,
                            "tbrl" or "vertical-rl" => TextDirectionValues.TopToBottomRightToLeft,
                            "lrtb" or "horizontal" => TextDirectionValues.LefToRightTopToBottom,
                            "tbrl-r" or "tb-rl-rotated" => TextDirectionValues.TopToBottomRightToLeftRotated,
                            "lrtb-r" or "lr-tb-rotated" => TextDirectionValues.LefttoRightTopToBottomRotated,
                            "tblr-r" or "tb-lr-rotated" => TextDirectionValues.TopToBottomLeftToRightRotated,
                            _ => throw new ArgumentException($"Invalid textDirection value: '{value}'. Valid values: lrtb, btlr, tbrl, horizontal, vertical.")
                        }
                    };
                    break;
                case "nowrap":
                    tcPr.NoWrap = IsTruthy(value) ? new NoWrap() : null;
                    break;
                case "vmerge":
                    // ST_Merge schema only defines "restart" — continuation is bare <w:vMerge/>.
                    tcPr.VerticalMerge = value.ToLowerInvariant() == "restart"
                        ? new VerticalMerge { Val = MergedCellValues.Restart }
                        : new VerticalMerge();
                    break;
                case "hmerge":
                    // Mirrors vmerge: ST_Merge schema only defines "restart" —
                    // continuation is bare <w:hMerge/>.
                    tcPr.HorizontalMerge = value.ToLowerInvariant() == "restart"
                        ? new HorizontalMerge { Val = MergedCellValues.Restart }
                        : new HorizontalMerge();
                    break;
                case var k when k.StartsWith("border"):
                    ApplyCellBorders(tcPr, key, value);
                    break;
                case "gridspan" or "colspan":
                    var newSpan = ParseHelpers.SafeParseInt(value, "gridspan");
                    if (newSpan <= 0)
                        throw new ArgumentException($"Invalid 'gridspan' value: '{value}'. Must be a positive integer (> 0).");
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
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetElementTableRow(TableRow row, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var trPr = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "height":
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
                            targetPara.AppendChild(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                    }
                    else if (key.Contains('.')
                        && Core.TypedAttributeFallback.TrySet(trPr, key, value))
                    {
                        // Generic dotted fallback (e.g. trHeight.* attrs).
                    }
                    else if (!GenericXmlQuery.TryCreateTypedChild(trPr, key, value))
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid row props: height, height.exact, header, cantSplit, c1, c2, ...)"
                            : key);
                    break;
            }
        }

        var affectedPara = row.Ancestors<Paragraph>().FirstOrDefault();
        if (affectedPara != null)
            affectedPara.TextId = GenerateParaId();
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    private List<string> SetElementTable(Table tbl, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var tblPr = tbl.GetFirstChild<TableProperties>() ?? tbl.PrependChild(new TableProperties());
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "style":
                    var tblStyle = tblPr.TableStyle ?? (tblPr.TableStyle = new TableStyle());
                    tblStyle.Val = value;
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
                        var pct = ParseHelpers.SafeParseInt(value.TrimEnd('%'), "width") * 50; // OOXML pct = percent * 50
                        tblPr.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                    }
                    else
                    {
                        // CONSISTENCY(spacing-units): accept unit-qualified lengths
                        // ('10cm', '5in', '12pt') alongside bare twips, matching
                        // Add and the cross-handler convention from
                        // root CLAUDE.md "Spacing input is lenient". Previous
                        // SafeParseUint-only path rejected '10cm'.
                        var twips = OfficeCli.Core.SpacingConverter.ParseWordSpacing(value);
                        tblPr.TableWidth = new TableWidth { Width = twips.ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "indent":
                    tblPr.TableIndentation = new TableIndentation { Width = ParseHelpers.SafeParseInt(value, "indent"), Type = TableWidthUnitValues.Dxa };
                    break;
                case "cellspacing":
                    tblPr.TableCellSpacing = new TableCellSpacing { Width = ParseHelpers.SafeParseUint(value, "cellspacing").ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                case "layout":
                    tblPr.TableLayout = new TableLayout
                    {
                        Type = value.ToLowerInvariant() == "fixed" ? TableLayoutValues.Fixed : TableLayoutValues.Autofit
                    };
                    break;
                case "padding":
                {
                    var dxa = value;
                    var cm = tblPr.TableCellMarginDefault ?? tblPr.AppendChild(new TableCellMarginDefault());
                    cm.TopMargin = new TopMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    var paddingVal = ParseHelpers.SafeParseInt(dxa, "padding");
                    cm.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    cm.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    break;
                }
                case "firstrow":
                case "lastrow":
                case "firstcol" or "firstcolumn":
                case "lastcol" or "lastcolumn":
                case "bandrow" or "bandedrows" or "bandrows":
                case "bandcol" or "bandedcols" or "bandcols":
                {
                    var tblLook = tblPr.GetFirstChild<TableLook>();
                    if (tblLook == null) { tblLook = new TableLook { Val = "04A0" }; tblPr.AppendChild(tblLook); }
                    var bv = IsTruthy(value);
                    switch (key.ToLowerInvariant())
                    {
                        case "firstrow": tblLook.FirstRow = bv; break;
                        case "lastrow": tblLook.LastRow = bv; break;
                        case "firstcol" or "firstcolumn": tblLook.FirstColumn = bv; break;
                        case "lastcol" or "lastcolumn": tblLook.LastColumn = bv; break;
                        case "bandrow" or "bandedrows" or "bandrows": tblLook.NoHorizontalBand = !bv; break;
                        case "bandcol" or "bandedcols" or "bandcols": tblLook.NoVerticalBand = !bv; break;
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
                            tblPr.AppendChild(tpp);
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
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.HorizontalAnchor = value.ToLowerInvariant() switch
                    {
                        "margin" => HorizontalAnchorValues.Margin,
                        "page" => HorizontalAnchorValues.Page,
                        "text" => HorizontalAnchorValues.Text,
                        _ => throw new ArgumentException($"Invalid horizontalAnchor: '{value}'. Valid: margin, page, text.")
                    };
                    break;
                }
                case "position.vanchor" or "position.verticalanchor":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.VerticalAnchor = value.ToLowerInvariant() switch
                    {
                        "margin" => VerticalAnchorValues.Margin,
                        "page" => VerticalAnchorValues.Page,
                        "text" => VerticalAnchorValues.Text,
                        _ => throw new ArgumentException($"Invalid verticalAnchor: '{value}'. Valid: margin, page, text.")
                    };
                    break;
                }
                case "position.leftfromtext" or "position.left":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.LeftFromText = (short)ParseTwips(value);
                    break;
                }
                case "position.rightfromtext" or "position.right":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.RightFromText = (short)ParseTwips(value);
                    break;
                }
                case "position.topfromtext" or "position.top":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.TopFromText = (short)ParseTwips(value);
                    break;
                }
                case "position.bottomfromtext" or "position.bottom":
                {
                    var tpp = EnsureTablePositionProperties(tblPr);
                    tpp.BottomFromText = (short)ParseTwips(value);
                    break;
                }
                case "overlap":
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
                        tblPr.AppendChild(new TableCaption { Val = value });
                    break;
                case "description":
                    tblPr.RemoveAllChildren<TableDescription>();
                    if (!string.IsNullOrEmpty(value))
                        tblPr.AppendChild(new TableDescription { Val = value });
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
                    var tblGrid = tbl.GetFirstChild<TableGrid>();
                    if (tblGrid == null)
                    {
                        tblGrid = new TableGrid();
                        tbl.InsertAfter(tblGrid, tblPr);
                    }
                    var gridCols = tblGrid.Elements<GridColumn>().ToList();
                    for (int ci = 0; ci < parts.Length; ci++)
                    {
                        var twips = ParseTwips(parts[ci].Trim());
                        if (ci < gridCols.Count)
                            gridCols[ci].Width = twips.ToString();
                        else
                            tblGrid.AppendChild(new GridColumn { Width = twips.ToString() });
                        // Also update cell widths in each row for this column
                        foreach (var tblRow in tbl.Elements<TableRow>())
                        {
                            var cells = tblRow.Elements<TableCell>().ToList();
                            if (ci < cells.Count)
                            {
                                var tcPr = cells[ci].TableCellProperties ?? cells[ci].PrependChild(new TableCellProperties());
                                tcPr.TableCellWidth = new TableCellWidth { Width = twips.ToString(), Type = TableWidthUnitValues.Dxa };
                            }
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
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

}
