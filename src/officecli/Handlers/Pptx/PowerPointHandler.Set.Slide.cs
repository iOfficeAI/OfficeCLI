// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for slide / master / layout / notes paths.
// Mechanically extracted from the original god-method Set(); each helper
// owns one path-pattern's full handling. No behavior change.
public partial class PowerPointHandler
{
    private List<string> SetNotesByPath(Match notesSetMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(notesSetMatch.Groups[1].Value);
        var slidePartsN = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slidePartsN.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slidePartsN.Count})");
        var notesPart = EnsureNotesSlidePart(slidePartsN[PathIndex.ToArrayIndex(slideIdx)]);
        var unsupportedN = new List<string>();
        // Pull the notes body shape (idx=1 placeholder) so run-level keys
        // (lang, lang.*, font, size, color, …) route through the same
        // SetRunOrShapeProperties pipeline as regular slide shapes.
        // CONSISTENCY(notes-shape-set): notes had its own bespoke key
        // handling that recognised only text/direction; other run keys
        // surfaced as UNSUPPORTED. The notes body is just a Shape — it
        // should accept the full run-attr surface.
        Shape? notesBody = null;
        var notesShapeTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
        if (notesShapeTree != null)
        {
            foreach (var sh in notesShapeTree.Elements<Shape>())
            {
                var ph = sh.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();
                if (ph?.Index?.Value == 1) { notesBody = sh; break; }
            }
        }

        var deferredRunProps = new Dictionary<string, string>();
        foreach (var (key, value) in properties)
        {
            if (key.Equals("text", StringComparison.OrdinalIgnoreCase))
            {
                XmlTextValidator.ValidateOrThrow(value, "text", allowSoftBreakChar: true);
                SetNotesText(notesPart, value);
            }
            else if (key.Equals("direction", StringComparison.OrdinalIgnoreCase)
                  || key.Equals("dir", StringComparison.OrdinalIgnoreCase)
                  || key.Equals("rtl", StringComparison.OrdinalIgnoreCase))
                ApplyNotesDirection(notesPart, value);
            else
                // Defer to SetRunOrShapeProperties — handles lang, lang.*,
                // sz, b, i, u, font, color, etc. on the notes body shape.
                deferredRunProps[key] = value;
        }
        if (deferredRunProps.Count > 0)
        {
            if (notesBody == null)
                unsupportedN.AddRange(deferredRunProps.Keys);
            else
            {
                var notesRuns = notesBody.Descendants<Drawing.Run>().ToList();
                unsupportedN.AddRange(SetRunOrShapeProperties(deferredRunProps, notesRuns, notesBody));
            }
        }
        notesPart.NotesSlide!.Save();
        return unsupportedN;
    }

    private List<string> SetMasterShapeByPath(Match masterShapeMatch, Dictionary<string, string> properties)
    {
        // CONSISTENCY(master-layout-shape-edit): partType is lowercased by
        // NormalizePptxPathSegmentCasing — compare case-insensitively.
        var partType = masterShapeMatch.Groups[1].Value;
        var partIdx = int.Parse(masterShapeMatch.Groups[2].Value);
        var presentationPart = _doc.PresentationPart!;

        OpenXmlPart ownerPart;
        OpenXmlPartRootElement rootEl;
        // CONSISTENCY(master-layout-path-aliases): accept both `slidemaster` and
        // short `master`; same for `slidelayout` / `layout`.
        var isMaster = partType.Equals("slidemaster", StringComparison.OrdinalIgnoreCase)
                    || partType.Equals("master", StringComparison.OrdinalIgnoreCase);
        if (isMaster)
        {
            var masters = presentationPart.SlideMasterParts.ToList();
            if (partIdx < 1 || partIdx > masters.Count)
                throw new ArgumentException($"SlideMaster {partIdx} not found (total: {masters.Count})");
            ownerPart = masters[partIdx - 1];
            rootEl = masters[partIdx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt slide master");
        }
        else
        {
            var layouts = PowerPointHandler.LayoutsInOrder(presentationPart);
            if (partIdx < 1 || partIdx > layouts.Count)
                throw new ArgumentException($"SlideLayout {partIdx} not found (total: {layouts.Count})");
            ownerPart = layouts[partIdx - 1];
            rootEl = layouts[partIdx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt slide layout");
        }

        return ApplyMasterLayoutShapeOrSelfProperties(masterShapeMatch, properties, ownerPart, rootEl);
    }

    // CONSISTENCY(master-layout-shape-edit): nested form
    // /slidemaster[N]/slidelayout[L]/shape[K] gets its own dispatcher so the
    // top-level flat regex (which captures part-type at group[1]) stays
    // unambiguous. Shared body via ApplyMasterLayoutShapeOrSelfProperties.
    private List<string> SetNestedMasterLayoutShapeByPath(Match nestedMatch, Dictionary<string, string> properties)
    {
        var mIdx = int.Parse(nestedMatch.Groups[1].Value);
        var lIdx = int.Parse(nestedMatch.Groups[2].Value);
        var presentationPart = _doc.PresentationPart!;
        var masters = presentationPart.SlideMasterParts.ToList();
        if (mIdx < 1 || mIdx > masters.Count)
            throw new ArgumentException($"SlideMaster {mIdx} not found (total: {masters.Count})");
        var layouts = masters[mIdx - 1].SlideLayoutParts.ToList();
        if (lIdx < 1 || lIdx > layouts.Count)
            throw new ArgumentException($"SlideLayout {lIdx} not found under master {mIdx} (total: {layouts.Count})");
        var lp = layouts[lIdx - 1];
        var rootEl = lp.SlideLayout
            ?? throw new InvalidOperationException("Corrupt slide layout");

        // Reuse the shape/self body by synthesising a match with groups[3]/[4]
        // shifted from the nested capture (groups[3]/[4] in nestedMatch).
        return ApplyMasterLayoutShapeOrSelfProperties(nestedMatch, properties, lp, rootEl, shapeTypeGroup: 3, shapeIdxGroup: 4);
    }

    private List<string> ApplyMasterLayoutShapeOrSelfProperties(
        Match m,
        Dictionary<string, string> properties,
        OpenXmlPart ownerPart,
        OpenXmlPartRootElement rootEl,
        int shapeTypeGroup = 3,
        int shapeIdxGroup = 4)
    {
        if (!m.Groups[shapeTypeGroup].Success)
        {
            // Set properties on the master/layout itself
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (key.Equals("name", StringComparison.OrdinalIgnoreCase))
                {
                    var csd = rootEl.GetFirstChild<CommonSlideData>();
                    if (csd != null) csd.Name = value;
                }
                else
                {
                    if (unsupported.Count == 0)
                        unsupported.Add($"{key} (valid master/layout props: name)");
                    else
                        unsupported.Add(key);
                }
            }
            rootEl.Save();
            return unsupported;
        }

        var elType = m.Groups[shapeTypeGroup].Value;
        var elIdx = int.Parse(m.Groups[shapeIdxGroup].Value);
        var shapeTree = rootEl.Descendants<ShapeTree>().FirstOrDefault()
            ?? throw new ArgumentException("No shape tree found");

        if (elType.Equals("shape", StringComparison.OrdinalIgnoreCase))
        {
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (elIdx < 1 || elIdx > shapes.Count)
                throw new ArgumentException($"Shape {elIdx} not found (total: {shapes.Count})");
            var shape = shapes[elIdx - 1];
            var allRuns = shape.Descendants<Drawing.Run>().ToList();
            // Pass the owning part so fill/image/effect helpers that need a
            // relationship anchor (e.g. picture fills) write to the correct part.
            var unsupp = SetRunOrShapeProperties(properties, allRuns, shape, ownerPart, unrecognizedLatex: LastUnrecognizedLatex);
            rootEl.Save();
            return unsupp;
        }

        throw new ArgumentException($"Unsupported element type: '{elType}' for master/layout. Valid types: shape.");
    }

    private List<string> SetMasterOrLayoutBackgroundByPath(Match masterBgMatch, Match layoutBgMatch, Dictionary<string, string> properties)
    {
        OpenXmlPart targetPart;
        OpenXmlPartRootElement targetRoot;
        if (masterBgMatch.Success)
        {
            var masterIdx = int.Parse(masterBgMatch.Groups[1].Value);
            var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
            if (masterIdx < 1 || masterIdx > masters.Count)
                throw new ArgumentException($"Slide master {masterIdx} not found (total: {masters.Count})");
            var mp = masters[masterIdx - 1];
            if (masterBgMatch.Groups[2].Success)
            {
                var lIdx = int.Parse(masterBgMatch.Groups[2].Value);
                var layouts = mp.SlideLayoutParts?.ToList() ?? [];
                if (lIdx < 1 || lIdx > layouts.Count)
                    throw new ArgumentException($"Slide layout {lIdx} not found under master {masterIdx} (total: {layouts.Count})");
                targetPart = layouts[lIdx - 1];
                targetRoot = layouts[lIdx - 1].SlideLayout
                    ?? throw new InvalidOperationException("Corrupt slide layout");
            }
            else
            {
                targetPart = mp;
                targetRoot = mp.SlideMaster
                    ?? throw new InvalidOperationException("Corrupt slide master");
            }
        }
        else
        {
            var lIdx = int.Parse(layoutBgMatch.Groups[1].Value);
            var allLayouts = (_doc.PresentationPart?.SlideMasterParts ?? Enumerable.Empty<SlideMasterPart>())
                .SelectMany(m => m.SlideLayoutParts ?? Enumerable.Empty<SlideLayoutPart>()).ToList();
            if (lIdx < 1 || lIdx > allLayouts.Count)
                throw new ArgumentException($"Slide layout {lIdx} not found (total: {allLayouts.Count})");
            targetPart = allLayouts[lIdx - 1];
            targetRoot = allLayouts[lIdx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt slide layout");
        }

        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "background":
                    ApplyBackground(targetPart, value, ReadBackgroundImageOptions(properties));
                    break;
                case "background.mode":
                case "background.alpha":
                case "background.scale":
                    break;
                case "name":
                {
                    var csd = targetRoot.GetFirstChild<CommonSlideData>();
                    if (csd != null)
                    {
                        XmlTextValidator.ValidateOrThrow(value, "name");
                        csd.Name = value;
                    }
                    break;
                }
                case "direction" or "dir" or "rtl":
                {
                    // Layout/master-level RTL. Two prongs:
                    //   1. Cascade <a:pPr rtl="1"/> onto every paragraph in every
                    //      placeholder shape on the layout (preserves direction on
                    //      placeholders that already have text).
                    //   2. Persist a default in the master's <p:txStyles>
                    //      bodyStyle/titleStyle/otherStyle Level1 paragraph
                    //      properties. Blank layouts have no placeholders, so
                    //      this is the only ancestor surface inheriting shapes
                    //      can probe — see ResolveInheritedDirection.
                    bool rtl = key.ToLowerInvariant() == "rtl"
                        ? IsTruthy(value)
                        : ParsePptDirectionRtl(value);
                    var csdShapes = targetRoot.GetFirstChild<CommonSlideData>()?.ShapeTree;
                    if (csdShapes != null)
                    {
                        foreach (var sp in csdShapes.Elements<Shape>())
                        {
                            foreach (var para in sp.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                            {
                                var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                                pProps.RightToLeft = rtl ? (bool?)true : null;
                            }
                        }
                    }
                    // Resolve the master that owns this layout (or self when targetPart
                    // is itself a SlideMasterPart) and write the default into txStyles.
                    SlideMasterPart? mp2 = targetPart switch
                    {
                        SlideLayoutPart lp2 => lp2.SlideMasterPart,
                        SlideMasterPart smp => smp,
                        _ => null,
                    };
                    if (mp2?.SlideMaster is SlideMaster sm)
                    {
                        var txStyles = sm.TextStyles ?? (sm.TextStyles = new TextStyles());
                        void Stamp<T>() where T : OpenXmlCompositeElement, new()
                        {
                            var st = txStyles.GetFirstChild<T>() ?? txStyles.AppendChild(new T());
                            var lvl1 = st.GetFirstChild<Drawing.Level1ParagraphProperties>()
                                ?? st.AppendChild(new Drawing.Level1ParagraphProperties());
                            lvl1.RightToLeft = rtl ? (bool?)true : null;
                        }
                        Stamp<BodyStyle>();
                        Stamp<TitleStyle>();
                        Stamp<OtherStyle>();
                    }
                    break;
                }
                default:
                    if (unsupported.Count == 0)
                        unsupported.Add($"{key} (valid slidemaster/slidelayout props: background, background.mode, background.alpha, background.scale, name, direction)");
                    else
                        unsupported.Add(key);
                    break;
            }
        }
        MaybeMutateExistingBackgroundImage(targetPart, properties);
        SaveBackgroundRoot(targetPart);
        return unsupported;
    }

    private List<string> SetSlideByPath(Match slideOnlyMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
        var slideParts2 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts2.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts2.Count})");
        var slidePart2 = slideParts2[PathIndex.ToArrayIndex(slideIdx)];
        var slide2 = GetSlide(slidePart2);

        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "background":
                    ApplyBackground(slidePart2, value, ReadBackgroundImageOptions(properties));
                    break;
                case "background.mode":
                case "background.alpha":
                case "background.scale":
                    // If paired with "background=", consumed inside the "background" case
                    // via ReadBackgroundImageOptions. Otherwise mutate the existing image
                    // fill in place — done once for the whole property batch, gated below.
                    break;
                case "transition":
                    ApplyTransition(slidePart2, value);
                    if (value.StartsWith("morph", StringComparison.OrdinalIgnoreCase))
                        AutoPrefixMorphNames(slidePart2);
                    else
                        AutoUnprefixMorphNames(slidePart2);
                    break;
                case "transitionspeed":
                    ApplyTransitionSpeed(slidePart2, value);
                    break;
                case "advancetime" or "advanceaftertime":
                    SetAdvanceTime(slide2, value);
                    break;
                case "advanceclick" or "advanceonclick":
                    SetAdvanceClick(slide2, IsTruthy(value));
                    break;
                case "notes":
                {
                    XmlTextValidator.ValidateOrThrow(value, "notes");
                    var notesPart = EnsureNotesSlidePart(slidePart2);
                    SetNotesText(notesPart, value);
                    break;
                }
                case "align":
                {
                    var targets = properties.GetValueOrDefault("targets");
                    AlignShapes(slidePart2, value, targets);
                    break;
                }
                case "distribute":
                {
                    var targets = properties.GetValueOrDefault("targets");
                    DistributeShapes(slidePart2, value, targets);
                    break;
                }
                case "targets":
                    break; // consumed by align/distribute
                case "hidden":
                {
                    // <p:sld show="0"> — hides the slide from slideshow.
                    // Default (Show=null) means visible.
                    if (IsTruthy(value))
                        slide2.Show = false;
                    else
                        slide2.Show = null;
                    break;
                }
                case "showmastershapes":
                {
                    // <p:sld showMasterSp="0"> — suppress master-level
                    // decoration shapes. Default (null) means show. false→"0".
                    if (IsTruthy(value))
                        slide2.ShowMasterShapes = null;
                    else
                        slide2.ShowMasterShapes = false;
                    break;
                }
                case "showfooter":
                case "showslidenumber":
                case "showdate":
                case "showheader":
                {
                    // Toggle header/footer visibility. OOXML CT_Slide does not
                    // permit p:hf as a child of p:sld (validate flags it as an
                    // invalid child element), so visibility is controlled
                    // exclusively by the master-level placeholder presence:
                    // a ftr/dt/sldNum placeholder on the slide master makes
                    // the corresponding footer-row text render; absent
                    // placeholders mean nothing renders. Inject placeholders
                    // on enable. Disable is a UI concept we don't model on the
                    // master (matches PowerPoint UI — "show/hide" doesn't
                    // delete the master shape); the existing master ph stays.
                    bool flag = IsTruthy(value);
                    if (flag && key.ToLowerInvariant() != "showheader")
                    {
                        EnsureMasterFooterPlaceholder(slidePart2, key.ToLowerInvariant());
                    }
                    // Clean up any pre-existing p:hf left on the slide by an
                    // older OfficeCli build — purely defensive so reload of an
                    // older file passes validate after a no-op Set.
                    var legacyHf = slide2.GetFirstChild<HeaderFooter>();
                    if (legacyHf != null) legacyHf.Remove();
                    break;
                }
                case "footertext":
                {
                    // R44 (deferred): we still write the text into the master
                    // ftr placeholder TextBody so the data persists, but real
                    // PowerPoint won't render it without the full placeholder
                    // chain (master ph → layout ph reference → slide-level
                    // <p:hf ftr=1> AND footer instance). All three layers are
                    // schema-constrained in ways that fight each other (p:hf
                    // is invalid on both p:sld AND p:presentation; layout-
                    // level inheritance needs its own placeholder shapes).
                    // Mirrors [[project_deferred_watermark_header_render]]:
                    // render-layer feature pending a dedicated design pass.
                    // For now, accept the input + persist + surface a clear
                    // advisory so callers don't think Set succeeded visually.
                    EnsureMasterFooterPlaceholder(slidePart2, "showfooter");
                    SetMasterFooterPlaceholderText(slidePart2,
                        PlaceholderValues.Footer, value);
                    unsupported.Add("footerText (deferred: text stored in master ftr placeholder, but real PowerPoint requires the full master/layout/slide placeholder chain to render — currently a visual no-op pending render-layer rework)");
                    break;
                }
                case "direction":
                case "dir":
                case "bidi":
                {
                    // R9-bt-3: PPT slides have no slide-level reading direction
                    // — direction is a paragraph-level (txBody/pPr) property.
                    // Reject with a clear pointer instead of silently accepting
                    // or surfacing the unsupported-list dump (which previously
                    // omitted i18n entries from the valid-prop summary).
                    throw new ArgumentException(
                        $"Slide-level '{key}' is not a PPT concept — reading direction is a paragraph property. " +
                        "Apply it per shape: " +
                        $"`set /slide[{slideIdx}]/shape[N]/text/p[M] --prop direction={value}` " +
                        "or set on the txBody bodyPr by setting `direction` on the shape itself.");
                }
                case "name":
                {
                    var csd = slide2.CommonSlideData;
                    if (csd != null)
                    {
                        XmlTextValidator.ValidateOrThrow(value, "name");
                        csd.Name = value;
                    }
                    break;
                }
                case "layout":
                {
                    // Change slide layout. Route through the single resolver so
                    // Set accepts the same grammar Add accepts: display name,
                    // OOXML type token (e.g. "objTx", "blank") or friendly
                    // alias, and 1-based numeric index. Available-list format
                    // is shared too — see ResolveSlideLayout / FormatAvailableLayouts.
                    var presentationPart = _doc.PresentationPart
                        ?? throw new InvalidOperationException("No presentation part");
                    var targetLayout = ResolveSlideLayout(presentationPart, value);
                    if (targetLayout == null)
                        throw new ArgumentException($"Layout '{value}' not found (no layouts defined).");
                    // Point the slide's layout relationship to the new layout
                    if (slidePart2.SlideLayoutPart != null)
                        slidePart2.DeletePart(slidePart2.SlideLayoutPart);
                    slidePart2.AddPart(targetLayout);
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(slide2, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid slide props: background, background.mode, background.alpha, background.scale, layout, transition, name, align, distribute, targets, showFooter, showSlideNumber, showDate, showHeader, showMasterShapes)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        MaybeMutateExistingBackgroundImage(slidePart2, properties);
        slide2.Save();
        return unsupported;
    }

    // When showFooter / showSlideNumber / showDate is toggled on, the slide
    // master must carry a matching placeholder shape — PowerPoint won't
    // render footer-area content from <p:hf> flags alone. Blank documents
    // created by BlankDocCreator have no footer placeholders in the master,
    // so without this helper the toggle was a silent visual no-op.
    //
    // Geometry / positions mirror PowerPoint's default footer-row layout
    // (per CT_HeaderFooter convention): three horizontally-spaced shapes
    // along the bottom of the slide. EMU values match the standard 16:9
    // template ("Office Theme") so the injection looks native.
    private static void EnsureMasterFooterPlaceholder(SlidePart slidePart, string flagKey)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        var masterPart = layoutPart?.SlideMasterPart;
        if (masterPart?.SlideMaster == null) return;

        var master = masterPart.SlideMaster;
        var shapeTree = master.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return;

        PlaceholderValues phType;
        long offsetX, offsetY, extentCx, extentCy;
        string phName;
        uint phIdx;
        // 16:9 widescreen master: cx=12192000 EMU (~33.87cm), cy=6858000 EMU (~19.05cm)
        // Footer row sits at y ≈ 6356000 EMU with cy ≈ 365125 (~1cm).
        switch (flagKey)
        {
            case "showfooter":
                phType = PlaceholderValues.Footer;
                phName = "Footer Placeholder";
                phIdx = 11;
                offsetX = 4040188; offsetY = 6356350;
                extentCx = 4111625; extentCy = 365125;
                break;
            case "showdate":
                phType = PlaceholderValues.DateAndTime;
                phName = "Date Placeholder";
                phIdx = 10;
                offsetX = 838200; offsetY = 6356350;
                extentCx = 2895600; extentCy = 365125;
                break;
            case "showslidenumber":
                phType = PlaceholderValues.SlideNumber;
                phName = "Slide Number Placeholder";
                phIdx = 12;
                offsetX = 8470900; offsetY = 6356350;
                extentCx = 2895600; extentCy = 365125;
                break;
            default:
                return;
        }

        // Skip if a placeholder of this type already lives on the master.
        var existing = shapeTree.Descendants<PlaceholderShape>()
            .Any(ph => ph.Type != null && ph.Type.Value == phType);
        if (existing) return;

        // Pick a fresh cNvPr id higher than anything else on the master so we
        // don't collide with existing shape ids (master placeholders typically
        // use small ids 2..N).
        uint nextId = 1;
        foreach (var nvDp in shapeTree.Descendants<NonVisualDrawingProperties>())
        {
            if (nvDp.Id?.Value is uint v && v >= nextId) nextId = v + 1;
        }
        if (nextId < 100) nextId = 100;

        var shape = new Shape();
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = nextId, Name = phName },
            new NonVisualShapeDrawingProperties(
                new DocumentFormat.OpenXml.Drawing.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(
                new PlaceholderShape { Type = phType, Index = phIdx, Size = PlaceholderSizeValues.Quarter })
        );
        shape.ShapeProperties = new ShapeProperties(
            new DocumentFormat.OpenXml.Drawing.Transform2D(
                new DocumentFormat.OpenXml.Drawing.Offset { X = offsetX, Y = offsetY },
                new DocumentFormat.OpenXml.Drawing.Extents { Cx = extentCx, Cy = extentCy }
            )
        );
        shape.TextBody = new TextBody(
            new DocumentFormat.OpenXml.Drawing.BodyProperties(),
            new DocumentFormat.OpenXml.Drawing.ListStyle(),
            new DocumentFormat.OpenXml.Drawing.Paragraph(
                new DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties { Language = "en-US" })
        );
        shapeTree.AppendChild(shape);
        master.Save();
    }

    // Stamp footer text into the master ftr placeholder's TextBody. The
    // placeholder shape must already exist (caller's responsibility — usually
    // by calling EnsureMasterFooterPlaceholder("showfooter") first). Replaces
    // any prior text content with a single run, so re-setting overwrites.
    private static void SetMasterFooterPlaceholderText(SlidePart slidePart,
        PlaceholderValues phType, string text)
    {
        var master = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster;
        if (master == null) return;
        var phShape = master.CommonSlideData?.ShapeTree?
            .Descendants<Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<PlaceholderShape>()?.Type?.Value == phType);
        if (phShape == null) return;
        var txBody = phShape.TextBody;
        if (txBody == null)
        {
            txBody = new TextBody(
                new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                new DocumentFormat.OpenXml.Drawing.ListStyle());
            phShape.TextBody = txBody;
        }
        // Strip existing paragraphs and emit a single <a:p><a:r><a:t>{text}</a:t></a:r></a:p>.
        // BodyProperties + ListStyle are preserved so any auto-fit / wrap
        // configuration survives.
        foreach (var p in txBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList())
            p.Remove();
        var run = new DocumentFormat.OpenXml.Drawing.Run(
            new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US" },
            new DocumentFormat.OpenXml.Drawing.Text(text));
        txBody.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph(run));
        master.Save();
    }

}
