// Copyright 2025 OfficeCLI (officecli.ai)
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
        var notesPart = EnsureNotesSlidePart(slidePartsN[slideIdx - 1]);
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
                XmlTextValidator.ValidateOrThrow(value, "text");
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
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
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
                throw new ArgumentException($"Shape {elIdx} not found");
            var shape = shapes[elIdx - 1];
            var allRuns = shape.Descendants<Drawing.Run>().ToList();
            // Pass the owning part so fill/image/effect helpers that need a
            // relationship anchor (e.g. picture fills) write to the correct part.
            var unsupp = SetRunOrShapeProperties(properties, allRuns, shape, ownerPart);
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
        var slidePart2 = slideParts2[slideIdx - 1];
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
                case "showfooter":
                case "showslidenumber":
                case "showdate":
                case "showheader":
                {
                    // Toggle header/footer visibility flags on the slide.
                    // Emits <p:hf ftr="1" sldNum="0" dt="1" hdr="0"/> as a
                    // direct child of <p:sld>. The OpenXml SDK models this
                    // via DocumentFormat.OpenXml.Presentation.HeaderFooter
                    // (local name "hf"). Although CT_Slide's published
                    // schema does not list hf, PowerPoint itself writes it
                    // on slides when the "Insert > Header & Footer" dialog
                    // toggles per-slide overrides — we mirror that.
                    var hf = slide2.GetFirstChild<HeaderFooter>() ?? new HeaderFooter();
                    bool isNew = hf.Parent == null;
                    bool flag = IsTruthy(value);
                    switch (key.ToLowerInvariant())
                    {
                        case "showfooter": hf.Footer = flag; break;
                        case "showslidenumber": hf.SlideNumber = flag; break;
                        case "showdate": hf.DateTime = flag; break;
                        case "showheader": hf.Header = flag; break;
                    }
                    if (isNew) slide2.AppendChild(hf);
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
                            unsupported.Add($"{key} (valid slide props: background, background.mode, background.alpha, background.scale, layout, transition, name, align, distribute, targets, showFooter, showSlideNumber, showDate, showHeader)");
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

}
