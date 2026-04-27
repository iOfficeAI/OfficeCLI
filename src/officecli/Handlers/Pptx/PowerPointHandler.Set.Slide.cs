// Copyright 2025 OfficeCli (officecli.ai)
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
        foreach (var (key, value) in properties)
        {
            if (key.Equals("text", StringComparison.OrdinalIgnoreCase))
                SetNotesText(notesPart, value);
            else
                unsupportedN.Add(key);
        }
        return unsupportedN;
    }

    private List<string> SetMasterShapeByPath(Match masterShapeMatch, Dictionary<string, string> properties)
    {
        var partType = masterShapeMatch.Groups[1].Value;
        var partIdx = int.Parse(masterShapeMatch.Groups[2].Value);
        var presentationPart = _doc.PresentationPart!;

        OpenXmlPartRootElement rootEl;
        if (partType == "slideMaster")
        {
            var masters = presentationPart.SlideMasterParts.ToList();
            if (partIdx < 1 || partIdx > masters.Count)
                throw new ArgumentException($"SlideMaster {partIdx} not found (total: {masters.Count})");
            rootEl = masters[partIdx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt slide master");
        }
        else
        {
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (partIdx < 1 || partIdx > layouts.Count)
                throw new ArgumentException($"SlideLayout {partIdx} not found (total: {layouts.Count})");
            rootEl = layouts[partIdx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt slide layout");
        }

        if (!masterShapeMatch.Groups[3].Success)
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

        // Set on a specific shape within master/layout
        var elType = masterShapeMatch.Groups[3].Value;
        var elIdx = int.Parse(masterShapeMatch.Groups[4].Value);
        var shapeTree = rootEl.Descendants<ShapeTree>().FirstOrDefault()
            ?? throw new ArgumentException("No shape tree found");

        if (elType == "shape")
        {
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (elIdx < 1 || elIdx > shapes.Count)
                throw new ArgumentException($"Shape {elIdx} not found");
            var shape = shapes[elIdx - 1];
            var allRuns = shape.Descendants<Drawing.Run>().ToList();
            var unsupp = SetRunOrShapeProperties(properties, allRuns, shape);
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
                    if (csd != null) csd.Name = value;
                    break;
                }
                default:
                    if (unsupported.Count == 0)
                        unsupported.Add($"{key} (valid slidemaster/slidelayout props: background, background.mode, background.alpha, background.scale, name)");
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
                case "layout":
                {
                    // Change slide layout
                    var presentationPart = _doc.PresentationPart
                        ?? throw new InvalidOperationException("No presentation part");
                    var allLayouts = presentationPart.SlideMasterParts
                        .SelectMany(m => m.SlideLayoutParts).ToList();
                    var targetLayout = allLayouts.FirstOrDefault(lp =>
                        lp.SlideLayout?.CommonSlideData?.Name?.Value?.Equals(value, StringComparison.OrdinalIgnoreCase) == true);
                    if (targetLayout == null)
                    {
                        var availableNames = allLayouts
                            .Select(lp => lp.SlideLayout?.CommonSlideData?.Name?.Value)
                            .Where(n => n != null)
                            .ToList();
                        throw new ArgumentException($"Layout '{value}' not found. Available layouts: {string.Join(", ", availableNames)}");
                    }
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
