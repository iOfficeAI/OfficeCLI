// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

/// <summary>
/// SlideComposition — the shared interpreter of PowerPoint's master/layout/slide
/// inheritance model (issue #466).
///
/// A slide is stored NOWHERE in the file: slide.xml holds only overrides, and
/// what a viewer displays is the composition of three shape trees painted
/// master (bottom) → layout → slide (top), with placeholder shapes inheriting
/// from the layout by idx and from the master by type. Until now each consumer
/// (HTML preview, SVG preview, get/query) re-implemented — or simply skipped —
/// that interpretation, and the copies drifted: the HTML preview resolved
/// inherited positions for shapes but not pictures, the SVG preview reached
/// into the HTML renderer's file for the same function, and get reported no
/// geometry at all for slot-bound shapes (issue #466, symptom 6).
///
/// This module is the one place that resolves the model; consumers read from
/// it instead of improvising. Scope of this first extraction: placeholder
/// slot matching, inherited-frame resolution and its provenance. Later steps
/// per issue #466: layer paint order, hf/showMasterSp/hidden gates, effective
/// fill inheritance.
///
/// Frame semantics (ECMA-376): &lt;a:xfrm&gt; is ATOMIC. A placeholder either
/// carries a complete xfrm (owns its frame) or none (inherits the whole frame
/// from the matching layout/master slot). A partial xfrm is not schema-valid
/// — off requires both x and y, ext requires both cx and cy — so "half
/// inherited" is not a state the model has.
/// </summary>
internal static class SlideComposition
{
    // ==================== Placeholder Slot Matching ====================

    /// <summary>
    /// Check if two placeholder shapes match by type and/or index.
    /// </summary>
    public static bool PlaceholderMatches(PlaceholderShape slidePh, PlaceholderShape layoutPh)
    {
        // Match by index first (most specific)
        if (slidePh.Index?.HasValue == true && layoutPh.Index?.HasValue == true)
            return slidePh.Index.Value == layoutPh.Index.Value;

        // Match by type — a family counts as the same type. ECMA-376 §19.7.10
        // spells the title-slide slots `ctrTitle` / `subTitle` in a LAYOUT and
        // `title` / `body` on the MASTER; each pair is one inheritance slot, and
        // PowerPoint and LibreOffice resolve the layout spelling through the
        // master's. Without this a layout's ctrTitle (which carries no idx)
        // never reaches the master's title placeholder, so that placeholder's
        // own lstStyle — a cascade layer — is silently skipped.
        if (slidePh.Type?.HasValue == true && layoutPh.Type?.HasValue == true)
            return SamePlaceholderFamily(slidePh.Type.Value, layoutPh.Type.Value);

        // R26-5: slide ph has idx but NO type, layout ph has type but NO idx.
        // OOXML: a <p:ph idx=N/> with no type defaults to type=body, so it
        // should inherit from a type=body (or object) layout/master placeholder.
        // Without this branch all inheritance silently drops for idx-only slide
        // placeholders bound to a typed layout placeholder.
        if (slidePh.Index?.HasValue == true && slidePh.Type?.HasValue != true
            && layoutPh.Type?.HasValue == true && layoutPh.Index?.HasValue != true)
        {
            var lt = layoutPh.Type.Value;
            return lt == PlaceholderValues.Body || lt == PlaceholderValues.Object;
        }

        // If slide ph has no type/idx, match by name or consider it a body placeholder
        // Default placeholder type (when type is omitted) is "body" per OOXML spec
        if (slidePh.Type?.HasValue != true && slidePh.Index?.HasValue != true)
        {
            // A typeless/indexless placeholder matches title if the layout has title,
            // or body/subtitle by convention
            if (layoutPh.Type?.HasValue == true)
            {
                var lt = layoutPh.Type.Value;
                return lt == PlaceholderValues.Title || lt == PlaceholderValues.CenteredTitle
                    || lt == PlaceholderValues.SubTitle || lt == PlaceholderValues.Body;
            }
        }

        return false;
    }

    /// <summary>
    /// Whether two placeholder types name the same inheritance slot.
    ///
    /// ECMA-376 §19.7.10 spells the title-slide slots differently in the layout
    /// and on the master: a layout declares &lt;p:ph type="ctrTitle"/&gt; and
    /// &lt;p:ph type="subTitle" idx="1"/&gt; where the master declares
    /// &lt;p:ph type="title"/&gt; and &lt;p:ph type="body" idx="1"/&gt;.
    /// PowerPoint and LibreOffice treat each pair as one slot, so a placeholder
    /// inherits across the two spellings. This is the same partition the cascade
    /// already uses to pick a master text style (titleStyle covers
    /// Title/CenteredTitle, bodyStyle covers Body/SubTitle/Object) — here it
    /// decides which master placeholder an inheritance step reaches.
    /// </summary>
    private static bool SamePlaceholderFamily(PlaceholderValues a, PlaceholderValues b)
    {
        if (a == b) return true;

        static bool IsTitleFamily(PlaceholderValues v)
            => v == PlaceholderValues.Title || v == PlaceholderValues.CenteredTitle;

        static bool IsBodyFamily(PlaceholderValues v)
            => v == PlaceholderValues.Body || v == PlaceholderValues.SubTitle;

        return (IsTitleFamily(a) && IsTitleFamily(b)) || (IsBodyFamily(a) && IsBodyFamily(b));
    }

    // ==================== Inherited Frame Resolution ====================

    /// <summary>
    /// A resolved frame plus WHERE it came from. get/query reports the
    /// provenance as <c>effective.*.src</c> (same convention StyleList uses for
    /// inherited text properties), and the write path will need the same source
    /// shape to materialize an explicit frame before applying a delta — xfrm is
    /// atomic, so a partial write is schema-invalid and must start from the
    /// inherited frame.
    /// </summary>
    public sealed record InheritedFrame(
        long X, long Y, long Cx, long Cy,
        Shape SourceShape, PlaceholderShape SourcePh, bool FromMaster);

    /// <summary>
    /// Resolve the frame a placeholder inherits from its layout/master slot.
    /// Shared entry point for shapes (&lt;p:sp&gt;) and pictures (&lt;p:pic&gt;) — a
    /// picture filled into a picture placeholder carries the same &lt;p:ph&gt; and an
    /// empty spPr, so it inherits its frame the same way. Walks the layout tree
    /// first, then the master tree, and returns the FIRST slot that matches AND
    /// carries an xfrm with both offset and extents. Returns null for
    /// non-placeholders or when no slot with a frame resolves.
    /// </summary>
    public static InheritedFrame? ResolveInheritedFrame(PlaceholderShape? ph, SlidePart slidePart)
    {
        if (ph == null) return null;

        var layoutShapeTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
        var masterShapeTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;

        var fromMaster = false;
        foreach (var tree in new[] { layoutShapeTree, masterShapeTree })
        {
            if (tree == null) continue;
            foreach (var candidate in tree.Elements<Shape>())
            {
                var candidatePh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                if (candidatePh == null) continue;

                if (!PlaceholderMatches(ph, candidatePh)) continue;

                var cxfrm = candidate.ShapeProperties?.Transform2D;
                if (cxfrm?.Offset != null && cxfrm?.Extents != null)
                {
                    return new InheritedFrame(
                        cxfrm.Offset.X?.Value ?? 0,
                        cxfrm.Offset.Y?.Value ?? 0,
                        cxfrm.Extents.Cx?.Value ?? 0,
                        cxfrm.Extents.Cy?.Value ?? 0,
                        candidate, candidatePh, fromMaster);
                }
            }
            fromMaster = true;
        }

        return null;
    }

    /// <summary>
    /// When a shape has no Transform2D, try to find position from matching placeholder
    /// on the slide layout or slide master (OOXML placeholder inheritance chain).
    /// </summary>
    public static (long x, long y, long cx, long cy)? ResolveInheritedPosition(Shape shape, OpenXmlPart part)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        // Only placeholder shapes can inherit position from layout/master
        if (ph == null) return null;

        var slidePart = part as SlidePart;
        if (slidePart == null) return null;

        var frame = ResolveInheritedFrame(ph, slidePart);
        return frame == null ? null : (frame.X, frame.Y, frame.Cx, frame.Cy);
    }

    /// <summary>
    /// R12-5: find the layout (then master) placeholder shape that the given
    /// slide placeholder inherits from. Same ph type/idx matching as
    /// ResolveInheritedPosition, but returns the whole shape so callers can
    /// read inherited spPr fill/etc. Returns null for non-placeholders.
    /// </summary>
    public static Shape? ResolveInheritedPlaceholderShape(Shape shape, OpenXmlPart part)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return null;

        var slidePart = part as SlidePart;
        if (slidePart == null) return null;

        var layoutShapeTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
        var masterShapeTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;

        foreach (var tree in new[] { layoutShapeTree, masterShapeTree })
        {
            if (tree == null) continue;
            foreach (var candidate in tree.Elements<Shape>())
            {
                var candidatePh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                if (candidatePh == null) continue;
                if (PlaceholderMatches(ph, candidatePh)) return candidate;
            }
        }

        return null;
    }

    // ==================== Own-Frame Completeness ====================

    /// <summary>
    /// Whether a shape's own xfrm is COMPLETE: off carries both x and y, ext
    /// carries both cx and cy (the schema-valid form — all four attributes are
    /// required). A partial xfrm (e.g. left behind by a piecemeal geometry Set)
    /// is not schema-valid: consumers treat the shape as NOT owning its frame
    /// and resolve the effective frame from the slot instead.
    /// </summary>
    public static bool HasCompleteOwnFrame(Drawing.Transform2D? xfrm)
        => xfrm?.Offset?.X?.HasValue == true && xfrm.Offset.Y?.HasValue == true
           && xfrm.Extents?.Cx?.HasValue == true && xfrm.Extents.Cy?.HasValue == true;

    // ==================== Provenance Vocabulary ====================

    /// <summary>
    /// The get/query provenance path for an inherited frame, e.g.
    /// <c>/slide[1]/layout/ph[@type=pic][@idx=1]</c> — the same vocabulary the
    /// text-style layer already uses for <c>effective.*.src</c>. Pointers cross
    /// parts on purpose: the owner of an inherited frame lives in the layout or
    /// master, not on the slide.
    /// </summary>
    public static string ProvenancePath(int slideNum, InheritedFrame frame)
    {
        var level = frame.FromMaster ? "master" : "layout";
        var predicate = "";
        var typeText = frame.SourcePh.Type?.InnerText;
        if (!string.IsNullOrEmpty(typeText))
            predicate += $"[@type={typeText}]";
        if (frame.SourcePh.Index?.HasValue == true)
            predicate += $"[@idx={frame.SourcePh.Index.Value}]";
        return $"/slide[{slideNum}]/{level}/ph{predicate}";
    }
}
