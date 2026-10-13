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
/// per issue #466: showMasterSp handling inside the module, effective
/// fill inheritance.
///
/// Visibility gates: a shape marked &lt;p:cNvPr hidden="1"&gt; is not drawn. The
/// attribute lives on &lt;a:CT_NonVisualDrawingProps&gt;, so it travels with the
/// shape at every layer of the composition — a hidden master decoration would
/// otherwise leak onto every slide — and it is not the same bit as the
/// slide-level &lt;p:sld show="0"&gt; that hides a slide from a show.
///
/// Frame semantics (ECMA-376): &lt;a:xfrm&gt; is ATOMIC. A placeholder either
/// carries a complete xfrm (owns its frame) or none (inherits the whole frame
/// from the matching layout/master slot). A partial xfrm is not schema-valid
/// — off requires both x and y, ext requires both cx and cy — so "half
/// inherited" is not a state the model has.
/// </summary>
internal static class SlideComposition
{
    // ==================== Visibility Gates ====================

    /// <summary>
    /// Whether a shape is marked hidden in its own <c>&lt;p:cNvPr hidden="1"&gt;</c>.
    /// The gate travels with the shape through the whole composition: a hidden
    /// master or layout decoration must not reach a slide, and a hidden shape on
    /// the slide itself is just as absent from PowerPoint's output. Consumers
    /// that paint shapes skip these; consumers that read the model (get/query)
    /// still see them, because the shape is in the file either way.
    /// </summary>
    public static bool IsHidden(OpenXmlElement? element)
    {
        var cNvPr = element switch
        {
            Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
            Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
            ConnectionShape connector => connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
            GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
            GroupShape group => group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
            _ => null,
        };
        return cNvPr?.Hidden?.Value == true;
    }

    /// <summary>
    /// Whether the master's own shapes show on this slide. <c>showMasterSp</c>
    /// on the slide wins, else the layout's, else shown. A layout authored with
    /// <c>showMasterSp="0"</c> hides the master decoration on every slide that
    /// uses it; the master's placeholders and the layout's shapes are unaffected.
    /// </summary>
    public static bool ShowMasterShapes(SlidePart slidePart)
    {
        var slideFlag = slidePart.Slide?.ShowMasterShapes;
        if (slideFlag?.HasValue == true) return slideFlag.Value;
        var layoutFlag = slidePart.SlideLayoutPart?.SlideLayout?.ShowMasterShapes;
        if (layoutFlag?.HasValue == true) return layoutFlag.Value;
        return true;
    }

    // ==================== Header/Footer Gates ====================

    /// <summary>
    /// The metadata slots (<c>dt</c>, <c>ftr</c>, <c>hdr</c>, <c>sldNum</c>) a slide
    /// does NOT display, as <c>type:&lt;ph type&gt;</c> skip keys for the inherited
    /// passes. &lt;p:hf&gt; on the layout overrides the master's; an absent flag
    /// means shown. The master's footer placeholder holds the footer TEXT, but
    /// &lt;p:hf ftr="0"/&gt; means "do not display it" — without this gate every
    /// slide inherited the text. A slide's own metadata shape is not gated: it
    /// was placed on that slide deliberately.
    /// </summary>
    public static IEnumerable<string> HiddenMetadataSlotKeys(SlidePart slidePart)
    {
        var layoutHf = slidePart.SlideLayoutPart?.SlideLayout?.GetFirstChild<HeaderFooter>();
        var masterHf = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.GetFirstChild<HeaderFooter>();
        static bool Shown(BooleanValue? layout, BooleanValue? master)
            => layout?.HasValue == true ? layout.Value : master?.HasValue != true || master.Value;
        if (!Shown(layoutHf?.DateTime, masterHf?.DateTime)) yield return "type:dt";
        if (!Shown(layoutHf?.Footer, masterHf?.Footer)) yield return "type:ftr";
        if (!Shown(layoutHf?.Header, masterHf?.Header)) yield return "type:hdr";
        if (!Shown(layoutHf?.SlideNumber, masterHf?.SlideNumber)) yield return "type:sldNum";
    }

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

        // Same partition as the cascade's bodyStyle (Body / SubTitle / Object)
        // and the idx-only branch above, so one family rule covers all three.
        static bool IsBodyFamily(PlaceholderValues v)
            => v == PlaceholderValues.Body || v == PlaceholderValues.SubTitle || v == PlaceholderValues.Object;

        return (IsTitleFamily(a) && IsTitleFamily(b)) || (IsBodyFamily(a) && IsBodyFamily(b));
    }

    /// <summary>
    /// The layout→master edge of slot matching: BY TYPE, not idx. The master
    /// numbers its placeholders independently of the layout, so idx never
    /// correlates across that edge — a layout's sldNum idx=12 inherits from
    /// the master's sldNum idx=4. Types map at the master level per
    /// PowerPoint: ctrTitle→title, subTitle/obj/pic/chart/tbl→body; the
    /// metadata types (dt/ftr/hdr/sldNum) keep their own.
    /// </summary>
    public static bool MatchesMasterSlot(PlaceholderShape ph, PlaceholderShape masterPh)
    {
        var a = ph.Type?.HasValue == true ? ph.Type.Value : PlaceholderValues.Object;
        var b = masterPh.Type?.HasValue == true ? masterPh.Type.Value : PlaceholderValues.Object;
        return MasterLevelType(a) == MasterLevelType(b);
    }

    // Master-level type mapping: the master's placeholder vocabulary is
    // coarser than the layout's. A typeless <p:ph> is obj (ECMA-376 default)
    // and lands on body. (PlaceholderValues is a struct of static properties
    // in the SDK, not an enum — switch cases need constants, so == it is.)
    private static PlaceholderValues MasterLevelType(PlaceholderValues t)
    {
        if (t == PlaceholderValues.CenteredTitle) return PlaceholderValues.Title;
        if (t == PlaceholderValues.SubTitle
            || t == PlaceholderValues.Object
            || t == PlaceholderValues.Picture
            || t == PlaceholderValues.Chart
            || t == PlaceholderValues.Table)
            return PlaceholderValues.Body;
        return t;
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
    /// slide placeholder inherits from. Same two edges as
    /// ResolveInheritedPosition — slide→layout by idx, layout→master by type —
    /// but returns the whole shape so callers can read inherited spPr fill/etc.
    /// Returns null for non-placeholders.
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

        foreach (var (tree, toMaster) in new[] { (layoutShapeTree, false), (masterShapeTree, true) })
        {
            if (tree == null) continue;
            foreach (var candidate in tree.Elements<Shape>())
            {
                var candidatePh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                if (candidatePh == null) continue;
                if (toMaster ? MatchesMasterSlot(ph, candidatePh) : PlaceholderMatches(ph, candidatePh))
                    return candidate;
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
