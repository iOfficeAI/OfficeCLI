// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string AddConnector(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Accept a slide (/slide[N]) or a nested-group parent
                // (/slide[N]/group[K]/…) so dump-emitted grouped connectors replay.
                var cxnParent = ResolveSlideOrGroupAddParent(parentPath)
                    ?? throw new ArgumentException("Connectors must be added to a slide: /slide[N]");
                var cxnSlidePart = cxnParent.slidePart;
                var cxnShapeTree = cxnParent.shapeTree;
                var cxnInsertContainer = cxnParent.insertContainer;
                var cxnReturnPrefix = cxnParent.returnPathPrefix;

                var cxnId = AcquireShapeId(cxnShapeTree, properties);
                var cxnName = properties.GetValueOrDefault("name", $"Connector {cxnShapeTree.Elements<ConnectionShape>().Count() + 1}");

                // Position: explicit x/y/width/height OR derived from connected shapes.
                // When from=/to= reference existing shapes and x/y/width/height are
                // omitted, compute the connector's bounding box from the two shapes'
                // centers so the rendered line actually spans the gap between them.
                // PowerPoint does NOT recompute connector geometry from stCxn/endCxn
                // at render time — it trusts our offset/extent — so a missing default
                // here paints the connector at a hard-coded stub near the slide center.
                var hasX = properties.ContainsKey("x") || properties.ContainsKey("left");
                var hasY = properties.ContainsKey("y") || properties.ContainsKey("top");
                var hasW = properties.ContainsKey("width");
                var hasH = properties.ContainsKey("height");
                // Look up a frame's (x,y,width,height) by OOXML shape ID across
                // every connectable container element (Shape, Picture, GraphicFrame,
                // ConnectionShape, GroupShape) — same set ResolveShapeId+AddGroup
                // accepts so connector from=/to= works against the full frame list.
                static (long x, long y, long cx, long cy)? GetFrameBoundsById(ShapeTree tree, uint id)
                    => FrameBoundsById(tree, id);

                var hasFrom = properties.ContainsKey("from") || properties.ContainsKey("startshape") || properties.ContainsKey("startShape");
                var hasTo = properties.ContainsKey("to") || properties.ContainsKey("endshape") || properties.ContainsKey("endShape");

                long cxnX = (properties.TryGetValue("x", out var cx1) || properties.TryGetValue("left", out cx1)) ? ParseEmu(cx1) : 2000000;
                long cxnY = (properties.TryGetValue("y", out var cy1) || properties.TryGetValue("top", out cy1)) ? ParseEmu(cy1) : 3000000;
                long cxnCx = properties.TryGetValue("width", out var cw) ? ParseEmu(cw) : 4000000;
                long cxnCy = properties.TryGetValue("height", out var ch) ? ParseEmu(ch) : 0;
                var cxnFlipH = false;
                var cxnFlipV = false;
                // Explicit flipH / flipV props mirror the shape Add convention
                // (Add.Shape routes "fliph"/"flipv" through SetRunOrShapeProperties).
                // Connectors previously dropped them as unsupported, so a
                // right-to-left diagonal connector could never be authored and the
                // SVG always drew top-left→bottom-right. The HTML renderer already
                // honors <a:xfrm flipH/flipV>; this just lets users set it.
                // Track explicit presence (not just truthiness) so an explicit
                // flipH/flipV survives — and can override — the geometry-derived
                // flip below. Without this an explicit value passed alongside
                // from/to+sides was silently clobbered by the derived ordering.
                bool hasFlipH = properties.TryGetValue("flipH", out var cxnFlipHRaw) || properties.TryGetValue("fliph", out cxnFlipHRaw);
                if (hasFlipH) cxnFlipH = IsTruthy(cxnFlipHRaw);
                bool hasFlipV = properties.TryGetValue("flipV", out var cxnFlipVRaw) || properties.TryGetValue("flipv", out cxnFlipVRaw);
                if (hasFlipV) cxnFlipV = IsTruthy(cxnFlipVRaw);
                if ((hasFrom || hasTo) && !(hasX && hasY && hasW && hasH))
                {
                    var startRef = properties.GetValueOrDefault("from")
                        ?? properties.GetValueOrDefault("startShape")
                        ?? properties.GetValueOrDefault("startshape");
                    var endRef = properties.GetValueOrDefault("to")
                        ?? properties.GetValueOrDefault("endShape")
                        ?? properties.GetValueOrDefault("endshape");
                    var startBox = startRef != null ? GetFrameBoundsById(cxnShapeTree, ResolveShapeId(startRef, cxnShapeTree)) : null;
                    var endBox = endRef != null ? GetFrameBoundsById(cxnShapeTree, ResolveShapeId(endRef, cxnShapeTree)) : null;
                    var pStart = startBox ?? endBox;
                    var pEnd = endBox ?? startBox;
                    if (pStart.HasValue && pEnd.HasValue)
                    {
                        var fromSide = NormalizeConnectorSide(properties, "fromSide", "fromside", "startSide", "startside");
                        var toSide = NormalizeConnectorSide(properties, "toSide", "toside", "endSide", "endside");
                        var (p1x, p1y, p2x, p2y) = ComputeConnectorEndpoints(
                            pStart.Value, pEnd.Value, fromSide, toSide);
                        if (!hasX) cxnX = Math.Min(p1x, p2x);
                        if (!hasY) cxnY = Math.Min(p1y, p2y);
                        if (!hasW) cxnCx = Math.Abs(p2x - p1x);
                        if (!hasH) cxnCy = Math.Abs(p2y - p1y);
                        // Encode start/end ordering via flipH/flipV (mirrors PowerPoint),
                        // unless the caller supplied an explicit flip — then honor theirs.
                        if (!hasFlipH) cxnFlipH = p2x < p1x;
                        if (!hasFlipV) cxnFlipV = p2y < p1y;
                    }
                }
                // CONSISTENCY(positive-size): mirror Add.Shape negative-size guard so picture
                // / chart / connector / media all reject inverted dimensions instead of silently
                // emitting negative cx/cy that PowerPoint draws as flipped or 0-sized boxes.
                if (cxnCx < 0) throw new ArgumentException($"Negative width is not allowed: '{cw}'.");
                if (cxnCy < 0) throw new ArgumentException($"Negative height is not allowed: '{ch}'.");
                // The derived span between two far-apart endpoints can exceed the
                // int32 EMU ceiling even when each shape's own coordinates are in
                // range — the SDK would then write an <a:ext> that fails schema
                // validation. Guard the derived extent the same way individual
                // shape coordinates are guarded.
                if (cxnCx > int.MaxValue || cxnCy > int.MaxValue)
                    throw new ArgumentException(
                        "Connector span exceeds the maximum supported coordinate (INT32_MAX EMU); "
                        + "the connected shapes are too far apart.");

                var connector = new ConnectionShape();
                var cxnNvProps = new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = cxnId, Name = cxnName },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );

                // Connect to shapes if specified
                var cxnDrawProps = cxnNvProps.NonVisualConnectorShapeDrawingProperties!;
                // bt-6: <a:stCxn idx="M"/> / <a:endCxn idx="M"/> pins the
                // connector to a specific anchor on the target shape; without
                // honoring fromIdx / toIdx (the dump→replay-aligned keys
                // PptxBatchEmitter emits) every connector landed on anchor 0,
                // breaking source-authored diagram routing.
                // NOTE(idx-from-side, deferred): fromSide/toSide drive the drawn
                // geometry (offset/extent) above but are NOT translated to a stCxn
                // idx here — the drawn line is already correct (PowerPoint renders
                // our offset/extent, not the idx), and a per-preset cxnLst ordering
                // map cannot be verified by rendering. Until that map exists, idx
                // comes only from an explicit fromIdx/toIdx.
                static uint ParseCxnIdx(Dictionary<string, string> p, params string[] keys)
                {
                    foreach (var k in keys)
                    {
                        if (!p.TryGetValue(k, out var raw)) continue;
                        var t = raw?.Trim();
                        if (string.IsNullOrEmpty(t)) continue;
                        // Reject non-numeric / negative / overflow input instead of
                        // silently coercing to 0 — mirror NormalizeConnectorSide's
                        // strict validation so fromIdx/toIdx aren't the lone lenient
                        // connector prop. (Bounds vs the preset's cxnLst site count
                        // is not checked here — that needs a per-preset map.)
                        if (uint.TryParse(t, System.Globalization.NumberStyles.Integer,
                                System.Globalization.CultureInfo.InvariantCulture, out var v))
                            return v;
                        throw new ArgumentException(
                            $"Invalid connector index '{raw}' for '{k}': expected a non-negative integer.");
                    }
                    return 0;
                }
                // A from/to reference must resolve to a real, reachable frame on
                // this container. ResolveShapeId's plain-int path deliberately
                // trusts an unmatched id (dump-replay forward refs), so a garbage
                // id — or a shape nested inside a group, which FrameBoundsById does
                // not descend into — would otherwise be written as a dangling
                // stCxn/endCxn and paint a degenerate connector collapsed onto the
                // one endpoint that did resolve. Fail loudly instead, matching the
                // @id/@name path forms which already throw when not found.
                // Only enforce when we derive geometry from the frames (no full
                // explicit box). Dump→batch replay always emits explicit x/y/w/h and
                // may legitimately reference a shape added later in the same batch, so
                // gating on the derivation path preserves forward-ref round-trips while
                // still catching interactive garbage refs.
                bool derivesGeometry = !(hasX && hasY && hasW && hasH);
                void RequireReachableFrame(uint id, string reference, string key)
                {
                    if (derivesGeometry && GetFrameBoundsById(cxnShapeTree, id) == null)
                        throw new ArgumentException(
                            $"Connector '{key}={reference}' does not resolve to a shape on this slide "
                            + "(no top-level frame with that id/index exists; note a shape nested inside a group cannot be a connector endpoint).");
                }
                // A connector is slide-local: stCxn/endCxn ids are scoped to the
                // connector's own slide, so a from/to path naming a different slide
                // (e.g. to=/slide[2]/shape[1]) can't be honored — ResolveShapeId
                // ignores the slide token and would silently bind to the same-index
                // shape on THIS slide. Reject it instead of misresolving.
                var cxnSlideNum = int.TryParse(
                    Regex.Match(parentPath, @"/slide\[(\d+)\]").Groups[1].Value, out var csn) ? csn : 1;
                void RejectCrossSlideRef(string reference, string key)
                {
                    var m = Regex.Match(reference, @"^/slide\[(\d+)\]/");
                    if (m.Success && int.Parse(m.Groups[1].Value) != cxnSlideNum)
                        throw new ArgumentException(
                            $"Connector '{key}={reference}' references a different slide; a connector can "
                            + $"only attach to shapes on its own slide (slide {cxnSlideNum}).");
                }
                if (properties.TryGetValue("startshape", out var startId) || properties.TryGetValue("startShape", out startId)
                    || properties.TryGetValue("from", out startId))
                {
                    RejectCrossSlideRef(startId!, "from");
                    var startIdVal = ResolveShapeId(startId!, cxnShapeTree);
                    RequireReachableFrame(startIdVal, startId!, "from");
                    var startIdxVal = ParseCxnIdx(properties, "fromIdx", "fromidx", "startIdx", "startidx");
                    cxnDrawProps.StartConnection = new Drawing.StartConnection { Id = startIdVal, Index = startIdxVal };
                }
                if (properties.TryGetValue("endshape", out var endId) || properties.TryGetValue("endShape", out endId)
                    || properties.TryGetValue("to", out endId))
                {
                    RejectCrossSlideRef(endId!, "to");
                    var endIdVal = ResolveShapeId(endId!, cxnShapeTree);
                    RequireReachableFrame(endIdVal, endId!, "to");
                    var endIdxVal = ParseCxnIdx(properties, "toIdx", "toidx", "endIdx", "endidx");
                    cxnDrawProps.EndConnection = new Drawing.EndConnection { Id = endIdVal, Index = endIdxVal };
                }

                // R53 bt-2: <a:cxnSpLocks noChangeShapeType="1"> — pinned
                // connector primitive (PowerPoint stamps it on inserted
                // connectors). Honor the `lockShapeType` input so dump→replay
                // round-trips the lock instead of silently dropping it.
                if (properties.TryGetValue("lockShapeType", out var cxnLockSt)
                    && IsTruthy(cxnLockSt))
                {
                    cxnDrawProps.AppendChild(new Drawing.ConnectionShapeLocks
                    {
                        NoChangeShapeType = true
                    });
                }

                connector.NonVisualConnectionShapeProperties = cxnNvProps;
                var cxnTransform = new Drawing.Transform2D(
                    new Drawing.Offset { X = cxnX, Y = cxnY },
                    new Drawing.Extents { Cx = cxnCx, Cy = cxnCy }
                );
                if (cxnFlipH) cxnTransform.HorizontalFlip = true;
                if (cxnFlipV) cxnTransform.VerticalFlip = true;
                // R53 bt-3: "line" is a distinct prstGeom value used by tools
                // that emit a connector as the bare-geometry primitive (no
                // <a:ln>). Aliasing it to StraightConnector1 silently rewrote
                // the prst attribute AND let the synthetic-outline branch
                // below stamp a default black 1pt stroke that the source
                // never had. Keep "line" as its own ShapeTypeValues.Line value
                // and signal `bareLine` so the outline injection skips when
                // no line.* props were supplied.
                var rawShapeKey = (properties.GetValueOrDefault("shape")
                                  ?? properties.GetValueOrDefault("preset", "straightConnector1"))
                                  .ToLowerInvariant();
                bool bareLine = rawShapeKey == "line";
                connector.ShapeProperties = new ShapeProperties(
                    cxnTransform,
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    {
                        // CONSISTENCY(canonical-key): canonical 'shape'; 'preset' legacy alias.
                        Preset = rawShapeKey switch
                        {
                            // Short canonical names + OOXML full names. "line" is the
                            // bare primitive (preserves prst="line" verbatim) — distinct
                            // from "straight"/"straightConnector1" which carries the
                            // canonical connector adjust list. The bent/curved families
                            // each have FOUR segment-count variants (2..5); map every
                            // OOXML name to its EXACT ShapeTypeValues so dump→replay keeps
                            // the source variant (collapsing e.g. curvedConnector4→3 would
                            // change the rendered bend). The friendly "elbow"/"curve"
                            // aliases default to the most common 3-segment form.
                            "straight" or "straightconnector1" => Drawing.ShapeTypeValues.StraightConnector1,
                            "line" => Drawing.ShapeTypeValues.Line,
                            "elbow" or "bentconnector3" => Drawing.ShapeTypeValues.BentConnector3,
                            "bentconnector2" => Drawing.ShapeTypeValues.BentConnector2,
                            "bentconnector4" => Drawing.ShapeTypeValues.BentConnector4,
                            "bentconnector5" => Drawing.ShapeTypeValues.BentConnector5,
                            "curve" or "curvedconnector3" => Drawing.ShapeTypeValues.CurvedConnector3,
                            "curvedconnector2" => Drawing.ShapeTypeValues.CurvedConnector2,
                            "curvedconnector4" => Drawing.ShapeTypeValues.CurvedConnector4,
                            "curvedconnector5" => Drawing.ShapeTypeValues.CurvedConnector5,
                            _ => throw new ArgumentException($"Invalid connector shape: '{properties.GetValueOrDefault("shape") ?? properties.GetValueOrDefault("preset", "straightConnector1")}'. Valid values: straight, elbow, curve, line (or OOXML full names: straightConnector1, bentConnector2-5, curvedConnector2-5).")
                        }
                    }
                );

                // R53 bt-3: when the source connector had no <a:ln> (the bare
                // prst="line" form) and the Add call carries no line.* / color
                // / arrow props, skip the synthetic outline so dump→replay
                // doesn't inject a default black 1pt stroke the source never
                // had. Any explicit line input still materializes <a:ln>.
                bool hasAnyLineInput =
                    properties.ContainsKey("line.gradient") || properties.ContainsKey("linegradient")
                    || properties.ContainsKey("lineColor") || properties.ContainsKey("linecolor")
                    || properties.ContainsKey("line") || properties.ContainsKey("color")
                    || properties.ContainsKey("line.color")
                    || properties.ContainsKey("linewidth") || properties.ContainsKey("lineWidth")
                    || properties.ContainsKey("line.width")
                    || properties.ContainsKey("lineDash") || properties.ContainsKey("linedash")
                    // R64 bt-3: lineDashRaw (<a:custDash> verbatim passthrough)
                    // — without this, replay of a connector that carries only
                    // a custom dash pattern (no width/color/preset-dash) would
                    // fall into the bareLine skip branch and drop <a:custDash>.
                    || properties.ContainsKey("lineDashRaw") || properties.ContainsKey("linedashraw")
                    || properties.ContainsKey("line.dashRaw") || properties.ContainsKey("line.dashraw")
                    || properties.ContainsKey("headEnd") || properties.ContainsKey("headend")
                    || properties.ContainsKey("tailEnd") || properties.ContainsKey("tailend")
                    // R58 bt-3: lineCap (<a:ln cap=...>) and cmpd (<a:ln cmpd=...>)
                    // also count as explicit line input — without these, replay
                    // of a source connector that only carried cap+cmpd (no width
                    // / color / dash) would fall into the bareLine skip branch
                    // and drop the attributes entirely.
                    || properties.ContainsKey("lineCap") || properties.ContainsKey("linecap")
                    || properties.ContainsKey("line.cap")
                    || properties.ContainsKey("cmpd") || properties.ContainsKey("compoundLine")
                    || properties.ContainsKey("compoundline") || properties.ContainsKey("line.compound")
                    // R61 bt-2: lineJoin (<a:round/>|<a:bevel/>|<a:miter/>) and miterLimit
                    // (<a:miter lim=...>) — without these, replay of a source connector
                    // that only carried join/limit (no width/color/dash) would fall into
                    // the bareLine skip branch and drop them entirely.
                    || properties.ContainsKey("lineJoin") || properties.ContainsKey("linejoin")
                    || properties.ContainsKey("line.join")
                    || properties.ContainsKey("miterLimit") || properties.ContainsKey("miterlimit")
                    || properties.ContainsKey("miter.limit") || properties.ContainsKey("line.miterlimit");
                bool skipOutline = bareLine && !hasAnyLineInput;

                // A connector whose colour/width comes from its <p:style> lnRef
                // (round-tripped as a raw-set append after this Add) must NOT get
                // a synthetic default black solidFill + 1pt width — the explicit
                // outline would override the theme reference and replay black.
                // The emitter sets `styledLine` only when a style is present and
                // no explicit line colour/width was supplied, so any other line.*
                // input (tailEnd, dash, cap, …) still lands on a bare <a:ln> that
                // inherits colour/width from lnRef.
                bool styledLine = IsTruthy(properties.GetValueOrDefault("styledLine"))
                                  || IsTruthy(properties.GetValueOrDefault("styledline"));

                // Line style
                var cxnOutline = new Drawing.Outline { Width = 12700 }; // 1pt default
                // line.gradient parity with Set side — accept gradient outline at Add time
                // so dump→replay round-trips. Gradient fill wins over solid color.
                if (properties.TryGetValue("line.gradient", out var cxnLineGrad)
                    || properties.TryGetValue("linegradient", out cxnLineGrad))
                {
                    cxnOutline.AppendChild(BuildGradientFill(NormalizeLineGradientSpec(cxnLineGrad)));
                }
                else if (properties.TryGetValue("lineColor", out var cxnColor2) || properties.TryGetValue("linecolor", out cxnColor2)
                    || properties.TryGetValue("line", out cxnColor2) || properties.TryGetValue("color", out cxnColor2)
                    || properties.TryGetValue("line.color", out cxnColor2))
                    cxnOutline.AppendChild(BuildSolidFill(cxnColor2));
                else if (!styledLine)
                    cxnOutline.AppendChild(BuildSolidFill("000000"));
                if (properties.TryGetValue("linewidth", out var lwVal) || properties.TryGetValue("lineWidth", out lwVal)
                    || properties.TryGetValue("line.width", out lwVal))
                    cxnOutline.Width = Core.EmuConverter.ParseLineWidth(lwVal);
                if (properties.TryGetValue("lineDash", out var cxnDash) || properties.TryGetValue("linedash", out cxnDash))
                {
                    cxnOutline.AppendChild(new Drawing.PresetDash { Val = ParseLineDashValue(cxnDash) });
                }
                // R64 bt-3: lineDashRaw — verbatim <a:custDash> install. Mirrors
                // shadowRaw / fillOverlayRaw passthrough strategy: lift attrs
                // (none on custDash itself) + InnerXml from the source XML and
                // append a fresh Drawing.CustomDash. Wins over lineDash since
                // CT_LineProperties accepts only one of prstDash / custDash
                // (EG_LineDashProperties choice).
                if (properties.TryGetValue("lineDashRaw", out var cxnDashRaw)
                    || properties.TryGetValue("linedashraw", out cxnDashRaw)
                    || properties.TryGetValue("line.dashRaw", out cxnDashRaw)
                    || properties.TryGetValue("line.dashraw", out cxnDashRaw))
                {
                    if (!string.IsNullOrWhiteSpace(cxnDashRaw))
                    {
                        cxnOutline.RemoveAllChildren<Drawing.PresetDash>();
                        cxnOutline.RemoveAllChildren<Drawing.CustomDash>();
                        cxnOutline.AppendChild(BuildCustomDashFromRaw(cxnDashRaw));
                    }
                }
                // R58 bt-3: lineCap (<a:ln cap="..."/>) and cmpd (<a:ln cmpd="..."/>)
                // — attributes on the outline element itself, not children.
                // Previously dropped silently on dump→replay; mirror the shape
                // Add aliases (lineCap/line.cap, cmpd/compoundLine/line.compound)
                // so both Add and Set paths accept the same vocabulary.
                if (properties.TryGetValue("lineCap", out var cxnCap)
                    || properties.TryGetValue("linecap", out cxnCap)
                    || properties.TryGetValue("line.cap", out cxnCap))
                {
                    cxnOutline.CapType = cxnCap.ToLowerInvariant() switch
                    {
                        "round" or "rnd" => Drawing.LineCapValues.Round,
                        "flat" => Drawing.LineCapValues.Flat,
                        "square" or "sq" => Drawing.LineCapValues.Square,
                        _ => throw new ArgumentException($"Invalid 'lineCap' value: '{cxnCap}'. Valid values: round, flat, square.")
                    };
                }
                if (properties.TryGetValue("cmpd", out var cxnCmpd)
                    || properties.TryGetValue("compoundLine", out cxnCmpd)
                    || properties.TryGetValue("compoundline", out cxnCmpd)
                    || properties.TryGetValue("line.compound", out cxnCmpd))
                {
                    cxnOutline.CompoundLineType = cxnCmpd switch
                    {
                        var s when s.Equals("sng", StringComparison.OrdinalIgnoreCase) || s.Equals("single", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Single,
                        var s when s.Equals("dbl", StringComparison.OrdinalIgnoreCase) || s.Equals("double", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Double,
                        var s when s.Equals("thickThin", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.ThickThin,
                        var s when s.Equals("thinThick", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.ThinThick,
                        var s when s.Equals("tri", StringComparison.OrdinalIgnoreCase) || s.Equals("triple", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Triple,
                        _ => throw new ArgumentException($"Invalid 'cmpd' value: '{cxnCmpd}'. Valid values: sng, dbl, thickThin, thinThick, tri.")
                    };
                }
                // R61 bt-2: lineJoin (<a:round/>|<a:bevel/>|<a:miter/>) and miterLimit
                // (<a:miter lim="N"/>) — previously silently dropped on connector
                // dump→replay even though shape Set/Add already accepted them.
                // Accept compound "miter:<lim>" form so a single key carries both.
                int? cxnMiterLim = null;
                if (properties.TryGetValue("miterLimit", out var cxnMiterLimRaw)
                    || properties.TryGetValue("miterlimit", out cxnMiterLimRaw)
                    || properties.TryGetValue("miter.limit", out cxnMiterLimRaw)
                    || properties.TryGetValue("line.miterlimit", out cxnMiterLimRaw))
                {
                    if (!int.TryParse(cxnMiterLimRaw, System.Globalization.NumberStyles.Integer,
                            System.Globalization.CultureInfo.InvariantCulture, out var cxnMiterLimParsed))
                        throw new ArgumentException($"Invalid 'miterLimit' value: '{cxnMiterLimRaw}'. Expected integer (1000ths of a percent, e.g. 800000 = 800%).");
                    cxnMiterLim = cxnMiterLimParsed;
                }
                if (properties.TryGetValue("lineJoin", out var cxnJoin)
                    || properties.TryGetValue("linejoin", out cxnJoin)
                    || properties.TryGetValue("line.join", out cxnJoin))
                {
                    var cxnJoinValue = cxnJoin;
                    var cxnJoinColon = cxnJoin.IndexOf(':');
                    if (cxnJoinColon > 0)
                    {
                        cxnJoinValue = cxnJoin.Substring(0, cxnJoinColon);
                        var cxnLimTok = cxnJoin.Substring(cxnJoinColon + 1).Trim();
                        if (!int.TryParse(cxnLimTok, System.Globalization.NumberStyles.Integer,
                                System.Globalization.CultureInfo.InvariantCulture, out var cxnLimParsed))
                            throw new ArgumentException($"Invalid 'lineJoin' miter limit token: '{cxnLimTok}'. Expected integer (1000ths of a percent, e.g. 800000 = 800%).");
                        cxnMiterLim = cxnLimParsed;
                    }
                    OpenXmlElement cxnJoinEl = cxnJoinValue.ToLowerInvariant() switch
                    {
                        "round" => new Drawing.Round(),
                        "bevel" => new Drawing.LineJoinBevel(),
                        "miter" => cxnMiterLim.HasValue
                            ? new Drawing.Miter { Limit = cxnMiterLim.Value }
                            : new Drawing.Miter(),
                        _ => throw new ArgumentException($"Invalid 'lineJoin' value: '{cxnJoinValue}'. Valid values: round, bevel, miter.")
                    };
                    cxnOutline.AppendChild(cxnJoinEl);
                }
                else if (cxnMiterLim.HasValue)
                {
                    // miterLimit alone implies miter join.
                    cxnOutline.AppendChild(new Drawing.Miter { Limit = cxnMiterLim.Value });
                }
                // Arrow head/tail
                if (properties.TryGetValue("headEnd", out var headVal) || properties.TryGetValue("headend", out headVal))
                {
                    cxnOutline.AppendChild(new Drawing.HeadEnd { Type = ParseLineEndType(headVal) });
                }
                if (properties.TryGetValue("tailEnd", out var tailVal) || properties.TryGetValue("tailend", out tailVal))
                {
                    cxnOutline.AppendChild(new Drawing.TailEnd { Type = ParseLineEndType(tailVal) });
                }

                // CONSISTENCY(shape-picture-parity): rotation lives on Transform2D
                // for shape/picture/connector/group; all four must parse the same
                // way. Shape (Add.Shape.cs) and Picture (Add.Media.cs) accept
                // fractional degrees (e.g. 22.5); connector previously used
                // int.TryParse and silently dropped non-integer values.
                if (properties.TryGetValue("rotation", out var cxnRot)
                    || properties.TryGetValue("rotate", out cxnRot))
                {
                    connector.ShapeProperties.Transform2D!.Rotation =
                        (int)(ParseHelpers.SafeParseRotationDegrees(cxnRot, "rotation") * 60000);
                }
                if (!skipOutline)
                    connector.ShapeProperties.AppendChild(cxnOutline);

                // R57 bt-4: in-line text label on the connector (<p:txBody>
                // child of <p:cxnSp>). PowerPoint and most flowchart authoring
                // tools attach a txBody to connectors that show a label
                // between their endpoints. The OOXML p:cxnSp schema does not
                // declare txBody, so we attach the typed
                // Presentation.TextBody as a permissive child — round-trips
                // through the SDK as an OpenXmlUnknownElement on reload
                // (NodeBuilder.ResolveConnectorTextBody reparses it back to
                // the typed form). Accept `text` for the single-paragraph
                // single-run inline case; multi-paragraph / multi-run
                // labels arrive via subsequent `add paragraph` / `add run`
                // ops against the connector path.
                if (properties.TryGetValue("text", out var cxnText) && !string.IsNullOrEmpty(cxnText))
                {
                    XmlTextValidator.ValidateOrThrow(cxnText, "text", allowSoftBreakChar: true);
                    var cxnRunProps = new Drawing.RunProperties { Language = "en-US" };
                    var cxnPara = new Drawing.Paragraph(new Drawing.Run(cxnRunProps,
                        MakePreservingText(cxnText)));
                    var cxnTxBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        cxnPara);
                    connector.AppendChild(cxnTxBody);
                }

                InsertAtPosition(cxnInsertContainer, connector, index);
                if (properties.TryGetValue("zorder", out var cxnZ)
                    || properties.TryGetValue("z-order", out cxnZ)
                    || properties.TryGetValue("order", out cxnZ))
                    ApplyZOrder(cxnSlidePart, connector, cxnZ);
                GetSlide(cxnSlidePart).Save();

                return $"{cxnReturnPrefix}/{BuildElementPathSegment("connector", connector, cxnInsertContainer.Elements<ConnectionShape>().Count())}";
    }

    // R57 bt-4: Resolve a connector under a slide by either positional index
    // ("3") or cNvPr id form ("@id=N" matching BuildElementPathSegment). Used
    // by AddParagraph / AddRun connector-parent branches.
    internal static ConnectionShape ResolveConnectorByToken(SlidePart slidePart, string token)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");
        var connectors = shapeTree.Elements<ConnectionShape>().ToList();
        var idMatch = Regex.Match(token, @"^@id=(\d+)$");
        if (idMatch.Success && uint.TryParse(idMatch.Groups[1].Value, out var id))
        {
            var match = connectors.FirstOrDefault(c =>
                c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value == id);
            if (match == null)
                throw new ArgumentException($"Connector with id={id} not found");
            return match;
        }
        if (int.TryParse(token, out var posIdx))
        {
            if (posIdx < 1 || posIdx > connectors.Count)
                throw new ArgumentException($"Connector {posIdx} not found (total: {connectors.Count})");
            return connectors[posIdx - 1];
        }
        throw new ArgumentException($"Invalid connector token: '{token}'");
    }

    internal static int ConnectorPositionalIndex(SlidePart slidePart, ConnectionShape cxn)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");
        return PathIndex.FromArrayIndex(shapeTree.Elements<ConnectionShape>().ToList().IndexOf(cxn));
    }

    // R57 bt-4: locate the connector's <p:txBody> (lazily creating one when
    // absent) and return it as a strongly-typed Presentation.TextBody whose
    // edits are committed back to the connector XML. The OpenXml SDK parses
    // the unknown txBody subtree on cxnSp as a raw OpenXmlUnknownElement; we
    // detect that form, replace it with a reparsed typed TextBody, and append
    // a fresh one when the connector carries no label yet. Subsequent
    // AddParagraph / AddRun edits land on the live typed element so changes
    // serialize correctly.
    internal static DocumentFormat.OpenXml.Presentation.TextBody ConnectorEnsureTextBody(ConnectionShape cxn)
    {
        var typed = cxn.GetFirstChild<DocumentFormat.OpenXml.Presentation.TextBody>();
        if (typed != null) return typed;

        var unk = cxn.ChildElements.OfType<OpenXmlUnknownElement>()
            .FirstOrDefault(e => e.LocalName == "txBody");
        if (unk != null)
        {
            var rebuilt = new DocumentFormat.OpenXml.Presentation.TextBody(unk.OuterXml);
            cxn.ReplaceChild(rebuilt, unk);
            return rebuilt;
        }

        var fresh = new DocumentFormat.OpenXml.Presentation.TextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle());
        cxn.AppendChild(fresh);
        return fresh;
    }

    /// <summary>
    /// Returns the slide-absolute bounding box (EMU) of a shape-tree frame by its
    /// NonVisualDrawingProperties Id, across all frame kinds (Shape, Picture,
    /// ConnectionShape, GraphicFrame, GroupShape). Shared by connector Add and Set
    /// (R14-4) so reconnecting endpoints recomputes the connector xfrm the same way
    /// Add does. Returns null if no frame with that id has a transform.
    /// </summary>
    private static (long x, long y, long cx, long cy)? FrameBoundsById(ShapeTree tree, uint id)
    {
        foreach (var el in tree.ChildElements)
        {
            Drawing.Transform2D? xf = null;
            uint? frameId = null;
            switch (el)
            {
                case Shape s:
                    frameId = s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                    xf = s.ShapeProperties?.Transform2D;
                    break;
                case Picture p:
                    frameId = p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                    xf = p.ShapeProperties?.Transform2D;
                    break;
                case ConnectionShape c:
                    frameId = c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                    xf = c.ShapeProperties?.Transform2D;
                    break;
                case GraphicFrame gf:
                    frameId = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value;
                    if (frameId == id && gf.Transform != null)
                        return (gf.Transform.Offset?.X?.Value ?? 0, gf.Transform.Offset?.Y?.Value ?? 0,
                                gf.Transform.Extents?.Cx?.Value ?? 0, gf.Transform.Extents?.Cy?.Value ?? 0);
                    break;
                case GroupShape g:
                    frameId = g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                    var gxf = g.GroupShapeProperties?.TransformGroup;
                    if (frameId == id && gxf != null)
                        return (gxf.Offset?.X?.Value ?? 0, gxf.Offset?.Y?.Value ?? 0,
                                gxf.Extents?.Cx?.Value ?? 0, gxf.Extents?.Cy?.Value ?? 0);
                    break;
            }
            if (frameId == id && xf != null)
                return (xf.Offset?.X?.Value ?? 0, xf.Offset?.Y?.Value ?? 0,
                        xf.Extents?.Cx?.Value ?? 0, xf.Extents?.Cy?.Value ?? 0);
        }
        return null;
    }

    /// <summary>
    /// Given the two connected frames' bounding boxes (EMU) and optional per-end
    /// side overrides, return the connector's start/end endpoint coordinates.
    ///
    /// PowerPoint does not recompute the visible line from &lt;a:stCxn/endCxn&gt; —
    /// it trusts our stored offset/extent — so the endpoint we pick here IS what
    /// gets drawn. Historically both endpoints were the shape centers, which paints
    /// a straight line through both boxes for same-row / same-column layouts.
    ///
    /// Auto (side == null): pick the edge midpoint on the axis that dominates the
    /// center-to-center vector — horizontal separation → right↔left, vertical →
    /// bottom↔top. Same-row boxes yield A-right → B-left; stacked boxes yield
    /// A-bottom → B-top. Explicit side (left/right/top/bottom) forces that edge;
    /// "center" reproduces the legacy through-the-box behavior on demand.
    /// </summary>
    private static (long p1x, long p1y, long p2x, long p2y) ComputeConnectorEndpoints(
        (long x, long y, long cx, long cy) a,
        (long x, long y, long cx, long cy) b,
        string? fromSide, string? toSide)
    {
        static (long px, long py) EdgePoint((long x, long y, long cx, long cy) f, string side) => side switch
        {
            "left"   => (f.x, f.y + f.cy / 2),
            "right"  => (f.x + f.cx, f.y + f.cy / 2),
            "top"    => (f.x + f.cx / 2, f.y),
            "bottom" => (f.x + f.cx / 2, f.y + f.cy),
            _        => (f.x + f.cx / 2, f.y + f.cy / 2), // center
        };

        // Auto side for each end from the center-to-center direction. The frame
        // farther right/down "faces" the other along whichever axis separates them
        // most, so A and B get opposite edges on that axis.
        long acx = a.x + a.cx / 2, acy = a.y + a.cy / 2;
        long bcx = b.x + b.cx / 2, bcy = b.y + b.cy / 2;
        long dx = bcx - acx, dy = bcy - acy;
        string autoFrom, autoTo;
        if (Math.Abs(dx) >= Math.Abs(dy))
        {
            autoFrom = dx >= 0 ? "right" : "left";
            autoTo   = dx >= 0 ? "left" : "right";
        }
        else
        {
            autoFrom = dy >= 0 ? "bottom" : "top";
            autoTo   = dy >= 0 ? "top" : "bottom";
        }

        var (p1x, p1y) = EdgePoint(a, fromSide ?? autoFrom);
        var (p2x, p2y) = EdgePoint(b, toSide ?? autoTo);
        return (p1x, p1y, p2x, p2y);
    }

    /// <summary>
    /// Normalize a user-supplied connector side value to the canonical
    /// top/bottom/left/right/center token, or null if unset/invalid. Accepts the
    /// same lenient casing as other Add/Set inputs.
    /// </summary>
    private static string? NormalizeConnectorSide(Dictionary<string, string> p, params string[] keys)
    {
        foreach (var k in keys)
        {
            if (!p.TryGetValue(k, out var raw) || raw == null) continue;
            var v = raw.Trim().ToLowerInvariant();
            switch (v)
            {
                case "left": case "right": case "top": case "bottom": case "center":
                    return v;
                case "middle": case "centre":
                    return "center";
                default:
                    throw new ArgumentException(
                        $"Invalid connector side '{raw}' for '{k}': expected top, bottom, left, right, or center.");
            }
        }
        return null;
    }

    /// <summary>
    /// Resolves a shape reference to an OOXML shape ID.
    /// Accepts: plain integer (shape ID), or DOM path like /slide[1]/shape[2] (resolves Nth shape's ID).
    /// </summary>
    private static uint ResolveShapeId(string value, ShapeTree shapeTree)
    {
        // Connector endpoints may be ANY top-level frame — shape, picture, table,
        // chart, connector, group, ole, media, model3d, zoom — not just `shape`
        // (FrameBoundsById supports the full list). id/name are unique across
        // frames, so @id/@name need no type filter; positional forms are scoped
        // to the named type; a bare integer resolves against the full frame list.
        (uint? id, string? name) FrameIdName(OpenXmlElement el) => el switch
        {
            Shape s => (s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                        s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value),
            Picture p => (p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
                          p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value),
            ConnectionShape c => (c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                                  c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value),
            GraphicFrame gf => (gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
                                gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value),
            GroupShape g => (g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                             g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value),
            _ => (null, null),
        };
        var allFrames = shapeTree.ChildElements.Where(e => FrameIdName(e).id.HasValue).ToList();

        // Plain integer: match an existing frame ID first (across ALL frame types,
        // not just shapes), else treat as a 1-based index into the full frame list,
        // else pass through as a literal id (forward-ref for dump replay).
        if (uint.TryParse(value, out var directId))
        {
            if (allFrames.Any(e => FrameIdName(e).id == directId))
                return directId;
            if (directId >= 1 && directId <= (uint)allFrames.Count)
                return FrameIdName(allFrames[(int)directId - 1]).id!.Value;
            return directId;
        }

        // @id path: /slide[N]/<type>[@id=M]
        var atIdMatch = Regex.Match(value, @"/slide\[\d+\]/\w+\[@id=(\d+)\]");
        if (atIdMatch.Success)
        {
            var atId = uint.Parse(atIdMatch.Groups[1].Value);
            if (!allFrames.Any(e => FrameIdName(e).id == atId))
                throw new ArgumentException($"Frame @id={atId} not found on this slide");
            return atId;
        }

        // @name path: /slide[N]/<type>[@name=Foo]
        var atNameMatch = Regex.Match(value, @"/slide\[\d+\]/\w+\[@name=([^\]]+)\]");
        if (atNameMatch.Success)
        {
            var atName = atNameMatch.Groups[1].Value;
            var matched = allFrames.FirstOrDefault(e => FrameIdName(e).name == atName);
            if (matched == null)
                throw new ArgumentException($"Frame @name={atName} not found on this slide");
            return FrameIdName(matched).id
                ?? throw new ArgumentException($"Frame @name={atName} has no ID");
        }

        // Positional path: /slide[N]/<type>[M] (1-based among that element type).
        var pathMatch = Regex.Match(value, @"/slide\[\d+\]/(\w+)\[(\d+)\]");
        if (pathMatch.Success)
        {
            var typeToken = pathMatch.Groups[1].Value.ToLowerInvariant();
            var idx = int.Parse(pathMatch.Groups[2].Value);
            // Unknown type tokens (typos like `shpae`, or `widget`) must be REJECTED,
            // not silently bucketed into GraphicFrame — otherwise a misspelled ref
            // resolves to the wrong endpoint with no error.
            // table/chart/diagram/ole are ALL <p:graphicFrame>; the positional index
            // must filter by graphicData URI so `chart[1]` counts charts (not the
            // first graphicFrame, which may be a table). Media/model3d are <p:pic>.
            static string GfUri(GraphicFrame gf) => gf.Graphic?.GraphicData?.Uri?.Value ?? "";
            IEnumerable<OpenXmlElement>? frames = typeToken switch
            {
                "shape" => shapeTree.Elements<Shape>(),
                "picture" or "media" or "model3d" => shapeTree.Elements<Picture>(),
                "connector" or "connection" => shapeTree.Elements<ConnectionShape>(),
                "group" => shapeTree.Elements<GroupShape>(),
                "table" => shapeTree.Elements<GraphicFrame>()
                    .Where(gf => GfUri(gf).Contains("/table", StringComparison.OrdinalIgnoreCase)),
                "chart" => shapeTree.Elements<GraphicFrame>()
                    .Where(gf => GfUri(gf).Contains("chart", StringComparison.OrdinalIgnoreCase)),
                "diagram" or "smartart" => shapeTree.Elements<GraphicFrame>()
                    .Where(gf => GfUri(gf).Contains("diagram", StringComparison.OrdinalIgnoreCase)),
                "ole" => shapeTree.Elements<GraphicFrame>()
                    .Where(gf => GfUri(gf).Contains("ole", StringComparison.OrdinalIgnoreCase)),
                // bare graphicFrame / zoom: any graphicFrame (best-effort, no type filter).
                "graphicframe" or "zoom" => shapeTree.Elements<GraphicFrame>(),
                _ => null,
            };
            if (frames == null)
                throw new ArgumentException(
                    $"Unknown frame type '{typeToken}' in reference '{value}'. Expected shape, picture, table, chart, group, connector, ole, media, model3d, or zoom.");
            var list = frames.ToList();
            if (idx < 1 || idx > list.Count)
                throw new ArgumentException($"{typeToken} index {idx} out of range (total: {list.Count})");
            return FrameIdName(list[PathIndex.ToArrayIndex(idx)]).id
                ?? throw new ArgumentException($"{typeToken} {idx} has no ID");
        }

        throw new ArgumentException($"Invalid shape reference: '{value}'. Expected a frame index (1, 2, ...), path (/slide[N]/<type>[M]), @id path (/slide[N]/<type>[@id=M]), or @name path (/slide[N]/<type>[@name=Foo]).");
    }

    private string AddGroup(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var grpSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                // CONSISTENCY(nested-group): accept a /slide[N]/group[K]... parent
                // chain so dump-replay of nested groups round-trips. AddEmptyGroup
                // inserts into the resolved container element; sibling lookups use
                // GroupShape children there instead of the slide-level shape tree.
                var grpNestedMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\](/group\[\d+\])+$");
                if (!grpSlideMatch.Success && !grpNestedMatch.Success)
                    throw new ArgumentException("Groups must be added to a slide or a nested group: /slide[N] or /slide[N]/group[K]");

                var grpSlideIdx = int.Parse((grpSlideMatch.Success ? grpSlideMatch : grpNestedMatch).Groups[1].Value);
                var grpSlideParts = GetSlideParts().ToList();
                if (grpSlideIdx < 1 || grpSlideIdx > grpSlideParts.Count)
                    throw new ArgumentException($"Slide {grpSlideIdx} not found (total: {grpSlideParts.Count})");

                var grpSlidePart = grpSlideParts[grpSlideIdx - 1];
                var grpSlideShapeTree = GetSlide(grpSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                // Resolve container: slide-level ShapeTree or a nested GroupShape.
                // For Add purposes either works (both expose ChildElements + can
                // host a new GroupShape via InsertAtPosition).
                OpenXmlCompositeElement grpShapeTree = grpSlideShapeTree;
                if (grpNestedMatch.Success)
                {
                    var nestedTokens = Regex.Matches(parentPath.Substring($"/slide[{grpSlideIdx}]".Length), @"/group\[(\d+)\]");
                    OpenXmlCompositeElement cursor = grpSlideShapeTree;
                    foreach (Match nm in nestedTokens)
                    {
                        var gi = int.Parse(nm.Groups[1].Value);
                        var nestedGroups = cursor.Elements<GroupShape>().ToList();
                        if (gi < 1 || gi > nestedGroups.Count)
                            throw new ArgumentException($"Group {gi} not found under {parentPath} (total: {nestedGroups.Count})");
                        cursor = nestedGroups[gi - 1];
                    }
                    grpShapeTree = cursor;
                }

                // ID allocation must scan the whole slide (shape IDs are slide-scoped),
                // not the container; use the slide-level shape tree even for nested groups.
                var grpId = AcquireShapeId(grpSlideShapeTree, properties);
                var grpName = properties.GetValueOrDefault("name", $"Group {grpShapeTree.Elements<GroupShape>().Count() + 1}");

                // Parse shape paths to group: shapes="1,2,3" (shape indices)
                if (!properties.TryGetValue("shapes", out var shapesStr))
                {
                    // CONSISTENCY(dump-replay-empty-group): dump emits
                    // `add group` (geometry only) followed by per-child
                    // `add shape parent=/slide/group[K]`. Without an empty-
                    // group mode here, dump-replay would lose every group.
                    // Required props: at least one of the geometry markers
                    // so this stays distinguishable from a mis-typed 'shapes'
                    // call ('groups must group something' was the old
                    // intent — that's still the message when geometry is
                    // also absent).
                    bool hasGeometry =
                        properties.ContainsKey("x") || properties.ContainsKey("y")
                        || properties.ContainsKey("width") || properties.ContainsKey("height")
                        || properties.ContainsKey("cx") || properties.ContainsKey("cy");
                    if (!hasGeometry)
                        throw new ArgumentException("'shapes' property required: comma-separated shape indices to group (e.g. shapes=1,2,3), or supply geometry (x,y,width,height) for an empty group to be filled by subsequent `add shape parent=/slide[N]/group[K]` calls.");

                    return AddEmptyGroup(grpSlidePart, grpShapeTree, grpSlideIdx, grpId, grpName, index, properties, parentPath);
                }

                // CONSISTENCY(query-path-roundtrip): help advertises @id=/@name=
                // path forms for shapes=; query shape returns @id form. Resolve
                // against the same heterogeneous frame list AddGroup uses below
                // so groups can include pictures / graphicFrames / connectors.
                var grpFrameList = grpShapeTree.ChildElements
                    .Where(c => c is Shape || c is GroupShape || c is Picture
                        || c is GraphicFrame || c is ConnectionShape)
                    .ToList();
                static uint? FrameId(OpenXmlElement e) => e switch
                {
                    Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                    GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                    Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
                    GraphicFrame gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
                    ConnectionShape c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                    _ => null,
                };
                static string? FrameName(OpenXmlElement e) => e switch
                {
                    Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                    GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                    Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value,
                    GraphicFrame gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value,
                    ConnectionShape c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                    _ => null,
                };

                // CONSISTENCY(group-frame-types): include all frame-like elements
                // (Shape, GroupShape, Picture, GraphicFrame, ConnectionShape) so
                // existing groups, pictures, charts, and connectors can also be
                // grouped together. Index space matches the shape-tree order
                // PowerPoint uses for sibling lookups (B13).
                //
                // CONSISTENCY(group-numeric-skip-placeholder): exclude placeholder
                // <p:sp> elements (those with a <p:ph> child) from the numeric
                // index so `shapes=1,2` aligns with the non-placeholder shape[N]
                // index space that Query/Get use. Users wanting to group a
                // placeholder can still target it explicitly via the @id= /
                // @name= / path forms below.
                var allShapes = grpShapeTree.ChildElements
                    .Where(c => c is Shape || c is GroupShape || c is Picture
                        || c is GraphicFrame || c is ConnectionShape)
                    .Where(c => !(c is Shape s
                        && s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                            ?.GetFirstChild<PlaceholderShape>() != null))
                    .ToList();

                var shapeParts = shapesStr.Split(',');
                // R29-2: resolve every part directly to its OpenXmlElement. @id/@name
                // path forms look up in grpFrameList (placeholders included, so an
                // explicitly-named placeholder can still be grouped); bare-numeric
                // and positional [M] forms index allShapes (placeholders excluded, the
                // shape[N] space Query/Get use). Resolving to the element — instead of
                // a position in one list re-looked-up in the other — keeps the two
                // index spaces from colliding when a placeholder precedes the target.
                var resolved = new List<OpenXmlElement>();
                foreach (var sp in shapeParts)
                {
                    var trimmed = sp.Trim();
                    if (trimmed.StartsWith("/"))
                    {
                        // CONSISTENCY(group-frame-paths): accept any frame-like
                        // element kind in the path (shape / group / picture / pic /
                        // connector / connection / chart / table / graphicframe),
                        // mirroring the heterogeneous frame list AddGroup operates
                        // on. Without this, `query group` / `query picture` paths
                        // round-tripped into `shapes=` were rejected even though
                        // the lookup index space supports them. The first element
                        // type is intentionally not validated against frame kind
                        // — id/name/positional lookup is by-position regardless of
                        // which kind name the user used.
                        const string frameKind =
                            @"(?:shape|group|picture|pic|connector|connection|chart|table|graphicframe|graphicFrame|ole|object|embed|video|audio)";

                        // @id path: /slide[N]/<kind>[@id=M] — round-trips from `query`
                        var atIdMatch = Regex.Match(trimmed,
                            $@"/slide\[\d+\]/{frameKind}\[@id=(\d+)\]",
                            RegexOptions.IgnoreCase);
                        if (atIdMatch.Success)
                        {
                            var atId = uint.Parse(atIdMatch.Groups[1].Value);
                            var el = grpFrameList.FirstOrDefault(e => FrameId(e) == atId);
                            if (el == null)
                                throw new ArgumentException($"Frame @id={atId} not found on this slide");
                            resolved.Add(el);
                            continue;
                        }
                        // @name path: /slide[N]/<kind>[@name=Foo]
                        var atNameMatch = Regex.Match(trimmed,
                            $@"/slide\[\d+\]/{frameKind}\[@name=([^\]]+)\]",
                            RegexOptions.IgnoreCase);
                        if (atNameMatch.Success)
                        {
                            var atName = atNameMatch.Groups[1].Value;
                            var el = grpFrameList.FirstOrDefault(e => FrameName(e) == atName);
                            if (el == null)
                                throw new ArgumentException($"Frame @name={atName} not found on this slide");
                            resolved.Add(el);
                            continue;
                        }
                        // Positional path: /slide[N]/<kind>[M] — indexes the
                        // placeholder-excluded allShapes space (shape[N] form).
                        var pathMatch = Regex.Match(trimmed,
                            $@"/slide\[\d+\]/{frameKind}\[(\d+)\]",
                            RegexOptions.IgnoreCase);
                        if (!pathMatch.Success)
                            throw new ArgumentException($"Invalid frame path: '{trimmed}'. Expected /slide[N]/<kind>[M], /slide[N]/<kind>[@id=ID], or /slide[N]/<kind>[@name=Foo] where <kind> is shape/group/picture/connector/chart/table/etc.");
                        var pIdx = int.Parse(pathMatch.Groups[1].Value);
                        if (pIdx < 1 || pIdx > allShapes.Count)
                            throw new ArgumentException($"Shape {pIdx} not found (total: {allShapes.Count})");
                        resolved.Add(allShapes[pIdx - 1]);
                    }
                    else if (int.TryParse(trimmed, out var idx))
                    {
                        if (idx < 1 || idx > allShapes.Count)
                            throw new ArgumentException($"Shape {idx} not found (total: {allShapes.Count})");
                        resolved.Add(allShapes[idx - 1]);
                    }
                    else
                    {
                        throw new ArgumentException($"Invalid 'shapes' value: '{trimmed}' is not a valid integer or DOM path. Expected comma-separated shape indices (e.g. shapes=1,2,3) or DOM paths (e.g. shapes=/slide[1]/shape[1],/slide[1]/shape[2]).");
                    }
                }

                // Collect shapes to group in shape-tree order (stable regardless of
                // the order parts were listed), de-duplicating repeated references.
                var toGroup = grpFrameList.Where(e => resolved.Contains(e)).ToList();

                // Calculate bounding box across heterogeneous frame elements.
                long minX = long.MaxValue, minY = long.MaxValue, maxX = long.MinValue, maxY = long.MinValue;
                bool hasTransform = false;
                foreach (var s in toGroup)
                {
                    long? sx = null, sy = null, scx = null, scy = null;
                    switch (s)
                    {
                        case Shape sp:
                            var xfrmSp = sp.ShapeProperties?.Transform2D;
                            sx = xfrmSp?.Offset?.X?.Value; sy = xfrmSp?.Offset?.Y?.Value;
                            scx = xfrmSp?.Extents?.Cx?.Value; scy = xfrmSp?.Extents?.Cy?.Value;
                            break;
                        case Picture pic:
                            var xfrmPic = pic.ShapeProperties?.Transform2D;
                            sx = xfrmPic?.Offset?.X?.Value; sy = xfrmPic?.Offset?.Y?.Value;
                            scx = xfrmPic?.Extents?.Cx?.Value; scy = xfrmPic?.Extents?.Cy?.Value;
                            break;
                        case ConnectionShape cs:
                            var xfrmCs = cs.ShapeProperties?.Transform2D;
                            sx = xfrmCs?.Offset?.X?.Value; sy = xfrmCs?.Offset?.Y?.Value;
                            scx = xfrmCs?.Extents?.Cx?.Value; scy = xfrmCs?.Extents?.Cy?.Value;
                            break;
                        case GroupShape gs:
                            var xfrmGs = gs.GroupShapeProperties?.TransformGroup;
                            sx = xfrmGs?.Offset?.X?.Value; sy = xfrmGs?.Offset?.Y?.Value;
                            scx = xfrmGs?.Extents?.Cx?.Value; scy = xfrmGs?.Extents?.Cy?.Value;
                            break;
                        case GraphicFrame gf:
                            var xfrmGf = gf.Transform;
                            sx = xfrmGf?.Offset?.X?.Value; sy = xfrmGf?.Offset?.Y?.Value;
                            scx = xfrmGf?.Extents?.Cx?.Value; scy = xfrmGf?.Extents?.Cy?.Value;
                            break;
                    }
                    if (sx == null || sy == null || scx == null || scy == null) continue;
                    hasTransform = true;
                    if (sx.Value < minX) minX = sx.Value;
                    if (sy.Value < minY) minY = sy.Value;
                    if (sx.Value + scx.Value > maxX) maxX = sx.Value + scx.Value;
                    if (sy.Value + scy.Value > maxY) maxY = sy.Value + scy.Value;
                }
                if (!hasTransform) { minX = 0; minY = 0; maxX = 0; maxY = 0; }

                var groupShape = new GroupShape();
                groupShape.NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = grpId, Name = grpName },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                groupShape.GroupShapeProperties = new GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = minX, Y = minY },
                        new Drawing.Extents { Cx = maxX - minX, Cy = maxY - minY },
                        new Drawing.ChildOffset { X = minX, Y = minY },
                        new Drawing.ChildExtents { Cx = maxX - minX, Cy = maxY - minY }
                    )
                );

                // Mirror SetGroupByPath: dump emits `rotation=<deg>` for groups whose
                // <a:xfrm rot="..."> was non-zero, so AddGroup must honor the same
                // input key. Without this the apply path silently dropped rot on
                // every dump→batch replay.
                if (properties.TryGetValue("rotation", out var grpRot)
                    || properties.TryGetValue("rotate", out grpRot))
                {
                    groupShape.GroupShapeProperties.TransformGroup!.Rotation =
                        (int)(ParseHelpers.SafeParseRotationDegrees(grpRot, "rotation") * 60000);
                }

                // Move shapes into group
                foreach (var s in toGroup)
                {
                    s.Remove();
                    groupShape.AppendChild(s);
                }

                InsertAtPosition(grpShapeTree, groupShape, index);

                // Optional click hyperlink on the group's cNvPr — same
                // contract as shape/picture so Add and Set agree on the
                // 'link' / 'tooltip' input keys at creation time.
                if (properties.TryGetValue("link", out var grpLinkVal) && !string.IsNullOrEmpty(grpLinkVal))
                {
                    var grpTipVal = properties.GetValueOrDefault("tooltip");
                    ApplyGroupHyperlink(grpSlidePart, groupShape, grpLinkVal, grpTipVal);
                }

                if (properties.TryGetValue("zorder", out var grpZ)
                    || properties.TryGetValue("z-order", out grpZ)
                    || properties.TryGetValue("order", out grpZ))
                    ApplyZOrder(grpSlidePart, groupShape, grpZ);

                GetSlide(grpSlidePart).Save();

                var grpCount = grpShapeTree.Elements<GroupShape>().Count();
                var remainingShapes = grpShapeTree.Elements<Shape>().Count();
                var resultPath = $"/slide[{grpSlideIdx}]/group[{grpCount}]";
                // Warn about re-indexing: grouped shapes are removed from the shape tree
                Console.Error.WriteLine($"  Note: {toGroup.Count} shapes moved into group. Remaining shape count: {remainingShapes}. Shape indices have been re-numbered.");
                return resultPath;
    }


    /// <summary>
    /// Create an empty <p:grpSp> on the slide so subsequent
    /// `add shape parent=/slide[N]/group[K]` calls have a container to
    /// attach to. Path back: /slide[N]/group[K] (1-based, positional within
    /// the slide's group list — same convention as the populated-group
    /// branch). Required for `dump | batch` round-trip: dump emits a
    /// geometry-only group followed by per-child shape adds.
    /// </summary>
    private string AddEmptyGroup(SlidePart grpSlidePart, OpenXmlCompositeElement grpShapeTree, int grpSlideIdx,
                                 uint grpId, string grpName, int? index,
                                 Dictionary<string, string> properties, string parentPath = "")
    {
        long emptyX = (properties.TryGetValue("x", out var ex) || properties.TryGetValue("left", out ex)) ? ParseEmu(ex) : 0;
        long emptyY = (properties.TryGetValue("y", out var ey) || properties.TryGetValue("top", out ey)) ? ParseEmu(ey) : 0;
        long emptyCx = (properties.TryGetValue("width", out var ew) || properties.TryGetValue("cx", out ew)) ? ParseEmu(ew) : 0;
        long emptyCy = (properties.TryGetValue("height", out var eh) || properties.TryGetValue("cy", out eh)) ? ParseEmu(eh) : 0;
        // R53 bt-4: explicit childOffset / childExtent input ("EMU_X,EMU_Y").
        // When omitted, default to the outer offset/extent (identity mapping)
        // — same behavior as before. When the source group declared an
        // asymmetric chOff/chExt (NodeBuilder emits childOffset / childExtent),
        // honoring it here keeps dump→replay byte-faithful so the inner
        // shapes' positions resolve through the original child coord system
        // instead of silently snapping to the outer rect.
        long emptyChX = emptyX, emptyChY = emptyY, emptyChCx = emptyCx, emptyChCy = emptyCy;
        if (properties.TryGetValue("childOffset", out var emptyChOffVal))
        {
            var parts = emptyChOffVal.Split(',');
            if (parts.Length == 2
                && long.TryParse(parts[0], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var chX)
                && long.TryParse(parts[1], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var chY))
            { emptyChX = chX; emptyChY = chY; }
        }
        if (properties.TryGetValue("childExtent", out var emptyChExtVal))
        {
            var parts = emptyChExtVal.Split(',');
            if (parts.Length == 2
                && long.TryParse(parts[0], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var chCx)
                && long.TryParse(parts[1], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var chCy))
            { emptyChCx = chCx; emptyChCy = chCy; }
        }

        var groupShape = new GroupShape();
        groupShape.NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
            new NonVisualDrawingProperties { Id = grpId, Name = grpName },
            new NonVisualGroupShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()
        );
        groupShape.GroupShapeProperties = new GroupShapeProperties(
            new Drawing.TransformGroup(
                new Drawing.Offset { X = emptyX, Y = emptyY },
                new Drawing.Extents { Cx = emptyCx, Cy = emptyCy },
                new Drawing.ChildOffset { X = emptyChX, Y = emptyChY },
                new Drawing.ChildExtents { Cx = emptyChCx, Cy = emptyChCy }
            )
        );

        // Honor `rotation` for empty groups too. Dump emits `add group` with the
        // group's rotation when its source <a:xfrm rot> was non-zero; replay
        // formerly built the TransformGroup without a Rot attribute.
        if (properties.TryGetValue("rotation", out var emptyRot)
            || properties.TryGetValue("rotate", out emptyRot))
        {
            groupShape.GroupShapeProperties.TransformGroup!.Rotation =
                (int)(ParseHelpers.SafeParseRotationDegrees(emptyRot, "rotation") * 60000);
        }

        InsertAtPosition(grpShapeTree, groupShape, index);

        if (properties.TryGetValue("link", out var emptyLink) && !string.IsNullOrEmpty(emptyLink))
        {
            var emptyTip = properties.GetValueOrDefault("tooltip");
            ApplyGroupHyperlink(grpSlidePart, groupShape, emptyLink, emptyTip);
        }
        if (properties.TryGetValue("zorder", out var emptyZ)
            || properties.TryGetValue("z-order", out emptyZ)
            || properties.TryGetValue("order", out emptyZ))
            ApplyZOrder(grpSlidePart, groupShape, emptyZ);

        GetSlide(grpSlidePart).Save();
        var emptyCount = grpShapeTree.Elements<GroupShape>().Count();
        var parentPrefix = string.IsNullOrEmpty(parentPath) || parentPath == $"/slide[{grpSlideIdx}]"
            ? $"/slide[{grpSlideIdx}]" : parentPath;
        return $"{parentPrefix}/group[{emptyCount}]";
    }

    // CONSISTENCY(add-dispatch-shape): mirrors AddGroup/AddShape resolution flow.
    // Emits a <p:sp> with <p:ph type="..."/> that binds to the layout's matching
    // placeholder. Leaves <p:spPr> empty so PowerPoint inherits geometry/font
    // from the layout placeholder. Optional --prop text=... prepopulates text.
    private string AddPlaceholder(string parentPath, int? index, Dictionary<string, string> properties)
    {
        var phSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
        if (!phSlideMatch.Success)
            throw new ArgumentException("Placeholders must be added to a slide: /slide[N]");

        var phSlideIdx = int.Parse(phSlideMatch.Groups[1].Value);
        var phSlideParts = GetSlideParts().ToList();
        if (phSlideIdx < 1 || phSlideIdx > phSlideParts.Count)
            throw new ArgumentException($"Slide {phSlideIdx} not found (total: {phSlideParts.Count})");

        var phSlidePart = phSlideParts[phSlideIdx - 1];
        var phShapeTree = GetSlide(phSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // Bare <p:ph/> round-trip: the source placeholder had NO type and NO idx
        // (see NodeBuilder's phBare marker). It must replay bare so it stays
        // UNBOUND to any layout slot — binding it (type="body"+idx) would
        // inherit the slot's bullet/formatting the source deliberately opted out
        // of (formatting-bullet-indent). Build a minimal shape carrying an empty
        // <p:ph/> and forward the remaining props through Set exactly like the
        // normal path.
        if ((properties.TryGetValue("phBare", out var phBareVal) || properties.TryGetValue("phbare", out phBareVal))
            && IsTruthy(phBareVal))
        {
            var bareId = AcquireShapeId(phShapeTree, properties);
            var bareName = properties.GetValueOrDefault("name", $"Placeholder {bareId}");
            var bareShape = new Shape
            {
                NonVisualShapeProperties = new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = bareId, Name = bareName },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                ShapeProperties = new ShapeProperties(),
                TextBody = new TextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.EndParagraphRunProperties { Language = "en-US" })),
            };
            InsertAtPosition(phShapeTree, bareShape, index);
            if (properties.TryGetValue("zorder", out var bz) || properties.TryGetValue("z-order", out bz)
                || properties.TryGetValue("order", out bz))
                ApplyZOrder(phSlidePart, bareShape, bz);
            GetSlide(phSlidePart).Save();
            var bareCount = phShapeTree.Elements<Shape>().Count();
            var barePath = $"/slide[{phSlideIdx}]/shape[{bareCount}]";
            var bareConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "phBare", "phbare", "phType", "phtype", "type", "phIndex", "phindex", "idx",
                "name", "id", "zorder", "z-order", "order", "text", "isTitle", "istitle",
                "geometry", "noGrp", "nogrp",
            };
            var barePass = properties
                .Where(kv => !bareConsumed.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            if (barePass.Count > 0) Set(barePath, barePass);
            return barePath;
        }

        if (!properties.TryGetValue("phType", out var phTypeStr)
            && !properties.TryGetValue("phtype", out phTypeStr)
            && !properties.TryGetValue("type", out phTypeStr))
            throw new ArgumentException("'phType' property required for placeholder type (e.g. phType=body|date|footer|slidenum|header|subtitle|title)");

        var phTypeVal = ParsePlaceholderType(phTypeStr)
            ?? throw new ArgumentException(
                $"Invalid placeholder type: '{phTypeStr}'. Valid: title, body, subtitle, date, footer, slidenum, header, picture, chart, table, diagram, media, obj, clipart.");

        // Title/ctrTitle are the only placeholder types PowerPoint treats as
        // unique-per-slide (ECMA-376 §19.3.1.36; UI auto-deduplicates). Body,
        // object, picture, chart, etc. can legitimately appear multiple times
        // on one slide — each binds to a different layout slot via @idx
        // (two-content / comparison layouts; also bare <p:ph/> defaulted to
        // body where the source had two unattributed placeholders). The
        // uniqueness check is therefore restricted to title-family, and
        // (idx, type) pairs are deduplicated only when both placeholders share
        // the same explicit @idx. Otherwise dump→replay of a slide with two
        // bare <p:ph/> elements (both canonicalize to phType=body) throws on
        // the second Add — false positive against ECMA-376.
        bool phTypeIsTitleFamily = phTypeVal == PlaceholderValues.Title
            || phTypeVal == PlaceholderValues.CenteredTitle;
        // dump→replay escape: a malformed-but-real source slide can carry two
        // title/ctrTitle placeholders (PowerPoint keeps them — it only
        // auto-dedups via the UI, not on load). The uniqueness guard below is a
        // convenience for interactive `add`; for a faithful round-trip the
        // emitter sets allowDuplicate so the second title placeholder replays
        // instead of aborting (which also cascaded the slide's later shapes via
        // the broken shape count). The flag is consumed here, never written to XML.
        bool phAllowDuplicate = (properties.TryGetValue("allowDuplicate", out var phAdup)
                                  || properties.TryGetValue("allowduplicate", out phAdup))
                                 && IsTruthy(phAdup);
        properties.Remove("allowDuplicate");
        properties.Remove("allowduplicate");
        if (phTypeIsTitleFamily && !phAllowDuplicate)
        {
            var existingTitle = phShapeTree.Elements<Shape>()
                .FirstOrDefault(s =>
                {
                    var existingType = s.NonVisualShapeProperties
                        ?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>()?.Type?.Value;
                    return existingType == PlaceholderValues.Title
                        || existingType == PlaceholderValues.CenteredTitle;
                });
            if (existingTitle != null)
                throw new ArgumentException(
                    $"Placeholder phType='{phTypeStr}' already exists on slide {phSlideIdx}. " +
                    "Use Set to update the existing placeholder, or Remove the existing one first.");
        }

        var phId = AcquireShapeId(phShapeTree, properties);
        var phName = properties.GetValueOrDefault("name", $"{phTypeStr} Placeholder {phId}");

        // ECMA-376 §19.3.1.36: every non-title placeholder needs an @idx so the
        // slide-layout slot can be located by PowerPoint. Without
        // idx, the placeholder defaults to idx=0 which collides with title and
        // strips geometry/font inheritance. Strategy:
        //   1. If user passed phIndex explicitly, honor it.
        //   2. Else if the layout has a matching phType slot with idx, copy it.
        //   3. Else allocate the smallest non-zero idx not already used on slide.
        // Title (and centeredTitle) keep no idx — per spec the default 0 binds
        // to the layout title slot.
        uint? phIdx = null;
        bool isTitleType = phTypeVal == PlaceholderValues.Title
            || phTypeVal == PlaceholderValues.CenteredTitle;
        // Track whether the placeholder will bind to a layout slot. When it
        // does not, PowerPoint renders nothing because we leave ShapeProperties
        // empty (geometry pulled from layout). Below, we synthesize a fallback
        // Transform2D for the unbound case so the shape is at least visible.
        bool boundToLayout = false;
        // Check layout for a matching slot regardless of phIdx source.
        var layoutPartCheck = phSlidePart.SlideLayoutPart;
        var titleLayoutSlot = isTitleType
            ? layoutPartCheck?.SlideLayout?.CommonSlideData?.ShapeTree
                ?.Elements<Shape>()
                .Select(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>())
                .FirstOrDefault(p => p?.Type?.Value == PlaceholderValues.Title
                    || p?.Type?.Value == PlaceholderValues.CenteredTitle)
            : null;
        // R27-1: PowerPoint inherits placeholder geometry/typography by STRICT
        // ph-type match. A slide <p:ph type="title"> does NOT inherit from a
        // layout <p:ph type="ctrTitle"> (and vice versa) — the title then has no
        // resolvable position, renders at (0,0) and overlaps the subtitle.
        // The OR match above intentionally treats title/ctrTitle as one family
        // for "does a title slot exist", but to actually inherit we must adopt
        // the layout slot's exact type. Re-point phTypeVal to match the slot the
        // title will bind to (e.g. ctrTitle on the "Title Slide" layout).
        if (isTitleType && titleLayoutSlot?.Type?.Value is { } layoutTitleType
            && layoutTitleType != phTypeVal)
        {
            phTypeVal = layoutTitleType;
        }
        // Detect whether the caller explicitly provided an idx — distinguishes
        // "user passed no idx, want bare <p:ph type='subTitle'/>" from "user
        // didn't bother and we should pick one". Dump→batch replay relies on
        // this: NodeBuilder emits phIndex only when the source XML had an
        // idx attribute, so the absence of the key on the prop bag carries
        // semantic weight for the round trip. Without this distinction, a
        // bare <p:ph type='subTitle'/> source replayed as
        // <p:ph type='subTitle' idx='1'/>, and the idx=1 binding inherited
        // body's default bullet style from the layout/master cascade.
        bool callerProvidedIdx =
            properties.ContainsKey("phIndex")
            || properties.ContainsKey("phindex")
            || properties.ContainsKey("idx");
        if (isTitleType)
        {
            boundToLayout = titleLayoutSlot != null;
        }
        else
        {
            if ((properties.TryGetValue("phIndex", out var phIdxStr)
                    || properties.TryGetValue("phindex", out phIdxStr)
                    || properties.TryGetValue("idx", out phIdxStr))
                && uint.TryParse(phIdxStr, out var parsedIdx))
            {
                phIdx = parsedIdx;
                // User-specified idx: bound only if layout has matching slot.
                var slot = layoutPartCheck?.SlideLayout?.CommonSlideData?.ShapeTree
                    ?.Elements<Shape>()
                    .Select(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>())
                    .FirstOrDefault(p => p?.Index?.Value == parsedIdx);
                boundToLayout = slot != null;
            }
            else if (phTypeVal == PlaceholderValues.SubTitle && !callerProvidedIdx)
            {
                // Subtitle bound by type alone — leave Index unset so the
                // emitted <p:ph type="subTitle"/> matches a source that had
                // no idx attribute. Layout binding still resolves via type.
                var layoutMatch = layoutPartCheck?.SlideLayout?.CommonSlideData?.ShapeTree
                    ?.Elements<Shape>()
                    .Select(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>())
                    .FirstOrDefault(p => p?.Type?.Value == phTypeVal);
                boundToLayout = layoutMatch != null;
            }
            else if (!callerProvidedIdx && properties.ContainsKey("geometry"))
            {
                // DRIFT-4 — caller supplied an explicit geometry= prop. That's
                // the dump→replay signature for a TextBox-style placeholder
                // (source had <p:ph type=None/> + its own prstGeom). Forcing
                // idx=1 here would gain a spurious phIndex on round-trip and
                // (worse) re-binding to a layout body slot drops the explicit
                // prstGeom. Leave idx unset; layout binding falls through to
                // type-only match if any.
                var layoutMatch = layoutPartCheck?.SlideLayout?.CommonSlideData?.ShapeTree
                    ?.Elements<Shape>()
                    .Select(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>())
                    .FirstOrDefault(p => p?.Type?.Value == phTypeVal);
                boundToLayout = layoutMatch != null;
            }
            else
            {
                var layoutMatch = layoutPartCheck?.SlideLayout?.CommonSlideData?.ShapeTree
                    ?.Elements<Shape>()
                    .Select(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>())
                    .FirstOrDefault(p => p?.Type?.Value == phTypeVal && p.Index?.HasValue == true);
                if (layoutMatch != null) { phIdx = layoutMatch.Index!.Value; boundToLayout = true; }
                else
                {
                    var usedIdx = phShapeTree.Elements<Shape>()
                        .Select(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                            ?.GetFirstChild<PlaceholderShape>()?.Index?.Value)
                        .Where(v => v.HasValue)
                        .Select(v => v!.Value)
                        .ToHashSet();
                    uint next = 1;
                    while (usedIdx.Contains(next)) next++;
                    phIdx = next;
                }
            }
        }

        var shape = new Shape();
        var appNvPr = new ApplicationNonVisualDrawingProperties();
        var phElem = new PlaceholderShape { Type = phTypeVal };
        if (phIdx.HasValue) phElem.Index = phIdx.Value;
        // Placeholder @sz (full|half|quarter) — a real OOXML attribute on
        // <p:ph>. Without recognising it here, the value falls through to
        // Set's font-size branch which throws ArgumentException on "half".
        if ((properties.TryGetValue("size", out var phSizeStr)
                || properties.TryGetValue("sz", out phSizeStr))
            && phSizeStr is not null)
        {
            var phSizeKey = phSizeStr.Trim().ToLowerInvariant();
            if (phSizeKey == "full") phElem.Size = PlaceholderSizeValues.Full;
            else if (phSizeKey == "half") phElem.Size = PlaceholderSizeValues.Half;
            else if (phSizeKey == "quarter") phElem.Size = PlaceholderSizeValues.Quarter;
        }
        appNvPr.AppendChild(phElem);
        // CONSISTENCY(splocks-round-trip): <p:cNvSpPr><a:spLocks noGrp="1"/>
        // </p:cNvSpPr> is the on-disk marker that the placeholder cannot be
        // ungrouped — present on every PowerPoint-authored placeholder, but
        // ABSENT on placeholders generated by other authoring tools and on
        // some hand-crafted templates. NodeBuilder surfaces Format["noGrp"]
        // only when the source carried the lock, so dump→replay preserves
        // exactly what was on disk (R43 4a670cdf injected the lock by
        // default, which silently added <a:spLocks noGrp="1"/> to every
        // placeholder that originally had nothing — drift against the
        // source). Default to no lock; honor an explicit `noGrp=true` to
        // opt in.
        var phCNvSpPr = new NonVisualShapeDrawingProperties();
        bool phNoGrp = false;
        if (properties.TryGetValue("noGrp", out var phNoGrpStr)
            || properties.TryGetValue("nogrp", out phNoGrpStr))
            phNoGrp = IsTruthy(phNoGrpStr);
        if (phNoGrp)
            phCNvSpPr.AppendChild(new Drawing.ShapeLocks { NoGrouping = true });
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = phId, Name = phName },
            phCNvSpPr,
            appNvPr
        );
        // R27-1: write an EXPLICIT <a:xfrm> on every placeholder so its position
        // is self-contained. Type-matching the layout slot is necessary but NOT
        // sufficient — real PowerPoint does not always resolve geometry purely by
        // inheritance on this Add path, so a slide placeholder with empty spPr can
        // render at (0,0) and overlap a sibling (e.g. title over subtitle on the
        // Title Slide layout — caught when rendered in real Office). When the
        // placeholder binds to a layout slot, copy that slot's resolved off/ext;
        // otherwise fall back to the standard slot rectangle below.
        shape.ShapeProperties = new ShapeProperties();
        var resolvedLayoutGeom = boundToLayout
            ? ResolveLayoutSlotGeometry(layoutPartCheck, phTypeVal, phIdx)
            : null;
        {
            (long x, long y, long cx, long cy) geom = resolvedLayoutGeom ?? phTypeVal switch
            {
                _ when phTypeVal == PlaceholderValues.Title
                    || phTypeVal == PlaceholderValues.CenteredTitle
                        => (838200L, 365125L, 10515600L, 1325563L),
                _ when phTypeVal == PlaceholderValues.SubTitle
                        => (1371600L, 3886200L, 6400800L, 1752600L),
                _ when phTypeVal == PlaceholderValues.DateAndTime
                        => (838200L, 6356350L, 2895600L, 365125L),
                _ when phTypeVal == PlaceholderValues.Footer
                        => (3884613L, 6356350L, 4351338L, 365125L),
                _ when phTypeVal == PlaceholderValues.SlideNumber
                        => (8506463L, 6356350L, 2847338L, 365125L),
                _ => (838200L, 1825625L, 10515600L, 4351338L), // body/header/picture/chart/...
            };
            shape.ShapeProperties.AppendChild(new Drawing.Transform2D(
                new Drawing.Offset { X = geom.x, Y = geom.y },
                new Drawing.Extents { Cx = geom.cx, Cy = geom.cy }
            ));
            // R24 — do NOT inject <a:prstGeom prst="rect"/> by default. PPT
            // falls back to a rectangle when no geometry
            // is declared on a placeholder's spPr (the placeholder slot is
            // inherently rectangular), so the explicit element is redundant
            // for rendering. The cost of emitting it is real: NodeBuilder
            // surfaces it as `geometry=rect` in dump, the batch emitter
            // forwards it through Set, and Set's geometry path seeds a
            // default outline (bbe1a0c8) — so an idempotent dump+replay
            // grows a 1pt border around every formerly-unbound placeholder.
        }
        // DRIFT-4 — when caller explicitly supplies a geometry prop, honor it:
        // dump→replay of a source with <p:sp><a:prstGeom prst=".."/>... wrapped
        // as a placeholder must preserve the prstGeom; otherwise geometry
        // silently disappears on the second dump.
        if (properties.TryGetValue("geometry", out var phGeom) && !string.IsNullOrEmpty(phGeom))
        {
            var prstName = phGeom.Trim();
            if (prstName.Equals("custom", StringComparison.OrdinalIgnoreCase))
                prstName = "rect";
            if (TryParsePresetShape(prstName, out var prstEnum))
            {
                shape.ShapeProperties.AppendChild(
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = prstEnum }
                );
            }
        }

        // Optional text prepopulation. Build a minimal TextBody so PowerPoint
        // still renders layout placeholder typography.
        // CONSISTENCY(text-newline-split): mirror Set --prop text=... behavior —
        // a literal "\n" (backslash-n) or actual LF in the value spawns one
        // paragraph per line. Without this, Add stored "A\nB" as a single run
        // while Set on the same shape produced two paragraphs (asymmetric).
        var textBody = new TextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle()
        );
        if (properties.TryGetValue("text", out var phText) && phText.Length > 0)
        {
            XmlTextValidator.ValidateOrThrow(phText, "text", allowSoftBreakChar: true);
            // CONSISTENCY(text-escape-boundary): \n / \t resolution is at the
            // CLI --prop boundary; phText already contains real newlines.
            var lines = phText.Split('\n');
            foreach (var line in lines)
            {
                var p = new Drawing.Paragraph();
                if (line.Length > 0)
                {
                    p.AppendChild(new Drawing.Run(
                        new Drawing.RunProperties { Language = "en-US" },
                        new Drawing.Text(line)
                    ));
                }
                else
                {
                    p.AppendChild(new Drawing.EndParagraphRunProperties { Language = "en-US" });
                }
                textBody.AppendChild(p);
            }
        }
        else
        {
            // Empty paragraph is valid — PowerPoint shows the layout prompt text.
            var p = new Drawing.Paragraph();
            p.AppendChild(new Drawing.EndParagraphRunProperties { Language = "en-US" });
            textBody.AppendChild(p);
        }
        shape.TextBody = textBody;

        InsertAtPosition(phShapeTree, shape, index);
        if (properties.TryGetValue("zorder", out var phZ)
            || properties.TryGetValue("z-order", out phZ)
            || properties.TryGetValue("order", out phZ))
            ApplyZOrder(phSlidePart, shape, phZ);
        GetSlide(phSlidePart).Save();

        var shapeCount = phShapeTree.Elements<Shape>().Count();
        var phPath = $"/slide[{phSlideIdx}]/shape[{shapeCount}]";

        // CONSISTENCY(placeholder-prop-passthrough): AddPlaceholder previously
        // consumed only phType/phIndex/name/id/zorder/text and silently
        // dropped every other caller-supplied prop. That broke
        // dump→batch→replay for any placeholder whose source carried explicit
        // x/y/width/height/fill/font/color/line/... (i.e. every placeholder
        // overriding its layout slot). On replay the batch reported success
        // but Get returned layout defaults. Forward the leftover props through
        // Set so the same code path Add uses for plain shapes applies.
        var consumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "phType", "phtype", "type",
            "phIndex", "phindex", "idx",
            "size", "sz",
            "name", "id",
            "zorder", "z-order", "order",
            "text",
            // isTitle is a discriminator on Get but a no-op here: phType already
            // determines title-ness. Drop without forwarding so Set doesn't see
            // an unknown key.
            "isTitle", "istitle",
            // geometry on a placeholder is implicit (rect) — AddPlaceholder
            // already injected a PresetGeometry where needed. Forwarding would
            // be a no-op at best, an unsupported_property warning at worst.
            "geometry",
            // CONSISTENCY(splocks-round-trip): consumed above when building
            // the cNvSpPr ShapeLocks element. Do not forward to Set.
            "noGrp", "nogrp",
        };
        var passthrough = properties
            .Where(kv => !consumed.Contains(kv.Key))
            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
        if (passthrough.Count > 0)
            Set(phPath, passthrough);

        return phPath;
    }

    // R27-1: resolve the explicit off/ext geometry of the layout (or master)
    // placeholder slot that a slide placeholder of (phType, phIdx) binds to, so
    // AddPlaceholder can stamp it onto the slide shape directly instead of
    // relying on inheritance (which real PowerPoint does not always honor on the
    // Add path). Title-family slots match by type; everything else matches by
    // type and, when the slide placeholder carries an idx, by idx too. If the
    // layout slot has no xfrm of its own, fall back to the master's slot for the
    // same type. Returns null when no slot or no resolvable xfrm is found — the
    // caller then uses its default-rectangle table.
    private static (long x, long y, long cx, long cy)? ResolveLayoutSlotGeometry(
        SlideLayoutPart? layoutPart, PlaceholderValues phType, uint? phIdx)
    {
        if (layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null) return null;

        bool isTitle = phType == PlaceholderValues.Title
            || phType == PlaceholderValues.CenteredTitle;

        static (long, long, long, long)? XfrmOf(Shape? s)
        {
            var xfrm = s?.ShapeProperties?.Transform2D;
            var off = xfrm?.Offset;
            var ext = xfrm?.Extents;
            if (off?.X is null || off.Y is null || ext?.Cx is null || ext.Cy is null)
                return null;
            return (off.X!.Value, off.Y!.Value, ext.Cx!.Value, ext.Cy!.Value);
        }

        Shape? MatchSlot(OpenXmlElement? tree)
        {
            return tree?.Elements<Shape>().FirstOrDefault(s =>
            {
                var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                if (ph == null) return false;
                var t = ph.Type?.Value;
                if (isTitle)
                    return t == PlaceholderValues.Title || t == PlaceholderValues.CenteredTitle;
                // Non-title: @idx is the authoritative binding key (ECMA-376
                // §19.3.1.36). When the slide placeholder has an idx, match the
                // layout slot with the SAME idx regardless of its @type — a
                // layout slot may OMIT @type (it defaults to body/obj) yet still
                // be the inheritance source (e.g. <p:ph sz="quarter" idx="13"/>
                // carrying a custom off/ext). Requiring t == phType here missed
                // such slots, so ResolveLayoutSlotGeometry returned null and
                // AddPlaceholder stamped a generic default rectangle instead of
                // the master/layout-inherited geometry — visibly shrinking and
                // shifting the placeholder on round-trip.
                if (phIdx.HasValue)
                    return ph.Index?.Value == phIdx.Value;
                // No idx on the slide placeholder: fall back to a same-type slot.
                return t == phType;
            });
        }

        // Layout slot first (its xfrm overrides the master's).
        var layoutSlot = MatchSlot(layoutPart.SlideLayout.CommonSlideData.ShapeTree);
        if (XfrmOf(layoutSlot) is { } lg) return lg;

        // Fall back to the master slot of the same type.
        var masterTree = layoutPart.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
        var masterSlot = MatchSlot(masterTree);
        if (XfrmOf(masterSlot) is { } mg) return mg;

        return null;
    }

    private string AddAnimation(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Add animation to a shape (/slide[N]/shape[M]) or chart graphicFrame
                // (/slide[N]/chart[M]). Chart targets accept the additional chartBuild
                // prop (per-series/category build) and emit <p:bldGraphic> instead of
                // <p:bldP> in the slide's <p:bldLst>.
                // CONSISTENCY(animation-target): the timing tree binds by spid only —
                // both element kinds resolve to one through GetAnimationTargetSpId.
                var animMatch = System.Text.RegularExpressions.Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
                var animChartMatch = System.Text.RegularExpressions.Regex.Match(parentPath, @"^/slide\[(\d+)\]/chart\[(\d+)\]$");
                if (!animMatch.Success && !animChartMatch.Success)
                    throw new ArgumentException(
                        "Animations must be added to a top-level shape or chart: /slide[N]/shape[M] or /slide[N]/chart[M]. "
                        + "Shapes inside a group cannot be animated individually (PowerPoint animates the group as a whole) — "
                        + "animate the group or ungroup first.");

                SlidePart animSlidePart;
                DocumentFormat.OpenXml.OpenXmlElement animTarget;
                bool isChartTarget = false;
                int parentSlideIdx;
                int parentPositionalIdx;
                string parentKind;
                if (animChartMatch.Success)
                {
                    parentSlideIdx = int.Parse(animChartMatch.Groups[1].Value);
                    parentPositionalIdx = int.Parse(animChartMatch.Groups[2].Value);
                    parentKind = "chart";
                    var (sp, gf, _, _) = ResolveChart(parentSlideIdx, parentPositionalIdx);
                    animSlidePart = sp;
                    animTarget = gf;
                    isChartTarget = true;
                }
                else
                {
                    parentSlideIdx = int.Parse(animMatch.Groups[1].Value);
                    parentPositionalIdx = int.Parse(animMatch.Groups[2].Value);
                    parentKind = "shape";
                    var (sp, sh) = ResolveShape(parentSlideIdx, parentPositionalIdx);
                    animSlidePart = sp;
                    animTarget = sh;
                    // chartBuild is meaningless on plain shapes — hard-reject up
                    // front so the user finds the mistake at Add time instead of
                    // ApplyShapeAnimation deep inside the call stack.
                    if (properties.ContainsKey("chartBuild") || properties.ContainsKey("chartbuild"))
                        throw new ArgumentException(
                            "chartBuild only applies to chart targets. Use /slide[N]/chart[M] "
                            + "or remove the chartBuild prop.");
                }

                // L3 sub-B: class=motion routes to motion-path animation instead
                // of preset entrance/exit/emphasis. Preset path lookup ("line",
                // "arc", "circle", ...) translates to OOXML <p:animMotion path>.
                // path=custom requires d= to supply raw SVG-like data.
                if (properties.TryGetValue("class", out var maybeMotionCls)
                    && maybeMotionCls.Equals("motion", StringComparison.OrdinalIgnoreCase))
                {
                    if (isChartTarget)
                        throw new ArgumentException(
                            "Motion-path animations on a chart graphicFrame are not supported. "
                            + "Use class=entrance/exit/emphasis with optional chartBuild=series|category|...");
                    return AddMotionAnimation(parentPath, animSlidePart, (Shape)animTarget, properties);
                }

                // Build animation value string from properties
                var effect = properties.GetValueOrDefault("effect", "fade");
                var explicitCls = properties.GetValueOrDefault("class");
                // bt-1 / fuzz-1 fix: detect class suffix on effect (fly-out,
                // zoom-in, wipe-entrance, fade-exit). If user did not pass an
                // explicit class= property, the suffix wins over the default
                // "entrance". Reject contradictory class tokens (fly-in-out)
                // rather than silently keeping the last one.
                var (effectStripped, suffixCls) = ParseEffectClassSuffix(effect);
                effect = effectStripped;
                var cls = explicitCls ?? suffixCls ?? "entrance";
                // Validate class enum up front — composite animValue parsing
                // silently falls back to entrance on unknown class tokens
                // (stderr warning only), so callers got success + wrong cls.
                // Mirror the hard-reject pattern used for trigger / effect.
                ValidateAnimationClass(cls);
                // CONSISTENCY(animation-dur-alias): accept "dur" as alias for
                // "duration" — mirrors the short name used elsewhere (transition
                // dur attribute) and matches user intuition.
                var duration = properties.GetValueOrDefault("duration")
                    ?? properties.GetValueOrDefault("dur", "500");
                // OOXML @dur is ST_PositiveUniversalMeasure (>= 0). Schema declares
                // duration as integer ms — reject unit suffixes (500ms), fractions
                // (500.7), non-numeric garbage, and bare negatives. The composite
                // animValue parser would silently default these to 400 with a
                // stderr-only warning.
                ValidateAnimationDuration(duration);
                var trigger = properties.GetValueOrDefault("trigger", "onclick");

                // Validate delay symmetrically with duration. The composite
                // animValue split('-') silently drops the minus sign on a
                // negative delay token, leaving delay=0 with no error.
                if (properties.TryGetValue("delay", out var rawDelay))
                    ValidateAnimationDelay(rawDelay);

                // L2 props (repeat, restart, autoReverse) — validate up front
                // for a hard error rather than relying on the composite parser
                // (which silently ignores unknown key=value segments).
                if (properties.TryGetValue("repeat", out var rawRepeat))
                    ValidateAnimationRepeat(rawRepeat);
                if (properties.TryGetValue("restart", out var rawRestart))
                    ValidateAnimationRestart(rawRestart);
                if (properties.TryGetValue("autoReverse", out var rawAutoRev)
                    || properties.TryGetValue("autoreverse", out rawAutoRev))
                    ValidateAnimationAutoReverse(rawAutoRev);

                // Map trigger property to animation format
                var triggerPart = trigger.ToLowerInvariant() switch
                {
                    "onclick" or "click" => "click",
                    "after" or "afterprevious" => "after",
                    "with" or "withprevious" => "with",
                    _ => throw new ArgumentException($"Invalid animation trigger: '{trigger}'. Valid values: onclick, click, after, afterprevious, with, withprevious.")
                };

                var animValue = $"{effect}-{cls}-{duration}-{triggerPart}";

                // Append delay/easing properties if specified
                if (properties.TryGetValue("delay", out var delay))
                    animValue += $"-delay={delay}";
                if (properties.TryGetValue("easein", out var easein))
                    animValue += $"-easein={easein}";
                if (properties.TryGetValue("easeout", out var easeout))
                    animValue += $"-easeout={easeout}";
                if (properties.TryGetValue("easing", out var easing))
                    animValue += $"-easing={easing}";
                if (properties.TryGetValue("direction", out var dir))
                    animValue += $"-{dir}";
                if (properties.TryGetValue("repeat", out var repProp))
                    animValue += $"-repeat={repProp}";
                if (properties.TryGetValue("restart", out var restartProp))
                    animValue += $"-restart={restartProp}";
                if (properties.TryGetValue("autoReverse", out var arProp)
                    || properties.TryGetValue("autoreverse", out arProp))
                    animValue += $"-autoReverse={arProp}";

                // Validate + thread chartBuild on chart targets. Routed through
                // the composite animValue string so ApplyShapeAnimation's existing
                // parser picks it up alongside repeat / restart / autoReverse.
                if (isChartTarget
                    && (properties.TryGetValue("chartBuild", out var rawChartBuild)
                        || properties.TryGetValue("chartbuild", out rawChartBuild)))
                {
                    ValidateAnimationChartBuild(rawChartBuild);
                    animValue += $"-chartBuild={rawChartBuild}";
                }
                // R14-bug6: thread buildType on plain-shape targets so dump→batch
                // round-trips an animation that iterates by paragraph (build="p").
                // ApplyShapeAnimation's parser picks it up from the composite
                // animValue and stamps the bldP element accordingly. Reject on
                // chart targets (chart-build vocabulary is chartBuild=*).
                if (!isChartTarget
                    && (properties.TryGetValue("buildType", out var rawBuildType)
                        || properties.TryGetValue("buildtype", out rawBuildType)))
                {
                    animValue += $"-buildType={rawBuildType}";
                }

                ApplyShapeAnimation(animSlidePart, animTarget, animValue);
                GetSlide(animSlidePart).Save();

                // Count animations on this target — must match Get's enumeration
                // (effect-bearing CommonTimeNodes), not raw ShapeTarget references.
                // CONSISTENCY(animation-index): mirror EnumerateShapeAnimationCTns
                // in Query.cs — counting ShapeTargets over-counts effects like
                // fly/swivel that emit multiple p:anim per single user effect,
                // returning a stale path like animation[2] for the first add.
                var animCount = EnumerateShapeAnimationCTns(animSlidePart, animTarget).Count;
                // CONSISTENCY(query-path-roundtrip): rebuild parent segment so the
                // returned path uses @id= form when the target has a cNvPr id —
                // matches AddShape/AddPicture/etc which always emit the id form
                // via BuildElementPathSegment. ResolveIdPath at Add.cs:37 already
                // collapsed @id= input to a positional [N] form before dispatch,
                // so reconstructing here is the only way to surface @id= back.
                var rebuiltParent = $"/slide[{parentSlideIdx}]/{BuildElementPathSegment(parentKind, animTarget, parentPositionalIdx)}";
                return $"{rebuiltParent}/animation[{animCount}]";
    }

    // L3 sub-B: motion-path animation handler (class=motion). Supports a small
    // set of preset paths (line / arc / circle / diamond / triangle / square)
    // with optional direction= for line/arc; custom path requires d=. Appends
    // to the shape's animation chain so animation[K] indexing remains uniform.
    // CONSISTENCY(animation-chain): mirrors AddAnimation's append behavior.
    private string AddMotionAnimation(string parentPath,
        DocumentFormat.OpenXml.Packaging.SlidePart slidePart,
        DocumentFormat.OpenXml.Presentation.Shape shape,
        Dictionary<string, string> properties)
    {
        var preset = properties.GetValueOrDefault("path");
        if (string.IsNullOrEmpty(preset))
            throw new ArgumentException(
                "class=motion requires path=<preset>. Valid presets: "
                + string.Join(", ", KnownMotionPresets())
                + ". Use path=custom with d=<SVG-like path data> for a custom motion path.");

        string pathString;
        if (preset.Equals("custom", StringComparison.OrdinalIgnoreCase))
        {
            if (!properties.TryGetValue("d", out var customD) || string.IsNullOrEmpty(customD))
                throw new ArgumentException(
                    "path=custom requires d=<SVG-like path data> (e.g. d='M 0 0 L 0.5 0 E'). "
                    + "Coords are relative to slide (0..1).");
            pathString = customD;
            // Ensure path is terminated with E so PowerPoint accepts it.
            if (!pathString.TrimEnd().EndsWith("E", StringComparison.OrdinalIgnoreCase))
                pathString = pathString.TrimEnd() + " E";
        }
        else
        {
            var direction = properties.GetValueOrDefault("direction");
            var resolved = GetMotionPresetPath(preset, direction);
            if (resolved == null)
                throw new ArgumentException(
                    $"Unknown motion path preset: '{preset}'. Valid presets: "
                    + string.Join(", ", KnownMotionPresets()) + ".");
            pathString = resolved;
        }

        var duration = properties.GetValueOrDefault("duration")
                       ?? properties.GetValueOrDefault("dur", "2000");
        ValidateAnimationDuration(duration);
        var durationMs = int.Parse(duration, System.Globalization.CultureInfo.InvariantCulture);

        var trigger = properties.GetValueOrDefault("trigger", "onclick");
        var triggerEnum = trigger.ToLowerInvariant() switch
        {
            "onclick" or "click"            => PowerPointHandler.AnimTrigger.OnClick,
            "after" or "afterprevious"      => PowerPointHandler.AnimTrigger.AfterPrevious,
            "with" or "withprevious"        => PowerPointHandler.AnimTrigger.WithPrevious,
            _ => throw new ArgumentException(
                $"Invalid animation trigger: '{trigger}'. Valid values: onclick, click, after, afterprevious, with, withprevious.")
        };

        int delayMs = 0, easingAccel = 0, easingDecel = 0;
        if (properties.TryGetValue("delay", out var dlyRaw))
        {
            ValidateAnimationDelay(dlyRaw);
            delayMs = int.Parse(dlyRaw, System.Globalization.CultureInfo.InvariantCulture);
        }
        if (properties.TryGetValue("easein", out var einRaw)
            && int.TryParse(einRaw, out var einV)) easingAccel = einV * 1000;
        if (properties.TryGetValue("easeout", out var eoutRaw)
            && int.TryParse(eoutRaw, out var eoutV)) easingDecel = eoutV * 1000;

        AppendMotionPathAnimation(slidePart, shape, pathString, durationMs,
            triggerEnum, delayMs, easingAccel, easingDecel);
        GetSlide(slidePart).Save();

        var animCount = EnumerateShapeAnimationCTns(slidePart, shape).Count;
        // CONSISTENCY(query-path-roundtrip): same as AddAnimation — rebuild
        // parent shape segment so @id= form survives ResolveIdPath collapse.
        var motionMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
        if (motionMatch.Success)
        {
            var motionSlideIdx = int.Parse(motionMatch.Groups[1].Value);
            var motionPosIdx = int.Parse(motionMatch.Groups[2].Value);
            var motionRebuilt = $"/slide[{motionSlideIdx}]/{BuildElementPathSegment("shape", shape, motionPosIdx)}";
            return $"{motionRebuilt}/animation[{animCount}]";
        }
        return $"{parentPath}/animation[{animCount}]";
    }


    private string AddZoom(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var zmSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!zmSlideMatch.Success)
                    throw new ArgumentException("Zoom must be added to a slide: /slide[N]");

                // Target slide (required)
                if (!properties.TryGetValue("target", out var targetStr) && !properties.TryGetValue("slide", out targetStr))
                    throw new ArgumentException("'target' property required for zoom type (target slide number, e.g. target=2)");
                if (!int.TryParse(targetStr, out var targetSlideNum))
                    throw new ArgumentException($"Invalid 'target' value: '{targetStr}'. Expected a slide number.");

                var zmSlideIdx = int.Parse(zmSlideMatch.Groups[1].Value);
                var zmSlideParts = GetSlideParts().ToList();
                if (zmSlideIdx < 1 || zmSlideIdx > zmSlideParts.Count)
                    throw new ArgumentException($"Slide {zmSlideIdx} not found (total: {zmSlideParts.Count})");
                if (targetSlideNum < 1 || targetSlideNum > zmSlideParts.Count)
                    throw new ArgumentException($"Target slide {targetSlideNum} not found (total: {zmSlideParts.Count})");

                var zmSlidePart = zmSlideParts[zmSlideIdx - 1];
                var zmShapeTree = GetSlide(zmSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");
                var targetSlidePart = zmSlideParts[targetSlideNum - 1];

                // Get target slide's SlideId from presentation.xml
                var zmPresentation = _doc.PresentationPart?.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                var zmSlideIdList = zmPresentation.GetFirstChild<SlideIdList>()
                    ?? throw new InvalidOperationException("No slides");
                var zmSlideIds = zmSlideIdList.Elements<SlideId>().ToList();
                var targetSldId = zmSlideIds[targetSlideNum - 1].Id!.Value;

                // Position and size (default: 8cm x 4.5cm, centered)
                long zmCx = 3048000; // ~8cm
                long zmCy = 1714500; // ~4.5cm
                if (properties.TryGetValue("width", out var zmW)) zmCx = ParseEmu(zmW);
                if (properties.TryGetValue("height", out var zmH)) zmCy = ParseEmu(zmH);
                var (zmSlideW, zmSlideH) = GetSlideSize();
                long zmX = (zmSlideW - zmCx) / 2;
                long zmY = (zmSlideH - zmCy) / 2;
                if (properties.TryGetValue("x", out var zmXStr)) zmX = ParseEmu(zmXStr);
                if (properties.TryGetValue("y", out var zmYStr)) zmY = ParseEmu(zmYStr);

                var returnToParent = properties.TryGetValue("returntoparent", out var rtp) && IsTruthy(rtp) ? "1" : "0";
                var transitionDur = properties.GetValueOrDefault("transitiondur", "1000");

                // Generate shape IDs
                var zmShapeId = AcquireShapeId(zmShapeTree, properties);
                var zmName = properties.GetValueOrDefault("name", $"Slide Zoom {GetZoomElements(zmShapeTree).Count + 1}");
                var zmGuid = Guid.NewGuid().ToString("B").ToUpperInvariant();
                var zmCreationId = Guid.NewGuid().ToString("B").ToUpperInvariant();

                // Create a minimal 1x1 gray placeholder PNG (PowerPoint regenerates the thumbnail on open)
                byte[] placeholderPng = GenerateZoomPlaceholderPng();
                var zmImagePart = zmSlidePart.AddImagePart(ImagePartType.Png);
                using (var ms = new MemoryStream(placeholderPng))
                    zmImagePart.FeedData(ms);
                var zmImageRelId = zmSlidePart.GetIdOfPart(zmImagePart);

                // Create slide-to-slide relationship for fallback hyperlink
                var zmSlideRelId = zmSlidePart.CreateRelationshipToPart(targetSlidePart);

                // Build mc:AlternateContent programmatically (same pattern as morph transition)
                var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                var pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
                var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
                var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var pslzNs = "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom";
                var p166Ns = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
                var a16Ns = "http://schemas.microsoft.com/office/drawing/2014/main";

                var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);

                // === mc:Choice (for clients that support Slide Zoom) ===
                var choiceElement = new OpenXmlUnknownElement("mc", "Choice", mcNs);
                choiceElement.SetAttribute(new OpenXmlAttribute("", "Requires", null!, "pslz"));
                choiceElement.AddNamespaceDeclaration("pslz", pslzNs);

                var gfElement = new OpenXmlUnknownElement("p", "graphicFrame", pNs);
                gfElement.AddNamespaceDeclaration("a", aNs);
                gfElement.AddNamespaceDeclaration("r", rNs);

                // nvGraphicFramePr
                var nvGfPr = new OpenXmlUnknownElement("p", "nvGraphicFramePr", pNs);
                var cNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
                cNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmShapeId.ToString()));
                cNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, zmName));
                // creationId extension
                var extLst = new OpenXmlUnknownElement("a", "extLst", aNs);
                var ext = new OpenXmlUnknownElement("a", "ext", aNs);
                ext.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
                var creationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
                creationId.SetAttribute(new OpenXmlAttribute("", "id", null!, zmCreationId));
                ext.AppendChild(creationId);
                extLst.AppendChild(ext);
                cNvPr.AppendChild(extLst);
                nvGfPr.AppendChild(cNvPr);

                var cNvGfSpPr = new OpenXmlUnknownElement("p", "cNvGraphicFramePr", pNs);
                var gfLocks = new OpenXmlUnknownElement("a", "graphicFrameLocks", aNs);
                gfLocks.SetAttribute(new OpenXmlAttribute("", "noChangeAspect", null!, "1"));
                cNvGfSpPr.AppendChild(gfLocks);
                nvGfPr.AppendChild(cNvGfSpPr);
                nvGfPr.AppendChild(new OpenXmlUnknownElement("p", "nvPr", pNs));
                gfElement.AppendChild(nvGfPr);

                // xfrm (position/size)
                var gfXfrm = new OpenXmlUnknownElement("p", "xfrm", pNs);
                var gfOff = new OpenXmlUnknownElement("a", "off", aNs);
                gfOff.SetAttribute(new OpenXmlAttribute("", "x", null!, zmX.ToString()));
                gfOff.SetAttribute(new OpenXmlAttribute("", "y", null!, zmY.ToString()));
                var gfExt = new OpenXmlUnknownElement("a", "ext", aNs);
                gfExt.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                gfExt.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                gfXfrm.AppendChild(gfOff);
                gfXfrm.AppendChild(gfExt);
                gfElement.AppendChild(gfXfrm);

                // graphic > graphicData > pslz:sldZm
                var graphic = new OpenXmlUnknownElement("a", "graphic", aNs);
                var graphicData = new OpenXmlUnknownElement("a", "graphicData", aNs);
                graphicData.SetAttribute(new OpenXmlAttribute("", "uri", null!, pslzNs));

                var sldZm = new OpenXmlUnknownElement("pslz", "sldZm", pslzNs);
                var sldZmObj = new OpenXmlUnknownElement("pslz", "sldZmObj", pslzNs);
                sldZmObj.SetAttribute(new OpenXmlAttribute("", "sldId", null!, targetSldId.ToString()));
                sldZmObj.SetAttribute(new OpenXmlAttribute("", "cId", null!, "0"));

                var zmPr = new OpenXmlUnknownElement("pslz", "zmPr", pslzNs);
                zmPr.AddNamespaceDeclaration("p166", p166Ns);
                zmPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmGuid));
                zmPr.SetAttribute(new OpenXmlAttribute("", "returnToParent", null!, returnToParent));
                zmPr.SetAttribute(new OpenXmlAttribute("", "transitionDur", null!, transitionDur));

                // blipFill (thumbnail)
                var blipFill = new OpenXmlUnknownElement("p166", "blipFill", p166Ns);
                var blip = new OpenXmlUnknownElement("a", "blip", aNs);
                blip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, zmImageRelId));
                blipFill.AppendChild(blip);
                var stretch = new OpenXmlUnknownElement("a", "stretch", aNs);
                stretch.AppendChild(new OpenXmlUnknownElement("a", "fillRect", aNs));
                blipFill.AppendChild(stretch);
                zmPr.AppendChild(blipFill);

                // spPr (shape properties inside zoom)
                var zmSpPr = new OpenXmlUnknownElement("p166", "spPr", p166Ns);
                var zmSpXfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
                var zmSpOff = new OpenXmlUnknownElement("a", "off", aNs);
                zmSpOff.SetAttribute(new OpenXmlAttribute("", "x", null!, "0"));
                zmSpOff.SetAttribute(new OpenXmlAttribute("", "y", null!, "0"));
                var zmSpExt = new OpenXmlUnknownElement("a", "ext", aNs);
                zmSpExt.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                zmSpExt.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                zmSpXfrm.AppendChild(zmSpOff);
                zmSpXfrm.AppendChild(zmSpExt);
                zmSpPr.AppendChild(zmSpXfrm);
                var prstGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
                prstGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
                prstGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
                zmSpPr.AppendChild(prstGeom);
                var zmLn = new OpenXmlUnknownElement("a", "ln", aNs);
                zmLn.SetAttribute(new OpenXmlAttribute("", "w", null!, "3175"));
                var zmLnFill = new OpenXmlUnknownElement("a", "solidFill", aNs);
                var zmLnClr = new OpenXmlUnknownElement("a", "prstClr", aNs);
                zmLnClr.SetAttribute(new OpenXmlAttribute("", "val", null!, "ltGray"));
                zmLnFill.AppendChild(zmLnClr);
                zmLn.AppendChild(zmLnFill);
                zmSpPr.AppendChild(zmLn);
                zmPr.AppendChild(zmSpPr);

                sldZmObj.AppendChild(zmPr);
                sldZm.AppendChild(sldZmObj);
                graphicData.AppendChild(sldZm);
                graphic.AppendChild(graphicData);
                gfElement.AppendChild(graphic);
                choiceElement.AppendChild(gfElement);

                // === mc:Fallback (pic + hyperlink for older clients) ===
                var fallbackElement = new OpenXmlUnknownElement("mc", "Fallback", mcNs);
                var fbPic = new OpenXmlUnknownElement("p", "pic", pNs);
                fbPic.AddNamespaceDeclaration("a", aNs);
                fbPic.AddNamespaceDeclaration("r", rNs);

                var fbNvPicPr = new OpenXmlUnknownElement("p", "nvPicPr", pNs);
                var fbCNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
                fbCNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmShapeId.ToString()));
                fbCNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, zmName));
                var hlinkClick = new OpenXmlUnknownElement("a", "hlinkClick", aNs);
                hlinkClick.SetAttribute(new OpenXmlAttribute("r", "id", rNs, zmSlideRelId));
                hlinkClick.SetAttribute(new OpenXmlAttribute("", "action", null!, "ppaction://hlinksldjump"));
                fbCNvPr.AppendChild(hlinkClick);
                // Same creationId
                var fbExtLst = new OpenXmlUnknownElement("a", "extLst", aNs);
                var fbExt = new OpenXmlUnknownElement("a", "ext", aNs);
                fbExt.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
                var fbCreationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
                fbCreationId.SetAttribute(new OpenXmlAttribute("", "id", null!, zmCreationId));
                fbExt.AppendChild(fbCreationId);
                fbExtLst.AppendChild(fbExt);
                fbCNvPr.AppendChild(fbExtLst);
                fbNvPicPr.AppendChild(fbCNvPr);

                var fbCNvPicPr = new OpenXmlUnknownElement("p", "cNvPicPr", pNs);
                var picLocks = new OpenXmlUnknownElement("a", "picLocks", aNs);
                foreach (var lockAttr in new[] { "noGrp", "noRot", "noChangeAspect", "noMove", "noResize",
                    "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType" })
                    picLocks.SetAttribute(new OpenXmlAttribute("", lockAttr, null!, "1"));
                fbCNvPicPr.AppendChild(picLocks);
                fbNvPicPr.AppendChild(fbCNvPicPr);
                fbNvPicPr.AppendChild(new OpenXmlUnknownElement("p", "nvPr", pNs));
                fbPic.AppendChild(fbNvPicPr);

                // Fallback blipFill
                var fbBlipFill = new OpenXmlUnknownElement("p", "blipFill", pNs);
                var fbBlip = new OpenXmlUnknownElement("a", "blip", aNs);
                fbBlip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, zmImageRelId));
                fbBlipFill.AppendChild(fbBlip);
                var fbStretch = new OpenXmlUnknownElement("a", "stretch", aNs);
                fbStretch.AppendChild(new OpenXmlUnknownElement("a", "fillRect", aNs));
                fbBlipFill.AppendChild(fbStretch);
                fbPic.AppendChild(fbBlipFill);

                // Fallback spPr
                var fbSpPr = new OpenXmlUnknownElement("p", "spPr", pNs);
                var fbXfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
                var fbOff = new OpenXmlUnknownElement("a", "off", aNs);
                fbOff.SetAttribute(new OpenXmlAttribute("", "x", null!, zmX.ToString()));
                fbOff.SetAttribute(new OpenXmlAttribute("", "y", null!, zmY.ToString()));
                var fbExtSz = new OpenXmlUnknownElement("a", "ext", aNs);
                fbExtSz.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                fbExtSz.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                fbXfrm.AppendChild(fbOff);
                fbXfrm.AppendChild(fbExtSz);
                fbSpPr.AppendChild(fbXfrm);
                var fbGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
                fbGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
                fbGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
                fbSpPr.AppendChild(fbGeom);
                var fbLn = new OpenXmlUnknownElement("a", "ln", aNs);
                fbLn.SetAttribute(new OpenXmlAttribute("", "w", null!, "3175"));
                var fbLnFill = new OpenXmlUnknownElement("a", "solidFill", aNs);
                var fbLnClr = new OpenXmlUnknownElement("a", "prstClr", aNs);
                fbLnClr.SetAttribute(new OpenXmlAttribute("", "val", null!, "ltGray"));
                fbLnFill.AppendChild(fbLnClr);
                fbLn.AppendChild(fbLnFill);
                fbSpPr.AppendChild(fbLn);
                fbPic.AppendChild(fbSpPr);

                fallbackElement.AppendChild(fbPic);

                acElement.AppendChild(choiceElement);
                acElement.AppendChild(fallbackElement);
                InsertAtPosition(zmShapeTree, acElement, index);
                GetSlide(zmSlidePart).Save();

                var zmCount = zmShapeTree.ChildElements
                    .Count(e => e.LocalName == "AlternateContent");
                return $"/slide[{zmSlideIdx}]/zoom[{zmCount}]";
    }


    private string AddDefault(string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
                // Try resolving logical paths (table/placeholder) first
                var logicalResult = ResolveLogicalPath(parentPath);
                SlidePart fbSlidePart;
                OpenXmlElement fbParent;

                if (logicalResult.HasValue)
                {
                    fbSlidePart = logicalResult.Value.slidePart;
                    fbParent = logicalResult.Value.element;
                }
                else
                {
                    // Generic fallback: navigate by XML localName
                    var allSegments = GenericXmlQuery.ParsePathSegments(parentPath);
                    if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                        throw new ArgumentException($"Generic add requires a path starting with /slide[N]: {parentPath}");

                    var fbSlideIdx = allSegments[0].Index!.Value;
                    var fbSlideParts = GetSlideParts().ToList();
                    if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                        throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

                    fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                    fbParent = GetSlide(fbSlidePart);
                    var remaining = allSegments.Skip(1).ToList();
                    if (remaining.Count > 0)
                    {
                        fbParent = GenericXmlQuery.NavigateByPath(fbParent, remaining)
                            ?? throw new ArgumentException(
                                parentPath.Contains("chart", StringComparison.OrdinalIgnoreCase) &&
                                (parentPath.Contains("series", StringComparison.OrdinalIgnoreCase) ||
                                 type.Equals("trendline", StringComparison.OrdinalIgnoreCase))
                                    ? $"Cannot add child elements to chart sub-paths via Add. " +
                                      $"To add trendlines, use: Set /slide[N]/chart[1] --prop series1.trendline=linear"
                                    : $"Parent element not found: {parentPath}");
                    }
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent, type, properties, index);
                if (created == null)
                    throw new CliException($"Unknown element type '{type}' for {parentPath}. " +
                        "Valid types: slide, shape, textbox, picture, table, chart, ole (object, embed), paragraph, run, connector, group, video, audio, model3d (3dmodel), equation, notes, zoom. " +
                        "Use 'officecli pptx add' for details.")
                        { Code = "invalid_type" };

                GetSlide(fbSlidePart).Save();

                // Build result path
                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = PathIndex.FromArrayIndex(siblings.IndexOf(created));
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
    }

    /// <summary>
    /// Parse trailing class-suffix tokens off an animation effect name.
    /// Returns the stripped effect plus the resolved class ("entrance"/"exit"/
    /// "emphasis") or null if no suffix is present. Throws when contradictory
    /// class tokens appear in the effect string (e.g. "fly-in-out").
    /// CONSISTENCY(animation-class-suffix): shared by AddAnimation and
    /// SetShapeAnimationByPath so Add and Set route class identically.
    /// </summary>
    private static (string effect, string? cls) ParseEffectClassSuffix(string effect)
    {
        if (string.IsNullOrEmpty(effect)) return (effect, null);

        static string? ClassOf(string seg) => seg switch
        {
            "in" or "entrance" or "entr" => "entrance",
            "out" or "exit" => "exit",
            "emph" or "emphasis" => "emphasis",
            _ => null
        };

        // Scan all dash-separated segments for class tokens. Reject any pair
        // of segments that resolve to different classes — silently keeping the
        // last token has bitten users (fuzz-1: fly-in-out vs fly-out-in).
        var segs = effect.Split('-');
        string? seenClass = null;
        string? seenToken = null;
        for (int i = 1; i < segs.Length; i++)
        {
            var c = ClassOf(segs[i].ToLowerInvariant());
            if (c == null) continue;
            if (seenClass != null && seenClass != c)
                throw new ArgumentException(
                    $"Animation effect '{effect}' has contradictory class tokens "
                    + $"'{seenToken}' ({seenClass}) and '{segs[i]}' ({c}). "
                    + "Pass exactly one of: in/out/entrance/exit/emphasis, "
                    + "or use the class= property.");
            seenClass = c;
            seenToken = segs[i];
        }

        // Strip only a trailing class suffix from the effect name (preserve
        // pre-existing direction/duration tokens that other parsers handle).
        // Exception: when the full-form name (normalized, dashes removed) is
        // a known template effect — e.g. "float-out" ↔ Float Out preset, "fade-out"
        // ↔ Fade Out exit — keep the full name so the registry lookup hits the
        // right template. CONSISTENCY(animation-template-name): registry keys
        // are normalized identifiers, so "float-out" and "floatout" both map.
        var dashIdx = effect.LastIndexOf('-');
        if (dashIdx > 0)
        {
            var tailCls = ClassOf(effect[(dashIdx + 1)..].ToLowerInvariant());
            if (tailCls != null)
            {
                // Probe both class buckets — if the full-form effect resolves to
                // a template under the suffix-implied class, do NOT strip.
                var classEnum = tailCls switch
                {
                    "exit" => DocumentFormat.OpenXml.Presentation.TimeNodePresetClassValues.Exit,
                    "entrance" => DocumentFormat.OpenXml.Presentation.TimeNodePresetClassValues.Entrance,
                    "emphasis" => DocumentFormat.OpenXml.Presentation.TimeNodePresetClassValues.Emphasis,
                    _ => (DocumentFormat.OpenXml.Presentation.TimeNodePresetClassValues?)null
                };
                if (classEnum.HasValue && TryGetEffectTemplate(effect, classEnum.Value) != null)
                    return (effect, tailCls);
                return (effect[..dashIdx], tailCls);
            }
        }
        return (effect, seenClass);
    }
}
