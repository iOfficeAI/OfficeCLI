// Copyright 2025 OfficeCLI (officecli.ai)
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
                var cxnSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!cxnSlideMatch.Success)
                    throw new ArgumentException("Connectors must be added to a slide: /slide[N]");

                var cxnSlideIdx = int.Parse(cxnSlideMatch.Groups[1].Value);
                var cxnSlideParts = GetSlideParts().ToList();
                if (cxnSlideIdx < 1 || cxnSlideIdx > cxnSlideParts.Count)
                    throw new ArgumentException($"Slide {cxnSlideIdx} not found (total: {cxnSlideParts.Count})");

                var cxnSlidePart = cxnSlideParts[cxnSlideIdx - 1];
                var cxnShapeTree = GetSlide(cxnSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

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

                var hasFrom = properties.ContainsKey("from") || properties.ContainsKey("startshape") || properties.ContainsKey("startShape");
                var hasTo = properties.ContainsKey("to") || properties.ContainsKey("endshape") || properties.ContainsKey("endShape");

                long cxnX = (properties.TryGetValue("x", out var cx1) || properties.TryGetValue("left", out cx1)) ? ParseEmu(cx1) : 2000000;
                long cxnY = (properties.TryGetValue("y", out var cy1) || properties.TryGetValue("top", out cy1)) ? ParseEmu(cy1) : 3000000;
                long cxnCx = properties.TryGetValue("width", out var cw) ? ParseEmu(cw) : 4000000;
                long cxnCy = properties.TryGetValue("height", out var ch) ? ParseEmu(ch) : 0;
                var cxnFlipH = false;
                var cxnFlipV = false;
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
                        var (sx, sy, scx, scy) = pStart.Value;
                        var (ex, ey, ecx, ecy) = pEnd.Value;
                        var p1x = sx + scx / 2;
                        var p1y = sy + scy / 2;
                        var p2x = ex + ecx / 2;
                        var p2y = ey + ecy / 2;
                        if (!hasX) cxnX = Math.Min(p1x, p2x);
                        if (!hasY) cxnY = Math.Min(p1y, p2y);
                        if (!hasW) cxnCx = Math.Abs(p2x - p1x);
                        if (!hasH) cxnCy = Math.Abs(p2y - p1y);
                        // Encode start/end ordering via flipH/flipV (mirrors PowerPoint).
                        cxnFlipH = p2x < p1x;
                        cxnFlipV = p2y < p1y;
                    }
                }
                // CONSISTENCY(positive-size): mirror Add.Shape negative-size guard so picture
                // / chart / connector / media all reject inverted dimensions instead of silently
                // emitting negative cx/cy that PowerPoint draws as flipped or 0-sized boxes.
                if (cxnCx < 0) throw new ArgumentException($"Negative width is not allowed: '{cw}'.");
                if (cxnCy < 0) throw new ArgumentException($"Negative height is not allowed: '{ch}'.");

                var connector = new ConnectionShape();
                var cxnNvProps = new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = cxnId, Name = cxnName },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );

                // Connect to shapes if specified
                var cxnDrawProps = cxnNvProps.NonVisualConnectorShapeDrawingProperties!;
                if (properties.TryGetValue("startshape", out var startId) || properties.TryGetValue("startShape", out startId)
                    || properties.TryGetValue("from", out startId))
                {
                    var startIdVal = ResolveShapeId(startId!, cxnShapeTree);
                    cxnDrawProps.StartConnection = new Drawing.StartConnection { Id = startIdVal, Index = 0 };
                }
                if (properties.TryGetValue("endshape", out var endId) || properties.TryGetValue("endShape", out endId)
                    || properties.TryGetValue("to", out endId))
                {
                    var endIdVal = ResolveShapeId(endId!, cxnShapeTree);
                    cxnDrawProps.EndConnection = new Drawing.EndConnection { Id = endIdVal, Index = 0 };
                }

                connector.NonVisualConnectionShapeProperties = cxnNvProps;
                var cxnTransform = new Drawing.Transform2D(
                    new Drawing.Offset { X = cxnX, Y = cxnY },
                    new Drawing.Extents { Cx = cxnCx, Cy = cxnCy }
                );
                if (cxnFlipH) cxnTransform.HorizontalFlip = true;
                if (cxnFlipV) cxnTransform.VerticalFlip = true;
                connector.ShapeProperties = new ShapeProperties(
                    cxnTransform,
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    {
                        // CONSISTENCY(canonical-key): canonical 'shape'; 'preset' legacy alias.
                        Preset = (properties.GetValueOrDefault("shape")
                                  ?? properties.GetValueOrDefault("preset", "straightConnector1")).ToLowerInvariant() switch
                        {
                            // Short canonical names + OOXML full names. "line" is a
                            // historical schema alias for the straight preset; bent/curved
                            // accept either the 2-segment or 3-segment OOXML variant
                            // (PowerPoint maps both to the same drawing primitive set).
                            "straight" or "straightconnector1" or "line" => Drawing.ShapeTypeValues.StraightConnector1,
                            "elbow" or "bentconnector3" or "bentconnector2" => Drawing.ShapeTypeValues.BentConnector3,
                            "curve" or "curvedconnector3" or "curvedconnector2" => Drawing.ShapeTypeValues.CurvedConnector3,
                            _ => throw new ArgumentException($"Invalid connector shape: '{properties.GetValueOrDefault("shape") ?? properties.GetValueOrDefault("preset", "straightConnector1")}'. Valid values: straight, elbow, curve (or OOXML full names: straightConnector1, bentConnector3, curvedConnector3).")
                        }
                    }
                );

                // Line style
                var cxnOutline = new Drawing.Outline { Width = 12700 }; // 1pt default
                // line.gradient parity with Set side — accept gradient outline at Add time
                // so dump→replay round-trips. Gradient fill wins over solid color.
                if (properties.TryGetValue("line.gradient", out var cxnLineGrad)
                    || properties.TryGetValue("linegradient", out cxnLineGrad))
                {
                    cxnOutline.AppendChild(BuildGradientFill(cxnLineGrad));
                }
                else if (properties.TryGetValue("lineColor", out var cxnColor2) || properties.TryGetValue("linecolor", out cxnColor2)
                    || properties.TryGetValue("line", out cxnColor2) || properties.TryGetValue("color", out cxnColor2)
                    || properties.TryGetValue("line.color", out cxnColor2))
                    cxnOutline.AppendChild(BuildSolidFill(cxnColor2));
                else
                    cxnOutline.AppendChild(BuildSolidFill("000000"));
                if (properties.TryGetValue("linewidth", out var lwVal) || properties.TryGetValue("lineWidth", out lwVal)
                    || properties.TryGetValue("line.width", out lwVal))
                    cxnOutline.Width = Core.EmuConverter.ParseLineWidth(lwVal);
                if (properties.TryGetValue("lineDash", out var cxnDash) || properties.TryGetValue("linedash", out cxnDash))
                {
                    cxnOutline.AppendChild(new Drawing.PresetDash { Val = ParseLineDashValue(cxnDash) });
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
                connector.ShapeProperties.AppendChild(cxnOutline);

                InsertAtPosition(cxnShapeTree, connector, index);
                if (properties.TryGetValue("zorder", out var cxnZ)
                    || properties.TryGetValue("z-order", out cxnZ)
                    || properties.TryGetValue("order", out cxnZ))
                    ApplyZOrder(cxnSlidePart, connector, cxnZ);
                GetSlide(cxnSlidePart).Save();

                return $"/slide[{cxnSlideIdx}]/{BuildElementPathSegment("connector", connector, cxnShapeTree.Elements<ConnectionShape>().Count())}";
    }

    /// <summary>
    /// Resolves a shape reference to an OOXML shape ID.
    /// Accepts: plain integer (shape ID), or DOM path like /slide[1]/shape[2] (resolves Nth shape's ID).
    /// </summary>
    private static uint ResolveShapeId(string value, ShapeTree shapeTree)
    {
        // Try plain integer first (shape ID)
        if (uint.TryParse(value, out var directId))
        {
            var shapes = shapeTree.Elements<Shape>().ToList();
            // If directId matches an actual shape ID, use it directly
            if (shapes.Any(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == directId))
                return directId;
            // Otherwise treat as 1-based shape index
            if (directId >= 1 && directId <= (uint)shapes.Count)
            {
                var shape = shapes[(int)directId - 1];
                return shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? directId;
            }
            return directId;
        }

        // Try @id path form first: /slide[N]/shape[@id=M] (as returned by `query shape`).
        // CONSISTENCY(query-path-roundtrip): query shape returns @id form; Add must accept it.
        var atIdMatch = Regex.Match(value, @"/slide\[\d+\]/shape\[@id=(\d+)\]");
        if (atIdMatch.Success)
        {
            var atId = uint.Parse(atIdMatch.Groups[1].Value);
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (!shapes.Any(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == atId))
                throw new ArgumentException($"Shape @id={atId} not found on this slide");
            return atId;
        }

        // Try @name path form: /slide[N]/shape[@name=Foo]
        // CONSISTENCY: every other PPTX op accepts @name= selectors; connector from=/to= must too.
        var atNameMatch = Regex.Match(value, @"/slide\[\d+\]/shape\[@name=([^\]]+)\]");
        if (atNameMatch.Success)
        {
            var atName = atNameMatch.Groups[1].Value;
            var shapes = shapeTree.Elements<Shape>().ToList();
            var matched = shapes.FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == atName);
            if (matched == null)
                throw new ArgumentException($"Shape @name={atName} not found on this slide");
            return matched.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                ?? throw new ArgumentException($"Shape @name={atName} has no ID");
        }

        // Try DOM path: /slide[N]/shape[M] (positional)
        var pathMatch = Regex.Match(value, @"/slide\[\d+\]/shape\[(\d+)\]");
        if (pathMatch.Success)
        {
            var shapeIdx = int.Parse(pathMatch.Groups[1].Value);
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (shapeIdx < 1 || shapeIdx > shapes.Count)
                throw new ArgumentException($"Shape index {shapeIdx} out of range (total: {shapes.Count})");
            return shapes[shapeIdx - 1].NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                ?? throw new ArgumentException($"Shape {shapeIdx} has no ID");
        }

        throw new ArgumentException($"Invalid shape reference: '{value}'. Expected a shape index (1, 2, ...), path (/slide[N]/shape[M]), @id path (/slide[N]/shape[@id=M]), or @name path (/slide[N]/shape[@name=Foo]).");
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

                var shapeParts = shapesStr.Split(',');
                var shapeIndices = new List<int>();
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
                        // — id/name/positional lookup is by-position within
                        // grpFrameList regardless of which kind name the user used.
                        const string frameKind =
                            @"(?:shape|group|picture|pic|connector|connection|chart|table|graphicframe|graphicFrame|ole|object|embed|video|audio)";

                        // @id path: /slide[N]/<kind>[@id=M] — round-trips from `query`
                        var atIdMatch = Regex.Match(trimmed,
                            $@"/slide\[\d+\]/{frameKind}\[@id=(\d+)\]",
                            RegexOptions.IgnoreCase);
                        if (atIdMatch.Success)
                        {
                            var atId = uint.Parse(atIdMatch.Groups[1].Value);
                            var idx = grpFrameList.FindIndex(e => FrameId(e) == atId);
                            if (idx < 0)
                                throw new ArgumentException($"Frame @id={atId} not found on this slide");
                            shapeIndices.Add(idx + 1);
                            continue;
                        }
                        // @name path: /slide[N]/<kind>[@name=Foo]
                        var atNameMatch = Regex.Match(trimmed,
                            $@"/slide\[\d+\]/{frameKind}\[@name=([^\]]+)\]",
                            RegexOptions.IgnoreCase);
                        if (atNameMatch.Success)
                        {
                            var atName = atNameMatch.Groups[1].Value;
                            var idx = grpFrameList.FindIndex(e => FrameName(e) == atName);
                            if (idx < 0)
                                throw new ArgumentException($"Frame @name={atName} not found on this slide");
                            shapeIndices.Add(idx + 1);
                            continue;
                        }
                        // Positional path: /slide[N]/<kind>[M]
                        var pathMatch = Regex.Match(trimmed,
                            $@"/slide\[\d+\]/{frameKind}\[(\d+)\]",
                            RegexOptions.IgnoreCase);
                        if (!pathMatch.Success)
                            throw new ArgumentException($"Invalid frame path: '{trimmed}'. Expected /slide[N]/<kind>[M], /slide[N]/<kind>[@id=ID], or /slide[N]/<kind>[@name=Foo] where <kind> is shape/group/picture/connector/chart/table/etc.");
                        shapeIndices.Add(int.Parse(pathMatch.Groups[1].Value));
                    }
                    else if (int.TryParse(trimmed, out var idx))
                    {
                        shapeIndices.Add(idx);
                    }
                    else
                    {
                        throw new ArgumentException($"Invalid 'shapes' value: '{trimmed}' is not a valid integer or DOM path. Expected comma-separated shape indices (e.g. shapes=1,2,3) or DOM paths (e.g. shapes=/slide[1]/shape[1],/slide[1]/shape[2]).");
                    }
                }
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
                // @name= / path forms above.
                var allShapes = grpShapeTree.ChildElements
                    .Where(c => c is Shape || c is GroupShape || c is Picture
                        || c is GraphicFrame || c is ConnectionShape)
                    .Where(c => !(c is Shape s
                        && s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                            ?.GetFirstChild<PlaceholderShape>() != null))
                    .ToList();

                // Collect shapes to group (in reverse order to maintain indices during removal)
                var toGroup = new List<OpenXmlElement>();
                foreach (var si in shapeIndices.OrderBy(i => i))
                {
                    if (si < 1 || si > allShapes.Count)
                        throw new ArgumentException($"Shape {si} not found (total: {allShapes.Count})");
                    toGroup.Add(allShapes[si - 1]);
                }

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
                new Drawing.ChildOffset { X = emptyX, Y = emptyY },
                new Drawing.ChildExtents { Cx = emptyCx, Cy = emptyCy }
            )
        );

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

        if (!properties.TryGetValue("phType", out var phTypeStr)
            && !properties.TryGetValue("phtype", out phTypeStr)
            && !properties.TryGetValue("type", out phTypeStr))
            throw new ArgumentException("'phType' property required for placeholder type (e.g. phType=body|date|footer|slidenum|header|subtitle|title)");

        var phTypeVal = ParsePlaceholderType(phTypeStr)
            ?? throw new ArgumentException(
                $"Invalid placeholder type: '{phTypeStr}'. Valid: title, body, subtitle, date, footer, slidenum, header, picture, chart, table, diagram, media, obj, clipart.");

        var phId = AcquireShapeId(phShapeTree, properties);
        var phName = properties.GetValueOrDefault("name", $"{phTypeStr} Placeholder {phId}");

        // ECMA-376 §19.3.1.36: every non-title placeholder needs an @idx so the
        // slide-layout slot can be located by PowerPoint / LibreOffice. Without
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
        appNvPr.AppendChild(phElem);
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = phId, Name = phName },
            new NonVisualShapeDrawingProperties(),
            appNvPr
        );
        // Leave ShapeProperties empty when layout supplies geometry — but when
        // the slide's layout has no matching <p:ph> slot (e.g. user added a
        // body placeholder to a Blank-layout slide), PowerPoint and LibreOffice
        // render NOTHING. Inject a sensible default rectangle so the shape is
        // at least visible. Coordinates picked to roughly mirror the standard
        // "Title and Content" layout slots (16:9 deck, EMU = 914400/inch).
        shape.ShapeProperties = new ShapeProperties();
        if (!boundToLayout)
        {
            (long x, long y, long cx, long cy) geom = phTypeVal switch
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
            // R24 — do NOT inject <a:prstGeom prst="rect"/>. PPT and
            // LibreOffice both fall back to a rectangle when no geometry is
            // declared on a placeholder's spPr (the placeholder slot is
            // inherently rectangular), so the explicit element is redundant
            // for rendering. The cost of emitting it is real: NodeBuilder
            // surfaces it as `geometry=rect` in dump, the batch emitter
            // forwards it through Set, and Set's geometry path seeds a
            // default outline (bbe1a0c8) — so an idempotent dump+replay
            // grows a 1pt border around every formerly-unbound placeholder.
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
            XmlTextValidator.ValidateOrThrow(phText, "text");
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
        };
        var passthrough = properties
            .Where(kv => !consumed.Contains(kv.Key))
            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
        if (passthrough.Count > 0)
            Set(phPath, passthrough);

        return phPath;
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
                    throw new ArgumentException("Animations must be added to a shape or chart: /slide[N]/shape[M] or /slide[N]/chart[M]");

                SlidePart animSlidePart;
                DocumentFormat.OpenXml.OpenXmlElement animTarget;
                bool isChartTarget = false;
                if (animChartMatch.Success)
                {
                    var slideIdx = int.Parse(animChartMatch.Groups[1].Value);
                    var chartIdx = int.Parse(animChartMatch.Groups[2].Value);
                    var (sp, gf, _, _) = ResolveChart(slideIdx, chartIdx);
                    animSlidePart = sp;
                    animTarget = gf;
                    isChartTarget = true;
                }
                else
                {
                    var animSlideIdx = int.Parse(animMatch.Groups[1].Value);
                    var animShapeIdx = int.Parse(animMatch.Groups[2].Value);
                    var (sp, sh) = ResolveShape(animSlideIdx, animShapeIdx);
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

                ApplyShapeAnimation(animSlidePart, animTarget, animValue);
                GetSlide(animSlidePart).Save();

                // Count animations on this target — must match Get's enumeration
                // (effect-bearing CommonTimeNodes), not raw ShapeTarget references.
                // CONSISTENCY(animation-index): mirror EnumerateShapeAnimationCTns
                // in Query.cs — counting ShapeTargets over-counts effects like
                // fly/swivel that emit multiple p:anim per single user effect,
                // returning a stale path like animation[2] for the first add.
                var animCount = EnumerateShapeAnimationCTns(animSlidePart, animTarget).Count;
                return $"{parentPath}/animation[{animCount}]";
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
                        "Valid types: slide, shape, textbox, picture, table, chart, ole (object, embed), paragraph, run, connector, group, video, audio, equation, notes, zoom. " +
                        "Use 'officecli pptx add' for details.")
                        { Code = "invalid_type" };

                GetSlide(fbSlidePart).Save();

                // Build result path
                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
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
        var dashIdx = effect.LastIndexOf('-');
        if (dashIdx > 0)
        {
            var tailCls = ClassOf(effect[(dashIdx + 1)..].ToLowerInvariant());
            if (tailCls != null)
                return (effect[..dashIdx], tailCls);
        }
        return (effect, seenClass);
    }
}
