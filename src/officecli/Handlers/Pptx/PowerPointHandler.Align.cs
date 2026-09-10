// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Align & Distribute ====================

    /// <summary>
    /// Align shapes on a slide along one axis.
    /// align: left | center | right | top | middle | bottom
    /// targets: comma-separated paths, e.g. "shape[1],shape[2],shape[3]"
    ///          If null/empty, all shapes on the slide are aligned.
    /// Alignment is relative to the bounding box of the selected shapes.
    /// Special values: "slide-left", "slide-center", etc. — align relative to slide.
    /// </summary>
    private void AlignShapes(SlidePart slidePart, string alignValue, string? targets)
    {
        var shapes = ResolveAlignTargets(slidePart, targets);
        if (shapes.Count < 1) return;

        var boxes = shapes.Select(GetTransform2D).ToList();

        var (slideWidth, slideHeight) = GetSlideSize();

        bool relative = alignValue.StartsWith("slide-", StringComparison.OrdinalIgnoreCase);
        var mode = relative ? alignValue[6..].ToLowerInvariant() : alignValue.ToLowerInvariant();

        // Bounding box of all selected shapes (for relative-to-selection alignment)
        long refLeft = relative ? 0 : boxes.Where(b => b != null).Min(b => b!.Offset?.X?.Value ?? 0);
        long refTop = relative ? 0 : boxes.Where(b => b != null).Min(b => b!.Offset?.Y?.Value ?? 0);
        long refRight = relative ? slideWidth : boxes.Where(b => b != null)
            .Max(b => (b!.Offset?.X?.Value ?? 0) + (b.Extents?.Cx?.Value ?? 0));
        long refBottom = relative ? slideHeight : boxes.Where(b => b != null)
            .Max(b => (b!.Offset?.Y?.Value ?? 0) + (b.Extents?.Cy?.Value ?? 0));
        long refCenterX = (refLeft + refRight) / 2;
        long refCenterY = (refTop + refBottom) / 2;

        for (int i = 0; i < shapes.Count; i++)
        {
            var xfrm = boxes[i];
            if (xfrm?.Offset == null || xfrm.Extents == null) continue;

            var w = xfrm.Extents.Cx?.Value ?? 0;
            var h = xfrm.Extents.Cy?.Value ?? 0;

            switch (mode)
            {
                case "left":
                    xfrm.Offset.X = refLeft;
                    break;
                case "center" or "hcenter" or "centerh":
                    xfrm.Offset.X = refCenterX - w / 2;
                    break;
                case "right":
                    xfrm.Offset.X = refRight - w;
                    break;
                case "top":
                    xfrm.Offset.Y = refTop;
                    break;
                case "middle" or "vcenter" or "centerv":
                    xfrm.Offset.Y = refCenterY - h / 2;
                    break;
                case "bottom":
                    xfrm.Offset.Y = refBottom - h;
                    break;
                default:
                    throw new ArgumentException(
                        $"Invalid align value: '{alignValue}'. Valid: left, center, right, top, middle, bottom, " +
                        "slide-left, slide-center, slide-right, slide-top, slide-middle, slide-bottom");
            }
        }

        // Re-glue any connector anchored to a moved shape — same "connector follows
        // shape" behavior the plain per-shape x/y set path triggers. Without this,
        // align/distribute silently detached glued connectors (stale connector xfrm).
        RerouteConnectorsForMovedShapes(slidePart, shapes);
    }

    /// <summary>
    /// Distribute shapes evenly on a slide.
    /// distribute: horizontal | vertical
    /// targets: comma-separated paths (need at least 3 shapes for meaningful distribution)
    /// Distributes shapes so gaps between them are equal.
    /// </summary>
    private void DistributeShapes(SlidePart slidePart, string distributeValue, string? targets)
    {
        var shapes = ResolveAlignTargets(slidePart, targets);
        if (shapes.Count < 3) return;

        var boxes = shapes.Select(GetTransform2D).ToList();
        var mode = distributeValue.ToLowerInvariant();

        if (mode is "horizontal" or "h" or "horiz")
        {
            // Sort shapes by their left edge
            var sorted = shapes.Zip(boxes)
                .Where(p => p.Second?.Offset != null && p.Second.Extents != null)
                .OrderBy(p => p.Second!.Offset!.X!.Value)
                .ToList();
            if (sorted.Count < 3) return;

            var first = sorted.First().Second!;
            var last = sorted.Last().Second!;
            long totalWidth = sorted.Sum(p => p.Second!.Extents!.Cx!.Value);
            long span = (last.Offset!.X!.Value + last.Extents!.Cx!.Value) - first.Offset!.X!.Value;
            long gap = (span - totalWidth) / (sorted.Count - 1);

            long cursor = first.Offset.X.Value;
            foreach (var (_, xfrm) in sorted)
            {
                if (xfrm?.Offset != null)
                    xfrm.Offset.X = cursor;
                cursor += (xfrm?.Extents?.Cx?.Value ?? 0) + gap;
            }
        }
        else if (mode is "vertical" or "v" or "vert")
        {
            var sorted = shapes.Zip(boxes)
                .Where(p => p.Second?.Offset != null && p.Second.Extents != null)
                .OrderBy(p => p.Second!.Offset!.Y!.Value)
                .ToList();
            if (sorted.Count < 3) return;

            var first = sorted.First().Second!;
            var last = sorted.Last().Second!;
            long totalHeight = sorted.Sum(p => p.Second!.Extents!.Cy!.Value);
            long span = (last.Offset!.Y!.Value + last.Extents!.Cy!.Value) - first.Offset!.Y!.Value;
            long gap = (span - totalHeight) / (sorted.Count - 1);

            long cursor = first.Offset.Y.Value;
            foreach (var (_, xfrm) in sorted)
            {
                if (xfrm?.Offset != null)
                    xfrm.Offset.Y = cursor;
                cursor += (xfrm?.Extents?.Cy?.Value ?? 0) + gap;
            }
        }
        else
        {
            throw new ArgumentException(
                $"Invalid distribute value: '{distributeValue}'. Valid: horizontal, vertical");
        }

        // Re-glue connectors anchored to the redistributed shapes (see AlignShapes).
        RerouteConnectorsForMovedShapes(slidePart, shapes);
    }

    /// <summary>
    /// After align/distribute repositions shapes, re-run the connector reroute for
    /// each moved shape so glued connectors track them — the same behavior the
    /// per-shape x/y set path already provides via RerouteConnectorsForShape.
    /// </summary>
    private void RerouteConnectorsForMovedShapes(SlidePart slidePart, List<Shape> shapes)
    {
        foreach (var s in shapes)
        {
            var id = s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id.HasValue)
                RerouteConnectorsForShape(slidePart, id.Value);
        }
    }

    /// <summary>
    /// Resolve target shapes from a comma-separated list of shape paths (relative to the slide).
    /// Accepts "shape[N]", "picture[N]", etc. or empty (= all shapes).
    /// </summary>
    private List<Shape> ResolveAlignTargets(SlidePart slidePart, string? targets)
    {
        var tree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (tree == null) return [];

        if (string.IsNullOrWhiteSpace(targets))
            return tree.Elements<Shape>().ToList();

        var result = new List<Shape>();
        var allShapes = tree.Elements<Shape>().ToList();

        foreach (var token in targets.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            // Accept "shape[@id=N]" (the path form get/query emit — e.g.
            // /slide[1]/shape[@id=100000]), "shape[N]" (1-based positional), or a
            // bare "N" (positional). The @id form lets callers pass the exact
            // paths they already hold from a selection without first mapping them
            // to positional indices (the positional order also shifts as shapes
            // are added/removed, so @id is the stable reference).
            var idMatch = Regex.Match(token, @"@id=(\d+)");
            if (idMatch.Success)
            {
                var id = idMatch.Groups[1].Value;
                var byId = allShapes.FirstOrDefault(s =>
                    s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString() == id);
                if (byId != null) result.Add(byId);
                continue;
            }
            var m = Regex.Match(token, @"shape\[(\d+)\]|^(\d+)$");
            if (m.Success)
            {
                var idx = int.Parse(m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value) - 1;
                if (idx >= 0 && idx < allShapes.Count)
                    result.Add(allShapes[idx]);
            }
        }
        return result;
    }

    private static Drawing.Transform2D? GetTransform2D(Shape shape) =>
        shape.ShapeProperties?.Transform2D;

    /// <summary>
    /// First-class geometric layout entry for the `layout` verb (CLI / batch /
    /// MCP / resident share this core): align and/or distribute the shapes on
    /// one slide in a single call. Thin wrapper over AlignShapes /
    /// DistributeShapes — the same engine the slide-level `set --prop
    /// align=/distribute=` path uses — with slide-path validation and a
    /// caller-readable summary. Throws ArgumentException on a non-slide path,
    /// missing operation, or invalid value (valid lists in the engine errors).
    /// </summary>
    public string LayoutSlide(string slidePath, string? align, string? distribute, string? targets)
    {
        var m = Regex.Match(slidePath, @"^/slide\[(\d+)\]$");
        if (!m.Success)
            throw new ArgumentException(
                $"'layout' path must be a slide path (/slide[N]). Got: '{slidePath}'. " +
                "Example: layout deck.pptx /slide[2] --align bottom");
        var idx = int.Parse(m.Groups[1].Value);
        var parts = GetSlideParts().ToList();
        if (idx < 1 || idx > parts.Count)
            throw new ArgumentException($"Slide {idx} not found (total: {parts.Count})");
        if (string.IsNullOrWhiteSpace(align) && string.IsNullOrWhiteSpace(distribute))
            throw new ArgumentException(
                "'layout' requires --align and/or --distribute. " +
                "Example: layout deck.pptx /slide[2] --align bottom --distribute horizontal");
        var slidePart = parts[PathIndex.ToArrayIndex(idx)];
        var msgs = new List<string>();
        if (!string.IsNullOrWhiteSpace(align))
        {
            var n = ResolveAlignTargets(slidePart, targets).Count;
            AlignShapes(slidePart, align, targets);
            msgs.Add($"aligned {n} shape(s) {align}");
        }
        if (!string.IsNullOrWhiteSpace(distribute))
        {
            var n = ResolveAlignTargets(slidePart, targets).Count;
            DistributeShapes(slidePart, distribute, targets);
            msgs.Add(n >= 3
                ? $"distributed {n} shape(s) {distribute}"
                : "distribute skipped: need at least 3 shapes with geometry");
        }
        return string.Join("; ", msgs) + $" on /slide[{idx}]";
    }
}
