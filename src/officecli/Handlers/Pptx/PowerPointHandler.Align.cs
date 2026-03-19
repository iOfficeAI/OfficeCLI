// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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

        var slideWidth = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>()?.Cx?.Value ?? 9144000L;
        var slideHeight = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>()?.Cy?.Value ?? 6858000L;

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
            // Accept "shape[N]" or just "N"
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
}
