// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== CSS Helper: Fill ====================

    // CONSISTENCY(shape-fill-css): this solidFill/gradFill/pattFill → CSS mapping
    // is ~70% structurally duplicated by WordHandler.ResolveShapeFillCss
    // (Word/WordHandler.HtmlPreview.Css.cs). They diverge on element access
    // (SDK-typed GetFirstChild here vs untyped LocalName scan there) and ride
    // different tint/shade extraction before both delegate to ColorMath. Deferred
    // Core consolidation (e.g. Core/ShapeFillCss) — do NOT land as a one-handler
    // special case; unify cross-handler in one pass once the docx fix-storm settles.
    private static string GetShapeFillCss(ShapeProperties? spPr, OpenXmlPart part, Dictionary<string, string> themeColors)
    {
        if (spPr == null) return "";

        // NoFill
        if (spPr.GetFirstChild<Drawing.NoFill>() != null)
            return "background:transparent";

        // Solid fill
        var solidFill = spPr.GetFirstChild<Drawing.SolidFill>();
        if (solidFill != null)
        {
            var color = ResolveFillColor(solidFill, themeColors);
            if (color != null) return $"background:{color}";
        }

        // Gradient fill
        var gradFill = spPr.GetFirstChild<Drawing.GradientFill>();
        if (gradFill != null)
            return $"background:{GradientToCss(gradFill, themeColors)}";

        // Image fill (blip)
        var blipFill = spPr.GetFirstChild<Drawing.BlipFill>();
        if (blipFill != null)
        {
            var dataUri = BlipToDataUri(blipFill, part);
            if (dataUri != null)
            {
                // R4-4: honor <a:tile> — repeat at native size rather than cover.
                var tile = blipFill.GetFirstChild<Drawing.Tile>();
                if (tile != null)
                {
                    // R49-01: honor sx/sy scale (×1000 percent of NATIVE pixel
                    // size) and algn. CSS percentages are container-relative, so
                    // compute pixel size from the decoded image's natural size.
                    var sx = tile.HorizontalRatio?.Value;
                    var sy = tile.VerticalRatio?.Value;
                    string bgSize = "auto";
                    var nonDefaultScale = (sx.HasValue && sx.Value != 100000)
                        || (sy.HasValue && sy.Value != 100000);
                    if (nonDefaultScale)
                    {
                        var nat = TryGetBlipDimensions(blipFill, part);
                        var sxPct = (sx ?? 100000) / 100000.0;
                        var syPct = (sy ?? 100000) / 100000.0;
                        bgSize = nat != null
                            ? $"{nat.Value.Width * sxPct:0.##}px {nat.Value.Height * syPct:0.##}px"
                            : $"{sxPct * 100:0.##}% {syPct * 100:0.##}%";
                    }
                    var bgPos = TileAlignToBackgroundPosition(tile.Alignment);
                    return $"background:url('{dataUri}') {bgPos} repeat;background-size:{bgSize}";
                }
                // R49-02: honor <a:stretch><a:fillRect> insets — inset the image
                // from each edge instead of full-cover.
                var stretch = blipFill.GetFirstChild<Drawing.Stretch>();
                var fillRect = stretch?.FillRectangle;
                if (fillRect != null)
                {
                    double fl = (fillRect.Left?.Value ?? 0) / 1000.0;
                    double ft = (fillRect.Top?.Value ?? 0) / 1000.0;
                    double fr = (fillRect.Right?.Value ?? 0) / 1000.0;
                    double fb = (fillRect.Bottom?.Value ?? 0) / 1000.0;
                    if (fl != 0 || ft != 0 || fr != 0 || fb != 0)
                    {
                        var sizeW = Math.Max(100.0 - fl - fr, 0.01).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        var sizeH = Math.Max(100.0 - ft - fb, 0.01).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        var dX = fl + fr; var dY = ft + fb;
                        var px = (dX != 0 ? fl / dX * 100.0 : 0.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        var py = (dY != 0 ? ft / dY * 100.0 : 0.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        return $"background:transparent url('{dataUri}') {px}% {py}%/{sizeW}% {sizeH}% no-repeat";
                    }
                }
                // R9-5: honor <a:srcRect> crop insets (l/t/r/b, units of 1/1000%).
                // Emulate the crop by zooming the background past the box and
                // shifting its origin so the visible region maps to the kept rect.
                var srcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>();
                if (srcRect != null)
                {
                    double l = (srcRect.Left?.Value ?? 0) / 1000.0;
                    double t = (srcRect.Top?.Value ?? 0) / 1000.0;
                    double r = (srcRect.Right?.Value ?? 0) / 1000.0;
                    double b = (srcRect.Bottom?.Value ?? 0) / 1000.0;
                    double keptW = 100.0 - l - r;
                    double keptH = 100.0 - t - b;
                    if (keptW > 0 && keptH > 0 && (l != 0 || t != 0 || r != 0 || b != 0))
                    {
                        // Scale so the kept fraction fills 100% of the box; offset
                        // so the cropped-away top/left is pushed out of view.
                        var sizeW = (100.0 / keptW * 100.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        var sizeH = (100.0 / keptH * 100.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        var posX = (keptW <= 0 ? 0 : l / (l + r == 0 ? 1 : (l + r)) * 100.0);
                        var posY = (keptH <= 0 ? 0 : t / (t + b == 0 ? 1 : (t + b)) * 100.0);
                        var px = posX.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        var py = posY.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                        return $"background:transparent url('{dataUri}') {px}% {py}%/{sizeW}% {sizeH}% no-repeat";
                    }
                }
                return $"background:transparent url('{dataUri}') center/cover no-repeat";
            }
        }

        // Pattern fill (a:pattFill) — approximate the preset pattern with a CSS
        // repeating-linear-gradient using the fg color over the bg color. Native
        // Office tiles the real preset bitmap; the gradient gives a recognisable
        // striped/cross texture and, crucially, surfaces the fg color so the
        // shape no longer renders as plain white. Mirrors the colorScale/dxf
        // approach of "approximate in CSS, don't leave it blank".
        var pattFill = spPr.GetFirstChild<Drawing.PatternFill>();
        if (pattFill != null)
            return PatternFillToCss(pattFill, themeColors);

        return "";
    }

    /// <summary>
    /// Convert an a:pattFill to a CSS repeating-linear-gradient approximation.
    /// fg = the pattern's foreground color (the lines), bg = background color.
    /// The preset name picks the gradient angle (diagonal / horizontal /
    /// vertical / grid); unrecognised presets fall back to a diagonal stripe.
    /// </summary>
    private static string PatternFillToCss(Drawing.PatternFill pattFill, Dictionary<string, string> themeColors)
    {
        var fg = ResolveWrappedColor(pattFill.GetFirstChild<Drawing.ForegroundColor>(), themeColors) ?? "#000000";
        var bg = ResolveWrappedColor(pattFill.GetFirstChild<Drawing.BackgroundColor>(), themeColors) ?? "#FFFFFF";
        var preset = pattFill.Preset?.HasValue == true ? pattFill.Preset.InnerText : "diagStripe";

        // Map preset family → gradient angle(s). Diagonal patterns use 45deg,
        // horizontal use 0deg, vertical 90deg, grid/cross layer both.
        var p = (preset ?? "").ToLowerInvariant();
        // Diagonal-cross presets render the × (45deg + 135deg): diagCross plus the
        // checkerboard presets (smCheck/lgCheck). Orthogonal-grid presets render
        // the + (0deg + 90deg): plain cross (without "diag"), grid, weave, trellis.
        bool isDiagCross = p.Contains("diagcross") || p.Contains("check");
        bool isOrthoGrid = (p.Contains("cross") && !p.Contains("diag"))
            || p.Contains("grid") || p.Contains("weave") || p.Contains("trellis");
        bool isGrid = isDiagCross || isOrthoGrid;
        bool isHorz = p.Contains("horz");
        bool isVert = p.Contains("vert");
        // Down-diagonal presets (ltDnDiag/dkDnDiag/wdDnDiag/dashDnDiag) draw `\`
        // (135deg); up-diagonal draws `/` (45deg). "dn" is the down marker.
        bool isDnDiag = p.Contains("dndiag") || (p.Contains("dn") && p.Contains("diag"));
        var diagAngle = isDnDiag ? "135deg" : "45deg";
        var angle = isHorz ? "0deg" : isVert ? "90deg" : diagAngle;
        // Base angle for a grid: orthogonal grids start at 0deg, diagonal crosses
        // at 45deg (down-diagonal grids would start at 135deg if any existed).
        if (isOrthoGrid) angle = "0deg";
        else if (isDiagCross) angle = "45deg";

        // 4px band: 2px fg line over 2px bg gap.
        var stripe = $"repeating-linear-gradient({angle},{fg} 0,{fg} 2px,{bg} 2px,{bg} 4px)";
        if (isGrid)
        {
            // Layer a perpendicular stripe to form a grid; comma-separated
            // backgrounds stack (first on top). Orthogonal grid pairs 0deg+90deg,
            // diagonal cross pairs 45deg+135deg.
            var crossAngle = isOrthoGrid ? "90deg" : "135deg";
            var cross = $"repeating-linear-gradient({crossAngle},{fg} 0,{fg} 2px,transparent 2px,transparent 4px)";
            return $"background:{cross},{stripe}";
        }
        return $"background:{stripe}";
    }

    /// <summary>
    /// Resolve a color from a wrapper element (a:fgClr / a:bgClr) that contains
    /// a srgbClr or schemeClr child — same resolution as a SolidFill body.
    /// </summary>
    private static string? ResolveWrappedColor(OpenXmlCompositeElement? wrapper, Dictionary<string, string> themeColors)
    {
        if (wrapper == null) return null;

        var rgbEl = wrapper.GetFirstChild<Drawing.RgbColorModelHex>();
        var rgb = rgbEl?.Val?.Value;
        if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
            // Apply lumMod/lumOff/tint/shade transforms (same as schemeClr path).
            return ApplyRgbColorTransforms(rgb[..6], rgbEl!);

        var schemeColor = wrapper.GetFirstChild<Drawing.SchemeColor>();
        if (schemeColor?.Val?.HasValue == true)
        {
            var schemeName = schemeColor.Val!.InnerText;
            if (schemeName != null && themeColors.TryGetValue(schemeName, out var themeHex))
                return ApplyColorTransforms(themeHex, schemeColor);
        }
        return null;
    }

    // ==================== CSS Helper: Custom Geometry ====================

    /// <summary>
    /// Convert OOXML CustomGeometry (a:custGeom) path data to CSS clip-path.
    /// Supports moveTo, lineTo, cubicBezTo, quadBezTo, close.
    /// Coordinates are in the path's own coordinate system (w/h),
    /// converted to percentages for clip-path.
    /// </summary>
    private static string CustomGeometryToClipPath(Drawing.CustomGeometry custGeom)
    {
        var pathList = custGeom.GetFirstChild<Drawing.PathList>();
        if (pathList == null) return "";

        var path = pathList.GetFirstChild<Drawing.Path>();
        if (path == null) return "";

        // Path coordinate system
        var pathW = path.Width?.HasValue == true ? path.Width.Value : 100000L;
        var pathH = path.Height?.HasValue == true ? path.Height.Value : 100000L;
        if (pathW == 0) pathW = 100000;
        if (pathH == 0) pathH = 100000;

        // Helper: parse Drawing.Point X/Y (StringValue) to double percentage
        static bool TryParsePoint(Drawing.Point? pt, double pw, double ph, out double px, out double py)
        {
            px = py = 0;
            if (pt?.X?.HasValue != true || pt?.Y?.HasValue != true) return false;
            if (!long.TryParse(pt.X.Value, out var xv) || !long.TryParse(pt.Y.Value, out var yv)) return false;
            px = xv * 100.0 / pw;
            py = yv * 100.0 / ph;
            return true;
        }

        // Try polygon first (only moveTo + lineTo + close = all straight lines)
        bool hasOnlyLines = true;
        foreach (var child in path.ChildElements)
        {
            if (child is Drawing.CubicBezierCurveTo or Drawing.QuadraticBezierCurveTo or Drawing.ArcTo)
            {
                hasOnlyLines = false;
                break;
            }
        }

        if (hasOnlyLines)
        {
            // Use clip-path: polygon() — better browser support
            var points = new List<string>();
            foreach (var child in path.ChildElements)
            {
                switch (child)
                {
                    case Drawing.MoveTo moveTo:
                        if (TryParsePoint(moveTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var mx, out var my))
                            points.Add($"{mx:0.##}% {my:0.##}%");
                        break;
                    case Drawing.LineTo lineTo:
                        if (TryParsePoint(lineTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var lx, out var ly))
                            points.Add($"{lx:0.##}% {ly:0.##}%");
                        break;
                    case Drawing.CloseShapePath:
                        break; // polygon implicitly closes
                }
            }
            if (points.Count >= 3)
                return $"clip-path:polygon({string.Join(",", points)})";
        }
        else
        {
            // Has curves — approximate with polygon() by sampling bezier curves
            // clip-path:path() uses pixel coordinates (not percentages), so we must
            // flatten curves into polygon points with percentage coordinates instead.
            var polyPoints = new List<string>();
            double curX = 0, curY = 0;
            const int bezierSegments = 8; // number of line segments per bezier curve

            foreach (var child in path.ChildElements)
            {
                switch (child)
                {
                    case Drawing.MoveTo moveTo:
                        if (TryParsePoint(moveTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var mx, out var my))
                        {
                            polyPoints.Add($"{mx:0.##}% {my:0.##}%");
                            curX = mx; curY = my;
                        }
                        break;
                    case Drawing.LineTo lineTo:
                        if (TryParsePoint(lineTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var lx, out var ly))
                        {
                            polyPoints.Add($"{lx:0.##}% {ly:0.##}%");
                            curX = lx; curY = ly;
                        }
                        break;
                    case Drawing.CubicBezierCurveTo cubicBez:
                    {
                        var pts = cubicBez.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 3
                            && TryParsePoint(pts[0], pathW, pathH, out var c1x, out var c1y)
                            && TryParsePoint(pts[1], pathW, pathH, out var c2x, out var c2y)
                            && TryParsePoint(pts[2], pathW, pathH, out var c3x, out var c3y))
                        {
                            // Sample cubic bezier: B(t) = (1-t)^3*P0 + 3(1-t)^2*t*P1 + 3(1-t)*t^2*P2 + t^3*P3
                            for (int i = 1; i <= bezierSegments; i++)
                            {
                                double t = i / (double)bezierSegments;
                                double u = 1 - t;
                                double px = u * u * u * curX + 3 * u * u * t * c1x + 3 * u * t * t * c2x + t * t * t * c3x;
                                double py = u * u * u * curY + 3 * u * u * t * c1y + 3 * u * t * t * c2y + t * t * t * c3y;
                                polyPoints.Add($"{px:0.##}% {py:0.##}%");
                            }
                            curX = c3x; curY = c3y;
                        }
                        break;
                    }
                    case Drawing.QuadraticBezierCurveTo quadBez:
                    {
                        var pts = quadBez.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 2
                            && TryParsePoint(pts[0], pathW, pathH, out var q1x, out var q1y)
                            && TryParsePoint(pts[1], pathW, pathH, out var q2x, out var q2y))
                        {
                            // Sample quadratic bezier: B(t) = (1-t)^2*P0 + 2(1-t)*t*P1 + t^2*P2
                            for (int i = 1; i <= bezierSegments; i++)
                            {
                                double t = i / (double)bezierSegments;
                                double u = 1 - t;
                                double px = u * u * curX + 2 * u * t * q1x + t * t * q2x;
                                double py = u * u * curY + 2 * u * t * q1y + t * t * q2y;
                                polyPoints.Add($"{px:0.##}% {py:0.##}%");
                            }
                            curX = q2x; curY = q2y;
                        }
                        break;
                    }
                    case Drawing.ArcTo arc:
                    {
                        // OOXML arcTo: an elliptical arc relative to the current point.
                        // wR/hR are the ellipse radii (path units); stAng/swAng are in
                        // 60000ths of a degree (clockwise from 3-o'clock). The arc STARTS
                        // at the current point (which lies on the ellipse at angle stAng),
                        // so the ellipse center is back-solved from the current point.
                        double wr = ParseArcRadius(arc.WidthRadius?.Value, pathW);
                        double hr = ParseArcRadius(arc.HeightRadius?.Value, pathH);
                        double stAng = Angle60kToDegrees(arc.StartAngle?.Value);
                        double swAng = Angle60kToDegrees(arc.SwingAngle?.Value);
                        // Center in percent space: c = current - (rx·cos st, ry·sin st).
                        double rxPct = wr * 100.0 / pathW;
                        double ryPct = hr * 100.0 / pathH;
                        double ccx = curX - rxPct * Math.Cos(stAng * Math.PI / 180.0);
                        double ccy = curY - ryPct * Math.Sin(stAng * Math.PI / 180.0);
                        int steps = Math.Max(2, (int)Math.Ceiling(Math.Abs(swAng) / 6.0));
                        for (int s = 1; s <= steps; s++)
                        {
                            double a = (stAng + swAng * s / steps) * Math.PI / 180.0;
                            double px = ccx + rxPct * Math.Cos(a);
                            double py = ccy + ryPct * Math.Sin(a);
                            polyPoints.Add($"{px:0.##}% {py:0.##}%");
                            curX = px; curY = py;
                        }
                        break;
                    }
                    case Drawing.CloseShapePath:
                        break; // polygon implicitly closes
                }
            }
            if (polyPoints.Count >= 3)
                return $"clip-path:polygon({string.Join(",", polyPoints)})";
        }

        return "";
    }

    /// <summary>
    /// R48: convert a multi-subpath / fill="none" custGeom to inline SVG path data.
    /// CSS clip-path:polygon() can only express ONE polygon, so a custGeom whose
    /// pathLst holds more than one &lt;a:path&gt; (or a path with fill="none") cannot
    /// round-trip through the clip-path branch. Build SVG path `d` segments in the
    /// path's own w×h coordinate space (used as the SVG viewBox), so the renderer can
    /// emit a scaled &lt;svg&gt; with a fill path (all fillable subpaths, evenodd) and a
    /// separate stroke path (all stroked subpaths). Returns false when there is no
    /// usable path data. <paramref name="fillD"/>/<paramref name="strokeD"/> may be
    /// empty individually (e.g. a fill="none" path has no fill segment).
    /// </summary>
    private static bool CustomGeometryToSvgPaths(
        Drawing.CustomGeometry custGeom,
        out string fillD, out string strokeD, out long pathW, out long pathH)
    {
        fillD = strokeD = "";
        pathW = pathH = 100000L;

        var pathList = custGeom.GetFirstChild<Drawing.PathList>();
        if (pathList == null) return false;

        var paths = pathList.Elements<Drawing.Path>().ToList();
        if (paths.Count == 0) return false;

        // viewBox uses the FIRST path's coordinate space (OOXML paths in one custGeom
        // share the shape's coordinate system; w/h are the same in practice).
        var first = paths[0];
        pathW = first.Width?.HasValue == true && first.Width.Value != 0 ? first.Width.Value : 100000L;
        pathH = first.Height?.HasValue == true && first.Height.Value != 0 ? first.Height.Value : 100000L;

        var fillSegs = new System.Text.StringBuilder();
        var strokeSegs = new System.Text.StringBuilder();

        foreach (var path in paths)
        {
            var pw = path.Width?.HasValue == true && path.Width.Value != 0 ? path.Width.Value : pathW;
            var ph = path.Height?.HasValue == true && path.Height.Value != 0 ? path.Height.Value : pathH;
            var seg = PathToSvgD(path, pw, ph, pathW, pathH);
            if (string.IsNullOrEmpty(seg)) continue;

            // fill="none" → excluded from fill geometry. Stroke="0"/false → excluded from stroke.
            var fillNone = path.Fill?.Value == Drawing.PathFillModeValues.None;
            var strokeOff = path.Stroke?.Value == false;
            if (!fillNone) fillSegs.Append(seg).Append(' ');
            if (!strokeOff) strokeSegs.Append(seg).Append(' ');
        }

        fillD = fillSegs.ToString().Trim();
        strokeD = strokeSegs.ToString().Trim();
        return fillD.Length > 0 || strokeD.Length > 0;
    }

    /// <summary>Convert a single &lt;a:path&gt;'s commands to an SVG path `d` segment in
    /// the viewBox (vbW×vbH) coordinate space. Curves are emitted exactly (C/Q);
    /// arcTo is flattened to line segments (same math as CustomGeometryToClipPath).</summary>
    private static string PathToSvgD(Drawing.Path path, long pw, long ph, long vbW, long vbH)
    {
        // Map path-unit coords into the viewBox space (handles paths declaring a
        // different w/h than the viewBox).
        static bool TryPt(Drawing.Point? pt, long pw, long ph, long vbW, long vbH, out double x, out double y)
        {
            x = y = 0;
            if (pt?.X?.HasValue != true || pt?.Y?.HasValue != true) return false;
            if (!long.TryParse(pt.X.Value, out var xv) || !long.TryParse(pt.Y.Value, out var yv)) return false;
            x = xv * (double)vbW / pw;
            y = yv * (double)vbH / ph;
            return true;
        }

        var d = new System.Text.StringBuilder();
        double curX = 0, curY = 0;
        foreach (var child in path.ChildElements)
        {
            switch (child)
            {
                case Drawing.MoveTo moveTo:
                    if (TryPt(moveTo.GetFirstChild<Drawing.Point>(), pw, ph, vbW, vbH, out var mx, out var my))
                    {
                        d.Append($"M{mx:0.##} {my:0.##} ");
                        curX = mx; curY = my;
                    }
                    break;
                case Drawing.LineTo lineTo:
                    if (TryPt(lineTo.GetFirstChild<Drawing.Point>(), pw, ph, vbW, vbH, out var lx, out var ly))
                    {
                        d.Append($"L{lx:0.##} {ly:0.##} ");
                        curX = lx; curY = ly;
                    }
                    break;
                case Drawing.CubicBezierCurveTo cubic:
                {
                    var pts = cubic.Elements<Drawing.Point>().ToList();
                    if (pts.Count >= 3
                        && TryPt(pts[0], pw, ph, vbW, vbH, out var c1x, out var c1y)
                        && TryPt(pts[1], pw, ph, vbW, vbH, out var c2x, out var c2y)
                        && TryPt(pts[2], pw, ph, vbW, vbH, out var c3x, out var c3y))
                    {
                        d.Append($"C{c1x:0.##} {c1y:0.##} {c2x:0.##} {c2y:0.##} {c3x:0.##} {c3y:0.##} ");
                        curX = c3x; curY = c3y;
                    }
                    break;
                }
                case Drawing.QuadraticBezierCurveTo quad:
                {
                    var pts = quad.Elements<Drawing.Point>().ToList();
                    if (pts.Count >= 2
                        && TryPt(pts[0], pw, ph, vbW, vbH, out var q1x, out var q1y)
                        && TryPt(pts[1], pw, ph, vbW, vbH, out var q2x, out var q2y))
                    {
                        d.Append($"Q{q1x:0.##} {q1y:0.##} {q2x:0.##} {q2y:0.##} ");
                        curX = q2x; curY = q2y;
                    }
                    break;
                }
                case Drawing.ArcTo arc:
                {
                    // Flatten the elliptical arc into line segments (same back-solve as
                    // CustomGeometryToClipPath, but in viewBox units instead of percent).
                    double wr = ParseArcRadius(arc.WidthRadius?.Value, pw) * (double)vbW / pw;
                    double hr = ParseArcRadius(arc.HeightRadius?.Value, ph) * (double)vbH / ph;
                    double stAng = Angle60kToDegrees(arc.StartAngle?.Value);
                    double swAng = Angle60kToDegrees(arc.SwingAngle?.Value);
                    double ccx = curX - wr * Math.Cos(stAng * Math.PI / 180.0);
                    double ccy = curY - hr * Math.Sin(stAng * Math.PI / 180.0);
                    int steps = Math.Max(2, (int)Math.Ceiling(Math.Abs(swAng) / 6.0));
                    for (int s = 1; s <= steps; s++)
                    {
                        double a = (stAng + swAng * s / steps) * Math.PI / 180.0;
                        double px = ccx + wr * Math.Cos(a);
                        double py = ccy + hr * Math.Sin(a);
                        d.Append($"L{px:0.##} {py:0.##} ");
                        curX = px; curY = py;
                    }
                    break;
                }
                case Drawing.CloseShapePath:
                    d.Append("Z ");
                    break;
            }
        }
        return d.ToString().Trim();
    }

    /// <summary>Parse an arcTo radius (path-unit StringValue). Falls back to a quarter
    /// of the path dimension when missing/invalid so a malformed arc still curves.</summary>
    private static double ParseArcRadius(string? raw, double fallbackDim)
        => long.TryParse(raw, out var v) ? v : fallbackDim / 4.0;

    /// <summary>Parse an OOXML angle (60000ths of a degree) to degrees. (Named to
    /// avoid colliding with Model3D's degrees→60000ths ParseAngle60k overload.)</summary>
    private static double Angle60kToDegrees(string? raw)
        => long.TryParse(raw, out var v) ? v / 60000.0 : 0;

    // ==================== CSS Helper: Gradient ====================

    /// <summary>
    /// R15-2: Build an SVG &lt;linearGradient&gt; element from an OOXML &lt;a:gradFill&gt;
    /// for use as a connector/line stroke (stroke="url(#id)"). Reuses the same stop
    /// color resolution as GradientToCss. <paramref name="firstStopColor"/> returns the
    /// first stop's color as a solid fallback. Returns "" when there are fewer than 2 stops.
    /// </summary>
    private static string BuildSvgLinearGradient(Drawing.GradientFill gradFill, string id,
        Dictionary<string, string> themeColors, out string? firstStopColor)
    {
        firstStopColor = null;
        var stops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stops == null || stops.Count < 2) return "";

        // Map the OOXML linear gradient angle to SVG x1/y1/x2/y2 in objectBoundingBox
        // space (0..1). OOXML <a:lin ang> convention: 0° = left→right, 90° = top→bottom
        // (clockwise). The unit direction vector is (cos θ, sin θ) with θ in OOXML
        // degrees, so 0°→(1,0) horizontal and 90°→(0,1) vertical top→bottom. Anchor
        // the gradient line at the center and project ±0.5 along that vector so the
        // line spans the box. Default 90° (top→bottom) matches the fill SVG default.
        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        var angleDeg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 90.0;
        var rad = angleDeg * Math.PI / 180.0;
        var dx = Math.Cos(rad) / 2.0;
        var dy = Math.Sin(rad) / 2.0;
        var x1 = 0.5 - dx; var y1 = 0.5 - dy;
        var x2 = 0.5 + dx; var y2 = 0.5 + dy;

        var sb = new System.Text.StringBuilder();
        sb.Append($"<linearGradient id=\"{id}\" x1=\"{x1:0.####}\" y1=\"{y1:0.####}\" x2=\"{x2:0.####}\" y2=\"{y2:0.####}\">");
        foreach (var gs in stops)
        {
            var color = ResolveGradientStopColor(gs, themeColors);
            firstStopColor ??= color;
            var pos = (gs.Position?.Value ?? 0) / 1000.0;
            sb.Append($"<stop offset=\"{pos:0.##}%\" stop-color=\"{CssSanitizeColor(color)}\"/>");
        }
        sb.Append("</linearGradient>");
        return sb.ToString();
    }

    /// <summary>Resolve a single gradient stop's color (shared by GradientToCss / BuildSvgLinearGradient).</summary>
    private static string ResolveGradientStopColor(Drawing.GradientStop gs, Dictionary<string, string> themeColors)
    {
        var color = ResolveFillColor(gs.GetFirstChild<Drawing.SolidFill>(), themeColors);
        if (color != null) return color;
        var rgbEl = gs.GetFirstChild<Drawing.RgbColorModelHex>();
        var rgb = rgbEl?.Val?.Value;
        if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
        {
            var alpha = rgbEl!.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (alpha.HasValue && alpha.Value < 100000)
            {
                var (r, g, b) = ColorMath.HexToRgb(rgb[..6]);
                return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";
            }
            return $"#{rgb[..6]}";
        }
        var schemeEl = gs.GetFirstChild<Drawing.SchemeColor>();
        var scheme = schemeEl?.Val?.InnerText;
        if (scheme != null && themeColors.TryGetValue(scheme, out var tc))
            return ApplyColorTransforms(tc, schemeEl!);
        // Preset/named color stop (<a:gs><a:prstClr val="red"/>): mirror ResolveFillColor's
        // prstClr support so a named-color gradient stop isn't dropped to transparent.
        var prstEl = gs.GetFirstChild<Drawing.PresetColor>();
        if (prstEl?.Val?.HasValue == true)
        {
            var h = ParseHelpers.TryGetNamedColorHex(prstEl.Val!.InnerText);
            if (h != null) return ApplyColorTransforms(h, prstEl);
        }
        var sysEl = gs.GetFirstChild<Drawing.SystemColor>();
        if (sysEl != null && SysColorHex(sysEl) is string sh)
            return ApplyColorTransforms(sh, sysEl);
        return "transparent";
    }

    private static string GradientToCss(Drawing.GradientFill gradFill, Dictionary<string, string> themeColors)
    {
        var stops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stops == null || stops.Count < 2) return "transparent";

        var cssStops = new List<string>();
        foreach (var gs in stops)
        {
            var color = ResolveFillColor(gs.GetFirstChild<Drawing.SolidFill>(), themeColors);
            if (color == null)
            {
                // Try direct color children. A gradient stop carries its color as a
                // direct <a:srgbClr>/<a:schemeClr> child (not wrapped in solidFill),
                // and that color may have an <a:alpha> child. Bug #7: read the alpha
                // and emit rgba() — the old path dropped it, losing transparency.
                var rgbEl = gs.GetFirstChild<Drawing.RgbColorModelHex>();
                var rgb = rgbEl?.Val?.Value;
                if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
                {
                    var alpha = rgbEl!.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
                    if (alpha.HasValue && alpha.Value < 100000)
                    {
                        var (r, g, b) = ColorMath.HexToRgb(rgb[..6]);
                        color = $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";
                    }
                    else
                        color = $"#{rgb[..6]}";
                }
                else
                {
                    var schemeEl = gs.GetFirstChild<Drawing.SchemeColor>();
                    var scheme = schemeEl?.Val?.InnerText;
                    if (scheme != null && themeColors.TryGetValue(scheme, out var tc))
                    {
                        // Apply lumMod/lumOff/tint/shade/alpha transforms (same as
                        // ResolveFillColor); without this the stops collapse to the base hex.
                        color = ApplyColorTransforms(tc, schemeEl!);
                    }
                    else
                    {
                        // Preset/named color stop (<a:prstClr val="...">) — was dropped to
                        // transparent, blanking the whole gradient fill.
                        var prstEl = gs.GetFirstChild<Drawing.PresetColor>();
                        var ph = prstEl?.Val?.HasValue == true ? ParseHelpers.TryGetNamedColorHex(prstEl.Val!.InnerText) : null;
                        var sysEl = gs.GetFirstChild<Drawing.SystemColor>();
                        color = ph != null ? ApplyColorTransforms(ph, prstEl!)
                            : sysEl != null && SysColorHex(sysEl) is string sh ? ApplyColorTransforms(sh, sysEl)
                            : "transparent";
                    }
                }
            }
            var pos = gs.Position?.Value;
            if (pos.HasValue)
                cssStops.Add($"{color} {pos.Value / 1000.0:0.##}%");
            else
                cssStops.Add(color);
        }

        // Radial or linear?
        var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
        if (pathGrad != null)
        {
            // OOXML <a:path path="circle"> with default fill rectangle fills to the shape
            // bounds (last stop at the edge). CSS default is `farthest-corner`, which overshoots
            // for square-ish shapes. `closest-side` lands the final stop at the nearer edge,
            // matching Office's rendering for rectangular shapes.
            // Bug #6: the gradient FOCUS comes from <a:fillToRect l/t/r/b> (1/1000%
            // units). The focus point is the rect's top-left (l, t): center =
            // 50/50/50/50 → at 50% 50%; tl = l0/t0 → at 0% 0%; br = l100000/t100000
            // → at 100% 100%; tr = l100000/t0 → at 100% 0%. Previously omitted, so
            // every focal variant rendered centered.
            var ftr = pathGrad.FillToRectangle;
            var cx = (ftr?.Left?.Value ?? 50000) / 1000.0;
            var cy = (ftr?.Top?.Value ?? 50000) / 1000.0;
            // Size keyword must track the focus: `closest-side` is only right for a
            // CENTERED focus (the nearer-edge radius that matches Office on a
            // rectangle). At a corner/edge focus, closest-side's nearest-edge
            // distance collapses to ~0 → the gradient degenerates to a point and
            // the focus colors vanish (only the final stop's color shows). Use
            // `farthest-corner` for off-center foci so the gradient fills the shape
            // out to its far corner, matching native's large color fill.
            var centered = Math.Abs(cx - 50) < 0.5 && Math.Abs(cy - 50) < 0.5;
            var sizeKeyword = centered ? "closest-side" : "farthest-corner";
            // <a:path path="rect"> draws rectangular isocontours that reach all four
            // edges of the shape simultaneously. CSS `ellipse` adapts its two radii to
            // the element's aspect ratio, matching that on a non-square shape; `circle`
            // keeps equal radii and only reaches the short edges. `circle`/`shape` paths
            // stay on the CSS `circle` keyword (exact match / best available).
            var shapeKeyword = pathGrad.Path?.Value == Drawing.PathShadeValues.Rectangle ? "ellipse" : "circle";
            return $"radial-gradient({shapeKeyword} {sizeKeyword} at {cx:0.##}% {cy:0.##}%, {string.Join(", ", cssStops)})";
        }

        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        var angleDeg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 90.0;
        // OOXML angle 0° = top→bottom (same as CSS 180deg), so CSS angle = OOXML + 90°
        // Actually OOXML: 0 = right, 90 = bottom; CSS: 0 = up, 90 = right
        // OOXML ang is clockwise from +x (east): 0=→, 90(5400000)=↓, measured in
        // 60000ths of a degree. CSS linear-gradient(Ndeg) is clockwise from "to top"
        // (0deg=up, 90deg=right, 180deg=down). So cssAngle = ooxmlAngle + 90:
        //   ooxml 0°(east) → css 90deg(right) ✓
        //   ooxml 90°(south) → css 180deg(down/to bottom) ✓
        //   ooxml 45°(SE) → css 135deg(to bottom right) ✓
        var cssAngle = angleDeg + 90;

        // scaled="1": the OOXML angle is measured in the shape's bounding-box
        // diagonal space, NOT in absolute screen degrees. Dual-render vs real
        // PowerPoint (R62):
        //   • CARDINAL angles (0°/90°/180°/270° → horizontal/vertical) stay
        //     axis-aligned regardless of aspect ratio — a 90° scaled gradient on
        //     ANY box is a clean vertical (red top → blue bottom). The old code
        //     snapped these to a CSS corner keyword, turning vertical into a
        //     diagonal. WRONG. They must emit the precise CSS angle.
        //   • TRUE DIAGONAL angles (45°/135°/... ) on a NON-SQUARE box run
        //     corner-to-corner of the actual box (full saturation lands exactly
        //     on the box corners), not along a fixed 135° screen line. CSS corner
        //     keywords (`to bottom right`, …) are aspect-aware by spec and match
        //     PowerPoint here; a fixed 135deg would miss the corners on a wide box.
        // So: corner-snap ONLY genuine diagonals; emit the precise angle for
        // cardinals (and for all unscaled gradients).
        var aNorm = ((angleDeg % 360) + 360) % 360;
        var isCardinal = Math.Abs(aNorm % 90) < 0.5 || Math.Abs(aNorm % 90 - 90) < 0.5;
        if (linear?.Scaled?.Value == true && !isCardinal)
        {
            var a = ((cssAngle % 360) + 360) % 360;
            var corner = a switch
            {
                >= 0 and < 90 => "to top right",
                >= 90 and < 180 => "to bottom right",
                >= 180 and < 270 => "to bottom left",
                _ => "to top left",
            };
            return $"linear-gradient({corner}, {string.Join(", ", cssStops)})";
        }

        return $"linear-gradient({cssAngle:0.##}deg, {string.Join(", ", cssStops)})";
    }

    // ==================== CSS Helper: Outline/Border ====================

    /// <summary>
    /// Parse outline into (widthPt, ooxmlDashType, color). Returns null if NoFill.
    /// </summary>
    private static (double widthPt, string dashType, string color, string cap, string cmpd, string join)? ParseOutline(Drawing.Outline outline, Dictionary<string, string> themeColors)
    {
        if (outline.GetFirstChild<Drawing.NoFill>() != null) return null;

        // Empty <a:ln/> (no fill child, no width) means "inherit/default" — for text
        // shapes PowerPoint treats this as no line. Without this guard we fall through
        // to dk1 default + 0.5pt and paint a phantom border on every plain text box.
        if (outline.GetFirstChild<Drawing.SolidFill>() == null
            && outline.GetFirstChild<Drawing.GradientFill>() == null
            && outline.Width?.HasValue != true)
            return null;

        var color = ResolveFillColor(outline.GetFirstChild<Drawing.SolidFill>(), themeColors)
            ?? (themeColors.TryGetValue("dk1", out var dk1Hex) ? $"#{dk1Hex}" : "#000000");
        var widthPt = outline.Width?.HasValue == true ? outline.Width.Value / EmuConverter.EmuPerPointF : 1.0;
        if (widthPt < 0.5) widthPt = 0.5;

        var dash = outline.GetFirstChild<Drawing.PresetDash>();
        var dashType = "solid";
        if (dash?.Val?.HasValue == true)
            dashType = dash.Val.InnerText ?? "solid";
        else
        {
            // <a:custDash> (mutually exclusive with prstDash): a list of <a:ds d sp/>
            // stops where d/sp are ST_PositivePercentage (100000 = 100% of line width).
            // Encode as "custom:<onMult>,<offMult>,..." (multiples of line width) so the
            // SVG dash overlay can scale them by stroke width like the preset dasharrays.
            var cust = outline.GetFirstChild<Drawing.CustomDash>();
            if (cust != null)
            {
                var ci = System.Globalization.CultureInfo.InvariantCulture;
                var segs = new System.Collections.Generic.List<string>();
                foreach (var ds in cust.Elements<Drawing.DashStop>())
                {
                    double d = (ds.DashLength?.Value ?? 0) / 100000.0;
                    double sp = (ds.SpaceLength?.Value ?? 0) / 100000.0;
                    if (d <= 0 && sp <= 0) continue;
                    segs.Add(d.ToString("0.##", ci));
                    segs.Add(sp.ToString("0.##", ci));
                }
                if (segs.Count > 0) dashType = "custom:" + string.Join(",", segs);
            }
        }

        // Line cap (<a:ln cap="rnd|sq|flat"/>) and compound type
        // (<a:ln cmpd="sng|dbl|thickThin|thinThick|tri"/>).
        var cap = outline.CapType?.HasValue == true ? (outline.CapType.InnerText ?? "flat") : "flat";
        var cmpd = outline.CompoundLineType?.HasValue == true ? (outline.CompoundLineType.InnerText ?? "sng") : "sng";

        // Line join (<a:ln> child <a:round/> | <a:bevel/> | <a:miter/>) → SVG
        // stroke-linejoin. PowerPoint's DEFAULT outline join is ROUND (verified: a
        // thick-outlined shape with no explicit join renders rounded corners), whereas
        // SVG's default is "miter" (sharp). So default to "round" here; the SVG stroke
        // sites that omit stroke-linejoin were rendering sharp corners where PowerPoint
        // rounds them.
        var join = outline.GetFirstChild<Drawing.LineJoinBevel>() != null ? "bevel"
            : outline.GetFirstChild<Drawing.Miter>() != null ? "miter"
            : "round"; // explicit <a:round/> or absent → round (PowerPoint default)

        return (widthPt, dashType, color, cap, cmpd, join);
    }

    /// <summary>
    /// Map OOXML line cap (<a:ln cap="rnd|sq|flat"/>) to the SVG stroke-linecap value.
    /// rnd→round (pill dash ends), sq→square, flat/default→butt.
    /// </summary>
    private static string CapToSvgLinecap(string cap) => cap switch
    {
        "rnd" => "round",
        "sq" => "square",
        _ => "butt",
    };

    // ==================== Style-matrix (p:style) resolution ====================
    //
    // A shape may carry a <p:style> with fillRef/lnRef/effectRef/fontRef that index
    // into the theme's <a:fmtScheme> (FormatScheme). When the shape's own spPr has no
    // explicit fill/outline/effect (or the run has no explicit color), these refs
    // supply the value. Explicit spPr/run values always WIN; we only resolve a ref
    // when the explicit value is absent.
    //
    // The FormatScheme is read fresh from the part chain each render (not cached) so
    // SDK-injected effectStyleLst entries are seen after reopen.

    /// <summary>
    /// Resolve the theme FormatScheme for a shape's part:
    /// SlidePart→SlideLayoutPart→SlideMasterPart→ThemePart→Theme→ThemeElements→FormatScheme.
    /// Read directly from the part each call (no caching) so post-reopen theme edits are honored.
    /// </summary>
    private static Drawing.FormatScheme? ResolveFormatScheme(OpenXmlPart part)
    {
        var theme = part switch
        {
            SlidePart sp => sp.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme,
            SlideLayoutPart lp => lp.SlideMasterPart?.ThemePart?.Theme,
            SlideMasterPart mp => mp.ThemePart?.Theme,
            _ => null,
        };
        return theme?.ThemeElements?.FormatScheme;
    }

    /// <summary>
    /// Resolve a style-matrix reference's <a:schemeClr> (the ref's own color slot,
    /// e.g. fillRef/lnRef/fontRef child) to a "#RRGGBB"/rgba() string via the theme
    /// color map, applying any lumMod/lumOff/tint/shade/alpha transforms.
    /// </summary>
    private static string? ResolveStyleRefSchemeColor(OpenXmlCompositeElement? styleRef, Dictionary<string, string> themeColors)
    {
        var schemeColor = styleRef?.GetFirstChild<Drawing.SchemeColor>();
        if (schemeColor?.Val?.HasValue != true) return null;
        var name = schemeColor.Val!.InnerText;
        if (name == null || !themeColors.TryGetValue(name, out var hex)) return null;
        return ApplyColorTransforms(hex, schemeColor);
    }

    /// <summary>
    /// Resolve a style-matrix ref's schemeClr to a bare 6-hex (no '#'), applying the
    /// ref's own transforms. Used as the "phClr" (placeholder color) base when an
    /// indexed theme fill (gradient) uses phClr stops — the gradient's own stop
    /// transforms (lumMod/tint/satMod) then layer on top of this value.
    /// </summary>
    private static string? ResolveStyleRefBareHex(OpenXmlCompositeElement? styleRef, Dictionary<string, string> themeColors)
    {
        var sc = styleRef?.GetFirstChild<Drawing.SchemeColor>();
        if (sc?.Val?.HasValue != true) return null;
        var name = sc.Val!.InnerText;
        if (name == null || !themeColors.TryGetValue(name, out var hex)) return null;
        var css = ApplyColorTransforms(hex, sc);
        // ApplyColorTransforms normally yields "#RRGGBB"; reduce to bare hex so the
        // gradient stop resolver (which expects a bare theme hex) can re-transform it.
        if (css.StartsWith("#") && css.Length >= 7) return css.Substring(1, 6);
        return hex; // rgba()/unexpected — fall back to the untransformed base hex
    }

    /// <summary>
    /// Style-matrix fill fallback: resolve <p:style>/<a:fillRef idx=N> against
    /// FormatScheme.FillStyleList[N] (1-based), blending the fillRef's schemeClr into
    /// the indexed fill style. Returns a "background:..." CSS string, or "" if none.
    /// </summary>
    private static string GetStyleFillRefCss(ShapeStyle? style, OpenXmlPart part, Dictionary<string, string> themeColors)
    {
        var fillRef = style?.FillReference;
        var idx = fillRef?.Index?.Value ?? 0;
        if (fillRef == null || idx == 0) return "";

        var refColor = ResolveStyleRefSchemeColor(fillRef, themeColors);

        var fmtScheme = ResolveFormatScheme(part);
        var fillStyle = fmtScheme?.FillStyleList?.ChildElements
            .OfType<OpenXmlElement>()
            .ElementAtOrDefault((int)idx - 1);

        // The indexed fill style entry is most commonly a <a:solidFill> referencing
        // <a:schemeClr val="phClr"/> (the placeholder color = the fillRef's schemeClr).
        // Emit the fillRef schemeClr directly — it is the resolved phClr.
        if (fillStyle is Drawing.SolidFill && refColor != null)
            return $"background:{refColor}";

        // Gradient indexed fill: the theme gradient's stops reference phClr (the
        // placeholder color = the fillRef's own schemeClr). themeColors has no
        // "phClr" key, so without substitution every stop resolves to "transparent"
        // and the shape renders invisible (PowerPoint shows the accent gradient).
        // Inject phClr = the resolved fillRef color so the stops' own transforms layer.
        if (fillStyle is Drawing.GradientFill gf)
        {
            var phHex = ResolveStyleRefBareHex(fillRef, themeColors);
            var patched = phHex != null
                ? new Dictionary<string, string>(themeColors) { ["phClr"] = phHex }
                : themeColors;
            var css = GradientToCss(gf, patched);
            return string.IsNullOrEmpty(css) || css == "transparent"
                ? (refColor != null ? $"background:{refColor}" : "")
                : $"background:{css}";
        }

        // Unknown/no indexed style entry — fall back to the bare ref color.
        return refColor != null ? $"background:{refColor}" : "";
    }

    /// <summary>
    /// Style-matrix outline fallback: resolve <p:style>/<a:lnRef idx=N> against
    /// FormatScheme.LineStyleList[N] (1-based), coloring it with the lnRef's schemeClr.
    /// Returns a "border:..." CSS string, or "" if none.
    /// </summary>
    private static string GetStyleLineRefCss(ShapeStyle? style, OpenXmlPart part, Dictionary<string, string> themeColors)
    {
        var lnRef = style?.LineReference;
        var idx = lnRef?.Index?.Value ?? 0;
        if (lnRef == null || idx == 0) return "";

        var refColor = ResolveStyleRefSchemeColor(lnRef, themeColors)
            ?? (themeColors.TryGetValue("dk1", out var dk1) ? $"#{dk1}" : "#000000");

        var fmtScheme = ResolveFormatScheme(part);
        var lineStyle = fmtScheme?.LineStyleList?.ChildElements
            .OfType<Drawing.Outline>()
            .ElementAtOrDefault((int)idx - 1);

        var widthPt = lineStyle?.Width?.HasValue == true ? lineStyle.Width!.Value / EmuConverter.EmuPerPointF : 1.0;
        if (widthPt < 0.5) widthPt = 0.5;

        return $"border:{widthPt:0.##}pt solid {refColor}";
    }

    /// <summary>
    /// Style-matrix stroke fallback for SVG connectors/lines: resolve
    /// <p:style>/<a:lnRef idx=N> to a (color, widthPt) stroke. This parallels
    /// GetStyleLineRefCss (which returns a CSS "border:" for box shapes) but yields
    /// the raw stroke color + width an SVG <line>/<polyline> needs. Returns null when
    /// there is no usable lnRef. A default PowerPoint connector carries NO explicit
    /// <a:ln>; its color/weight come entirely from this lnRef against the theme
    /// FormatScheme.LineStyleList, so without this the line renders theme-tx1 black.
    /// </summary>
    private static (string color, double widthPt)? ResolveStyleLineRefStroke(
        ShapeStyle? style, OpenXmlPart? part, Dictionary<string, string> themeColors)
    {
        var lnRef = style?.LineReference;
        var idx = lnRef?.Index?.Value ?? 0;
        if (lnRef == null || idx == 0 || part == null) return null;

        var refColor = ResolveStyleRefSchemeColor(lnRef, themeColors)
            ?? (themeColors.TryGetValue("dk1", out var dk1) ? $"#{dk1}" : "#000000");

        var fmtScheme = ResolveFormatScheme(part);
        var lineStyle = fmtScheme?.LineStyleList?.ChildElements
            .OfType<Drawing.Outline>()
            .ElementAtOrDefault((int)idx - 1);
        var widthPt = lineStyle?.Width?.HasValue == true
            ? lineStyle.Width!.Value / EmuConverter.EmuPerPointF : 1.0;
        return (refColor, widthPt);
    }

    /// <summary>
    /// Style-matrix effect fallback: resolve <p:style>/<a:effectRef idx=N> against
    /// FormatScheme.EffectStyleList[N] (1-based), reusing EffectListToShadowCss on the
    /// indexed effect style's <a:effectLst>. Returns the shadow CSS, or "" if none.
    /// </summary>
    private static string GetStyleEffectRefCss(ShapeStyle? style, OpenXmlPart part, Dictionary<string, string> themeColors)
        => EffectListToShadowCss(ResolveStyleEffectRefList(style, part), themeColors);

    /// <summary>
    /// Resolve the <a:effectLst> from a shape's <p:style>/<a:effectRef idx=N>
    /// against the theme FormatScheme.EffectStyleList[N] (1-based). Returns null
    /// when there is no effectRef or no matching style entry. Callers can pass the
    /// result to the shadow/glow/reflection converters so effectRef-only shapes
    /// surface all three effect kinds (R14-1), not just shadow.
    /// </summary>
    private static Drawing.EffectList? ResolveStyleEffectRefList(ShapeStyle? style, OpenXmlPart part)
    {
        var effectRef = style?.EffectReference;
        var idx = effectRef?.Index?.Value ?? 0;
        if (effectRef == null || idx == 0) return null;

        var fmtScheme = ResolveFormatScheme(part);
        var effectStyle = fmtScheme?.EffectStyleList?.ChildElements
            .OfType<Drawing.EffectStyle>()
            .ElementAtOrDefault((int)idx - 1);
        return effectStyle?.GetFirstChild<Drawing.EffectList>();
    }

    private static string OutlineToCss(Drawing.Outline outline, Dictionary<string, string> themeColors)
    {
        var parsed = ParseOutline(outline, themeColors);
        if (parsed == null) return "";
        var (widthPt, dashType, color, _, cmpd, _) = parsed.Value;

        var borderStyle = dashType switch
        {
            "dash" or "lgDash" or "sysDash" => "dashed",
            "dot" or "sysDot" => "dotted",
            "dashDot" or "lgDashDot" or "sysDashDot" or "sysDashDotDot" => "dashed",
            _ => "solid"
        };
        if (dashType.StartsWith("custom:", StringComparison.Ordinal))
            borderStyle = "dashed";
        // Compound (dbl/thickThin/thinThick/tri) draws multiple parallel lines.
        // CSS `double` renders two parallel lines when the border is wide enough.
        if (cmpd != "sng" && dashType == "solid")
            borderStyle = "double";

        return $"border:{widthPt:0.##}pt {borderStyle} {color}";
    }

    /// <summary>
    /// Convert OOXML dash type to SVG stroke-dasharray relative to stroke width.
    /// </summary>
    private static string DashTypeToSvgDasharray(string dashType, double strokeWidth)
    {
        var w = strokeWidth;
        // custom:<onMult>,<offMult>,... — a <a:custDash> pattern in multiples of line width.
        if (dashType.StartsWith("custom:", StringComparison.Ordinal))
        {
            var ci = System.Globalization.CultureInfo.InvariantCulture;
            var outp = new System.Collections.Generic.List<string>();
            foreach (var part in dashType.Substring(7).Split(','))
                if (double.TryParse(part, System.Globalization.NumberStyles.Float, ci, out var m))
                    outp.Add((m * w).ToString("0.##", ci));
            return string.Join(" ", outp);
        }
        return dashType switch
        {
            // Dot is a visible short segment (length = stroke width) with linecap=butt
            // so the dot renders as a square of side w. Prior implementation used "0.1"
            // as a zero-length segment relying on stroke-linecap=round to paint a cap;
            // that collapses when linecap=butt or when stroke-width rounds down.
            "solid" => "",
            // dot: on 1w, off 3w (clear separation); sysDot: on 1w, off 1w (tight,
            // dots nearly touching). Measured against real PowerPoint — the two must
            // render distinctly (previously both emitted "w 2w" and looked identical).
            "dot" => $"{w:0.##} {w * 3:0.##}",
            "sysDot" => $"{w:0.##} {w:0.##}",
            "dash" => $"{w * 4:0.##} {w * 3:0.##}",
            "lgDash" => $"{w * 8:0.##} {w * 3:0.##}",
            "sysDash" => $"{w * 3:0.##} {w * 1:0.##}",
            "dashDot" => $"{w * 4:0.##} {w * 2:0.##} {w:0.##} {w * 2:0.##}",
            "lgDashDot" => $"{w * 8:0.##} {w * 2:0.##} {w:0.##} {w * 2:0.##}",
            "sysDashDot" => $"{w * 3:0.##} {w * 1.5:0.##} {w:0.##} {w * 1.5:0.##}",
            "sysDashDotDot" => $"{w * 3:0.##} {w * 1.5:0.##} {w:0.##} {w * 1.5:0.##} {w:0.##} {w * 1.5:0.##}",
            "lgDashDotDot" => $"{w * 8:0.##} {w * 2:0.##} {w:0.##} {w * 2:0.##} {w:0.##} {w * 2:0.##}",
            _ => ""
        };
    }

    // ==================== CSS Helper: Shadow ====================

    /// <summary>a:blur — a Gaussian blur over the ENTIRE shape/picture (distinct from
    /// a:softEdge, which only feathers the edge region). Maps to CSS filter:blur().
    /// Returns just the filter function (e.g. "blur(8pt)") or "" when absent. The
    /// caller merges it into the combined filter: property alongside drop-shadow/glow.</summary>
    private static string EffectListToBlurCss(Drawing.EffectList? effectList)
    {
        var blur = effectList?.GetFirstChild<Drawing.Blur>();
        if (blur?.Radius?.HasValue != true || blur.Radius.Value <= 0) return "";
        // PowerPoint's a:blur rad is a Gaussian blur RADIUS; CSS filter:blur() takes a
        // standard deviation (sigma), and a Gaussian's visible spread is ~2*sigma. Using
        // the radius directly as sigma over-blurs badly (text becomes unreadable);
        // sigma ≈ radius/2 matches PowerPoint's milder render.
        var blurPt = blur.Radius.Value / EmuConverter.EmuPerPointF / 2.0;
        return $"blur({blurPt:0.##}pt)";
    }

    private static string EffectListToShadowCss(Drawing.EffectList? effectList, Dictionary<string, string> themeColors)
    {
        if (effectList == null) return "";

        var shadow = effectList.GetFirstChild<Drawing.OuterShadow>();
        if (shadow != null)
        {
            var color = ResolveShadowColor(shadow, themeColors);
            var blurPt = shadow.BlurRadius?.HasValue == true ? shadow.BlurRadius.Value / EmuConverter.EmuPerPointF : 0;
            var distPt = shadow.Distance?.HasValue == true ? shadow.Distance.Value / EmuConverter.EmuPerPointF : 0;
            var angleDeg = shadow.Direction?.HasValue == true ? shadow.Direction.Value / 60000.0 : 0;
            var angleRad = angleDeg * Math.PI / 180;
            var offsetX = distPt * Math.Cos(angleRad);
            var offsetY = distPt * Math.Sin(angleRad);
            return $"filter:drop-shadow({offsetX:0.##}pt {offsetY:0.##}pt {blurPt:0.##}pt {color})";
        }

        // a:innerShdw has no CSS filter equivalent; render as an inset box-shadow.
        var inner = effectList.GetFirstChild<Drawing.InnerShadow>();
        if (inner != null)
        {
            var color = ResolveShadowColor(inner, themeColors);
            var blurPt = inner.BlurRadius?.HasValue == true ? inner.BlurRadius.Value / EmuConverter.EmuPerPointF : 0;
            var distPt = inner.Distance?.HasValue == true ? inner.Distance.Value / EmuConverter.EmuPerPointF : 0;
            var angleDeg = inner.Direction?.HasValue == true ? inner.Direction.Value / 60000.0 : 0;
            var angleRad = angleDeg * Math.PI / 180;
            var offsetX = distPt * Math.Cos(angleRad);
            var offsetY = distPt * Math.Sin(angleRad);
            return $"box-shadow:inset {offsetX:0.##}pt {offsetY:0.##}pt {blurPt:0.##}pt {color}";
        }

        // R9-1: a:prstShdw (preset shadow). No CSS filter equivalent for the
        // preset bitmap; approximate from its dist/dir/color as a box-shadow.
        var preset = effectList.GetFirstChild<Drawing.PresetShadow>();
        if (preset != null)
        {
            var color = ResolveShadowColor(preset, themeColors);
            var distPt = preset.Distance?.HasValue == true ? preset.Distance.Value / EmuConverter.EmuPerPointF : 0;
            var angleDeg = preset.Direction?.HasValue == true ? preset.Direction.Value / 60000.0 : 0;
            var angleRad = angleDeg * Math.PI / 180;
            var offsetX = distPt * Math.Cos(angleRad);
            var offsetY = distPt * Math.Sin(angleRad);
            // Emit as filter:drop-shadow so it merges with the glow/shadow filter
            // chain in RenderShape (box-shadow can't combine there). Blur ≈ half dist.
            return $"filter:drop-shadow({offsetX:0.##}pt {offsetY:0.##}pt {(distPt / 2):0.##}pt {color})";
        }

        return "";
    }

    /// <summary>
    /// Resolve a shadow's color (rgba) from its srgbClr/schemeClr child, applying
    /// lumMod/lumOff/tint/shade/alpha transforms. When the color carries no
    /// &lt;a:alpha&gt;, the shadow is fully opaque (OOXML default = 100%), matching
    /// how PowerPoint renders an alpha-less shadow (solid, not half-transparent).
    /// </summary>
    private static string ResolveShadowColor(OpenXmlCompositeElement shadow, Dictionary<string, string> themeColors)
    {
        var alpha = shadow.Descendants<Drawing.Alpha>().FirstOrDefault()?.Val?.Value ?? 100000;
        var opacity = alpha / 100000.0;
        var rgbEl = shadow.GetFirstChild<Drawing.RgbColorModelHex>();
        var rgb = rgbEl?.Val?.Value;
        if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
        {
            var transformed = ApplyRgbColorTransforms(rgb[..6], rgbEl!);
            var (r, g, b) = ColorMath.HexToRgb(transformed.StartsWith('#') ? transformed[1..] : transformed);
            return $"rgba({r},{g},{b},{opacity:0.##})";
        }

        var schemeEl = shadow.GetFirstChild<Drawing.SchemeColor>();
        var schemeName = schemeEl?.Val?.InnerText;
        if (schemeName != null && themeColors.TryGetValue(schemeName, out var sc))
        {
            // ApplyTransforms returns #RRGGBB (or rgba when alpha given); strip to hex
            // then re-apply the shadow opacity uniformly. Read by local name so both
            // typed and OpenXmlUnknownElement transform children resolve.
            var transformed = ColorMath.ApplyTransforms(sc,
                tint: ReadTransformVal(schemeEl!, "tint"),
                shade: ReadTransformVal(schemeEl!, "shade"),
                lumMod: ReadTransformVal(schemeEl!, "lumMod"),
                lumOff: ReadTransformVal(schemeEl!, "lumOff"),
                satMod: ReadTransformVal(schemeEl!, "satMod"));
            var hex = transformed.StartsWith('#') ? transformed[1..] : transformed;
            var (r, g, b) = ColorMath.HexToRgb(hex);
            return $"rgba({r},{g},{b},{opacity:0.##})";
        }

        return $"rgba(0,0,0,{opacity:0.##})";
    }

    /// <summary>
    /// Apply lumMod/lumOff/tint/shade transforms (if present as children) to an
    /// srgbClr hex. Returns a #RRGGBB hex (alpha handled separately by callers).
    /// </summary>
    private static string ApplyRgbColorTransforms(string hex, Drawing.RgbColorModelHex rgbEl)
    {
        // Alpha is handled by the caller; pass only lum/tint/shade so the result
        // stays a #RRGGBB hex. Reads by local name to support both typed and
        // OpenXmlUnknownElement children (see ApplyColorTransforms remarks).
        return ColorMath.ApplyTransforms(hex,
            tint: ReadTransformVal(rgbEl, "tint"),
            shade: ReadTransformVal(rgbEl, "shade"),
            lumMod: ReadTransformVal(rgbEl, "lumMod"),
            lumOff: ReadTransformVal(rgbEl, "lumOff"),
            satMod: ReadTransformVal(rgbEl, "satMod"));
    }

    // ==================== CSS Helper: Glow ====================

    private static string EffectListToGlowCss(Drawing.EffectList? effectList, Dictionary<string, string> themeColors)
    {
        if (effectList == null) return "";

        var glow = effectList.GetFirstChild<Drawing.Glow>();
        if (glow == null) return "";

        var alpha = glow.Descendants<Drawing.Alpha>().FirstOrDefault()?.Val?.Value ?? 40000;
        var opacity = alpha / 100000.0;
        var radiusPt = glow.Radius?.HasValue == true ? glow.Radius.Value / EmuConverter.EmuPerPointF : 5;

        // Resolve the glow color WITH its lumMod/lumOff/tint/shade/satMod transforms —
        // every built-in glow preset uses e.g. <a:schemeClr val="accent1"><a:lumMod
        // val="75000"/>, and ignoring the transform renders the raw (too-bright) accent.
        // Mirrors ResolveShadowColor (which already applies them); the glow path didn't.
        var rgbEl = glow.GetFirstChild<Drawing.RgbColorModelHex>();
        var rgb = rgbEl?.Val?.Value;
        (int r, int g, int b)? rgbTuple = null;
        if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
        {
            var t = ApplyRgbColorTransforms(rgb[..6], rgbEl!);
            rgbTuple = ColorMath.HexToRgb(t.StartsWith('#') ? t[1..] : t);
        }
        else
        {
            var schemeEl = glow.GetFirstChild<Drawing.SchemeColor>();
            var schemeName = schemeEl?.Val?.InnerText;
            if (schemeName != null && themeColors.TryGetValue(schemeName, out var sc))
            {
                var t = ColorMath.ApplyTransforms(sc,
                    tint: ReadTransformVal(schemeEl!, "tint"),
                    shade: ReadTransformVal(schemeEl!, "shade"),
                    lumMod: ReadTransformVal(schemeEl!, "lumMod"),
                    lumOff: ReadTransformVal(schemeEl!, "lumOff"),
                    satMod: ReadTransformVal(schemeEl!, "satMod"));
                rgbTuple = ColorMath.HexToRgb(t.StartsWith('#') ? t[1..] : t);
            }
            else
            {
                // No color specified — use theme accent1 or transparent
                var acc1 = themeColors.TryGetValue("accent1", out var a1) ? a1 : null;
                if (acc1 != null)
                    rgbTuple = ColorMath.HexToRgb(acc1);
            }
        }

        if (rgbTuple == null)
            return ""; // no resolvable color — emit nothing rather than an invisible shadow

        var (gr, gg, gb) = rgbTuple.Value;

        // A single low-alpha drop-shadow is barely visible on a white slide.
        // Real PowerPoint paints a dense saturated halo, so stack several
        // drop-shadow layers at progressively wider radii. Each layer composites
        // over the previous, building up the colored halo to a visible density
        // matching native. Per-layer alpha is boosted relative to the OOXML alpha
        // (clamped) since drop-shadow blur disperses the color heavily.
        double layerAlpha = Math.Min(0.9, Math.Max(0.45, opacity + 0.35));
        string col = $"rgba({gr},{gg},{gb},{layerAlpha:0.##})";
        var layers = new[]
        {
            $"drop-shadow(0 0 {radiusPt * 0.4:0.##}pt {col})",
            $"drop-shadow(0 0 {radiusPt:0.##}pt {col})",
            $"drop-shadow(0 0 {radiusPt:0.##}pt {col})",
            $"drop-shadow(0 0 {radiusPt * 1.6:0.##}pt {col})",
        };
        return $"filter:{string.Join(" ", layers)}";
    }

    // ==================== CSS Helper: Reflection ====================

    /// <summary>
    /// Generates CSS -webkit-box-reflect for an OOXML reflection effect.
    /// Uses the reflection's StartOpacity, EndAlpha, EndPosition, Distance, and BlurRadius
    /// to build an appropriate linear-gradient fade.
    /// </summary>
    private static string EffectListToReflectionCss(Drawing.EffectList? effectList)
    {
        if (effectList == null) return "";

        var refl = effectList.GetFirstChild<Drawing.Reflection>();
        if (refl == null) return "";

        // Distance between shape bottom and reflection start (EMU → pt)
        var distPt = refl.Distance?.HasValue == true ? refl.Distance.Value / EmuConverter.EmuPerPointF : 0;

        // StartOpacity: initial opacity of reflected image (thousandths of a percent)
        var startOpacity = refl.StartOpacity?.HasValue == true ? refl.StartOpacity.Value / 100000.0 : 0.52;

        // EndAlpha: final opacity (thousandths of a percent)
        var endOpacity = refl.EndAlpha?.HasValue == true ? refl.EndAlpha.Value / 100000.0 : 0.0;

        // EndPosition: how much of the shape height is reflected (thousandths of a percent → CSS percentage).
        // In -webkit-box-reflect, 0% is the top of the reflection (closest to the source shape) and
        // 100% is the far edge. The reflection should be most opaque at the top (startOpacity) and
        // fade to endOpacity at endPos%, then fully transparent beyond endPos.
        var endPos = refl.EndPosition?.HasValue == true ? Math.Clamp(refl.EndPosition.Value / 1000.0, 0, 100) : 90.0;

        var startStop = $"rgba(255,255,255,{startOpacity:0.###}) 0%";
        var endStop = $"rgba(255,255,255,{endOpacity:0.###}) {endPos:0.#}%";
        var tailStop = endPos < 100 ? $",transparent 100%" : "";

        return $"-webkit-box-reflect:below {distPt:0.##}pt linear-gradient({startStop},{endStop}{tailStop})";
    }

    // ==================== CSS Helper: Preset Geometry ====================

    /// <summary>Plus/cross polygon with arm width proportional to min(w,h).</summary>
    // plus/cross: the single adj is the corner-notch inset as a fraction of each
    // dimension (ECMA x1=w*adj, y1=h*adj), default 25000 (25%). Arm spans
    // [inset, 100-inset] on both axes. Small adj => wide arms / tiny notches;
    // large adj => thin needle-cross. The old code hardcoded 25% (ignored adj).
    private static string PlusPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var adj = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, 50000);
        var inset = adj / 1000.0;
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        var l = P(inset); var r = P(100 - inset); var t = l; var b = r;
        return $"clip-path:polygon({l}% 0,{r}% 0,{r}% {t}%,100% {t}%,100% {b}%,{r}% {b}%,{r}% 100%,{l}% 100%,{l}% {b}%,0 {b}%,0 {t}%,{l}% {t}%)";
    }

    private static string PresetGeometryToCss(string preset) =>
        PresetGeometryToCss(preset, 0, 0, null);

    /// <summary>
    /// Read an adjustment value from PresetGeometry's AdjustValueList (OOXML "val NNNNN" formula).
    /// </summary>
    private static long ReadAdjValueCss(Drawing.PresetGeometry? presetGeom, int index, long defaultValue)
    {
        var avList = presetGeom?.GetFirstChild<Drawing.AdjustValueList>();
        if (avList == null) return defaultValue;
        var guides = avList.Elements<Drawing.ShapeGuide>().ToList();
        if (index >= guides.Count) return defaultValue;
        var formula = guides[index].Formula?.Value;
        if (formula != null && formula.StartsWith("val "))
        {
            if (long.TryParse(formula.AsSpan(4), out var parsed))
                return parsed;
        }
        return defaultValue;
    }

    /// <summary>
    /// Build a clip-path polygon for rightArrow honoring OOXML avLst.
    /// adj1 = tail height relative to shape height (0..100000, default 50000 = 50%)
    /// adj2 = head width relative to min(w,h) (0..100000, default 50000)
    /// </summary>
    private static string RightArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var adj1 = ReadAdjValueCss(presetGeom, 0, 50000);
        var adj2 = ReadAdjValueCss(presetGeom, 1, 50000);
        // Clamp avLst values to sane range
        if (adj1 < 0) adj1 = 0; if (adj1 > 100000) adj1 = 100000;
        if (adj2 < 0) adj2 = 0; if (adj2 > 100000) adj2 = 100000;

        // Tail vertical extent (centered on midline): adj1 fraction of height
        var tailTop = (100000.0 - adj1) / 2000.0;   // e.g. 25%
        var tailBot = 100.0 - tailTop;              // e.g. 75%

        // Head width measured from the right edge. Fallback to square assumption if dims missing.
        double headStartX;
        if (widthEmu > 0 && heightEmu > 0)
        {
            var minSide = Math.Min(widthEmu, heightEmu);
            var headWidthEmu = minSide * adj2 / 100000.0;
            if (headWidthEmu > widthEmu) headWidthEmu = widthEmu;
            headStartX = (widthEmu - headWidthEmu) / (double)widthEmu * 100.0;
        }
        else
        {
            headStartX = 100.0 - adj2 / 1000.0; // fallback: treat adj2 as % of width
        }

        return $"clip-path:polygon(0 {tailTop:0.##}%,{headStartX:0.##}% {tailTop:0.##}%,{headStartX:0.##}% 0,100% 50%,{headStartX:0.##}% 100%,{headStartX:0.##}% {tailBot:0.##}%,0 {tailBot:0.##}%)";
    }

    /// <summary>
    /// Build a directional arrow clip-path honoring avLst, derived from the
    /// rightArrow geometry by mirroring/transposing the point coordinates.
    /// dir: "left" mirrors X, "up" transposes (swap X/Y), "down" transposes
    /// then mirrors Y. adj1/adj2 carry the same meaning as rightArrow.
    /// </summary>
    private static string DirectionalArrowPolygon(string dir, long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        // For up/down the head extends along height, so the head-width adj is
        // measured against the perpendicular dimension — swap dims when transposing.
        var (w, h) = (dir == "up" || dir == "down") ? (heightEmu, widthEmu) : (widthEmu, heightEmu);
        var baseCss = RightArrowPolygon(w, h, presetGeom); // "clip-path:polygon(...)"
        var inner = System.Text.RegularExpressions.Regex.Match(baseCss, @"polygon\(([^)]+)\)").Groups[1].Value;
        var pts = inner.Split(',');
        var outPts = new List<string>();
        foreach (var p in pts)
        {
            var xy = p.Trim().Split(' ');
            double x = ParsePct(xy[0]);
            double y = ParsePct(xy[1]);
            double nx, ny;
            switch (dir)
            {
                case "left": nx = 100 - x; ny = y; break;
                case "up":   nx = y; ny = 100 - x; break;  // rotate rightArrow 90deg CCW -> tip at top
                case "down": nx = 100 - y; ny = x; break;  // rotate rightArrow 90deg CW -> tip at bottom
                default:     nx = x; ny = y; break;
            }
            outPts.Add($"{nx:0.##}% {ny:0.##}%");
        }
        return $"clip-path:polygon({string.Join(",", outPts)})";
    }

    private static double ParsePct(string s) =>
        double.TryParse(s.TrimEnd('%').Trim(), System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;

    /// <summary>
    /// Build a clip-path polygon for a 5-point star honoring OOXML adj value.
    /// adj = inner radius fraction * 50000 (default 19098, giving inner ratio ~0.382).
    /// Star is stretched to fill bounding box (outer radius = min(w,h)/2 scaled independently to w,h).
    /// </summary>
    private static string Star5Polygon(Drawing.PresetGeometry? presetGeom)
    {
        var adj = ReadAdjValueCss(presetGeom, 0, 19098);
        if (adj < 0) adj = 0; if (adj > 50000) adj = 50000;
        var innerRatio = adj / 50000.0;

        var pts = new List<string>();
        // 10 points around the center, alternating outer (radius=0.5) and inner (radius=0.5*innerRatio).
        // Start at top (angle = -90°), step = 36° = PI/5. Scale x,y to 0..100%.
        for (int i = 0; i < 10; i++)
        {
            var angle = -Math.PI / 2 + Math.PI * i / 5;
            var r = (i % 2 == 0) ? 0.5 : 0.5 * innerRatio;
            var x = 50.0 + r * Math.Cos(angle) * 100.0;
            var y = 50.0 + r * Math.Sin(angle) * 100.0;
            pts.Add($"{x:0.##}% {y:0.##}%");
        }
        return $"clip-path:polygon({string.Join(",", pts)})";
    }

    /// <summary>
    /// R19 BUG A: build an N-point star clip-path honoring OOXML adj1 (inner-radius
    /// ratio), mirroring <see cref="Star5Polygon"/> but for arbitrary point counts.
    /// adj1 ranges 0..50000 mapping to inner ratio 0..1 (same scaling as star5,
    /// whose guide max is 50000). 2N vertices alternate outer radius (0.5) and
    /// inner radius (0.5·ratio), starting at the top (-90°).
    /// </summary>
    private static string StarNPolygon(int points, long defaultAdj, Drawing.PresetGeometry? presetGeom)
    {
        var adj = ReadAdjValueCss(presetGeom, 0, defaultAdj);
        if (adj < 0) adj = 0; if (adj > 50000) adj = 50000;
        var innerRatio = adj / 50000.0;
        var pts = new List<string>();
        for (int i = 0; i < points * 2; i++)
        {
            var angle = -Math.PI / 2 + Math.PI * i / points;
            var r = (i % 2 == 0) ? 0.5 : 0.5 * innerRatio;
            var x = 50.0 + r * Math.Cos(angle) * 100.0;
            var y = 50.0 + r * Math.Sin(angle) * 100.0;
            pts.Add($"{x:0.##}% {y:0.##}%");
        }
        return $"clip-path:polygon({string.Join(",", pts)})";
    }

    /// <summary>R19 BUG A: parallelogram honoring adj1 (top-left x offset fraction
    /// of width, ×100000; default 25000). polygon (a,0)(100,0)(100-a,100)(0,100).</summary>
    private static string ParallelogramPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000) / 100000.0 * 100.0, 0, 100);
        return $"clip-path:polygon({a:0.##}% 0,100% 0,{100 - a:0.##}% 100%,0 100%)";
    }

    /// <summary>R19 BUG A: trapezoid honoring adj1 (top-edge inset from each side,
    /// ×100000; default 25000). polygon (a,0)(100-a,0)(100,100)(0,100).</summary>
    private static string TrapezoidPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000) / 100000.0 * 100.0, 0, 50);
        return $"clip-path:polygon({a:0.##}% 0,{100 - a:0.##}% 0,100% 100%,0 100%)";
    }

    /// <summary>R19 BUG A: chevron (right-pointing) honoring adj1 (notch depth
    /// fraction, ×100000; default 50000).
    /// polygon (0,0)(100-a,0)(100,50)(100-a,100)(0,100)(a,50).</summary>
    private static string ChevronPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000) / 100000.0 * 100.0, 0, 100);
        return $"clip-path:polygon(0 0,{100 - a:0.##}% 0,100% 50%,{100 - a:0.##}% 100%,0 100%,{a:0.##}% 50%)";
    }

    /// <summary>R19 BUG A: hexagon honoring adj1 (corner inset fraction, ×100000;
    /// default 25000). polygon (a,0)(100-a,0)(100,50)(100-a,100)(a,100)(0,50).</summary>
    private static string HexagonPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000) / 100000.0 * 100.0, 0, 100);
        return $"clip-path:polygon({a:0.##}% 0,{100 - a:0.##}% 0,100% 50%,{100 - a:0.##}% 100%,{a:0.##}% 100%,0 50%)";
    }

    // ---- R18 BUG B/D/E: sector / ring / segment geometry --------------------
    // OOXML angle units are 60000ths of a degree, measured clockwise from the
    // 3-o'clock direction. In a 0..100% clip-path box, +x is right and +y is
    // down, so a point at angle θ on the unit circle maps to
    // (50 + 50·cosθ, 50 + 50·sinθ). We sample the arc into enough points for a
    // smooth silhouette, mirroring how custGeom beziers are flattened to
    // clip-path polygons elsewhere in this file.
    private static (double x, double y) ArcPointPct(double deg, double rx = 50, double ry = 50,
        double cx = 50, double cy = 50)
    {
        var rad = deg * Math.PI / 180.0;
        return (cx + rx * Math.Cos(rad), cy + ry * Math.Sin(rad));
    }

    private static IEnumerable<(double x, double y)> SampleArc(double startDeg, double endDeg,
        double rx, double ry, double cx, double cy)
    {
        // Sweep clockwise (increasing angle) from start to end.
        var sweep = endDeg - startDeg;
        if (sweep <= 0) sweep += 360;
        int steps = Math.Max(2, (int)Math.Ceiling(sweep / 6.0)); // ~6° per segment
        for (int i = 0; i <= steps; i++)
        {
            var d = startDeg + sweep * i / steps;
            yield return ArcPointPct(d, rx, ry, cx, cy);
        }
    }

    private static string PtsToPolygon(IEnumerable<(double x, double y)> pts)
    {
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        return $"clip-path:polygon({string.Join(",", pts.Select(p => $"{P(p.x)}% {P(p.y)}%"))})";
    }

    /// <summary>pie: filled sector from adj1 (start) to adj2 (end angle).</summary>
    private static string PieSectorPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a1 = ReadAdjValueCss(presetGeom, 0, 0) / 60000.0;
        var a2 = ReadAdjValueCss(presetGeom, 1, 16200000) / 60000.0;
        var pts = new List<(double, double)> { (50, 50) };
        pts.AddRange(SampleArc(a1, a2, 50, 50, 50, 50));
        return PtsToPolygon(pts);
    }

    /// <summary>arc: open sector outline (no center join) — approximate as a
    /// thin sector wedge so the silhouette traces the arc band.</summary>
    private static string ArcWedgePolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a1 = ReadAdjValueCss(presetGeom, 0, 16200000) / 60000.0;
        var a2 = ReadAdjValueCss(presetGeom, 1, 0) / 60000.0;
        // Outer arc start→end, then inner arc end→start (thin band ~ full radius
        // to 0.92 radius) so a stroke-less fill still reads as an arc curve.
        var outer = SampleArc(a1, a2, 50, 50, 50, 50).ToList();
        var inner = SampleArc(a1, a2, 46, 46, 50, 50).ToList();
        inner.Reverse();
        return PtsToPolygon(outer.Concat(inner));
    }

    /// <summary>chord: circular segment between an arc and its chord.</summary>
    private static string ChordPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a1 = ReadAdjValueCss(presetGeom, 0, 2700000) / 60000.0;
        var a2 = ReadAdjValueCss(presetGeom, 1, 16200000) / 60000.0;
        // Just the arc points; the polygon auto-closes start→end with the chord.
        return PtsToPolygon(SampleArc(a1, a2, 50, 50, 50, 50));
    }

    /// <summary>blockArc: thick ring segment — outer arc then inner arc back.
    /// adj1 start, adj2 end (60000ths deg), adj3 inner radius fraction (x100000).</summary>
    private static string BlockArcPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var a1 = ReadAdjValueCss(presetGeom, 0, 10800000) / 60000.0;
        var a2 = ReadAdjValueCss(presetGeom, 1, 0) / 60000.0;
        var innerFrac = ReadAdjValueCss(presetGeom, 2, 25000) / 100000.0;
        innerFrac = Math.Clamp(innerFrac, 0, 1);
        var ir = 50 * innerFrac;
        var outer = SampleArc(a1, a2, 50, 50, 50, 50).ToList();
        var inner = SampleArc(a1, a2, ir, ir, 50, 50).ToList();
        inner.Reverse();
        return PtsToPolygon(outer.Concat(inner));
    }

    /// <summary>snipRoundRect: top-left corner ROUNDED (adj2), top-right corner
    /// CHAMFERED/snipped (adj1); bottom-left and bottom-right square. Matches the
    /// real-PowerPoint silhouette (ground truth). Approximate the round
    /// corner with sample points and chamfer the snip with a single diagonal
    /// vertex.</summary>
    private static string SnipRoundRectPolygon(long widthEmu, long heightEmu,
        Drawing.PresetGeometry? presetGeom)
    {
        long minSideEmu = Math.Min(widthEmu, heightEmu);
        var a1 = ReadAdjValueCss(presetGeom, 0, 16667); // snip (top-right)
        var a2 = ReadAdjValueCss(presetGeom, 1, 16667); // round (top-left)
        double Pct(long adj, bool horizontal)
        {
            var sizeEmu = minSideEmu * adj / 100000.0;
            var axis = horizontal ? widthEmu : heightEmu;
            var pct = axis > 0 ? sizeEmu / axis * 100.0 : adj / 100000.0 * 100.0;
            return Math.Clamp(pct, 0, 50);
        }
        var shx = Pct(a1, true);   // snip horizontal inset
        var svy = Pct(a1, false);  // snip vertical inset
        var rhx = Pct(a2, true);   // round horizontal radius
        var rvy = Pct(a2, false);  // round vertical radius

        var pts = new List<(double x, double y)>();
        // top-left ROUNDED corner: quarter-circle (center at (rhx, rvy)) swept
        // from the left edge (180°) round to the top edge (270°). These points
        // sit in the top-left region — the silhouette's defining feature.
        foreach (var p in SampleArc(180, 270, rhx, rvy, rhx, rvy))
            pts.Add(p);
        // top edge to start of top-right snip
        pts.Add((100 - shx, 0));
        // chamfer down the right edge (top-right snip)
        pts.Add((100, svy));
        // right edge down to bottom-right (square)
        pts.Add((100, 100));
        // bottom edge to bottom-left (square)
        pts.Add((0, 100));
        // left edge back up to the start of the top-left round
        return PtsToPolygon(pts);
    }

    /// <summary>donut: full disc with a centered circular hole. OOXML adj1 is the
    /// ring thickness as a fraction of the box (×100000), so the hole radius as a
    /// percent of the box is holeRadiusPct = 50 - adj1/1000 (adj1=25000→25%,
    /// 45000→5%, 10000→40%). Default adj1=25000 when absent.</summary>
    private static string DonutCss(Drawing.PresetGeometry? presetGeom)
    {
        var adj1 = ReadAdjValueCss(presetGeom, 0, 25000);
        var holePct = Math.Clamp(50.0 - adj1 / 1000.0, 0, 50);
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        var stop = $"{P(holePct)}%";
        return $"border-radius:50%;-webkit-mask-image:radial-gradient(circle,transparent {stop},black {stop});"
             + $"mask-image:radial-gradient(circle,transparent {stop},black {stop})";
    }

    /// <summary>
    /// Picture-frame clip-path honoring OOXML avLst. adj1 = border thickness as a
    /// fraction of the shorter side (default 12500 = 12.5%). The border is equal
    /// absolute thickness on all sides, so the per-axis % differs for non-square
    /// shapes; adj1 large enough (inset >= 50% of a side) collapses the hole to a
    /// solid rectangle — matching PowerPoint.
    /// </summary>
    private static string FramePolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var adj1 = ReadAdjValueCss(presetGeom, 0, 12500);
        var minSide = Math.Min(widthEmu, heightEmu);
        double insetEmu = minSide * adj1 / 100000.0;
        var hx = widthEmu > 0 ? Math.Clamp(insetEmu / widthEmu * 100.0, 0, 50) : 12.5;
        var vy = heightEmu > 0 ? Math.Clamp(insetEmu / heightEmu * 100.0, 0, 50) : 12.5;
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        return $"clip-path:polygon(0 0,100% 0,100% 100%,0 100%,0 {P(vy)}%,{P(hx)}% {P(vy)}%,"
             + $"{P(hx)}% {P(100 - vy)}%,{P(100 - hx)}% {P(100 - vy)}%,{P(100 - hx)}% {P(vy)}%,0 {P(vy)}%)";
    }

    /// <summary>
    /// notchedRightArrow clip-path honoring avLst. adj1 = tail (shaft) height as a
    /// fraction of shape height (default 50000); adj2 = arrowhead width as a fraction
    /// of the shorter side (default 50000). The back-edge notch mirrors the
    /// arrowhead slope: its tip x = headWidthX% * (tail-height fraction). The old
    /// hardcoded polygon ignored both adj.
    /// </summary>
    private static string NotchedRightArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var adj1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000), 0, 100000);
        var adj2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 50000), 0, 100000);
        var minSide = Math.Min(widthEmu, heightEmu);
        var headWidthX = widthEmu > 0 ? Math.Clamp((double)adj2 / 100000.0 * minSide / widthEmu * 100.0, 0, 100) : 50.0;
        var headStartX = 100 - headWidthX;
        var tailTop = (100 - adj1 / 1000.0) / 2.0;
        var tailBottom = 100 - tailTop;
        var notchX = headWidthX * adj1 / 100000.0;
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        return $"clip-path:polygon(0 {P(tailTop)}%,{P(headStartX)}% {P(tailTop)}%,{P(headStartX)}% 0,"
             + $"100% 50%,{P(headStartX)}% 100%,{P(headStartX)}% {P(tailBottom)}%,0 {P(tailBottom)}%,{P(notchX)}% 50%)";
    }

    /// <summary>
    /// upDownArrow clip-path honoring avLst. adj1 = shaft width as a fraction of
    /// WIDTH (default 50000); adj2 = each arrowhead's length as a fraction of the
    /// SHORTER side (default 50000). The arrowhead base always spans the full
    /// width; the arrowhead depth (in %height) clamps to 50% so the two heads meet
    /// (a diamond) rather than overflow. The old hardcoded polygon ignored both.
    /// Mapping verified empirically against real PowerPoint.
    /// </summary>
    private static string UpDownArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var adj1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000), 0, 100000);
        var adj2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 50000), 0, 100000);
        var minSide = Math.Min(widthEmu, heightEmu);
        var depthY = Math.Clamp((double)adj2 / 100000.0 * minSide / heightEmu * 100.0, 0, 50);
        var shaftHalf = adj1 / 2000.0;
        var shaftL = Math.Clamp(50 - shaftHalf, 0, 50);
        var shaftR = Math.Clamp(50 + shaftHalf, 50, 100);
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        var d2 = 100 - depthY;
        return $"clip-path:polygon(50% 0,100% {P(depthY)}%,{P(shaftR)}% {P(depthY)}%,{P(shaftR)}% {P(d2)}%,"
             + $"100% {P(d2)}%,50% 100%,0 {P(d2)}%,{P(shaftL)}% {P(d2)}%,{P(shaftL)}% {P(depthY)}%,0 {P(depthY)}%)";
    }

    /// <summary>
    /// leftRightArrow clip-path honoring avLst — the horizontal mirror of
    /// UpDownArrowPolygon. adj1 = shaft thickness as a fraction of HEIGHT (default
    /// 50000); adj2 = each arrowhead's length as a fraction of the SHORTER side
    /// (default 50000). The arrowhead base spans the full height; the arrowhead
    /// depth (in %width) clamps to 50% so the heads meet (a horizontal diamond).
    /// The old hardcoded polygon ignored both adj. Verified empirically.
    /// </summary>
    private static string LeftRightArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var adj1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000), 0, 100000);
        var adj2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 50000), 0, 100000);
        var minSide = Math.Min(widthEmu, heightEmu);
        var depthX = Math.Clamp((double)adj2 / 100000.0 * minSide / widthEmu * 100.0, 0, 50);
        var shaftHalf = adj1 / 2000.0;
        var shaftT = Math.Clamp(50 - shaftHalf, 0, 50);
        var shaftB = Math.Clamp(50 + shaftHalf, 50, 100);
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        var d2 = 100 - depthX;
        return $"clip-path:polygon(0 50%,{P(depthX)}% 0,{P(depthX)}% {P(shaftT)}%,{P(d2)}% {P(shaftT)}%,"
             + $"{P(d2)}% 0,100% 50%,{P(d2)}% 100%,{P(d2)}% {P(shaftB)}%,{P(depthX)}% {P(shaftB)}%,{P(depthX)}% 100%)";
    }

    /// <summary>
    /// halfFrame (top-left L-bracket) clip-path honoring avLst. adj1 = top-arm
    /// thickness (fraction of HEIGHT, default 33333); adj2 = left-arm thickness
    /// (fraction of WIDTH). The two open ends are 45° mitered (equal absolute
    /// offset in EMU), so the miter %s are aspect-adjusted. The old hardcoded
    /// polygon used 15% square corners (no miters) and ignored adj. Verified
    /// empirically against real PowerPoint.
    /// </summary>
    private static string HalfFramePolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var adj1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 33333), 0, 100000);
        var adj2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 33333), 0, 100000);
        var armTopH = adj1 / 1000.0;                              // % of height
        var armLeftW = adj2 / 1000.0;                             // % of width
        var aspect = (double)heightEmu / widthEmu;
        var miterTopX = Math.Clamp(100 - armTopH * aspect, 0, 100);   // 45° miter (equal EMU offset)
        var miterBottomY = Math.Clamp(100 - armLeftW / aspect, 0, 100);
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        return $"clip-path:polygon(0 0,100% 0,{P(miterTopX)}% {P(armTopH)}%,{P(armLeftW)}% {P(armTopH)}%,"
             + $"{P(armLeftW)}% {P(miterBottomY)}%,0 100%)";
    }

    /// <summary>
    /// octagon clip-path honoring avLst. adj = corner-cut size as a fraction of the
    /// SHORTER side (default 29289 ≈ the regular octagon's 29.29%). For a square,
    /// corner% = adj/1000; adj=50000 collapses to a diamond. The old hardcoded 29%
    /// happened to match the default but ignored adj. Verified empirically.
    /// </summary>
    private static string OctagonPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var adj = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 29289), 0, 50000);
        var minSide = Math.Min(widthEmu, heightEmu);
        var cx = Math.Clamp((double)adj / 100000.0 * minSide / widthEmu * 100.0, 0, 50);
        var cy = Math.Clamp((double)adj / 100000.0 * minSide / heightEmu * 100.0, 0, 50);
        string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        return $"clip-path:polygon({P(cx)}% 0,{P(100 - cx)}% 0,100% {P(cy)}%,100% {P(100 - cy)}%,"
             + $"{P(100 - cx)}% 100%,{P(cx)}% 100%,0 {P(100 - cy)}%,0 {P(cy)}%)";
    }

    // moon: a crescent. adj = the rightward extent of the inner (bite) arc at the
    // vertical center, as a fraction of width (default 50000 = 50%). Both the outer
    // (left) boundary and the inner (bite) boundary are half-ellipses of vertical
    // radius = half the height, sharing the top/bottom tips at (100%,0)/(100%,100%).
    // The outer arc reaches the left edge (0%) at mid; the inner arc reaches adj% at
    // mid. The old hardcoded 30-point polygon locked the shape to ~adj=50000 (and
    // even then put the bite at 42%), so any non-default adj was invisible. Both
    // arcs verified empirically against real PowerPoint (R167).
    private static string MoonPolygon(Drawing.PresetGeometry? presetGeom)
    {
        var adjPct = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000) / 1000.0, 0, 100);
        var innerRx = 100.0 - adjPct; // inner-arc horizontal radius (bulge left from x=100)
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string Pt(double x, double y) => $"{x.ToString("0.##", ci)}% {y.ToString("0.##", ci)}%";
        var pts = new System.Collections.Generic.List<string>();
        const int N = 18;
        // inner (right/bite) arc: top tip -> bottom tip
        for (int i = 0; i <= N; i++)
        {
            double y = 100.0 * i / N;
            double k = Math.Sqrt(Math.Max(0, 1 - Math.Pow((y - 50) / 50.0, 2)));
            pts.Add(Pt(100.0 - innerRx * k, y));
        }
        // outer (left) arc: bottom tip -> top tip
        for (int i = N; i >= 0; i--)
        {
            double y = 100.0 * i / N;
            double k = Math.Sqrt(Math.Max(0, 1 - Math.Pow((y - 50) / 50.0, 2)));
            pts.Add(Pt(100.0 * (1 - k), y));
        }
        return "clip-path:polygon(" + string.Join(",", pts) + ")";
    }

    // mathMinus: a centered horizontal minus bar. adj1 (default 23520) is the bar's
    // full height as a fraction of the shape height (dy1 = adj1/200000 is the half
    // height). The horizontal extent is FIXED by ECMA at 73490/200000 (=36.745%)
    // either side of center, so the bar spans 13.26%-86.74% of width with white
    // gutters. The old literal spanned the full width and used a fixed 30% height,
    // ignoring adj1. Geometry verified empirically against real PowerPoint (R169).
    private static string MathMinusCss(Drawing.PresetGeometry? presetGeom)
    {
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 23520), 0, 100000);
        var dy = a1 / 200000.0 * 100.0;                 // half bar height, % of h
        const double dx = 73490.0 / 200000.0 * 100.0;   // 36.745% half-width, % of w
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string P(double d) => d.ToString("0.##", ci);
        double x1 = 50 - dx, x2 = 50 + dx, y1 = 50 - dy, y2 = 50 + dy;
        return $"clip-path:polygon({P(x1)}% {P(y1)}%,{P(x2)}% {P(y1)}%,{P(x2)}% {P(y2)}%,{P(x1)}% {P(y2)}%)";
    }

    // mathPlus: a centered plus/cross. The four arm tips touch the bounding-box
    // edges (0%/100% on both axes — verified in real PowerPoint). adj1 (default
    // 23520, pinned to [0,73490]) is the arm half-thickness: an ABSOLUTE length
    // = ss*adj1/146980 (ss = shorter side), so adj1=73490 fills the box solid.
    // The old literal hardcoded the arm notch at 33%/67% and ignored adj1 entirely.
    // Arm-thickness law verified empirically across adj1 in {10000,23520,50000,73490}.
    private static string MathPlusCss(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 23520), 0, 73490);
        double ss = Math.Min(widthEmu, heightEmu);
        double armAbs = ss * a1 / 146980.0;            // arm half-thickness, EMU
        double hx = armAbs / widthEmu * 100.0;          // as % of width
        double hy = armAbs / heightEmu * 100.0;         // as % of height
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string P(double d) => d.ToString("0.##", ci);
        double xl = 50 - hx, xr = 50 + hx, yt = 50 - hy, yb = 50 + hy;
        return "clip-path:polygon("
             + $"{P(xl)}% 0,{P(xr)}% 0,{P(xr)}% {P(yt)}%,100% {P(yt)}%,100% {P(yb)}%,"
             + $"{P(xr)}% {P(yb)}%,{P(xr)}% 100%,{P(xl)}% 100%,{P(xl)}% {P(yb)}%,"
             + $"0 {P(yb)}%,0 {P(yt)}%,{P(xl)}% {P(yt)}%)";
    }

    // mathMultiply: a diagonal multiplication X pointing to the four corners. adj1
    // (default 23520, pinned [0,51965]) is the arm thickness: th = ss*adj1/100000.
    // The 12 vertices come straight from the ECMA-376 guide chain (atan2 diagonal
    // angle, half-diagonal arm length rw = dl*0.51965). The old literal was a fixed
    // corner-X that ignored adj1 (and was a touch too wide). Even at max adj1 all
    // vertices stay inside the box. Formula verified against the ECMA definition and
    // real PowerPoint (default + adj1=40000).
    private static string MathMultiplyCss(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu;
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 23520), 0, 51965);
        double ss = Math.Min(w, h);
        double th = ss * a1 / 100000.0;
        double a = Math.Atan2(h, w);
        double sa = Math.Sin(a), ca = Math.Cos(a), ta = Math.Tan(a);
        double hc = w / 2, vc = h / 2;
        double dl = Math.Sqrt(w * w + h * h);
        double rw = dl * 51965 / 100000.0;
        double lM = dl - rw;
        double xM = ca * lM / 2, yM = sa * lM / 2;
        double dxAM = sa * th / 2, dyAM = ca * th / 2;
        double xA = xM - dxAM, yA = yM + dyAM;
        double xB = xM + dxAM, yB = yM - dyAM;
        double yC = (hc - xB) * ta + yB;
        double xD = w - xB, xE = w - xA;
        double yFE = vc - yA, xFE = yFE / ta;
        double xF = xE - xFE, xL = xA + xFE;
        double yG = h - yA, yH = h - yB, yI = h - yC;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string Pt(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"{Pt(xA, yA)},{Pt(xB, yB)},{Pt(hc, yC)},{Pt(xD, yB)},{Pt(xE, yA)},{Pt(xF, vc)},"
             + $"{Pt(xE, yG)},{Pt(xD, yH)},{Pt(hc, yI)},{Pt(xB, yH)},{Pt(xA, yG)},{Pt(xL, vc)})";
    }

    // quadArrow: a four-headed arrow cross (N/S/E/W). It was entirely missing from
    // the preset switch, so it fell through to no clip-path and rendered as a plain
    // rectangle. adj1 = shaft thickness, adj2 = arrowhead half-width, adj3 = head
    // inset/notch depth (all default 22500). The ECMA guide chain pins them
    // (a1<=2*a2, a3<=(100000-2*a2)/2). 24 straight-edge vertices, single connected.
    // Path + guides taken verbatim from the ECMA-376 preset definition; verified
    // against real PowerPoint at default and adj1=10000/adj2=30000/adj3=30000.
    private static string QuadArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 22500), 0, 50000);
        var maxAdj1 = a2 * 2;
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 22500), 0, maxAdj1);
        var maxAdj3 = (100000 - maxAdj1) / 2;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 22500), 0, maxAdj3);
        double x1 = ss * a3 / 100000.0;       // inset depth
        double dx2 = ss * a2 / 100000.0;      // arrowhead half-width
        double dx3 = ss * a1 / 200000.0;      // shaft half-width
        double hc = w / 2, vc = h / 2;
        double x2 = hc - dx2, x5 = hc + dx2, x3 = hc - dx3, x4 = hc + dx3, x6 = w - x1;
        double y2 = vc - dx2, y5 = vc + dx2, y3 = vc - dx3, y4 = vc + dx3, y6 = h - x1;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"{P(0, vc)},{P(x1, y2)},{P(x1, y3)},{P(x3, y3)},{P(x3, x1)},{P(x2, x1)},"
             + $"{P(hc, 0)},{P(x5, x1)},{P(x4, x1)},{P(x4, y3)},{P(x6, y3)},{P(x6, y2)},"
             + $"{P(w, vc)},{P(x6, y5)},{P(x6, y4)},{P(x4, y4)},{P(x4, y6)},{P(x5, y6)},"
             + $"{P(hc, h)},{P(x2, y6)},{P(x3, y6)},{P(x3, y4)},{P(x1, y4)},{P(x1, y5)})";
    }

    // rightArrowCallout: a rectangular callout body with a right-pointing arrow on the
    // right edge. It was entirely missing from the preset switch, so it emitted no
    // clip-path and rendered as a plain rectangle. adj1 = shaft half-height, adj2 =
    // arrowhead half-height, adj3 = arrowhead depth, adj4 = body right-edge x (all of
    // ss/w bases per the ECMA guide chain, with per-axis pins). 11 straight-edge
    // vertices, single connected. Verified against real PowerPoint (default + non-default).
    private static string RightArrowCalloutPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        double maxAdj2 = 50000.0 * h / ss;
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, maxAdj2);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, a2 * 2);
        double maxAdj3 = 100000.0 * w / ss;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double q2 = a3 * ss / w, maxAdj4 = 100000 - q2;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 64977), 0, maxAdj4);
        double vc = h / 2;
        double dy1 = ss * a2 / 100000.0, dy2 = ss * a1 / 200000.0;
        double y1 = vc - dy1, y2 = vc - dy2, y3 = vc + dy2, y4 = vc + dy1;
        double dx3 = ss * a3 / 100000.0, x3 = w - dx3;
        double x2 = w * a4 / 100000.0;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 0,{P(x2, 0)},{P(x2, y2)},{P(x3, y2)},{P(x3, y1)},{P(w, vc)},"
             + $"{P(x3, y4)},{P(x3, y3)},{P(x2, y3)},{P(x2, h)},0 100%)";
    }

    // leftArrowCallout: rectangular body on the right with a left-pointing arrow. Was
    // missing from the switch (rendered as a plain rectangle). Mirror of
    // rightArrowCallout: adj1=shaft half-height, adj2=arrowhead half-height, adj3=
    // arrowhead depth (from left), adj4=body width. 11 straight-edge vertices, single
    // connected. Verified against real PowerPoint (default + non-default).
    private static string LeftArrowCalloutPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        double maxAdj2 = 50000.0 * h / ss;
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, maxAdj2);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, a2 * 2);
        double maxAdj3 = 100000.0 * w / ss;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double q2 = a3 * ss / w, maxAdj4 = 100000 - q2;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 64977), 0, maxAdj4);
        double vc = h / 2;
        double dy1 = ss * a2 / 100000.0, dy2 = ss * a1 / 200000.0;
        double y1 = vc - dy1, y2 = vc - dy2, y3 = vc + dy2, y4 = vc + dy1;
        double x1 = ss * a3 / 100000.0;       // arrowhead depth from left
        double x2 = w - w * a4 / 100000.0;    // body left edge
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 50%,{P(x1, y1)},{P(x1, y2)},{P(x2, y2)},{P(x2, 0)},{P(w, 0)},"
             + $"{P(w, h)},{P(x2, h)},{P(x2, y3)},{P(x1, y3)},{P(x1, y4)})";
    }

    // upArrowCallout: rectangular body at the bottom with an up-pointing arrow on top.
    // Was missing from the switch (rendered as a plain rectangle). 90deg variant of
    // rightArrowCallout: adj1=shaft half-width, adj2=arrowhead half-width, adj3=
    // arrowhead depth (from top), adj4=body height. 11 straight-edge vertices, single
    // connected. Verified against real PowerPoint (default + non-default).
    private static string UpArrowCalloutPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        double maxAdj2 = 50000.0 * w / ss;
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, maxAdj2);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, a2 * 2);
        double maxAdj3 = 100000.0 * h / ss;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double q2 = a3 * ss / h, maxAdj4 = 100000 - q2;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 64977), 0, maxAdj4);
        double hc = w / 2;
        double dx1 = ss * a2 / 100000.0, dx2 = ss * a1 / 200000.0;
        double x1 = hc - dx1, x2 = hc - dx2, x3 = hc + dx2, x4 = hc + dx1;
        double y1 = ss * a3 / 100000.0;       // arrowhead depth from top
        double y2 = h - h * a4 / 100000.0;    // body top edge
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 {Y(y2)}%,{P(x2, y2)},{P(x2, y1)},{P(x1, y1)},{P(hc, 0)},{P(x4, y1)},"
             + $"{P(x3, y1)},{P(x3, y2)},{P(w, y2)},{P(w, h)},0 100%)";
    }

    // downArrowCallout: rectangular body at the top with a down-pointing arrow below.
    // Was missing from the switch (rendered as a plain rectangle). adj1=shaft half-
    // width, adj2=arrowhead half-width, adj3=arrowhead depth (from bottom), adj4=body
    // height. 11 straight-edge vertices, single connected. Verified against real
    // PowerPoint (default + non-default).
    private static string DownArrowCalloutPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        double maxAdj2 = 50000.0 * w / ss;
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, maxAdj2);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, a2 * 2);
        double maxAdj3 = 100000.0 * h / ss;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double q2 = a3 * ss / h, maxAdj4 = 100000 - q2;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 64977), 0, maxAdj4);
        double hc = w / 2;
        double dx1 = ss * a2 / 100000.0, dx2 = ss * a1 / 200000.0;
        double x1 = hc - dx1, x2 = hc - dx2, x3 = hc + dx2, x4 = hc + dx1;
        double y3 = h - ss * a3 / 100000.0;   // arrowhead base (near bottom)
        double y2 = h * a4 / 100000.0;        // body bottom edge
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 0,{P(w, 0)},{P(w, y2)},{P(x3, y2)},{P(x3, y3)},{P(x4, y3)},{P(hc, h)},"
             + $"{P(x1, y3)},{P(x2, y3)},{P(x2, y2)},0 {Y(y2)}%)";
    }

    // leftRightArrowCallout: a centered rectangular body with arrows pointing both left
    // and right. Was missing from the switch (rendered as a plain rectangle). adj1=shaft
    // half-height, adj2=arrowhead half-height, adj3=arrowhead depth, adj4=body half-width.
    // 18 straight-edge vertices, single connected. Verified against real PowerPoint.
    private static string LeftRightArrowCalloutPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        double maxAdj2 = 50000.0 * h / ss;
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, maxAdj2);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, a2 * 2);
        double maxAdj3 = 50000.0 * w / ss;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double q2 = a3 * ss / (w / 2), maxAdj4 = 100000 - q2;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 48123), 0, maxAdj4);
        double hc = w / 2, vc = h / 2;
        double dy1 = ss * a2 / 100000.0, dy2 = ss * a1 / 200000.0;
        double y1 = vc - dy1, y2 = vc - dy2, y3 = vc + dy2, y4 = vc + dy1;
        double x1 = ss * a3 / 100000.0, x4 = w - x1;
        double dx2 = w * a4 / 200000.0, x2 = hc - dx2, x3 = hc + dx2;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 50%,{P(x1, y1)},{P(x1, y2)},{P(x2, y2)},{P(x2, 0)},{P(x3, 0)},{P(x3, y2)},"
             + $"{P(x4, y2)},{P(x4, y1)},{P(w, vc)},{P(x4, y4)},{P(x4, y3)},{P(x3, y3)},"
             + $"{P(x3, h)},{P(x2, h)},{P(x2, y3)},{P(x1, y3)},{P(x1, y4)})";
    }

    // upDownArrowCallout: a centered rectangular body with arrows pointing both up and
    // down. Was missing from the switch (rendered as a plain rectangle). adj1=shaft
    // half-width, adj2=arrowhead half-width, adj3=arrowhead depth, adj4=body half-height.
    // 18 straight-edge vertices, single connected. Verified against real PowerPoint.
    private static string UpDownArrowCalloutPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        double maxAdj2 = 50000.0 * w / ss;
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, maxAdj2);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, a2 * 2);
        double maxAdj3 = 50000.0 * h / ss;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double q2 = a3 * ss / (h / 2), maxAdj4 = 100000 - q2;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 48123), 0, maxAdj4);
        double hc = w / 2, vc = h / 2;
        double dx1 = ss * a2 / 100000.0, dx2 = ss * a1 / 200000.0;
        double x1 = hc - dx1, x2 = hc - dx2, x3 = hc + dx2, x4 = hc + dx1;
        double y1 = ss * a3 / 100000.0, y4 = h - y1;
        double dy2 = h * a4 / 200000.0, y2 = vc - dy2, y3 = vc + dy2;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 {Y(y2)}%,{P(x2, y2)},{P(x2, y1)},{P(x1, y1)},{P(hc, 0)},{P(x4, y1)},{P(x3, y1)},"
             + $"{P(x3, y2)},{P(w, y2)},{P(w, y3)},{P(x3, y3)},{P(x3, y4)},{P(x4, y4)},"
             + $"{P(hc, h)},{P(x1, y4)},{P(x2, y4)},{P(x2, y3)},0 {Y(y3)}%)";
    }

    // quadArrowCallout: a central rectangular body with arrows pointing all four ways.
    // Was missing from the switch (rendered as a plain rectangle). adj1=shaft half-
    // width, adj2=arrowhead half-width, adj3=arrowhead depth, adj4=body half-size (note
    // its pin lower-bound is a1, not 0). 32 straight-edge vertices, single connected.
    // Verified against real PowerPoint (default + non-default).
    private static string QuadArrowCalloutPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 18515), 0, 50000);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 18515), 0, a2 * 2);
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 18515), 0, 50000 - a2);
        double maxAdj4 = 100000 - a3 * 2;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 48123), a1, maxAdj4);
        double hc = w / 2, vc = h / 2;
        double dx2 = ss * a2 / 100000.0, dx3 = ss * a1 / 200000.0, ah = ss * a3 / 100000.0;
        double dx1 = w * a4 / 200000.0, dy1 = h * a4 / 200000.0;
        double x8 = w - ah, x2 = hc - dx1, x7 = hc + dx1, x3 = hc - dx2, x6 = hc + dx2, x4 = hc - dx3, x5 = hc + dx3;
        double y8 = h - ah, y2 = vc - dy1, y7 = vc + dy1, y3 = vc - dx2, y6 = vc + dx2, y4 = vc - dx3, y5 = vc + dx3;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 {Y(vc)}%,{P(ah, y3)},{P(ah, y4)},{P(x2, y4)},{P(x2, y2)},{P(x4, y2)},{P(x4, ah)},{P(x3, ah)},"
             + $"{P(hc, 0)},{P(x6, ah)},{P(x5, ah)},{P(x5, y2)},{P(x7, y2)},{P(x7, y4)},{P(x8, y4)},{P(x8, y3)},"
             + $"{P(w, vc)},{P(x8, y6)},{P(x8, y5)},{P(x7, y5)},{P(x7, y7)},{P(x5, y7)},{P(x5, y8)},{P(x6, y8)},"
             + $"{P(hc, h)},{P(x3, y8)},{P(x4, y8)},{P(x4, y7)},{P(x2, y7)},{P(x2, y5)},{P(ah, y5)},{P(ah, y6)})";
    }

    // bentUpArrow: an L-shaped arrow — horizontal stem from the left rising into a
    // vertical shaft on the right with an up-pointing arrowhead. Was missing from the
    // switch (rendered as a plain rectangle). adj1=stem/shaft thickness, adj2=arrowhead
    // width, adj3=elbow height (all fractions of ss). 9 straight-edge vertices, single
    // connected. Verified against real PowerPoint (default + non-default).
    private static string BentUpArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, 50000);
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, 50000);
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, 50000);
        double y1 = ss * a3 / 100000.0;
        double x1 = w - ss * a2 / 50000.0;
        double x3 = w - ss * a2 / 100000.0;
        double dx2 = ss * a1 / 200000.0, x2 = x3 - dx2, x4 = x3 + dx2;
        double y2 = h - ss * a1 / 100000.0;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 {Y(y2)}%,{P(x2, y2)},{P(x2, y1)},{P(x1, y1)},{P(x3, 0)},{P(w, y1)},"
             + $"{P(x4, y1)},{P(x4, h)},0 100%)";
    }

    // leftUpArrow: a two-pronged L-shaped double arrow (one arm points left, one points
    // up). Was missing from the switch (rendered as a plain rectangle). adj1=stem width,
    // adj2=arrowhead half-width, adj3=stem length (fractions of ss). 12 straight-edge
    // vertices, single connected. Verified against real PowerPoint (default + non-default).
    private static string LeftUpArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, 50000);
        double maxAdj1 = a2 * 2;
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, maxAdj1);
        double maxAdj3 = 100000 - maxAdj1;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double x1 = ss * a3 / 100000.0;
        double dx2 = ss * a2 / 50000.0, x2 = w - dx2, y2 = h - dx2;
        double dx4 = ss * a2 / 100000.0, x4 = w - dx4, y4 = h - dx4;
        double dx3 = ss * a1 / 200000.0, x3 = x4 - dx3, x5 = x4 + dx3, y3 = y4 - dx3, y5 = y4 + dx3;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 {Y(y4)}%,{P(x1, y2)},{P(x1, y3)},{P(x3, y3)},{P(x3, x1)},{P(x2, x1)},"
             + $"{P(x4, 0)},{P(w, x1)},{P(x5, x1)},{P(x5, y5)},{P(x1, y5)},{P(x1, h)})";
    }

    // leftRightUpArrow: a three-headed arrow (left + right + up) sharing a central hub.
    // Was missing from the switch (rendered as a plain rectangle). adj1=shaft thickness,
    // adj2=arrowhead half-width, adj3=arrowhead depth (fractions of ss). 17 straight-edge
    // vertices, single connected. Verified against real PowerPoint (default + non-default).
    private static string LeftRightUpArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, 50000);
        double maxAdj1 = a2 * 2;
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, maxAdj1);
        double maxAdj3 = (100000 - maxAdj1) / 2;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double hc = w / 2;
        double x1 = ss * a3 / 100000.0;
        double dx2 = ss * a2 / 100000.0, x2 = hc - dx2, x5 = hc + dx2;
        double dx3 = ss * a1 / 200000.0, x3 = hc - dx3, x4 = hc + dx3;
        double x6 = w - x1;
        double dy2 = ss * a2 / 50000.0, y2 = h - dy2;
        double y4 = h - dx2, y3 = y4 - dx3, y5 = y4 + dx3;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        string P(double x, double y) => $"{X(x)}% {Y(y)}%";
        return "clip-path:polygon("
             + $"0 {Y(y4)}%,{P(x1, y2)},{P(x1, y3)},{P(x3, y3)},{P(x3, x1)},{P(x2, x1)},"
             + $"{P(hc, 0)},{P(x5, x1)},{P(x4, x1)},{P(x4, y3)},{P(x6, y3)},{P(x6, y2)},"
             + $"{P(w, y4)},{P(x6, h)},{P(x6, y5)},{P(x1, y5)},{P(x1, h)})";
    }

    // nonIsoscelesTrapezoid: a trapezoid (wide bottom, narrow top) whose two top vertices
    // are inset independently. Was missing from the switch (rendered as a plain rectangle).
    // adj1=left inset, adj2=right inset (both default 25000). Insets are ABSOLUTE lengths
    // (ss-based), so for a non-square shape they are NOT simply adj% of width. 4 straight-
    // edge vertices, single connected. Verified against real PowerPoint (default + asymmetric).
    private static string NonIsoscelesTrapezoidPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        double maxAdj = 50000.0 * w / ss;
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, maxAdj);
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, maxAdj);
        double x2 = ss * a1 / 100000.0;
        double x3 = w - ss * a2 / 100000.0;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        return $"clip-path:polygon(0 100%,{X(x2)}% 0,{X(x3)}% 0,100% 100%)";
    }

    // plaque: a rectangle whose four corners are scooped inward by concave quarter-circle
    // arcs (centered at the OUTER corners, radius = ss*adj/100000, default adj 16667). Was
    // missing from the switch (rendered as a plain rectangle). At large adj the scoops nearly
    // meet, yielding a 4-pointed concave star. Each arc sampled into 7 points; verified
    // against real PowerPoint (default + adj=40000).
    private static string PlaquePolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var adj = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 16667), 0, 50000);
        double r = ss * adj / 100000.0;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        var pts = new System.Collections.Generic.List<string>();
        const int N = 6;
        void Arc(double cx, double cy, double a0, double a1)
        {
            for (int i = 0; i <= N; i++)
            {
                double a = (a0 + (a1 - a0) * i / N) * Math.PI / 180.0;
                pts.Add($"{X(cx + r * Math.Cos(a))}% {Y(cy + r * Math.Sin(a))}%");
            }
        }
        Arc(0, 0, 90, 0);      // top-left scoop
        Arc(w, 0, 180, 90);    // top-right scoop
        Arc(w, h, 270, 180);   // bottom-right scoop
        Arc(0, h, 360, 270);   // bottom-left scoop
        return "clip-path:polygon(" + string.Join(",", pts) + ")";
    }

    // pieWedge: a quarter-ellipse wedge. No adjusts. The curved edge is a quarter of the
    // box-inscribed ellipse centered at the bottom-right corner (radii = full w,h), swept
    // from the bottom-left corner (180deg) to the top-right corner (270deg); the right and
    // bottom edges are straight. Was missing from the switch (rendered as a rectangle). The
    // shape is aspect-independent in %. Verified against real PowerPoint.
    private static string PieWedgeCss()
    {
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        var pts = new System.Collections.Generic.List<string>();
        const int N = 10;
        for (int i = 0; i <= N; i++)
        {
            double a = (180 + 90.0 * i / N) * Math.PI / 180.0;
            double x = (1 + Math.Cos(a)) * 100, y = (1 + Math.Sin(a)) * 100;
            pts.Add($"{x.ToString("0.##", ci)}% {y.ToString("0.##", ci)}%");
        }
        pts.Add("100% 100%");
        return "clip-path:polygon(" + string.Join(",", pts) + ")";
    }

    // teardrop: the box-inscribed ellipse with its upper-right quarter pulled into a point
    // toward the top-right. adj (default 100000, range 0..200000) sets the tail length: the
    // tail tip is at (hc + wd2*adj/100000, vc - hd2*adj/100000) — at the corner for the
    // default, beyond the box past 100000, no tail at 0. Built from 3 ellipse quarters plus
    // two quadratic beziers for the tail. Was missing (rendered as a rectangle). Verified
    // against real PowerPoint (default + adj=50000).
    private static string TeardropPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 100000), 0, 200000);
        double dx1 = wd2 * adj / 100000.0, dy1 = hd2 * adj / 100000.0;
        double x1 = hc + dx1, y1 = vc - dy1;        // tail tip
        double x2 = (hc + x1) / 2, y2 = (vc + y1) / 2;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        var pts = new System.Collections.Generic.List<string>();
        void Pt(double x, double y) => pts.Add($"{X(x)}% {Y(y)}%");
        const int N = 8;
        void Arc(double a0, double a1) { for (int i = 0; i <= N; i++) { double a = (a0 + (a1 - a0) * i / N) * Math.PI / 180.0; Pt(hc + wd2 * Math.Cos(a), vc + hd2 * Math.Sin(a)); } }
        void Quad(double px0, double py0, double cx, double cy, double px1, double py1)
        { for (int i = 1; i <= N; i++) { double t = (double)i / N, u = 1 - t; Pt(u * u * px0 + 2 * u * t * cx + t * t * px1, u * u * py0 + 2 * u * t * cy + t * t * py1); } }
        Arc(180, 270);                       // upper-left quarter: (0,vc) -> (hc,0)
        Quad(hc, 0, x2, 0, x1, y1);          // top-mid -> tail tip
        Quad(x1, y1, w, y2, w, vc);          // tail tip -> right-mid
        Arc(0, 90);                          // lower-right quarter: (w,vc) -> (hc,h)
        Arc(90, 180);                        // lower-left quarter: (hc,h) -> (0,vc)
        return "clip-path:polygon(" + string.Join(",", pts) + ")";
    }

    // swooshArrow: a curved swoosh sweeping up from the bottom-left to an arrowhead near
    // the top-right. adj1 (default 25000) = arrowhead height, adj2 (default 16667) =
    // arrowhead inset from the right. Single connected path: 2 quadratic beziers (the
    // curved body edges) + 4 straight lines (the arrowhead). Was missing (rendered as a
    // rectangle). Guide chain from the ECMA preset definition; verified against PowerPoint.
    private static string SwooshArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 1, 75000);
        double maxAdj2 = 70000.0 * w / ss;
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 16667), 0, maxAdj2);
        double ad1 = h * a1 / 100000.0, ad2 = ss * a2 / 100000.0, ssd8 = ss / 8.0;
        double xB = w - ad2, yB = ssd8;
        double alfa = (Math.PI / 2) / 14.0;
        double dx0 = ssd8 * Math.Tan(alfa), xC = xB - dx0;
        double dx1 = ad1 * Math.Tan(alfa), yF = yB + ad1, xF = xB + dx1, xE = xF + dx0, yE = yF + ssd8;
        double yD = yE / 2 - h / 20.0;
        double yP1 = h / 3.0, xP1 = w / 6.0;
        double yP2 = yF + (h / 6.0) / 2.0, xP2 = w / 4.0;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        var pts = new System.Collections.Generic.List<string>();
        void Pt(double x, double y) => pts.Add($"{X(x)}% {Y(y)}%");
        const int N = 10;
        void Quad(double px0, double py0, double cx, double cy, double px1, double py1)
        { for (int i = 1; i <= N; i++) { double tt = (double)i / N, u = 1 - tt; Pt(u * u * px0 + 2 * u * tt * cx + tt * tt * px1, u * u * py0 + 2 * u * tt * cy + tt * tt * py1); } }
        Pt(0, h);                              // bottom-left start
        Quad(0, h, xP1, yP1, xB, yB);          // upper body edge -> (xB,yB)
        Pt(xC, 0); Pt(w, yD); Pt(xE, yE); Pt(xF, yF);   // arrowhead
        Quad(xF, yF, xP2, yP2, 0, h);          // lower body edge back to start
        return "clip-path:polygon(" + string.Join(",", pts) + ")";
    }

    // uturnArrow: a U-turn arrow — up the left arm, semicircular bend over the top, down
    // the right arm into a downward arrowhead. 5 adjusts (shaft thickness, arrowhead width,
    // arrowhead length, bend radius, arrowhead-tip height). Single connected path: 4 circular
    // arcs (outer + inner bend halves) + straight arms/arrowhead. Was missing (rectangle).
    // Guide chain from the ECMA preset definition; verified against PowerPoint.
    private static string UturnArrowPolygon(long widthEmu, long heightEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = widthEmu, h = heightEmu, ss = Math.Min(w, h);
        var a2 = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 25000), 0, 25000);
        double maxAdj1 = a2 * 2;
        var a1 = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000), 0, maxAdj1);
        double q2 = a1 * ss / h, q3 = 100000 - q2, maxAdj3 = q3 * h / ss;
        var a3 = Math.Clamp(ReadAdjValueCss(presetGeom, 2, 25000), 0, maxAdj3);
        double q1 = a3 - a1, minAdj5 = q1 * ss / h;
        var a5 = Math.Clamp(ReadAdjValueCss(presetGeom, 4, 75000), minAdj5, 100000);
        double th = ss * a1 / 100000.0, aw2 = ss * a2 / 100000.0, th2 = th / 2, dh2 = aw2 - th2;
        double y5 = h * a5 / 100000.0, ah = ss * a3 / 100000.0, y4 = y5 - ah;
        double x9 = w - dh2, bw = x9 / 2, bs = Math.Min(bw, y4), maxAdj4 = bs * 100000 / ss;
        var a4 = Math.Clamp(ReadAdjValueCss(presetGeom, 3, 43750), 0, maxAdj4);
        double bd = ss * a4 / 100000.0, bd2 = Math.Max(bd - th, 0);
        double x3 = th + bd2, x8 = w - aw2, x6 = x8 - aw2, x7 = x6 + dh2, x4 = x9 - bd;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        string X(double v) => (v / w * 100).ToString("0.##", ci);
        string Y(double v) => (v / h * 100).ToString("0.##", ci);
        var pts = new System.Collections.Generic.List<string>();
        void Pt(double x, double y) => pts.Add($"{X(x)}% {Y(y)}%");
        const int N = 8;
        void Arc(double cx, double cy, double r, double d0, double d1)
        { for (int i = 0; i <= N; i++) { double a = (d0 + (d1 - d0) * i / N) * Math.PI / 180.0; Pt(cx + r * Math.Cos(a), cy + r * Math.Sin(a)); } }
        Pt(0, h); Pt(0, bd);
        Arc(bd, bd, bd, 180, 270);     // outer bend, left half
        Pt(x4, 0);
        Arc(x4, bd, bd, 270, 360);     // outer bend, right half
        Pt(x9, y4); Pt(w, y4); Pt(x8, y5); Pt(x6, y4); Pt(x7, y4); Pt(x7, x3);
        Arc(x7 - bd2, x3, bd2, 0, -90);   // inner bend, right half
        Pt(x3, th);
        Arc(x3, x3, bd2, 270, 180);    // inner bend, left half
        Pt(th, h);
        return "clip-path:polygon(" + string.Join(",", pts) + ")";
    }

    // flowChartMagneticTape: the box-inscribed ellipse with a small rectangular tab at the
    // bottom-right (where the circle meets the baseline). No adjusts. The ellipse is swept
    // from the bottom (90deg) almost all the way around to ang1=atan2(h,w) past 0deg; then a
    // horizontal line out to the right edge, down to the bottom-right corner, and back along
    // the bottom. Was missing (rendered as a rectangle). Verified against real PowerPoint.
    private static string FlowChartMagneticTapeCss(long widthEmu, long heightEmu)
    {
        double w = widthEmu, h = heightEmu;
        double ang1 = Math.Atan2(h, w) * 180.0 / Math.PI;
        double ibPct = 50.0 * (1 + Math.Sin(Math.PI / 4));   // ib as % of h
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        var pts = new System.Collections.Generic.List<string>();
        const int N = 28;
        double d0 = 90, d1 = 360 + ang1;
        for (int i = 0; i <= N; i++)
        {
            double a = (d0 + (d1 - d0) * i / N) * Math.PI / 180.0;
            double x = 50 * (1 + Math.Cos(a)), y = 50 * (1 + Math.Sin(a));
            pts.Add($"{x.ToString("0.##", ci)}% {y.ToString("0.##", ci)}%");
        }
        pts.Add($"100% {ibPct.ToString("0.##", ci)}%");
        pts.Add("100% 100%");
        return "clip-path:polygon(" + string.Join(",", pts) + ")";
    }

    private static string PresetGeometryToCss(string preset, long widthEmu, long heightEmu,
        Drawing.PresetGeometry? presetGeom)
    {
        // Wedge callouts: tail tip follows adj1/adj2 (ECMA preset formulas).
        if (IsWedgeCallout(preset)) return WedgeCalloutClipPath(preset, widthEmu, heightEmu, presetGeom) ?? "";
        // R18: sector / ring / segment presets honoring avLst angles
        if (preset == "pie") return PieSectorPolygon(presetGeom);
        if (preset == "arc") return ArcWedgePolygon(presetGeom);
        if (preset == "chord") return ChordPolygon(presetGeom);
        if (preset == "blockArc") return BlockArcPolygon(presetGeom);
        if (preset == "snipRoundRect") return SnipRoundRectPolygon(widthEmu, heightEmu, presetGeom);

        // Parametric arrows honoring avLst
        if (preset == "rightArrow")
            return RightArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "leftArrow")
            return DirectionalArrowPolygon("left", widthEmu, heightEmu, presetGeom);
        if (preset == "upArrow")
            return DirectionalArrowPolygon("up", widthEmu, heightEmu, presetGeom);
        if (preset == "downArrow")
            return DirectionalArrowPolygon("down", widthEmu, heightEmu, presetGeom);
        // Parametric stars honoring avLst (inner-radius ratio)
        if (preset == "star5")
            return Star5Polygon(presetGeom);
        if (preset == "star4") return StarNPolygon(4, 38250, presetGeom);
        if (preset == "star6") return StarNPolygon(6, 28868, presetGeom);
        if (preset == "star7") return StarNPolygon(7, 34601, presetGeom);
        if (preset == "star8") return StarNPolygon(8, 37500, presetGeom);
        if (preset == "star10") return StarNPolygon(10, 42533, presetGeom);
        if (preset == "star12") return StarNPolygon(12, 37500, presetGeom);
        if (preset == "star16") return StarNPolygon(16, 37500, presetGeom);
        if (preset == "star24") return StarNPolygon(24, 37500, presetGeom);
        if (preset == "star32") return StarNPolygon(32, 37500, presetGeom);

        // irregularSeal1/2 ("explosion" starbursts): fixed jagged polygons with no
        // adjusts (hard-coded vertices in the ECMA 21600 grid). Were missing from the
        // switch → rendered as a plain rectangle. Vertices converted straight from the
        // ECMA preset definition; verified against real PowerPoint.
        if (preset == "irregularSeal1")
            return "clip-path:polygon(50% 26.85%,67.23% 0%,65.53% 24.65%,85.09% 20.63%,77.32% 33.87%,97.67% 37.67%,81.51% 48.5%,100% 61.53%,77.95% 59.92%,84% 83.77%,64.91% 66.93%,61.33% 91.38%,48.76% 69.14%,39.28% 100%,35.72% 72.35%,22.05% 81.56%,26.24% 64.52%,0.62% 67.53%,17.23% 54.51%,0% 39.88%,21.42% 35.26%,1.71% 10.62%,33.85% 29.26%,38.67% 10.62%)";
        if (preset == "irregularSeal2")
            return "clip-path:polygon(53.06% 20.1%,68.47% 0%,67.25% 26.75%,83.37% 14.69%,75.83% 30.24%,100% 30.76%,78.63% 43.53%,84.58% 52.27%,75.83% 56.99%,87.39% 72.37%,67.78% 66.44%,69.18% 80.42%,56.39% 73.77%,53.76% 87.23%,45.7% 80.42%,40.28% 91.26%,34.85% 83.91%,22.76% 100%,22.25% 84.44%,5.95% 82.52%,15.42% 71.16%,0% 59.62%,18.22% 53.67%,5.43% 38.29%,24.87% 36.19%,20.84% 16.78%,39.58% 29.55%,45.01% 8.74%)";

        // R19 BUG A: parametric quads/polys honoring avLst (slant/notch/inset)
        if (preset == "parallelogram") return ParallelogramPolygon(presetGeom);
        if (preset == "trapezoid") return TrapezoidPolygon(presetGeom);
        if (preset == "chevron") return ChevronPolygon(presetGeom);
        if (preset == "hexagon") return HexagonPolygon(presetGeom);
        // triangle/isosTriangle: adj = apex horizontal position (fraction*100000,
        // default 50000 = centered). A non-default adj shifts the apex, making a
        // scalene triangle; the old hardcoded 50% ignored it.
        if (preset is "triangle" or "isosTriangle")
        {
            var apexPct = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000) / 1000.0, 0, 100);
            return $"clip-path:polygon({apexPct:0.##}% 0,100% 100%,0 100%)";
        }
        // diagStripe: a diagonal band (bottom-left→top-right). adj = stripe width as
        // a fraction of each axis (default 50000). The band edges run corner-to-corner
        // and parallel, so a pure-% polygon holds across aspect ratios. Was entirely
        // missing from the switch → rendered as a solid rectangle.
        if (preset == "diagStripe")
        {
            var adjPct = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000) / 1000.0, 0, 100);
            return $"clip-path:polygon({adjPct:0.##}% 0,100% 0,0 100%,0 {adjPct:0.##}%)";
        }
        // frame: adj1 = border thickness as a fraction of the shorter side (default
        // 12500 = 12.5%). The border is EQUAL absolute thickness on all sides, so the
        // per-axis percentage differs for non-square shapes; a large adj (e.g. 50000)
        // collapses the hole to a solid rectangle. The old hardcoded 12%/12% ignored adj.
        if (preset == "frame")
            return FramePolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "notchedRightArrow")
            return NotchedRightArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "upDownArrow" && widthEmu > 0 && heightEmu > 0)
            return UpDownArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "leftRightArrow" && widthEmu > 0 && heightEmu > 0)
            return LeftRightArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "halfFrame" && widthEmu > 0 && heightEmu > 0)
            return HalfFramePolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "octagon" && widthEmu > 0 && heightEmu > 0)
            return OctagonPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "moon")
            return MoonPolygon(presetGeom);
        if (preset == "mathPlus" && widthEmu > 0 && heightEmu > 0)
            return MathPlusCss(widthEmu, heightEmu, presetGeom);
        if (preset == "mathMultiply" && widthEmu > 0 && heightEmu > 0)
            return MathMultiplyCss(widthEmu, heightEmu, presetGeom);
        if (preset == "quadArrow" && widthEmu > 0 && heightEmu > 0)
            return QuadArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "rightArrowCallout" && widthEmu > 0 && heightEmu > 0)
            return RightArrowCalloutPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "leftArrowCallout" && widthEmu > 0 && heightEmu > 0)
            return LeftArrowCalloutPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "upArrowCallout" && widthEmu > 0 && heightEmu > 0)
            return UpArrowCalloutPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "downArrowCallout" && widthEmu > 0 && heightEmu > 0)
            return DownArrowCalloutPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "leftRightArrowCallout" && widthEmu > 0 && heightEmu > 0)
            return LeftRightArrowCalloutPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "upDownArrowCallout" && widthEmu > 0 && heightEmu > 0)
            return UpDownArrowCalloutPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "quadArrowCallout" && widthEmu > 0 && heightEmu > 0)
            return QuadArrowCalloutPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "bentUpArrow" && widthEmu > 0 && heightEmu > 0)
            return BentUpArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "leftUpArrow" && widthEmu > 0 && heightEmu > 0)
            return LeftUpArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "leftRightUpArrow" && widthEmu > 0 && heightEmu > 0)
            return LeftRightUpArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "nonIsoscelesTrapezoid" && widthEmu > 0 && heightEmu > 0)
            return NonIsoscelesTrapezoidPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "plaque" && widthEmu > 0 && heightEmu > 0)
            return PlaquePolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "pieWedge")
            return PieWedgeCss();
        if (preset == "teardrop" && widthEmu > 0 && heightEmu > 0)
            return TeardropPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "swooshArrow" && widthEmu > 0 && heightEmu > 0)
            return SwooshArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "uturnArrow" && widthEmu > 0 && heightEmu > 0)
            return UturnArrowPolygon(widthEmu, heightEmu, presetGeom);
        if (preset == "flowChartMagneticTape" && widthEmu > 0 && heightEmu > 0)
            return FlowChartMagneticTapeCss(widthEmu, heightEmu);
        // corner (L-shape): adj1 = bottom (horizontal) arm height %, adj2 = left
        // (vertical) arm width %; both default 50000. Inner corner at (adj2, 100-adj1).
        // The old hardcoded 50/50 ignored both, so a thin-armed L looked fat.
        // homePlate (pentagon arrow): adj = point depth as a fraction of width
        // (default 25000 = 25%). The body occupies (100-adj)%, the triangular point
        // the rest. The old hardcoded 75%/25% ignored adj.
        if (preset == "homePlate")
        {
            var ptPct = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 25000) / 1000.0, 0, 100);
            var bodyPct = 100 - ptPct;
            string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            return $"clip-path:polygon(0 0,{P(bodyPct)}% 0,100% 50%,{P(bodyPct)}% 100%,0 100%)";
        }
        if (preset == "corner")
        {
            var armH = Math.Clamp(ReadAdjValueCss(presetGeom, 0, 50000) / 1000.0, 0, 100);
            var armW = Math.Clamp(ReadAdjValueCss(presetGeom, 1, 50000) / 1000.0, 0, 100);
            string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            var innerY = 100 - armH;
            return $"clip-path:polygon(0 0,{P(armW)}% 0,{P(armW)}% {P(innerY)}%,100% {P(innerY)}%,100% 100%,0 100%)";
        }

        // Calculate roundRect corner radius from avLst or default (16.667% of shorter side)
        if (preset is "roundRect" or "round1Rect" or "round2SameRect" or "round2DiagRect")
        {
            var minSide = Math.Min(widthEmu, heightEmu);
            // The round2* presets carry TWO adjust guides: adj1 = first corner pair,
            // adj2 = second corner pair. Read both like the snip2* branch below;
            // previously only adj1 was read and the second pair was hardcoded to 0,
            // so adj2 was ignored and that pair always rendered square.
            var avList = presetGeom?.GetFirstChild<Drawing.AdjustValueList>();
            var gds = avList?.Elements<Drawing.ShapeGuide>().ToList() ?? new List<Drawing.ShapeGuide>();
            long ReadRadiusAdj(int i, long dflt)
            {
                if (i < gds.Count && gds[i].Formula?.Value is string f && f.StartsWith("val ")
                    && long.TryParse(f.AsSpan(4), out var v)) return v;
                return dflt;
            }
            // adj1 defaults to 16.667%; adj2 defaults to 0 (sharp) per the ECMA preset
            // definition, so a round2* with no adj2 is unchanged from the old output.
            long avVal = ReadRadiusAdj(0, 16667);
            long avVal2 = ReadRadiusAdj(1, 0);
            string Radius(long adj)
            {
                if (minSide <= 0) return "6pt"; // fallback if no dimensions
                return $"{Units.EmuToPt(minSide * adj / 100000):0.##}pt";
            }
            var r = Radius(avVal);
            var r2 = Radius(avVal2);

            // CSS border-radius order: top-left top-right bottom-right bottom-left.
            return preset switch
            {
                "roundRect" => $"border-radius:{r}",
                "round1Rect" => $"border-radius:0 {r} 0 0",  // PowerPoint rounds the top-right corner
                // round2SameRect: adj1 = top pair, adj2 = bottom pair.
                "round2SameRect" => $"border-radius:{r} {r} {r2} {r2}",
                // round2DiagRect: adj1 = TL+BR diagonal, adj2 = TR+BL diagonal.
                "round2DiagRect" => $"border-radius:{r} {r2} {r} {r2}",
                _ => ""
            };
        }

        // Parametric snip rectangles honoring avLst. The snipped corner size is
        // adj/100000 of the shorter side (matches roundRect's adj semantics).
        // Default adj for snip presets is 16667 (16.667%); the old code hardcoded
        // 8%/92% and ignored avLst entirely, so a custom adj never moved the
        // corner. Read adj1 (and adj2 for snip2Same/DiagRect) like roundRect does.
        if (preset is "snip1Rect" or "snip2SameRect" or "snip2DiagRect")
        {
            var avList = presetGeom?.GetFirstChild<Drawing.AdjustValueList>();
            var gds = avList?.Elements<Drawing.ShapeGuide>().ToList() ?? new List<Drawing.ShapeGuide>();
            long ReadAdj(int i, long dflt)
            {
                if (i < gds.Count && gds[i].Formula?.Value is string f && f.StartsWith("val ")
                    && long.TryParse(f.AsSpan(4), out var v)) return v;
                return dflt;
            }
            // adj is a fraction of the shorter side; convert to per-axis percent so
            // the snip is square (a 50% adj on a non-square box snips minSide/2 on
            // both axes). Clamp to [0,50] — a corner snip cannot exceed half a side.
            long minSideEmu = Math.Min(widthEmu, heightEmu);
            double AdjPct(long adj, bool horizontal)
            {
                var snipEmu = minSideEmu * adj / 100000.0;
                var axis = horizontal ? widthEmu : heightEmu;
                var pct = axis > 0 ? snipEmu / axis * 100.0 : adj / 100000.0 * 100.0;
                return Math.Clamp(pct, 0, 50);
            }
            var a1 = ReadAdj(0, 16667);
            var hx = AdjPct(a1, true);   // horizontal inset %
            var vy = AdjPct(a1, false);  // vertical inset %
            string P(double d) => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            switch (preset)
            {
                case "snip1Rect":
                    // top-right corner snipped
                    return $"clip-path:polygon(0 0,{P(100 - hx)}% 0,100% {P(vy)}%,100% 100%,0 100%)";
                case "snip2SameRect":
                {
                    // adj1 = top corner pair (TL+TR), adj2 = bottom corner pair (BL+BR).
                    // Previously only adj1 was read, so the bottom corners always
                    // rendered square. Read adj2 like the snip2DiagRect sibling and,
                    // when present, snip the bottom pair too (full 8-point octagon).
                    var a2 = ReadAdj(1, 0);
                    if (a2 == 0)
                        return $"clip-path:polygon({P(hx)}% 0,{P(100 - hx)}% 0,100% {P(vy)}%,100% 100%,0 100%,0 {P(vy)}%)";
                    var hx2 = AdjPct(a2, true);
                    var vy2 = AdjPct(a2, false);
                    return $"clip-path:polygon({P(hx)}% 0,{P(100 - hx)}% 0,100% {P(vy)}%,100% {P(100 - vy2)}%,{P(100 - hx2)}% 100%,{P(hx2)}% 100%,0 {P(100 - vy2)}%,0 {P(vy)}%)";
                }
                case "snip2DiagRect":
                {
                    // top-left + bottom-right snipped (diagonal)
                    var a2 = ReadAdj(1, 0);
                    var hx2 = AdjPct(a2 == 0 ? a1 : a2, true);
                    var vy2 = AdjPct(a2 == 0 ? a1 : a2, false);
                    return $"clip-path:polygon({P(hx)}% 0,100% 0,100% {P(100 - vy2)}%,{P(100 - hx2)}% 100%,0 100%,0 {P(vy)}%)";
                }
            }
        }

        return preset switch
        {
            // Rectangles
            "rect" => "",

            // Ellipses
            "ellipse" => "border-radius:50%",

            // Triangles
            "rtTriangle" => "clip-path:polygon(0 0,100% 100%,0 100%)",

            // Diamonds and parallelograms
            "diamond" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            // parallelogram/trapezoid handled above (parametric, honor avLst)

            // Polygons
            "pentagon" => "clip-path:polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%)",
            // hexagon handled above (parametric, honors avLst)
            // vertex-up regular heptagon, bbox-normalized (side vertices at y=64.3%, base 72.3%-27.7%)
            "heptagon" => "clip-path:polygon(50% 0,90.1% 19.8%,100% 64.3%,72.3% 100%,27.7% 100%,0 64.3%,9.9% 19.8%)",
            // edge-up regular decagon (10 vertices): flat top/bottom, pointed left/right
            "decagon" => "clip-path:polygon(35% 0,65% 0,90% 19%,100% 50%,90% 81%,65% 100%,35% 100%,10% 81%,0 50%,10% 19%)",
            "dodecagon" => "clip-path:polygon(37% 0,63% 0,87% 13%,100% 37%,100% 63%,87% 87%,63% 100%,37% 100%,13% 87%,0 63%,0 37%,13% 13%)",

            // Stars
            "star4" => "clip-path:polygon(50% 0,62% 38%,100% 50%,62% 62%,50% 100%,38% 62%,0 50%,38% 38%)",
            "star5" => "clip-path:polygon(50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)",
            "star6" => "clip-path:polygon(50% 0,63% 25%,100% 25%,75% 50%,100% 75%,63% 75%,50% 100%,37% 75%,0 75%,25% 50%,0 25%,37% 25%)",
            "star8" => "clip-path:polygon(50% 0,62% 19%,85% 15%,81% 38%,100% 50%,81% 62%,85% 85%,62% 81%,50% 100%,38% 81%,15% 85%,19% 62%,0 50%,19% 38%,15% 15%,38% 19%)",
            "star10" => "clip-path:polygon(50% 0,59% 19%,79% 5%,74% 27%,97% 25%,84% 43%,100% 50%,84% 57%,97% 75%,74% 73%,79% 95%,59% 81%,50% 100%,41% 81%,21% 95%,26% 73%,3% 75%,16% 57%,0 50%,16% 43%,3% 25%,26% 27%,21% 5%,41% 19%)",
            "star12" => "clip-path:polygon(50% 0,57% 15%,75% 7%,71% 25%,93% 25%,84% 42%,100% 50%,84% 58%,93% 75%,71% 75%,75% 93%,57% 85%,50% 100%,43% 85%,25% 93%,29% 75%,7% 75%,16% 58%,0 50%,16% 42%,7% 25%,29% 25%,25% 7%,43% 15%)",

            // Arrows
            "rightArrow" => "clip-path:polygon(0 20%,70% 20%,70% 0,100% 50%,70% 100%,70% 80%,0 80%)",
            "bentArrow" => "clip-path:polygon(0 20%,60% 20%,60% 0,100% 35%,60% 70%,60% 50%,20% 50%,20% 100%,0 100%)",
            "chevron" => "clip-path:polygon(0 0,80% 0,100% 50%,80% 100%,0 100%,20% 50%)",
            "stripedRightArrow" => "clip-path:polygon(10% 20%,12% 20%,12% 80%,10% 80%,10% 20%,15% 20%,70% 20%,70% 0,100% 50%,70% 100%,70% 80%,15% 80%)",

            // Callouts — rectangle/rounded-rect/ellipse body with a wedge tail pointing down-left

            // Crosses and plus — arm width scales with aspect ratio
            "plus" or "cross" => PlusPolygon(presetGeom),

            // Heart (polygon approximation)
            "heart" => "clip-path:polygon(50% 18%, 53% 12%, 57% 6%, 62% 2%, 68% 0%, 75% 0%, 82% 0%, 89% 3%, 94% 8%, 98% 14%, 100% 21%, 100% 28%, 99% 35%, 95% 43%, 90% 51%, 84% 59%, 77% 67%, 69% 75%, 60% 84%, 50% 100%, 40% 84%, 31% 75%, 23% 67%, 16% 59%, 10% 51%, 5% 43%, 1% 35%, 0% 28%, 0% 21%, 2% 14%, 6% 8%, 11% 3%, 18% 0%, 25% 0%, 32% 0%, 38% 2%, 43% 6%, 47% 12%)",

            // Cloud — SVG-based clip-path for realistic cloud bumps
            "cloud" => "clip-path:polygon(25% 80%,18% 80%,12% 78%,7% 74%,5% 69%,4% 64%,5% 60%,3% 56%,1% 51%,1% 47%,3% 42%,7% 38%,11% 36%,15% 35%,14% 29%,14% 23%,17% 19%,21% 16%,26% 15%,30% 15%,31% 10%,34% 6%,38% 3%,43% 1%,48% 0%,55% 5%,61% 2%,67% 1%,72% 2%,76% 6%,78% 15%,82% 12%,87% 11%,91% 13%,94% 17%,95% 22%,95% 30%,97% 33%,99% 37%,100% 42%,99% 47%,97% 52%,93% 55%,90% 55%,93% 59%,96% 64%,97% 68%,96% 73%,92% 76%,88% 78%,85% 78%,84% 82%,82% 87%,78% 90%,73% 92%,68% 92%,63% 90%,60% 90%,56% 93%,51% 96%,46% 97%,41% 96%,38% 93%,35% 90%)",
            "cloudCallout" => "clip-path:polygon(25% 80%,18% 80%,12% 78%,7% 74%,5% 69%,4% 64%,5% 60%,3% 56%,1% 51%,1% 47%,3% 42%,7% 38%,11% 36%,15% 35%,14% 29%,14% 23%,17% 19%,21% 16%,26% 15%,30% 15%,31% 10%,34% 6%,38% 3%,43% 1%,48% 0%,55% 5%,61% 2%,67% 1%,72% 2%,76% 6%,78% 15%,82% 12%,87% 11%,91% 13%,94% 17%,95% 22%,95% 30%,97% 33%,99% 37%,100% 42%,99% 47%,97% 52%,93% 55%,90% 55%,93% 59%,96% 64%,97% 68%,96% 73%,92% 76%,88% 78%,85% 78%,84% 82%,82% 87%,78% 90%,73% 92%,68% 92%,63% 90%,60% 90%,56% 93%,51% 96%,46% 97%,41% 96%,38% 93%,35% 90%)",

            // Smiley (circle)
            "smileyFace" or "smiley" => "border-radius:50%",

            // Sun — circle with triangular rays
            "sun" => "clip-path:polygon(50% 0,56% 15%,70% 3%,66% 19%,85% 15%,74% 27%,93% 30%,80% 38%,97% 45%,82% 48%,97% 55%,80% 62%,93% 70%,74% 73%,85% 85%,66% 81%,70% 97%,56% 85%,50% 100%,44% 85%,30% 97%,34% 81%,15% 85%,26% 73%,7% 70%,20% 62%,3% 55%,18% 48%,3% 45%,20% 38%,7% 30%,26% 27%,15% 15%,34% 19%,30% 3%,44% 15%)",

            // Moon (crescent) — outer arc minus inner arc
            "moon" => MoonPolygon(presetGeom),

            // Gear (polygon approximation of 6-tooth gear)
            "gear6" => "clip-path:polygon(50% 0,61% 10%,75% 3%,80% 18%,97% 25%,88% 38%,100% 50%,88% 62%,97% 75%,80% 82%,75% 97%,61% 90%,50% 100%,39% 90%,25% 97%,20% 82%,3% 75%,12% 62%,0 50%,12% 38%,3% 25%,20% 18%,25% 3%,39% 10%)",
            "gear9" => "clip-path:polygon(50% 0,56% 8%,65% 2%,68% 12%,78% 9%,78% 20%,88% 20%,85% 30%,95% 35%,90% 44%,100% 50%,90% 56%,95% 65%,85% 70%,88% 80%,78% 80%,78% 91%,68% 88%,65% 98%,56% 92%,50% 100%,44% 92%,35% 98%,32% 88%,22% 91%,22% 80%,12% 80%,15% 70%,5% 65%,10% 56%,0 50%,10% 44%,5% 35%,15% 30%,12% 20%,22% 20%,22% 9%,32% 12%,35% 2%,44% 8%)",

            // 3D-like shapes (rendered flat)
            "cube" => "clip-path:polygon(10% 0,100% 0,100% 85%,90% 100%,0 100%,0 15%)",
            "can" or "cylinder" => "border-radius:50%/10%",
            "bevel" => "border:3px outset currentColor",
            "foldedCorner" => "clip-path:polygon(0 0,85% 0,100% 15%,100% 100%,0 100%)",
            "lightningBolt" => "clip-path:polygon(35% 0,55% 35%,100% 30%,45% 55%,80% 100%,25% 60%,0 80%,30% 45%)",

            // Misc shapes
            "donut" => DonutCss(presetGeom),
            "noSmoking" => "border-radius:50%",
            // pie/arc/chord/blockArc/snipRoundRect handled above (parametric)

            // Ribbons/banners
            "ribbon" or "ribbon2" or "wave" or "doubleWave" => "",
            "horizontalScroll" or "verticalScroll" => "border-radius:4px",

            // Flowchart
            "flowChartProcess" => "",
            "flowChartAlternateProcess" => "border-radius:8px",
            "flowChartDecision" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            "flowChartInputOutput" or "flowChartData" => "clip-path:polygon(20% 0,100% 0,80% 100%,0 100%)",
            "flowChartPredefinedProcess" => "border-left:3px double currentColor;border-right:3px double currentColor",
            "flowChartDocument" => "",
            "flowChartMultidocument" => "",
            "flowChartTerminator" => "border-radius:50%/100%",
            "flowChartPreparation" => "clip-path:polygon(17% 0,83% 0,100% 50%,83% 100%,17% 100%,0 50%)",
            "flowChartManualInput" => "clip-path:polygon(0 20%,100% 0,100% 100%,0 100%)",
            "flowChartManualOperation" => "clip-path:polygon(0 0,100% 0,85% 100%,15% 100%)",
            "flowChartMerge" => "clip-path:polygon(0 0,100% 0,50% 100%)",
            "flowChartExtract" => "clip-path:polygon(50% 0,100% 100%,0 100%)",
            "flowChartSort" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            "flowChartCollate" => "clip-path:polygon(0 0,100% 0,50% 50%,100% 100%,0 100%,50% 50%)",
            "flowChartDelay" => "border-radius:0 50% 50% 0",
            "flowChartDisplay" => "clip-path:polygon(0 50%,15% 0,85% 0,100% 50%,85% 100%,15% 100%)",
            "flowChartPunchedCard" => "clip-path:polygon(15% 0,100% 0,100% 100%,0 100%,0 15%)",
            "flowChartPunchedTape" => "",
            "flowChartOnlineStorage" => "border-radius:50% 0 0 50%",
            "flowChartOfflineStorage" => "clip-path:polygon(0 0,100% 0,50% 100%)",
            "flowChartMagneticDisk" => "border-radius:50%/20%",
            "flowChartConnector" or "flowChartOffpageConnector" => "border-radius:50%",

            // Block arrows (curved)
            "curvedRightArrow" => "clip-path:polygon(0% 85%,0% 55%,2% 40%,6% 28%,12% 19%,20% 13%,30% 10%,70% 10%,70% 0%,100% 20%,70% 40%,70% 30%,40% 30%,32% 33%,26% 38%,22% 45%,20% 55%,20% 85%)",
            "curvedLeftArrow" => "clip-path:polygon(100% 85%,100% 55%,98% 40%,94% 28%,88% 19%,80% 13%,70% 10%,30% 10%,30% 0%,0% 20%,30% 40%,30% 30%,60% 30%,68% 33%,74% 38%,78% 45%,80% 55%,80% 85%)",
            "curvedUpArrow" => "clip-path:polygon(85% 100%,55% 100%,40% 98%,28% 94%,19% 88%,13% 80%,10% 70%,10% 30%,0% 30%,20% 0%,40% 30%,30% 30%,30% 60%,33% 68%,38% 74%,45% 78%,55% 80%,85% 80%)",
            "curvedDownArrow" => "clip-path:polygon(85% 0%,55% 0%,40% 2%,28% 6%,19% 12%,13% 20%,10% 30%,10% 70%,0% 70%,20% 100%,40% 70%,30% 70%,30% 40%,33% 32%,38% 26%,45% 22%,55% 20%,85% 20%)",
            // circular-arrow family: rendered as a circle (border-radius) — a crude but
            // CONSISTENT approximation. The faithful annular-sector-with-arrowhead geometry
            // (5 adjusts, 200+ guides) is a deferred focused task; a circle is far closer to
            // the true silhouette than the bare-rectangle fallback the left variants had.
            "circularArrow" or "leftCircularArrow" or "leftRightCircularArrow" => "border-radius:50%",

            // Math
            "mathPlus" => "clip-path:polygon(33% 0,67% 0,67% 33%,100% 33%,100% 67%,67% 67%,67% 100%,33% 100%,33% 67%,0 67%,0 33%,33% 33%)",
            "mathMinus" => MathMinusCss(presetGeom),
            "mathMultiply" => "clip-path:polygon(20% 0,50% 30%,80% 0,100% 20%,70% 50%,100% 80%,80% 100%,50% 70%,20% 100%,0 80%,30% 50%,0 20%)",
            "mathDivide" => "",
            "mathEqual" => "clip-path:polygon(0 25%,100% 25%,100% 40%,0 40%,0 60%,100% 60%,100% 75%,0 75%)",
            "mathNotEqual" => "",

            // Default: render as rectangle
            _ => ""
        };
    }

    // ==================== Color Resolution ====================

    private static string? ResolveFillColor(Drawing.SolidFill? solidFill, Dictionary<string, string> themeColors)
    {
        if (solidFill == null) return null;

        var rgb = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
        {
            var hexPart = rgb[..6]; // Only use first 6 hex chars, ignore any trailing data
            var rgbEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()!;
            // Apply lumMod/lumOff/tint/shade if present (same transforms as schemeClr).
            var transformed = ApplyRgbColorTransforms(hexPart, rgbEl);
            hexPart = transformed.StartsWith('#') ? transformed[1..] : transformed;
            var alpha = rgbEl.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (alpha.HasValue && alpha.Value < 100000)
            {
                var (r, g, b) = ColorMath.HexToRgb(hexPart);
                return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";
            }
            return $"#{hexPart}";
        }

        var schemeColor = solidFill.GetFirstChild<Drawing.SchemeColor>();
        if (schemeColor?.Val?.HasValue == true)
        {
            var schemeName = schemeColor.Val!.InnerText;
            if (schemeName != null && themeColors.TryGetValue(schemeName, out var themeHex))
            {
                // Check for lumMod/lumOff/tint/shade transforms
                var color = ApplyColorTransforms(themeHex, schemeColor);
                return color;
            }
            return null; // Unknown scheme color
        }

        // Preset/named color (<a:prstClr val="red"/>): a valid CT_Color choice that other
        // tools (and hand-authored files) use. ResolveFillColor handled only srgbClr and
        // schemeClr, so a prstClr fill fell through to null and the shape/run rendered with
        // no fill (PowerPoint paints the named color). OOXML preset names are lowercase CSS
        // color names, which TryGetNamedColorHex maps to hex.
        var prstColor = solidFill.GetFirstChild<Drawing.PresetColor>();
        if (prstColor?.Val?.HasValue == true)
        {
            var hex = ParseHelpers.TryGetNamedColorHex(prstColor.Val!.InnerText);
            if (hex != null)
            {
                // Alpha is folded in by the caller below (same as the srgbClr
                // branch). ApplyColorTransforms folds it too and returns an
                // rgba() string, which the hex parse below then rejected —
                // one translucent preset color aborted the whole HTML render.
                var transformed = ApplyOpaqueColorTransforms(hex, prstColor);
                var solid = transformed.StartsWith('#') ? transformed[1..] : transformed;
                var alpha = prstColor.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
                if (alpha.HasValue && alpha.Value < 100000)
                {
                    var (r, g, b) = ColorMath.HexToRgb(solid);
                    return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";
                }
                return $"#{solid}";
            }
        }

        // System color (<a:sysClr>): resolve via lastClr/standard slot, then transforms+alpha.
        var sysColor = solidFill.GetFirstChild<Drawing.SystemColor>();
        if (sysColor != null)
        {
            var sysHex = SysColorHex(sysColor);
            if (sysHex != null)
            {
                var transformed = ApplyOpaqueColorTransforms(sysHex, sysColor);
                var solid = transformed.StartsWith('#') ? transformed[1..] : transformed;
                var alpha = sysColor.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
                if (alpha.HasValue && alpha.Value < 100000)
                {
                    var (r, g, b) = ColorMath.HexToRgb(solid);
                    return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";
                }
                return $"#{solid}";
            }
        }

        return null;
    }

    /// <summary>
    /// Resolve a color from a <c>&lt;a:buClr&gt;</c> (CT_Color) element, whose color
    /// child (<c>srgbClr</c>/<c>schemeClr</c>) sits directly inside it rather than
    /// being wrapped in <c>&lt;a:solidFill&gt;</c>. Honours lumMod/lumOff/tint/shade/alpha
    /// transforms via the same helpers as ResolveFillColor.
    /// </summary>
    private static string? ResolveBulletColor(Drawing.BulletColor? buClr, Dictionary<string, string> themeColors)
    {
        if (buClr == null) return null;

        var rgbEl = buClr.GetFirstChild<Drawing.RgbColorModelHex>();
        var rgb = rgbEl?.Val?.Value;
        if (rgbEl != null && rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
        {
            var hexPart = rgb[..6];
            var transformed = ApplyRgbColorTransforms(hexPart, rgbEl);
            hexPart = transformed.StartsWith('#') ? transformed[1..] : transformed;
            var alpha = rgbEl.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (alpha.HasValue && alpha.Value < 100000)
            {
                var (r, g, b) = ColorMath.HexToRgb(hexPart);
                return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";
            }
            return $"#{hexPart}";
        }

        var schemeColor = buClr.GetFirstChild<Drawing.SchemeColor>();
        if (schemeColor?.Val?.HasValue == true)
        {
            var schemeName = schemeColor.Val!.InnerText;
            if (schemeName != null && themeColors.TryGetValue(schemeName, out var themeHex))
                return ApplyColorTransforms(themeHex, schemeColor);
        }

        // Preset/named bullet color (<a:buClr><a:prstClr val="red"/>): was dropped to null,
        // so the bullet glyph inherited the text color instead of its named color.
        var prstColor = buClr.GetFirstChild<Drawing.PresetColor>();
        if (prstColor?.Val?.HasValue == true)
        {
            var hex = ParseHelpers.TryGetNamedColorHex(prstColor.Val!.InnerText);
            if (hex != null) return ApplyColorTransforms(hex, prstColor);
        }

        var sysColor = buClr.GetFirstChild<Drawing.SystemColor>();
        if (sysColor != null && SysColorHex(sysColor) is string sh)
            return ApplyColorTransforms(sh, sysColor);

        return null;
    }

    // System color (<a:sysClr val="windowText" lastClr="RRGGBB"/>): a renderer that cannot
    // query the OS uses the cached lastClr (what PowerPoint last rendered); for the two common
    // system slots fall back to their standard values when lastClr is absent. Returns bare hex.
    private static string? SysColorHex(Drawing.SystemColor sys)
    {
        var last = sys.LastColor?.Value;
        if (last != null && last.Length == 6 && last.All(char.IsAsciiHexDigit)) return last;
        return sys.Val?.InnerText switch { "windowText" => "000000", "window" => "FFFFFF", _ => null };
    }

    private static string ApplyColorTransforms(string hex, Drawing.SchemeColor schemeColor)
        => ApplyColorTransforms(hex, (OpenXmlElement)schemeColor);

    /// <summary>
    /// <see cref="ApplyColorTransforms(string, OpenXmlElement)"/> minus alpha:
    /// for callers that fold alpha in themselves and need a #RRGGBB back.
    /// </summary>
    private static string ApplyOpaqueColorTransforms(string hex, OpenXmlElement colorEl)
    {
        return ColorMath.ApplyTransforms(hex,
            tint: ReadTransformVal(colorEl, "tint"),
            shade: ReadTransformVal(colorEl, "shade"),
            lumMod: ReadTransformVal(colorEl, "lumMod"),
            lumOff: ReadTransformVal(colorEl, "lumOff"),
            satMod: ReadTransformVal(colorEl, "satMod"));
    }

    /// <summary>
    /// Apply lumMod/lumOff/tint/shade/alpha child transforms on any color element
    /// (schemeClr or srgbClr). Reads by local name so it works for both strongly-typed
    /// children (round-tripped from disk) AND OpenXmlUnknownElement children (built
    /// in-memory by DrawingColorBuilder, which appends transforms as unknown elements).
    /// </summary>
    private static string ApplyColorTransforms(string hex, OpenXmlElement colorEl)
    {
        return ColorMath.ApplyTransforms(hex,
            tint: ReadTransformVal(colorEl, "tint"),
            shade: ReadTransformVal(colorEl, "shade"),
            lumMod: ReadTransformVal(colorEl, "lumMod"),
            lumOff: ReadTransformVal(colorEl, "lumOff"),
            alpha: ReadTransformVal(colorEl, "alpha"),
            satMod: ReadTransformVal(colorEl, "satMod"));
    }

    private static int? ReadTransformVal(OpenXmlElement colorEl, string localName)
    {
        foreach (var child in colorEl.ChildElements)
        {
            if (!child.LocalName.Equals(localName, StringComparison.Ordinal)) continue;
            var raw = child.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (int.TryParse(raw, out var v)) return v;
        }
        return null;
    }

    /// <summary>
    /// Build a map of scheme color names to hex values from the presentation theme.
    /// </summary>
    private Dictionary<string, string> ResolveThemeColorMap()
    {
        // Use the same theme-part resolution as Get/Set (GetThemePart prefers the
        // presentationPart's own theme, then the first master's) so a Set hlink/
        // accent color is reflected here instead of reading a stale master theme.
        var colorScheme = GetColorScheme()
            ?? _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault()
                ?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        return ThemeColorResolver.BuildColorMap(colorScheme, includePptAliases: true);
    }

    /// <summary>
    /// R9-2: apply a slide's &lt;p:clrMapOvr&gt;&lt;p:overrideClrMapping&gt; on top of
    /// the global theme color map. The override remaps the 12 color slots
    /// (e.g. accent1="accent2" makes the accent1 slot resolve to accent2's hex).
    /// Returns the base map unchanged when the slide carries no override mapping.
    /// </summary>
    private static Dictionary<string, string> ApplySlideColorMapOverride(
        SlidePart slidePart, Dictionary<string, string> baseMap)
    {
        var clrMapOvr = slidePart.Slide?.GetFirstChild<ColorMapOverride>();
        // The overrideClrMapping element may be authored under either the a: or p:
        // namespace (real PowerPoint emits a:overrideClrMapping); read attributes
        // generically by local name to be namespace-agnostic. A masterClrMapping
        // child (no attributes) means "inherit" — nothing to remap.
        var ovr = clrMapOvr?.ChildElements
            .FirstOrDefault(e => e.LocalName == "overrideClrMapping");
        if (ovr == null || !ovr.HasAttributes) return baseMap;

        var remapped = new Dictionary<string, string>(baseMap, StringComparer.OrdinalIgnoreCase);
        // Each attribute (bg1/tx1/bg2/tx2/accent1..6/hlink/folHlink) names a slot
        // and its value is the scheme token it now points at (e.g. accent1="accent2"
        // makes the accent1 slot resolve to accent2's hex).
        foreach (var attr in ovr.GetAttributes())
        {
            var slot = attr.LocalName;
            var token = attr.Value;
            if (string.IsNullOrEmpty(token)) continue;
            if (baseMap.TryGetValue(token!, out var hex))
                remapped[slot] = hex;
        }
        return remapped;
    }

    // ==================== Image Helpers ====================

    private static string? BlipToDataUri(Drawing.BlipFill blipFill, OpenXmlPart part)
    {
        var blip = blipFill.GetFirstChild<Drawing.Blip>();
        if (blip == null) return null;
        // R4-1: an SVG picture stores BOTH a raster fallback (blip.Embed → 1x1
        // PNG) AND the real vector via the asvg:svgBlip extension. Prefer the SVG
        // rel-id so the preview shows the actual artwork, not the 1x1 fallback.
        var svgRelId = OfficeCli.Core.SvgImageHelper.GetSvgRelId(blip);
        if (!string.IsNullOrEmpty(svgRelId))
        {
            try
            {
                var svgPart = part.GetPartById(svgRelId!);
                using var stream = svgPart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                return $"data:image/svg+xml;base64,{Convert.ToBase64String(ms.ToArray())}";
            }
            catch { /* unresolved SVG rel — fall back to the raster embed */ }
        }
        if (blip.Embed?.HasValue != true) return null;
        return HtmlPreviewHelper.PartToDataUri(part, blip.Embed.Value!);
    }

    // R49-01: decode a blip's natural pixel dimensions (PNG/JPEG/GIF/BMP) so
    // tile sx/sy scale can be expressed as image-relative px in CSS.
    private static (int Width, int Height)? TryGetBlipDimensions(Drawing.BlipFill blipFill, OpenXmlPart part)
    {
        var blip = blipFill.GetFirstChild<Drawing.Blip>();
        if (blip?.Embed?.HasValue != true) return null;
        try
        {
            var imgPart = part.GetPartById(blip.Embed.Value!);
            using var stream = imgPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            ms.Position = 0;
            return OfficeCli.Core.ImageSource.TryGetDimensions(ms);
        }
        catch { return null; }
    }

    // R49-01: map <a:tile algn> to a CSS background-position keyword pair.
    private static string TileAlignToBackgroundPosition(DocumentFormat.OpenXml.EnumValue<Drawing.RectangleAlignmentValues>? algn)
    {
        var v = algn?.Value;
        if (v == null) return "left top";
        if (v == Drawing.RectangleAlignmentValues.TopLeft) return "left top";
        if (v == Drawing.RectangleAlignmentValues.Top) return "center top";
        if (v == Drawing.RectangleAlignmentValues.TopRight) return "right top";
        if (v == Drawing.RectangleAlignmentValues.Left) return "left center";
        if (v == Drawing.RectangleAlignmentValues.Center) return "center";
        if (v == Drawing.RectangleAlignmentValues.Right) return "right center";
        if (v == Drawing.RectangleAlignmentValues.BottomLeft) return "left bottom";
        if (v == Drawing.RectangleAlignmentValues.Bottom) return "center bottom";
        if (v == Drawing.RectangleAlignmentValues.BottomRight) return "right bottom";
        return "left top";
    }

    // ==================== Utility ====================

    // Unit conversions moved to shared Units class (Core/Units.cs).

    // CONSISTENCY(html-encode): shared plain entity-encoder lives in Core/HtmlPreviewHelper.
    private static string HtmlEncode(string text) => HtmlPreviewHelper.HtmlEncode(text);

    /// <summary>
    /// Sanitize a value for use inside a CSS style attribute.
    /// Strips characters that could break out of the style context.
    /// </summary>
    private static readonly string[] CjkFallbacks = { "PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC", "Hiragino Sans GB" };

    private static string CssFontFamilyWithFallback(string font)
    {
        var sanitized = CssSanitize(font);
        var fallbacks = string.Join(",", CjkFallbacks
            .Where(f => !f.Equals(font, StringComparison.OrdinalIgnoreCase))
            .Select(f => $"'{f}'"));
        return $"font-family:'{sanitized}',{fallbacks},sans-serif";
    }

    /// <summary>
    /// Returns true if the hex color is dark (low luminance).
    /// </summary>
    private static bool IsColorDark(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length < 6) return false;
        var (r, g, b) = ColorMath.HexToRgb(hex);
        // Relative luminance approximation
        return (r * 0.299 + g * 0.587 + b * 0.114) < 128;
    }

    private static string CssSanitize(string value)
    {
        // Remove characters that could escape the style attribute or inject HTML
        return value.Replace("\"", "").Replace("'", "").Replace("<", "").Replace(">", "")
            .Replace(";", "").Replace("{", "").Replace("}", "");
    }

    /// <summary>
    /// Sanitize a color value for safe embedding in CSS.
    /// Only allows hex colors (#RRGGBB), rgb/rgba() functions, and named CSS colors.
    /// </summary>
    private static string CssSanitizeColor(string color)
    {
        if (string.IsNullOrEmpty(color)) return "transparent";
        // Allow: #hex, rgb(), rgba(), named colors (alphanumeric only)
        var trimmed = color.Trim();
        if (trimmed.StartsWith('#') && trimmed.Length <= 9 && trimmed[1..].All(char.IsAsciiHexDigit))
            return trimmed;
        if (trimmed.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
            return CssSanitize(trimmed);
        if (trimmed.All(c => char.IsLetterOrDigit(c) || c == '.'))
            return trimmed;
        return "transparent";
    }

    /// <summary>
    /// Sanitize a MIME content type for safe embedding in a data URI.
    /// </summary>
    private static string SanitizeContentType(string contentType)
    {
        if (string.IsNullOrEmpty(contentType)) return "image/png";
        // Only allow alphanumeric, '/', '+', '-', '.'
        if (contentType.All(c => char.IsLetterOrDigit(c) || c is '/' or '+' or '-' or '.'))
            return contentType;
        return "image/png";
    }
}
