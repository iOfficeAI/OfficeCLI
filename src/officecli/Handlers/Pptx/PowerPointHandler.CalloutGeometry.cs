// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

/// <summary>
/// Wedge callout outlines (wedgeRectCallout / wedgeRoundRectCallout /
/// wedgeEllipseCallout) computed from the ECMA-376 preset definitions, so the
/// tail follows adj1/adj2 (tail tip offset from the centre, as a fraction of
/// width/height ×100000) instead of a fixed bottom-left wedge. Shared by the
/// HTML and SVG previews. The tip may lie outside the w×h box.
/// </summary>
public partial class PowerPointHandler
{
    internal static bool IsWedgeCallout(string preset) =>
        preset is "wedgeRectCallout" or "wedgeRoundRectCallout" or "wedgeEllipseCallout";

    /// <summary>Adjust value by guide name (adj1/adj2/…); falls back to
    /// <paramref name="defaultValue"/> when absent or not a plain `val N`.</summary>
    private static long ReadNamedAdj(Drawing.PresetGeometry? presetGeom, string name, long defaultValue)
    {
        var avList = presetGeom?.GetFirstChild<Drawing.AdjustValueList>();
        if (avList == null) return defaultValue;
        foreach (var gd in avList.Elements<Drawing.ShapeGuide>())
        {
            if (!string.Equals(gd.Name?.Value, name, StringComparison.Ordinal)) continue;
            var f = gd.Formula?.Value;
            if (f != null && f.StartsWith("val ") && long.TryParse(f.AsSpan(4), out var v)) return v;
            return defaultValue;
        }
        return defaultValue;
    }

    /// <summary>Outline vertices of a wedge callout in a w×h coordinate space,
    /// or null when <paramref name="preset"/> is not a wedge callout.</summary>
    internal static List<(double X, double Y)>? WedgeCalloutOutline(
        string preset, double w, double h, Drawing.PresetGeometry? presetGeom)
    {
        if (!IsWedgeCallout(preset) || w <= 0 || h <= 0) return null;

        double adj1 = ReadNamedAdj(presetGeom, "adj1", -20833);
        double adj2 = ReadNamedAdj(presetGeom, "adj2", 62500);
        double hc = w / 2, vc = h / 2;
        double dxPos = w * adj1 / 100000.0;
        double dyPos = h * adj2 / 100000.0;
        double xPos = hc + dxPos, yPos = vc + dyPos;

        if (preset == "wedgeEllipseCallout")
        {
            // pang = at2(dxPos*h, dyPos*w); tail base at pang ± 11° on the ellipse.
            double pang = Math.Atan2(dyPos * w, dxPos * h);
            double half = 11.0 * Math.PI / 180.0;
            double st = pang + half;
            var pts = new List<(double, double)> { (xPos, yPos) };
            const int steps = 48;
            double sweep = 2 * Math.PI - 2 * half;
            for (int i = 0; i <= steps; i++)
            {
                double a = st + sweep * i / steps;
                pts.Add((hc + w / 2 * Math.Cos(a), vc + h / 2 * Math.Sin(a)));
            }
            return pts;
        }

        // Rect / round-rect: pick the side the tail leaves from.
        double dq = dxPos * h / w;
        double dz = Math.Abs(dyPos) - Math.Abs(dq); // > 0 → tail on top/bottom
        double x1 = w * (dxPos > 0 ? 7 : 2) / 12.0, x2 = w * (dxPos > 0 ? 10 : 5) / 12.0;
        double y1 = h * (dyPos > 0 ? 7 : 2) / 12.0, y2 = h * (dyPos > 0 ? 10 : 5) / 12.0;

        double xl = dz > 0 ? 0 : (dxPos > 0 ? 0 : xPos);
        double yl = dz > 0 ? y1 : (dxPos > 0 ? y1 : yPos);
        double xt = dz > 0 ? (dyPos > 0 ? x1 : xPos) : x1;
        double yt = dz > 0 ? (dyPos > 0 ? 0 : yPos) : 0;
        double xr = dz > 0 ? w : (dxPos > 0 ? xPos : w);
        double yr = dz > 0 ? y1 : (dxPos > 0 ? yPos : y1);
        double xb = dz > 0 ? (dyPos > 0 ? xPos : x1) : x1;
        double yb = dz > 0 ? (dyPos > 0 ? yPos : h) : h;

        double u = 0;
        if (preset == "wedgeRoundRectCallout")
            u = Math.Min(w, h) * Math.Clamp(ReadNamedAdj(presetGeom, "adj3", 16667), 0, 50000) / 100000.0;

        var o = new List<(double, double)>();
        void Corner(double cx, double cy, double startDeg)
        {
            if (u <= 0) { o.Add((cx, cy)); return; }
            for (int i = 0; i <= 6; i++)
            {
                double a = (startDeg + 90.0 * i / 6) * Math.PI / 180.0;
                o.Add((cx + u * Math.Cos(a), cy + u * Math.Sin(a)));
            }
        }
        // top edge (left → right)
        Corner(u, u, 180);
        o.Add((x1, 0)); o.Add((xt, yt)); o.Add((x2, 0));
        Corner(w - u, u, 270);
        // right edge (top → bottom)
        o.Add((w, y1)); o.Add((xr, yr)); o.Add((w, y2));
        Corner(w - u, h - u, 0);
        // bottom edge (right → left)
        o.Add((x2, h)); o.Add((xb, yb)); o.Add((x1, h));
        Corner(u, h - u, 90);
        // left edge (bottom → top)
        o.Add((0, y2)); o.Add((xl, yl)); o.Add((0, y1));
        return o;
    }

    /// <summary>SVG path `d` for the callout outline (closed).</summary>
    internal static string? WedgeCalloutSvgPath(string preset, double w, double h, Drawing.PresetGeometry? presetGeom)
    {
        var pts = WedgeCalloutOutline(preset, w, h, presetGeom);
        if (pts == null) return null;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < pts.Count; i++)
            sb.Append(i == 0 ? "M" : " L").Append(pts[i].X.ToString("0.##", ci)).Append(',').Append(pts[i].Y.ToString("0.##", ci));
        return sb.Append(" Z").ToString();
    }

    /// <summary>CSS clip-path polygon (percent of the box) for the callout. Points
    /// outside 0–100% are kept; the box clips them, which is the best a clip-path
    /// can do (shapes route through the SVG overlay instead so the tail shows).</summary>
    internal static string? WedgeCalloutClipPath(string preset, long wEmu, long hEmu, Drawing.PresetGeometry? presetGeom)
    {
        double w = wEmu > 0 ? wEmu : 100000, h = hEmu > 0 ? hEmu : 100000;
        var pts = WedgeCalloutOutline(preset, w, h, presetGeom);
        if (pts == null) return null;
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        return "clip-path:polygon(" + string.Join(",", pts.Select(p =>
            (p.X / w * 100).ToString("0.##", ci) + "% " + (p.Y / h * 100).ToString("0.##", ci) + "%")) + ")";
    }
}
