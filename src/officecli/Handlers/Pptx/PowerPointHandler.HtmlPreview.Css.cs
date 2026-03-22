// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== CSS Helper: Fill ====================

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
                return $"background:url('{dataUri}') center/cover no-repeat";
        }

        return "";
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
            if (child is Drawing.CubicBezierCurveTo or Drawing.QuadraticBezierCurveTo)
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
            // Has curves — use clip-path: path() with SVG path syntax
            var svgPath = new StringBuilder();
            foreach (var child in path.ChildElements)
            {
                switch (child)
                {
                    case Drawing.MoveTo moveTo:
                        if (TryParsePoint(moveTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var mx, out var my))
                            svgPath.Append($"M {mx:0.##} {my:0.##} ");
                        break;
                    case Drawing.LineTo lineTo:
                        if (TryParsePoint(lineTo.GetFirstChild<Drawing.Point>(), pathW, pathH, out var lx, out var ly))
                            svgPath.Append($"L {lx:0.##} {ly:0.##} ");
                        break;
                    case Drawing.CubicBezierCurveTo cubicBez:
                    {
                        var pts = cubicBez.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 3
                            && TryParsePoint(pts[0], pathW, pathH, out var c1x, out var c1y)
                            && TryParsePoint(pts[1], pathW, pathH, out var c2x, out var c2y)
                            && TryParsePoint(pts[2], pathW, pathH, out var c3x, out var c3y))
                            svgPath.Append($"C {c1x:0.##} {c1y:0.##},{c2x:0.##} {c2y:0.##},{c3x:0.##} {c3y:0.##} ");
                        break;
                    }
                    case Drawing.QuadraticBezierCurveTo quadBez:
                    {
                        var pts = quadBez.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 2
                            && TryParsePoint(pts[0], pathW, pathH, out var q1x, out var q1y)
                            && TryParsePoint(pts[1], pathW, pathH, out var q2x, out var q2y))
                            svgPath.Append($"Q {q1x:0.##} {q1y:0.##},{q2x:0.##} {q2y:0.##} ");
                        break;
                    }
                    case Drawing.CloseShapePath:
                        svgPath.Append("Z ");
                        break;
                }
            }
            var pathStr = svgPath.ToString().Trim();
            if (!string.IsNullOrEmpty(pathStr))
                return $"clip-path:path('{pathStr}')";
        }

        return "";
    }

    // ==================== CSS Helper: Gradient ====================

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
                // Try direct color children
                var rgb = gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (rgb != null && rgb.Length >= 6 && rgb[..6].All(char.IsAsciiHexDigit))
                    color = $"#{rgb[..6]}";
                else
                {
                    var scheme = gs.GetFirstChild<Drawing.SchemeColor>()?.Val?.InnerText;
                    color = scheme != null && themeColors.TryGetValue(scheme, out var tc) ? $"#{tc}" : "transparent";
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
            return $"radial-gradient(circle, {string.Join(", ", cssStops)})";

        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        var angleDeg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 90.0;
        // OOXML angle 0° = top→bottom (same as CSS 180deg), so CSS angle = OOXML + 90°
        // Actually OOXML: 0 = right, 90 = bottom; CSS: 0 = up, 90 = right
        var cssAngle = angleDeg + 90;

        return $"linear-gradient({cssAngle:0.##}deg, {string.Join(", ", cssStops)})";
    }

    // ==================== CSS Helper: Outline/Border ====================

    private static string OutlineToCss(Drawing.Outline outline, Dictionary<string, string> themeColors)
    {
        if (outline.GetFirstChild<Drawing.NoFill>() != null) return "";

        var color = ResolveFillColor(outline.GetFirstChild<Drawing.SolidFill>(), themeColors) ?? "#000000";
        var widthPt = outline.Width?.HasValue == true ? outline.Width.Value / 12700.0 : 1.0;
        if (widthPt < 0.5) widthPt = 0.5;

        var dash = outline.GetFirstChild<Drawing.PresetDash>();
        var borderStyle = "solid";
        if (dash?.Val?.HasValue == true)
        {
            borderStyle = dash.Val.InnerText switch
            {
                "dash" or "lgDash" or "sysDash" => "dashed",
                "dot" or "sysDot" => "dotted",
                "dashDot" or "lgDashDot" or "sysDashDot" or "sysDashDotDot" => "dashed",
                _ => "solid"
            };
        }

        return $"border:{widthPt:0.##}pt {borderStyle} {color}";
    }

    // ==================== CSS Helper: Shadow ====================

    private static string EffectListToShadowCss(Drawing.EffectList? effectList, Dictionary<string, string> themeColors)
    {
        if (effectList == null) return "";

        var shadow = effectList.GetFirstChild<Drawing.OuterShadow>();
        if (shadow == null) return "";

        var alpha = shadow.Descendants<Drawing.Alpha>().FirstOrDefault()?.Val?.Value ?? 50000;
        var opacity = alpha / 100000.0;
        var rgb = shadow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        string color;
        if (rgb != null)
        {
            var r = Convert.ToInt32(rgb[..2], 16);
            var g = Convert.ToInt32(rgb[2..4], 16);
            var b = Convert.ToInt32(rgb[4..6], 16);
            color = $"rgba({r},{g},{b},{opacity:0.##})";
        }
        else
        {
            // Try scheme color
            var schemeColor = shadow.GetFirstChild<Drawing.SchemeColor>()?.Val?.InnerText;
            var resolved = schemeColor != null && themeColors.TryGetValue(schemeColor, out var sc) ? sc : null;
            if (resolved != null)
            {
                var r = Convert.ToInt32(resolved[..2], 16);
                var g = Convert.ToInt32(resolved[2..4], 16);
                var b = Convert.ToInt32(resolved[4..6], 16);
                color = $"rgba({r},{g},{b},{opacity:0.##})";
            }
            else
            {
                color = $"rgba(0,0,0,{opacity:0.##})";
            }
        }

        var blurPt = shadow.BlurRadius?.HasValue == true ? shadow.BlurRadius.Value / 12700.0 : 4;
        var distPt = shadow.Distance?.HasValue == true ? shadow.Distance.Value / 12700.0 : 3;
        var angleDeg = shadow.Direction?.HasValue == true ? shadow.Direction.Value / 60000.0 : 45;
        var angleRad = angleDeg * Math.PI / 180;
        var offsetX = distPt * Math.Cos(angleRad);
        var offsetY = distPt * Math.Sin(angleRad);

        return $"box-shadow:{offsetX:0.##}pt {offsetY:0.##}pt {blurPt:0.##}pt {color}";
    }

    // ==================== CSS Helper: Preset Geometry ====================

    private static string PresetGeometryToCss(string preset) =>
        PresetGeometryToCss(preset, 0, 0, null);

    private static string PresetGeometryToCss(string preset, long widthEmu, long heightEmu,
        Drawing.PresetGeometry? presetGeom)
    {
        // Calculate roundRect corner radius from avLst or default (16.667% of shorter side)
        if (preset is "roundRect" or "round1Rect" or "round2SameRect" or "round2DiagRect")
        {
            var minSide = Math.Min(widthEmu, heightEmu);
            // Default adjustment value is 16667 (= 16.667%)
            long avVal = 16667;
            var avList = presetGeom?.GetFirstChild<Drawing.AdjustValueList>();
            var gd = avList?.GetFirstChild<Drawing.ShapeGuide>();
            if (gd?.Formula?.Value != null && gd.Formula.Value.StartsWith("val "))
            {
                if (long.TryParse(gd.Formula.Value.AsSpan(4), out var parsed))
                    avVal = parsed;
            }
            var radiusEmu = minSide * avVal / 100000;
            var radiusCm = radiusEmu / 360000.0;
            var r = $"{radiusCm:0.##}cm";
            if (minSide <= 0) r = "8px"; // fallback if no dimensions

            return preset switch
            {
                "roundRect" => $"border-radius:{r}",
                "round1Rect" => $"border-radius:{r} 0 0 0",
                "round2SameRect" => $"border-radius:{r} {r} 0 0",
                "round2DiagRect" => $"border-radius:{r} 0 {r} 0",
                _ => ""
            };
        }

        return preset switch
        {
            // Rectangles
            "rect" => "",
            "snip1Rect" => "clip-path:polygon(0 0,92% 0,100% 8%,100% 100%,0 100%)",
            "snip2SameRect" => "clip-path:polygon(8% 0,92% 0,100% 8%,100% 100%,0 100%,0 8%)",
            "snip2DiagRect" => "clip-path:polygon(8% 0,100% 0,100% 92%,92% 100%,0 100%,0 8%)",

            // Ellipses
            "ellipse" => "border-radius:50%",

            // Triangles
            "triangle" or "isosTriangle" => "clip-path:polygon(50% 0,100% 100%,0 100%)",
            "rtTriangle" => "clip-path:polygon(0 0,100% 100%,0 100%)",

            // Diamonds and parallelograms
            "diamond" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            "parallelogram" => "clip-path:polygon(15% 0,100% 0,85% 100%,0 100%)",
            "trapezoid" => "clip-path:polygon(20% 0,80% 0,100% 100%,0 100%)",

            // Polygons
            "pentagon" => "clip-path:polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%)",
            "hexagon" => "clip-path:polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%)",
            "heptagon" => "clip-path:polygon(50% 0,90% 20%,100% 60%,75% 100%,25% 100%,0 60%,10% 20%)",
            "octagon" => "clip-path:polygon(29% 0,71% 0,100% 29%,100% 71%,71% 100%,29% 100%,0 71%,0 29%)",
            "decagon" => "clip-path:polygon(35% 0,65% 0,90% 12%,100% 38%,100% 62%,90% 88%,65% 100%,35% 100%,10% 88%,0 62%,0 38%,10% 12%)",
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
            "leftArrow" => "clip-path:polygon(30% 0,30% 20%,100% 20%,100% 80%,30% 80%,30% 100%,0 50%)",
            "upArrow" => "clip-path:polygon(20% 30%,50% 0,80% 30%,80% 100%,20% 100%)",
            "downArrow" => "clip-path:polygon(20% 0,80% 0,80% 70%,100% 70%,50% 100%,0 70%,20% 70%)",
            "leftRightArrow" => "clip-path:polygon(0 50%,15% 20%,15% 35%,85% 35%,85% 20%,100% 50%,85% 80%,85% 65%,15% 65%,15% 80%)",
            "upDownArrow" => "clip-path:polygon(50% 0,80% 15%,65% 15%,65% 85%,80% 85%,50% 100%,20% 85%,35% 85%,35% 15%,20% 15%)",
            "notchedRightArrow" => "clip-path:polygon(0 20%,70% 20%,70% 0,100% 50%,70% 100%,70% 80%,0 80%,10% 50%)",
            "bentArrow" => "clip-path:polygon(0 20%,60% 20%,60% 0,100% 35%,60% 70%,60% 50%,20% 50%,20% 100%,0 100%)",
            "chevron" => "clip-path:polygon(0 0,80% 0,100% 50%,80% 100%,0 100%,20% 50%)",
            "homePlate" => "clip-path:polygon(0 0,85% 0,100% 50%,85% 100%,0 100%)",
            "stripedRightArrow" => "clip-path:polygon(10% 20%,12% 20%,12% 80%,10% 80%,10% 20%,15% 20%,70% 20%,70% 0,100% 50%,70% 100%,70% 80%,15% 80%)",

            // Callouts
            "wedgeRoundRectCallout" => "border-radius:6px",
            "wedgeRectCallout" or "wedgeEllipseCallout" => "",
            "cloudCallout" => "border-radius:50%",

            // Crosses and plus
            "plus" or "cross" => "clip-path:polygon(33% 0,67% 0,67% 33%,100% 33%,100% 67%,67% 67%,67% 100%,33% 100%,33% 67%,0 67%,0 33%,33% 33%)",

            // Heart (polygon approximation)
            "heart" => "clip-path:polygon(50% 18%,65% 0,85% 0,100% 15%,100% 35%,50% 100%,0 35%,0 15%,15% 0,35% 0)",

            // Cloud (rounded blob)
            "cloud" or "cloudCallout" => "border-radius:50% 50% 45% 55% / 60% 40% 55% 45%",

            // Smiley (circle)
            "smileyFace" or "smiley" => "border-radius:50%",

            // Gear (polygon approximation of 6-tooth gear)
            "gear6" => "clip-path:polygon(50% 0,61% 10%,75% 3%,80% 18%,97% 25%,88% 38%,100% 50%,88% 62%,97% 75%,80% 82%,75% 97%,61% 90%,50% 100%,39% 90%,25% 97%,20% 82%,3% 75%,12% 62%,0 50%,12% 38%,3% 25%,20% 18%,25% 3%,39% 10%)",
            "gear9" => "clip-path:polygon(50% 0,56% 8%,65% 2%,68% 12%,78% 9%,78% 20%,88% 20%,85% 30%,95% 35%,90% 44%,100% 50%,90% 56%,95% 65%,85% 70%,88% 80%,78% 80%,78% 91%,68% 88%,65% 98%,56% 92%,50% 100%,44% 92%,35% 98%,32% 88%,22% 91%,22% 80%,12% 80%,15% 70%,5% 65%,10% 56%,0 50%,10% 44%,5% 35%,15% 30%,12% 20%,22% 20%,22% 9%,32% 12%,35% 2%,44% 8%)",

            // 3D-like shapes (rendered flat)
            "cube" => "",
            "can" or "cylinder" => "border-radius:50%/10%",
            "bevel" => "border:3px outset currentColor",
            "foldedCorner" => "clip-path:polygon(0 0,85% 0,100% 15%,100% 100%,0 100%)",
            "lightningBolt" => "clip-path:polygon(35% 0,55% 35%,100% 30%,45% 55%,80% 100%,25% 60%,0 80%,30% 45%)",

            // Misc shapes
            "frame" => "clip-path:polygon(0 0,100% 0,100% 100%,0 100%,0 12%,12% 12%,12% 88%,88% 88%,88% 12%,0 12%)",
            "donut" => "border-radius:50%", // approximate — real donut has inner hole
            "noSmoking" => "border-radius:50%",
            "halfFrame" => "clip-path:polygon(0 0,100% 0,100% 15%,15% 15%,15% 100%,0 100%)",
            "corner" => "clip-path:polygon(0 0,50% 0,50% 50%,100% 50%,100% 100%,0 100%)",
            "pie" or "arc" => "border-radius:50%",

            // Ribbons/banners
            "ribbon" or "ribbon2" or "wave" or "doubleWave" => "",
            "horizontalScroll" or "verticalScroll" => "border-radius:4px",

            // Flowchart
            "flowChartProcess" => "",
            "flowChartAlternateProcess" => "border-radius:8px",
            "flowChartDecision" => "clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)",
            "flowChartInputOutput" or "flowChartData" => "clip-path:polygon(15% 0,100% 0,85% 100%,0 100%)",
            "flowChartPredefinedProcess" => "border-left:3px double currentColor;border-right:3px double currentColor",
            "flowChartDocument" => "",
            "flowChartMultidocument" => "",
            "flowChartTerminator" => "border-radius:50%/100%",
            "flowChartPreparation" => "clip-path:polygon(17% 0,83% 0,100% 50%,83% 100%,17% 100%,0 50%)",
            "flowChartManualInput" => "clip-path:polygon(0 15%,100% 0,100% 100%,0 100%)",
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
            "flowChartOfflineStorage" => "clip-path:polygon(10% 0,90% 0,50% 100%)",
            "flowChartMagneticDisk" => "border-radius:50%/20%",
            "flowChartConnector" or "flowChartOffpageConnector" => "border-radius:50%",

            // Block arrows
            "curvedRightArrow" or "curvedLeftArrow" or "curvedUpArrow" or "curvedDownArrow" => "",
            "circularArrow" => "border-radius:50%",

            // Math
            "mathPlus" => "clip-path:polygon(33% 0,67% 0,67% 33%,100% 33%,100% 67%,67% 67%,67% 100%,33% 100%,33% 67%,0 67%,0 33%,33% 33%)",
            "mathMinus" => "clip-path:polygon(0 35%,100% 35%,100% 65%,0 65%)",
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
            var alpha = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (alpha.HasValue && alpha.Value < 100000)
            {
                var r = Convert.ToInt32(hexPart[..2], 16);
                var g = Convert.ToInt32(hexPart[2..4], 16);
                var b = Convert.ToInt32(hexPart[4..6], 16);
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

        return null;
    }

    private static string ApplyColorTransforms(string hex, Drawing.SchemeColor schemeColor)
    {
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        var lumMod = schemeColor.GetFirstChild<Drawing.LuminanceModulation>()?.Val?.Value;
        var lumOff = schemeColor.GetFirstChild<Drawing.LuminanceOffset>()?.Val?.Value;
        var tint = schemeColor.GetFirstChild<Drawing.Tint>()?.Val?.Value;
        var shade = schemeColor.GetFirstChild<Drawing.Shade>()?.Val?.Value;
        var alpha = schemeColor.GetFirstChild<Drawing.Alpha>()?.Val?.Value;

        // OOXML spec: tint blends toward white, shade blends toward black
        if (tint.HasValue)
        {
            var t = tint.Value / 100000.0;
            r = (int)(r + (255 - r) * (1 - t));
            g = (int)(g + (255 - g) * (1 - t));
            b = (int)(b + (255 - b) * (1 - t));
        }

        if (shade.HasValue)
        {
            var s = shade.Value / 100000.0;
            r = (int)(r * s);
            g = (int)(g * s);
            b = (int)(b * s);
        }

        // OOXML spec: lumMod/lumOff operate in HSL space
        if (lumMod.HasValue || lumOff.HasValue)
        {
            var mod = (lumMod ?? 100000) / 100000.0;
            var off = (lumOff ?? 0) / 100000.0;
            RgbToHsl(r, g, b, out var h, out var s, out var l);
            l = Math.Clamp(l * mod + off, 0, 1);
            HslToRgb(h, s, l, out r, out g, out b);
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);

        if (alpha.HasValue && alpha.Value < 100000)
            return $"rgba({r},{g},{b},{alpha.Value / 100000.0:0.##})";

        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private static void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
    {
        var rf = r / 255.0;
        var gf = g / 255.0;
        var bf = b / 255.0;
        var max = Math.Max(rf, Math.Max(gf, bf));
        var min = Math.Min(rf, Math.Min(gf, bf));
        var delta = max - min;

        l = (max + min) / 2.0;

        if (delta < 1e-10)
        {
            h = 0;
            s = 0;
            return;
        }

        s = l < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);

        if (Math.Abs(max - rf) < 1e-10)
            h = ((gf - bf) / delta + (gf < bf ? 6 : 0)) / 6.0;
        else if (Math.Abs(max - gf) < 1e-10)
            h = ((bf - rf) / delta + 2) / 6.0;
        else
            h = ((rf - gf) / delta + 4) / 6.0;
    }

    private static void HslToRgb(double h, double s, double l, out int r, out int g, out int b)
    {
        if (s < 1e-10)
        {
            r = g = b = (int)Math.Round(l * 255);
            return;
        }

        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;

        r = (int)Math.Round(HueToRgb(p, q, h + 1.0 / 3) * 255);
        g = (int)Math.Round(HueToRgb(p, q, h) * 255);
        b = (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3) * 255);
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    /// <summary>
    /// Build a map of scheme color names to hex values from the presentation theme.
    /// </summary>
    private Dictionary<string, string> ResolveThemeColorMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var theme = _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault()?.ThemePart?.Theme;
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return map;

        void Add(string name, OpenXmlCompositeElement? color)
        {
            if (color == null) return;
            var rgb = color.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            var sys = color.GetFirstChild<Drawing.SystemColor>();
            var srgb = sys?.LastColor?.Value ?? sys?.Val?.InnerText;
            var hex = rgb ?? srgb;
            if (hex != null) map[name] = hex;
        }

        Add("dk1", colorScheme.Dark1Color);
        Add("dk2", colorScheme.Dark2Color);
        Add("lt1", colorScheme.Light1Color);
        Add("lt2", colorScheme.Light2Color);
        Add("accent1", colorScheme.Accent1Color);
        Add("accent2", colorScheme.Accent2Color);
        Add("accent3", colorScheme.Accent3Color);
        Add("accent4", colorScheme.Accent4Color);
        Add("accent5", colorScheme.Accent5Color);
        Add("accent6", colorScheme.Accent6Color);
        Add("hlink", colorScheme.Hyperlink);
        Add("folHlink", colorScheme.FollowedHyperlinkColor);

        // Aliases
        if (map.TryGetValue("dk1", out var dk1)) { map["tx1"] = dk1; map["dark1"] = dk1; map["text1"] = dk1; }
        if (map.TryGetValue("dk2", out var dk2)) { map["dark2"] = dk2; map["text2"] = dk2; map["tx2"] = dk2; }
        if (map.TryGetValue("lt1", out var lt1)) { map["bg1"] = lt1; map["light1"] = lt1; map["background1"] = lt1; }
        if (map.TryGetValue("lt2", out var lt2)) { map["bg2"] = lt2; map["light2"] = lt2; map["background2"] = lt2; }

        return map;
    }

    // ==================== Image Helpers ====================

    private static string? BlipToDataUri(Drawing.BlipFill blipFill, OpenXmlPart part)
    {
        var blip = blipFill.GetFirstChild<Drawing.Blip>();
        if (blip?.Embed?.HasValue != true) return null;

        try
        {
            var imgPart = part.GetPartById(blip.Embed.Value!);
            using var stream = imgPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());
            return $"data:{imgPart.ContentType ?? "image/png"};base64,{base64}";
        }
        catch
        {
            return null;
        }
    }

    // ==================== Utility ====================

    private static double EmuToCm(long emu)
    {
        return Math.Round(emu / 360000.0, 3);
    }

    private static string HtmlEncode(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }

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
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);
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
