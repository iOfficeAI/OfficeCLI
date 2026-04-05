// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Shared chart SVG rendering logic used by both PowerPoint and Excel HTML preview.
/// </summary>
internal class ChartSvgRenderer
{
    // Default chart colors matching Office theme accent colors
    public static readonly string[] DefaultColors = [
        "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
        "#264478", "#9E480E", "#636363", "#997300", "#255E91", "#43682B"
    ];

    // Chart styling — configurable per chart instance
    public string ValueColor { get; set; } = "#D0D8E0";
    public string CatColor { get; set; } = "#C8D0D8";
    public string AxisColor { get; set; } = "#B0B8C0";
    public string SecondaryAxisColor { get; set; } = "#aaa";
    public string GridColor { get; set; } = "#333";
    public string AxisLineColor { get; set; } = "#555";
    public int ValFontPx { get; set; } = 9;
    public int CatFontPx { get; set; } = 9;
    public int DataLabelFontPx { get; set; } = 8;
    public int AxisTickCount { get; set; } = 4;

    public static string HtmlEncode(string text) =>
        text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
            .Replace("\"", "&quot;").Replace("'", "&#39;");

    public void RenderBarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph,
        bool horizontal, bool stacked = false, bool percentStacked = false,
        double? ooxmlMax = null, double? ooxmlMin = null, double? ooxmlMajorUnit = null,
        int? ooxmlGapWidth = null, int valFontSize = 9, int catFontSize = 9,
        bool showDataLabels = false, string? valNumFmt = null, string? plotFillColor = null)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;
        if (percentStacked) stacked = true;

        double maxVal;
        if (percentStacked) maxVal = 100;
        else if (stacked)
        {
            maxVal = 0;
            for (int c = 0; c < catCount; c++)
            {
                var sum = series.Sum(s => c < s.values.Length ? s.values[c] : 0);
                if (sum > maxVal) maxVal = sum;
            }
        }
        else maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;

        double niceMax, tickStep;
        int nTicks;
        if (!percentStacked)
        {
            if (ooxmlMax.HasValue && ooxmlMajorUnit.HasValue)
            {
                niceMax = ooxmlMax.Value;
                tickStep = ooxmlMajorUnit.Value;
                nTicks = (int)Math.Round(niceMax / tickStep);
            }
            else (niceMax, tickStep, nTicks) = ComputeNiceAxis(ooxmlMax ?? maxVal);
        }
        else { niceMax = 100; nTicks = 5; tickStep = 20; }

        if (horizontal)
        {
            // Estimate label width from longest category name (approx 0.5 × fontSize per char)
            var maxLabelLen = categories.Length > 0 ? categories.Max(c => c.Length) : 0;
            var hLabelMargin = (int)(maxLabelLen * catFontSize * 0.5) + 4;
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;

            // Plot area background starts at the Y-axis (plotOx), labels are outside
            if (plotFillColor != null)
                sb.AppendLine($"        <rect x=\"{plotOx}\" y=\"{oy}\" width=\"{plotPw}\" height=\"{ph}\" fill=\"#{plotFillColor}\"/>");

            var groupH = (double)ph / Math.Max(catCount, 1);
            var gapPct = (ooxmlGapWidth ?? 150) / 100.0;
            double barH, gap;
            if (stacked) { barH = groupH / (1 + gapPct); gap = (groupH - barH) / 2; }
            else { barH = groupH / (serCount + gapPct); gap = barH * gapPct / 2; }

            for (int t = 1; t <= nTicks; t++)
            {
                var gx = plotOx + (double)plotPw * t / nTicks;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                var dataIdx = catCount - 1 - c;
                double stackX = 0;
                var catSum = percentStacked ? series.Sum(s => dataIdx < s.values.Length ? s.values[dataIdx] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = dataIdx < series[s].values.Length ? series[s].values[dataIdx] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    var barW = (val / niceMax) * plotPw;
                    if (stacked)
                    {
                        var bx = plotOx + (stackX / niceMax) * plotPw;
                        var by = oy + c * groupH + gap;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        stackX += val;
                    }
                    else
                    {
                        var bx = plotOx;
                        var by = oy + c * groupH + gap + (serCount - 1 - s) * barH;
                        sb.AppendLine($"        <rect x=\"{bx}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                    }
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var dataIdx = catCount - 1 - c;
                var label = dataIdx < categories.Length ? categories[dataIdx] : "";
                var ly = oy + c * groupH + groupH / 2;
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"{CatColor}\" font-size=\"{catFontSize}\" text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= nTicks; t++)
            {
                var val = tickStep * t;
                var label = percentStacked ? $"{(int)val}%" : FormatAxisValue(val, valNumFmt);
                var tx = plotOx + (double)plotPw * t / nTicks;
                sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{AxisColor}\" font-size=\"{valFontSize}\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var gapPct = (ooxmlGapWidth ?? 150) / 100.0;
            double barW, gap;
            if (stacked) { barW = groupW / (1 + gapPct); gap = (groupW - barW) / 2; }
            else { barW = groupW / (serCount + gapPct); gap = barW * gapPct / 2; }

            for (int t = 1; t <= nTicks; t++)
            {
                var gy = oy + ph - (double)ph * t / nTicks;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
            }
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                double stackY = 0;
                var catSum = percentStacked ? series.Sum(s => c < s.values.Length ? s.values[c] : 0) : 1;
                for (int s = 0; s < serCount; s++)
                {
                    var rawVal = c < series[s].values.Length ? series[s].values[c] : 0;
                    var val = percentStacked && catSum > 0 ? (rawVal / catSum) * 100 : rawVal;
                    var barH = (val / niceMax) * ph;
                    if (stacked)
                    {
                        var bx = ox + c * groupW + gap;
                        var by = oy + ph - (stackY / niceMax) * ph - barH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        stackY += val;
                    }
                    else
                    {
                        var bx = ox + c * groupW + gap + s * barW;
                        var by = oy + ph - barH;
                        sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                        if (showDataLabels)
                        {
                            var vlabel = rawVal % 1 == 0 ? $"{(int)rawVal}" : $"{rawVal:0.#}";
                            sb.AppendLine($"        <text x=\"{bx + barW / 2:0.#}\" y=\"{by - 3:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\">{vlabel}</text>");
                        }
                    }
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{catFontSize}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= nTicks; t++)
            {
                var val = tickStep * t;
                var label = percentStacked ? $"{(int)val}%" : FormatAxisValue(val, valNumFmt);
                var ty = oy + ph - (double)ph * t / nTicks;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{valFontSize}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    public void RenderLineChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool showDataLabels = false)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var (niceMax, tickStep, nTicks) = ComputeNiceAxis(maxVal);

        for (int t = 1; t <= nTicks; t++)
        {
            var gy = oy + ph - (double)ph * t / nTicks;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\" stroke-dasharray=\"none\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / niceMax) * ph;
                points.Add($"{px:0.#},{py:0.#}");
            }
            if (points.Count > 0)
            {
                var lineColor = colors[s % colors.Count];
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{lineColor}\" stroke-width=\"2\"/>");
                for (int p = 0; p < points.Count; p++)
                {
                    var parts = points[p].Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{lineColor}\"/>");
                    if (showDataLabels)
                    {
                        var val = series[s].values[p];
                        var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                        sb.AppendLine($"        <text x=\"{parts[0]}\" y=\"{double.Parse(parts[1]) - 6:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\">{vlabel}</text>");
                    }
                }
            }
        }
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= nTicks; t++)
        {
            var val = tickStep * t;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / nTicks;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderPieChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, double holeRatio = 0.0, bool showDataLabels = false)
    {
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.42;
        var innerR = r * holeRatio;
        var startAngle = -Math.PI / 2;

        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : DefaultColors[i % DefaultColors.Length];

            if (values.Length == 1 && holeRatio <= 0)
                sb.AppendLine($"        <circle cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" r=\"{r:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            else if (holeRatio > 0)
            {
                var ox1 = cx + r * Math.Cos(startAngle); var oy1 = cy + r * Math.Sin(startAngle);
                var ox2 = cx + r * Math.Cos(endAngle); var oy2 = cy + r * Math.Sin(endAngle);
                var ix1 = cx + innerR * Math.Cos(endAngle); var iy1 = cy + innerR * Math.Sin(endAngle);
                var ix2 = cx + innerR * Math.Cos(startAngle); var iy2 = cy + innerR * Math.Sin(startAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {ox1:0.#},{oy1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {ox2:0.#},{oy2:0.#} L {ix1:0.#},{iy1:0.#} A {innerR:0.#},{innerR:0.#} 0 {largeArc},0 {ix2:0.#},{iy2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            else
            {
                var x1 = cx + r * Math.Cos(startAngle); var y1 = cy + r * Math.Sin(startAngle);
                var x2 = cx + r * Math.Cos(endAngle); var y2 = cy + r * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {cx:0.#},{cy:0.#} L {x1:0.#},{y1:0.#} A {r:0.#},{r:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
            startAngle = endAngle;
        }
        if (showDataLabels)
        {
            var labelAngle = -Math.PI / 2;
            var labelR = holeRatio > 0 ? r * (1 + holeRatio) / 2 : r * 0.65;
            for (int i = 0; i < values.Length; i++)
            {
                var sliceAngle = 2 * Math.PI * values[i] / total;
                var midAngle = labelAngle + sliceAngle / 2;
                var lx = cx + labelR * Math.Cos(midAngle);
                var ly = cy + labelR * Math.Sin(midAngle);
                var pct = values[i] / total * 100;
                var label = pct >= 5 ? $"{pct:0}%" : "";
                if (!string.IsNullOrEmpty(label))
                    sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"#fff\" font-size=\"9\" font-weight=\"bold\" text-anchor=\"middle\" dominant-baseline=\"central\">{label}</text>");
                labelAngle += sliceAngle;
            }
        }
    }

    public void RenderAreaChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool stacked = false)
    {
        if (series.Count == 0) return;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        if (catCount == 0) return;

        var cumulative = new double[series.Count, catCount];
        for (int c = 0; c < catCount; c++)
        {
            double runningSum = 0;
            for (int s = 0; s < series.Count; s++)
            {
                var val = c < series[s].values.Length ? series[s].values[c] : 0;
                runningSum += stacked ? val : 0;
                cumulative[s, c] = stacked ? runningSum : val;
            }
        }
        var allAreaVals = series.SelectMany(s => s.values).DefaultIfEmpty(0).ToArray();
        var maxVal = 0.0;
        var minVal = 0.0;
        if (stacked) { for (int c = 0; c < catCount; c++) maxVal = Math.Max(maxVal, cumulative[series.Count - 1, c]); }
        else { maxVal = allAreaVals.Max(); minVal = Math.Min(0.0, allAreaVals.Min()); }
        if (maxVal <= minVal) maxVal = minVal + 1;
        var (niceMax, tickInterval, tickCount) = ComputeNiceAxis(Math.Abs(maxVal) > Math.Abs(minVal) ? maxVal : -minVal);
        // For non-stacked charts with negative values, expand the axis to cover minVal
        var niceMin = minVal < 0 ? -ComputeNiceAxis(-minVal).niceMax : 0.0;
        var axisRange = niceMax - niceMin;

        // Helper: map a data value to a y-coordinate within [oy, oy+ph]
        double DataToY(double v) => oy + ph - (v - niceMin) / axisRange * ph;
        double ZeroY() => DataToY(0.0);

        for (int t = 1; t <= tickCount; t++)
        {
            var gy = oy + ph - (double)ph * t / tickCount;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        if (stacked)
        {
            for (int s = series.Count - 1; s >= 0; s--)
            {
                var topPoints = new List<string>();
                var bottomPoints = new List<string>();
                for (int c = 0; c < catCount; c++)
                {
                    var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                    topPoints.Add($"{px:0.#},{oy + ph - (cumulative[s, c] / niceMax) * ph:0.#}");
                    var bottomVal = s > 0 ? cumulative[s - 1, c] : 0;
                    bottomPoints.Add($"{px:0.#},{oy + ph - (bottomVal / niceMax) * ph:0.#}");
                }
                bottomPoints.Reverse();
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", topPoints)} {string.Join(" ", bottomPoints)}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
            }
        }
        else
        {
            var baseY = ZeroY();
            var renderOrder = Enumerable.Range(0, series.Count).OrderByDescending(s => series[s].values.DefaultIfEmpty(0).Max()).ToList();
            foreach (var s in renderOrder)
            {
                var topPoints = new List<string>();
                for (int c = 0; c < catCount; c++)
                {
                    var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                    var val = c < series[s].values.Length ? series[s].values[c] : 0;
                    topPoints.Add($"{px:0.#},{DataToY(val):0.#}");
                }
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastIdx = Math.Min(series[s].values.Length - 1, catCount - 1);
                var lastX = ox + (catCount > 1 ? (double)pw * lastIdx / (catCount - 1) : pw / 2.0);
                sb.AppendLine($"        <polygon points=\"{firstX:0.#},{baseY:0.#} {string.Join(" ", topPoints)} {lastX:0.#},{baseY:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
            }
        }
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= tickCount; t++)
        {
            var val = tickInterval * t;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / tickCount;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderRadarChartSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH, int catLabelFontSize = 0)
    {
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        if (catCount < 3) return;
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max();
        if (maxVal <= 0) maxVal = 1;

        var labelSize = catLabelFontSize > 0 ? catLabelFontSize : 11;
        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var r = Math.Min(svgW, svgH) * 0.33;

        for (int ring = 1; ring <= 5; ring++)
        {
            var rr = r * ring / 5;
            var gridPoints = new List<string>();
            for (int c = 0; c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                gridPoints.Add($"{cx + rr * Math.Cos(angle):0.#},{cy + rr * Math.Sin(angle):0.#}");
            }
            sb.AppendLine($"        <polygon points=\"{string.Join(" ", gridPoints)}\" fill=\"none\" stroke=\"#ccc\" stroke-width=\"0.5\"/>");
        }
        for (int c = 0; c < catCount; c++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            sb.AppendLine($"        <line x1=\"{cx:0.#}\" y1=\"{cy:0.#}\" x2=\"{cx + r * Math.Cos(angle):0.#}\" y2=\"{cy + r * Math.Sin(angle):0.#}\" stroke=\"#ccc\" stroke-width=\"0.5\"/>");
        }
        for (int s = 0; s < series.Count; s++)
        {
            var points = new List<string>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
                var val = series[s].values[c] / maxVal * r;
                points.Add($"{cx + val * Math.Cos(angle):0.#},{cy + val * Math.Sin(angle):0.#}");
            }
            if (points.Count > 0)
            {
                var serColor = colors[s % colors.Count];
                sb.AppendLine($"        <polygon points=\"{string.Join(" ", points)}\" fill=\"{serColor}\" fill-opacity=\"0.2\" stroke=\"{serColor}\" stroke-width=\"2\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{serColor}\"/>");
                }
            }
        }
        foreach (var frac in new[] { 0.2, 0.4, 0.6, 0.8, 1.0 })
        {
            var val = maxVal * frac;
            var tickLabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{cx + 2:0.#}\" y=\"{cy - r * frac:0.#}\" fill=\"{AxisColor}\" font-size=\"8\" dominant-baseline=\"middle\">{tickLabel}</text>");
        }
        var labelOffset = Math.Max(18, r * 0.15);
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var angle = -Math.PI / 2 + 2 * Math.PI * c / catCount;
            var lx = cx + (r + labelOffset) * Math.Cos(angle);
            var ly = cy + (r + labelOffset) * Math.Sin(angle);
            var anchor = Math.Abs(Math.Cos(angle)) < 0.1 ? "middle" : (Math.Cos(angle) > 0 ? "start" : "end");
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"{CatColor}\" font-size=\"{labelSize}\" text-anchor=\"{anchor}\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
        }
    }

    public void RenderBubbleChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> series, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        var bubbleSeries = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(e => e.LocalName == "ser" && e.Parent?.LocalName == "bubbleChart").ToList();

        var allX = new List<double>(); var allY = new List<double>(); var allSize = new List<double>();
        var seriesData = new List<(double[] x, double[] y, double[] size)>();

        for (int s = 0; s < bubbleSeries.Count; s++)
        {
            var ser = bubbleSeries[s];
            var xVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "xVal")) ?? [];
            var yVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "yVal")) ?? [];
            var sizeVals = ChartHelper.ReadNumericData(ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "bubbleSize")) ?? yVals;
            seriesData.Add((xVals, yVals, sizeVals));
            allX.AddRange(xVals); allY.AddRange(yVals); allSize.AddRange(sizeVals);
        }
        if (seriesData.Count == 0)
        {
            foreach (var s in series)
            {
                var xVals = Enumerable.Range(0, s.values.Length).Select(i => (double)i).ToArray();
                seriesData.Add((xVals, s.values, s.values));
                allX.AddRange(xVals); allY.AddRange(s.values); allSize.AddRange(s.values);
            }
        }
        if (allY.Count == 0) return;
        var minX = allX.Min(); var maxX = allX.Max(); if (maxX <= minX) maxX = minX + 1;
        var minY = allY.Min(); var maxY = allY.Max(); if (maxY <= minY) maxY = minY + 1;
        var maxSz = allSize.Count > 0 ? allSize.Max() : 1; if (maxSz <= 0) maxSz = 1;
        var bubbleScaleEl = plotArea.Descendants<BubbleScale>().FirstOrDefault();
        var bubbleScale = bubbleScaleEl?.Val?.HasValue == true ? bubbleScaleEl.Val.Value / 100.0 : 1.0;
        var maxRadius = Math.Min(pw, ph) * 0.12 * bubbleScale;

        for (int t = 1; t <= 4; t++)
        {
            var gy = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\"/>");
        }
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        for (int s = 0; s < seriesData.Count; s++)
        {
            var (xVals, yVals, sizeVals) = seriesData[s];
            var count = Math.Min(xVals.Length, yVals.Length);
            for (int i = 0; i < count; i++)
            {
                var bx = ox + ((xVals[i] - minX) / (maxX - minX)) * pw;
                var by = oy + ph - ((yVals[i] - minY) / (maxY - minY)) * ph;
                var sz = i < sizeVals.Length ? sizeVals[i] : yVals[i];
                var r = Math.Sqrt(Math.Max(0, sz) / maxSz) * maxRadius + maxRadius * 0.15;
                sb.AppendLine($"        <circle cx=\"{bx:0.#}\" cy=\"{by:0.#}\" r=\"{r:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.6\"/>");
            }
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = minX + (maxX - minX) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox + (double)pw * t / 4:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{label}</text>");
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = minY + (maxY - minY) * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / 4:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public void RenderComboChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> seriesList, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        var barIndices = new HashSet<int>();
        var lineIndices = new HashSet<int>();
        var areaIndices = new HashSet<int>();
        var secondaryIndices = new HashSet<int>(); // series on secondary Y-axis

        // Detect which axis IDs are secondary (right-side value axis)
        var secondaryAxIds = new HashSet<uint>();
        var valAxes = plotArea.Elements<ValueAxis>().ToList();
        if (valAxes.Count >= 2)
        {
            // The secondary value axis is the one with axPos="r"
            // Use .InnerText because AxisPositionValues.ToString() is broken in Open XML SDK v3+
            foreach (var va in valAxes)
            {
                var posText = va.GetFirstChild<AxisPosition>()?.Val?.InnerText;
                if (posText == "r")
                {
                    var id = va.GetFirstChild<AxisId>()?.Val?.Value;
                    if (id.HasValue) secondaryAxIds.Add(id.Value);
                }
            }
            // Fallback: if no explicit right axis found, treat 2nd valAx as secondary
            if (secondaryAxIds.Count == 0 && valAxes.Count >= 2)
            {
                var id = valAxes[1].GetFirstChild<AxisId>()?.Val?.Value;
                if (id.HasValue) secondaryAxIds.Add(id.Value);
            }
        }

        var idx = 0;
        foreach (var chartEl in plotArea.ChildElements)
        {
            var serElements = chartEl.Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser").ToList();
            if (serElements.Count == 0) continue;
            var localName = chartEl.LocalName.ToLowerInvariant();
            var isBar = localName.Contains("bar");
            var isArea = localName.Contains("area");

            // Check if this chart group uses a secondary axis
            var axIds = chartEl.ChildElements
                .Where(e => e.LocalName == "axId")
                .Select(e => e.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value)
                .Where(v => v != null)
                .Select(v => uint.TryParse(v, out var u) ? u : 0)
                .ToHashSet();
            var isSecondary = axIds.Overlaps(secondaryAxIds);

            foreach (var _ in serElements)
            {
                if (isBar) barIndices.Add(idx);
                else if (isArea) areaIndices.Add(idx);
                else lineIndices.Add(idx);
                if (isSecondary) secondaryIndices.Add(idx);
                idx++;
            }
        }

        // Separate primary and secondary values for independent axis scaling
        var primaryValues = seriesList.Where((_, i) => !secondaryIndices.Contains(i)).SelectMany(s => s.values).ToArray();
        var secondaryValues = seriesList.Where((_, i) => secondaryIndices.Contains(i)).SelectMany(s => s.values).ToArray();
        if (primaryValues.Length == 0 && secondaryValues.Length == 0) return;

        var priMax = primaryValues.Length > 0 ? primaryValues.Max() : 0; if (priMax <= 0) priMax = 1;
        var (priNiceMax, _, _) = ComputeNiceAxis(priMax);
        var hasSecondary = secondaryValues.Length > 0;
        double secNiceMax = 1;
        if (hasSecondary)
        {
            var secMax = secondaryValues.Max(); if (secMax <= 0) secMax = 1;
            (secNiceMax, _, _) = ComputeNiceAxis(secMax);
        }

        var catCount = Math.Max(categories.Length, seriesList.Max(s => s.values.Length));

        // Axes
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        // Bar series (primary axis)
        var barSeries = barIndices.Where(i => i < seriesList.Count).ToList();
        if (barSeries.Count > 0)
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.5 / barSeries.Count;
            var gap = (groupW - barSeries.Count * barW) / 2;
            for (int bi = 0; bi < barSeries.Count; bi++)
            {
                var s = barSeries[bi];
                var axMax = secondaryIndices.Contains(s) ? secNiceMax : priNiceMax;
                for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
                {
                    var val = seriesList[s].values[c];
                    var barH = (val / axMax) * ph;
                    sb.AppendLine($"        <rect x=\"{ox + c * groupW + gap + bi * barW:0.#}\" y=\"{oy + ph - barH:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.85\"/>");
                }
            }
        }
        // Area series
        foreach (var s in areaIndices.Where(i => i < seriesList.Count))
        {
            var axMax = secondaryIndices.Contains(s) ? secNiceMax : priNiceMax;
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                points.Add($"{px:0.#},{oy + ph - (seriesList[s].values[c] / axMax) * ph:0.#}");
            }
            if (points.Count > 0)
            {
                var firstX = ox + (catCount > 1 ? 0 : pw / 2.0);
                var lastX = ox + (catCount > 1 ? (double)pw * (seriesList[s].values.Length - 1) / (catCount - 1) : pw / 2.0);
                sb.AppendLine($"        <polygon points=\"{firstX:0.#},{oy + ph} {string.Join(" ", points)} {lastX:0.#},{oy + ph}\" fill=\"{colors[s % colors.Count]}\" opacity=\"0.3\"/>");
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2\"/>");
            }
        }
        // Line series (may use secondary axis)
        foreach (var s in lineIndices.Where(i => i < seriesList.Count))
        {
            var axMax = secondaryIndices.Contains(s) ? secNiceMax : priNiceMax;
            var points = new List<string>();
            for (int c = 0; c < seriesList[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                points.Add($"{px:0.#},{oy + ph - (seriesList[s].values[c] / axMax) * ph:0.#}");
            }
            if (points.Count > 0)
            {
                sb.AppendLine($"        <polyline points=\"{string.Join(" ", points)}\" fill=\"none\" stroke=\"{colors[s % colors.Count]}\" stroke-width=\"2.5\"/>");
                foreach (var pt in points)
                {
                    var parts = pt.Split(',');
                    sb.AppendLine($"        <circle cx=\"{parts[0]}\" cy=\"{parts[1]}\" r=\"3\" fill=\"{colors[s % colors.Count]}\"/>");
                }
            }
        }
        // Category labels
        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (double)pw * c / Math.Max(catCount, 1) + (double)pw / Math.Max(catCount, 1) / 2;
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        // Primary Y-axis labels (left)
        for (int t = 0; t <= AxisTickCount; t++)
        {
            var val = priNiceMax * t / AxisTickCount;
            var label = FormatAxisValue(val);
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / AxisTickCount:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
        // Secondary Y-axis labels (overlaid on left in lighter color)
        if (hasSecondary)
        {
            var secFontPx = Math.Max(ValFontPx - 1, CatFontPx);
            for (int t = 0; t <= AxisTickCount; t++)
            {
                var val = secNiceMax * t / AxisTickCount;
                var label = FormatAxisValue(val);
                sb.AppendLine($"        <text x=\"{ox + 2}\" y=\"{oy + ph - (double)ph * t / AxisTickCount:0.#}\" fill=\"{SecondaryAxisColor}\" font-size=\"{secFontPx}\" text-anchor=\"start\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    private static string FormatAxisValue(double val, string? numFmt = null)
    {
        if (!string.IsNullOrEmpty(numFmt) && numFmt != "General")
            return ApplyNumFmt(val, numFmt);
        if (val == 0) return "0";
        if (Math.Abs(val) >= 1_000_000) return $"{val / 1_000_000:0.#}M";
        if (Math.Abs(val) >= 1_000) return $"{val / 1_000:0.#}K";
        return val % 1 == 0 ? $"{(long)val}" : $"{val:0.#}";
    }

    /// <summary>Apply an OOXML number format code to a value for axis display.</summary>
    private static string ApplyNumFmt(double val, string fmt)
    {
        var prefix = "";
        var suffix = "";
        var f = fmt;

        // Extract literal prefix (e.g. "$")
        if (f.Length > 0 && !char.IsDigit(f[0]) && f[0] != '#' && f[0] != '0' && f[0] != '.')
        {
            prefix = f[0].ToString();
            f = f[1..];
        }
        // Extract literal suffix (e.g. "%")
        if (f.Length > 0 && f[^1] == '%')
        {
            suffix = "%";
            f = f[..^1];
            val *= 100;
        }

        // Determine decimal places from format
        var decIdx = f.IndexOf('.');
        int decimals = decIdx >= 0 ? f[(decIdx + 1)..].Count(c => c is '0' or '#') : 0;

        // Check if thousands separator is used (#,##0 pattern)
        bool useThousands = f.Contains(",##") || f.Contains("#,#");

        string formatted;
        if (useThousands)
            formatted = decimals > 0
                ? val.ToString($"N{decimals}")
                : ((long)val).ToString("N0");
        else
            formatted = decimals > 0
                ? val.ToString($"F{decimals}")
                : (val % 1 == 0 ? $"{(long)val}" : $"{val:0.#}");

        return prefix + formatted + suffix;
    }

    public void RenderStockChartSvg(StringBuilder sb, PlotArea plotArea,
        List<(string name, double[] values)> series, string[] categories, List<string> colors,
        int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var maxVal = allValues.Max(); var minVal = allValues.Min();
        if (maxVal <= minVal) maxVal = minVal + 1;
        var range = maxVal - minVal;
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        var upColor = "#2ECC71"; var downColor = "#E74C3C";
        var stockChart = plotArea.GetFirstChild<StockChart>();
        if (stockChart != null)
        {
            var upFill = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "upBars")
                ?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (upFill != null) upColor = $"#{upFill}";
            var downFill = stockChart.Descendants<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == "downBars")
                ?.Descendants<Drawing.SolidFill>().FirstOrDefault()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (downFill != null) downColor = $"#{downFill}";
        }

        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        var groupW = (double)pw / Math.Max(catCount, 1);
        if (series.Count >= 4)
        {
            for (int c = 0; c < catCount; c++)
            {
                var open = c < series[0].values.Length ? series[0].values[c] : 0;
                var high = c < series[1].values.Length ? series[1].values[c] : 0;
                var low = c < series[2].values.Length ? series[2].values[c] : 0;
                var close = c < series[3].values.Length ? series[3].values[c] : 0;
                var ccx = ox + c * groupW + groupW / 2;
                var yHigh = oy + ph - ((high - minVal) / range) * ph;
                var yLow = oy + ph - ((low - minVal) / range) * ph;
                var yOpen = oy + ph - ((open - minVal) / range) * ph;
                var yClose = oy + ph - ((close - minVal) / range) * ph;
                var color = close >= open ? upColor : downColor;
                var barW = groupW * 0.5;
                sb.AppendLine($"        <line x1=\"{ccx:0.#}\" y1=\"{yHigh:0.#}\" x2=\"{ccx:0.#}\" y2=\"{yLow:0.#}\" stroke=\"{color}\" stroke-width=\"1.5\"/>");
                var bodyTop = Math.Min(yOpen, yClose); var bodyH = Math.Max(Math.Abs(yOpen - yClose), 1);
                sb.AppendLine($"        <rect x=\"{ccx - barW / 2:0.#}\" y=\"{bodyTop:0.#}\" width=\"{barW:0.#}\" height=\"{bodyH:0.#}\" fill=\"{color}\" opacity=\"0.85\"/>");
            }
        }
        else { RenderLineChartSvg(sb, series, categories, colors, ox, oy, pw, ph); return; }

        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            sb.AppendLine($"        <text x=\"{ox + c * groupW + groupW / 2:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }
        for (int t = 0; t <= 4; t++)
        {
            var val = minVal + range * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{oy + ph - (double)ph * t / 4:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }

    public static (double niceMax, double tickStep, int nTicks) ComputeNiceAxis(double maxVal)
    {
        if (maxVal <= 0) maxVal = 1;
        // Guard against subnormal/denormal values where Log10 returns -Infinity
        if (!double.IsFinite(maxVal) || maxVal < 1e-10) maxVal = 1;
        var mag = Math.Pow(10, Math.Floor(Math.Log10(maxVal)));
        if (!double.IsFinite(mag) || mag == 0) mag = 1;
        var res = maxVal / mag;
        var tickStep = res <= 1.5 ? 0.2 * mag : res <= 4 ? 0.5 * mag : res <= 8 ? 1.0 * mag : 2.0 * mag;
        var niceMax = Math.Ceiling(maxVal / tickStep) * tickStep;
        if (niceMax < maxVal * 1.05) niceMax += tickStep;
        var nTicks = (int)Math.Round(niceMax / tickStep);
        if (nTicks < 2) nTicks = 2;
        return (niceMax, tickStep, nTicks);
    }

    // ==================== Shared Chart Info & Rendering ====================

    /// <summary>All metadata extracted from an OOXML chart, used by the shared rendering pipeline.</summary>
    public class ChartInfo
    {
        /// <summary>Original PlotArea element, needed by combo/bubble/stock renderers.</summary>
        public PlotArea? PlotArea { get; set; }
        public string ChartType { get; set; } = "column";
        public string[] Categories { get; set; } = [];
        public List<(string name, double[] values)> Series { get; set; } = [];
        public List<string> Colors { get; set; } = [];
        public string? Title { get; set; }
        public string TitleFontSize { get; set; } = "10pt";
        public bool ShowDataLabels { get; set; }
        public double HoleRatio { get; set; }
        public bool IsStacked { get; set; }
        public bool IsPercent { get; set; }
        public bool Is3D { get; set; }
        public double? AxisMax { get; set; }
        public double? AxisMin { get; set; }
        public double? MajorUnit { get; set; }
        public int? GapWidth { get; set; }
        public string? ValAxisTitle { get; set; }
        public int ValAxisTitleFontPx { get; set; } = 9;
        public bool ValAxisTitleBold { get; set; }
        public string? CatAxisTitle { get; set; }
        public int CatAxisTitleFontPx { get; set; } = 9;
        public bool CatAxisTitleBold { get; set; }
        public string? PlotFillColor { get; set; }
        public string? ChartFillColor { get; set; }
        public bool HasLegend { get; set; }
        public string LegendFontSize { get; set; } = "8pt";
        public string? LegendFontColor { get; set; }
        public int ValFontPx { get; set; } = 9;
        public string? ValFontColor { get; set; }
        public int CatFontPx { get; set; } = 9;
        public string? CatFontColor { get; set; }
        public string? ValNumFmt { get; set; }
        public string? TitleFontColor { get; set; }
        public string? GridlineColor { get; set; }
        public string? AxisLineColor { get; set; }
        public int DataLabelFontPx { get; set; } = 8;
    }

    /// <summary>Extract all chart metadata from OOXML PlotArea and Chart elements.</summary>
    public static ChartInfo ExtractChartInfo(OpenXmlElement plotArea, OpenXmlElement? chart)
    {
        var info = new ChartInfo();
        info.PlotArea = plotArea as PlotArea;
        if (info.PlotArea == null) return info;

        // Chart type, categories, series
        info.ChartType = ChartHelper.DetectChartType(info.PlotArea) ?? "column";
        info.Categories = ChartHelper.ReadCategories(info.PlotArea) ?? [];
        info.Series = ChartHelper.ReadAllSeries(info.PlotArea);
        if (info.Series.Count == 0) return info;

        info.Is3D = info.ChartType.Contains("3d");
        info.IsStacked = info.ChartType.Contains("stacked") || info.ChartType.Contains("Stacked");
        info.IsPercent = info.ChartType.Contains("percent") || info.ChartType.Contains("Percent");

        // Locate chart type element (barChart, lineChart, pieChart, etc.)
        var chartTypeEl = plotArea.Elements().FirstOrDefault(e =>
            e.LocalName is "barChart" or "bar3DChart" or "lineChart" or "line3DChart"
                or "pieChart" or "pie3DChart" or "doughnutChart" or "areaChart" or "area3DChart"
                or "scatterChart" or "radarChart" or "bubbleChart" or "ofPieChart"
                or "stockChart");

        // Colors
        var isPieType = info.ChartType.Contains("pie") || info.ChartType.Contains("doughnut");
        var serElements = chartTypeEl?.Elements().Where(e => e.LocalName == "ser").ToList() ?? [];
        info.Colors = ExtractColors(serElements, info.Series, isPieType, info.ChartType);

        // Title
        var titleEl = chart?.Elements().FirstOrDefault(e => e.LocalName == "title");
        if (titleEl != null)
        {
            var titleRuns = titleEl.Descendants<Drawing.Run>()
                .Select(r => r.GetFirstChild<Drawing.Text>()?.Text)
                .Where(t => t != null);
            info.Title = string.Join("", titleRuns);
            var titleRPr = titleEl.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (titleRPr?.FontSize?.HasValue == true)
                info.TitleFontSize = $"{titleRPr.FontSize.Value / 100.0:0.##}pt";
            info.TitleFontColor = ExtractFontColor(titleRPr);
        }

        // Data labels
        var dLbls = chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "dLbls")
            ?? plotArea.Descendants().FirstOrDefault(e => e.LocalName == "dLbls");
        if (dLbls != null)
        {
            info.ShowDataLabels = dLbls.Elements().Any(e =>
                (e.LocalName is "showVal" or "showPercent" or "showCatName")
                && e.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "1");
        }

        // Doughnut hole size
        if (info.ChartType.Contains("doughnut"))
        {
            var holeSizeEl = chartTypeEl?.Elements().FirstOrDefault(e => e.LocalName == "holeSize");
            var holeSizeVal = holeSizeEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.HoleRatio = (holeSizeVal != null && int.TryParse(holeSizeVal, out var hs) ? hs : 50) / 100.0;
        }

        // Axis info
        var valAxis = plotArea.Elements().FirstOrDefault(e => e.LocalName == "valAx");
        var catAxis = plotArea.Elements().FirstOrDefault(e => e.LocalName == "catAx");

        if (valAxis != null)
        {
            var valTitleEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "title");
            info.ValAxisTitle = valTitleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
            var valTitleRPr = valTitleEl?.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (valTitleRPr?.FontSize?.HasValue == true)
                info.ValAxisTitleFontPx = (int)(valTitleRPr.FontSize.Value / 100.0);
            if (valTitleRPr?.Bold?.Value == true)
                info.ValAxisTitleBold = true;
            var scaling = valAxis.Elements().FirstOrDefault(e => e.LocalName == "scaling");
            if (scaling != null)
            {
                var maxEl = scaling.Elements().FirstOrDefault(e => e.LocalName == "max");
                var minEl = scaling.Elements().FirstOrDefault(e => e.LocalName == "min");
                if (maxEl != null && double.TryParse(maxEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var maxV))
                    info.AxisMax = maxV;
                if (minEl != null && double.TryParse(minEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var minV))
                    info.AxisMin = minV;
            }
            var majorUnit = valAxis.Elements().FirstOrDefault(e => e.LocalName == "majorUnit");
            if (majorUnit != null && double.TryParse(majorUnit.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value, out var mu))
                info.MajorUnit = mu;

            // Use txPr > defRPr for tick label font (not title's RunProperties)
            var valTxPr = valAxis.Elements().FirstOrDefault(e => e.LocalName == "txPr");
            var valDefRPr = valTxPr?.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            if (valDefRPr?.FontSize?.HasValue == true)
                info.ValFontPx = (int)(valDefRPr.FontSize.Value / 100.0);
            info.ValFontColor = ExtractFontColor(valDefRPr);

            // Gridline color
            var majorGridlines = valAxis.Elements().FirstOrDefault(e => e.LocalName == "majorGridlines");
            var gridSpPr = majorGridlines?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
            info.GridlineColor = ExtractLineColor(gridSpPr);

            // Axis line color
            var valSpPr = valAxis.Elements().FirstOrDefault(e => e.LocalName == "spPr");
            info.AxisLineColor = ExtractLineColor(valSpPr);

            // Value axis number format (e.g. "$#,##0")
            var numFmtEl = valAxis.Elements().FirstOrDefault(e => e.LocalName == "numFmt");
            var fmtCode = numFmtEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "formatCode").Value;
            if (!string.IsNullOrEmpty(fmtCode) && fmtCode != "General")
                info.ValNumFmt = fmtCode;
        }
        if (catAxis != null)
        {
            var catTitleEl = catAxis.Elements().FirstOrDefault(e => e.LocalName == "title");
            info.CatAxisTitle = catTitleEl?.Descendants<Drawing.Text>().FirstOrDefault()?.Text;
            var catTitleRPr = catTitleEl?.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (catTitleRPr?.FontSize?.HasValue == true)
                info.CatAxisTitleFontPx = (int)(catTitleRPr.FontSize.Value / 100.0);
            if (catTitleRPr?.Bold?.Value == true)
                info.CatAxisTitleBold = true;
            // Use txPr > defRPr for tick label font (not title's RunProperties)
            var catTxPr = catAxis.Elements().FirstOrDefault(e => e.LocalName == "txPr");
            var catDefRPr = catTxPr?.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            if (catDefRPr?.FontSize?.HasValue == true)
                info.CatFontPx = (int)(catDefRPr.FontSize.Value / 100.0);
            info.CatFontColor = ExtractFontColor(catDefRPr);
        }

        // Data label font size
        if (dLbls != null)
        {
            var dLblDefRPr = dLbls.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            var dLblFontSize = dLblDefRPr?.FontSize ?? dLbls.Descendants<Drawing.RunProperties>().FirstOrDefault()?.FontSize;
            if (dLblFontSize?.HasValue == true)
                info.DataLabelFontPx = (int)(dLblFontSize.Value / 100.0);
        }

        // Gap width
        var gapWidthEl = plotArea.Descendants().FirstOrDefault(e => e.LocalName == "gapWidth");
        if (gapWidthEl != null)
        {
            var gv = gapWidthEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (gv != null && int.TryParse(gv, out var gw)) info.GapWidth = gw;
        }

        // Plot / chart fill
        var plotSpPr = plotArea.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        info.PlotFillColor = ExtractFillColor(plotSpPr);
        var chartSpPr = chart?.Parent?.Elements().FirstOrDefault(e => e.LocalName == "spPr");
        info.ChartFillColor = ExtractFillColor(chartSpPr);

        // Legend
        var legendEl = chart?.Elements().FirstOrDefault(e => e.LocalName == "legend");
        if (legendEl != null)
        {
            var deleteEl = legendEl.Elements().FirstOrDefault(e => e.LocalName == "delete");
            var delVal = deleteEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            info.HasLegend = delVal != "1";
            var legendRPr = legendEl.Descendants<Drawing.RunProperties>().FirstOrDefault();
            if (legendRPr?.FontSize?.HasValue == true)
                info.LegendFontSize = $"{legendRPr.FontSize.Value / 100.0:0.##}pt";
            info.LegendFontColor = ExtractFontColor(legendRPr);
        }
        else
        {
            info.HasLegend = info.Series.Count > 1 || isPieType;
        }

        return info;
    }

    /// <summary>Extract series colors (per-point for pie/doughnut, stroke for line/scatter, fill for others).</summary>
    private static List<string> ExtractColors(List<OpenXmlElement> serElements, List<(string name, double[] values)> series,
        bool isPieType, string chartType)
    {
        var colors = new List<string>();

        if (isPieType && serElements.Count > 0)
        {
            // Pie/doughnut: colors are per data point (dPt), not per series
            var ser = serElements[0];
            var dPts = ser.Elements().Where(e => e.LocalName == "dPt").ToList();
            var catCount = series.FirstOrDefault().values?.Length ?? 0;
            for (int i = 0; i < catCount; i++)
            {
                var dPt = dPts.FirstOrDefault(d =>
                {
                    var idxEl = d.Elements().FirstOrDefault(e => e.LocalName == "idx");
                    if (idxEl == null) return false;
                    return idxEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == i.ToString();
                });
                var rgb = ExtractFillColor(dPt?.Elements().FirstOrDefault(e => e.LocalName == "spPr"));
                colors.Add(rgb != null ? $"#{rgb}" : DefaultColors[i % DefaultColors.Length]);
            }
        }
        else
        {
            // Detect line/scatter series for stroke color extraction
            var isLineType = chartType.Contains("line") || chartType == "scatter";
            for (int i = 0; i < series.Count; i++)
            {
                string? rgb = null;
                if (i < serElements.Count)
                {
                    var spPr = serElements[i].Elements().FirstOrDefault(e => e.LocalName == "spPr");
                    if (isLineType)
                    {
                        // For line/scatter, prefer stroke color from a:ln > a:solidFill
                        var ln = spPr?.Elements().FirstOrDefault(e => e.LocalName == "ln");
                        rgb = ExtractFillColor(ln);
                    }
                    // Fallback to solidFill
                    rgb ??= ExtractFillColor(spPr);
                }
                colors.Add(rgb != null ? $"#{rgb}" : DefaultColors[i % DefaultColors.Length]);
            }
        }
        return colors;
    }

    /// <summary>Extract hex color (without #) from solidFill > srgbClr inside an spPr or ln element.</summary>
    private static string? ExtractFillColor(OpenXmlElement? container)
    {
        if (container == null) return null;
        var solidFill = container.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        return srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
    }

    /// <summary>Extract font color from RunProperties or DefaultRunProperties (solidFill > srgbClr).</summary>
    private static string? ExtractFontColor(OpenXmlElement? rPr)
    {
        if (rPr == null) return null;
        var solidFill = rPr.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var val = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return val != null ? $"#{val}" : null;
    }

    /// <summary>Extract line/outline color from spPr (ln > solidFill > srgbClr).</summary>
    private static string? ExtractLineColor(OpenXmlElement? spPr)
    {
        if (spPr == null) return null;
        var ln = spPr.Elements().FirstOrDefault(e => e.LocalName == "ln");
        if (ln == null) return null;
        var solidFill = ln.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
        var srgb = solidFill?.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
        var val = srgb?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return val != null ? $"#{val}" : null;
    }

    /// <summary>Render the chart SVG content (inside an already-opened svg tag) based on ChartInfo.</summary>
    public void RenderChartSvgContent(StringBuilder sb, ChartInfo info, int svgW, int svgH,
        int marginLeft = 45, int marginTop = 10, int marginRight = 15, int marginBottom = 30)
    {
        // Sync instance font sizes and colors from ChartInfo
        ValFontPx = info.ValFontPx;
        CatFontPx = info.CatFontPx;
        if (info.ValFontColor != null) AxisColor = info.ValFontColor;
        if (info.CatFontColor != null) CatColor = info.CatFontColor;
        if (info.GridlineColor != null) GridColor = info.GridlineColor;
        if (info.AxisLineColor != null) AxisLineColor = info.AxisLineColor;
        DataLabelFontPx = info.DataLabelFontPx;

        // Increase right margin for long axis labels (e.g. "$1,000,000")
        if (!string.IsNullOrEmpty(info.ValNumFmt) && marginRight < 30)
            marginRight = 30;

        var plotW = svgW - marginLeft - marginRight;
        var plotH = svgH - marginTop - marginBottom;
        if (plotW < 10 || plotH < 10) return;

        var chartType = info.ChartType;

        // Plot area background — for horizontal bar charts, defer to RenderBarChartSvg (labels are outside plot)
        var isHorizBarType = chartType.Contains("bar") && !chartType.Contains("column");
        if (info.PlotFillColor != null && !isHorizBarType)
            sb.AppendLine($"    <rect x=\"{marginLeft}\" y=\"{marginTop}\" width=\"{plotW}\" height=\"{plotH}\" fill=\"#{info.PlotFillColor}\"/>");

        if (chartType.Contains("pie") || chartType.Contains("doughnut"))
        {
            if (info.Is3D)
                RenderPie3DSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH);
            else
                RenderPieChartSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH, info.HoleRatio, info.ShowDataLabels);
        }
        else if (chartType.Contains("area"))
        {
            var areaW = plotW - (int)(plotW * 0.03);
            RenderAreaChartSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, areaW, plotH, info.IsStacked);
        }
        else if (chartType == "combo")
        {
            RenderComboChartSvg(sb, info.PlotArea!, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH);
        }
        else if (chartType.Contains("radar"))
        {
            RenderRadarChartSvg(sb, info.Series, info.Categories, info.Colors, svgW, svgH, CatFontPx);
        }
        else if (chartType == "bubble")
        {
            RenderBubbleChartSvg(sb, info.PlotArea!, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH);
        }
        else if (chartType == "stock")
        {
            RenderStockChartSvg(sb, info.PlotArea!, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH);
        }
        else if (chartType.Contains("line") || chartType == "scatter")
        {
            if (info.Is3D)
                RenderLine3DSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH);
            else
                RenderLineChartSvg(sb, info.Series, info.Categories, info.Colors, marginLeft, marginTop, plotW, plotH, info.ShowDataLabels);
        }
        else
        {
            // Column/bar variants
            var isHorizontal = chartType.Contains("bar") && !chartType.Contains("column");
            // Horizontal bars have their own hLabelMargin inside, so reduce outer marginLeft
            var barMarginLeft = isHorizontal ? 5 : marginLeft;
            var barPlotW = isHorizontal ? svgW - barMarginLeft - marginRight : plotW;
            if (info.Is3D && !info.IsStacked)
                RenderBar3DSvg(sb, info.Series, info.Categories, info.Colors, barMarginLeft, marginTop, barPlotW, plotH, isHorizontal);
            else
                RenderBarChartSvg(sb, info.Series, info.Categories, info.Colors, barMarginLeft, marginTop, barPlotW, plotH,
                    isHorizontal, info.IsStacked, info.IsPercent, info.AxisMax, info.AxisMin, info.MajorUnit,
                    info.GapWidth, ValFontPx, CatFontPx, info.ShowDataLabels, info.ValNumFmt,
                    isHorizontal ? info.PlotFillColor : null);
        }

        // Axis titles inside SVG — for horizontal bar charts, value axis is on bottom and category axis is on left
        var isHorizBar = chartType.Contains("bar") && !chartType.Contains("column");
        var bottomTitle = isHorizBar ? info.ValAxisTitle : info.CatAxisTitle;
        var bottomTitleFont = isHorizBar ? info.ValAxisTitleFontPx : info.CatAxisTitleFontPx;
        var bottomTitleBold = isHorizBar ? info.ValAxisTitleBold : info.CatAxisTitleBold;
        var leftTitle = isHorizBar ? info.CatAxisTitle : info.ValAxisTitle;
        var leftTitleFont = isHorizBar ? info.CatAxisTitleFontPx : info.ValAxisTitleFontPx;
        var leftTitleBold = isHorizBar ? info.CatAxisTitleBold : info.ValAxisTitleBold;
        if (!string.IsNullOrEmpty(leftTitle))
            sb.AppendLine($"    <text x=\"10\" y=\"{svgH / 2}\" fill=\"{AxisColor}\" font-size=\"{leftTitleFont}\"{(leftTitleBold ? " font-weight=\"bold\"" : "")} text-anchor=\"middle\" dominant-baseline=\"middle\" transform=\"rotate(-90,10,{svgH / 2})\">{HtmlEncode(leftTitle)}</text>");
        if (!string.IsNullOrEmpty(bottomTitle))
            sb.AppendLine($"    <text x=\"{svgW / 2}\" y=\"{svgH - 2}\" fill=\"{AxisColor}\" font-size=\"{bottomTitleFont}\"{(bottomTitleBold ? " font-weight=\"bold\"" : "")} text-anchor=\"middle\">{HtmlEncode(bottomTitle)}</text>");
    }

    /// <summary>Render chart legend HTML (outside the svg tag).</summary>
    public void RenderLegendHtml(StringBuilder sb, ChartInfo info, string fontColor = "#555")
    {
        if (!info.HasLegend) return;
        var isPieType = info.ChartType.Contains("pie") || info.ChartType.Contains("doughnut");
        sb.Append($"<div style=\"display:flex;justify-content:center;gap:16px;padding:4px 0;font-size:{info.LegendFontSize};color:{fontColor}\">");
        if (isPieType && info.Categories.Length > 0)
        {
            for (int i = 0; i < info.Categories.Length; i++)
            {
                var color = i < info.Colors.Count ? info.Colors[i] : DefaultColors[i % DefaultColors.Length];
                sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><span style=\"display:inline-block;width:12px;height:12px;background:{color};border-radius:1px\"></span>{HtmlEncode(info.Categories[i])}</span>");
            }
        }
        else
        {
            for (int i = 0; i < info.Series.Count; i++)
            {
                var color = i < info.Colors.Count ? info.Colors[i] : DefaultColors[i % DefaultColors.Length];
                sb.Append($"<span style=\"display:inline-flex;align-items:center;gap:4px\"><span style=\"display:inline-block;width:12px;height:12px;background:{color};border-radius:1px\"></span>{HtmlEncode(info.Series[i].name)}</span>");
            }
        }
        sb.AppendLine("</div>");
    }

    // ==================== 3D Chart Helpers ====================

    /// <summary>Darken or lighten a hex color by a factor (0.0-2.0, 1.0=unchanged)</summary>
    private static string AdjustColor(string hexColor, double factor)
    {
        var hex = hexColor.TrimStart('#');
        if (hex.Length < 6) return hexColor;
        var r = (int)Math.Clamp(int.Parse(hex[..2], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        var g = (int)Math.Clamp(int.Parse(hex[2..4], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        var b = (int)Math.Clamp(int.Parse(hex[4..6], System.Globalization.NumberStyles.HexNumber) * factor, 0, 255);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    // 3D isometric offsets
    private const double Depth3D = 12;
    private const double DxIso = 8;
    private const double DyIso = -6;

    private void RenderBar3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph, bool horizontal)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var (maxVal, _, _) = ComputeNiceAxis(allValues.Max());
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));
        var serCount = series.Count;

        if (horizontal)
        {
            // Estimate label width from longest category name (approx 0.5 × fontSize per char)
            var maxLabelLen = categories.Length > 0 ? categories.Max(c => c.Length) : 0;
            var hLabelMargin = (int)(maxLabelLen * CatFontPx * 0.5) + 4;
            var plotOx = ox + hLabelMargin;
            var plotPw = pw - hLabelMargin;
            var groupH = (double)ph / Math.Max(catCount, 1);
            var barH = groupH * 0.5 / serCount;
            var gap = groupH * 0.2;

            for (int t = 1; t <= 4; t++)
            {
                var gx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <line x1=\"{gx:0.#}\" y1=\"{oy}\" x2=\"{gx:0.#}\" y2=\"{oy + ph}\" stroke=\"{GridColor}\" stroke-width=\"0.5\" stroke-dasharray=\"none\"/>");
            }
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy}\" x2=\"{plotOx}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{plotOx}\" y1=\"{oy + ph}\" x2=\"{plotOx + plotPw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int s = 0; s < serCount; s++)
            {
                var color = colors[s % colors.Count];
                var sideColor = AdjustColor(color, 0.7);
                var topColor = AdjustColor(color, 1.3);
                for (int c = 0; c < series[s].values.Length && c < catCount; c++)
                {
                    var val = series[s].values[c];
                    var barW = (val / maxVal) * plotPw;
                    var bx = plotOx;
                    var by = oy + c * groupH + gap + s * barH;
                    sb.AppendLine($"        <polygon points=\"{bx:0.#},{by:0.#} {bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + DxIso:0.#},{by + DyIso:0.#}\" fill=\"{topColor}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"{color}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <polygon points=\"{bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + barW + DxIso:0.#},{by + barH + DyIso:0.#} {bx + barW:0.#},{by + barH:0.#}\" fill=\"{sideColor}\" opacity=\"0.9\"/>");
                    var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                    sb.AppendLine($"        <text x=\"{bx + barW + DxIso + 4:0.#}\" y=\"{by + barH / 2:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"start\" dominant-baseline=\"middle\">{vlabel}</text>");
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var ly = oy + c * groupH + groupH / 2;
                sb.AppendLine($"        <text x=\"{plotOx - 4}\" y=\"{ly:0.#}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var tx = plotOx + (double)plotPw * t / 4;
                sb.AppendLine($"        <text x=\"{tx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"middle\">{label}</text>");
            }
        }
        else
        {
            var groupW = (double)pw / Math.Max(catCount, 1);
            var barW = groupW * 0.5 / serCount;
            var gap = groupW * 0.2;

            for (int t = 1; t <= 4; t++)
            {
                var gy = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{gy:0.#}\" x2=\"{ox + pw}\" y2=\"{gy:0.#}\" stroke=\"{GridColor}\" stroke-width=\"0.5\" stroke-dasharray=\"none\"/>");
            }
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
            sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

            for (int c = 0; c < catCount; c++)
            {
                for (int s = 0; s < serCount; s++)
                {
                    if (c >= series[s].values.Length) continue;
                    var val = series[s].values[c];
                    var color = colors[s % colors.Count];
                    var sideColor = AdjustColor(color, 0.65);
                    var topColor = AdjustColor(color, 1.25);
                    var barH2 = (val / maxVal) * ph;
                    var bx = ox + c * groupW + gap + s * barW;
                    var by = oy + ph - barH2;

                    sb.AppendLine($"        <rect x=\"{bx:0.#}\" y=\"{by:0.#}\" width=\"{barW:0.#}\" height=\"{barH2:0.#}\" fill=\"{color}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <polygon points=\"{bx:0.#},{by:0.#} {bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + DxIso:0.#},{by + DyIso:0.#}\" fill=\"{topColor}\" opacity=\"0.9\"/>");
                    sb.AppendLine($"        <polygon points=\"{bx + barW:0.#},{by:0.#} {bx + barW + DxIso:0.#},{by + DyIso:0.#} {bx + barW + DxIso:0.#},{oy + ph + DyIso:0.#} {bx + barW:0.#},{oy + ph:0.#}\" fill=\"{sideColor}\" opacity=\"0.9\"/>");
                    var vlabel = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                    sb.AppendLine($"        <text x=\"{bx + barW / 2 + DxIso / 2:0.#}\" y=\"{by + DyIso - 3:0.#}\" fill=\"{ValueColor}\" font-size=\"{DataLabelFontPx}\" text-anchor=\"middle\">{vlabel}</text>");
                }
            }
            for (int c = 0; c < catCount; c++)
            {
                var label = c < categories.Length ? categories[c] : "";
                var lx = ox + c * groupW + groupW / 2;
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
            }
            for (int t = 0; t <= 4; t++)
            {
                var val = maxVal * t / 4;
                var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
                var ty = oy + ph - (double)ph * t / 4;
                sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
            }
        }
    }

    private void RenderPie3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int svgW, int svgH)
    {
        var values = series.FirstOrDefault().values ?? [];
        if (values.Length == 0) return;
        var total = values.Sum();
        if (total <= 0) return;

        var cx = svgW / 2.0;
        var cy = svgH / 2.0;
        var rx = Math.Min(svgW, svgH) * 0.35;
        var ry = rx * 0.55;
        var depth = rx * 0.15;
        var startAngle = -Math.PI / 2;

        var slices = new List<(int idx, double start, double end, string color)>();
        var angle = startAngle;
        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var color = i < colors.Count ? colors[i] : DefaultColors[i % DefaultColors.Length];
            slices.Add((i, angle, angle + sliceAngle, color));
            angle += sliceAngle;
        }

        foreach (var (idx, start, end, color) in slices)
        {
            var sideColor = AdjustColor(color, 0.6);
            if (start < Math.PI && end > 0)
            {
                var clampedStart = Math.Max(start, -0.01);
                var clampedEnd = Math.Min(end, Math.PI + 0.01);
                var steps = Math.Max(8, (int)((clampedEnd - clampedStart) / 0.1));
                var pathPoints = new StringBuilder();
                pathPoints.Append($"M {cx + rx * Math.Cos(clampedStart):0.#},{cy + ry * Math.Sin(clampedStart):0.#} ");
                for (int step = 0; step <= steps; step++)
                {
                    var a = clampedStart + (clampedEnd - clampedStart) * step / steps;
                    pathPoints.Append($"L {cx + rx * Math.Cos(a):0.#},{cy + ry * Math.Sin(a):0.#} ");
                }
                for (int step = steps; step >= 0; step--)
                {
                    var a = clampedStart + (clampedEnd - clampedStart) * step / steps;
                    pathPoints.Append($"L {cx + rx * Math.Cos(a):0.#},{cy + ry * Math.Sin(a) + depth:0.#} ");
                }
                pathPoints.Append("Z");
                sb.AppendLine($"        <path d=\"{pathPoints}\" fill=\"{sideColor}\" opacity=\"0.9\"/>");
            }
        }

        startAngle = -Math.PI / 2;
        for (int i = 0; i < values.Length; i++)
        {
            var sliceAngle = 2 * Math.PI * values[i] / total;
            var endAngle = startAngle + sliceAngle;
            var color = i < colors.Count ? colors[i] : DefaultColors[i % DefaultColors.Length];

            if (values.Length == 1)
                sb.AppendLine($"        <ellipse cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" rx=\"{rx:0.#}\" ry=\"{ry:0.#}\" fill=\"{color}\" opacity=\"0.9\"/>");
            else
            {
                var x1 = cx + rx * Math.Cos(startAngle);
                var y1 = cy + ry * Math.Sin(startAngle);
                var x2 = cx + rx * Math.Cos(endAngle);
                var y2 = cy + ry * Math.Sin(endAngle);
                var largeArc = sliceAngle > Math.PI ? 1 : 0;
                sb.AppendLine($"        <path d=\"M {cx:0.#},{cy:0.#} L {x1:0.#},{y1:0.#} A {rx:0.#},{ry:0.#} 0 {largeArc},1 {x2:0.#},{y2:0.#} Z\" fill=\"{color}\" opacity=\"0.9\"/>");
            }

            var midAngle = startAngle + sliceAngle / 2;
            var lx = cx + rx * 0.55 * Math.Cos(midAngle);
            var ly = cy + ry * 0.55 * Math.Sin(midAngle);
            var label = i < categories.Length ? categories[i] : "";
            if (!string.IsNullOrEmpty(label))
                sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" fill=\"white\" font-size=\"9\" text-anchor=\"middle\" dominant-baseline=\"middle\">{HtmlEncode(label)}</text>");

            startAngle = endAngle;
        }
    }

    private void RenderLine3DSvg(StringBuilder sb, List<(string name, double[] values)> series,
        string[] categories, List<string> colors, int ox, int oy, int pw, int ph)
    {
        var allValues = series.SelectMany(s => s.values).ToArray();
        if (allValues.Length == 0) return;
        var (maxVal, _, _) = ComputeNiceAxis(allValues.Max());
        var catCount = Math.Max(categories.Length, series.Max(s => s.values.Length));

        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy}\" x2=\"{ox}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");
        sb.AppendLine($"        <line x1=\"{ox}\" y1=\"{oy + ph}\" x2=\"{ox + pw}\" y2=\"{oy + ph}\" stroke=\"{AxisLineColor}\" stroke-width=\"1\"/>");

        for (int s = series.Count - 1; s >= 0; s--)
        {
            var color = colors[s % colors.Count];
            var shadowColor = AdjustColor(color, 0.5);
            var points = new List<(double x, double y)>();
            for (int c = 0; c < series[s].values.Length && c < catCount; c++)
            {
                var px = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
                var py = oy + ph - (series[s].values[c] / maxVal) * ph;
                points.Add((px, py));
            }
            if (points.Count > 1)
            {
                var ribbon = new StringBuilder();
                ribbon.Append("M ");
                for (int p = 0; p < points.Count; p++)
                    ribbon.Append($"{points[p].x:0.#},{points[p].y:0.#} L ");
                for (int p = points.Count - 1; p >= 0; p--)
                    ribbon.Append($"{points[p].x + DxIso:0.#},{points[p].y + DyIso:0.#} L ");
                ribbon.Length -= 2;
                ribbon.Append(" Z");
                sb.AppendLine($"        <path d=\"{ribbon}\" fill=\"{shadowColor}\" opacity=\"0.4\"/>");

                var linePoints = string.Join(" ", points.Select(p => $"{p.x:0.#},{p.y:0.#}"));
                sb.AppendLine($"        <polyline points=\"{linePoints}\" fill=\"none\" stroke=\"{color}\" stroke-width=\"2.5\"/>");
                foreach (var pt in points)
                    sb.AppendLine($"        <circle cx=\"{pt.x:0.#}\" cy=\"{pt.y:0.#}\" r=\"3\" fill=\"{color}\"/>");
            }
        }

        for (int c = 0; c < catCount; c++)
        {
            var label = c < categories.Length ? categories[c] : "";
            var lx = ox + (catCount > 1 ? (double)pw * c / (catCount - 1) : pw / 2.0);
            sb.AppendLine($"        <text x=\"{lx:0.#}\" y=\"{oy + ph + 16}\" fill=\"{CatColor}\" font-size=\"{CatFontPx}\" text-anchor=\"middle\">{HtmlEncode(label)}</text>");
        }

        // Y-axis value labels
        for (int t = 0; t <= 4; t++)
        {
            var val = maxVal * t / 4;
            var label = val % 1 == 0 ? $"{(int)val}" : $"{val:0.#}";
            var ty = oy + ph - (double)ph * t / 4;
            sb.AppendLine($"        <text x=\"{ox - 4}\" y=\"{ty:0.#}\" fill=\"{AxisColor}\" font-size=\"{ValFontPx}\" text-anchor=\"end\" dominant-baseline=\"middle\">{label}</text>");
        }
    }
}
